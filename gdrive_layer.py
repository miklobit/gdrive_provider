import csv
import shutil
import os
import io
import sys
import collections
import base64
import zlib
import _thread
import traceback

from tempfile import NamedTemporaryFile
from PyQt5.QtXml import QDomDocument
from PyQt5.QtWidgets import QProgressBar
from PyQt5.QtCore import QObject, pyqtSignal, QVariant, Qt

from qgis.core import (QgsVectorLayer, QgsFeature, QgsGeometry, QgsExpression, QgsField,
                       QgsProject, QgsFeatureRequest, QgsMessageLog,QgsCoordinateReferenceSystem)

from qgis.gui import QgsMessageBar

import qgis.core

logger = lambda msg: QgsMessageLog.logMessage(msg, 'Googe Drive Provider', 1)


from .services import google_authorization, service_drive, service_spreadsheet


from .utils import slugify


class progressBar:
    def __init__(self, parent, msg = ''):
        self.iface = parent.iface
        widget = self.iface.messageBar().createMessage("GooGIS plugin:",msg)
        progressBar = QProgressBar()
        progressBar.setRange(0,0) #(1,steps)
        progressBar.setValue(0)
        progressBar.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        widget.layout().addWidget(progressBar)
        self.iface.messageBar().pushWidget(widget, QgsMessageBar.INFO, 50)

    def stop(self, msg = ''):
        self.iface.messageBar().clearWidgets()
        message = self.iface.messageBar().createMessage("GooGIS plugin:",msg)
        self.iface.messageBar().pushWidget(message, QgsMessageBar.SUCCESS, 3)

class GoogleDriveLayer(QObject):
    """ Pretend we are a data provider """

    invalidEdit = pyqtSignal()
    dirty = False
    doing_attr_update = False
    geom_types = ("Point", "LineString", "Polygon","Unknown","NoGeometry")


    def __init__(self, parent, authorization, layer_name, spreadsheet_id = None, loading_layer = None, importing_layer = None, crs_def = None, geom_type = None, test = None):
        """ Initialize the layer by reading the Google drive sheet, creating a memory
        layer, and adding records to it """

        super(GoogleDriveLayer, self).__init__()
        # Save the path to the file soe we can update it in response to edits
        self.test = test
        self.parent = parent
        self.iface = parent.iface
        bar = progressBar(self, 'loading google drive layer')
        self.service_drive = service_drive(authorization)
        self.client_id = authorization.client_id
        self.authorization = authorization
        if spreadsheet_id:
            self.spreadsheet_id = spreadsheet_id
            self.service_sheet = service_spreadsheet(authorization, self.spreadsheet_id)
        elif importing_layer:
            layer_as_list = self.qgis_layer_to_list(importing_layer)
            self.service_sheet = service_spreadsheet(authorization, new_sheet_name=importing_layer.name(), new_sheet_data = layer_as_list)
            self.spreadsheet_id = self.service_sheet.spreadsheetId()
            self.service_sheet.set_crs(importing_layer.crs().authid())
            self.service_sheet.set_geom_type(self.geom_types[importing_layer.geometryType()])
            self.service_sheet.set_style(self.layer_style_to_xml(importing_layer))
            self.saveFieldTypes(importing_layer.fields())

        self.reader = self.service_sheet.get_sheet_values()
        self.header = self.reader[0]

        self.crs_def = self.service_sheet.crs()
        self.geom_type = self.service_sheet.geom_type()
        logger("LOADED GOOGLE SHEET LAYER: %s CRS_ID:%s GEOM_type:%s" % (self.service_sheet.name,self.crs_def, self.geom_type))
        #logger( "GEOM_type "+ self.geom_type)
        # Build up the URI needed to create memory layer
        if loading_layer:
            self.lyr = loading_layer
            print ("LOADIN", self.lyr, loading_layer)
            attrIds = [i for i in range(0, self.lyr.fields().count())]
            self.lyr.dataProvider().deleteAttributes(attrIds)
            self.lyr.updateFields()
        else:
            self.uri = self.uri = "Multi%s?crs=%s&index=yes" % (self.geom_type, self.crs_def)
            #logger(self.uri)
            # Create the layer
            self.lyr = QgsVectorLayer(self.uri, layer_name, 'memory')
            self.lyr.setCustomProperty("googleDriveId", self.spreadsheet_id)
        fields_types = self.service_sheet.get_line("ROWS", 1, sheet="settings")
        attributes = []
        for i in range(2,len(self.header)):
            if self.header[i][:8] != 'DELETED_':
                type_pack = fields_types[i].split("|")
                attributes.append(QgsField(name=self.header[i],type=int(type_pack[0]), len=int(type_pack[1]), prec=int(type_pack[2])))
                #self.uri += u'&field={}:{}'.format(fld.decode('utf8'), field_name_types[fld])
        self.lyr.dataProvider().addAttributes(attributes)
        self.lyr.updateFields()

        self.xml_to_layer_style(self.lyr,self.service_sheet.style())
        self.lyr.rendererChanged.connect(self.style_changed)

        self.add_records()

        # Make connections
        self.makeConnections(self.lyr)

        # Add the layer the map
        if not loading_layer:
            QgsProject.instance().addMapLayer(self.lyr)
        else:
            pass
            #self.lyr.editingStarted.emit()
        bar.stop("Layer %s succesfully loaded" % layer_name)

    def makeConnections(self,lyr):
        lyr.editingStarted.connect(self.editing_started)
        lyr.editingStopped.connect(self.editing_stopped)
        lyr.committedAttributesDeleted.connect(self.attributes_deleted)
        lyr.committedAttributesAdded[list].connect(self.attributes_added)
        lyr.committedFeaturesAdded.connect(self.features_added)
        lyr.committedGeometriesChanges.connect(self.geometry_changed)
        lyr.committedAttributeValuesChanges.connect(self.attributes_changed)
        lyr.layerDeleted.connect(self.unsubscribe)
        lyr.beforeCommitChanges.connect(self.inspect_changes)

    def add_records(self):
        """ Add records to the memory layer by reading the Google Sheet """
        self.lyr.startEditing()

        for i, row in enumerate(self.reader[1:]):
            flds = collections.OrderedDict(zip(self.header, row))
            status = flds.pop('STATUS')

            if status != 'D': #non caricare i deleted
                wkt_geom = str(zlib.decompress(base64.b64decode(flds.pop('WKTGEOMETRY'))))
                #fid = int(flds.pop('FEATUREID'))
                feature = QgsFeature()
                geometry = QgsGeometry.fromWkt(wkt_geom)
                feature.setGeometry(geometry)
                cleared_row = [] #[fid]
                for field, attribute in flds.items():
                    if field[:8] != 'DELETED_': #skip deleted fields
                        if attribute == '()':
                            cleared_row.append(qgis.core.NULL)
                        else:
                            cleared_row.append(attribute)
                    else:
                        logger( "DELETED " + field)
                feature.setAttributes(cleared_row)
                self.lyr.addFeature(feature)
        self.lyr.commitChanges()

    def style_changed(self):
        logger( "style changed")
        self.service_sheet.set_style(self.layer_style_to_xml(self.lyr))

    def renew_connection(self):
        '''
        when connection stay alive too long we have to rebuild service
        '''
        try:
            self.service_sheet.sheet_cell("A1")
        except:
            print ("renew authorization")
            self.service_sheet.get_service()


    def update_from_subscription(self):
        bar = progressBar(self, 'updating local layer from remote')
        print ("canEdit", self.service_sheet.canEdit)
        if self.service_sheet.canEdit:
            updates = self.service_sheet.get_line('COLUMNS','A', sheet=self.client_id)
        else:
            new_changes_log_rows = self.service_sheet.get_line("COLUMNS",'A',sheet="changes_log")
            if len(new_changes_log_rows) > len(self.service_sheet.changes_log_rows):
                updates = new_changes_log_rows[-len(new_changes_log_rows)+len(self.service_sheet.changes_log_rows):]
                self.service_sheet.changes_log_rows = new_changes_log_rows
            else:
                updates = []
        if updates:
            self.service_sheet.erase_cells(self.client_id)
        print ("UPDATES", updates)
        for update in updates:
            decode_update = update.split("|")
            if decode_update[0] in ('new_feature', 'delete_feature', 'update_geometry', 'update_attributes'):
                sheet_feature = self.service_sheet.get_line('ROWS',decode_update[1])
                if decode_update[0] == 'new_feature':
                    feat = QgsFeature()
                    geom = QgsGeometry().fromWkt(zlib.decompress(base64.b64decode(sheet_feature[0])))
                    feat.setGeometry(geom)
                    feat.setAttributes(sheet_feature[2:])
                    logger(( "updating from subscription, new_feature: " + str(self.lyr.dataProvider().addFeatures([feat]))))
                else:
                    sheet_feature_id = decode_update[1]
                    feat = self.lyr.getFeatures(QgsFeatureRequest(QgsExpression(' "FEATUREID" = %s' % sheet_feature_id))).next()
                    if   decode_update[0] == 'delete_feature':
                        print ("updating from subscription, delete_feature: " + str(self.lyr.dataProvider().deleteFeatures([feat.id()])))
                    elif decode_update[0] == 'update_geometry':
                        update_set = {feat.id(): QgsGeometry().fromWkt(zlib.decompress(base64.b64decode(sheet_feature[0])))}
                        print ("update_set", update_set)
                        print ("updating from subscription, update_geometry: " + str(self.lyr.dataProvider().changeGeometryValues(update_set)))
                    elif decode_update[0] == 'update_attributes':
                        new_attributes = sheet_feature_id[2:]
                        attributes_map = {}
                        for i in range(0, len(new_attributes)):
                            attributes_map[i] = new_attributes[i]
                        update_map = {feat.id(): attributes_map,}
                        print ("update_map", update_map)
                        print ("updating from subscription, update_attributes: " +(self.lyr.dataProvider().changeAttributeValues(update_map)))
            elif decode_update[0] == 'add_field':
                field_a1_notation = self.service_sheet.header_map[decode_update[1]]
                type_def = self.service_sheet.sheet_cell('settings!%s1' % field_a1_notation)
                type_def_decoded = type_def.split("|")
                new_field = QgsField(name=decode_update[1],type=int(type_def_decoded[0]), len=int(type_def_decoded[1]), prec=int(type_def_decoded[2]))
                print ("updating from subscription, add_field: ", + (self.lyr.dataProvider().addAttributes([new_field])))
                self.lyr.updateFields()
            elif decode_update[0] == 'delete_field':
                print ("updating from subscription, delete_field: " + str(self.lyr.dataProvider().deleteAttributes([self.lyr.dataProvider().fields().fieldNameIndex(decode_update[1])])))
                self.lyr.updateFields()
        bar.stop("local layer updated")

    def editing_started(self):
        """ Connect to the edit buffer so we can capture geometry and attribute
        changes """
        self.renew_connection()
        self.update_from_subscription()
        self.bar = None
        if self.service_sheet.canEdit:
            self.activeThreads = 0
            self.lockedFeatures = []
            self.editing = True
            self.lyr.geometryChanged.connect(self.buffer_geometry_changed)
            self.lyr.attributeValueChanged.connect(self.buffer_attributes_changed)
            self.lyr.beforeCommitChanges.connect(self.catch_deleted)
            self.lyr.beforeRollBack.connect(self.rollBack)
            self.invalidEdit.connect(self.rollBack)
            self.changes_log=[]
        else: #refuse editing if file is read only
            self.lyr.rollBack()

    def buffer_geometry_changed(self,fid,geom):
        if self.editing:
            if self.test:
                self.lock_feature(fid)
            else:
                _thread.start_new_thread(self.lock_feature, (fid,))
            #logger("active threads: "+ str( self.activeThreads))
            #self.lock_feature(fid)

    def buffer_attributes_changed(self,fid,attr_id,value):
        if self.editing:
            if self.test:
                self.lock_feature(fid)
            else:
                _thread.start_new_thread(self.lock_feature, (fid,))
            #logger("active threads: "+ str( self.activeThreads))
            #self.lock_feature(fid)


    def lock_feature(self,fid):
        """
        the row in google sheet linked to feature that has been modified is locked
        filling the the STATUS column with the client_id
        """
        self.activeThreads += 1
        if fid >= 0 and self.activeThreads < 2: # fid <0 means that the change relates to newly created features not yet present in the sheet
            feature_locking = self.lyr.getFeatures(QgsFeatureRequest(fid)).next()
            locking_row_id = feature_locking[0]
            status = self.service_sheet.cell('STATUS', locking_row_id)
            if status in (None,''):
                self.service_sheet.set_cell('STATUS', locking_row_id, self.client_id)
                #logger("feature #%s locked by %s" % (locking_row_id, self.client_id))
                self.lockedFeatures.append(locking_row_id)
            #else:
                #pass
                #logger( "LOCK ERROR", "FEATURE %s IS LOCKED BY: %s, EDITS WILL NOT BE SAVED" % (locking_row_id ,status))
                #self.iface.messageBar().pushMessage("LOCK ERROR", "FEATURE %s IS LOCKED BY: %s, EDITS WILL NOT BE SAVED" % (locking_row_id ,status),level=QgsMessageBar.CRITICAL)
                #self.invalidEdit.emit()
                #self.lyr.rollBack()

        else:
            logger( "editing newly created feature %s" % fid)
        self.activeThreads -= 1

    def rollBack(self):
        """
        before rollback changes status field is cleared
        """
        self.renew_connection()
        print ("ROLLBACK", self.lockedFeatures)
        mods = []
        for row_id in self.lockedFeatures:
            logger("rollback changes on feature #"+str(row_id))
            mods.append(('STATUS', row_id, "()"))
        if mods:
            self.service_sheet.set_multicell(mods)
        self.lyr.beforeRollBack.disconnect(self.rollBack)

        #self.lyr.geometryChanged.disconnect(self.buffer_geometry_changed)
        #self.lyr.attributeValueChanged.disconnect(self.buffer_attributes_changed)
        self.editing = False

    def editing_stopped(self):
        """ Update the file if changes were committed """
        #self.lyr.geometryChanged.disconnect(self.buffer_geometry_changed)
        #self.lyr.attributeValueChanged.disconnect(self.buffer_attributes_changed)
        self.renew_connection()
        if self.service_sheet.canEdit:
            self.service_sheet.advertise(self.changes_log)
        self.editing = False
        if self.bar:
            self.bar.stop("update to remote finished")

    def attributes_changed(self, layer, changes):
        """ Attribute values changed; set the dirty flag """
        if not self.doing_attr_update:
            logger("attributes changed")
            value_mods = []
            status_mods = []
            for fid,attrib_change in changes.items():
                feature_changing = self.lyr.getFeatures(QgsFeatureRequest(fid)).next()
                row_id = feature_changing[0]
                logger ( "changing row: %s" % row_id)
                lock = self.service_sheet.cell('STATUS', row_id)
                if lock and lock != self.client_id:
                    logger( "cant apply edits to feature, locked by " + lock)
                    continue
                for attrib_idx, new_value in attrib_change.items():
                    fieldName = QgsProject.instance().mapLayer(layer).fields().field(attrib_idx).name()
                    if fieldName == 'FEATUREID':
                        logger("can't modify FEATUREID")
                        continue
                    try:
                        cleaned_value = new_value.toString(format = Qt.ISODate)
                    except:
                        if not new_value or new_value == qgis.core.NULL:
                            cleaned_value = '()'
                        else:
                            cleaned_value = new_value
                    value_mods.append((fieldName,row_id, cleaned_value))
                    status_mods.append(("STATUS",row_id, "()"))
                    self.changes_log.append('%s|%s' % ('update_attributes', str(row_id)))
            value_mods_result = self.service_sheet.set_multicell(value_mods,lockBy=self.client_id)
            if value_mods_result:
                self.service_sheet.set_multicell(status_mods)


    def inspect_changes(self):
        '''
        here we can inspect changes before commit them
        self.deleted_list = []
        for deleted in self.lyr.editBuffer().deletedAttributeIds():
            self.deleted_list.append(self.lyr.fields().at(deleted).name())

        logger("attributes_added")
        for field in self.lyr.editBuffer().addedAttributes():
            self.service_sheet.add_column([field.name()], fill_with_null = True)
        '''
        pass

    def attributes_added(self, layer, added):
        """
        Attribute added;
        """
        logger("attributes_added")
        for field in added:
            logger( "ADDED FIELD %s" % field.name())
            self.service_sheet.add_column([field.name()], fill_with_null = True)
            self.service_sheet.add_column(["%d|%d|%d" % (field.type(), field.length(), field.precision())],child_sheet="settings", fill_with_null = None)
            self.changes_log.append('%s|%s' % ('add_field', field.name()))

    def attributes_deleted(self, layer, deleted_ids):
        """ Attribute deleted;"""
        logger("attributes_deleted")
        for deleted in deleted_ids:
            deleted_name = self.service_sheet.mark_field_as_deleted(deleted)
            self.changes_log.append('%s|%s' % ('delete_field', deleted_name))


    def features_added(self, layer, features):
        """ Features added; update the X and Y attributes for each and set the
        dirty flag
        """
        logger("features added")
        for feature in features:
            new_fid = self.service_sheet.new_fid()
            self.lyr.dataProvider().changeAttributeValues({feature.id() : {0: new_fid}})
            feature.setAttribute(0, new_fid)
            new_row_dict = {}.fromkeys(self.service_sheet.header,'()')
            new_row_dict['WKTGEOMETRY'] = base64.b64encode(zlib.compress(feature.geometry().exportToWkt()))
            new_row_dict['STATUS'] = '()'
            print (feature.attributes())
            for i,item in enumerate(feature.attributes()):
                fieldName = self.lyr.fields().at(i).name()
                try:
                    new_row_dict[fieldName] = item.toString(format = Qt.ISODate)
                except:
                    if not item or item == qgis.core.NULL:
                        new_row_dict[fieldName] = '()'
                    else:
                        new_row_dict[fieldName] = item
            result = self.service_sheet.add_row(new_row_dict)
            sheet_new_row = int(result['updates']['updatedRange'].split('!A')[1].split(':')[0])
            self.changes_log.append('%s|%s' % ('new_feature', str(new_fid)))

    def catch_deleted(self):
        self.bar = progressBar(self, 'updating local edits to remote')
        """ Features removed; but before commit """
        deleted_ids = self.lyr.editBuffer().deletedFeatureIds()
        if deleted_ids:
            mods = []
            for fid in deleted_ids:
                removed_feat = self.lyr.dataProvider().getFeatures(QgsFeatureRequest(fid)).next()
                removed_row = removed_feat[0]
                print ("deleting row:",removed_row,removed_feat)
                lock = self.service_sheet.cell('STATUS', removed_row)
                if lock and lock != self.client_id:
                    logger( "feature locked by " + lock)
                    continue
                print ("REMOVED",removed_feat,removed_row)
                mods.append(("STATUS",removed_row,'D'))
                self.changes_log.append('%s|%s' % ('delete_feature', str(removed_row)))
            self.service_sheet.set_multicell(mods,lockBy=self.client_id)


    def geometry_changed(self, layer, geom_map):
        """ Geometry for a feature changed; update the X and Y attributes for each """
        for fid,geom in geom_map.items():
            feature_changing = self.lyr.getFeatures(QgsFeatureRequest(fid)).next()
            row_id = feature_changing[0]
            lock = self.service_sheet.cell('STATUS', row_id)
            if lock and lock != self.client_id:
                logger( "FEATUREID %s not changed. locked by %s" %(row_id,lock))
                continue
            wkt = geom.exportToWkt(precision=10)
            self.service_sheet.set_cell('WKTGEOMETRY',row_id, base64.b64encode(zlib.compress(wkt)))
            self.service_sheet.set_cell('STATUS',row_id, '()')
            logger ("Updated FEATUREID %s geometry" % row_id)
            self.changes_log.append('%s|%s' % ('update_geometry', str(row_id)))

    def unsubscribe(self):
        self.renew_connection()
        self.service_sheet.unsubscribe()

    def qgis_layer_to_csv(self,qgis_layer):
        stream = io.BytesIO()
        writer = csv.writer(stream, delimiter=',', quotechar='"', lineterminator='\n')
        row = ["WKTGEOMETRY","FEATUREID","STATUS"]
        for feat in qgis_layer.getFeatures():
            for field in feat.fields().toList():
                row.append(field.name().encode("utf-8"))
            break
        writer.writerow(row)
        for feat in qgis_layer.getFeatures():
            row = [base64.b64encode(zlib.compress(feat.geometry().exportToWkt(precision=10))),feat.id(),"()"]
            for field in feat.fields().toList():
                if feat[field.name()] == qgis.core.NULL:
                    content = "()"
                else:
                    if type(feat[field.name()]) == unicode:
                        content = feat[field.name()].encode("utf-8")
                    else:
                        content = feat[field.name()]
                row.append(content)
            writer.writerow(row)
        stream.seek(0)
        #csv.reader(stream, delimiter=',', quotechar='"', lineterminator='\n')
        return stream

    def qgis_layer_to_list(self,qgis_layer):
        row = ["WKTGEOMETRY","STATUS","FEATUREID"]
        for feat in qgis_layer.getFeatures():
            for field in feat.fields().toList():
                row.append(unicode(field.name()).encode("utf-8"))# slugify(field.name())
            break
        rows = [row]
        for i,feat in enumerate(qgis_layer.getFeatures()):
            row = [base64.b64encode(zlib.compress(feat.geometry().exportToWkt(precision=10))),"()",i+2] #i+2 to make equal featureid and row
            if len(row[0]) > 10000:
                print (feat.id, len(row))
            if len(row[0]) > 50000: # ignore features with geometry > 50000 bytes zipped
                continue
            for field in feat.fields().toList():
                if feat[field.name()] == qgis.core.NULL:
                    content = "()"
                else:
                    if type(feat[field.name()]) == unicode:
                        content = feat[field.name()].encode("utf-8")
                    elif field.typeName() in ('Date', 'Time'):
                        content = feat[field.name()].toString(format = Qt.ISODate)
                    else:
                        content = feat[field.name()]
                row.append(content)
            rows.append(row)
        #csv.reader(stream, delimiter=',', quotechar='"', lineterminator='\n')
        return rows

    def saveFieldTypes(self,fields):
        types_array = ["","","4|4|0"] #default featureId type to longint
        for field in fields.toList():
            types_array.append("%d|%d|%d" % (field.type(), field.length(), field.precision()))
        print ("FIELDTYPES",self.service_sheet.update_cells('settings!A1',types_array))

    def layer_style_to_xml(self,qgis_layer):
        XMLDocument = QDomDocument("qgis_style")
        XMLStyleNode = XMLDocument.createElement("style")
        XMLDocument.appendChild(XMLStyleNode)
        error = None
        qgis_layer.writeSymbology(XMLStyleNode, XMLDocument, error)
        xmldoc = XMLDocument.toString(1)
        return xmldoc

    def xml_to_layer_style(self,qgis_layer,xml):
        XMLDocument = QDomDocument()
        error = None
        XMLDocument.setContent(xml)
        XMLStyleNode = XMLDocument.namedItem("style")
        qgis_layer.readSymbology(XMLStyleNode, error)

    def get_gdrive_id(self):
        return self.spreadsheet_id

    def get_service_drive(self):
        return self.service_drive

    def get_service_sheet(self):
        return self.service_sheet