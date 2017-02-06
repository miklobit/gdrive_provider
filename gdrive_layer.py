import csv
import shutil
import os
import io
import StringIO
import collections

from tempfile import NamedTemporaryFile

from qgis.core import (QgsVectorLayer, QgsFeature, QgsGeometry, QgsPoint,
                       QgsMapLayerRegistry, QgsFeatureRequest, QgsMessageLog,QgsCoordinateReferenceSystem)

logger = lambda msg: QgsMessageLog.logMessage(msg, 'Googe Drive Provider', 1)


from services import google_authorization, service_drive, service_sheet

class GoogleDriveLayer():
    """ Pretend we are a data provider """

    dirty = False
    doing_attr_update = False
    geom_types = ("Point", "LineString", "Polygon")


    def __init__(self, authorization, layer_name, sheet_id = None, qgis_layer = None, crs_def = None, geom_type = None, fields = None):
        """ Initialize the layer by reading the Google drive sheet, creating a memory
        layer, and adding records to it """

        # Save the path to the file soe we can update it in response to edits
        
        self.service_drive = service_drive(authorization)
        if sheet_id:
            self.sheet_id = sheet_id
        elif qgis_layer:
            result = self.service_drive.upload_sheet(sheetName=qgis_layer.name(), csv_file_obj = self.qgis_layer_to_csv(qgis_layer))
            self.sheet_id = result['id']
        self.reader = self.service_drive.download_sheet(self.sheet_id)
        if crs_def:
            self.crs_def = crs_def
        elif qgis_layer:
            self.crs_def = qgis_layer.crs()
        else: # default to lon/lat geometry
            self.crs_def = QgsCoordinateReferenceSystem("EPSG:4857")
        self.header = self.reader.next()
        logger(str(self.header))
        # Get sample
        sample = self.reader.next()
        self.field_sample = dict(zip(self.header, sample))
        logger("sample %s" % str(self.field_sample))
        field_name_types = {}
        self.geom_type = geom_type
        # create dict of fieldname:type
        for key in self.field_sample.keys():
            if self.field_sample[key].isdigit():
                field_type = 'integer'
            else:
                try:
                    float(self.field_sample[key])
                    field_type = 'real'
                except ValueError:
                    field_type = 'string'
            field_name_types[key] = field_type
        #geometry recognization
        for geom_type in self.geom_types:
            if geom_type.upper() in self.field_sample['WKT'].upper():
                self.geom_type = geom_type
        logger(str(field_name_types))
        # Build up the URI needed to create memory layer
        self.uri = self.uri = "Multi%s?crs=%s" % (self.geom_type, self.crs_def.authid())
        for fld in self.header:
            if fld != 'WKT':
                self.uri += '&field={}:{}'.format(fld, field_name_types[fld])

        logger(self.uri)
        # Create the layer
        self.lyr = QgsVectorLayer(self.uri, layer_name, 'memory')
        self.add_records()

        # Make connections
        self.lyr.editingStarted.connect(self.editing_started)
        self.lyr.editingStopped.connect(self.editing_stopped)
        self.lyr.committedAttributeValuesChanges.connect(self.attributes_changed)
        self.lyr.committedFeaturesAdded.connect(self.features_added)
        self.lyr.committedFeaturesRemoved.connect(self.features_removed)
        self.lyr.geometryChanged.connect(self.geometry_changed)
        
        if qgis_layer:
            self.lyr.setRendererV2(qgis_layer.rendererV2())
        # Add the layer the map
        QgsMapLayerRegistry.instance().addMapLayer(self.lyr)

    def add_records(self):
        """ Add records to the memory layer by reading the Google Drive Sheet """
        # Return to beginning of csv file
        self.reader = self.service_drive.download_sheet(self.sheet_id)
        # Skip the header
        self.reader.next()
        self.lyr.startEditing()

        for row in self.reader:
            flds = dict(zip(self.header, row))
            # logger("This row: %s" % flds)
            feature = QgsFeature()
            wkt_geom = flds.pop('WKT')
            geometry = QgsGeometry.fromWkt(wkt_geom)
            feature.setGeometry(geometry)
            cleared_row = []
            for item in row:
                if not self.geom_type.upper() in item.upper():
                    logger("setting attribute for |%s|" % item)
                    cleared_row.append(item)
            feature.setAttributes(cleared_row)
            self.lyr.addFeature(feature, True)
        self.lyr.commitChanges()

    def editing_started(self):
        """ Connect to the edit buffer so we can capture geometry and attribute
        changes """
        self.lyr.editBuffer().committedAttributeValuesChanges.connect(
            self.attributes_changed)

    def editing_stopped(self):
        """ Update the CSV file if changes were committed """
        if self.dirty:
            logger("Updating the CSV")
            features = self.lyr.getFeatures()
            tempfile = StringIO.StringIO()#NamedTemporaryFile(mode='w', delete=False)
            writer = csv.writer(tempfile, delimiter=',', quotechar='"', lineterminator='\n')
            # write the header
            writer.writerow(self.header)
            for feature in features:
                row = []
                for fld in self.header:
                    if fld == 'WKT':
                        geom = feature.geometry()
                        geom.convertToMultiType()
                        row.append(geom.exportToWkt())
                    else:
                        row.append(feature[feature.fieldNameIndex(fld)])
                print row
                writer.writerow(row)
            #tempfile.close()
            #shutil.move(tempfile.name, self.csv_path)
            print "UPLOAD: ",self.service_drive.upload_sheet(csv_file_obj=tempfile, update_sheetId=self.sheet_id)
            self.dirty = False

    def attributes_changed(self, layer, changes):
        """ Attribute values changed; set the dirty flag """
        if not self.doing_attr_update:
            logger("attributes changed")
            self.dirty = True

    def features_added(self, layer, features):
        """ Features added; update the X and Y attributes for each and set the
        dirty flag
        """
        logger("features added")
        for feature in features:
            self.geometry_changed(feature.id(), feature.geometry())
        self.dirty = True

    def features_removed(self, layer, feature_ids):
        """ Features removed; set the dirty flag """
        logger("features removed")
        self.dirty = True

    def geometry_changed(self, fid, geom):
        """ Geometry for a feature changed; update the X and Y attributes for each """
        feature = self.lyr.getFeatures(QgsFeatureRequest(fid)).next()
        wkt = geom.exportToWkt()
        logger("Updating feature {} WKT attributes to: {}".format(fid, wkt))
        #self.lyr.changeAttributeValue(fid, feature.fieldNameIndex('WKT'),wkt)
        self.dirty = True

    def qgis_layer_to_csv(self,qgis_layer):
        stream = io.BytesIO()
        writer = csv.writer(stream, delimiter=',', quotechar='"', lineterminator='\n')
        row = ["WKT"]
        for feat in qgis_layer.getFeatures():
            for field in feat.fields().toList():
                row.append(field.name().encode("utf-8"))
            break
        writer.writerow(row)
        for feat in qgis_layer.getFeatures():
            row = [feat.geometry().exportToWkt(precision=10)]
            for field in feat.fields().toList():
                if type(feat[field.name()]) == unicode:
                    content = feat[field.name()].encode("utf-8")
                else:
                    content = feat[field.name()]
                row.append(content)
            writer.writerow(row)
        # print stream.getvalue()
        stream.seek(0)
        #csv.reader(stream, delimiter=',', quotechar='"', lineterminator='\n')
        return stream

