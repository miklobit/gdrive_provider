import csv
import shutil
import os
from tempfile import NamedTemporaryFile

from qgis.core import (QgsVectorLayer, QgsFeature, QgsGeometry, QgsPoint,
                       QgsMapLayerRegistry, QgsFeatureRequest, QgsMessageLog,QgsCoordinateReferenceSystem)

logger = lambda msg: QgsMessageLog.logMessage(msg, 'CSV Provider Example', 1)


class CsvLayer():
    """ Pretend we are a data provider """

    dirty = False
    doing_attr_update = False
    geom_types = ("Point", "LineString", "Polygon")


    def __init__(self, csv_path, crs_def = None, geom_type = None, fields = None):
        """ Initialize the layer by reading the CSV file, creating a memory
        layer, and adding records to it """

        # Save the path to the file soe we can update it in response to edits
        #
        if crs_def:
            self.crs_def = crs_def
        else:
            self.crs_def = QgsCoordinateReferenceSystem("EPSG:4857")
        self.csv_path = csv_path
        self.csv_file = open(csv_path, 'rb')
        self.reader = csv.reader(self.csv_file,delimiter=';', quotechar='"')
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
        self.lyr = QgsVectorLayer(self.uri, os.path.basename(self.csv_path), 'memory')
        self.add_records()
        # done with the csv file
        self.csv_file.close()

        # Make connections
        self.lyr.editingStarted.connect(self.editing_started)
        self.lyr.editingStopped.connect(self.editing_stopped)
        self.lyr.committedAttributeValuesChanges.connect(self.attributes_changed)
        self.lyr.committedFeaturesAdded.connect(self.features_added)
        self.lyr.committedFeaturesRemoved.connect(self.features_removed)
        self.lyr.geometryChanged.connect(self.geometry_changed)

        # Add the layer the map
        QgsMapLayerRegistry.instance().addMapLayer(self.lyr)

    def add_records(self):
        """ Add records to the memory layer by reading the CSV file """
        # Return to beginning of csv file
        self.csv_file.seek(0)
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
            tempfile = NamedTemporaryFile(mode='w', delete=False)
            writer = csv.writer(tempfile, delimiter=';', quotechar='"', lineterminator='\n')
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

            tempfile.close()
            shutil.move(tempfile.name, self.csv_path)

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
