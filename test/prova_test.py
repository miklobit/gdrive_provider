#!/usr/bin/env python

import unittest

import os
import sys

# configure python to play nicely with qgis
osgeo4w_root = r'C:\\OSGeo4W64'
#os.environ['PATH'] = '{}/bin{}{}'.format(osgeo4w_root, os.pathsep, os.environ['PATH'])
#sys.path.insert(0, '{}/apps/qgis/python'.format(osgeo4w_root))
#sys.path.insert(1, '{}/apps/python27/lib/site-packages'.format(osgeo4w_root))

# import PyQGIS
from qgis.core import *
from qgis.gui import *

# import Qt
from PyQt4 import QtCore, QtGui, QtTest
from PyQt4.QtCore import Qt

# disable debug messages
os.environ['QGIS_DEBUG'] = '-1'

def setUpModule():
    # load qgis providers
    QgsApplication.setPrefixPath('{}\\apps\\qgis\\'.format(osgeo4w_root), True)
    qgs = QgsApplication(sys.argv, False)
    qgs.initQgis()


# FIXME: this seems to throw errors
#def tearDownModule():
#    QgsApplication.exitQgis()

# dummy instance to replace qgis.utils.iface
class QgisInterfaceDummy(object):
    def __getattr__(self, name):
        # return an function that accepts any arguments and does nothing
        def dummy(*args, **kwargs):
            return None
        return dummy

class ExamplePluginTest(unittest.TestCase):
    def setUp(self):
        # create a new application instance
        print "SETUP1"
        self.app = app = QtGui.QApplication(sys.argv)
        print "SETUP2"
        # create a map canvas widget
        self.canvas = canvas = QgsMapCanvas()
        canvas.setCanvasColor(QtGui.QColor('white'))
        canvas.enableAntiAliasing(True)
        print "SETUP3"

        # load a shapefile
        self.test_dir = os.path.dirname(__file__)
        self.dataset_dir = os.path.join(self.test_dir,'dataset')
        layers = {}
        for layerName in ['c0601016_SistemiEcorelazionali', 'c0601037_SpecieArboree.shp', 'c0509028_LocSitiContaminati.shp']:
            layers[layerName] = QgsVectorLayer(os.path.join(self.dataset_dir,layerName,'ogr'))
        print "SETUP4"

        # add the layer to the canvas and zoom to it
        QgsMapLayerRegistry.instance().addMapLayer(layer)
        canvas.setLayerSet([QgsMapCanvasLayer(layer)])
        canvas.setExtent(layer.extent())
        print "SETUP5"

        # display the map canvas widget
        #canvas.show()

        print "SETUP6"
        iface = QgisInterfaceDummy()

        # import the plugin to be tested
        from gdrive_provider import Google_Drive_Provider
        from gdrive_layer import GoogleDriveLayer
        from services import google_authorization, service_drive, service_spreadsheet
        print "SETUP7"
        self.plugin = Google_Drive_Provider(iface)
        #self.plugin.initGui()
        #self.dlg = self.plugin.dlg
        #self.dlg.show()
        print "SETUP_FINE"

    def test_dlg_name(self):
        self.assertEqual(self.dlg.windowTitle(), 'Testing')

    def tearDown(self):
        self.plugin.unload()
        del(self.plugin)
        del(self.app) # do not forget this

if __name__ == "__main__":
    unittest.main()

print "PROVATEST"
