# coding=utf-8
"""Dialog test.

.. note:: This program is free software; you can redistribute it and/or modify
     it under the terms of the GNU General Public License as published by
     the Free Software Foundation; either version 2 of the License, or
     (at your option) any later version.

"""

__author__ = 'enricofer@gmail.com'
__date__ = '2017-03-24'
__copyright__ = 'Copyright 2017, Enrico Ferreguti'

import unittest
import os
import httplib2
import site
site.addsitedir(os.path.join(os.getcwd(),'extlibs'))

from utilities import get_qgis_app
QGIS_APP, CANVAS, IFACE, PARENT = get_qgis_app()
from qgis.core import QgsVectorLayer
from qgis.gui import QgsMapCanvas

from PyQt4.QtGui import QDialogButtonBox, QDialog, QWidget
from PyQt4.QtCore import QSettings

#from gdrive_provider import Google_Drive_Provider, SCOPES, CLIENT_SECRET_FILE, APPLICATION_NAME
from gdrive_layer import GoogleDriveLayer
from services import google_authorization, service_drive, service_spreadsheet

#!/usr/bin/env python

from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

import certifi

CLIENT_SECRET = 'client_secret.json'
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive'
APPLICATION_NAME = 'GooGIS plugin'


# dummy instance to replace qgis.utils.iface
class QgisInterfaceDummy(object):
    def __getattr__(self, name):
        # return an function that accepts any arguments and does nothing
        def dummy(*args, **kwargs):
            return None
        return dummy

class test_auth:

    def __init__(self):
        self.credential_dir = os.path.join(os.getcwd(), 'credentials')
        self.store = file.Storage(os.path.join(self.credential_dir,'enricofer_googis plugin.json'))
        self.creds = None #self.store.get()
        self.creds = self.get_credentials()
        self.client_id = 'enricofer@gmail.com'


    def get_credentials(self):
        if not self.creds or self.creds.invalid:
            flow = client.flow_from_clientsecrets(os.path.join(self.credential_dir,CLIENT_SECRET), SCOPES)
            self.creds = tools.run_flow(flow, self.store, tools.argparser.parse_args([]))
        print "INVALID", self.creds.invalid
        return self.creds

    def authorize(self):
        s = QSettings()
        proxyEnabled = 'true' #s.value("proxy/proxyEnabled", "")
        proxyType = 'HttpProxy' #s.value("proxy/proxyType", "")
        proxyHost = '172.20.0.252' #s.value("proxy/proxyHost", "")
        proxyPort = '3128' #s.value("proxy/proxyPort", "")
        proxyUser = 'ferregutie' #s.value("proxy/proxyUser", "")
        proxyPassword = '0an1malOO' #s.value("proxy/proxyPassword", "")
        if proxyEnabled == "true" and proxyType == 'HttpProxy':  # test if there are proxy settings
            proxyConf = httplib2.ProxyInfo(httplib2.socks.PROXY_TYPE_HTTP, proxyHost, int(proxyPort), proxy_user=proxyUser,
                                           proxy_pass=proxyPassword)
        else:
            proxyConf = None
        print "proxyConf",proxyConf
        #httpConnection = httplib2.Http(proxy_info=proxyConf, ca_certs=os.path.join(self.credential_dir, 'cacerts.txt'))
        httpConnection = httplib2.Http(proxy_info=proxyConf)
        auth = self.creds.authorize(httpConnection)
        print "auth", auth
        return


class googisDialogTest(unittest.TestCase):
    """Test dialog works."""

    def setUp(self):
        """Runs before each test."""
        print "setUP_start"
        #self.plugin = Google_Drive_Provider(IFACE)
        #self.plugin.initGui()
        #self.dlg = self.plugin.dlg
        self.iface = IFACE
        print self.iface.messageBar()
        print QGIS_APP, CANVAS, IFACE, PARENT
        self.test_dir = os.path.dirname(__file__)
        self.dataset_dir = os.path.join(self.test_dir,'dataset')
        print os.path.join(os.getcwd(), 'credentials')
        self.layers = {}
        for layerName in ['c0601016_SistemiEcorelazionali', 'c0601037_SpecieArboree.shp', 'c0509028_LocSitiContaminati.shp']:
            self.layers[layerName] = QgsVectorLayer(os.path.join(self.dataset_dir,layerName,'ogr'))
        #self.authorization = google_authorization(self, SCOPES, os.path.join(os.getcwd(), 'credentials'), APPLICATION_NAME, 'enricofer@gmail.com')
        self.authorization = test_auth()
        self.myDrive = service_drive(self.authorization)
        print self.myDrive.list_files()
        self.mySheet = service_spreadsheet(self.authorization,spreadsheetId='1SfHvVxvVRR65pPWRQauAWN4VWUZDGl-yErd3sDizCoA')
        print "setUP_end"

    def tearDown(self):
        """Runs after each test."""
        self.dlg = None

    def test_import_layers(self):
        for layerName,layer in self.layers.iteritems():
            GoogleDriveLayer(self, self.authorization, layerName, importing_layer=layer, test=True)

    def test_populated(self):
        """Test we can click OK."""
        self.assertTrue(self.plugin.available_sheets != [] )


if __name__ == "__main__":
    suite = unittest.makeSuite(googisDialogTest)
    runner = unittest.TextTestRunner(verbosity=3)
    runner.run(suite)

