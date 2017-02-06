# -*- coding: utf-8 -*-
"""
/***************************************************************************
                                 A QGIS plugin
 Example of "faking" a data provider with PyQGIS
                              -------------------
        begin                : 2015-03-13
        git sha              : $Format:%H$
        copyright            : (C) 2015 by GeoApt LLC
        email                : gsherman@geoapt.com
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""
from PyQt4.QtCore import QSettings, QTranslator, qVersion, QCoreApplication, QTimer
from PyQt4.QtGui import QAction, QIcon, QDialog
from qgis.core import QgsMapLayer
# Initialize Qt resources from file resources.py
import resources_rc
# Import the code for the dialog
from ui_internal_browser import Ui_InternalBrowser
from gdrive_provider_dialog import GoogleDriveProviderDialog
from gdrive_layer import GoogleDriveLayer


import os
import sys
import requests
import httplib2


#dependencies

from apiclient import discovery
from apiclient.http import MediaFileUpload
from oauth2client import client, GOOGLE_TOKEN_URI
from oauth2client import tools
from oauth2client.file import Storage


from services import google_authorization, service_drive, service_sheet


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'GooGIS_client_secret.json'
APPLICATION_NAME = 'GooGIS plugin'
CLIENT_ID = 'enricofer@gmail.com'


class OAuth2Verify(QDialog, Ui_InternalBrowser):

    def __init__(self, target, parent = None):
        super(OAuth2Verify, self).__init__(parent)
        self.setupUi(self)
        #self.webView.page().setNetworkAccessManager(QgsNetworkAccessManager.instance())
        if target[0:4] == 'http':
            self.setWindowTitle('Help')
            self.webView.setUrl(QUrl(target))
        else:
            self.setWindowTitle('Auth')
            self.webView.setHtml(target)
            self.timer = QTimer()
            self.timer.setInterval(250)
            self.timer.timeout.connect(self.codeProbe)
            self.timer.start()
            self.auth_code = None
            self.show()
            self.raise_()

    def codeProbe(self):
        frame = self.webView.page().mainFrame()
        frame.evaluateJavaScript('document.getElementById("code").value')
        codeElement = frame.findFirstElement("#code")
        #val = codeElement.evaluateJavaScript("this.value") # redirect urn:ietf:wg:oauth:2.0:oob
        val = self.webView.title().split('=')
        if val[0] == 'Success code':
            self.auth_code = val[1]
            self.accept()
        else:
            self.auth_code = None

    def patchLoginHint(self,loginHint):
        frame = self.webView.page().mainFrame()
        frame.evaluateJavaScript('document.getElementById("Email").value = "%s"' % loginHint)

    @staticmethod
    def getCode(html,loginHint, title=""):
        dialog = OAuth2Verify(html)
        dialog.patchLoginHint(loginHint)
        result = dialog.exec_()
        dialog.timer.stop()
        if result == QDialog.Accepted:
            return dialog.auth_code
        else:
            return None


class Google_Drive_Provider:
    """QGIS Plugin Implementation."""

    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'CSVProvider_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)

            if qVersion() > '4.3.3':
                QCoreApplication.installTranslator(self.translator)

        # Create the dialog (after translation) and keep reference
        self.dlg = GoogleDriveProviderDialog()

        # Declare instance attributes
        self.actions = []
        self.menu = self.tr(u'&Google Drive Provider')
        # TODO: We are going to let the user set this up in a future iteration
        self.toolbar = self.iface.addToolBar(u'GoogleDriveProvider')
        self.toolbar.setObjectName(u'GoogleDriveProvider')

    # noinspection PyMethodMayBeStatic
    def tr(self, message):
        """Get the translation for a string using Qt translation API.

        We implement this ourselves since we do not inherit QObject.

        :param message: String for translation.
        :type message: str, QString

        :returns: Translated version of message.
        :rtype: QString
        """
        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('CSVProvider', message)


    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):
        """Add a toolbar icon to the toolbar.

        :param icon_path: Path to the icon for this action. Can be a resource
            path (e.g. ':/plugins/foo/bar.png') or a normal file system path.
        :type icon_path: str

        :param text: Text that should be shown in menu items for this action.
        :type text: str

        :param callback: Function to be called when the action is triggered.
        :type callback: function

        :param enabled_flag: A flag indicating if the action should be enabled
            by default. Defaults to True.
        :type enabled_flag: bool

        :param add_to_menu: Flag indicating whether the action should also
            be added to the menu. Defaults to True.
        :type add_to_menu: bool

        :param add_to_toolbar: Flag indicating whether the action should also
            be added to the toolbar. Defaults to True.
        :type add_to_toolbar: bool

        :param status_tip: Optional text to show in a popup when mouse pointer
            hovers over the action.
        :type status_tip: str

        :param parent: Parent widget for the new action. Defaults None.
        :type parent: QWidget

        :param whats_this: Optional text to show in the status bar when the
            mouse pointer hovers over the action.

        :returns: The action that was created. Note that the action is also
            added to self.actions list.
        :rtype: QAction
        """

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToVectorMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""

        icon_path = ':/plugins/GoogleDriveProvider/icon.png'
        self.add_action(
            icon_path,
            text=self.tr(u'Google Drive Provider '),
            callback=self.run,
            parent=self.iface.mainWindow())
        self.dlg.listWidget.itemDoubleClicked.connect(self.run)
        #add contextual menu
        self.dup_to_google_drive_action = QAction(QIcon(icon_path), "Duplicate to Google drive layer", self.iface.legendInterface() )
        self.iface.legendInterface().addLegendLayerAction(self.dup_to_google_drive_action, "","01", QgsMapLayer.VectorLayer,True)
        self.dup_to_google_drive_action.triggered.connect(self.dup_to_google_drive)
        #authorize plugin
        self.authorization = google_authorization(SCOPES,os.path.join(self.plugin_dir,'credentials'),APPLICATION_NAME,CLIENT_ID)


    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginVectorMenu(
                self.tr(u'&Google Drive Provider'),
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar
        self.iface.legendInterface().removeLegendLayerAction(self.dup_to_google_drive_action)


    def run(self):
        """Run method that performs all the real work"""
        # show the dialog
        
        self.myDrive = service_drive(self.authorization)
        self.available_sheets = self.myDrive.list_files()
        print self.available_sheets
        self.dlg.listWidget.clear()
        self.dlg.listWidget.addItems(self.available_sheets.keys())
        #self.myDrive.create_googis_sheet_from_csv("/home/enrico/Scrivania/temp9.csv")

        self.dlg.show()
        # Run the dialog event loop
        result = self.dlg.exec_()
        # See if OK was pressed
        if result and self.dlg.listWidget.selectedItems():
            self.load_sheet(self.dlg.listWidget.selectedItems()[0])


            # Make connections
            #self.lyr.editingStarted.connect(self.editing_started)
            #self.lyr.editingStopped.connect(self.editing_stopped)
            #self.lyr.committedAttributeValuesChanges.connect(self.attributes_changed)
            #self.lyr.committedFeaturesAdded.connect(self.features_added)
            #self.lyr.committedFeaturesRemoved.connect(self.features_removed)
            #self.lyr.geometryChanged.connect(self.geometry_changed)

    def load_sheet(self,item):
        sheet_name = item.text()
        sheet_id = self.available_sheets[sheet_name]['id']
        print sheet_id
        self.gdrive_layer = GoogleDriveLayer(self.authorization, sheet_name, sheet_id=sheet_id)

    def dup_to_google_drive(self):
        currentLayer = self.iface.legendInterface().currentLayer()
        print currentLayer.name()
        self.gdrive_layer = GoogleDriveLayer(self.authorization, currentLayer.name(), qgis_layer=currentLayer)
        

    def dum(self):

        '''
        s = QSettings() #getting proxy from qgis options settings
        token = s.value("Google_API_access_token", "")
        print "STORED TOKEN:",token

        #credentials = client.AccessTokenCredentials(token,'qt4-user-agent/1.0',None)
        #credentials = client.GoogleCredentials(token,
                                               'enricofer@gmail.com',
                                               'cdgsfo',
                                               None,
                                               None,
                                               GOOGLE_TOKEN_URI,
                                               'pyqt4-user-agent/1.0',
                                               revoke_uri = None)
        #print "SCOPES",credentials.retrieve_scopes(httpConnection)

        try:
            print "SCOPES",credentials.retrieve_scopes(httpConnection)
        except:
            print "TOKEN INVALID: asking new credentials"
            token = self.get_credentials()
            print "NEW TOKEN:",token
            credentials = client.AccessTokenCredentials(token,'qt4-user-agent/1.0',None)
            s.setValue("Google_API_access_token", token)

        if credentials.access_token_expired:
            print "TOKEN EXPIRED: refreshing"
            credentials.refresh(http)
        elif token == '' or credentials is None or credentials.invalid:
            print "TOKEN INVALID: asking new credentials"
            token = self.get_credentials()
            print "NEW TOKEN:",token
            credentials = client.AccessTokenCredentials(token,'qt4-user-agent/1.0',None)
            s.setValue("Google_API_access_token", token)
        else:
            print "access token ok:",credentials.to_json()
            #print credentials.retrieve_scopes(httpConnection)
        '''

        media_body = MediaFileUpload(csv_path, mimetype='text/csv', resumable=None)
        body = {
            'name': os.path.basename(csv_path),
            'description': 'GooGIS sheet',
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        file = service_drive.files().create(body=body, media_body=media_body).execute()
        print file

        #gs = Sheets.from_files(os.path.join(os.path.dirname(__file__), 'credentials','sheets.googleapis.com-python-quickstart.json'))
        #print gs
        #tab = gs[file['id']]
        #print tab

        list = service_sheets.spreadsheets().values().get(spreadsheetId=file['id'], range='B2:B', valueRenderOption = 'UNFORMATTED_VALUE').execute()
        print list

        sheet_metadata = service_sheets.spreadsheets().get(spreadsheetId=file['id']).execute()
        print sheet_metadata
        update_body = {
            "requests": [{
                "findReplace": {
                               # Finds and replaces data in cells over a range, sheet, or all sheets. # Finds and replaces occurrences of some text with other text.
                                   #"includeFormulas": True or False,
                               # True if the search should include cells with formulas.
                                   # False to skip cells with formulas.
                                   "matchEntireCell": True, #True or False,  # True if the find value should match the entire cell.
                                   "allSheets": None,  # True to find/replace over all sheets.
                                   #"matchCase": True or False,  # True if the search is case sensitive.
                                   "find": '8',  # The value to search.
                                   "range": {
                                        "sheetId": sheet_metadata['sheets'][0]['properties']['sheetId'], # The sheet this range is on.
                                        "startRowIndex": 1, # The start row (inclusive) of the range, or not set if unbounded.
                                        #"endRowIndex": 42, # The end row (exclusive) of the range, or not set if unbounded.
                                        "startColumnIndex": 1, # The start column (inclusive) of the range, or not set if unbounded.
                                        "endColumnIndex": 2, # The end column (exclusive) of the range, or not set if unbounded.
                                    },
                                   #"searchByRegex": None,  # True if the find value is a regex.
                                   # The regular expression and replacement should follow Java regex rules
                                   # at https://docs.oracle.com/javase/8/docs/api/java/util/regex/Pattern.html.
                                   # The replacement string is allowed to refer to capturing groups.
                                   # For example, if one cell has the contents `"Google Sheets"` and another
                                   # has `"Google Docs"`, then searching for `"o.* (.*)"` with a replacement of
                                   # `"$1 Rocks"` would change the contents of the cells to
                                   # `"GSheets Rocks"` and `"GDocs Rocks"` respectively.
                                   #"sheetId": 42,  # The sheet to find/replace over.
                                   #"replacement": '8',  # The value to use as the replacement.
                               }
            }]
        }

        find = service_sheets.spreadsheets().batchUpdate(spreadsheetId=file['id'],body=update_body).execute()
        print find