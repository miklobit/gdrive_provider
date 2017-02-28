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
from PyQt5.QtCore import QSettings, QTranslator, qVersion, QCoreApplication, QTimer, QUrl
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QAction, QProgressBar, QDialog
from qgis.core import QgsMapLayer, QgsVectorLayer, QgsProject
from qgis.utils import plugins
# Initialize Qt resources from file resources.py
import resources_rc
# Import the code for the dialog
from ui_internal_browser import Ui_InternalBrowser
from gdrive_provider_dialog import GoogleDriveProviderDialog
from gdrive_layer import progressBar, GoogleDriveLayer


import os
import sys
import json
import io

from services import google_authorization, service_drive, service_spreadsheet


# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly https://www.googleapis.com/auth/drive.readonly https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive'
CLIENT_SECRET_FILE = 'GooGIS_client_secret.json'
APPLICATION_NAME = 'GooGIS plugin'
CLIENT_ID = 'fasulloef@gmail.com'
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
        self.add_action(
            ':/plugins/GoogleDriveProvider/test.png',
            text=self.tr(u'Google Drive Provider test '),
            callback=self.test_suite,
            parent=self.iface.mainWindow())

        self.dlg.listWidget.itemDoubleClicked.connect(self.run)
        self.dlg.refreshButton.clicked.connect(self.refresh_available)
        #add contextual menu
        self.dup_to_google_drive_action = QAction(QIcon(icon_path), "Duplicate to Google drive layer", self.iface.legendInterface() )
        self.iface.legendInterface().addLegendLayerAction(self.dup_to_google_drive_action, "","01", QgsMapLayer.VectorLayer,True)
        self.dup_to_google_drive_action.triggered.connect(self.dup_to_google_drive)
        #authorize plugin
        self.authorization = google_authorization(SCOPES,os.path.join(self.plugin_dir,'credentials'),APPLICATION_NAME,CLIENT_ID)
        #QgsProject.instance().layerLoaded.connect(self.loadGDriveLayers)
        QgsProject.instance().readProject.connect(self.loadGDriveLayers)


    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        QgsProject.instance().readProject.disconnect(self.loadGDriveLayers)
        for action in self.actions:
            self.iface.removePluginVectorMenu(
                self.tr(u'&Google Drive Provider'),
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar
        self.iface.legendInterface().removeLegendLayerAction(self.dup_to_google_drive_action)

    def GooGISLayers(self):
        for layer in QgsProject.instance().mapLayers().values():
            if self.isGooGISLayer(layer):
                yield layer

    def isGooGISLayer(self, layer):
        if layer.type() != QgsMapLayer.VectorLayer:
            return
        use = layer.customProperty("googleDriveId")
        return not (use == False)

    def loadGDriveLayers(self,dom):
        for layer in self.GooGISLayers():
            google_id = layer.customProperty("googleDriveId", defaultValue=None)
            if google_id:
                self.gdrive_layer = GoogleDriveLayer(self, self.authorization, layer.name(), spreadsheet_id=google_id, loading_layer=layer)
                print ("reading", google_id, layer.id(), self.gdrive_layer.lyr.id())
                #glayer.makeConnections(layer)
                layer.editingStarted.connect(self.gdrive_layer.editing_started)
                layer.updateExtents()



    def test_suite(self):
        #self.sheet_layer = GoogleDriveLayer(self.authorization, sheet_name, sheet_id='1hC8iT7IutoYDVDLlEWF8_op2viNRsUdv8tTVo9RlPkE')
        #gdrive = service_drive(self.authorization)
        #gsheet = service_sheet(self.authorization,'1hC8iT7IutoYDVDLlEWF8_op2viNRsUdv8tTVo9RlPkE')
        layer_a = QgsVectorLayer(os.path.join(self.plugin_dir,'test','dataset','c0601016_SistemiEcorelazionali.shp'), "layer_a", 'ogr')
        layer_b = QgsVectorLayer(os.path.join(self.plugin_dir,'test','dataset','c0601037_SpecieArboree.shp'), "layer_b", 'ogr')
        layer_c = QgsVectorLayer(os.path.join(self.plugin_dir,'test','dataset','c0509028_LocSitiContaminati.shp'), "layer_c", 'ogr')
        lv = plugins['layerVersion']
        for layer in  (layer_a, layer_b, layer_c ):
            print ("LAYER", layer.name())
            glayer = GoogleDriveLayer(self, self.authorization, layer.name(), importing_layer=layer, test=True)
            gsheet = glayer.get_service_sheet()

            print ("T1", gsheet.cell('Shape_Area',25))
            print ("T2", gsheet.set_cell('Shape_Area',24,234.500))
            print ("T3", gsheet.set_cell('Shape_Area',23,1000))
            print ("T4", gsheet.set_cell('Shape_Leng',22,'CIAOOOOO!'))
            print ("T5", gsheet.set_cell('Shape_Leng',21,None))
            print ("T6", gsheet.cell('Shape_Area',23))
            print ("T6", gsheet.cell('Shape_Leng',24))
            gsheet.add_sheet('byebye')
            gsheet.set_sheet_cell('byebye!A1', 'ciao')
            print ("FORMULA =SUM(SHEET!F2:F30):",gsheet.evaluate_formula('=SUM(SHEET!F2:F30)')))
            print ("FORMULA =MAX(SHEET!C2:C):",gsheet.evaluate_formula('=MAX(SHEET!C2:C)')))
            # gsheet.set_cell('barabao',33, 'ciao')
            fid = gsheet.new_fid()
            print ("NEW FID", fid)
            update_fieds = list(set(gsheet.header) - set(['WKTGEOMETRY','STATUS']))
            print ("update_fieds", update_fieds)
            print ("UPDATE DICT",dict(zip(update_fieds,["UNO",fid,34234,665.345,455.78,"HH"])))
            print ("APPEND_ROW", gsheet.add_row(dict(zip(update_fieds,['10000',"UNO",fid,34234,665.345,455.78,"HH"]))))
            print ("APPEND_COLUMN", gsheet.add_column(["UNO",fid,34234,665.345,455.78,"HH"]))
            print ("CRS", gsheet.crs())
            print ("NEW_FID", gsheet.new_fid())
            print ("DELETED FIELD 5", gsheet.mark_field_as_deleted(5))
            print (glayer.service_drive.trash_spreadsheet(glayer.get_gdrive_id()))
        print ("TEST ENDED")

    def load_available_sheets(self):
        bak_available_list_filepath = os.path.join(self.plugin_dir,'credentials','available_sheets.json')
        if os.path.exists(bak_available_list_filepath):
            with open(bak_available_list_filepath) as available_file:
                self.available_sheets = json.load(available_file)
        else:
            self.refresh_available()

    def refresh_available(self):
        available_list_filepath = os.path.join(self.plugin_dir,'credentials','available_sheets.json')
        self.available_sheets = self.myDrive.list_files()
        with io.open(available_list_filepath, 'w', encoding='utf-8') as available_file:
            available_file.write(unicode(json.dumps(self.available_sheets, ensure_ascii=False)))
        self.dlg.listWidget.clear()
        self.dlg.listWidget.addItems(self.available_sheets.keys())

    def run(self):
        """Run method that performs all the real work"""
        # show the dialog
        
        self.myDrive = service_drive(self.authorization)
        self.refresh_available()
        '''
        self.available_sheets = self.myDrive.list_files()
        self.available_sheets = {
            'TEST1 DELIMITAZIONI': {'id':'1aSI0qrC_mDrkffWK-crxtHugX5TqkjeASv_paovmKpA'},
            'TEST2 PUA': {'id':'1zwXxp6xMMzSYgbgFpryooF2PY5KWmnsXc0muCESXhQA'},
            'TEST3 strutVendita': {'id':'1TenhNLcCOJzunqLF2kvlJ3lCdxkJvSyhYqEnWc6pS28'},
            'TEST2 infra': {'id':'1AnGkUVXzNWt0k9850QvkDjtwoKdrntxe2t1aU7_G4tM'},
        }
        '''
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
        sheet_id = self.available_sheets[sheet_name]
        print (sheet_id)
        self.gdrive_layer = GoogleDriveLayer(self, self.authorization, sheet_name, spreadsheet_id=sheet_id)

    def dup_to_google_drive(self):
        currentLayer = self.iface.legendInterface().currentLayer()
        print (currentLayer.name())
        self.gdrive_layer = GoogleDriveLayer(self, self.authorization, currentLayer.name(), importing_layer=currentLayer)
        #update available list without refreshing
        try:
            self.available_sheets[currentLayer.name()] = self.gdrive_layer.spreadsheet_id
            self.dlg.listWidget.clear()
            self.dlg.listWidget.addItems(self.available_sheets.keys())
        except:
            pass