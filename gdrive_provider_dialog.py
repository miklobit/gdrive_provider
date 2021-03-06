# -*- coding: utf-8 -*-
"""
/***************************************************************************
                                 A QGIS plugin
 A plugin for using Google drive sheets as QGIS layer shared between concurrent users
 portions of code are from https://github.com/g-sherman/pseudo_csv_provider
                              -------------------
        begin                : 2015-03-13
        git sha              : $Format:%H$
        copyright            : (C)2017 Enrico Ferreguti (C)2015 by GeoApt LLC gsherman@geoapt.com
        email                : enricofer@gmail.com
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

import os
import webbrowser

from PyQt4 import QtGui, QtCore, uic
from PyQt4.QtGui import QListWidgetItem, QIcon
from PyQt4.QtCore import Qt
from qgis.gui import QgsMapLayerComboBox, QgsMapLayerProxyModel
from qgis.core import QgsMapLayer, QgsNetworkAccessManager

FORM_CLASS1, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'gdrive_provider_dialog_base.ui'))

FORM_CLASS2, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'settings.ui'))

FORM_CLASS3, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'comboDialog.ui'))

FORM_CLASS4, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'importFromID.ui'))

FORM_CLASS5, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'internalBrowser.ui'))

FORM_CLASS6, _ = uic.loadUiType(os.path.join(
    os.path.dirname(__file__), 'webMap.ui'))

try:
    from pydevd import *
except:
    None
    
class GoogleDriveProviderDialog(QtGui.QDialog, FORM_CLASS1):
    def __init__(self, parent=None):
        """Constructor."""
        super(GoogleDriveProviderDialog, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)


class accountDialog(QtGui.QDialog, FORM_CLASS2):
    def __init__(self, parent=None, account='', error=None):
        """Constructor."""
        super(accountDialog, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.buttonBox.accepted.connect(self.acceptedAction)
        self.buttonBox.rejected.connect(self.rejectedAction)
        self.gdriveAccount.setText(account)
        self.gdriveAccount.textEdited.connect(self.gdriveAccountChanging)
        self.labelTerms.setText('<a href="https://developers.google.com/terms/">Google API terms of service</a>')
        self.labelTerms.setTextFormat(QtCore.Qt.RichText)
        self.labelTerms.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        self.labelTerms.setOpenExternalLinks(True)


        s = QtCore.QSettings()
        TOSAgreement = self.client_id = s.value("GooGIS/term_of_service_agreement",  defaultValue =  None)
        if TOSAgreement and self.gdriveAccount.text()!="":
            self.licenceCheck.setChecked(True)
            self.buttonBox.setEnabled(True)
        else:
            self.licenceCheck.setChecked(False)
            self.buttonBox.setEnabled(False)
            
        self.licenceCheck.stateChanged.connect(self.licenceCheckAction)

        if error:
            self.label = 'Invalid Google Drive Account'
            self.gdriveAccount.selectAll()
        else:
            self.label = 'Google Drive Account'
        self.acceptedFlag = None

    def licenceCheckAction(self, state):
        
        if self.gdriveAccount.text()=="":
            self.buttonBox.setEnabled(False)
            return
        self.buttonBox.setEnabled(self.licenceCheck.isChecked())
        s = QtCore.QSettings()
        s.setValue("GooGIS/term_of_service_agreement",self.licenceCheck.isChecked())

    def gdriveAccountChanging(self,txt):
        self.licenceCheck.setChecked(False)
        self.buttonBox.setEnabled(False)
        s = QtCore.QSettings()
        s.setValue("GooGIS/term_of_service_agreement",self.licenceCheck.isChecked())


    def acceptedAction(self):
        self.result = self.gdriveAccount.text()
        self.close()
        self.acceptedFlag = True

    def rejectedAction(self):
        self.close()
        self.acceptedFlag = None

    @staticmethod
    def get_new_account(account, error=None):
        dialog = accountDialog(account=account, error=error)
        dialog.setWindowFlags(Qt.WindowSystemMenuHint | Qt.WindowTitleHint) 
        
        result = dialog.exec_()
        dialog.show()
        if dialog.acceptedFlag:
            return (dialog.result)
        else:
            return None


class comboDialog(QtGui.QDialog, FORM_CLASS3):
    def __init__(self,layerMap , parent=None, current=None):
        """Constructor."""
        super(comboDialog, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.comboBox.clear()
        print "current",current
        for layerName,layer in layerMap.iteritems():
            if layer.type() == QgsMapLayer.VectorLayer:
                self.comboBox.addItem(layer.name(),layer)
        if current:
            self.comboBox.setCurrentIndex(self.comboBox.findData(current))
        self.buttonBox.accepted.connect(self.acceptedAction)
        self.buttonBox.rejected.connect(self.rejectedAction)
        self.acceptedFlag = None

    def acceptedAction(self):
        if self.comboBox.currentText() != '':
            self.result = self.comboBox.itemData(self.comboBox.currentIndex())
            self.acceptedFlag = True
        else:
            self.acceptedFlag = None
        self.close()

    def rejectedAction(self):
        self.close()
        self.acceptedFlag = None

    @staticmethod
    def select(layerMap,current=None):
        dialog = comboDialog(layerMap,current=current)
        dialog.setWindowFlags(Qt.WindowSystemMenuHint | Qt.WindowTitleHint) 
        result = dialog.exec_()
        dialog.show()
        if dialog.acceptedFlag:
            return (dialog.result)
        else:
            return None


class importFromIdDialog(QtGui.QDialog, FORM_CLASS4):
    def __init__(self, parent=None, layer=''):
        """Constructor."""
        super(importFromIdDialog, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.buttonBox.accepted.connect(self.acceptedAction)
        self.buttonBox.rejected.connect(self.rejectedAction)
        self.acceptedFlag = None
          
        #Signal QLineEdit   
        self.lineEdit.textChanged.connect(self.ImportByIdChange)

    def ImportByIdChange(self,string):
        if string!="":
            self.buttonBox.setEnabled(True)
            return
        self.buttonBox.setEnabled(False)  
        return   
   
    def acceptedAction(self):
        if self.lineEdit.text() != '':
            self.result = self.lineEdit.text()
            self.acceptedFlag = True
        else:
            self.acceptedFlag = None
        self.close()

    def rejectedAction(self):
        self.close()
        self.acceptedFlag = None

    @staticmethod
    def getNewId():
        dialog = importFromIdDialog()
        dialog.setWindowFlags(Qt.WindowSystemMenuHint | Qt.WindowTitleHint) 
        result = dialog.exec_()
        dialog.show()
        if dialog.acceptedFlag:
            return (dialog.result)
        else:
            return None


class internalBrowser(QtGui.QDialog, FORM_CLASS5):
    def __init__(self, target = '', title = '', parent = None):
        """Constructor."""
        super(internalBrowser, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.setWindowTitle(title)
        self.webView.page().setNetworkAccessManager(QgsNetworkAccessManager.instance())
        self.webView.setUrl(QtCore.QUrl(target))


class webMapDialog(QtGui.QDialog, FORM_CLASS6):
    def __init__(self,availableMaps, parent=None):
        """Constructor."""
        super(webMapDialog, self).__init__(parent)
        # Set up the user interface from Designer.
        # After setupUI you can access any designer object by doing
        # self.<objectname>, and you can use autoconnect slots - see
        # http://qt-project.org/doc/qt-4.8/designer-using-a-ui-file.html
        # #widgets-and-dialogs-with-auto-connect
        self.setupUi(self)
        self.availableMaps = availableMaps
        self.buttonBox.accepted.connect(self.acceptedAction)
        self.buttonBox.rejected.connect(self.rejectedAction)
        for map_name, map_metadata in availableMaps.iteritems():
            permissions = self.get_permissions(map_metadata)
            if "anyone" in permissions:
                publicMapItem = QListWidgetItem(QIcon(os.path.join(os.path.dirname(__file__),'globe.png')),map_name,self.publicWebMapsList, QListWidgetItem.UserType)
                self.publicWebMapsList.addItem(publicMapItem)

    def get_permissions(self, metadata):
        '''
        returns a simplified list of permissions from the downloaded metadata
        :param metadata: the downloaded file metadata
        '''
        permissions = {}
        if 'permissions' in metadata:
            for permission in metadata['permissions']:
                if permission['type'] == 'anyone':
                    permissions['anyone'] = permission['role']
        return permissions

    def acceptedAction(self):
        selectedIds = []
        for map_item in self.publicWebMapsList.selectedItems():
            selectedIds.append(self.availableMaps[map_item.text()]['id'])
        url = "https://enricofer.github.io/GooGIS2CSV/converter.html?spreadsheet_id="
        for map_id in selectedIds:
            url += map_id+','
        url = url[:-1]
        webbrowser.open(url, new=2, autoraise=True)
        self.result = url
        self.close()
        self.acceptedFlag = True

    def rejectedAction(self):
        self.close()
        self.acceptedFlag = None

    @staticmethod
    def get_web_link(availableMaps):
        dialog = webMapDialog(availableMaps)
        dialog.setWindowFlags(Qt.WindowSystemMenuHint | Qt.WindowTitleHint)

        result = dialog.exec_()
        dialog.show()
        if dialog.acceptedFlag:
            return (dialog.result)
        else:
            return None

