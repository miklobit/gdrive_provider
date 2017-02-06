# -*- coding: utf-8 -*-
"""
/***************************************************************************

                                 A QGIS plugin

                              -------------------
        begin                :
        git sha              :
        copyright            :
        email                :
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


#QT4 specific
from PyQt4.QtCore import QSettings

#QGIS specific

#Standard modules
from apiclient.http import MediaFileUpload, MediaIoBaseUpload
import httplib2
import os
import sys
import StringIO
import csv

#Google API
from apiclient import discovery
from apiclient.http import MediaFileUpload, MediaIoBaseUpload
from oauth2client import client, GOOGLE_TOKEN_URI
from oauth2client import tools
from oauth2client.file import Storage

#Plugin modules
from utils import slugify


class google_authorization:

    def __init__(self, scopes, credential_dir, application_name, client_id, client_secret_file = 'client_secret.json' ):
        self.credential_dir = os.path.abspath(credential_dir)
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        self.credential_path = os.path.join(credential_dir,slugify(application_name)+'.json')
        self.secret_path = os.path.join(self.credential_dir,client_secret_file)
        self.store = Storage(self.credential_path)
        self.scopes = scopes
        self.client_id = client_id
        self.application_name = application_name

        try:
            import argparse
            self.flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
        except ImportError:
            self.flags = None


    def get_credentials(self):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """

        credentials = self.store.get()
        print credentials
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self.secret_path, self.scopes)
            flow.user_agent = self.application_name
            if self.flags:
                credentials = tools.run_flow(flow, self.store, self.flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, self.store)
            print 'Storing credentials to ' + self.credential_path
        return credentials

    def authorize(self):
        s = QSettings()
        proxyEnabled = s.value("proxy/proxyEnabled", "")
        proxyType = s.value("proxy/proxyType", "" )
        proxyHost = s.value("proxy/proxyHost", "" )
        proxyPort = s.value("proxy/proxyPort", "" )
        proxyUser = s.value("proxy/proxyUser", "" )
        proxyPassword = s.value("proxy/proxyPassword", "" )
        if proxyEnabled == "true" and proxyType == 'HttpProxy': # test if there are proxy settings
            proxyConf = httplib2.ProxyInfo(httplib2.socks.PROXY_TYPE_HTTP, proxyHost, int(proxyPort), proxy_user = proxyUser, proxy_pass = proxyPassword)
        else:
            proxyConf =  None
        self.httpConnection = httplib2.Http(proxy_info = proxyConf, ca_certs=os.path.join(self.credential_dir,'cacerts.txt'))
        return self.get_credentials().authorize(self.httpConnection)



class service_drive:

    #last_query = {'id':None,'type':None ,'result':None}

    def __init__(self,credentials):
        self.service = discovery.build('drive', 'v3', http=credentials.authorize())

    def getFileMetadata(self, fileId, cacheQuery = True):
        if cacheQuery and hasattr(self, 'lastQuery') and self.lastQuery['id'] == fileId and self.lastQuery['type'] == 'getFileInfo':
            return self.lastQuery['metadata']
        else:
            metadata = self.service.files().get(fileId=fileId,
                                                fields="name, mimeType, id, description, shared").execute()
            self.lastQuery = {
                'id': fileId,
                'type': 'getFileInfo',
                'metadata': metadata
            }
            return metadata

    def list_files(self, mimeTypeFilter = 'application/vnd.google-apps.spreadsheet', shared = None):
        raw_list = self.service.files().list().execute()
        clean_dict = {}
        for item in raw_list['files']:
            if item['mimeType'] == mimeTypeFilter:
                id = item['id']
                if self.isGooGisSheet(id):
                    if shared and self.isFileShared(id):
                        clean_dict[item['name']] = self.getFileMetadata(id)
                    elif not shared:
                        clean_dict[item['name']] = self.getFileMetadata(id)
        return clean_dict

    def isFileShared(self,fileId):
        return 'shared' in self.getFileMetadata(fileId).keys() and getFileMetadata(fileId)['shared']

    def isGooGisSheet(self,fileId):
        return 'description' in self.getFileMetadata(fileId).keys() and 'GOOGIS' in self.getFileMetadata(fileId)['description'].upper()

    def download_file(self,fileId):
        media_obj = self.service.files().export(fileId=fileId, mimeType='text/csv').execute()
        return media_obj

    def download_sheet(self,fileId):
        media_obj = self.service.files().export(fileId=fileId, mimeType='text/csv').execute()
        csv_txt = self.download_file(fileId)
        print csv_txt
        csv_file = StringIO.StringIO(csv_txt)
        csv_obj = csv.reader(csv_file,delimiter=',', quotechar='"')
        print csv_obj
        return csv_obj

    def upload_sheet(self, sheetName='GooGIS', body = {}, csv_file_obj = None, csv_path = None, update_sheetId = None):
        """
        """
        body['mimeType'] = 'application/vnd.google-apps.spreadsheet'

        if csv_path or csv_file_obj:
            if csv_path:
                media_body = MediaFileUpload(csv_path, mimetype='text/csv', resumable=None)
            elif csv_file_obj:
                media_body = MediaIoBaseUpload(csv_file_obj, mimetype='text/csv', resumable=None)
            if update_sheetId:
                return self.service.files().update(fileId=update_sheetId, media_body=media_body).execute()
            else:
                body['description'] = 'GooGIS sheet'
                body['name'] = sheetName
                return self.service.files().create(body=body, media_body=media_body).execute()
        else:
            return None

    def create_googis_sheet_from_csv(self,csv_path):

        media_body = MediaFileUpload(csv_path, mimetype='text/csv', resumable=None)
        body = {
            'name': os.path.basename(csv_path),
            'description': 'GooGIS sheet',
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        file = self.service.files().create(body=body, media_body=media_body).execute()
        print "CREATE:",file


class service_sheet:

    def __init__(self, sheetId,credentials):
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
        self.service = discovery.build('sheets', 'v4', http=credentials.authorize(), discoveryServiceUrl=discoveryUrl)
        self.sheetId = sheetId
