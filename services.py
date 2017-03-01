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
from PyQt5.QtCore import QSettings

#QGIS specific
import qgis.core

#Standard modules
import httplib2
import os
import sys
#import StringIO
import csv
import collections
import uuid
import base64
import zlib
from string import ascii_uppercase

#Google API
from apiclient import discovery
from apiclient.http import MediaFileUpload, MediaIoBaseUpload
from oauth2client import client, GOOGLE_TOKEN_URI
from oauth2client import tools
from oauth2client.file import Storage
from socks import socks

#Plugin modules
from .utils import slugify


logger = lambda msg: qgis.core.QgsMessageLog.logMessage(msg, 'Googe Drive Provider', 1)

def int_to_a1(n):
    if n < 1:
        return ''
    if n < 27:
        return ascii_uppercase[n-1]
    else:
        q, r = divmod(n, 26)
        return int_to_a1(q) + ascii_uppercase[r-1]

class google_authorization:

    def __init__(self, scopes, credential_dir, application_name, client_id, client_secret_file = 'client_secret.json' ):
        print ("authorizing:",client_id)
        self.credential_dir = os.path.abspath(credential_dir)
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        self.credential_path = os.path.join(credential_dir,client_id.split("@")[0]+"_"+slugify(application_name)+'.json')
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
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self.secret_path, self.scopes)
            print ("FLOW", flow)
            flow.user_agent = self.application_name
            if self.flags:
                credentials = tools.run_flow(flow, self.store, self.flags)
            else: # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, self.store)
            logger( 'Storing credentials to ' + self.credential_path)
        return credentials

    def authorize(self):
        s = QSettings()
        proxyEnabled = s.value("proxy/proxyEnabled", "")
        proxyType = s.value("proxy/proxyType", "" )
        proxyHost = s.value("proxy/proxyHost", "" )
        proxyPort = s.value("proxy/proxyPort", "" )
        proxyUser = s.value("proxy/proxyUser", "" )
        proxyPassword = s.value("proxy/proxyPassword", "" )
        self.httpConnection = httplib2.Http(ca_certs=os.path.join(self.credential_dir,'cacerts.txt'))
        if proxyEnabled == "true" and proxyType == 'HttpProxy': # test if there are proxy settings
            socks.setdefaultproxy(socks.PROXY_TYPE_HTTP, proxyHost, int(proxyPort), username = proxyUser, password = proxyPassword)
            socks.wrapmodule(httplib2)
            #proxyConf = httplib2.ProxyInfo(socks.PROXY_TYPE_HTTP_NO_TUNNEL, proxyHost, int(proxyPort), proxy_user = proxyUser, proxy_pass = proxyPassword)
        #else:
            #proxyConf =  None
        return self.get_credentials().authorize(self.httpConnection)



class service_drive:

    #last_query = {'id':None,'type':None ,'result':None}

    def __init__(self,credentials):
        self.credentials = credentials
        self.get_service()

    def get_service(self):
        self.service = discovery.build('drive', 'v3', http=self.credentials.authorize())

    def getFileMetadata(self, fileId, cacheQuery = True):
        if cacheQuery and hasattr(self, 'lastQuery') and self.lastQuery['id'] == fileId and self.lastQuery['type'] == 'getFileInfo':
            return self.lastQuery['metadata']
        else:
            metadata = self.service.files().get(fileId=fileId,
                                                fields="name, mimeType, id, description, shared, trashed").execute()
            self.lastQuery = {
                'id': fileId,
                'type': 'getFileInfo',
                'metadata': metadata
            }
            return metadata

    def isFileShared(self,fileId):
        return 'shared' in self.getFileMetadata(fileId).keys() and self.getFileMetadata(fileId)['shared']

    def isFileTrashed(self,fileId):
        return 'trashed' in self.getFileMetadata(fileId).keys() and self.getFileMetadata(fileId)['trashed']

    def isGooGisSheet(self,fileId):
        return 'description' in self.getFileMetadata(fileId).keys() and 'GOOGIS' in self.getFileMetadata(fileId)['description'].upper()

    def list_files(self, mimeTypeFilter = 'application/vnd.google-apps.spreadsheet', shared = None):
        query = 'trashed = false and appProperties has { key="isGOOGISsheet" and value="OK" }'
        query = "mimeType = '%s' and trashed = false and appProperties has { key='isGOOGISsheet' and value='OK' }" % mimeTypeFilter
        raw_list = self.service.files().list(orderBy="modifiedTime desc", q=query).execute()
        print ("list_files", raw_list)
        clean_dict = collections.OrderedDict()
        for item in raw_list['files']:
            if list(clean_dict.keys()).count(item['name']) > 0:
                key = "%s (%s)" % (item['name'], str(list(clean_dict.keys()).count(item['name'])))
            else:
                key = item['name']
            clean_dict[key] = item["id"]
        return clean_dict


    def mark_as_GooGIS_sheet(self,fileId):
        update_body = {
          "appProperties": {
            "isGOOGISsheet": "OK"
          }
        }
        result = self.service.files().update(fileId=fileId, body=update_body).execute()
        return result

    def download_file(self,fileId):
        media_obj = self.service.files().export(fileId=fileId, mimeType='text/csv').execute()
        return media_obj

    def ex_download_sheet(self,fileId):
        csv_txt = self.download_file(fileId)
        csv_file = StringIO.StringIO(csv_txt)
        csv_obj = csv.reader(csv_file,delimiter=',', quotechar='"')
        return csv_obj
    
    def file_property(self, fileId, property):
        metadata = self.service.files().get(fileId=fileId,fields=property).execute()
        return metadata[property]

    def set_spreadsheet_property(self,fileId, property, value):
        update_body = {
            property: value
        }
        result = self.service.files().update(fileId=fileId, body=update_body).execute()
        return result


    def create_sheet(self, sheetName='GooGIS', data = None):
        """
        """

    def trash_spreadsheet(self, spreadsheetId):
        """
        """
        self.set_spreadsheet_property(spreadsheetId,'trashed',True)
        

    def create_googis_sheet_from_csv(self,csv_path):

        media_body = MediaFileUpload(csv_path, mimetype='text/csv', resumable=None)
        body = {
            'name': os.path.basename(csv_path),
            'description': 'GooGIS sheet',
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        file = self.service.files().create(body=body, media_body=media_body).execute()


class service_spreadsheet:

    def __init__(self, credentials, spreadsheetId = None, new_sheet_name=None, new_sheet_data=None):
        self.credentials = credentials
        self.get_service()
        self.drive = service_drive(credentials)
        if spreadsheetId:
            self.name = self.drive.file_property(spreadsheetId, 'name') #the name is used for the first child sheet
            self.spreadsheetId = spreadsheetId
        elif new_sheet_name and new_sheet_data:
            create_body ={
                "properties": {
                    "title": new_sheet_name,
                    "locale": "en"
                },
            "sheets": [
                {
                    "properties": {
                        "title": new_sheet_name,
                    },
                }
            ]
            }
            result = self.service.spreadsheets().create(body=create_body).execute()
            self.spreadsheetId = result["spreadsheetId"]
            self.name = new_sheet_name
            update_range = new_sheet_name+"!A1"
            update_body = {
                "range": update_range,
                "values": new_sheet_data,
            }
            result = self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheetId,
                                                                 range=update_range,
                                                                 body=update_body,
                                                                 valueInputOption='USER_ENTERED').execute()
        else:
            raise Exception("service_sheet error: no sheet parameters provided")
            return

        capabilities = self.drive.file_property(self.spreadsheetId,"capabilities")
        self.update_header()
        self.canEdit = capabilities['canEdit']
        if self.canEdit:
            self.add_sheet('settings', hidden=False)
            self.add_sheet('changes_log', hidden=False)
            self.subscription = self.subscribe()
            self.drive.set_spreadsheet_property(self.spreadsheetId, "description", 'GooGIS layer')
            self.drive.mark_as_GooGIS_sheet(self.spreadsheetId)
        else:
            self.changes_log_rows = self.get_line("COLUMNS",'A',sheet="changes_log")

    def get_service(self):
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
        self.service = discovery.build('sheets', 'v4', http=self.credentials.authorize(), discoveryServiceUrl=discoveryUrl)

    def spreadsheetId(self):
        return self.spreadsheetId

    def subscribe(self):
        if not self.credentials.client_id in self.get_sheets():
            subscription = self.add_sheet(self.credentials.client_id, hidden=False)
            print ("subscription",subscription)
            return subscription
        else:
            print ("error multiple session on the same sheet!")
            self.erase_cells(self.credentials.client_id)
            return self.get_sheets()[self.credentials.client_id]

    def unsubscribe(self):
        update_body ={
            "requests":{
                "deleteSheet":{
                    "sheetId": self.subscription,
                }
            }
        }
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=update_body).execute()
        print ('result',result)
        return result

    def advertise(self,changes):
        for sheet_name,sheet_id in self.get_sheets().items():
            if not sheet_name in (self.name,'settings',self.credentials.client_id):
                print ('advertise', sheet_name)
                append_body = {
                    "range": sheet_name+"!A:A",
                    "majorDimension":'COLUMNS',
                    "values": [changes]
                }
                result = self.service.spreadsheets().values().append(spreadsheetId=self.spreadsheetId,
                                                                     range=sheet_name+"!A:A",
                                                                     body=append_body,
                                                                     valueInputOption='USER_ENTERED').execute()


    def update_header(self):
        result = self.service.spreadsheets().values().batchGet(spreadsheetId=self.spreadsheetId, ranges='1:1').execute()
        self.header_map = {}
        self.header = []
        for i, value in enumerate(result['valueRanges'][0]['values'][0]):
            self.header_map[value] = int_to_a1(i+1)
            self.header.append(value)
    
    def get_sheets(self):
        result = {}
        metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        for sheet in  metadata['sheets']:
            result[sheet['properties']['title']] = sheet['properties']['sheetId']
        return result
    
    def cell(self,field,row):
        if field in self.header_map.keys():
            A1_coords = self.header_map[field]+str(row)
            return self.sheet_cell(A1_coords)
        
    def sheet_cell(self,A1_coords):
        result = self.service.spreadsheets().values().batchGet(spreadsheetId=self.spreadsheetId, ranges=A1_coords, valueRenderOption='UNFORMATTED_VALUE').execute()
        try:
            cell_value = result['valueRanges'][0]['values'][0][0]
        except:
            cell_value = '()'
        if cell_value == '()':
            cell_value = None
        return cell_value
        
    def set_cell(self, field, row, value):
        if field in self.header_map.keys():
            A1_coords = self.header_map[field]+str(row)
            result = self.set_sheet_cell(A1_coords,value)
            if row == 1: #if row 1 is header so update stored header list
                self.update_header()
            return result
        else:
            #raise Exception("field %s not found") % field
            pass

    def set_multicell(self, mods, lockBy = None):
        locked = None
        if lockBy:
            ranges = []
            for (field, row, value) in mods:
                ranges.append( self.header_map['STATUS'] + str(row))
            print (ranges)
            query = self.service.spreadsheets().values().batchGet(spreadsheetId=self.spreadsheetId, ranges=ranges).execute()
            for value in query['valueRanges']:
                if not value["values"][0][0] in ('', '()','D', lockBy):
                    locked = value["values"][0][0]
                    break
        if locked:
            print ("Multi cell update is locked by "+locked)
            return None

        update_body = {
            "valueInputOption": 'USER_ENTERED',
            "data": []
        }
        for (field, row, value) in mods:
            if field in self.header_map.keys():
                if not value or value == qgis.core.NULL:
                    cleared_value = "()"
                else:
                    cleared_value = value
                valueRange = {
                    "range": self.header_map[field] + str(row),
                    "values": [[cleared_value]]
                }
                update_body['data'].append(valueRange)
            else:
                continue
        result = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body=update_body).execute()
        return result

    def get_line(self, majorDimension, line, sheet = None):
        if not sheet:
            sheet = self.name
        ranges = "%s!%s:%s" % (sheet, line, line)
        print ("line ranges:",ranges)
        result = self.service.spreadsheets().values().batchGet(spreadsheetId=self.spreadsheetId,
                                                               ranges=ranges,
                                                               majorDimension=majorDimension,
                                                               valueRenderOption='UNFORMATTED_VALUE').execute()
        if not 'values' in result['valueRanges'][0]:
            return [] #if cells required are void return a void list
        line_values = []
        for value in result['valueRanges'][0]['values'][0]:
            if value == "()":
                line_values.append(None)
            else:
                line_values.append(value)
        return line_values

    def get_sheet_values(self, child_sheet = None):
        if not child_sheet:
            child_sheet = self.name
        ranges = child_sheet
        result = self.service.spreadsheets().values().batchGet(spreadsheetId=self.spreadsheetId,
                                                               ranges=ranges,
                                                               majorDimension="ROWS",
                                                               valueRenderOption='UNFORMATTED_VALUE').execute()
        if not 'values' in result['valueRanges'][0]:
            return [] #if cells required are void return a void list
        array_values = []
        for row in result['valueRanges'][0]['values']:
            line_values = []
            for value in row:
                if value == "()":
                    line_values.append(None)
                else:
                    line_values.append(value)
            array_values.append(line_values)
        return array_values


    def update_cells(self,a1_origin,values,dimension='ROWS'):
        update_body = {
            "valueInputOption": 'USER_ENTERED',
            "data": [{
                "range": a1_origin,
                "majorDimension": dimension,
                "values": [values],
            }]
        }
        return self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.spreadsheetId, body=update_body).execute()
        
    def set_sheet_cell(self,A1_coords, value):
        if not value or value == qgis.core.NULL:
            value = "()"
        body = {
            "range": A1_coords,
            "values": [[value,],],
        }

        return self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheetId,
                                                           range=A1_coords,
                                                           body=body,
                                                           valueInputOption='USER_ENTERED').execute()

    def set_crs(self,crs):
        self.set_sheet_cell("settings!A2",crs)

    def crs(self):
        return self.sheet_cell("settings!A2")

    def set_geom_type(self,crs):
        self.set_sheet_cell("settings!B2",crs)

    def geom_type(self):
        return self.sheet_cell("settings!B2")

    def set_style(self,xmlstyle):
        xmlstyle_zip =  base64.b64encode(zlib.compress(xmlstyle.encode("utf-8")))
        self.set_sheet_cell("settings!A3",xmlstyle_zip)

    def style(self):
        xmlstyle_zip = self.sheet_cell("settings!A3")
        return zlib.decompress(base64.b64decode(xmlstyle_zip))
    
    def evaluate_formula(self,formula): #SHEET stay for the table layer
        formula = formula.replace('SHEET',self.name)
        self.set_sheet_cell("settings!A1",formula)
        return self.sheet_cell("settings!A1")
    
    def new_fid(self):
        return self.evaluate_formula('=MAX(SHEET!C2:C)') +1


    def erase_cells(self,range):
        erase_body = {
            "ranges":[
                range,
            ]
        }
        return self.service.spreadsheets().values().batchClear(spreadsheetId=self.spreadsheetId, body=erase_body).execute()

    def mark_field_as_deleted(self,fieldPos):
        '''
        '''
        cleaned_header = []
        for field in self.header[2:]:
            if field[:8] != 'DELETED_':
                cleaned_header.append(field)
        self.set_cell(cleaned_header[fieldPos],1,"DELETED_"+cleaned_header[fieldPos])
        return cleaned_header[fieldPos]

    def add_row(self,values_dict,childSheet = None):
        '''
        '''
        #reorder value_imput - not all fields could be present
        values = []
        for field in self.header:
            if field in values_dict.keys():
                values.append(values_dict[field])
            else:
                values.append('')
    
        append_body = {
            "majorDimension": "ROWS",
            "values": [values]
        }
        if childSheet:
            range = childSheet+"!A:ZZZ"
        else:
            range = "A:ZZZ"
        return self.service.spreadsheets().values().append(spreadsheetId=self.spreadsheetId,
                                                           range=range,
                                                           body=append_body,
                                                           valueInputOption='USER_ENTERED').execute()

    def add_column(self,values,child_sheet = None,fill_with_null = None):
        if not child_sheet:
            child_sheet = self.name
        '''
        metadata = self.service.spreadsheets().get(spreadsheetId=self.sheetId).execute()
        width = None
        height = None
        for sheet in  metadata['sheets']:
            if sheet['properties']['title'] == child_sheet:
                width = sheet['properties']['gridProperties']['columnCount']
                height = sheet['properties']['gridProperties']['rowCount']
                break
        '''
        width = self.evaluate_formula('=COUNTA(%s!1:1)' % self.name)

        if fill_with_null:
            for k in range(0,self.evaluate_formula('=COUNT(%s!C:C)' % self.name)-1): #fill the column with null values
                values.append('()')
        append_body = {
            "majorDimension": "COLUMNS",
            "values": [values]
        }
        A1_new_col = int_to_a1(width+1)
        cell_range = "%s!%s1:%s" % (child_sheet,A1_new_col, A1_new_col)
        result = self.service.spreadsheets().values().append(spreadsheetId=self.spreadsheetId,
                                                             range=cell_range,
                                                             body=append_body,
                                                             valueInputOption='USER_ENTERED').execute()
        self.update_header()
        return result

    
    def add_sheet(self, title, hidden=False):
        #check if settings exists
        metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        check_sheet_exists = None
        for sheet in  metadata['sheets']:
            if sheet['properties']['title'] == title:
                check_sheet_exists = sheet['properties']['sheetId']
                break
        if check_sheet_exists:
            return check_sheet_exists
        update_body = {
            "requests": [{
                "addSheet": {
                    "properties": {
                        "title": title,
                        "hidden": hidden,
                    }
                },
            }]
        }
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=update_body).execute()
        return result['replies'][0]['addSheet']['properties']['sheetId']
        
