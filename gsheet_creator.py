
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
	import argparse
	flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
	flags = None

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = '/Users/mhand/Downloads/google_client.json'
APPLICATION_NAME = ''

class GSheetEditor():
	def __init__(self):
		self.service = self.setup_service()


	def get_credentials(self):
		"""Gets valid user credentials from storage.

		If nothing has been stored, or if the stored credentials are invalid,
		the OAuth2 flow is completed to obtain the new credentials.

		Returns:
			Credentials, the obtained credential.
		"""
		home_dir = os.path.expanduser('~')
		credential_dir = os.path.join(home_dir, '.credentials')
		if not os.path.exists(credential_dir):
			os.makedirs(credential_dir)
		credential_path = os.path.join(credential_dir,
									   'sheets.googleapis.com-python.json')

		store = Storage(credential_path)
		credentials = store.get()
		if not credentials or credentials.invalid:
			flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
			flow.user_agent = APPLICATION_NAME
			if flags:
				credentials = tools.run_flow(flow, store, flags)
			else: # Needed only for compatibility with Python 2.6
				credentials = tools.run(flow, store)
			print('Storing credentials to ' + credential_path)
		return credentials


	def setup_service(self):
		credentials = self.get_credentials()
		http = credentials.authorize(httplib2.Http())
		discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
						'version=v4')
		service = discovery.build('sheets', 'v4', http=http,
								  discoveryServiceUrl=discoveryUrl)
		return service


	def update_doc_title(self, sheet_id, doc_title):
		update = {
			'updateSpreadsheetProperties': {
				'properties': {
					'title': doc_title
				},
				'fields': 'title'
			}
		}
		return update


	def execute_batch_request(self, sheet_id, batch):
		body = {
			'requests': batch
		}
		response = self.service.spreadsheets().batchUpdate(spreadsheetId=sheet_id,
													   body=body).execute()
		return response


	def setup_row_data(self, list_of_values):
		cells = []
		for value in list_of_values:
			cells.append({
				"userEnteredValue": {'stringValue': str(value)},
				"userEnteredFormat": {'wrapStrategy': 'WRAP', 'verticalAlignment': 'MIDDLE'}
				})
		return {'values': cells}


	def add_sheet_data(self, sheet, sheet_data):
		# Matrix of lists. Fills data down.
		row_data = []
		for row_index, row in enumerate(sheet_data):
			add = {
			'startRow':row_index,
			'rowData': self.setup_row_data(row)
			}
			row_data.append(add)
		sheet['data'] = row_data


	def create_sheet(self, title, rows=20, columns=12, sheet_data=[]):
		sheet = {"properties": {
			  "title": title,
			  "gridProperties": {
				"rowCount": rows,
				"columnCount": columns
				}
			}
		}
		if len(sheet_data):
			self.add_sheet_data(sheet, sheet_data)
		return sheet


	def ranges_from_indexes(self, row_index_start, row_index_end=None, column_index_start=None, column_index_end=None, sheet_id=None):
		# End ranges are exclusive
		add = {'startRowIndex': row_index_start}
		if row_index_end:
			add['endRowIndex'] = row_index_end
		if column_index_start:
			add['startColumnIndex'] = column_index_start
		if column_index_end:
			add['endColumnIndex'] = column_index_end
		if sheet_id:
			add['sheetId']= sheet_id 
		return add


	def merge_sheet_ranges(self, ranges, merge_type="MERGE_ALL"):
		# MERGE_ROWS, MERGE_COLUMNS, MERGE_ALL
		merges= []
		# Ranges are exclusive for the end value
		for r in ranges:
			merges.append({"mergeCells":
				{'range': r,
				"mergeType": merge_type
				}
			})
		return merges


	def update_cell_background(self, update_range, background_color):
		# {"red": 0-1, "green": 0-1, "blue": 0-1, "alpha": 0-1}
		update = {"repeatCell": 
			{'range': update_range,
			'cell':{
				"userEnteredFormat":{
					"backgroundColor": background_color
					}
			},
			"fields": "userEnteredFormat(backgroundColor)"
			}
		}
		return update


	def update_column_width(self, sheet_id, column_index_start, column_index_end, width=100):
		update = {
		  "updateDimensionProperties": {
			"range": {
			  "sheetId": sheet_id,
			  "dimension": "COLUMNS",
			  "startIndex": column_index_start,
			  "endIndex": column_index_end
			},
			"properties": {
			  "pixelSize": width
			},
			"fields": "pixelSize"
		  }
		}
		return update

	
	def auto_resize_columns(self, sheet_id, column_index_start, column_index_end):
		update = {
		  "autoResizeDimensions": {
			"dimensions": {
			  "sheetId": sheet_id,
			  "dimension": "COLUMNS",
			  "startIndex": column_index_start,
			  "endIndex": column_index_end
				}
			}
		}
		return update


	def freeze_rows(self, sheet_id, frozen_number):
		update = {
		  "updateSheetProperties": {
			"properties": {
			  "sheetId": sheet_id,
			  "gridProperties": {
				"frozenRowCount": frozen_number
			  }
			},
			"fields": "gridProperties.frozenRowCount"
			}
		}
		return update


	def create_spreadsheet(self, title, sheet_title="", sheets=[]):
		if sheet_title:
			sheets.append(self.create_sheet(sheet_title))
		spreadsheet_body = {
			"properties": {
				"title": title
				},
			"sheets":sheets
			}
		request = self.service.spreadsheets().create(body=spreadsheet_body)
		response = request.execute()
		return response


if __name__ == '__main__':
	s = GSheetEditor()
	test_data = [["heelp", 2, 3, 4, 5, 6, 7, 8], ['What', 'is', 'the', 'deal', 'with', 'airlines'], ['This', 'is', 'the', 'third', 'row'], ['How', 'is', 'life', 'today']]
	sheet = s.create_sheet('NEW SHEET', rows=len(test_data), columns=len(test_data[0]), sheet_data=test_data)
	resp = s.create_spreadsheet('THIS IS A TEST', sheets=[sheet])
	print(resp)
	document_id = resp['spreadsheetId']
	sheet_id = resp['sheets'][0]['properties']['sheetId']
	merge_range = s.ranges_from_indexes(1, 3, 0, 1, sheet_id)
	
	updates = [s.update_cell_background(merge_range, {'red':0, 'blue':0, 'green':0})]
	updates += s.merge_sheet_ranges([merge_range])
	updates.append(s.freeze_rows(sheet_id, 1))
	updates.append(s.update_column_width(sheet_id, 4, 6, 200))
	s.execute_batch_request(document_id, updates)
