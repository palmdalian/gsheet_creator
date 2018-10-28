
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

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations']
CLIENT_SECRET_FILE = os.path.expanduser('~/gsheet_creator/google_client.json')
APPLICATION_NAME = ''
DOMAIN = 'example.com' # Used to share with members of an org

class GSheetEditor():
	def __init__(self):
		self.colors = self.get_colors()
		self.service = self.setup_service()
		self.drive_service = self.setup_drive_service()
		self.slides_service = self.setup_slides_service()


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
									   'googleapis.com-python.json')

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

	def setup_slides_service(self):
		credentials = self.get_credentials()
		http = credentials.authorize(httplib2.Http())
		discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
						'version=v1')
		service = discovery.build('slides', 'v1', http=http)
		return service

	def setup_drive_service(self):
		credentials = self.get_credentials()
		http = credentials.authorize(httplib2.Http())
		discoveryUrl = ('https://drive.googleapis.com/$discovery/rest?'
						'version=v3')
		service = discovery.build('drive', 'v3', http=http)
		return service

	def get_colors(self):
		colors ={}
		unmapped = {'Green': '(217,234,211)', 'Pink': '(234,209,220)', 'Purple': '(217,210,233)', 'Blue': '(207,226,243)', 'Green2': '(217,234,211)', 'Yellow': '(255,255,0)'}
		for name, color in unmapped.items():
			red, green, blue = color.strip('(').strip(')').split(',')
			colors[name] = {'red': float(red)/128, 'blue': float(blue)/128, 'green': float(green)/128}
		return colors


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
				"userEnteredFormat": {'wrapStrategy': 'WRAP', 'verticalAlignment': 'MIDDLE', "horizontalAlignment": "CENTER"}
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


	def update_row_height(self, sheet_id, row_index_start, row_index_end, height=25):
		update = {
		  "updateDimensionProperties": {
			"range": {
			  "sheetId": sheet_id,
			  "dimension": "ROWS",
			  "startIndex": row_index_start,
			  "endIndex": row_index_end
			},
			"properties": {
			  "pixelSize": height
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


	def hide_columns(self, sheet_id, column_index_start, column_index_end):
		update = {'updateDimensionProperties': {
			    "range": {
			      "sheetId": sheet_id,
			      "dimension": 'COLUMNS',
			      "startIndex": column_index_start,
			      "endIndex": column_index_end,
			    },
			    "properties": {
			      "hiddenByUser": True,
			    },
			    "fields": 'hiddenByUser'
			}}
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


	def get_all_spreadsheet_data(self, spreadsheet_id):
		request = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id, includeGridData=True)
		response = request.execute()
		return response


	def format_sheet_data(self, sheet):
		sheet_data = []
		for row in sheet['data']:
			if not 'rowData' in row:
				continue
			for row_data in row['rowData']:
				data = []
				if not 'values' in row_data:
					continue
				for cell in row_data['values']:
					if 'effectiveValue' in cell:
						data.append(cell['formattedValue'])
					else:
						data.append('')
				sheet_data.append(data)
		return sheet_data


	def get_formatted_sheets_data(self, spreadsheet_id):
		all_data = self.get_all_spreadsheet_data(spreadsheet_id)
		formatted = []
		sheets = all_data['sheets']
		for sheet in sheets:
			sheet_data = self.format_sheet_data(sheet)
			new_sheet = {}
			new_sheet['data'] = sheet_data
			new_sheet['title'] = sheet['properties']['title']
			new_sheet['sheet_id'] = sheet['properties']['sheetId']
			formatted.append(new_sheet)
		return formatted


	# def callback(self, request_id, response, exception):
	# 	if exception:
	# 	    # Handle error
	# 	    print (exception)
	# 	else:
	# 	    print ("Permission Id: %s" % response.get('id'))


	def share_with_domain(self, file_id):
		batch = self.drive_service.new_batch_http_request(callback=None)
		domain_permission = {
		    'type': 'domain',
		    'role': 'writer',
		    'domain': DOMAIN
		}
		batch.add(self.drive_service.permissions().create(
			    fileId=file_id,
			    body=domain_permission,
			    fields='id',
			))
		batch.execute()

	def move_to_folder(self, file_id, folder_id):
		# Retrieve the existing parents to remove
		file = self.drive_service.files().get(fileId=file_id,
		                                 fields='parents').execute();
		previous_parents = ",".join(file.get('parents'))
		# Move the file to the new folder
		file = self.drive_service.files().update(fileId=file_id,
		                                    addParents=folder_id,
		                                    removeParents=previous_parents,
		                                    fields='id, parents').execute()


	def execute_presentation_batch(self, presentation_id, batch):
		body = {
			'requests': batch
		}
		response = self.slides_service.presentations().batchUpdate(presentationId=presentation_id,
													   body=body).execute()
		return response


	def create_presentation(self, title):
		body = {'title': title,}
		presentation = self.slides_service.presentations().create(body=body).execute()
		return presentation.get('presentationId')

	def add_slide(self):
		update = {'createSlide': {'slideLayoutReference': {}}}
		return update

	def add_text_formatting(self, object_id, font_size):
		update = { "updateTextStyle": {
				    "objectId": object_id,
				    "style": {"fontSize": pt14},
				    "textRange": { "type": "ALL" },
				    "fields": "fontSize"}}
		return update


	# This currently doesn't work well
	# def add_text_to_slide(self, object_id, title, text):
	# 	title_id = object_id + "title"
	# 	text_id = object_id + "text"
	# 	pt350 = {
	# 		'magnitude': 350,
	# 		'unit': 'PT'
	# 		}
	# 	pt14 = {
	# 		'magnitude': 14,
	# 		'unit': 'PT'
	# 		}
	# 	update = [
	# 	{'createShape': {
	# 		'objectId': title_id,
	# 		'shapeType': 'TEXT_BOX',
	# 		'elementProperties': {
	# 			'pageObjectId': object_id,
	# 			'size': {
	# 				'height': pt350,
	# 				'width': pt350
	# 			},
	# 				'transform': {
	# 					'scaleX': 1,
	# 					'scaleY': 1,
	# 					'translateX': 1,
	# 					'translateY': 1,
	# 					'unit': 'PT'
	# 					}
	# 				}
	# 			}
	# 		},
	# 		{
	# 		'insertText': {
	# 		'objectId': title_id,
	# 		'insertionIndex': 0,
	# 		'text': title}},
	# 	{'createShape': {
	# 		'objectId': text_id,
	# 		'shapeType': 'TEXT_BOX',
	# 		'elementProperties': {
	# 			'pageObjectId': object_id,
	# 			'size': {
	# 				'height': pt350,
	# 				'width': pt350
	# 			},
	# 				'transform': {
	# 					'scaleX': 1,
	# 					'scaleY': 1,
	# 					'translateX': 350,
	# 					'translateY': 100,
	# 					'unit': 'PT'
	# 					}
	# 				}
	# 			}
	# 		},
	# 		{
	# 		'insertText': {
	# 			'objectId': text_id,
	# 			'insertionIndex': 0,
	# 			'text': text}}
	# 		]
	# 	return update


if __name__ == '__main__':
	s = GSheetEditor()
	# Create a test spreadsheet
	test_data = [["heelp", 2, 3, 4, 5, 6, 7, 8], ['What', 'is', 'the', 'deal', 'with', 'airlines'], ['This', 'is', 'the', 'third', 'row'], ['How', 'is', 'life', 'today']]
	sheet = s.create_sheet('NEW SHEET', rows=len(test_data), columns=len(test_data[0]), sheet_data=test_data)
	resp = s.create_spreadsheet('THIS IS A TEST', sheets=[sheet])
	print(resp)

	# Apply formatting to the doc
	document_id = resp['spreadsheetId']
	sheet_id = resp['sheets'][0]['properties']['sheetId']
	merge_range = s.ranges_from_indexes(1, 3, 0, 1, sheet_id)
	updates = [s.update_cell_background(merge_range, s.colors['Blue'])]
	updates += s.merge_sheet_ranges([merge_range])
	updates.append(s.freeze_rows(sheet_id, 1))
	updates.append(s.update_column_width(sheet_id, 4, 6, 200))
	updates.append(s.update_row_height(sheet_id, 0, 1, 150))
	updates.append(s.hide_columns(sheet_id, 3, 4))
	s.execute_batch_request(document_id, updates)

	# WIP: Create a presentation
	presentation = s.create_presentation('testing')
	updates = []
	for i in xrange(0,10):
		updates.append(s.add_slide())
	resp = s.execute_presentation_batch(presentation, updates)
	updates = []
	for i, slide in enumerate(resp['replies']):
		object_id = slide['createSlide']['objectId']
		# updates.append(s.add_text_to_slide(object_id, str(i), "THIS IS A TEST"))
	resp = s.execute_presentation_batch(presentation, updates)
	print(resp)