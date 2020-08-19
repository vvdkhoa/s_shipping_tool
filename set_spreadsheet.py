import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from time import sleep

##################################################
# https://github.com/burnash/gspread
# https://console.developers.google.com/cloud-resource-manager?pli=1
# Need share sheet for: ...
# For o_o spreadsheet
def set_spreadsheet(SPREADSHEET_KEY, SHEET_NAME):

	try:
		JSON_FILE = 'data_base_eBaygetPrices2.json'
		scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
		credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_FILE, scope)
		gc = gspread.authorize(credentials)
		spreadsheet = gc.open_by_key(SPREADSHEET_KEY) #Open Spreadsheet
		sheet = spreadsheet.worksheet(SHEET_NAME) # Open sheet by name

		return sheet

	except:
		sleep_time = 60
		print('    set_spreadsheet error. Try again after {}s'.format(sleep_time))
		sleep(sleep_time)
		set_spreadsheet(SPREADSHEET_KEY, SHEET_NAME)


##################################################
# For o_o spreadsheet
def set_spreadsheet_o_o(SHEET_NAME):

	SPREADSHEET_KEY = '' # Spreadsheets ID
	return set_spreadsheet(SPREADSHEET_KEY, SHEET_NAME)

##################################################
# For 0. 差出表_作業指示書
def set_spreadsheet_ship_table(SHEET_NAME):

	SPREADSHEET_KEY = '' # Spreadsheets ID
	return set_spreadsheet(SPREADSHEET_KEY, SHEET_NAME)


##################################################
def sheet_write(SHEET_NAME, Start_row, Start_col, End_row, End_col, update_list, worksheet):

	try:
		# worksheet = set_spreadsheet_o_o(SHEET_NAME)
		cell_list = worksheet.range(Start_row, Start_col, End_row, End_col)

		if len(cell_list) == len(update_list):
			for i in range(len(update_list)):
				cell_list[i].value = update_list[i]
			worksheet.update_cells(cell_list)
		else:
			print('Please check: List length = {} and total cells = {} are not same'
				  .format(len(update_list), len(cell_list)))
	except:
		sleep_time = 60
		print('    sheet_write error. Try again after {}s'.format(sleep_time))
		sleep(sleep_time)
		sheet_write(SHEET_NAME, Start_row, Start_col, End_row, End_col, update_list, worksheet)


##################################################
# Write to sheet  (example A1:B2: A1>B1>A2>B2)
# (SHEET_NAME (str), Start_row, Start_col, End_row, End_col, update_list (list))
# For o_o spreadsheet
def sheet_write_o_o(SHEET_NAME, Start_row, Start_col, End_row, End_col, update_list):
	worksheet = set_spreadsheet_o_o(SHEET_NAME)
	sheet_write(SHEET_NAME, Start_row, Start_col, End_row, End_col, update_list, worksheet)


##################################################
# For eBay_getPrices_2 spreadsheet
def sheet_write_ship_table(SHEET_NAME, Start_row, Start_col, End_row, End_col, update_list):
	worksheet = set_spreadsheet_ship_table(SHEET_NAME)
	sheet_write(SHEET_NAME, Start_row, Start_col, End_row, End_col, update_list, worksheet)


##################################################
# Clean all except the fist row
def clean_all(SPREADSHEET_KEY, SHEET_NAME):

	worksheet = set_spreadsheet(SPREADSHEET_KEY, SHEET_NAME)
	while True:
		try:
			data = worksheet.get_all_values()
			cell_list = worksheet.range(2, 1, len(data), len(data[0]))
			for cell in cell_list:
				cell.value = ''
			worksheet.update_cells(cell_list)
			break
		except:
			print('clean_all error, sleep 60s and try again.')
			sleep(60)


##################################################
# Clean all except the fist row for eBay_getPrices_2
def clean_eBay_getPrice_2(SHEET_NAME):
	SPREADSHEET_KEY = ''
	clean_all(SPREADSHEET_KEY, SHEET_NAME)


##################################################
def main_set_spreadsheet():
	print('Access to Google Spreadsheets')

##################################################
if __name__ == '__main__':
	main_set_spreadsheet()
