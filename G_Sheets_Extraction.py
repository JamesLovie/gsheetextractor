# James Lovie 2019
# Google Sheet Extraction Script

from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
import os
import time

class Gsheets_Extractor:

	# Constructor
	def __init__(self, scope, workbook_key):
            self.scope = scope
            self.workbook_key = workbook_key

	def google_api_connector(self):
            data_folder = os.path.join("API_Key_Google")

            GOOGLE_KEY_FILE = os.path.join(data_folder, "api-key.json")

            credentials = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_KEY_FILE, self.scope)
            gc = gspread.authorize(credentials)
            workbook = gc.open_by_key(self.workbook_key)
            return workbook

	def extract_data(self, workbook):
            start_page_number = 148
            data_final = pd.DataFrame()

            total_sheets = len(workbook.worksheets())
            print(f'There are: {total_sheets} sheets in the workbook.')

            sheet = workbook.get_worksheet(start_page_number)
            values = sheet.get_all_values()
            data = pd.DataFrame(values[1:],columns=values[0])
            data_final = pd.concat([data_final, data], ignore_index=True, sort=False)

            for x in range(start_page_number, total_sheets):
                # Google API limits: https://developers.google.com/sheets/api/limits
                time.sleep(2)
                sheet = workbook.get_worksheet(x)
                values = sheet.get_all_values()
                data = pd.DataFrame(values[1:],columns=values[0])
                # Append the nth record data to the master dataframe.
                data_final = pd.concat([data_final, data], ignore_index=True, sort=False)
                print(f'Completed sheet # {x}')

            self.total_sheets = total_sheets
            return data_final


	def save_data(self, dataframe, filename):
		dataframe.to_csv(filename, index=False)
		print(f'File has been sucessfully exported to {filename}')

def main():
    # Instantiating the class (for constants)
    extractor = Gsheets_Extractor('https://www.googleapis.com/auth/spreadsheets', 'xy')
    # Establish Google Sheets API connection.
    workbook = extractor.google_api_connector()

    # Call the function extract_data.
    data_final = extractor.extract_data(workbook)
    print(f'Have extracted {extractor.total_sheets} sheets sucessfully')

    # Call the function save_data.
    extractor.save_data(data_final, 'data_november_2019.csv')

if __name__ == '__main__':
    main()
