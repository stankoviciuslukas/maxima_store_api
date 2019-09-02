from __future__ import print_function
import pickle
import os.path
import logging
import logging.config
import json
from datetime import datetime
from string import ascii_uppercase
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


class SheetsApi:
    """
    Initiates Google Sheets API which adds values and notes to specified cells.
    """

    def __init__(self):
        """
        Loads logging configuration file.
        Start Sheets API initiation process.
        """
        with open('logging.conf', 'r') as lc:
            log = json.load(lc)
            logging.config.dictConfig(log)
        self.log = logging.getLogger('SheetsApi')
        self.date = datetime.today().strftime('%Y-%m-%d')
        self.month = datetime.today().strftime('%Y-%m')
        self.__init_api()

    def __init_api(self):
        """
        Configures service API which is defined by variable api
        :param str api. Service name
        :return service. Instance of configured service (gmail or sheets)
        """
        api = 'sheets'
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.spreadsheet_id = '1eeWDcQ63Bsz6CU1PBcJ8_wKC4h5xpPhBHEpe1SOqyRg'
        version = 'v4'

        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(f'creds/token_{api}.pickle'):
            with open(f'creds/token_{api}.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    f'creds/credentials_{api}.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(f'creds/token_{api}.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build(api, version, credentials=creds)
        self.log.debug(f'Setup completed with {api}')

    def __loop_through(self, service):
        """
        By specified service goes to spreadsheet cell until finds a match
        of cell value and current date. Then that cell value is returned.
        :param service: Google Sheets API service instance
        :return: cell. Cell which matches with todays date.
        """
        date = datetime.today().strftime('%Y-%m-%d')
        month = datetime.today().strftime('%Y-%m')
        for number in [3, 10, 17, 24]:  # Week rows
            for letter in ascii_uppercase[7:14]:  # Range from H to N
                cell = f'{letter}{number}'
                result = service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                             range=f'{month}!{cell}').execute()
                date_in_spreadsheet = result.get('values', [])
                if date_in_spreadsheet[0][0] == date:
                    self.log.debug(f'Got cell {cell}')
                    return cell

    def __increase_cell_number(self, cell, number):
        """
        Increases sheet cell number by times which user provides
        :param cell. Sheet cell in string.
        :param number. Int number by what should cell be increased.
        :returns str. Increased cell number (e.g from J24 to J25)
        """
        cell = [i for i in cell]
        temp_number = int(''.join(cell[1:])) + int(number)
        return cell[0] + str(temp_number)

    def __get_cell_range(self, cell):
        """
        Gets cell range by provided cell number. Cell range is required 
        for sheets API to understand where user note should be placed.
        :param cell. Sheet cell in string.
        :return int. startRowIndex, endRowIndex, startColumnIndex, endColumnIndex
        """
        cell_splitted = [i for i in cell]
        for number, letter in enumerate(ascii_uppercase):
            if letter == cell_splitted[0]:
                startColumnIndex = number
                endColumnIndex = number + 1
                startRowIndex = int(''.join(cell_splitted[1:])) - 1
                endRowIndex = int(''.join(cell_splitted[1:]))
                self.log.debug(f'Got cell range: startRowIndex {startRowIndex}, endRowIndex {endRowIndex}')
                self.log.debug(f'startColumnIndex {startColumnIndex}, endColumnIndex {endColumnIndex}')
                return startRowIndex, endRowIndex, startColumnIndex, endColumnIndex
    
    def __update_sheet(self, body, c_range=None):
        """
        Updating spreadsheet cel y specified body param. 
        :param body. Dictionary with specifies sheets API whether to update values or add note to cell.
        :param c_range. String which specifies on which cell should values be updated.
        :return nothing.
        """
        if c_range is not None:
            self.log.debug(f'Entering receipt price to cell - {c_range}')
            self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheet_id,
                range=c_range, valueInputOption='USER_ENTERED', body=body).execute()
        else:
            self.log.debug('Entering note to cell')
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()

    def get_weekly_balance(self):
        """
        Gets weekly balance from spreadsheet. Finds week start date and week end date compares it with
        todays date. Day specifies which row to select for correct weekly balance.
        :return str. Weekly balance.
        """
        
        day = int(datetime.today().strftime('%d'))

        for row in [19, 20, 21, 22]:
            w_start = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                            range=f'{self.month}!B{row}').execute()
            w_end = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                            range=f'{self.month}!C{row}').execute()
            w_start = int(w_start.get('values')[0][0].split('-')[-1])
            w_end = int(w_end.get('values')[0][0].split('-')[-1])

            if w_start <= day <= w_end:
                cell = f'{self.month}!E{row}'
        
        balance = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheet_id,
                                                            range=cell).execute().get('values')
        return balance[0][0].replace(' ','')

    def write_to_sheet(self, receipts):
        """
        Writes parsed receipts from messages to sheets api.
        If receipt date do not match current date - receipt is skipped.
        First cost written in receipt is added to cell. Note with full
        purchase list is added to selected cell.
        :param receipts. Founded receipts in first 30 email messages
        :return nothing.
        """
        self.log.debug('Writting to sheet ..')
        cell_range = self.__loop_through(self.service)  # Find week
        cell_and_number = self.__increase_cell_number(cell_range, 2)  # Find start position
        month = datetime.today().strftime('%Y-%m')

        for receipt in receipts:
            cost = []
            if not receipt[0] == datetime.today().strftime('%Y-%m-%d'):
                continue
            for items in receipt[1:]:
                if isinstance(items, list):
                    startRowIndex, endRowIndex, startColumnIndex, endColumnIndex = self.__get_cell_range(cell_and_number)
                    body = {
                        "requests": [
                            {
                                "repeatCell": {
                                    "range": {
                                        "sheetId": 1502489095,  # this is the end bit of the url
                                        "startRowIndex": startRowIndex,
                                        "endRowIndex": endRowIndex,
                                        "startColumnIndex": startColumnIndex,
                                        "endColumnIndex": endColumnIndex,
                                    },
                                    "cell": {"note": '\n'.join(items)},
                                    "fields": "note",
                                }
                            }
                        ]
                    }
                    self.__update_sheet(body)
                else:
                    cost.append([items])
                    body = {
                       'values': cost
                    }
                    self.log.debug('Updating sheet .. ')
                    self.__update_sheet(body, f'{month}!{cell_and_number}')
            cell_and_number = self.__increase_cell_number(cell_and_number, 1)
