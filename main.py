from gmail_api import GmailApi
from sheets_api import SheetsApi
from threading import Timer
from datetime import datetime
from bs4 import BeautifulSoup
import requests
import logging
import logging.config
import json
import re
import random

def get_fortune():
    """
    Gets todays fortune by executing curl to myfortunecookie.co.uk site
    Parses request for fortune string output
    :return str. Fortune cookie string
    """
    rnd_num = random.randint(1,152)
    req = requests.get(f'http://www.myfortunecookie.co.uk/fortunes/{rnd_num}/')
    soup = BeautifulSoup(req.text, 'lxml')
    try:
        fortune = soup.find_all('div', {'class':'fortune'})[0]
    except IndexError:
        get_fortune()
    clean = re.compile('<.*?>')
    return re.sub(clean, '', str(fortune))


class MainApp:

    def __init__(self, start_time):
        """
        Initiates Gmail API class which is responsible for getting messages from specified email.
        Initiates Sheets API class which is responsible for writing values to spreadsheet cell.
        """
        self.gmail = GmailApi()
        self.sheets = SheetsApi()
        self.is_running = False
        self.start_time = start_time
        with open('logging.conf', 'r') as lc:
            log = json.load(lc)
            logging.config.dictConfig(log)
        self.log = logging.getLogger('MainApp')

    def __run(self):
        """
        Gets receipts from gmail in list. Then provides following receipt list to write_to_sheet method
        which writes values to correct cell.
        :return: nothing.
        """
        self.log.info('Starting to get receipts from Gmail API ..')
        self.is_running = False
        receipts = self.gmail.get_receipts()
        self.log.info('Got receipts!')
        self.log.info('Starting to write to Google Sheets spreadsheet .. ')
        self.sheets.write_to_sheet(receipts)
        self.log.info('Writting to spreadsheet completed!')
        balance = self.sheets.get_weekly_balance()
        fortune = get_fortune()
        text = f'Labas,\nSavaitės likutis: {balance}\n\nŠios dienos palinkėjimas: {fortune}'
        mes = self.gmail.create_message('maxima.test.api@gmail.com', 'lukas.stankovicius@gmail.com', 'FINANSAI: Likęs balansas savaitei', text)
        self.gmail.send_message('me', mes)
        self.start()
    
    def __time_to_seconds(self, time):
        """
        Converts specified time into seconds
        :param time. Specified time (e.g 12:00:01)
        :return int. Time in seconds
        """
        h, m, s = time.split(':')
        return int(h) * 3600 + int(m) * 60 + int(s)

    def __get_wait_time(self):
        """
        Gets required amount of time to wait before starting main applications sequence.
        :return int. Required amount of time to wait in seconds
        """
        current_time = datetime.now().strftime('%H:%M:%S')
        wait_time = self.__time_to_seconds(self.start_time) - self.__time_to_seconds(current_time)

        if wait_time < 0:
            wait_time = 24 * 3600 + wait_time

        return wait_time

    def start(self, wait=True):
        """
        Gets required wait time to start application by calling __get_wait_time() method.
        Initiates Timer() class with specified wait interval.
        :param wait. Boolean indicating whether wait for start time or not.
        :return nothing.
        """
        self.log.info('Getting required amount of wait time ..')
        wait_time = self.__get_wait_time()
        self.log.debug(f'Wait time in seconds - {wait_time}')
        if wait is False:
            self.__run()
        if not self.is_running:
            self.__timer = Timer(wait_time, self.__run)
            self.__timer.start()
            self.is_running = True
            self.log.info(f'Waiting for {self.start_time}')


if __name__ == '__main__':
    m = MainApp("22:25:00")
    m.start(False)
