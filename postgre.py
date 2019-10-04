import psycopg2 as pg2
import re
from configparser import ConfigParser
from datetime import datetime

class PostgreSQL:


    def __init__(self):

        self.params = self.__config()
        conn = pg2.connect(**self.params)
        cur = conn.cursor()
        query = '''
                CREATE TABLE IF NOT EXISTS payments (
                date date, 
                amount VARCHAR(20),
                items VARCHAR(500));
                '''
        cur.execute(query)
        cur.close()
        conn.commit()

    def __config(self, filename='database.ini', section='postgresql'):

        parser = ConfigParser()
        parser.read(filename)
        db = {}
        params = parser.items(section)
        for param in params:
            db[param[0]] = param[1]
        return db

    def __filter_data(self, data):

        date = datetime.today().strftime('%Y-%m-%d')                                                                                                                                               
        for receipt in data:
            if not receipt[0] == '2019-09-30':
                continue
            return receipt
            
    def insert_data(self, data):

        conn = pg2.connect(**self.params)
        cur = conn.cursor()
        values = self.__filter_data(data)
        date = values[0]
        amount = values[1].replace(',','.')
        for item in values[2]:
            m = re.match('^([^.])\D*', item).group(0).lower()
            query = f'''
            INSERT INTO payments(date, amount, items)
            VALUES(to_date('{date}', 'YYYY-MM-DD'), {amount}, '{m}');
            '''
            cur.execute(query)
        query = '''SELECT * FROM payments'''
        cur.execute(query)
        cur.fetchall()
