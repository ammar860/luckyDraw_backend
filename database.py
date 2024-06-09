from datetime import datetime

import pyodbc


class Database:
    def __init__(self, config):
        self.config = config

    def connect(self):
        try:
            connection_string = 'Driver={%s};Server=%s;Database=%s;UID=%s;PWD=%s' % (
                self.config['dbconfig']['driver'], self.config['dbconfig']['server'],
                self.config['dbconfig']['database'], self.config['dbconfig']['username'],
                self.config['dbconfig']['password']
            )
            conn = pyodbc.connect(connection_string)
            return conn
        except Exception as e:
            print(e)
            return None
