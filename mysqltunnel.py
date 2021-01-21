import MySQLdb
import bgtunnel
import os
import sys
from dotenv import load_dotenv

load_dotenv('.env-local')

class TunnelSQL(object):
    conn = False
    cursor = False
    connected = False

    def __init__(self, silent):
        self.ssh_user=os.getenv('SSH_USER')
        self.ssh_address=os.getenv('SSH_HOST')
        self.ssh_port=os.getenv('SSH_PORT')
        self.host_port=os.getenv('SSH_HOST_PORT') or 3306
        self.bind_port=os.getenv('SSH_BIND_PORT') or 3308
        self.silent=silent

        self.user = os.getenv('DB_USER')
        self.passwd = os.getenv('DB_PASS')
        self.dbname = os.getenv('DB_NAME')
        self.port = os.getenv('DB_PORT') or 3308
        self.host = ('127.0.0.1', os.getenv('DB_HOST'))[self.ssh_address == None]

    def connect(self):
        self.forwarder = bgtunnel.open(
                         ssh_user=self.ssh_user,
                         ssh_address=self.ssh_address,
                         ssh_port=self.ssh_port,
                         host_port=self.host_port,
                         bind_port=self.bind_port,
                         silent=self.silent,
        )

        if not self.silent:
            print("Starting mysql with: " +
                  "MySQLdb.connect(host='{}',port='{}', user='{}', passwd='{}', db='{}')".format(
                  self.host, self.port, self.user, self.passwd, self.dbname)
            )
        self.conn = MySQLdb.connect(host=self.host, port=self.port,
                                    user=self.user, passwd=self.passwd,
                                    db=self.dbname)
        self.connected = True
        if not self.silent:
            print(("Connection Failed", "Connected")[self.connected])
            print("Cursor Created")
        self.cursor = self.conn.cursor()

    def close(self):
        if self.cursor:
            if not self.silent:
                print("Closing Cursor")
            self.cursor.close()
            self.cursor = False
        if self.conn:
            self.conn.close()
            if not self.silent:
                print("Closing Connection")
            self.conn = False
        if not self.silent:
            print("Closing Forwarder")
        self.forwarder.close()

    def query(self, sql, data=[]):
        if not self.conn:
            self.connect()
        self.conn.query(sql, data)
        return self.conn.use_result()

    def execute(self, sql, data=[]):
        if not self.silent:
            print("Executing: " + sql)
            print(data)
        if not self.conn:
            self.connect()
        self.cursor.execute(sql, data)
        return self.cursor.fetchall()

    def executemany(self, sql, data=[]):
        if not self.conn:
            self.connect()
        self.cursor.executemany(sql, data)
        self.conn.commit()

    def info(self):
       return self.conn.info()

    def insert_id(self):
        return self.conn.insert_id()

