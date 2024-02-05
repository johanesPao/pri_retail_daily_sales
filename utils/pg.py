import psycopg2
import pandas.io.sql as sqlio
from .kesalahan import tangkap_kesalahan_koneksi, tangkap_kesalahan_kueri


class PostgreSQL:
    def __init__(self, host: str, port: str, dbname: str, username: str, pwd: str):
        self.host = host
        self.port = port
        self.dbname = dbname
        self.username = username
        self.pwd = pwd

    def buka_koneksi(self) -> psycopg2.connect:
        try:
            return psycopg2.connect(
                "host='{}' port={} dbname='{}' user={} password={}".format(
                    self.host, self.port, self.dbname, self.username, self.pwd
                )
            )
        except Exception as e:
            tangkap_kesalahan_koneksi(e)

    def lakukan_kueri(self, kueri):
        try:
            return sqlio.read_sql_query(kueri, self.buka_koneksi())
        except Exception as e:
            tangkap_kesalahan_kueri(e)
