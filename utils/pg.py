from urllib.parse import quote
from sqlalchemy import create_engine, Engine
import pandas as pd
from .kesalahan import tangkap_kesalahan_koneksi, tangkap_kesalahan_kueri


class PostgreSQL:
    def __init__(self, host: str, port: str, dbname: str, username: str, pwd: str):
        self.host = host
        self.port = port
        self.dbname = dbname
        self.username = username
        self.pwd = pwd

    def buka_koneksi(self) -> Engine:
        try:
            # parse dan encode password
            encode_password = quote(self.pwd, safe="")
            return create_engine(
                f"postgresql://{self.username}:{encode_password}@{self.host}:{self.port}/{self.dbname}"
            )
        except Exception as e:
            tangkap_kesalahan_koneksi(e)

    def lakukan_kueri(self, kueri):
        try:
            return pd.read_sql(sql=kueri, con=self.buka_koneksi())
        except Exception as e:
            tangkap_kesalahan_kueri(e)
