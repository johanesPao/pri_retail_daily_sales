import os
from utils.pg import PostgreSQL
from dotenv import load_dotenv

# load environment
load_dotenv(".config.env")

# assign konfigurasi
konfigurasi = {
    "HOST": os.getenv("HOST"),
    "PORT": os.getenv("PORT"),
    "DBNAME": os.getenv("DBNAME"),
    "USER": os.getenv("USER"),
    "PASSWORD": os.getenv("PASSWORD"),
}

# inisiasi kelas PostgreSQL
sql = PostgreSQL(
    host=konfigurasi["HOST"],
    port=konfigurasi["PORT"],
    dbname=konfigurasi["DBNAME"],
    username=konfigurasi["USER"],
    pwd=konfigurasi["PASSWORD"],
)

print(
    sql.lakukan_kueri(
        f'select distinct "Global Dimension 1 Code" from "pnt live$value entry$437dbf0e-84ff-417a-965d-ed2bb9650972"'
    )
)
