import os
from utils.pg import PostgreSQL
from dotenv import load_dotenv
from datetime import date, timedelta
from utils.kueri import kueri_target, kueri_sales

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

# konstruksi tanggal laporan
hari_ini = date.today()
kemarin = hari_ini - timedelta(days=1)

# inisiasi kelas PostgreSQL
sql = PostgreSQL(
    host=konfigurasi["HOST"],
    port=konfigurasi["PORT"],
    dbname=konfigurasi["DBNAME"],
    username=konfigurasi["USER"],
    pwd=konfigurasi["PASSWORD"],
)

# konstruksi dataframe
df_target = sql.lakukan_kueri(kueri_target(kemarin))
df_sales = sql.lakukan_kueri(kueri_sales(kemarin))
print(df_target)
print(df_sales)
