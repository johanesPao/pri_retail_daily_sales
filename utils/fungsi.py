import os
import math
import pandas as pd
import xlwings as xw
from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import date, timedelta
from pathlib import Path

from utils.kueri import kueri_sales, kueri_target
from utils.pg import PostgreSQL

file_konfig = ".config.env"
direktori = Path(os.path.dirname(__file__)).parent
file_path = os.path.join(direktori, file_konfig)

load_dotenv(file_path)
# O365 env
TOKEN_FILE = os.getenv("O365_TOKEN_FILE")
O365_CLIENT_ID = os.getenv("O365_CLIENT_ID")
O365_CLIENT_SECRET = os.getenv("O365_CLIENT_SECRET")
# PostgreSQL env
PG_HOST = os.getenv("HOST")
PG_PORT = os.getenv("PORT")
PG_DB = os.getenv("DBNAME")
PG_USER = os.getenv("USER")
PG_PWD = os.getenv("PASSWORD")


def persen_ach(numerator: float, denominator: float) -> float:
    # jika denominator 0 atau NaN
    if denominator == 0.0 or math.isnan(denominator):
        return 0.0
    else:
        return numerator / denominator


def autentikasi_o365() -> Account:
    token_eksis = os.path.isfile(TOKEN_FILE)
    credentials = (O365_CLIENT_ID, O365_CLIENT_SECRET)
    token_backend = (
        FileSystemTokenBackend(token_filename=TOKEN_FILE) if token_eksis else None
    )
    account = Account(
        credentials, scopes=["basic", "message_all"], token_backend=token_backend
    )
    if not token_eksis:
        # metode ini akan memunculkan prompt cli untuk autentikasi aplikasi
        # cari cara untuk mengautomatisasi proses auth ini
        account.authenticate()
    return account


def kirim_email() -> None:
    account = autentikasi_o365()
    mail = account.new_message()
    mail.to.add("johanes.pao@pri.co.id")
    mail.subject = "Testing!"
    mail.body = "Halo halo, ini cuma testing"
    mail.send()


def generate_df_utama() -> pd.DataFrame:
    # file konfigurasi environment
    file_konfig = ".config.env"
    direktori = Path(os.path.dirname(__file__)).parent
    file_path = os.path.join(direktori, file_konfig)
    # load konfigurasi database pada .config.env
    load_dotenv(file_path)
    # konstruksi tanggal laporan
    tanggal_laporan = date.today() - timedelta(days=1)
    # inisiasi kelas PostgreSQL
    sql = PostgreSQL(
        host=PG_HOST,
        port=PG_PORT,
        dbname=PG_DB,
        username=PG_USER,
        pwd=PG_PWD,
    )
    # konstruksi dataframe
    # target
    df_target = sql.lakukan_kueri(kueri_target(tanggal_laporan))
    # sales
    df_sales = sql.lakukan_kueri(kueri_sales(tanggal_laporan))
    # join dataframe
    df_merge = df_target.merge(df_sales, left_on="Toko", right_on="Toko")
    # Target Achievement
    df_merge["Daily Ach."] = df_merge.apply(
        lambda baris: persen_ach(baris["Daily Sales"], baris["Daily Target"]), axis=1
    )
    df_merge["WTD Ach."] = df_merge.apply(
        lambda baris: persen_ach(baris["WTD Sales"], baris["WTD Target"]), axis=1
    )
    df_merge["MTD Ach."] = df_merge.apply(
        lambda baris: persen_ach(baris["MTD Sales"], baris["MTD Target"]), axis=1
    )
    # LY Achievement
    df_merge["WTD LY Ach."] = df_merge.apply(
        lambda baris: persen_ach(baris["WTD Sales"], baris["WTD LY Sales"]), axis=1
    )
    df_merge["MTD LY Ach."] = df_merge.apply(
        lambda baris: persen_ach(baris["MTD Sales"], baris["MTD LY Sales"]), axis=1
    )
    # reorder dataframe
    df_merge = df_merge[
        [
            "Toko",
            "Nama Toko",
            "Daily Sales",
            "Daily Target",
            "Daily Ach.",
            "WTD Sales",
            "WTD Target",
            "WTD Ach.",
            "WTD LY Sales",
            "WTD LY Ach.",
            "MTD Sales",
            "MTD Target",
            "MTD Ach.",
            "MTD LY Sales",
            "MTD LY Ach.",
        ]
    ]
    # return df_merge seabgai dataframe utama
    return df_merge


def generate_report(
    path_template: Path, path_output: Path, dataframe: pd.DataFrame
) -> None:
    data = dict(title="PRI Daily Retail Report", df_utama=dataframe)
    with xw.App(visible=False) as app:
        book = app.render_template(
            path_template / "report_template.xlsx",
            path_output / "myreport.xlsx",
            **data
        )
