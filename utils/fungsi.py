import os
import math
from typing import Literal
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


def generate_df_utama(
    tgl_laporan: date = date.today() - timedelta(days=1),
) -> pd.DataFrame:
    # file konfigurasi environment
    file_konfig = ".config.env"
    direktori = Path(os.path.dirname(__file__)).parent
    file_path = os.path.join(direktori, file_konfig)
    # load konfigurasi database pada .config.env
    load_dotenv(file_path)
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
    df_target = sql.lakukan_kueri(kueri_target(tgl_laporan))
    # sales
    df_sales = sql.lakukan_kueri(kueri_sales(tgl_laporan))
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

def generate_nama_sbu(
    sbu: Literal["ODD", "Fisik", "Bazaar"],
    tgl: date
) -> str:
    match sbu:
        case "ODD":
            nama_sbu = "Our Daily Dose".upper()
        case "Fisik":
            nama_sbu = "Fisik".upper()
        case _:
            nama_sbu = "Bazaar".upper()
    return f"{nama_sbu} NATIONAL SALES {tgl.strftime("%d %B %Y").upper()}"

def generate_df_single_sbu(
    data: pd.DataFrame, 
    sbu: Literal["ODD", "Fisik", "Bazaar"]
) -> pd.DataFrame:
    match sbu:
        case "ODD":
            df_sbu = data[data["Toko"].str.startswith("OD")]
        case "Fisik":
            df_sbu = data[
                (data["Toko"].str.startswith("FS")) | 
                (data["Toko"].str.startswith("FF")) | 
                (data["Toko"].str.startswith("FO"))
            ]
        case _:
            df_sbu = data[data["Toko"].str.startswith("BZ")]
    df_sbu.index = range(1, len(df_sbu) + 1)
    print(df_sbu)
    return df_sbu

def generate_report(
    path_template: Path, 
    path_output: Path, 
    dataframe: pd.DataFrame, 
    tgl: date
) -> None:
    data = dict(
        title="PRI Daily Retail Report",
        df_utama=dataframe.reset_index(),
        judul_odd=generate_nama_sbu("ODD", tgl),
        df_odd=generate_df_single_sbu(dataframe, "ODD").reset_index(),
        judul_fisik=generate_nama_sbu("Fisik", tgl),
        df_fisik=generate_df_single_sbu(dataframe, "Fisik").reset_index(),
        judul_bazaar=generate_nama_sbu("Bazaar", tgl),
        df_bazaar=generate_df_single_sbu(dataframe, "Bazaar").reset_index() if 
            generate_df_single_sbu(dataframe, "Bazaar").shape[0] > 0 else
            ["Data tidak ditemukan"],
    )
    with xw.App(visible=False) as app:
        book = app.render_template(
            path_template / "report_template.xlsx",
            path_output / f"PRI_Retail_Daily_Sales_Report_{tgl.strftime("%d_%b_%Y")}.xlsx",
            **data,
        )
