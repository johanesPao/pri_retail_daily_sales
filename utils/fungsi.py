import os
import math
from typing import Literal
from numpy import isnan
import pandas as pd
from O365 import Account, FileSystemTokenBackend
from dotenv import load_dotenv
from datetime import date, timedelta
from pathlib import Path
from utils.excel import excel_daily_retail_report

from utils.kueri import kueri_area_toko, kueri_sales, kueri_target
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
    # df_merge["Daily Ach."] = df_merge.apply(
    #     lambda baris: persen_ach(baris["Daily Sales"], baris["Daily Target"]), axis=1
    # )
    # df_merge["WTD Ach."] = df_merge.apply(
    #     lambda baris: persen_ach(baris["WTD Sales"], baris["WTD Target"]), axis=1
    # )
    # df_merge["MTD Ach."] = df_merge.apply(
    #     lambda baris: persen_ach(baris["MTD Sales"], baris["MTD Target"]), axis=1
    # )
    # # LY Achievement
    # df_merge["WTD LY Ach."] = df_merge.apply(
    #     lambda baris: persen_ach(baris["WTD Sales"], baris["WTD LY Sales"]), axis=1
    # )
    # df_merge["MTD LY Ach."] = df_merge.apply(
    #     lambda baris: persen_ach(baris["MTD Sales"], baris["MTD LY Sales"]), axis=1
    # )
    # reorder dataframe
    df_merge = df_merge[
        [
            "Toko",
            "Nama Toko",
            "Daily Sales",
            "Daily Target",
            # "Daily Ach.",
            "WTD Sales",
            "WTD Target",
            # "WTD Ach.",
            "WTD LY Sales",
            # "WTD LY Ach.",
            "MTD Sales",
            "MTD Target",
            # "MTD Ach.",
            "MTD LY Sales",
            # "MTD LY Ach.",
        ]
    ]
    # return df_merge sebagai dataframe utama
    return df_merge

def generate_df_sbu(
        data: pd.DataFrame
) -> pd.DataFrame:
    # create SBU categorization column
    data['SBU'] = data["Toko"].apply(
        lambda toko: 'Our Daily Dose' if toko.startswith("OD") else 
        'Fisik' if (
            toko.startswith("FS") | 
            toko.startswith("FF") | 
            toko.startswith("FO")
        ) else 'Bazaar')
    # drop kolom Toko dan Nama Toko
    df_sbu = data.copy().drop(['Toko', 'Nama Toko'], axis=1)
    # pandas groupby SBU
    df_sbu = df_sbu.groupby(['SBU'], as_index=False).sum()
    # reset index menjadi 1
    df_sbu.index = range(1, len(df_sbu) + 1)
    # append Comp Stores, Non Comp Stores dan Stores total
    df_sbu.loc['COMP STORES TOTAL'] = data.copy().query('`MTD LY Sales`.notna() and Toko.notna()').sum(numeric_only=True, axis=0)
    df_sbu.loc['NON COMP STORES TOTAL'] = data.copy().query('`MTD LY Sales`.isna() and Toko.notna()').sum(numeric_only=True, axis=0)
    df_sbu.loc['STORES TOTAL'] = data.copy().sum(numeric_only=True, axis=0)
    # return dataframe sbu
    return df_sbu

def generate_df_area(
        data: pd.DataFrame,
        tgl_laporan: date = date.today() - timedelta(days=1),
) -> pd.DataFrame:
    # file konfigurasi environment
    file_konfig = ".config.env"
    direktori = Path(os.path.dirname(__file__)).parent
    file_path = os.path.join(direktori, file_konfig)
    # load konf database pada .config.env
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
    # area_toko
    df_area = sql.lakukan_kueri(kueri_area_toko(tgl_laporan))
    # join dataframe df_area dan df_utama
    df_merge = data.copy().merge(df_area, left_on="Toko", right_on="Toko")
    # drop kolom Toko dan Nama Toko
    df_area = df_merge.drop(['Toko', 'Nama Toko'], axis=1)
    # pandas groupby Area
    df_area = df_area.groupby(['Area'], as_index=False).sum()
    # reset index menjadi 1
    df_area.index = range(1, len(df_area) + 1)
    # append Comp Stores, Non Comp Stores, dan Stores total
    df_area.loc['COMP STORES TOTAL'] = data.copy().query('`MTD LY Sales`.notna() and Toko.notna()').sum(numeric_only=True, axis=0)
    df_area.loc['NON COMP STORES TOTAL'] = data.copy().query('`MTD LY Sales`.isna() and Toko.notna()').sum(numeric_only=True, axis=0)
    df_area.loc['STORES TOTAL'] = data.copy().sum(numeric_only=True, axis=0)
    # return df_area
    return df_area

def generate_df_cnc(
        data: pd.DataFrame
) -> pd.DataFrame:
    # fungsi kategorisasi cnc
    def kategorisasi_cnc(kode_toko: str, mtd_ly_sales: int) -> str:
        match kode_toko[:2]:
            case "OD":
                return "Comp Stores ODD" if not(isnan(mtd_ly_sales)) else "Non Comp Stores ODD"
            case "FS" | "FF" | "FO":
                return "Comp Stores Fisik" if not(isnan(mtd_ly_sales)) else "Non Comp Stores Fisik"
            case "BZ":
                return "Bazaar Stores"
            case _:
                return "Comp Stores Others" if not(isnan(mtd_ly_sales)) else "Non Comp Store Others"
            
    # kategorisasi comp-non comp berdasar ketersediaan data MTD LY Sales
    data['CNC'] = data.apply(
        lambda baris: kategorisasi_cnc(baris['Toko'], baris['MTD LY Sales']),
        axis=1
        )
    # drop kolom Toko dan Nama Toko
    df_cnc = data.copy().drop(["Toko", "Nama Toko"], axis=1)
    # groupby CNC
    df_cnc = df_cnc.groupby(['CNC'], as_index=False).sum()
    # reset index dari 1
    df_cnc.index = range(1, len(df_cnc) + 1)
    # append Comp Stores, Non Comp Stores, Stores total
    df_cnc.loc['COMP STORES TOTAL'] = data.copy().query('`MTD LY Sales`.notna() and Toko.notna()').sum(numeric_only=True, axis=0)
    df_cnc.loc['NON COMP STORES TOTAL'] = data.copy().query('`MTD LY Sales`.isna() and Toko.notna()').sum(numeric_only=True, axis=0)
    df_cnc.loc['STORES TOTAL'] = data.copy().sum(numeric_only=True, axis=0)
    # return df_cnc
    return df_cnc

def generate_nama_sbu(
    sbu: Literal["odd", "fisik", "bazaar"],
    tgl: date
) -> str:
    match sbu:
        case "odd":
            nama_sbu = "Our Daily Dose".upper()
        case "fisik":
            nama_sbu = "Fisik".upper()
        case "bazaar":
            nama_sbu = "Bazaar".upper()
        case _:
            nama_sbu = "Others".upper()
    return f"{nama_sbu} NATIONAL SALES {tgl.strftime("%d %B %Y").upper()}"

def generate_df_single_sbu(
    data: pd.DataFrame, 
    sbu: Literal["odd", "fisik", "bazaar"]
) -> pd.DataFrame:
    match sbu:
        case "odd":
            # sbu main dataframe
            df_sbu = data[data["Toko"].str.startswith("OD")].copy()
            # buat total jika row df lebih besar dari 0
            if df_sbu.shape[0] > 0:
                # sbu reset index to 1
                df_sbu.index = range(1, len(df_sbu) + 1)
                # append Comp Stores Total, Non Comp Stores dan Stores Total
                df_sbu.loc['COMP STORES TOTAL'] = df_sbu.query('`MTD LY Sales`.notna() and Toko.notna()').sum(numeric_only=True, axis=0)
                df_sbu.loc['NON COMP STORES TOTAL'] = df_sbu.query('`MTD LY Sales`.isna() and Toko.notna()').sum(numeric_only=True, axis=0)
                df_sbu.loc['STORES TOTAL'] = df_sbu.query('Toko.notna()').sum(numeric_only=True, axis=0)
        case "fisik":
            # buat dataframe per sub sbu
            df_fs = data.copy().loc[data["Toko"].str.startswith("FS")]
            df_ff = data.copy().loc[data["Toko"].str.startswith("FF")]
            df_fo = data.copy().loc[data["Toko"].str.startswith("FO")]
            # buat subtotal jika row df lebih besar dari 0
            if df_fs.shape[0] > 0:
                # sbu reset index to 1
                df_fs.index = range(1, len(df_fs) + 1)
                df_fs.loc['FISIK SPORT TOTAL'] = df_fs.sum(numeric_only=True, axis=0)
            if df_ff.shape[0] > 0:
                # sbu reset index to 1
                df_ff.index = range(1, len(df_ff) + 1)
                df_ff.loc['FISIK FOOTBALL TOTAL'] = df_ff.sum(numeric_only=True, axis=0)
            if df_fo.shape[0] > 0:
                # sbu reset index to 1
                df_fo.index = range(1, len(df_fo) + 1)
                df_fo.loc['FACTORY OUTLET TOTAL'] = df_fo.sum(numeric_only=True, axis=0)
            # concat ketiga df sub sbu
            df_sbu = pd.concat([df_fs, df_ff, df_fo], ignore_index=False)
            # append Comp Stores, Non Comp Stores dan Stores Total
            df_sbu.loc['COMP STORES TOTAL'] = df_sbu.query('`MTD LY Sales`.notna() and Toko.notna()').sum(numeric_only=True, axis=0)
            df_sbu.loc['NON COMP STORES TOTAL'] = df_sbu.query('`MTD LY Sales`.isna() and Toko.notna()').sum(numeric_only=True, axis=0)
            df_sbu.loc['STORES TOTAL'] = df_sbu.query('Toko.notna()').sum(numeric_only=True, axis=0)
        case _:
            df_sbu = data.copy().loc[data["Toko"].str.startswith("BZ")]
            # buat total jika row df lebih besar dari 0
            if df_sbu.shape[0] > 0:
                # sbu reset index to 1
                df_sbu.index = range(1, len(df_sbu) + 1)
                # append Comp Stores, Non Comp Stores dan Stores Total
                df_sbu.loc['COMP STORES TOTAL'] = df_sbu.query('`MTD LY Sales`.notna() and Toko.notna()').sum(numeric_only=True, axis=0)
                df_sbu.loc['NON COMP STORES TOTAL'] = df_sbu.query('`MTD LY Sales`.isna() and Toko.notna()').sum(numeric_only=True, axis=0)
                df_sbu.loc['STORES TOTAL'] = df_sbu.query('Toko.notna()').sum(numeric_only=True, axis=0)
    # return sbu
    return df_sbu

def generate_daily_retail_report( 
    path_output: Path,
    tgl: date,
    nama_sheet: str | list[str] = "Sheet1"
) -> None:
    # nama file
    nama_file = f"PRI_Retail_Daily_Sales_Report_{tgl.strftime("%d_%b_%Y")}.xlsx"
    # dataframe utama
    dataframe = generate_df_utama(tgl)
    # objek dataframe report
    objek_dataframe_report = {
        "sbu": generate_df_sbu(dataframe),
        "area": generate_df_area(dataframe, tgl),
        "cnc": generate_df_cnc(dataframe),
        "odd": generate_df_single_sbu(dataframe, "odd"),
        "fisik": generate_df_single_sbu(dataframe, "fisik"),
        "bazaar": generate_df_single_sbu(dataframe, "bazaar"),
    }
    # buat daily retail report
    excel_daily_retail_report(path_output, nama_file, objek_dataframe_report, nama_sheet)
