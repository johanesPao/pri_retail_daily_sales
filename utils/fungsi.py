import os
import math
from typing import Literal
from numpy import isnan
import pandas as pd
from O365 import Account, FileSystemTokenBackend
from datetime import date, timedelta
from pathlib import Path
from utils.excel import excel_rdsr

from utils.kueri import Kueri
from utils.pg import PostgreSQL


def persen_ach(numerator: float, denominator: float) -> float:
    # jika denominator 0 atau NaN
    if denominator == 0.0 or math.isnan(denominator):
        return 0.0
    else:
        return numerator / denominator


def autentikasi_o365(
    o365_env: dict[str, str]
) -> Account:
    print("Melakukan autentikasi tenant O365...")
    try:
        token_eksis = os.path.isfile(o365_env["TOKEN_FILE"])
        credentials = (o365_env["CLIENT_ID"], o365_env["CLIENT_SECRET"])
        token_backend = FileSystemTokenBackend(token_filename=o365_env["TOKEN_FILE"]) if token_eksis else None
        account = Account(
            credentials, scopes=["basic", "message_all"], token_backend=token_backend
        )
        if not token_eksis:
            # metode ini akan memunculkan prompt cli untuk autentikasi aplikasi
            # cari cara untuk mengautomatisasi proses auth ini
            account.authenticate()
        return account
    except:
        raise Exception("Gagal melakukan autentikasi dan mengembalikan Account tenant O365")


def kirim_email(env: dict[str, dict[str, str | None]]) -> None:
    print("Memulai pengiriman email...")
    try:
        account = autentikasi_o365(env["O365"])
        mail = account.new_message()
        mail.to.add("johanes.pao@pri.co.id")
        mail.subject = "Testing!"
        mail.body = "Halo halo, ini cuma testing"
        mail.send()
    except:
        raise Exception("Gagal melakukan pengiriman email")


def generate_df_utama(
    postgre_env: dict[str, str],
    parameter_kueri: dict[str, any],
    tgl_laporan: date = date.today() - timedelta(days=1),
) -> pd.DataFrame:
    print("Men-generate DataFrame utama untuk Retail Daily Sales Report...")
    # inisiasi kelas PostgreSQL
    sql = PostgreSQL(
        host=postgre_env["HOST"],
        port=postgre_env["PORT"],
        dbname=postgre_env["DB"],
        username=postgre_env["USER"],
        pwd=postgre_env["PWD"],
    )
    kueri = Kueri(parameter_kueri)
    # konstruksi dataframe
    # target
    df_target = sql.lakukan_kueri(kueri.kueri_target(tgl_laporan))
    # sales
    df_sales = sql.lakukan_kueri(kueri.kueri_sales(tgl_laporan))
    # join dataframe
    df_merge = df_target.merge(df_sales, left_on="Toko", right_on="Toko")
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
    print("Membentuk DataFrame SBU untuk Retail Daily Sales Report...")
    # kopi data
    df_sbu = data.copy()
    # create SBU categorization column
    df_sbu['SBU'] = df_sbu["Toko"].apply(
        lambda toko: 'Our Daily Dose' if toko.startswith("OD") else 
        'Fisik' if (
            toko.startswith("FS") | 
            toko.startswith("FF") | 
            toko.startswith("FO")
        ) else 'Bazaar')
    # drop kolom Toko dan Nama Toko
    df_sbu = df_sbu.drop(['Toko', 'Nama Toko'], axis=1)
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
    postgre_env: dict[str, str],
    parameter_kueri: dict[str, str],
    data: pd.DataFrame,
    tgl_laporan: date = date.today() - timedelta(days=1),
) -> pd.DataFrame:
    print("Men-generate DataFrame Area untuk Retail Daily Sales Report...")
    # inisiasi kelas PostgreSQL
    sql = PostgreSQL(
        host=postgre_env["HOST"],
        port=postgre_env["PORT"],
        dbname=postgre_env["DB"],
        username=postgre_env["USER"],
        pwd=postgre_env["PWD"],
    )
    # inisiasi kelas Kueri
    kueri = Kueri(parameter_kueri)
    # konstruksi dataframe
    # area_toko
    df_area = sql.lakukan_kueri(kueri.kueri_area_toko(tgl_laporan))
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
    print("Membentuk DataFrame Comp Non Comp Stores untuk Retail Daily Sales Report...")
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
            
    # kopi data
    df_cnc = data.copy()
    # kategorisasi comp-non comp berdasar ketersediaan data MTD LY Sales
    df_cnc['CNC'] = df_cnc.apply(
        lambda baris: kategorisasi_cnc(baris['Toko'], baris['MTD LY Sales']),
        axis=1
        )
    # drop kolom Toko dan Nama Toko
    df_cnc = df_cnc.drop(["Toko", "Nama Toko"], axis=1)
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
    print(f"Membentuk DataFrame sbu {sbu} untuk Retail Daily Sales Report...")
    # kopi data
    df_kopi = data.copy()
    match sbu:
        case "odd":
            # sbu main dataframe
            df_sbu = df_kopi.copy()[df_kopi["Toko"].str.startswith("OD")]
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
            df_fs = df_kopi.copy().loc[data["Toko"].str.startswith("FS")]
            df_ff = df_kopi.copy().loc[data["Toko"].str.startswith("FF")]
            df_fo = df_kopi.copy().loc[data["Toko"].str.startswith("FO")]
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
            df_sbu = df_kopi.copy().loc[data["Toko"].str.startswith("BZ")]
            # buat total jika row df lebih besar dari 0
            if df_sbu.shape[0] > 0:
                # sbu reset index to 1
                df_sbu.index = range(1, len(df_sbu) + 1)
                # append Comp Stores, Non Comp Stores dan Stores Total
                df_sbu.loc['COMP STORES TOTAL'] = df_sbu.copy().query('`MTD LY Sales`.notna() and Toko.notna()').sum(numeric_only=True, axis=0)
                df_sbu.loc['NON COMP STORES TOTAL'] = df_sbu.copy().query('`MTD LY Sales`.isna() and Toko.notna()').sum(numeric_only=True, axis=0)
                df_sbu.loc['STORES TOTAL'] = df_sbu.copy().query('Toko.notna()').sum(numeric_only=True, axis=0)
    # return sbu
    return df_sbu

def persentase_df_daily_retail_report(
    dataframe: pd.DataFrame,
    mode_df: Literal["sbu", "area", "cnc", "odd", "fisik", "bazaar"]
) -> pd.DataFrame:
    # kopi dataframe original
    df_kopi = dataframe.copy()
    # fix kolom
    kolom_fix = [
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
        "MTD LY Ach."
    ]
    # fungsi pengolahan persentase dataframe
    def hitung_target_ly_ach(dataframe: pd.DataFrame) -> pd.DataFrame:
        # fungsi perhitungan persentase
        def persen(pembilang: float, penyebut: float) -> float:
            if penyebut == 0.0:
                return 0.0
            else:
                return pembilang/penyebut
            
        # Target Achievement
        # Daily Target Ach.
        dataframe["Daily Ach."] = dataframe.apply(lambda baris: persen(baris["Daily Sales"], baris["Daily Target"]), axis=1)
        # WTD Target Ach.
        dataframe["WTD Ach."] = dataframe.apply(lambda baris: persen(baris["WTD Sales"], baris["WTD Target"]), axis=1)
        # MTD Target Ach.
        dataframe["MTD Ach."] = dataframe.apply(lambda baris: persen(baris["MTD Sales"], baris["MTD Target"]), axis=1)
        # LY Achievement
        # WTD LY Ach.
        dataframe["WTD LY Ach."] = dataframe.apply(lambda baris: persen(baris["WTD Sales"], baris["WTD LY Sales"]), axis=1)
        # MTD LY Ach.
        dataframe["MTD LY Ach."] = dataframe.apply(lambda baris: persen(baris["MTD Sales"], baris["MTD LY Sales"]), axis=1)
        # kembalikan dataframe
        return dataframe
    
    # match case mode_df
    match mode_df:
        case "sbu":
            # tambahkan kolom persentase
            df = hitung_target_ly_ach(df_kopi)
            # define kolom paling kiri pada mode_df
            kolom = ["SBU"]
            # extend dengan kolom_fix
            kolom.extend(kolom_fix)
            # reorder kolom
            df = df[kolom]
            # kembalikan dataframe
            return df
        case "area":
            # tambahkan kolom persentase
            df = hitung_target_ly_ach(df_kopi)
            # define kolom paling kiri pada mode_df
            kolom = ["Area"]
            # extend dengan kolom_fix
            kolom.extend(kolom_fix)
            # reorder kolom
            df = df[kolom]
            # kembalikan dataframe
            return df
        case "cnc":
            # tambahkan kolom persentase
            df = hitung_target_ly_ach(df_kopi)
            # define kolom paling kiri pada mode_df
            kolom = ["CNC"]
            # extend dengan kolom_fix
            kolom.extend(kolom_fix)
            # reorder kolom
            df = df[kolom]
            # kembalikan dataframe
            return df
        case "odd" | "fisik" | "bazaar":
            # tambahkan kolom persentase
            df = hitung_target_ly_ach(df_kopi)
            # define kolom paling kiri pada mode_df
            kolom = ["Toko", "Nama Toko"]
            # extend dengan kolom_fix
            kolom.extend(kolom_fix)
            # reorder kolom
            df = df[kolom]
            # kembalikan dataframe
            return df
        # raise TypeError jika mode_df tidak dikenali
        case _:
            raise TypeError(f"mode_df '{mode_df}' tidak dikenali")

def generate_daily_retail_report( 
    path_output: Path,
    tgl: date,
    env: dict[str, dict[str, str | None]],
    parameter_kueri: dict[str, any],
    nama_sheet: str | dict[str, str] = "Sheet1"
) -> None:
    # nama file
    nama_file = f"PRI_Retail_Daily_Sales_Report_{tgl.strftime("%d_%b_%Y")}.xlsx"
    # dataframe utama
    dataframe = generate_df_utama(env["POSTGRESQL"], parameter_kueri, tgl)
    # objek dataframe report
    objek_dataframe_report = {
        "sbu": persentase_df_daily_retail_report(generate_df_sbu(dataframe), "sbu"),
        "area": persentase_df_daily_retail_report(generate_df_area(env["POSTGRESQL"], parameter_kueri, dataframe, tgl), "area"),
        "cnc": persentase_df_daily_retail_report(generate_df_cnc(dataframe), "cnc"),
        "odd": persentase_df_daily_retail_report(generate_df_single_sbu(dataframe, "odd"), "odd"),
        "fisik": persentase_df_daily_retail_report(generate_df_single_sbu(dataframe, "fisik"), "fisik"),
        "bazaar": persentase_df_daily_retail_report(generate_df_single_sbu(dataframe, "bazaar"), "bazaar"),
    }
    # buat daily retail report
    excel_rdsr(path_output, nama_file, objek_dataframe_report, tgl, nama_sheet)
