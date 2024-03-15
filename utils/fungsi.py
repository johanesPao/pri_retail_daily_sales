import locale
import os
import math
from typing import Literal
from numpy import isnan
import pandas as pd
from O365 import Account, FileSystemTokenBackend, MSGraphProtocol
from datetime import date, timedelta, datetime
from pytz import timezone
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
        # autentikasi account perlu dilakukan terlebih dahulu satu kali menggunakan metode "on behalf of user",
        # dikarenakan metode auh lainnya mungkin memerlukan setup pada organisasi.
        # seperti yang telah dijelaskan di bawah, jika tidak memiliki file dengan konten token o365, maka kita
        # perlu menjalankan satu kali fungsi
        # account = Account(credentials)
        # account.authenticate(["basic", "mailbox", "message_all"])
        # kedua baris di atas akan menghasilkan file .txt dengan konten scopes, token akses dan token refresh
        # seperti ini:
        # {
        #  "token_type": "Bearer",
        #  "scope": [
        #   "profile",
        #   "openid",
        #   "email",
        #   "https://graph.microsoft.com/Mail.Read",
        #   "https://graph.microsoft.com/Mail.ReadWrite",
        #   "https://graph.microsoft.com/Mail.Send",
        #   "https://graph.microsoft.com/User.Read"
        #  ],
        #  "expires_in": 5371,
        #  "ext_expires_in": 5371,
        #  "access_token": akses_token,
        #  "refresh_token": refresh_token,
        #  "expires_at": 1710223536.265685
        # }
        account = Account(
            credentials, scopes=["basic", "mailbox", "message_all"], token_backend=token_backend
        )
        if not token_eksis:
            # metode ini akan memunculkan prompt cli untuk autentikasi aplikasi
            # cari cara untuk mengautomatisasi proses auth ini
            account = Account(credentials)
            account.authenticate(["basic", "mailbox", "message_all"])
        return account
    except:
        raise Exception("Gagal melakukan autentikasi dan mengembalikan Account tenant O365")


def kirim_email(
    env: dict[str, dict[str, str | None]],
    file: list[str] | None,
    from_email: str | None,
    to_email: list[str] | None,
    cc_email: list[str] | None,
    judul_email: str | None,
    body_email: str | None
) -> None:
    try:
        account = autentikasi_o365(env["O365"])
        print("[email] Setting pengirim pesan...")
        mailbox = account.mailbox(from_email if from_email != None else 'noname@nowhere.id')
        pesan = mailbox.new_message()
        print("[email] Setting TO pesan...")
        pesan.to.add(to_email if to_email != None else ['johanes.pao@pri.co.id'])
        print("[email] Setting CC pesan...")
        pesan.cc.add(cc_email if cc_email != None else ['johanes.pao@pri.co.id'])
        print("[email] Setting subyek pesan...")
        pesan.subject = judul_email if judul_email != None else 'No Subject'
        print("[email] Setting body pesan...")
        pesan.body = body_email if body_email != None else '<p style="color:red; margin: 20px;">Tidak ada email body</p>'
        pesan.save_message()
        print("[email] Setting attachment pesan...")
        if file != None:
            for indeks, attachment in file:
                pesan.attachments.add(attachment)
                print(f"[email] {pesan.attachments[indeks]}")
        else:
            print("[email] Tidak ada file yang dilampirkan...")
        print("[email] Mengirimkan pesan.")
        pesan.send()
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
            # replace NaN dengan 0 pada kolom_fix
            df.update(df.copy()[kolom_fix].fillna(0))
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
            # replace NaN dengan 0 pada kolom_fix
            df.update(df.copy()[kolom_fix].fillna(0))
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
            # replace NaN dengan 0 pada kolom_fix
            df.update(df.copy()[kolom_fix].fillna(0))
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
            # replace NaN dengan 0 pada kolom_fix
            df.update(df.copy()[kolom_fix].fillna(0))
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
) -> dict[str, dict[str, pd.DataFrame]]:
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
    # untuk saat ini jika dalam CI, cukup print kunci dan length dari df
    # kita belum menemukan cara melakukan workbook.close() pada runner github action yang benar
    # if env["CI"]:
    #     for kunci, df in objek_dataframe_report.items():
    #         print(f"Panjang data {kunci}: {len(df)} baris")
    #     # print path_output untuk debugging pada proses output dalam CI
    #     print(f"path_output: {path_output}")
    # else:
    #     # buat daily retail report
    #     excel_rdsr(path_output, nama_file, objek_dataframe_report, tgl, nama_sheet)

    # release this part, don't compile on github action, just run the script
    for kunci, df in objek_dataframe_report.items():
        print(f"Panjang data {kunci}: {len(df)} baris")
    # print path_output untuk debugging pada proses output dalam CI
    print(f"path_output: {path_output}")
    # buat daily retail report
    excel_rdsr(path_output, nama_file, objek_dataframe_report, tgl, nama_sheet)
    return {
        "lokasi_file": path_output / nama_file,
        "data": objek_dataframe_report
    }

def body_email_rdsr(data: dict[str, pd.DataFrame], tgl: date) -> str:
    def html_one_field_td_string(df: pd.DataFrame, df_key: Literal["SBU", "Area", "CNC"], mode: Literal["Daily", "WTD", "MTD"]) -> str:
        output_teks = ""
        # cek jika df_key tidak ada dalam kolom df maka raise exception dan exit
        if df_key not in df.columns.to_numpy().tolist():
            raise Exception("'df_key' tidak cocok dengan kolom pada dataframe")
        # assignment teks kolom berdasar mode
        match mode:
            case "Daily":
                sales = "Daily Sales"
                target = "Daily Target"
                target_ach = "Daily Ach."
            case "WTD":
                sales = "WTD Sales"
                target = "WTD Target"
                target_ach = "WTD Ach."
                ly_sales = "WTD LY Sales"
                ly_ach = "WTD LY Ach."
            case "MTD":
                sales = "MTD Sales"
                target = "MTD Target"
                target_ach = "MTD Ach."
                ly_sales = "MTD LY Sales"
                ly_ach = "MTD LY Ach."
            case _:
                raise Exception("'mode' tidak dikenali")
        # drop baris dengan nilai NaN pada kolom df_key
        df = df.dropna(subset=[df_key])
        # select specific kolom
        kolom = [df_key, sales, target, target_ach] if mode == "Daily" else [df_key, sales, target, target_ach, ly_sales, ly_ach]
        df_disp = df.copy()[kolom]
        # drop baris dengan semua nilai ada kolom sama dengan 0 pada kolom [1:]
        kolom_nilai = kolom[1:]
        df_disp = df_disp.loc[(df_disp[kolom_nilai] != 0.0).any(axis=1)]
        for _, baris in df_disp.iterrows():
            # styling percentage berdasar nilainya
            ach_style = '#FF004D' if baris.loc[target_ach] <= 0.5 else '#F6D776' if baris.loc[target_ach] <= 1 else '#B7E1B5'
            ly_ach_style = ('#FF004D' if baris.loc[ly_ach] <= 0.5 else '#F6D776' if baris.loc[ly_ach] <= 1 else '#B7E1B5') if mode != "Daily" else None
            # concatenate output baris dalam html
            if mode != "Daily":
                output_teks = output_teks + f"""
                    <tr style="background-color: '#35374B'">
                        <td style="padding: 10px; text-align: left; font-weight: bold">{baris.loc[df_key]}</td>
                        <td style="padding: 10px; text-align: center;">{locale.currency(baris.loc[sales], grouping=True)[:-3]}</td>
                        <td style="padding: 10px; text-align: center;">{locale.currency(baris.loc[target], grouping=True)[:-3]}</td>
                        <td style="padding: 10px; text-align: center;"><span style="color: {ach_style};">{'{:.2%}'.format(baris.loc[target_ach])}</span></td>
                        <td style="padding: 10px; text-align: center;">{locale.currency(baris.loc[ly_sales], grouping=True)[:-3]}</td>
                        <td style="padding: 10px; text-align: center;"><span style="color: {ly_ach_style};">{'{:.2%}'.format(baris.loc[ly_ach])}</span></td>
                    </tr>
                """
            else:
                output_teks = output_teks + f"""
                    <tr style="background-color: '#35374B'">
                        <td style="padding: 10px; text-align: left; font-weight: bold">{baris.loc[df_key]}</td>
                        <td style="padding: 10px; text-align: center;">{locale.currency(baris.loc[sales], grouping=True)[:-3]}</td>
                        <td style="padding: 10px; text-align: center;">{locale.currency(baris.loc[target], grouping=True)[:-3]}</td>
                        <td style="padding: 10px; text-align: center;"><span style="color: {ach_style};">{'{:.2%}'.format(baris.loc[target_ach])}</span></td>
                    </tr>
                """
        # return output_teks yang merupakan html
        return output_teks
    
    sbu_daily_string = html_one_field_td_string(data["sbu"], "SBU", "Daily") 
    area_daily_string = html_one_field_td_string(data["area"], "Area", "Daily")
    cnc_daily_string = html_one_field_td_string(data["cnc"], "CNC", "Daily")
    sbu_mtd_string = html_one_field_td_string(data["sbu"], "SBU", "MTD")
    area_mtd_string = html_one_field_td_string(data["area"], "Area", "MTD")
    cnc_mtd_string = html_one_field_td_string(data["cnc"], "CNC", "MTD")
    style_css = """
        <!--[if mso]>
        <style type="text/css">
            /* style CSS regular untuk desktop email client */
            body {
                background-color: #6C7B95;
                color: #EEEEEE;
                text-align: center;
            }
            .skala-tabel {
                background-color: #222831;
                text-align: left;
            }
            /* style CSS untuk tampilan mobile (Tab) */
            @media only screen and (max-width: 640px) {
                body {
                    background-color: #6C7B95;
                    color: #EEEEEE;
                    text-align: center;
                }
                .skala-tabel {
                    width: 440px !important;
                    margin: 0 !important;
                    background-color: #222831 !important;
                    text-align: left;
                }
            }
            /* style CSS untuk tampilan mobile (Phone) */
            @media only screen and (max-width: 479px) {
                body {
                    background-color: #6C7B95;
                    color: #EEEEEE;
                    text-align: center;
                }
                .skala-tabel {
                    width: 100% !important;
                    margin: 0 !important;
                    background-color: #222831 !important;
                    text-align: left;
                }
            }
        </style>
        <![endif]-->
    """
    return f"""
        <html xmlns="https://www.w3.org/1999/xhtml" lang="id" xml:lang="id"
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:o="urn:schemas-microsoft-com:office:office">
            <head>
                {style_css}
                <!--[if gte mso 9]><xml>
                <o:OfficeDocumentSettings>
                <o:AllowPNG/>
                <o:PixelsPerInch>96</o:PixelsPerInch>
                </o:OfficeDocumentSettings>
                </xml><![endif]-->
            </head>
            <body>
                <table cellpadding="10" cellspacing="0" class="skala-tabel">
                    <tr>
                        <td align="center">
                            <span style="font-size: 36px; color: #1597BB;">Selamat pagi!</span>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table width="100%" cellpadding="5" style="text-align: left;">
                                <tr>
                                    <td>
                                        Berikut adalah rangkuman OMNAS untuk data yang berakhir pada tanggal {tgl.strftime("%d %B %Y")}. 
                                        Data ini dikompilasi dari database PostgreSQL pada tanggal {datetime.now(timezone('Asia/Jakarta')).strftime('%d %B %Y pukul %H:%M:%S')}.<br><br>
                                        Untuk detail penjualan per toko dapat dilihat pada file terlampir pada email ini.<br><br>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table width="100%" cellpadding="10" style="text-align: left;">
                                <tr>
                                    <td>
                                        <strong style="font-size: 36px; color: #FFD523;">Daily Sales</strong>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table width="100%">
                                            <tr>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Strategic Business Unit</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% Target</th>
                                            </tr>
                                            {sbu_daily_string}
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table width="100%">
                                            <tr>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Area</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% Target</th>
                                            </tr>
                                            {area_daily_string}
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table width="100%"
                                            <tr>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Comp Stores Type</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% Target</th>
                                            </tr>
                                            {cnc_daily_string}
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <strong style="font-size: 36px; color: #FFD523;">Month to Date Sales</strong>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table width="100%">
                                            <tr>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Strategic Business Unit</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Last Year Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% Last Year</th>
                                            </tr>
                                            {sbu_mtd_string}
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table width="100%">
                                            <tr>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Area</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Last Year Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% Last Year</th>
                                            </tr>
                                            {area_mtd_string}
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <table width="100%"
                                            <tr>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Comp Stores Type</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% to Target</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">Last Year Sales</th>
                                                <th style="padding: 10px; text-align: center; font-size: 16px; background-color: #007F73;">% to Last Year</th>
                                            </tr>
                                            {cnc_mtd_string}
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <span style="color: #EC625F; font-size: 12px;"><i>* file terlampir dan konten pada email ini di-generate menggunakan python script yang running on schedule pada github actions, jika terjadi ketidaksesuaian data atau hal lainnya, mohon hubungi pengirim email.</i></span>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </body>
        </html>
    """
