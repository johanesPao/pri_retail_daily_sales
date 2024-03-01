from typing import Literal
from datetime import date
import xlsxwriter as xlsx
from xlsxwriter.utility import xl_range
import pandas as pd


def buat_header_rdsr(
    tgl: date,
    objek_buku: xlsx.Workbook,
    objek_sheet: xlsx.Workbook.worksheet_class,
    kunci_df: Literal["sbu", "area", "cnc", "odd", "fisik", "bazaar"],
    jenis_df_nt: bool,
    panjang_data: int,
) -> None:
    # define kunci_props berdasar kunci_df
    # tes
    match kunci_df:
        case "sbu":
            kunci_props = {
                "judul": f"NATIONAL SALES BY SBU {tgl.strftime("%d %B %Y").upper()} [NON-PPN]"
            }
        case "area":
            kunci_props = {
                "judul": f"NATIONAL SALES BY AREA {tgl.strftime("%d %B %Y").upper()} [NON-PPN]"
            }
        case "cnc":
            kunci_props = {
                "judul": f"NATIONAL SALES BY COMP STORES {tgl.strftime("%d %B %Y").upper()} [NON-PPN]"
            }
        case "odd":
            kunci_props = {
                "judul": f"OUR DAILY DOSE NATIONAL SALES {tgl.strftime("%d %B %Y").upper()} [NON-PPN]"
            }
        case "fisik":
            kunci_props = {
                "judul": f"FISIK NATIONAL SALES {tgl.strftime("%d %B %Y").upper()} [NON-PPN]"
            }
        case "bazaar":
            kunci_props = {
                "judul": f"BAZAAR NATIONAL SALES {tgl.strftime("%d %B %Y").upper()} [NON-PPN]"
            }
        case _:
            raise TypeError(f"kunci_df '{kunci_df}' tidak dapat dikenali")
    # jika jenis df adalah non toko (jenis_df_nt True) kita akan melakukan merge sepanjang
    # kolom pada dataframe dikurangi 1 (panjang_data - 1) karena R1C1 pada xlsxwriter
    # zero-indexing, untuk jenis_df_nt False kita akan menggunakan panjang_judul = panjang_data
    panjang_judul = (panjang_data - 1) if jenis_df_nt else panjang_data
    # posisi kolom dan baris awal
    kolom = 0
    baris = 0
    # define range untuk judul
    range_judul = xl_range(baris, kolom, baris, panjang_judul)
    # format untuk judul
    format_judul = objek_buku.add_format(
        {
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "fg_color": "#000000",
            "font_size": 16,
            "font_color": "#FFFFFF",
            "border": 1,
        }
    )
    # format untuk judul field
    format_field = objek_buku.add_format(
        {
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 14,
            "border": 1
        }
    )
    # format untuk judul pada fix_header
    format_judul_fh = objek_buku.add_format(
        {
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 14,
            "border": 1,
            "bottom": 7
        }
    )
    # format untuk subjudul pada fix_header, tergantung posisi subjudul, ada 3 jenis format
    format_subjudul_fh_first = objek_buku.add_format(
        {
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 14,
            "border": 7,
            "bottom": 1,
            "left": 1
        }
    )
    format_subjudul_fh_mid = objek_buku.add_format(
        {
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 14,
            "border": 7,
            "bottom": 1
        }
    )
    format_subjudul_fh_last = objek_buku.add_format(
        {
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "font_size": 14,
            "border": 7,
            "right": 1,
            "bottom": 1
        }
    )
    # tes
    # tulis judul
    objek_sheet.merge_range(range_judul, kunci_props["judul"], format_judul)
    # tulis field header dan set offset_kolom untuk fix_header berdasar jenis_df_nt
    if jenis_df_nt:
        # offset kolom untuk fix_header
        offset_kolom = kolom + 1
        # range field spesifik untuk jenis_df_nt True
        range_field = xl_range(baris + 1, kolom, baris + 2, kolom)
        # match case untuk judul_field berdasar jenis df
        match kunci_df:
            case "sbu":
                judul_field = "Retail SBU"
            case "area":
                judul_field = "Stores Area"
            case "cnc":
                judul_field = "Comp/Non Comp Stores"
        # set judul field
        objek_sheet.merge_range(range_field, judul_field, format_field)
    else:
        # offset kolom untuk fix_header
        offset_kolom = kolom + 3
        # range field spesifik untuk jenis_df_nt False
        range_field_no = xl_range(baris + 1, kolom, baris + 2, kolom)
        range_field_kode_toko = xl_range(baris + 1, kolom + 1, baris + 2, kolom + 1)
        range_field_lokasi = xl_range(baris + 1, kolom + 2, baris + 2, kolom + 2)
        # judul field berdasar range
        judul_field_no = "No."
        judul_field_kode_toko = "Store Code"
        judul_field_lokasi = "Location"
        # set judul field
        objek_sheet.merge_range(range_field_no, judul_field_no, format_field)
        objek_sheet.merge_range(range_field_kode_toko, judul_field_kode_toko, format_field)
        objek_sheet.merge_range(range_field_lokasi, judul_field_lokasi, format_field)
    # tulis fix header
    # fix_header judul merge range dan judul fh
    # daily
    fh_daily_range = xl_range(baris + 1, offset_kolom, baris + 1, offset_kolom + 2)
    fh_daily_judul = "Daily"
    # wtd
    fh_wtd_range = xl_range(baris + 1, offset_kolom + 3, baris + 1, offset_kolom + 7)
    fh_wtd_judul = "Week to Date"
    # mtd
    fh_mtd_range = xl_range(baris + 1, offset_kolom + 8, baris + 1, offset_kolom + 12)
    fh_mtd_judul = "Month to Date"
    # tulis judul fix header
    objek_sheet.merge_range(fh_daily_range, fh_daily_judul, format_judul_fh)
    objek_sheet.merge_range(fh_wtd_range, fh_wtd_judul, format_judul_fh)
    objek_sheet.merge_range(fh_mtd_range, fh_mtd_judul, format_judul_fh)
    # define sub judul fix header
    fh_sales_subjudul = "Sales"
    fh_target_subjudul = "Target"
    fh_ach_subjudul = "Achieve"
    fh_ly_sales_subjudul = "LY Sales"
    fh_ly_persen_subjudul = "%LY"
    # tulis subjudul fix header
    # daily
    objek_sheet.write(baris + 2, offset_kolom, fh_sales_subjudul, format_subjudul_fh_first)
    objek_sheet.write(baris + 2, offset_kolom + 1, fh_target_subjudul, format_subjudul_fh_mid)
    objek_sheet.write(baris + 2, offset_kolom + 2, fh_ach_subjudul, format_subjudul_fh_last)
    # wtd
    objek_sheet.write(baris + 2, offset_kolom + 3, fh_sales_subjudul, format_subjudul_fh_first)
    objek_sheet.write(baris + 2, offset_kolom + 4, fh_target_subjudul, format_subjudul_fh_mid)
    objek_sheet.write(baris + 2, offset_kolom + 5, fh_ach_subjudul, format_subjudul_fh_mid)
    objek_sheet.write(baris + 2, offset_kolom + 6, fh_ly_sales_subjudul, format_subjudul_fh_mid)
    objek_sheet.write(baris + 2, offset_kolom + 7, fh_ly_persen_subjudul, format_subjudul_fh_last)
    # mtd
    objek_sheet.write(baris + 2, offset_kolom + 8, fh_sales_subjudul, format_subjudul_fh_first)
    objek_sheet.write(baris + 2, offset_kolom + 9, fh_target_subjudul, format_subjudul_fh_mid)
    objek_sheet.write(baris + 2, offset_kolom + 10, fh_ach_subjudul, format_subjudul_fh_mid)
    objek_sheet.write(baris + 2, offset_kolom + 11, fh_ly_sales_subjudul, format_subjudul_fh_mid)
    objek_sheet.write(baris + 2, offset_kolom + 12, fh_ly_persen_subjudul, format_subjudul_fh_last)

def tulis_data_rdsr(
    file: str,
    objek_df: dict[str, pd.DataFrame],
    tgl: date,
    nama_sheet: str | dict[str, str] = "Sheet1",
) -> None:
    # secara keseluruhan setidaknya terdapat dua jenis dataframe pada objek_df. jenis dataframe pertama
    # adalah dataframe non toko seperti sbu, area dan cnc yang akan menampilkan field sbu, nama area dan
    # kategori comp non comp stores. sedangkan dataframe kedua adalah dataframe yang menampilkan field
    # daftar toko. perlakuan kedua jenis dataframe ini akan cukup berbeda, sehingga kita perlu untuk
    # memisahkan kedua jenis dataframe ini ke dalam dua kelompok.
    # untuk list kunci semua dataframe dalam objek_df, merujuk pada fungsi.py baris 385-392
    # buat list kunci dalam objek_df
    kunci_df = list(objek_df.keys())
    # buat list kunci masing - masing jenis dataframe
    kunci_df_nt = [kunci_df[0], kunci_df[1], kunci_df[2]]
    kunci_df_t = [kunci_df[3], kunci_df[4], kunci_df[5]]
    # buat workbook report
    buku_report = xlsx.Workbook(file)
    # loop pada objek_df dan lakukan penulisan data
    for kunci, data in objek_df.items():
        # untuk debugging
        print(f"{kunci}: ", data.shape)
        # jika kunci terdapat dalam nama_sheet.keys()
        if kunci in nama_sheet.keys():
            # buat sheet dengan nilai dimana kunci pada objek_df sebagai kunci pada nama_sheet
            sheet_report = buku_report.add_worksheet(nama_sheet[kunci])
            # set zoom 80%
            sheet_report.set_zoom(80)
            # sembunyikan gridline
            sheet_report.hide_gridlines(2)
            # define tinggi dan lebar data (data.shape tidak menghitung header/nama kolom dan index pada
            # dataframe)
            tinggi_data, lebar_data = data.shape
            # buat header
            buat_header_rdsr(
                tgl, buku_report, sheet_report, kunci, kunci in kunci_df_nt, data.shape[1]
            )

    # tutup dan output workbook report
    buku_report.close()


def excel_rdsr(
    path_output: str,
    nama_file: str,
    objek_df: dict[str, pd.DataFrame],
    tgl: date,
    nama_sheet: str | dict[str, str] = "Sheet1",
) -> None:
    print("Memulai penulisan retail daily sales report pada file excel...")
    # tulis objek_df pada masing - masing sheet
    tulis_data_rdsr(path_output / nama_file, objek_df, tgl, nama_sheet)
