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

def set_kolom_field_rdsr(
    objek_sheet: xlsx.Workbook.worksheet_class,
    jenis_df_nt: bool        
) -> None:
    if jenis_df_nt:
        objek_sheet.set_column(0,0,29)
    else:
        objek_sheet.set_column(0,0, 4.57)
        objek_sheet.set_column(1,1, 13.14)
        objek_sheet.set_column(2,2, 51)

def set_kolom_data_rdsr(
    objek_sheet: xlsx.Workbook.worksheet_class,
    kolom_awal: int,
    jumlah_kolom: int = 13,
    lebar_kolom: float = 15.57
) -> None:
    # set lebar kolom
    objek_sheet.set_column(kolom_awal, jumlah_kolom + kolom_awal - 1, lebar_kolom)
    # set freeze pane
    objek_sheet.freeze_panes(3, kolom_awal)

def buat_list_kolom_persen(
    jumlah_kolom_field: int = 1,
) -> list[int]:
    return [jumlah_kolom_field + 2, jumlah_kolom_field + 5, jumlah_kolom_field + 7, jumlah_kolom_field + 10, jumlah_kolom_field + 12]


def loop_dataframe_rdsr(
    kunci_df: Literal["sbu", "area", "cnc", "odd", "fisik", "bazaar"],
    jenis_df_nt: bool,
    objek_buku: xlsx.Workbook,
    objek_sheet: xlsx.Workbook.worksheet_class,
    df: pd.DataFrame,
) -> None:
    print(f"Setup styling {kunci_df}...")
    # judul_cnc
    judul_cnc = ["COMP STORES TOTAL", "NON COMP STORES TOTAL", "STORES TOTAL"]
    # format_judul_field
    format_judul_field = objek_buku.add_format({
        "bold": 1,
        "fg_color": "#E2EFDA"
    })
    format_judul_cnc = objek_buku.add_format({
        "bold": 1,
        "align": "center",
        "fg_color": "#FFF2CC"
    })
    print(f"Melakukan iterasi data {kunci_df}...")
    # set baris awal
    baris = 3
    # format persen dalam fungsi
    def format_persen(
        nilai: float, 
        objek_buku: xlsx.Workbook, 
        mode_cnc: bool
    ) -> xlsx.workbook.Format:
        return objek_buku.add_format({
            "bold": 1,
            "num_format": "0.0%",
            "fg_color": "#FFF2CC" if mode_cnc else "#E2EFDA",
            "color": "#FF204E" if nilai <= 0.5 else "#E8751A" if nilai <= 1 else "#4CCD99"
        })
    # format nominal
    format_nominal = objek_buku.add_format({
        "num_format": "#,##0.00,,_)jt;(#,##0.00,,)jt",
        "fg_color": "#E2EFDA"
    })
    format_nominal_cnc = objek_buku.add_format({
        "num_format": "#,##0.00,,_)jt;(#,##0.00,,)jt",
        "fg_color": "#FFF2CC"
    })
    # set kolom dengan num_format persen (zero-based)
    kolom_persen = buat_list_kolom_persen(1 if jenis_df_nt else 3)
    # iterasi pada dataframe df
    for indeks, baris_data in df.iterrows():
        # jika jenis_df_nt True maka kita tidak perlu untuk melakukan evaluasi indeks pada dataframe    
        if jenis_df_nt:
            for hitung in range(len(baris_data)):
                # pada kolom 0 kita ingin mengevaluasi apakah indeks adalah str atau tidak
                if hitung == 0:
                    if isinstance(indeks, str):
                        # tulis indeks pada kolom 0
                        objek_sheet.write(baris, hitung, indeks, format_judul_cnc if indeks in judul_cnc else format_judul_field)
                    else:
                        objek_sheet.write(baris, hitung, baris_data.iloc[hitung], format_judul_field)
                # jika bukan kolom 0
                else:
                    if isinstance(indeks, str) and indeks in judul_cnc:
                        objek_sheet.write(baris, hitung, baris_data.iloc[hitung], format_persen(baris_data.iloc[hitung], objek_buku, True) if hitung in kolom_persen else format_nominal_cnc)
                    else:
                        objek_sheet.write(baris, hitung, baris_data.iloc[hitung], format_persen(baris_data.iloc[hitung], objek_buku, False) if hitung in kolom_persen else format_nominal)
        else:
            # dalam kasus df bukanlah jenis_df_nt, maka kita akan melakukan pengisian tiga kolom awal dengan
            # No, Store Code dan Location
            # kita juga akan mengoffset loop + 1 untuk mengkompensasi penulisan indeks sebagai No atau CNC
            offset_kolom_regular = 1
            offset_kolom_cnc = 2 # zero-based
            for hitung in range(len(baris_data) + offset_kolom_regular):
                # jika hitung == 0:
                if hitung == 0:
                    # cek jika indeks adalah str
                    if isinstance(indeks, str):
                        # merge range cell dengan R1C1 1,3
                        objek_sheet.merge_range(baris, 0, baris, 2, indeks, format_judul_cnc if indeks in judul_cnc else format_judul_field)
                    else:
                        # tulis indeks tanpa merge
                        objek_sheet.write(baris, hitung, indeks, format_judul_field)
                elif hitung > 0 and hitung <= offset_kolom_cnc:
                    # cek jika indeks bukan instance of str, jika instance of str, do nothing
                    if not isinstance(indeks, str):
                        objek_sheet.write(baris, hitung, baris_data.iloc[hitung - offset_kolom_regular], format_judul_field)
                # jika hitung lebih besar atau sama dengan offset_kolom_cnc
                else:
                    if isinstance(indeks, str) and indeks in judul_cnc:
                        objek_sheet.write(baris, hitung, baris_data.iloc[hitung - offset_kolom_regular], format_persen(baris_data.iloc[hitung - offset_kolom_regular], objek_buku, True) if hitung in kolom_persen else format_nominal_cnc)
                    else:
                        objek_sheet.write(baris, hitung, baris_data.iloc[hitung - offset_kolom_regular], format_persen(baris_data.iloc[hitung - offset_kolom_regular], objek_buku, False) if hitung in kolom_persen else format_nominal)
        # increment baris
        baris += 1
        # lebar kolom dan freeze pane
        set_kolom_field_rdsr(objek_sheet, jenis_df_nt)
        set_kolom_data_rdsr(objek_sheet, 1 if jenis_df_nt else 3)
        # sembunyikan baris yang tidak dipergunakan
        objek_sheet.set_default_row(hide_unused_rows=True)
        # kolom akhir pada excel (zero-based R1C1)
        kolom_akhir_excel = 16383
        # kolom_awal_sembunyi (zero-based R1C1, kondisional terhadap jenis_df_nt)
        kolom_awal_sembunyi = len(baris_data) if jenis_df_nt else len(baris_data) + 1
        # sembunyikan kolom yang tidak dipergunakan
        objek_sheet.set_column(kolom_awal_sembunyi, kolom_akhir_excel, None, None, {"hidden": True})


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
    # buat list kunci dataframe jenis nt
    kunci_df_nt = [kunci_df[0], kunci_df[1], kunci_df[2]]
    # buat workbook report
    buku_report = xlsx.Workbook(file)
    # loop pada objek_df dan lakukan penulisan data
    for kunci, data in objek_df.items():
        # jika kunci terdapat dalam nama_sheet.keys()
        if kunci in nama_sheet.keys():
            # buat sheet dengan nilai dimana kunci pada objek_df sebagai kunci pada nama_sheet
            sheet_report = buku_report.add_worksheet(nama_sheet[kunci])
            # set zoom 80%
            sheet_report.set_zoom(80)
            # sembunyikan gridline
            sheet_report.hide_gridlines(2)
            # buat header
            buat_header_rdsr(
                tgl, buku_report, sheet_report, kunci, kunci in kunci_df_nt, data.shape[1]
            )
            # looping dalam dataframe
            loop_dataframe_rdsr(kunci, kunci in kunci_df_nt, buku_report, sheet_report, data)

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
