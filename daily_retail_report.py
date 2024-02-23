from pathlib import Path
import os
import argparse
from datetime import date, timedelta
from typing import Literal

from utils.fungsi import generate_df_utama, generate_df_single_sbu, generate_report

# command-line parser
parser = argparse.ArgumentParser()
# argument untuk manual run script untuk tanggal tertentu dalam kasus manual re-run diperlukan
# menerima argumen command line -t atau --tanggal
# jika script dijalankan tanpa argumen maka default ke tanggal kemarin
parser.add_argument(
    "-t",
    "--tanggal",
    type=date.fromisoformat,
    default=date.today() - timedelta(days=1),
    help="tanggal dalam iso format YYYYMMDD, contoh 20240131",
)
# parse argumen atau default ke tanggal kemarin
tgl_laporan = parser.parse_args().tanggal

# generate dataframe utama
df_utama = generate_df_utama(tgl_laporan)

# konfigurasi
konfig = [
    {"nama_tabel": "sbu", "nama_sheet": "Dashboard SBU"},
    {"nama_tabel": "area", "nama_sheet": "Dashboard Area"},
    {"nama_tabel": "cnc", "nama_sheet": "Dashboard Comp & Non Comp Store"},
    {"nama_tabel": "odd", "nama_sheet": "Our Daily Dose"},
    {"nama_tabel": "fisik", "nama_sheet": "Fisik"},
    {"nama_tabel": "bazaar", "nama_sheet": "Bazaar"},
    {"nama_tabel": "detailed", "nama_sheet": "Detailed Daily Sales & Est."},
]

# setup worksheet list
worksheet_list = [
    "Dashboard SBU",
    "Dashboard Area",
    "Dashboard Comp & Non Comp Store",
    "Our Daily Dose",
    "Fisik",
    "Bazaar",
    "Detailed Daily Sales & Est.",
]

# generate tabel untuk dimuat pada masing - masing sheet


# path report template dan direktori_output
report_dirname = "reports"
dir_ini = Path(__file__).resolve().parent
dir_report = Path(os.path.join(dir_ini, report_dirname)).resolve()

# generate report
generate_report(dir_ini, dir_report, konfig, df_utama, tgl_laporan)
