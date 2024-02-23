from pathlib import Path
import os
import argparse
from datetime import date, timedelta

from utils.fungsi import (
    generate_daily_retail_report,
)

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

# path report output dan direktori_output
report_dirname = "reports"
dir_ini = Path(__file__).resolve().parent
dir_report = Path(os.path.join(dir_ini, report_dirname)).resolve()

# generate report
generate_daily_retail_report(dir_report, tgl_laporan, worksheet_list)
