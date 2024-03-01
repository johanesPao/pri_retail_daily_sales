from pathlib import Path
import os
import argparse
from datetime import date, timedelta
from utils.inisiasi import memuat_env, memuat_parameter_bc, memuat_parameter_pd

from utils.fungsi import (
    generate_daily_retail_report,
)

# command-line parser
print("Parsing argumen...")
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
print(f"Proses dimulai untuk penarikan data {tgl_laporan}...")

# muat environment dan simpan pada variabel env
print("Memuat environment...")
env = memuat_env()

# memuat parameter untuk kueri
print("Memuat parameter kueri...")
parameter = {
    "bc": memuat_parameter_bc(env["MONGO"]),
    "pd": memuat_parameter_pd(env["MONGO"])
}

# setup worksheet list
worksheet_list = {
    "sbu": "Dashboard SBU",
    "area": "Dashboard Area",
    "cnc": "Dashboard Comp & Non Comp Store",
    "odd": "Our Daily Dose",
    "fisik": "Fisik",
    "bazaar": "Bazaar",
    "detailed": "Detailed Daily Sales & Est.",
}

# path report output dan direktori_output
report_dirname = "reports"
dir_ini = Path(__file__).resolve().parent
dir_report = Path(os.path.join(dir_ini, report_dirname)).resolve()

# generate report
print("Memulai pembuatan report...")
generate_daily_retail_report(dir_report, tgl_laporan, env, parameter, worksheet_list)
