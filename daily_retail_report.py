import locale
from pathlib import Path
import os
import argparse
from datetime import date, timedelta

from utils.inisiasi import memuat_env, memuat_parameter_bc, memuat_parameter_pd
from utils.fungsi import (
    body_email_rdsr,
    generate_daily_retail_report,
    kirim_email,
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
tgl_laporan: date = parser.parse_args().tanggal
print(f"Proses dimulai untuk penarikan data {tgl_laporan}...")

# muat environment dan simpan pada variabel env
print("Memuat environment...")
env = memuat_env()

# set lokalisasi
locale.setlocale(locale.LC_ALL, 'id_ID')

# memuat parameter untuk kueri
print("Memuat parameter kueri...")
parameter = {"bc": memuat_parameter_bc(env["MONGO"]), "pd": memuat_parameter_pd(env["MONGO"])}

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
konten = generate_daily_retail_report(dir_report, tgl_laporan, env, parameter, worksheet_list)

# generate atribut email
atribut_email = env["EMAIL"]
from_email = atribut_email["FROM"]
# define true jika script berjalan dalam environment Continuous Integration 
# dan di-trigger oleh workflow dengan branch dev
run_in_dev_ci = env["CI"]["IN_CI"] and env["CI"]["WORKFLOW_BRANCH"] == 'dev'
# bypass assignment dua variabel di bawah menjadi from_email jika script dijalankan 
# dalam CI dengan workflow branch 'dev'
to_email = from_email if run_in_dev_ci else atribut_email["TO"]
cc_email = from_email if run_in_dev_ci else atribut_email["CC"]
judul_email = f"PRI Retail Daily Sales Summary ({tgl_laporan.strftime("%d %B %Y")})"
body = body_email_rdsr(konten["data"], tgl_laporan)

# kirim email ke pengguna
print("Mengirimkan email ke pengguna...")
kirim_email(env, [konten["lokasi_file"]], from_email, to_email, cc_email, judul_email, body)

print("Script selesai dieksekusi!")
