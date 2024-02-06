from pathlib import Path
import os
import xlwings as xw

from utils.fungsi import (
    generate_df_utama,
    generate_report,
)

# generate dataframe utama
df_utama = generate_df_utama()

# path report template dan direktori_output
report_dirname = "reports"
dir_ini = Path(__file__).resolve().parent
dir_report = Path(os.path.join(dir_ini, report_dirname)).resolve()

# generate report
generate_report(dir_ini, dir_report, df_utama)
