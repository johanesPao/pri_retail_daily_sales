import xlsxwriter as xlsx
import pandas as pd


def buat_file_excel(
    filepath: str, nama_file: str, sheet: str | list[str] = "Sheet1"
) -> None:
    workbook = xlsx.Workbook(filepath / nama_file)
    if isinstance(sheet, str):
        workbook.add_worksheet(sheet)
    if isinstance(sheet, list) and all(isinstance(elemen, str) for elemen in sheet):
        for nama_sheet in sheet:
            workbook.add_worksheet(nama_sheet)
    workbook.close()


def excel_daily_retail_report(
    path_output: str,
    nama_file: str,
    objek_df: dict[str, pd.DataFrame],
    nama_sheet: str | list[str] = "Sheet1",
) -> None:
    print(objek_df)
    # buat file excel
    buat_file_excel(path_output, nama_file, nama_sheet)
