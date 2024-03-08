import os
from pathlib import Path
from dotenv import load_dotenv

from utils.mongo import ambil_dokumen_by_id, ambil_field_dokumen


def memuat_env() -> dict[str, dict[str, str | None]]:
    try:
        file_konfig = ".config.env"
        direktori = Path(os.path.dirname(__file__)).parent
        file_path = os.path.join(direktori, file_konfig)

        load_dotenv(file_path)

        # return environment
        return {
            # O365 env
            "O365": {
                "TOKEN_FILE": os.getenv("O365_TOKEN_FILE"),
                "CLIENT_ID": os.getenv("O365_CLIENT_ID"),
                "CLIENT_SECRET": os.getenv("O365_CLIENT_SECRET"),
            },
            # PostgreSQL env
            "POSTGRESQL": {
                "HOST": os.getenv("HOST"),
                "PORT": os.getenv("PORT"),
                "DB": os.getenv("DBNAME"),
                "USER": os.getenv("USER"),
                "PWD": os.getenv("PASSWORD"),
            },
            # MongoDB URI
            "MONGO": {
                "URI": os.getenv("MONGO_URI"),
                "DB": os.getenv("MONGO_DB"),
                "KOLEKSI_PARAMETER": os.getenv("KOLEKSI_PARAMETER"),
                "ID_PARAM_BC": os.getenv("ID_PARAM_BC"),
                "ID_PARAM_PD": os.getenv("ID_PARAM_PD"),
            },
        }
    except:
        raise Exception("Gagal inisialisasi memuat environment")


def memuat_parameter_bc(mongo_env: dict[str, str]) -> dict[str, any]:
    try:
        dokumen_parameter_bc = ambil_dokumen_by_id(
            mongo_env["URI"],
            mongo_env["DB"],
            mongo_env["KOLEKSI_PARAMETER"],
            mongo_env["ID_PARAM_BC"],
        )
        return {
            "tabel": ambil_field_dokumen(dokumen_parameter_bc, "tabel_bc"),
            "kolom": ambil_field_dokumen(dokumen_parameter_bc, "kolom_bc"),
            "argumen": ambil_field_dokumen(dokumen_parameter_bc, "argumen_bc"),
        }
    except:
        raise Exception("Gagal inisialisasi memuat parameter kueri")


def memuat_parameter_pd(mongo_env: dict[str, str]) -> dict[str, any]:
    try:
        dokumen_parameter_pd = ambil_dokumen_by_id(
            mongo_env["URI"],
            mongo_env["DB"],
            mongo_env["KOLEKSI_PARAMETER"],
            mongo_env["ID_PARAM_PD"],
        )
        return {
            "nama": ambil_field_dokumen(dokumen_parameter_pd, "nama"),
            "tabel": ambil_field_dokumen(dokumen_parameter_pd, "tabel"),
            "kolom": ambil_field_dokumen(dokumen_parameter_pd, "kolom"),
        }
    except:
        raise Exception("Gagal inisialisasi memuat parameter pd")
