from pymongo import MongoClient
from bson.objectid import ObjectId

def buka_koneksi(uri: str) -> MongoClient:
    try:
        print("Membuka koneksi ke MongoDB...")
        return MongoClient(uri)
    except:
        raise Exception('Gagal membuka koneksi ke MongoDB')
    
def hubungkan_koleksi(uri: str, database: str, koleksi: str) -> any:
    try:
        klien = buka_koneksi(uri)
        print(f"Menghubungkan dengan koleksi {koleksi}...")
        db = klien[database]
        return db[koleksi]
    except:
        raise Exception("Gagal terhubung dengan koleksi pada MongoDB")
    
def ambil_dokumen_by_id(uri: str, database: str, koleksi: str, id: str) -> dict[str, any]:
    try:
        kol = hubungkan_koleksi(uri, database, koleksi)
        print(f"Melakukan kueri find_one() pada koleksi {koleksi} untuk _id: ObjectId('{id}')...")
        return kol.find_one({'_id': ObjectId(id)})
    except:
        raise Exception(f"Gagal melakukan pencarian dokumen dalam koleksi {koleksi}")
    
def ambil_field_dokumen(dict: dict[str, any], field: str) -> dict[str, any]:
    try:
        print(f"Melakukan ekstraksi field {field}...")
        return dict[field]
    except:
        raise Exception(f"Tidak dapat menemukan field '{field}' dalam dokumen")