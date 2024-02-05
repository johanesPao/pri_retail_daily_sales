class KesalahanKoneksi(Exception):
    pass


class KesalahanKueri(Exception):
    pass


def tangkap_kesalahan_koneksi(exception):
    print(f"Terjadi kesalahan koneksi: {exception}")
    raise KesalahanKoneksi


def tangkap_kesalahan_kueri(exception):
    print(f"Terjadi kesalahan kueri: {exception}")
    raise KesalahanKueri
