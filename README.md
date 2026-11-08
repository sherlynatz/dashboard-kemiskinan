# Dashboard Kemiskinan dan Ketimpangan di Lima Provinsi, 2019–2023

Dashboard interaktif (Python, Streamlit, Plotly) untuk Tugas Proyek Akhir: Perancangan dan Implementasi Dashboard Data Ekonomi Interaktif.

**Tautan dashboard:** https://dashboard-kemiskinan-dg8c8p2zvcdgd9ueqev2ku.streamlit.app/

**Disusun oleh:** Sherly Natasya Zahwa

## Indikator
- Persentase penduduk miskin (P0), total, perkotaan, dan perdesaan (%)
- Jumlah penduduk miskin (ribu orang)
- Garis kemiskinan (Rp per kapita per bulan, nominal)
- Gini Ratio (indeks 0–1)

Cakupan: Jawa Timur, Jawa Tengah, Jawa Barat, DKI Jakarta, dan Bali; periode Maret 2019–2023; pembanding nasional.

## Fitur
- 4 KPI di bagian atas; P0 gabungan dihitung tertimbang penduduk.
- 3 filter interaktif: rentang tahun, pilihan provinsi, dan pembanding nasional.
- 5 grafik: line chart (2), grouped bar, scatter dengan garis tren OLS, dan bar horizontal.
- Palet Okabe-Ito yang ramah buta warna; provinsi juga dibedakan dengan bentuk marker.

## Menjalankan di komputer sendiri
```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
pip install -r requirements.txt
streamlit run app.py
```

## Struktur berkas
| Berkas | Fungsi |
|---|---|
| `app.py` | Kode dashboard |
| `data_kemiskinan_dashboard.xlsx` | Dataset bersih (sheet Data Provinsi, Data Nasional, Catatan) |
| `requirements.txt` | Daftar pustaka |
| `.streamlit/config.toml` | Tema terang |

## Sumber data dan pra-pemrosesan
Sumber: Badan Pusat Statistik (BPS), periode Maret (Semester 1). Tidak ada angka yang diubah; nilai P0 pada data bersih sudah dicocokkan dengan data mentah. P0 perdesaan DKI Jakarta tidak tersedia di sumber, sehingga dikosongkan dan tidak dihitung (bukan nol). Detail lengkap ada pada sheet *Catatan* di file Excel.
