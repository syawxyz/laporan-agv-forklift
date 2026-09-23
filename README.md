# Laporan Proyek AGV Forklift Otonom — dokumen dan data pendukung

Repositori ini hanya berisi berkas pelaporan (laporan .docx, data ringkasan baterai, grafik, dan skrip pembuat laporan). Kode navigasi, algoritma, dan firmware robot **tidak** disertakan.

| Berkas | Isi |
|---|---|
| `Laporan-Proyek-AGV-Forklift-2026-09-23.docx` | Laporan pengembangan 6 Agustus – 18 September 2026 (Bab 1–5, Lampiran A–C) + **Lampiran D** bukti data aktual dari repositori + **Lampiran E** laporan konsumsi daya baterai. |
| `baterai/ringkasan_sesi.csv` | Statistik tiap sesi data_logger (`logs/agv_*.csv`, 314 sesi 11/8–18/9): durasi, jarak, tegangan DC-link awal/akhir/min/maks, arus traksi rata/maks, energi V·I·Δt gerak/diam. |
| `baterai/per_hari.csv`, `baterai/per_sesi.csv` | Agregat per hari dan per sesi (jarak odometri yang melompat sudah dinolkan; pembacaan < 40 V dibuang). |
| `baterai/deret_tegangan_1hz.csv.gz` | Deret waktu 1 Hz: sesi, detik, tegangan, arus L+R, jarak — bahan untuk menghitung ulang grafik. |
| `baterai/hasil.json` | Angka kunci yang dikutip Lampiran E. |
| `baterai/g*.png` | Grafik Lampiran E. |
| `skrip/scan_baterai.py` | Pemindai `logs/agv_*.csv` → `ringkasan_sesi.csv` + deret 1 Hz (paralel, bisa dilanjutkan). |
| `skrip/analisis_baterai.py` | Ringkasan → tabel, grafik, `hasil.json`. |
| `skrip/tambah_bukti.py`, `skrip/tambah_baterai.py`, `skrip/docx_alat2.py` | Penyisip Lampiran D dan E ke docx (python-docx) tanpa mengubah isi lama. |

