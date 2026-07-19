# Starcom Executive Platform (SEP) V4 — AI Presenter

Paket ini siap diunggah ke GitHub dan dideploy melalui Streamlit Community Cloud.

## Isi repository

- `app.py` — aplikasi utama SEP V4 dan AI Presenter
- `data/SEP_V4_Final_Project_Productivity_Template4.xlsx` — data default
- `assets/logo_company.png` — logo perusahaan
- `.streamlit/config.toml` — tema dan konfigurasi Streamlit
- `.streamlit/secrets.example.toml` — contoh konfigurasi API key
- `requirements.txt` — dependensi Python
- `runtime.txt` — versi Python untuk deployment

## Deploy ke Streamlit Community Cloud

1. Upload seluruh isi folder ini ke root repository GitHub.
2. Di Streamlit Community Cloud, pilih **Create app**.
3. Pilih repository dan branch `main`.
4. Isi **Main file path** dengan `app.py`.
5. Buka **App settings > Secrets** dan masukkan:

```toml
OPENAI_API_KEY = "sk-API_KEY_ANDA"
```

6. Klik **Deploy** atau **Reboot app**.

## Fitur AI Presenter

- Tombol AI Presenter mengambang di kanan atas.
- Pertanyaan melalui teks.
- Pertanyaan suara melalui mikrofon browser.
- Transkripsi Bahasa Indonesia.
- Jawaban berbasis data aktif SEP V4.
- Jawaban audio melalui text-to-speech.
- Visual fokus sesuai topik pertanyaan.
- Fallback jawaban lokal untuk pertanyaan teks jika API key belum dipasang.

## Catatan keamanan

Jangan unggah API key ke GitHub. API key hanya dimasukkan pada menu **Secrets** di Streamlit Cloud.

## Pemeriksaan data

Aplikasi membaca file default dari:

`data/SEP_V4_Final_Project_Productivity_Template4.xlsx`

File Excel harus mempertahankan nama sheet wajib: `PARAMETERS`, `TARGETS`, `FJA_MAPPING`, `DIM_DEPARTMENT`, `REVENUE_YR`, `HEADCOUNT_YR`, dan `PAYROLL_YR`.
