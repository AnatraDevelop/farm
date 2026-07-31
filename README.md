# Anatra Farm — Aplikasi Nota & Laporan (Google Apps Script)

Aplikasi web internal untuk input nota, stok gudang, laporan keuntungan, dan master
harga modal telur. Dibangun di atas **Google Apps Script + Google Sheets** sebagai
database.

## ⚠️ Penting: GitHub hanya untuk menyimpan source code

Google Apps Script (`SpreadsheetApp`, `HtmlService`, `LockService`, dll) hanya bisa
dieksekusi di server milik Google. GitHub (termasuk GitHub Pages) **tidak bisa
menjalankan** kode ini — GitHub Pages hanya melayani file statis (HTML/CSS/JS murni
di browser), sedangkan aplikasi ini butuh backend server-side untuk baca/tulis ke
Google Sheets.

Jadi alur yang benar:

```
GitHub (source code / version control)  --->  clasp push  --->  Google Apps Script (tempat aplikasi benar-benar berjalan)
```

Ada 2 cara deploy dari repo ini ke Apps Script:

## Opsi A — Manual (paling simpel, tanpa install apa pun)

1. Buka [script.google.com](https://script.google.com) → buat project baru.
2. Copy-paste isi `Code.gs` ke file kode (`Code.gs`) di editor Apps Script.
3. Buat file HTML baru bernama **`Index`** (tanpa `.html`), copy-paste isi `Index.html`.
4. Buka menu **Project Settings** (ikon gerigi) → **Script Properties** → tambahkan:
   - Key: `SPREADSHEET_ID`
   - Value: ID Google Sheet database Anda (bagian di URL Sheet, antara `/d/` dan `/edit`)
5. Klik **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone** (atau sesuai kebutuhan)
6. Salin URL `.../exec` yang diberikan — itulah alamat aplikasi Anda.

## Opsi B — Otomatis pakai `clasp` (direkomendasikan jika sering update dari GitHub)

`clasp` adalah CLI resmi Google untuk sinkronisasi kode antara komputer/GitHub dan
Apps Script.

```bash
npm install -g @google/clasp
clasp login

# Jika project Apps Script SUDAH ada:
clasp clone <SCRIPT_ID>

# Jika BELUM ada, buat baru dari repo ini:
clasp create --type webapp --title "Anatra Farm"

# Salin .clasp.json.example -> .clasp.json lalu isi scriptId Anda
cp .clasp.json.example .clasp.json

# Push kode dari repo ke Apps Script
clasp push

# Deploy jadi web app
clasp deploy
```

Setelah itu, setiap kali ada perubahan di GitHub, tinggal `git pull` lalu
`clasp push` untuk mengirim ke Apps Script — tidak perlu copy-paste manual lagi.

> `.clasp.json` sengaja di-`.gitignore`-kan karena berisi `scriptId` milik Apps
> Script pribadi Anda dan tidak perlu ikut di-commit ke repo publik.

## Struktur file

| File | Keterangan |
|---|---|
| `Code.gs` | Seluruh logic backend (baca/tulis Google Sheets, generate nota, laporan, dll) |
| `Index.html` | Seluruh UI (HTML + CSS + JS front-end), dirender oleh `HtmlService` |
| `appsscript.json` | Manifest wajib Apps Script (timezone, izin web app, dll) |
| `.clasp.json.example` | Contoh konfigurasi `clasp`, copy jadi `.clasp.json` lalu isi `scriptId` |

## Keamanan sebelum push ke repo publik

- **ID Spreadsheet** sudah dipindah agar diambil dari **Script Properties**
  (`SPREADSHEET_ID`) — bukan hardcode di kode. Ini penting kalau repo di-publish
  publik di GitHub, supaya orang lain tidak langsung tahu ID Sheet database Anda.
- Kredensial user (`USERS` sheet: username/password) **tetap di Google Sheets**,
  bukan di kode — jadi aman untuk di-commit ke GitHub selama kode tidak membocorkan
  isi sheet.
- Jangan pernah commit `.clasp.json` atau `.clasprc.json` (sudah di `.gitignore`).
