# MATERIAL-USAGE

Web form rasmi untuk merekod penggunaan atau pinjaman material.

## Teknologi

- React 18 (melalui CDN)
- Babel Standalone (untuk JSX di dalam fail HTML)
- Integrasi Google Apps Script sebagai backend

## Ciri Utama

- Antara muka formal dan responsif
- Tambah atau buang baris material secara dinamik
- Autocadangan nama material melalui `datalist`
- Penghantaran rekod ke endpoint Google Apps Script

## Cara Guna

1. Buka `index.html` dalam pelayar (atau guna Five Server di VS Code).
2. Isi nama pengguna dan tujuan penggunaan.
3. Isi satu atau lebih material bersama kuantiti.
4. Klik butang **Hantar Rekod**.

## Endpoint Backend

URL backend ditetapkan dalam kod JavaScript pada `index.html`:

`https://script.google.com/macros/s/AKfycbyP8rKrQFT4Ho0g5lyI51dQZ-VBA-a9iKmELX_qk1D6xKkZ_bYwL9PYxgEKU09qO04f/exec`
