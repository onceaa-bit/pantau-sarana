# Pantau Sarana HCGA

Isi repo ini:

- `www/index.html`  aplikasi (web pengawas + aplikasi driver), satu berkas
- `gas/Code.gs`, `gas/appsscript.json`  backend Google Apps Script
- `.github/workflows/web.yml`  menerbitkan web (GitHub Pages)
- `.github/workflows/apk.yml`  membangun APK Android
- `scripts/`, `ci/`, `package.json`, `capacitor.config.json`  bahan pembangun APK

Langkah lengkap ada di dokumen "Panduan Pasang Driver Monitor HCGA".
Sebelum menjalankan workflow, isi variabel repo `API_URL` (Settings > Secrets and variables > Actions > Variables)
dengan URL Web App Apps Script yang berakhiran `/exec`.
