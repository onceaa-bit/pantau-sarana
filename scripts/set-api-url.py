#!/usr/bin/env python3
"""Isi alamat server (URL Web App Apps Script) ke www/index.html sebelum dibangun/diterbitkan."""
import os, re, sys, pathlib

url = os.environ.get('API_URL', '').strip()
if not re.fullmatch(r'https://script\.google\.com/.+/exec', url):
    sys.exit('GAGAL: variabel API_URL belum diisi atau salah.\n'
             'Buka Settings > Secrets and variables > Actions > tab Variables > New repository variable,\n'
             'nama API_URL, isi dengan URL Web App Apps Script yang berakhiran /exec, lalu jalankan ulang.')
p = pathlib.Path(__file__).resolve().parent.parent / 'www' / 'index.html'
s = p.read_text(encoding='utf-8')
old = "window.DM_CONFIG = { API_URL: '' };"
if old not in s:
    sys.exit('GAGAL: baris DM_CONFIG tidak ditemukan di www/index.html (file sudah diubah?).')
p.write_text(s.replace(old, "window.DM_CONFIG = { API_URL: '%s' };" % url.replace("'", '')), encoding='utf-8')
print('API_URL terisi:', url[:45] + '...')
