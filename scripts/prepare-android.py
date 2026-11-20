#!/usr/bin/env python3
"""Dijalankan setelah `npx cap add android`: tambah izin, ganti ikon dan warna aplikasi."""
import pathlib, re, sys
from PIL import Image, ImageDraw

root = pathlib.Path(__file__).resolve().parent.parent / 'android' / 'app' / 'src' / 'main'
manifest = root / 'AndroidManifest.xml'
if not manifest.exists():
    sys.exit('AndroidManifest.xml tidak ada. Jalankan `npx cap add android` dulu.')

# --- izin ---
m = manifest.read_text(encoding='utf-8')
need = ['VIBRATE', 'WAKE_LOCK', 'ACCESS_FINE_LOCATION', 'ACCESS_COARSE_LOCATION', 'POST_NOTIFICATIONS', 'FOREGROUND_SERVICE', 'FOREGROUND_SERVICE_LOCATION']
add = ''.join('    <uses-permission android:name="android.permission.%s" />\n' % n for n in need if 'android.permission.%s"' % n not in m)
m = m.replace('</manifest>', add + '</manifest>')
manifest.write_text(m, encoding='utf-8')

# --- ikon: pin lokasi kuning di atas dasar gelap ---
INK, YEL = (18, 32, 42, 255), (246, 216, 12, 255)
def pin(size, scale):
    im = Image.new('RGBA', (size * 4, size * 4), (0, 0, 0, 0)); d = ImageDraw.Draw(im); S = size * 4
    cx, cy, r = S / 2, S * 0.42, S * 0.20 * scale
    d.polygon([(cx - r * 0.92, cy + r * 0.45), (cx + r * 0.92, cy + r * 0.45), (cx, cy + r * 2.0)], fill=YEL)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=YEL)
    d.ellipse([cx - r * 0.42, cy - r * 0.42, cx + r * 0.42, cy + r * 0.42], fill=INK)
    return im.resize((size, size), Image.LANCZOS)
def legacy(size, round_):
    bg = Image.new('RGBA', (size, size), (0, 0, 0, 0)); d = ImageDraw.Draw(bg)
    (d.ellipse if round_ else lambda b, fill: d.rounded_rectangle(b, radius=size * 0.18, fill=fill))([0, 0, size - 1, size - 1], fill=INK)
    bg.alpha_composite(pin(size, 1.0)); return bg
for dens, px in {'mdpi': 48, 'hdpi': 72, 'xhdpi': 96, 'xxhdpi': 144, 'xxxhdpi': 192}.items():
    d = root / 'res' / ('mipmap-' + dens)
    if not d.exists(): continue
    legacy(px, False).save(d / 'ic_launcher.png'); legacy(px, True).save(d / 'ic_launcher_round.png')
    fg = px * 108 // 48                                   # lapisan depan ikon adaptif = 108dp
    pin(fg, 0.78).save(d / 'ic_launcher_foreground.png')
bgc = root / 'res' / 'values' / 'ic_launcher_background.xml'
if bgc.exists():
    bgc.write_text(re.sub(r'#[0-9A-Fa-f]{6}', '#12202A', bgc.read_text()), encoding='utf-8')
print('Siap: izin dan ikon dipasang.')
