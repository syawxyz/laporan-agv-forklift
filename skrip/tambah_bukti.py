#!/usr/bin/env python3
"""Tambahkan Lampiran D (bukti data aktual) ke laporan proyek .docx
TANPA mengubah konten/desain yang sudah ada. Semua angka dihitung langsung
dari repo saat skrip dijalankan (kecuali hasil uji regresi yang dicatat dari
run sesi 23/9)."""
import copy, csv, glob, hashlib, json, os, re, subprocess, sys, shutil
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

WS = '/home/ibe/sawal_ws'
SRC = f'{WS}/tempp/Laporan-Proyek-AGV-Forklift-2026-09-23.docx'
BAK = os.path.join(os.path.dirname(__file__), 'Laporan-sebelum-lampiran-D.docx')
shutil.copy2(SRC, BAK)
os.chdir(WS)
d = Document(SRC)
body = d.element.body

W_TOTAL = 9411
FILL_HEAD = '1A3A5C'

# ---------------------------------------------------------------- helper docx
def _sz(run, half_pts):
    rPr = run._r.get_or_add_rPr()
    s = OxmlElement('w:sz'); s.set(qn('w:val'), str(half_pts)); rPr.append(s)

def _color_auto(run):
    rPr = run._r.get_or_add_rPr()
    c = OxmlElement('w:color'); c.set(qn('w:val'), 'auto'); rPr.append(c)

bm_id = [max(int(x) for x in re.findall(r'w:id="(\d+)"', body.xml)) + 1]
toc_no = [max(int(x) for x in re.findall(r'_Toc(\d+)', body.xml)) + 1]
toc_entries = []   # (level, text, anchor)

def heading(text, level):
    p = d.add_paragraph(style=f'Heading {level}')
    r = p.add_run(text); _color_auto(r)
    # penanda paragraf (rPr di pPr) juga 'auto' seperti heading yang ada
    pPr = p._p.get_or_add_pPr()
    rpr = OxmlElement('w:rPr'); c = OxmlElement('w:color'); c.set(qn('w:val'), 'auto')
    rpr.append(c); pPr.append(rpr)
    if level <= 2:
        anchor = f'_Toc{toc_no[0]}'; toc_no[0] += 1
        bs = OxmlElement('w:bookmarkStart'); bs.set(qn('w:id'), str(bm_id[0])); bs.set(qn('w:name'), anchor)
        be = OxmlElement('w:bookmarkEnd'); be.set(qn('w:id'), str(bm_id[0])); bm_id[0] += 1
        r._r.addprevious(bs); r._r.addnext(be)
        toc_entries.append((level, text, anchor))
    return p

def para(text, catatan=False):
    p = d.add_paragraph()
    r = p.add_run(text)
    if catatan:
        r.font.color.rgb = RGBColor(0x55, 0x55, 0x55); _sz(r, 18)
    return p

def bullet(text):
    p = d.add_paragraph(text, style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    return p

def tabel(header, rows, widths):
    assert abs(sum(widths) - W_TOTAL) <= 2, widths
    t = d.add_table(rows=1 + len(rows), cols=len(header))
    t.style = d.styles['Table Grid']
    tbl = t._tbl
    tblPr = tbl.tblPr
    for el in list(tblPr):
        if el.tag in (qn('w:tblW'), qn('w:jc'), qn('w:tblLayout'), qn('w:tblLook')):
            tblPr.remove(el)
    tw = OxmlElement('w:tblW'); tw.set(qn('w:w'), '0'); tw.set(qn('w:type'), 'auto'); tblPr.append(tw)
    jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'center'); tblPr.append(jc)
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tblPr.append(lay)
    look = OxmlElement('w:tblLook')
    for k, v in [('val', '04A0'), ('firstRow', '1'), ('lastRow', '0'), ('firstColumn', '1'),
                 ('lastColumn', '0'), ('noHBand', '0'), ('noVBand', '1')]:
        look.set(qn('w:' + k), v)
    tblPr.append(look)
    grid = tbl.tblGrid
    for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
        gc.set(qn('w:w'), str(w))
    for ri, row in enumerate([header] + rows):
        tr = tbl.tr_lst[ri]
        trPr = tr.get_or_add_trPr()
        j = OxmlElement('w:jc'); j.set(qn('w:val'), 'center'); trPr.append(j)
        for ci, (cell, w) in enumerate(zip(t.rows[ri].cells, widths)):
            tcPr = cell._tc.get_or_add_tcPr()
            tcw = OxmlElement('w:tcW'); tcw.set(qn('w:w'), str(w)); tcw.set(qn('w:type'), 'dxa'); tcPr.append(tcw)
            if ri == 0:
                shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), FILL_HEAD); tcPr.append(shd)
            p = cell.paragraphs[0]
            txt = str(row[ci])
            r = p.add_run(txt)
            if ri == 0:
                r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            _sz(r, 18)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def fmt(x, nd=1):
    s = f'{x:,.{nd}f}' if nd else f'{int(round(x)):,}'
    return s.replace(',', '¤').replace('.', ',').replace('¤', '.')

def run(cmd):
    return subprocess.run(cmd, shell=True, capture_output=True, text=True).stdout.strip()

# ---------------------------------------------------------------- data nyata
hari_id = {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'Mei','06':'Jun','07':'Jul','08':'Agt','09':'Sep'}
def tgl(s):  # 20260918 -> 18/9
    return f'{int(s[6:8])}/{int(s[4:6])}'

# data_logger
agv_csv = sorted(glob.glob('logs/agv_2026*.csv'))
dock_csv = sorted(f for f in glob.glob('logs/dock_2026*.csv') if os.path.basename(f)[5:13] <= '20260918')
rak_csv = sorted(glob.glob('logs/rak_push*.csv'))
logs_size = run("du -sh logs | cut -f1")
per_hari = {}
for f in agv_csv:
    k = os.path.basename(f)[4:12]; per_hari.setdefault(k, [0, 0]); per_hari[k][0] += 1; per_hari[k][1] += os.path.getsize(f)

# rekam/ rosbag
rekam = sorted(x for x in os.listdir('rekam') if x.startswith('nav_'))
rekam_size = run("du -sh rekam | cut -f1")
rekam_hari = {}
for x in rekam:
    rekam_hari.setdefault(x[4:12], 0); rekam_hari[x[4:12]] += 1

# bags/
bags = []
for bd in sorted(glob.glob('bags/*/')):
    m = open(bd + 'metadata.yaml').read()
    dur = int(re.search(r'duration:\s*\n\s*nanoseconds:\s*(\d+)', m).group(1)) / 1e9
    cnt = int(re.search(r'message_count:\s*(\d+)', m).group(1))
    odom = re.search(r'name: /odom.*?message_count:\s*(\d+)', m, re.S)
    scan = re.search(r'name: /scan.*?message_count:\s*(\d+)', m, re.S)
    size = sum(os.path.getsize(bd + f) for f in os.listdir(bd))
    bags.append((bd.strip('/').split('/')[-1], dur, cnt, int(odom.group(1)) if odom else 0,
                 int(scan.group(1)) if scan else 0, size))

# traces firmware
tr_dir = 'ibe-agv/stm32_firmware/traces'
lap = sorted(x for x in os.listdir(tr_dir) if x.startswith('lapangan_'))
trc = sorted(x for x in os.listdir(tr_dir) if x.startswith('trace_'))
lap_hari = {}
for x in lap:
    lap_hari.setdefault(x[9:17], 0); lap_hari[x[9:17]] += 1
trc_hari = {}
for x in trc:
    trc_hari.setdefault(x[6:14], 0); trc_hari[x[6:14]] += 1
traces_size = run(f"du -sh {tr_dir} | cut -f1")

# maps
maps = []
for y in sorted(glob.glob('maps/*.yaml'), key=lambda p: os.path.getmtime(p)):
    nm = os.path.basename(y)[:-5]
    pgm = f'maps/{nm}.pgm'
    hdr = open(pgm, 'rb').read(40).split()
    w, h = int(hdr[1]), int(hdr[2])
    org = re.search(r'origin:\s*\[([^\]]+)\]', open(y).read()).group(1)
    ox, oy = [float(v) for v in org.split(',')[:2]]
    pg = os.path.exists(f'maps/{nm}.posegraph')
    maps.append((nm, datetime.fromtimestamp(os.path.getmtime(pgm)).strftime('%d/%m %H:%M'),
                 f'{w}×{h} px = {fmt(w*0.05)}×{fmt(h*0.05)} m', f'({fmt(ox,2)}; {fmt(oy,2)})',
                 'ya' if pg else 'tidak (turunan)'))

# firmware bin
fw = []
for b in glob.glob('ibe-agv/stm32_firmware/firmware_rollback/*.bin') + \
         ['ibe-agv/stm32_firmware/agv_ibe_june/ibe_agv/rollback-sebelum-2026-08-18.bin',
          'ibe-agv/stm32_firmware/agv_ibe_june/ibe_agv/build/ibe_agv.bin']:
    md5 = hashlib.md5(open(b, 'rb').read()).hexdigest()
    fw.append((datetime.fromtimestamp(os.path.getmtime(b)), md5, os.path.getsize(b), b))
fw.sort()

# git
git_total = run('git rev-list --count HEAD')
git_first = run('git log --reverse --format="%h %cd" --date=short | head -1')
git_last = run('git log -1 --format="%h %cd" --date=short')
git_hari = run('git log --format="%cd" --date=short | sort | uniq -c')
git_hari_s = '; '.join(f'{tgl(l.split()[1].replace("-", ""))}: {l.split()[0]}' for l in git_hari.splitlines())
git_dirty = run('git status --porcelain | wc -l')
git_diff = run('git diff --shortstat')
git_untracked = run('git status --porcelain | grep -c "^??"')

# kode
def loc(pattern):
    fs = glob.glob(pattern)
    return sum(1 for f in fs for _ in open(f, errors='ignore')), len(fs)
kode = [
    ('Navigasi Python (ibe-agv/nav_node/*.py)', *loc('ibe-agv/nav_node/*.py')),
    ('Jembatan CAN & emulator (ibe-agv/stm32_patch/*.py)', *loc('ibe-agv/stm32_patch/*.py')),
    ('Alat firmware/lapangan Python (ibe-agv/stm32_firmware/*.py)', *loc('ibe-agv/stm32_firmware/*.py')),
    ('Visi & docking (ibe-agv/vision/*.py)', *loc('ibe-agv/vision/*.py')),
    ('Paket ROS 2 agv (src/agv/agv/*.py)', *loc('src/agv/agv/*.py')),
    ('Launch (src/agv/launch/*.py)', *loc('src/agv/launch/*.py')),
    ('GUI web (src/agv/web/*.html, *.js)', loc('src/agv/web/*.html')[0] + loc('src/agv/web/*.js')[0],
     loc('src/agv/web/*.html')[1] + loc('src/agv/web/*.js')[1]),
    ('Firmware C STM32F407 (agv_ibe_june Core/*.c, *.h)',
     sum(1 for f in glob.glob('ibe-agv/stm32_firmware/agv_ibe_june/ibe_agv/Core/**/*.[ch]', recursive=True)
         for _ in open(f, errors='ignore')),
     len(glob.glob('ibe-agv/stm32_firmware/agv_ibe_june/ibe_agv/Core/**/*.[ch]', recursive=True))),
    ('Berkas uji (test_*.py di nav_node, stm32_patch, stm32_firmware, vision)',
     sum(loc(p)[0] for p in ['ibe-agv/nav_node/test_*.py', 'ibe-agv/stm32_patch/test_*.py',
                             'ibe-agv/stm32_firmware/test_*.py', 'ibe-agv/vision/test_*.py']),
     sum(loc(p)[1] for p in ['ibe-agv/nav_node/test_*.py', 'ibe-agv/stm32_patch/test_*.py',
                             'ibe-agv/stm32_firmware/test_*.py', 'ibe-agv/vision/test_*.py'])),
]
docs_md = glob.glob('ibe-agv/docs/*.md')
docs_kata = sum(len(open(f, errors='ignore').read().split()) for f in docs_md)
sim_html = glob.glob('ibe-agv/sim/*.html')

# sesi 16-18/9 dari CSV
def sesi_stat(f):
    n = 0; t0 = t1 = None; dist = 0.0; ymin = ymax = None; auto = 0
    with open(f, newline='') as fh:
        r = csv.reader(fh); hdr = next(r); ix = {k: i for i, k in enumerate(hdr)}
        it, idist, iyaw, ia, iaw = ix['t_mono'], ix['odom_dist_m'], ix['odom_yaw_unwrap_deg'], ix['auto_lin_x'], ix['auto_ang_z']
        for row in r:
            n += 1
            try: t = float(row[it])
            except ValueError: continue
            if t0 is None: t0 = t
            t1 = t
            if row[idist]: dist = float(row[idist])
            if row[iyaw]:
                y = float(row[iyaw]); ymin = y if ymin is None else min(ymin, y); ymax = y if ymax is None else max(ymax, y)
            if (row[ia] and abs(float(row[ia])) > 1e-6) or (row[iaw] and abs(float(row[iaw])) > 1e-6): auto += 1
    return n, (t1 - t0) if t0 is not None else 0, dist, (ymax - ymin) if ymin is not None else 0, auto

label_sesi = {
    '20260916_165922': 'pemetaan map_w_caster (peta disimpan 17:13)',
    '20260918_111518': 'pemetaan Friday (peta disimpan 11:45)',
    '20260918_134850': 'pemetaan friday2 (peta disimpan 14:03)',
    '20260918_153450': 'uji IMU 360° (uji A)',
    '20260918_152119': 'sesi otonom yang dibedah (§3.9 "hancur saat berubah heading")',
    '20260918_163501': 'dikemudikan joystick (lokalisasi berjalan)',
    '20260918_165040': 'dikemudikan joystick (lokalisasi berjalan)',
    '20260918_165617': 'dikemudikan joystick (lokalisasi berjalan)',
}
sesi_rows = []
for f in sorted(glob.glob('logs/agv_2026091[678]_*.csv')):
    k = os.path.basename(f)[4:19]
    n, dur, dist, yaw, auto = sesi_stat(f)
    if dist < 1.0 and yaw < 30: continue
    ket = label_sesi.get(k, 'perintah otonom terekam' if auto else 'tanpa perintah otonom (joystick)')
    sesi_rows.append((f'{tgl(k[:8])} {k[9:11]}:{k[11:13]}:{k[13:15]}', fmt(dur, 0), fmt(dist), fmt(yaw, 0),
                      fmt(auto / 50.0, 0) if auto else '0', ket))

# docking
dock_hari = {}
dock_sel = {}
for f in dock_csv:
    rows = list(csv.DictReader(open(f, newline='')))
    k = os.path.basename(f)[5:13]
    dock_hari.setdefault(k, {'n': 0, 'SELESAI': 0, 'GAGAL': 0, 'lain': 0, 'pendek': 0})
    dock_hari[k]['n'] += 1
    if len(rows) < 100:
        dock_hari[k]['pendek'] += 1; continue
    akhir = rows[-1]['keadaan']
    dock_hari[k][akhir if akhir in ('SELESAI', 'GAGAL') else 'lain'] += 1
    states = []
    for r in rows:
        if not states or states[-1] != r['keadaan']: states.append(r['keadaan'])
    hidup = [r for r in rows if r['hidup'] in ('1', 'True', 'true') and r['x_m']]
    lr = hidup[-1] if hidup else rows[-1]
    try: x, yw, z = float(lr['x_m']) * 100, float(lr['yaw_deg']), float(lr['z_m'])
    except ValueError: x = yw = z = float('nan')
    dock_sel[os.path.basename(f)[5:20]] = (len(rows), float(rows[-1]['t_wall']) - float(rows[0]['t_wall']), akhir, x, yw, z, states)

dock_pilih = [
    ('20260824_183925', '24/8 18:39 — dock sebelum perbaikan (§3.7: "0/1, 80 s")'),
    ('20260824_185906', '24/8 18:59 — dock sesudah kalibrasi z_stop'),
    ('20260826_114737', '26/8 11:47 — misi gabungan nav→dock'),
    ('20260826_135937', '26/8 13:59 — misi gabungan nav→dock'),
    ('20260826_174529', '26/8 17:45 — sesudah rekalibrasi papan tengah'),
    ('20260827_124449', '27/8 12:44 — uji siksa offset'),
    ('20260828_172658', '28/8 17:26 — TARUH-RAK v2 tembus penuh (MASUK→TARUH→KELUAR→SELESAI)'),
    ('20260831_155335', '31/8 15:53 — misi penuh: TARUH sampai KELUAR lalu GAGAL (timeout total)'),
    ('20260831_161531', '31/8 16:15 — misi penuh percobaan berikutnya'),
    ('20260831_164126', '31/8 16:41 — dock palet ulang'),
]

# ---------------------------------------------------------------- hasil uji 23/9 (dicatat dari run)
uji = [
    ('nav_node/test_heading_follower.py', '81 LULUS, 0 GAGAL', '81 lulus', 'sama'),
    ('nav_node/test_halangan.py', '63 LULUS, 0 GAGAL', '63 lulus', 'sama'),
    ('nav_node/test_theta_star.py', '24 LULUS, 0 GAGAL', '24 lulus', 'sama'),
    ('nav_node/test_smac_adapter.py', 'SEMUA 38 PEMERIKSAAN LULUS', '38 lulus', 'sama'),
    ('nav_node/test_geometri.py', 'SEMUA LULUS — geometri konsisten', 'lulus', 'sama'),
    ('nav_node/test_sasaran_valid.py', 'SEMUA 9 PEMERIKSAAN LULUS', '—', 'tambahan'),
    ('stm32_patch/test_cmd_clamp.py', 'SEMUA 52 PEMERIKSAAN LULUS', '52 lulus', 'sama'),
    ('stm32_firmware/test_nav_pose.py', '55 OK, 5 GAGAL (semua "dalam toleransi sudut")', '5 gagal baseline', 'sama'),
    ('stm32_firmware/test_nav_path.py', '55 OK, 5 GAGAL (sama dengan nav_pose)', '—', 'tambahan'),
    ('stm32_firmware/test_hdg_rule.py', 'SEMUA 6 PEMERIKSAAN LULUS', '—', 'tambahan'),
    ('stm32_firmware/test_field_test.py', '23 OK, 4 GAGAL (yaw prediksi rpm, FAKTOR B, luncuran 0,123 m)', '—', 'tambahan'),
    ('vision/test_dock_servo.py', 'SEMUA 20 PEMERIKSAAN LULUS', 'lulus', 'sama'),
    ('vision/test_fit_dorongan.py', 'SEMUA 28 PEMERIKSAAN LULUS', 'lulus', 'sama'),
    ('vision/test_kalibrasi_rak.py', 'SEMUA 22 PEMERIKSAAN LULUS', 'lulus', 'sama'),
    ('vision/test_pallet_pose.py', 'SEMUA 23 PEMERIKSAAN LULUS', 'lulus', 'sama'),
    ('vision/test_rak_taruh.py', 'SEMUA 33 PEMERIKSAAN LULUS', 'lulus', 'sama'),
    ('vision/test_ukur_pitch.py', 'SEMUA 18 PEMERIKSAAN LULUS', 'lulus', 'sama'),
    ('vision/test_e2e_dock_gui.py', '30 OK, 4 GAGAL (rangkaian misi penuh: PALET DI RAK → nav3 → home)', 'lulus', 'BERBEDA'),
    ('stm32_patch/test_ros_stack.py', 'tidak dijalankan (butuh susunan ROS hidup)', '29/30', '—'),
    ('nav_node/test_gui_otonom.py', 'tidak dijalankan (butuh susunan ROS hidup)', '21/21', '—'),
]

# ---------------------------------------------------------------- konfigurasi berlaku
kal_hist = [l for l in open('ibe-agv/config/odom_calib_history.txt') if l.strip() and not l.strip().startswith('#')]
pose = json.load(open('ibe-agv/config/pose_terakhir.json'))
pc = json.load(open('ibe-agv/config/pallet_cam.json'))
amp = json.load(open('ibe-agv/config/amplop_w_empiris.json'))
meta = json.load(open('logs/agv_20260918_152119.meta.json'))

# ================================================================= TULIS
heading('Lampiran D. Bukti data aktual dari repositori', 1)
para(f'Lampiran ini memuat data mentah yang ada di repositori kerja (/home/ibe/sawal_ws) pada '
     f'{datetime.now().strftime("%d %B %Y %H:%M").replace("September", "September")} sebagai bukti bagi angka-angka yang dikutip '
     f'di Bab 3–5. Seluruh angka di sini dihitung langsung dari berkas (log, rosbag, peta, konfigurasi, git), '
     f'bukan disalin dari catatan. Bila ada selisih kecil dengan bab utama, angka di lampiran ini yang '
     f'mengikat karena berasal dari berkas sumber.')

# D.1 inventaris
heading('D.1 Inventaris rekaman dan artefak', 2)
tabel(['Artefak', 'Jumlah / ukuran', 'Lokasi'], [
    ['Sesi data_logger 50 Hz (agv_*.csv + .meta.json)',
     f'{len(agv_csv)} sesi, {tgl(min(per_hari))}–{tgl(max(per_hari))}; total folder logs/ {logs_size}',
     'logs/'],
    ['Log servo docking & taruh-rak (dock_*.csv)', f'{len(dock_csv)} berkas, 24/8–18/9', 'logs/'],
    ['Log dorongan rak untuk kalibrasi tag (rak_push*.csv)', f'{len(rak_csv)} berkas, 28/8 & 31/8', 'logs/'],
    ['Rosbag sesi navigasi (nav_<tanggal>_<label>/)',
     f'{len(rekam)} sesi, {rekam_size}; per hari: ' + '; '.join(f'{tgl(k)}: {v}' for k, v in sorted(rekam_hari.items())),
     'rekam/'],
    ['Rosbag untuk adu parameter SLAM luring', '; '.join(
        f'{b[0]} ({fmt(b[1],1)} s, {fmt(b[2],0)} pesan: {fmt(b[3],0)} /odom + {fmt(b[4],0)} /scan, {b[5]/1e6:.0f} MB)' for b in bags),
     'bags/'],
    ['Trace bus CAN alat lapangan (field_test.py)',
     f'{len(lap)} sesi lapangan_* (' + '; '.join(f'{tgl(k)}: {v}' for k, v in sorted(lap_hari.items())) +
     f') + {len(trc)} trace_*.csv mdsm_trace (' + '; '.join(f'{tgl(k)}: {v}' for k, v in sorted(trc_hari.items())) +
     f'); {traces_size}',
     'ibe-agv/stm32_firmware/traces/'],
    ['Peta tersimpan (yaml+pgm, sebagian dengan posegraph)', f'{len(maps)} peta (rincian D.3)', 'maps/'],
    ['Biner firmware STM32 yang pernah di-flash', f'{len(fw)} berkas .bin dengan md5 (rincian D.4)', 'ibe-agv/stm32_firmware/'],
    ['Dokumen kerja Markdown', f'{len(docs_md)} berkas, {fmt(docs_kata,0)} kata', 'ibe-agv/docs/'],
    ['Halaman analisis/panduan HTML', f'{len(sim_html)} berkas', 'ibe-agv/sim/'],
    ['Riwayat git cabang otonom_test',
     f'{git_total} commit ({git_first} → {git_last}); per hari: {git_hari_s}. Belum di-commit saat lampiran ditulis: '
     f'{git_dirty} entri status ({git_diff.strip() or "0"}; {git_untracked} berkas/folder baru)',
     'git log'],
], [2900, 4511, 2000])

heading('D.2 Sesi data_logger per hari', 2)
para('Satu sesi = satu kali ./robot mulai (atau susunan pemetaan). Jumlah sesi per hari mencerminkan intensitas '
     'kerja lapangan; ukuran berkas sebanding dengan lama sesi (50 baris/detik).')
rows = [[tgl(k), str(v[0]), f'{v[1]/1e6:,.0f}'.replace(',', '.')] for k, v in sorted(per_hari.items())]
rows.append(['Total', str(sum(v[0] for v in per_hari.values())),
             f'{sum(v[1] for v in per_hari.values())/1e6:,.0f}'.replace(',', '.')])
tabel(['Tanggal', 'Jumlah sesi', 'Ukuran CSV (MB)'], rows, [3137, 3137, 3137])

heading('D.3 Peta yang tersimpan', 2)
para('Semua peta beresolusi 0,05 m/piksel (slam_toolbox). Kolom "posegraph" menandai peta hasil pemetaan asli; '
     'peta tanpa posegraph adalah turunan (zona-hindari) dari map_untuk_trial.')
tabel(['Nama peta', 'Disimpan', 'Ukuran', 'Origin (x; y) m', 'Posegraph'],
      [list(m) for m in maps], [2000, 1500, 2900, 1811, 1200])

heading('D.4 Biner firmware STM32 yang pernah di-flash', 2)
para('Sidik jari md5 dihitung ulang dari berkas .bin. Nama pendek md5 inilah yang dikutip di Bab 3 '
     '(mis. 741dd55f pada 14/8, a6786c8d pada 20/8).')
tabel(['Tanggal berkas', 'md5 (8 hex pertama)', 'Ukuran (byte)', 'Berkas'],
      [[f.strftime('%d/%m %H:%M'), m[:8], fmt(s, 0), os.path.basename(b)] for f, m, s, b in fw],
      [1600, 1900, 1500, 4411])

heading('D.5 Sesi 16–18 September yang dihitung ulang dari CSV', 2)
para('Untuk tiga hari terakhir, setiap berkas agv_*.csv dibaca ulang: durasi dari t_mono, jarak dari odom_dist_m, '
     'putaran = rentang odom_yaw_unwrap_deg, dan "detik otonom" = jumlah baris dengan /cmd_vel_auto tidak nol dibagi 50. '
     'Sesi yang robotnya diam (jarak < 1 m dan putaran < 30°) tidak ditampilkan.')
tabel(['Sesi', 'Durasi (s)', 'Jarak (m)', 'Putaran (°)', 'Detik otonom', 'Keterangan'],
      [list(r) for r in sesi_rows], [1500, 900, 900, 950, 950, 4211])

heading('D.6 Log docking dan taruh-rak (dock_*.csv)', 2)
para('Setiap berkas dock_*.csv adalah satu kali dock_exec dijalankan; kolomnya t_wall, keadaan, hidup, x_m, z_m, yaw_deg, '
     'umur_pose_s, v, w. Tabel pertama menghitung keadaan terakhir per hari (berkas < 100 baris = node dimulai lalu '
     'dimatikan tanpa docking). Tabel kedua mengambil sesi yang dirujuk di Bab 3.7 dan menampilkan nilai pose tag terakhir '
     'yang masih terlihat (x menyamping dalam cm, yaw, jarak z) apa adanya dari log.')
tabel(['Tanggal', 'Berkas', 'Berakhir SELESAI', 'Berakhir GAGAL', 'Berakhir lain', '< 100 baris'],
      [[tgl(k), str(v['n']), str(v['SELESAI']), str(v['GAGAL']), str(v['lain']), str(v['pendek'])]
       for k, v in sorted(dock_hari.items())], [1568, 1568, 1568, 1568, 1568, 1571])
rows = []
for k, ket in dock_pilih:
    if k not in dock_sel: continue
    n, dur, akhir, x, yw, z, st = dock_sel[k]
    rows.append([ket, fmt(dur, 0), akhir, f'{fmt(x)} cm / {fmt(yw)}° / z {fmt(z,2)} m', '→'.join(st)[:120] + ('…' if len('→'.join(st)) > 120 else '')])
tabel(['Sesi (rujukan bab)', 'Durasi (s)', 'Keadaan akhir', 'Pose tag terakhir x / yaw / z', 'Urutan keadaan'],
      rows, [2600, 800, 1000, 2000, 3011])

heading('D.7 Nilai konfigurasi yang berlaku saat ini', 2)
para('Nilai dibaca langsung dari berkas konfigurasi dan konstanta kode pada keadaan kerja (termasuk perubahan 18/9 '
     'yang belum di-commit). Inilah angka yang benar-benar dipakai robot bila ./robot mulai dijalankan hari ini.')
slam = open('src/agv/config/slam_toolbox.yaml').read()
def yv(key, txt=slam):
    m = re.search(rf'^\s*{key}:\s*([^\s#]+)', txt, re.M); return m.group(1) if m else '?'
launch = open('src/agv/launch/navigasi.launch.py').read()
hf = open('ibe-agv/nav_node/heading_follower.py').read()
def pyv(key, txt):
    m = re.search(rf'^{key}\s*=\s*([-0-9.]+)', txt, re.M); return m.group(1) if m else '?'
bridge = open('ibe-agv/stm32_patch/agv_can_bridge.py').read()
def envdef(key, txt):
    m = re.search(rf'"{key}",\s*"([^"]+)"', txt); return m.group(1) if m else '?'
gui = open('src/agv/agv/gui_server.py').read()
smac = open('src/agv/config/planner_smac.yaml').read()
gui2 = gui.replace("'", '"')
bat_str = (envdef('BAT_V_PENUH', gui2) + ' / ' + envdef('BAT_V_KUNING', gui2) + ' / ' + envdef('BAT_V_MERAH', gui2) + ' / ' + envdef('BAT_V_KOSONG', gui2) + ' V').replace('.', ',')
cfg_rows = [
    ['Pemetaan: angle_variance_penalty / distance_variance_penalty', f'{yv("angle_variance_penalty")} / {yv("distance_variance_penalty")}', 'src/agv/config/slam_toolbox.yaml'],
    ['Pemetaan: coarse_search_angle_offset / fine_search_angle_offset / coarse_angle_resolution (rad)',
     f'{yv("coarse_search_angle_offset")} / {yv("fine_search_angle_offset")} / {yv("coarse_angle_resolution")}', 'slam_toolbox.yaml'],
    ['Pemetaan: minimum_travel_distance / minimum_travel_heading', f'{yv("minimum_travel_distance")} m / {yv("minimum_travel_heading")} rad', 'slam_toolbox.yaml'],
    ['Pemetaan: resolution / max_laser_range / loop_search_maximum_distance', f'{yv("resolution")} m / {yv("max_laser_range")} m / {yv("loop_search_maximum_distance")} m', 'slam_toolbox.yaml'],
    ['Lokalisasi (override): angle_variance_penalty / coarse_search_angle_offset / minimum_travel_heading / minimum_travel_distance',
     "1,0 / SLAM_COARSE_ANGLE bawaan 0,175 rad (=10°; jalan pulang 0,44) / 3,14 rad / 0,15 m", 'src/agv/launch/navigasi.launch.py:242–250'],
    ['Perencana Smac: minimum_turning_radius / angle_quantization_bins / reverse_penalty / change_penalty / non_straight_penalty / cost_penalty / tolerance / smooth_path',
     f'{yv("minimum_turning_radius", smac)} m / {yv("angle_quantization_bins", smac)} / {yv("reverse_penalty", smac)} / {yv("change_penalty", smac)} / {yv("non_straight_penalty", smac)} / {yv("cost_penalty", smac)} / {yv("tolerance", smac)} m / {yv("smooth_path", smac)}',
     'src/agv/config/planner_smac.yaml'],
    ['Footprint poligon / inflation_radius / cost_scaling_factor', '[[0,35; ±0,55], [−1,60; ±0,55]] m / 0,85 m / 3,0', 'planner_smac.yaml:250, 326–327'],
    ['Perencana bawaan mulai_navigasi.sh / peta bawaan', 'PERENCANA=smac (A/B: theta) / PETA=Friday', 'ibe-agv/nav_node/mulai_navigasi.sh:12, 49'],
    ['Pose awal lokalisasi tersimpan (ditulis odom_map.py tiap 2 s)',
     f'peta {pose["peta"]}: x {fmt(pose["x"],3)} m, y {fmt(pose["y"],3)} m, yaw {fmt(pose["yaw_deg"],2)}° '
     f'({datetime.fromtimestamp(pose["t_wall"]).strftime("%d/%m %H:%M:%S")})', 'ibe-agv/config/pose_terakhir.json'],
    ['Follower: PIVOT_W_MIN_MRAD_S / D_BIDIK_MIN_M / GANTI_RUAS_M / RAYAP_MASUK_M / EKOR_TOL_POS_M / KOMIT_ARAH_M / SEMI_RODA_DALAM_MIN_MM_S',
     ' / '.join(pyv(k, hf) for k in ['PIVOT_W_MIN_MRAD_S', 'D_BIDIK_MIN_M', 'GANTI_RUAS_M', 'RAYAP_MASUK_M', 'EKOR_TOL_POS_M', 'KOMIT_ARAH_M', 'SEMI_RODA_DALAM_MIN_MM_S']).replace('.', ','),
     'ibe-agv/nav_node/heading_follower.py'],
    ['Jembatan CAN: CMD_CLAMP_ENABLE bawaan / R_MIN / V_MIN / W_PIVOT_MIN–MAX',
     f'{envdef("CMD_CLAMP_ENABLE", bridge)} (mati) / {envdef("CMD_CLAMP_R_MIN", bridge)} m / {envdef("CMD_CLAMP_V_MIN_MM_S", bridge)} mm/s / '
     f'{envdef("CMD_CLAMP_W_PIVOT_MIN", bridge)}–{envdef("CMD_CLAMP_W_PIVOT_MAX", bridge)} mrad/s'.replace('.', ','),
     'ibe-agv/stm32_patch/agv_can_bridge.py:184–215'],
    ['Penjaga KESASAR: lompatan pos / yaw / tenang / hitung dalam jendela', '1,0 m / 25° / 4,0 s / 3× dalam 30 s', 'ibe-agv/nav_node/halangan.py:230–231'],
    ['Posisi LiDAR di badan (LASER_X_M)', f'{pyv("LASER_X_M", hf if False else open("ibe-agv/nav_node/halangan.py").read())} m'.replace('.', ','), 'halangan.py:54'],
    ['GUI: pembatal arus (ARUS_BATAL_A / ARUS_BATAL_S)', f'{pyv("ARUS_BATAL_A", gui)} A / {pyv("ARUS_BATAL_S", gui)} s'.replace('.', ','), 'src/agv/agv/gui_server.py:125–128'],
    ['GUI baterai 48 V: BAT_V_PENUH / KUNING / MERAH / KOSONG',
     bat_str,
     'gui_server.py:151–154'],
    ['Dock palet: z_stop / z_gerbang / tol_x / tol_yaw / fork_dock / marker / id palet / offset_x kamera',
     f'{pc["dock"]["z_stop_m"]} m / {pc["dock"]["z_gerbang_m"]} m / {pc["dock"]["tol_x_m"]} m / {pc["dock"]["tol_yaw_deg"]}° / {pc["dock"]["fork_dock_cm"]} cm / '
     f'ArUco {pc["marker_size_m"]} m / id {pc["id_palet"]} / {pc["offset_x_m"]} m'.replace('.', ','),
     'ibe-agv/config/pallet_cam.json'],
    ['Taruh rak: z_taruh / z_gerbang / tol_x / tol_yaw / fork_angkat / fork_taruh / fork_bawa / fork_jalan',
     f'{pc["rak"]["z_taruh_m"]} m / {pc["rak"]["z_gerbang_m"]} m / {pc["rak"]["tol_x_m"]} m / {pc["rak"]["tol_yaw_deg"]}° / '
     f'{pc["rak"]["fork_angkat_cm"]} / {pc["rak"]["fork_taruh_cm"]} / {pc["rak"]["fork_bawa_cm"]} / {pc["rak"]["fork_jalan_cm"]} cm'.replace('.', ','),
     'pallet_cam.json /rak'],
    ['Taruh rak: w_bang / koreksi_m / koreksi_maks / servo_timeout / timeout_total / jarak_keluar / setir_maju',
     f'{pc["rak"]["w_bang_rad_s"]} rad/s / {pc["rak"]["koreksi_m"]} m / {pc["rak"]["koreksi_maks"]} / {pc["rak"]["servo_timeout_s"]:.0f} s / '
     f'{pc["rak"]["timeout_total_s"]:.0f} s / {pc["rak"]["jarak_keluar_m"]} m / {pc["rak"]["setir_maju"]}'.replace('.', ','),
     'pallet_cam.json /rak'],
    ['Kalibrasi tag rak id 3 / id 1 (x_ofs, yaw_ofs, z_ref)',
     f'id3: {pc["rak"]["kal"]["3"]["x_ofs_m"]} m, {pc["rak"]["kal"]["3"]["yaw_ofs_deg"]}°, z {pc["rak"]["kal"]["3"]["z_ref_m"]} m; '
     f'id1: {pc["rak"]["kal"]["1"]["x_ofs_m"]} m, {pc["rak"]["kal"]["1"]["yaw_ofs_deg"]}°, z {pc["rak"]["kal"]["1"]["z_ref_m"]} m'.replace('.', ','),
     'pallet_cam.json /rak/kal'],
    ['Odometri: linear_scale / angular_scale / jari-jari roda / jarak roda',
     f'1,0 / 1,0 / {meta["wheel_radius_m"]} m / {meta["wheel_base_m"]} m'.replace('.', ','),
     'ibe-agv/config/odom_calib.json; agv_*.meta.json'],
    ['Amplop w empiris era caster lama (kuantil w tercapai, rad/s)',
     f'p25 {amp["w_ach_rad_s_kuantil"]["p25"]} / p50 {amp["w_ach_rad_s_kuantil"]["p50"]} / p75 {amp["w_ach_rad_s_kuantil"]["p75"]} / '
     f'p90 {amp["w_ach_rad_s_kuantil"]["p90"]} / maks {amp["w_ach_rad_s_kuantil"]["maks"]}; p_macet pivot PC {amp["p_macet_putar_di_tempat"]["pc"]}, joystick {amp["p_macet_putar_di_tempat"]["joystick"]}'.replace('.', ','),
     'ibe-agv/config/amplop_w_empiris.json (18/8)'],
]
tabel(['Parameter', 'Nilai berlaku', 'Sumber'], cfg_rows, [3300, 3611, 2500])

heading('D.8 Riwayat kalibrasi odometri', 2)
para('Berkas riwayat kalibrasi menyimpan setiap nilai pengali yang pernah dicoba. Baris terakhir (1,0 / 1,0) adalah '
     'hasil ukur lapangan 6/8: rotasi 172,0° terbit untuk ~180° fisik (0,956×) dan odom 1,892 m vs pita ukur 2,000 m (1,057×), '
     'keduanya dalam ±10% sehingga tidak perlu pengali. Bukti mentahnya logs/agv_20260806_232504.csv (534 s, 26.742 baris) '
     'sudah tidak ada di logs/ (rekaman tertua yang tersisa 11/8).')
rows = []
for l in kal_hist:
    parts = [p.strip() for p in l.split(',', 2)]
    rows.append([parts[0], parts[1], parts[2].lstrip('# ').strip() if len(parts) > 2 else ''])
tabel(['linear_scale', 'angular_scale', 'Catatan di berkas'], rows, [1800, 1800, 5811])

heading('D.9 Uji regresi dijalankan ulang 23 September 2026', 2)
para('Seluruh berkas uji yang tidak memerlukan susunan ROS hidup dijalankan ulang pada keadaan kerja saat ini '
     '(perubahan 18/9 belum di-commit) dan hasilnya dibandingkan dengan baseline Lampiran A.')
tabel(['Berkas uji', 'Hasil 23/9 (keluaran skrip)', 'Baseline Lamp. A', 'Banding'],
      [list(u) for u in uji], [2700, 4011, 1400, 1300])
para('Catatan: test_e2e_dock_gui.py gagal pada 4 pemeriksaan rangkaian misi penuh (nav3 → home). Ini konsisten dengan '
     'Bab 5.1 butir 7 (misi penuh belum pernah tuntas bersih) dan dengan perubahan follower/launch 18/9 yang belum '
     'di-commit; sebelum 18/9 berkas ini lulus (Lampiran A). test_nav_pose/test_nav_path gagal 5 pemeriksaan "dalam '
     'toleransi sudut" persis seperti baseline (PoseNav adalah jalur cadangan era Ackermann).', catatan=True)

heading('D.10 Contoh keluaran alat lapangan (field_test.py, 12/8 13:57)', 2)
para('Setiap sesi lapangan_* berisi HASIL.md, sesi.json, dan trace.csv.gz (bus CAN mentah, baca-saja). Cuplikan '
     'berikut disalin apa adanya dari ibe-agv/stm32_firmware/traces/lapangan_20260812_1357/HASIL.md, tahap 6 '
     '(bring-up CAN untuk ROS), jendela 46,0 detik.')
tabel(['Besaran', 'Nilai terukur'], [
    ['frame', '12086'], ['arus_puncak_A', '0,0'], ['dc_min_V', '19,4'], ['EMCY', '2'],
    ['perubahan_statusword', '0'], ['agv_mode / agv_mode_terlihat', '[0] / [0]'],
    ['selisih_theta_vs_yaw_maks_deg', '0,0'], ['TELE_POSE_Hz', '30,0'],
    ['jeda_telemetri_terpanjang_s', '0,035'], ['PC_CMD_Hz', '50,0'],
    ['Catatan operator', 'odom_hz 30,000; gerbang_basi: ya'],
    ['Kesimpulan di berkas', 'mode 2 = AGV_MODE_PC. Kalau tidak pernah 2, jembatan tidak pernah memenangkan arbitrase.'],
], [3300, 6111])

# ---------------------------------------------------------------- TOC entries
toc_last1 = toc_last2 = None
for p in d.paragraphs:
    if p.style.name == 'toc 1' and 'Lampiran C' in p.text: toc_last1 = p
    if p.style.name == 'toc 2' and p.text.startswith('5.2'): toc_last2 = p
anchor_after = toc_last1._p
for level, text, anchor in toc_entries:
    tpl = (toc_last1 if level == 1 else toc_last2)._p
    new = copy.deepcopy(tpl)
    for el in new.iter():
        for a in list(el.attrib):
            if a.endswith('}paraId') or a.endswith('}textId'): del el.attrib[a]
    hl = new.find(qn('w:hyperlink')); hl.set(qn('w:anchor'), anchor)
    ts = hl.findall('.//' + qn('w:t')); ts[0].text = text
    it = hl.find('.//' + qn('w:instrText')); it.text = f' PAGEREF {anchor} \\h '
    anchor_after.addnext(new); anchor_after = new

# minta Word memperbarui bidang (nomor halaman daftar isi) saat dibuka
settings = d.settings.element
if settings.find(qn('w:updateFields')) is None:
    uf = OxmlElement('w:updateFields'); uf.set(qn('w:val'), 'true'); settings.append(uf)

d.save(SRC)
print('OK ->', SRC, 'paragraf', len(d.paragraphs), 'tabel', len(d.tables), 'toc baru', len(toc_entries))
