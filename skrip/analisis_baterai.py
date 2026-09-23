#!/usr/bin/env python3
"""Analisis konsumsi daya baterai AGV dari ringkasan pindaian logs/agv_*.csv.
Masukan : <dir_pindai>/ringkasan_sesi.csv, deret_tegangan_1hz.csv
Keluaran: <dir_keluar>/ tabel CSV, grafik PNG, hasil.json (angka kunci utk docx)."""
import csv, json, os, sys, math
from collections import OrderedDict
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

SRC, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT, exist_ok=True)
plt.rcParams.update({'font.size': 9, 'axes.spines.top': False, 'axes.spines.right': False,
                     'figure.dpi': 150, 'savefig.dpi': 150})
C1, C2, C3, C4 = '#1A3A5C', '#B3541E', '#2C5578', '#8A8A8A'

def fnum(s):
    try: return float(s)
    except (TypeError, ValueError): return None

def tgl(k): return f'{int(k[6:8])}/{int(k[4:6])}'

rows = sorted(csv.DictReader(open(os.path.join(SRC, 'ringkasan_sesi.csv'))), key=lambda r: r['sesi'])
# deret 1 Hz dibaca dulu: dipakai utk V-min sah (>= 40 V) dan V vs I
V = []; I = []; T = {}
with open(os.path.join(SRC, 'deret_tegangan_1hz.csv')) as f:
    for r in csv.DictReader(f):
        if r['bat_v'] and r['i_a']:
            v, i = float(r['bat_v']), float(r['i_a'])
            if v >= 40: V.append(v); I.append(i)
        T.setdefault(r['sesi'], []).append((float(r['t_s']), fnum(r['bat_v']), fnum(r['i_a']), fnum(r['jarak_m'])))
VMIN40 = {k: min((a[1] for a in d if a[1] is not None and a[1] >= 40), default=None) for k, d in T.items()}
VFIRST40 = {k: next((a[1] for a in d if a[1] is not None and a[1] >= 40), None) for k, d in T.items()}
VLAST40 = {k: next((a[1] for a in reversed(d) if a[1] is not None and a[1] >= 40), None) for k, d in T.items()}
NLOW = {k: sum(1 for a in d if a[1] is not None and 0 < a[1] < 40) for k, d in T.items()}
n_sesi_low = sum(1 for k, v in NLOW.items() if v > 0)
for r in rows:
    for k in list(r):
        if k not in ('sesi', 'tanggal', 'sumber_arus'): r[k] = fnum(r[k])
    # jarak odometri sah: rata-rata <= 0,7 m/s (plafon fisik 0,64) dan < 2 km per sesi; selebihnya lompatan odom
    r['jarak_sah'] = (r['jarak_m'] or 0) if (r['durasi_s'] or 0) > 0 and (r['jarak_m'] or 0) / max(r['durasi_s'], 1) <= 0.7 and (r['jarak_m'] or 0) < 2000 else 0.0
    r['jarak_asli'] = r['jarak_m']; r['jarak_m'] = r['jarak_sah']
    if VMIN40.get(r['sesi']) is not None: r['bat_min_v'] = VMIN40[r['sesi']]
    if VFIRST40.get(r['sesi']) is not None: r['bat_awal_v'] = VFIRST40[r['sesi']]
    if VLAST40.get(r['sesi']) is not None: r['bat_akhir_v'] = VLAST40[r['sesi']]
    r['detik_dclink_rendah'] = NLOW.get(r['sesi'], 0)
n_jarak_buang = sum(1 for r in rows if (r['jarak_asli'] or 0) > 0 and r['jarak_sah'] == 0)
sesi_v = [r for r in rows if r['n_bat'] and r['n_bat'] > 0 and r['bat_min_v'] is not None and r['bat_min_v'] >= 40]
sesi_i = [r for r in sesi_v if r['i_rata_a'] is not None and r['durasi_s'] >= 30]

# ---------------------------------------------------------------- per hari
hari = OrderedDict()
for r in sesi_v:
    h = hari.setdefault(r['tanggal'], {'n': 0, 'n_i': 0, 'v_awal': r['bat_awal_v'], 'v_akhir': r['bat_akhir_v'],
                                       'v_min': r['bat_min_v'], 'v_maks': r['bat_maks_v'], 'jarak': 0.0, 'durasi': 0.0,
                                       'gerak': 0.0, 'wh': 0.0, 'wh_gerak': 0.0, 'wh_diam': 0.0, 'jarak_i': 0.0,
                                       't_awal': r['sesi'][9:13], 't_akhir': r['sesi'][9:13]})
    h['n'] += 1; h['jarak'] += r['jarak_m'] or 0; h['durasi'] += r['durasi_s'] or 0
    h['v_akhir'] = r['bat_akhir_v']; h['t_akhir'] = r['sesi'][9:13]
    h['v_min'] = min(h['v_min'], r['bat_min_v']); h['v_maks'] = max(h['v_maks'], r['bat_maks_v'])
    if r['i_rata_a'] is not None:
        h['n_i'] += 1; h['gerak'] += r['detik_gerak'] or 0; h['wh'] += r['energi_wh'] or 0
        h['wh_gerak'] += r['energi_gerak_wh'] or 0; h['wh_diam'] += r['energi_diam_wh'] or 0
        h['jarak_i'] += r['jarak_m'] or 0
with open(os.path.join(OUT, 'per_hari.csv'), 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['tanggal', 'sesi', 'sesi_dgn_arus', 'jam_awal', 'jam_akhir', 'v_awal', 'v_akhir', 'v_min', 'v_maks', 'selisih_v',
                'jam_tercatat', 'jarak_m', 'menit_gerak', 'wh_traksi', 'wh_gerak', 'wh_diam', 'wh_per_m'])
    for k, h in hari.items():
        w.writerow([k, h['n'], h['n_i'], h['t_awal'][:2] + ':' + h['t_awal'][2:], h['t_akhir'][:2] + ':' + h['t_akhir'][2:],
                    f"{h['v_awal']:.1f}", f"{h['v_akhir']:.1f}", f"{h['v_min']:.1f}", f"{h['v_maks']:.1f}",
                    f"{h['v_awal'] - h['v_akhir']:+.1f}", f"{h['durasi'] / 3600:.2f}", f"{h['jarak']:.0f}",
                    f"{h['gerak'] / 60:.1f}", f"{h['wh']:.1f}", f"{h['wh_gerak']:.1f}", f"{h['wh_diam']:.1f}",
                    f"{h['wh_gerak'] / h['jarak_i']:.2f}" if h['jarak_i'] > 5 else ''])

# ---------------------------------------------------------------- per sesi (yang bermakna)
sesi_utama = [r for r in sesi_i if (r['jarak_m'] or 0) >= 5.0]
sesi_utama.sort(key=lambda r: -(r['energi_wh'] or 0))
with open(os.path.join(OUT, 'per_sesi.csv'), 'w', newline='') as f:
    w = csv.writer(f)
    w.writerow(['sesi', 'durasi_s', 'jarak_m', 'v_awal', 'v_akhir', 'v_min', 'i_rata_gerak_a', 'i_rata_diam_a', 'i_maks_a',
                'wh_total', 'wh_gerak', 'wh_diam', 'wh_per_m_gerak', 'w_rata_gerak', 'detik_otonom', 'sumber_arus'])
    for r in sorted(sesi_i, key=lambda r: r['sesi']):
        w.writerow([r['sesi'], f"{r['durasi_s']:.0f}", f"{r['jarak_m']:.1f}", f"{r['bat_awal_v']:.1f}", f"{r['bat_akhir_v']:.1f}",
                    f"{r['bat_min_v']:.1f}", r['i_rata_gerak_a'] if r['i_rata_gerak_a'] is not None else '',
                    r['i_rata_diam_a'] if r['i_rata_diam_a'] is not None else '', r['i_maks_a'],
                    f"{r['energi_wh']:.2f}", f"{r['energi_gerak_wh']:.2f}", f"{r['energi_diam_wh']:.2f}",
                    f"{r['energi_gerak_wh'] / r['jarak_m']:.3f}" if (r['jarak_m'] or 0) >= 5 else '',
                    f"{r['energi_gerak_wh'] * 3600 / r['detik_gerak']:.0f}" if (r['detik_gerak'] or 0) >= 10 else '',
                    f"{r['detik_otonom']:.0f}", r['sumber_arus']])

# ---------------------------------------------------------------- agregat
tot_jarak = sum(r['jarak_m'] or 0 for r in sesi_v)
tot_jam = sum(r['durasi_s'] or 0 for r in sesi_v) / 3600
jarak_i = sum(r['jarak_m'] or 0 for r in sesi_i)
wh_gerak = sum(r['energi_gerak_wh'] or 0 for r in sesi_i)
wh_diam = sum(r['energi_diam_wh'] or 0 for r in sesi_i)
det_gerak = sum(r['detik_gerak'] or 0 for r in sesi_i)
det_diam = sum(r['detik_diam'] or 0 for r in sesi_i)
wh_per_m = [r['energi_gerak_wh'] / r['jarak_m'] for r in sesi_utama if r['jarak_m'] >= 10 and r['energi_gerak_wh'] > 0]
w_gerak = [r['energi_gerak_wh'] * 3600 / r['detik_gerak'] for r in sesi_utama if (r['detik_gerak'] or 0) >= 30]
i_gerak = [r['i_rata_gerak_a'] for r in sesi_utama if r['i_rata_gerak_a'] is not None]
i_diam = [r['i_rata_diam_a'] for r in sesi_i if r['i_rata_diam_a'] is not None and (r['detik_diam'] or 0) >= 30]
i_maks = max(r['i_maks_a'] for r in sesi_i)
sesi_imaks = max(sesi_i, key=lambda r: r['i_maks_a'])['sesi']

# otonom vs joystick (sesi bergerak >= 10 m)
oto = [r for r in sesi_utama if r['jarak_m'] >= 10 and (r['detik_otonom'] or 0) > 0.5 * (r['detik_gerak'] or 1)]
joy = [r for r in sesi_utama if r['jarak_m'] >= 10 and (r['detik_otonom'] or 0) == 0]
def med(xs): return float(np.median(xs)) if xs else None
banding = {
    'otonom': {'n': len(oto), 'wh_per_m': med([r['energi_gerak_wh'] / r['jarak_m'] for r in oto]),
               'i_gerak': med([r['i_rata_gerak_a'] for r in oto if r['i_rata_gerak_a'] is not None]),
               'v_ms': med([r['jarak_m'] / r['detik_gerak'] for r in oto if (r['detik_gerak'] or 0) > 0])},
    'joystick': {'n': len(joy), 'wh_per_m': med([r['energi_gerak_wh'] / r['jarak_m'] for r in joy]),
                 'i_gerak': med([r['i_rata_gerak_a'] for r in joy if r['i_rata_gerak_a'] is not None]),
                 'v_ms': med([r['jarak_m'] / r['detik_gerak'] for r in joy if (r['detik_gerak'] or 0) > 0])},
}

# ---------------------------------------------------------------- V vs I (deret 1 Hz)
V = np.array(V); I = np.array(I)
sag = None
if len(V) > 100:
    A = np.vstack([np.ones_like(I), I]).T
    coef, *_ = np.linalg.lstsq(A, V, rcond=None)
    v0, slope = coef
    # binned untuk tampilan
    bins = np.arange(0, min(I.max(), 40) + 2, 2)
    idx = np.digitize(I, bins)
    bin_v = [(bins[j - 1] + 1, float(np.median(V[idx == j])), int((idx == j).sum())) for j in range(1, len(bins)) if (idx == j).sum() >= 20]
    sag = {'v0': float(v0), 'ohm': float(-slope), 'n': int(len(V)), 'r': float(np.corrcoef(I, V)[0, 1]), 'bin': bin_v}

# ---------------------------------------------------------------- grafik
# 1. Tegangan awal/akhir/min per hari
ks = list(hari)
x = np.arange(len(ks))
fig, ax = plt.subplots(figsize=(7.2, 3.0))
ax.plot(x, [hari[k]['v_awal'] for k in ks], 'o-', color=C1, label='awal hari')
ax.plot(x, [hari[k]['v_akhir'] for k in ks], 's-', color=C2, label='akhir hari')
ax.plot(x, [hari[k]['v_min'] for k in ks], 'v--', color=C4, label='terendah')
ax.set_xticks(x); ax.set_xticklabels([tgl(k) for k in ks], rotation=45)
ax.set_ylabel('Tegangan DC-link (V)'); ax.grid(alpha=0.3); ax.legend(loc='lower left', ncol=3, frameon=False)
ax.set_title('Tegangan baterai (DC-link drive) per hari kerja')
fig.tight_layout(); fig.savefig(os.path.join(OUT, 'g1_tegangan_per_hari.png')); plt.close(fig)

# 2. Deret tegangan+arus hari 18/9 (sesi berurutan disambung)
def plot_hari(tanggal, nama):
    ss = sorted(k for k in T if k.startswith(tanggal))
    if not ss: return
    fig, ax1 = plt.subplots(figsize=(7.2, 3.2)); ax2 = ax1.twinx()
    t_ofs = 0.0; ticks = []
    for k in ss:
        d = T[k]
        if len(d) < 30: continue
        t = np.array([a[0] for a in d]) + t_ofs
        v = np.array([a[1] if a[1] is not None else np.nan for a in d])
        i = np.array([a[2] if a[2] is not None else np.nan for a in d])
        ax1.plot(t / 60, v, color=C1, lw=0.8)
        ax2.plot(t / 60, i, color=C2, lw=0.4, alpha=0.7)
        ticks.append((t_ofs / 60, k[9:11] + ':' + k[11:13]))
        t_ofs += d[-1][0] + 60
    ax1.set_ylim(44, 52); ax1.set_ylabel('Tegangan DC-link (V)', color=C1); ax2.set_ylabel('Arus traksi L+R (A)', color=C2)
    ax1.set_xlabel('menit (sesi disambung, jeda 1 menit antar sesi)')
    for tx, lab in ticks: ax1.axvline(tx, color=C4, lw=0.5, ls=':')
    ax1.set_xticks([a for a, _ in ticks]); ax1.set_xticklabels([b for _, b in ticks], rotation=60, fontsize=7)
    ax1.set_title(f'Tegangan dan arus sepanjang {tgl(tanggal)} (mulai sesi ditandai jam)')
    ax1.grid(alpha=0.3); fig.tight_layout(); fig.savefig(os.path.join(OUT, nama)); plt.close(fig)
plot_hari('20260918', 'g2_deret_18sep.png')
plot_hari('20260821', 'g2b_deret_21agt.png')

# 3. V vs I
if sag:
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    sel = np.random.default_rng(0).choice(len(I), size=min(20000, len(I)), replace=False)
    ax.scatter(I[sel], V[sel], s=2, alpha=0.15, color=C3, label='sampel 1 Hz')
    xs = np.linspace(0, min(I.max(), 40), 50)
    ax.plot(xs, sag['v0'] - sag['ohm'] * xs, color=C2, lw=1.5, label=f"garis lurus: V = {sag['v0']:.2f} − {sag['ohm']:.3f}·I")
    ax.plot([b[0] for b in sag['bin']], [b[1] for b in sag['bin']], 'ko', ms=3, label='median per 2 A')
    ax.set_xlabel('Arus traksi L+R (A)'); ax.set_ylabel('Tegangan DC-link (V)'); ax.set_xlim(0, min(I.max(), 40))
    ax.set_ylim(max(V.min(), 40), V.max() + 0.5); ax.grid(alpha=0.3); ax.legend(frameon=False, fontsize=8)
    ax.set_title('Jatuh tegangan terhadap arus (seluruh sesi dengan data arus)')
    fig.tight_layout(); fig.savefig(os.path.join(OUT, 'g3_v_vs_i.png')); plt.close(fig)

# 4. Wh per meter per sesi (>= 10 m)
su = [r for r in sesi_utama if r['jarak_m'] >= 10]
su.sort(key=lambda r: r['sesi'])
if su:
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    cols = [C1 if (r['detik_otonom'] or 0) > 0.5 * (r['detik_gerak'] or 1) else C4 for r in su]
    ax.bar(range(len(su)), [r['energi_gerak_wh'] / r['jarak_m'] for r in su], color=cols)
    # label: satu per tanggal (di bar pertama tanggal itu), pemisah tanggal garis tipis
    tk = []; lb = []; prev = None
    for j, r in enumerate(su):
        if r['sesi'][:8] != prev:
            tk.append(j); lb.append(tgl(r['sesi'][:8])); prev = r['sesi'][:8]
            if j: ax.axvline(j - 0.5, color=C4, lw=0.5, ls=':')
    ax.set_xticks(tk); ax.set_xticklabels(lb, rotation=60, fontsize=7); ax.set_xlabel('sesi berurutan, dikelompokkan per tanggal')
    ax.set_ylabel('Wh per meter (saat bergerak)'); ax.grid(axis='y', alpha=0.3)
    ax.set_title('Energi traksi per meter tiap sesi ≥ 10 m (biru = dominan otonom, abu = joystick)')
    fig.tight_layout(); fig.savefig(os.path.join(OUT, 'g4_wh_per_m.png')); plt.close(fig)

hasil = {
    'n_sesi_total': len(rows), 'n_sesi_tegangan': len(sesi_v), 'n_sesi_arus': len(sesi_i),
    'n_sesi_arus_trq': sum(1 for r in sesi_i if r['sumber_arus'].startswith('trq')),
    'n_sesi_arus_eff': sum(1 for r in sesi_i if r['sumber_arus'].startswith('eff')),
    'hari_pertama': min(hari), 'hari_terakhir': max(hari), 'n_hari': len(hari),
    'v_tertinggi': max(r['bat_maks_v'] for r in sesi_v), 'v_terendah': min(r['bat_min_v'] for r in sesi_v),
    'sesi_v_terendah': min(sesi_v, key=lambda r: r['bat_min_v'])['sesi'],
    'n_sesi_dclink_rendah': n_sesi_low, 'n_jarak_buang': n_jarak_buang,
    'tot_jarak_m': tot_jarak, 'tot_jam_log': tot_jam, 'jarak_dgn_arus_m': jarak_i,
    'wh_gerak': wh_gerak, 'wh_diam': wh_diam, 'jam_gerak': det_gerak / 3600, 'jam_diam': det_diam / 3600,
    'wh_per_m_median': med(wh_per_m), 'wh_per_m_p25': float(np.percentile(wh_per_m, 25)) if wh_per_m else None,
    'wh_per_m_p75': float(np.percentile(wh_per_m, 75)) if wh_per_m else None,
    'w_gerak_median': med(w_gerak), 'i_gerak_median': med(i_gerak), 'i_diam_median': med(i_diam),
    'i_maks': i_maks, 'sesi_i_maks': sesi_imaks,
    'w_diam_median': (med(i_diam) or 0) * 48.0,
    'banding': banding, 'sag': sag,
    'per_hari': {k: hari[k] for k in hari},
    'sesi_utama': sesi_utama[:12],
}
json.dump(hasil, open(os.path.join(OUT, 'hasil.json'), 'w'), indent=1, ensure_ascii=False, default=float)
print(json.dumps({k: v for k, v in hasil.items() if k not in ('per_hari', 'sesi_utama', 'sag')}, indent=1, ensure_ascii=False, default=float))
print('sag', {k: v for k, v in (sag or {}).items() if k != 'bin'})
