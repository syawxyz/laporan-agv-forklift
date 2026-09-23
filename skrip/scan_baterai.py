#!/usr/bin/env python3
"""Pindai semua logs/agv_*.csv (paralel): statistik battery_v, arus, jarak, durasi per sesi.
Keluaran: ringkasan_sesi.csv + deret_tegangan_1hz.csv. Bisa dilanjutkan (sesi yang sudah ada dilewati)."""
import csv, glob, os, sys
from multiprocessing import Pool
OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)
KOLOM = ['sesi', 'tanggal', 'baris', 'durasi_s', 'jarak_m', 'n_bat', 'bat_awal_v', 'bat_akhir_v', 'bat_min_v', 'bat_maks_v', 'bat_rata_v',
         'i_rata_a', 'i_maks_a', 'i_rata_gerak_a', 'i_rata_diam_a', 'detik_gerak', 'detik_diam',
         'p_rata_w', 'energi_wh', 'energi_gerak_wh', 'energi_diam_wh', 'detik_otonom', 'sumber_arus']

def fl(s):
    try: return float(s)
    except (ValueError, TypeError): return None

def sesi(f):
    k = os.path.basename(f)[4:19]
    n = 0; t0 = tl = None; dist = 0.0; nb = 0
    b0 = b1 = bmin = bmax = None; bsum = 0.0
    isum = 0.0; imax = 0.0; ni = 0; ig = 0.0; ng = 0; idm = 0.0; ndm = 0
    e = eg = ed = 0.0; auto = 0; last_b = None; sumber_i = ''
    next_sample = 0.0; ts = []
    with open(f, newline='') as fh:
        r = csv.reader(line.replace('\0', '') for line in fh)
        try: hdr = next(r)
        except StopIteration: return None
        ix = {c: i for i, c in enumerate(hdr)}
        if 'battery_v' not in ix: return None
        it, ib, idist, itl, itr, iol, ia, iaw, iel, ier = (ix['t_mono'], ix['battery_v'], ix['odom_dist_m'], ix['trq_l'], ix['trq_r'],
                                                           ix['odom_lin'], ix['auto_lin_x'], ix['auto_ang_z'], ix['wheel_l_eff'], ix['wheel_r_eff'])
        need = max(it, ib, idist, itl, itr, iol, ia, iaw, iel, ier) + 1
        for row in r:
            n += 1
            if len(row) < need: continue
            t = fl(row[it])
            if t is None: continue
            if t0 is None: t0 = t
            dt = (t - tl) if tl is not None else 0.0
            tl = t
            v = fl(row[idist])
            if v is not None: dist = v
            b = fl(row[ib])
            if b is not None and b > 0:
                nb += 1; bsum += b; b1 = b
                if b0 is None: b0 = b
                bmin = b if bmin is None else min(bmin, b); bmax = b if bmax is None else max(bmax, b)
                last_b = b
            i = None
            tlv, trv = fl(row[itl]), fl(row[itr])
            if tlv is not None and trv is not None:
                i = abs(tlv) + abs(trv); sumber_i = sumber_i or 'trq(/motor_torque)'
            else:
                elv, erv = fl(row[iel]), fl(row[ier])
                if elv is not None and erv is not None:
                    i = abs(elv) + abs(erv); sumber_i = sumber_i or 'eff(/joint_states)'
            ol = fl(row[iol]); gerak = ol is not None and abs(ol) > 0.01
            if i is not None:
                isum += i; ni += 1; imax = max(imax, i)
                if gerak: ig += i; ng += 1
                else: idm += i; ndm += 1
                if last_b is not None and 0 < dt < 1.0:
                    de = last_b * i * dt / 3600.0
                    e += de
                    if gerak: eg += de
                    else: ed += de
            al, aw = fl(row[ia]), fl(row[iaw])
            if (al is not None and abs(al) > 1e-6) or (aw is not None and abs(aw) > 1e-6): auto += 1
            if t - t0 >= next_sample:
                ts.append([k, f'{t - t0:.0f}', '' if last_b is None else f'{last_b:.2f}', '' if i is None else f'{i:.2f}', f'{dist:.2f}'])
                next_sample += 1.0
    dur = (tl - t0) if t0 is not None else 0
    row = [k, k[:8], n, f'{dur:.0f}', f'{dist:.1f}', nb,
           *(['' if v is None else f'{v:.2f}' for v in (b0, b1, bmin, bmax)]),
           f'{bsum / nb:.2f}' if nb else '',
           f'{isum / ni:.2f}' if ni else '', f'{imax:.2f}',
           f'{ig / ng:.2f}' if ng else '', f'{idm / ndm:.2f}' if ndm else '',
           f'{ng / 50:.0f}', f'{ndm / 50:.0f}',
           f'{e * 3600 / dur:.1f}' if dur > 0 and e else '', f'{e:.3f}', f'{eg:.3f}', f'{ed:.3f}', f'{auto / 50:.0f}', sumber_i]
    return row, ts

if __name__ == '__main__':
    p_sum = os.path.join(OUT, 'ringkasan_sesi.csv'); p_ts = os.path.join(OUT, 'deret_tegangan_1hz.csv')
    sudah = set()
    if os.path.exists(p_sum):
        sudah = {r['sesi'] for r in csv.DictReader(open(p_sum))}
    files = [f for f in sorted(glob.glob('logs/agv_2026*.csv')) if os.path.basename(f)[4:19] not in sudah]
    files.sort(key=lambda f: -os.path.getsize(f))   # besar duluan supaya beban seimbang
    baru = not sudah
    fs = open(p_sum, 'a', newline=''); ws = csv.writer(fs)
    ft = open(p_ts, 'a', newline=''); wt = csv.writer(ft)
    if baru:
        ws.writerow(KOLOM); wt.writerow(['sesi', 't_s', 'bat_v', 'i_a', 'jarak_m'])
    n_done = 0
    with Pool(int(sys.argv[2]) if len(sys.argv) > 2 else 4) as pool:
        for res in pool.imap_unordered(sesi, files):
            n_done += 1
            if res is None: continue
            row, ts = res
            ws.writerow(row); wt.writerows(ts); fs.flush(); ft.flush()
            print(f'{n_done}/{len(files)} {row[0]}', flush=True)
    print('selesai', flush=True)
