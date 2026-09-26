# -*- coding: utf-8 -*-
"""
eta_star_summary.py
eta_star/ フォルダ内の eta_star_v5_*.xlsx から η*(および Re(M)/Im(M))を集計し、
社内で使っている配列形式(D×f の一覧表、周波数ごとの D/a_amp/η*/Re/Im 表)でまとめる。

入力 : ETASTAR_DIR (既定 "eta_star" フォルダ) 内の eta_star_v5_<label>_<f>Hz_<level>_<date>.xlsx
        - box(空ボックス)ラベルは粒径Dが無いため、この集計からは除外する。
        - η*, Re(M), Im(M) はシートのキャッシュ値には頼らず、effective_mass_batch.py と同じ式で
          TimeData から再計算する(生成直後でExcelが未計算のファイルでも正しく読めるようにするため)。
          η* はここでは空ボックス差し引き前の値(そのファイル1本の総質量ベース)。
        - 同じ D・A(=level×10)・f の条件で日付違いなど複数ファイルがある場合は平均値を使う。

出力 : OUTPUT_DIR/eta_star_summary.xlsx
        - シート "D_f_table"  : 添付画像1と同じ形式。行=(D[mm], A[m/s2])、列=f[Hz]、値=η*
        - シート "by_frequency": 添付画像2と同じ形式。f[Hz]ごとに D/a_amp/η*/Re/Im の表を横に並べる
        - シート "raw"        : 集計前の生データ(ファイルごとの1行、検算用)
"""
import re
import glob
import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# ===================== 設定 =====================
INPUT_DIR = r"C:\Users\USER\Desktop\calcREIM"
ETASTAR_DIR = os.path.join(INPUT_DIR, "eta_star")   # eta_star_v5_*.xlsx の場所
OUTPUT_DIR = "./output"
EMPTY_LABELS = ["box"]      # 空ボックスのラベル(この集計では除外)
LEVEL_TO_ACC = 10           # A[m/s^2] ≈ level × LEVEL_TO_ACC (プロジェクト内の既定の対応)
# ===============================================

FNAME_RE = re.compile(
    r"eta_star_v\d+_(?P<label>.+)_(?P<f>\d+(?:\.\d+)?)Hz_(?P<level>[^_.]+)(?:_(?P<date>\d+))?\.xlsx$",
    re.IGNORECASE)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def size_um(label):
    d = re.search(r"(\d+(?:\.\d+)?)", str(label))
    return float(d.group(1)) if d else np.nan


def read_one(path):
    m = FNAME_RE.search(os.path.basename(path))
    if not m:
        print(f"[skip] ファイル名を解釈できない: {os.path.basename(path)}")
        return None
    label = m.group("label")
    if label.lower() in EMPTY_LABELS:
        return None  # 空ボックスはD(粒径)が無いためこの集計では対象外

    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["EtaStar_Acc"]
    f_sheet, m_p, m_c = ws["B6"].value, ws["B7"].value or 0.0, ws["B8"].value or 0.0
    t, F, a = [], [], []
    for row in wb["TimeData"].iter_rows(min_row=2, max_col=8, values_only=True):
        row = list(row) + [None] * (8 - len(row))
        ti, Fi, ai = row[0], row[1], row[3]
        if not num(ti):
            continue
        if not num(ai):
            continue
        if not num(Fi):
            continue
        t.append(ti); F.append(Fi); a.append(ai)
    wb.close()
    if not t:
        print(f"[skip] TimeDataが読めない: {os.path.basename(path)}")
        return None
    t, F, a = map(np.asarray, (t, F, a))
    f = float(f_sheet) if num(f_sheet) else float(m.group("f"))
    e = np.exp(-2j * np.pi * f * t)
    X, Y = np.sum(a * e), np.sum(F * e)
    M = np.conj(X) * Y / abs(X) ** 2
    m_tot = (m_p + m_c) / 1000.0

    level = m.group("level")
    level_num = pd.to_numeric(level, errors="coerce")
    D_um = size_um(label)

    return dict(file=os.path.basename(path), label=label,
                D_mm=D_um / 1000.0 if pd.notna(D_um) else np.nan,
                f=f, level=level, A=level_num * LEVEL_TO_ACC if pd.notna(level_num) else np.nan,
                m_p_g=m_p, m_c_g=m_c,
                ReM_g=M.real * 1000, ImM_g=M.imag * 1000,
                eta=(-2 * np.pi * M.imag / m_tot) if m_tot > 0 else np.nan)


def build_d_f_table(agg):
    """添付画像1形式: 行=(D[mm], A[m/s2])、列=f[Hz]、値=η*"""
    piv = agg.pivot_table(index=["D_mm", "A"], columns="f", values="eta_mean").sort_index()
    piv = piv.reindex(sorted(piv.columns), axis=1)
    return piv


def write_d_f_table_sheet(wb, piv):
    ws = wb.create_sheet("D_f_table")
    freqs = list(piv.columns)
    n_f = len(freqs)

    ws.cell(row=1, column=1, value="D [mm]")
    ws.cell(row=1, column=2, value="A [m/s2]")
    ws.cell(row=1, column=3, value="f [Hz]")
    if n_f > 1:
        ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=2 + n_f)
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    for j, f_val in enumerate(freqs):
        ws.cell(row=2, column=3 + j, value=f_val)

    for i, (idx, row) in enumerate(piv.iterrows(), start=3):
        D_mm, A = idx
        ws.cell(row=i, column=1, value=round(D_mm, 4))
        ws.cell(row=i, column=2, value=A)
        for j, f_val in enumerate(freqs):
            v = row[f_val]
            if pd.notna(v):
                ws.cell(row=i, column=3 + j, value=round(float(v), 4))

    for col in range(1, 3 + n_f):
        ws.column_dimensions[get_column_letter(col)].width = 10


def write_by_frequency_sheet(wb, agg):
    """添付画像2形式: f[Hz]ごとに D/a_amp/η*/Re/Im の表を横に並べる"""
    ws = wb.create_sheet("by_frequency")
    freqs = sorted(agg.f.unique())
    da_pairs = (agg[["D_mm", "A"]].drop_duplicates().sort_values(["D_mm", "A"]).values.tolist())

    total_cols = len(freqs) * 6 - 1  # 5列+区切り1列、最後は区切りなし
    ws.cell(row=1, column=1, value="η* vs 加速度振幅(周波数・粒径別)")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(total_cols, 1))

    for b, f_val in enumerate(freqs):
        c0 = b * 6 + 1  # このブロックの先頭列(D列)
        ws.cell(row=2, column=c0, value=f"f={f_val:g}Hz")
        ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + 4)
        for k, h in enumerate(["D", "a_amp", "η*", "Re", "Im"]):
            ws.cell(row=3, column=c0 + k, value=h)

        sub = agg[agg.f == f_val].set_index(["D_mm", "A"])
        for i, (D_mm, A) in enumerate(da_pairs, start=4):
            ws.cell(row=i, column=c0 + 0, value=round(D_mm, 4))
            ws.cell(row=i, column=c0 + 1, value=A)
            if (D_mm, A) in sub.index:
                r = sub.loc[(D_mm, A)]
                ws.cell(row=i, column=c0 + 2, value=round(float(r.eta_mean), 4))
                ws.cell(row=i, column=c0 + 3, value=round(float(r.ReM_g_mean), 4))
                ws.cell(row=i, column=c0 + 4, value=round(float(r.ImM_g_mean), 4))

        for k in range(5):
            ws.column_dimensions[get_column_letter(c0 + k)].width = 9


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    files = sorted(p for p in glob.glob(os.path.join(ETASTAR_DIR, "*.xlsx"))
                   if not os.path.basename(p).startswith("~$"))
    rows = [r for r in (read_one(p) for p in files) if r]
    if not rows:
        print("対象ファイルなし"); return
    df = pd.DataFrame(rows)

    missing = df[df.D_mm.isna() | df.A.isna()]
    if len(missing):
        print("[warn] D または level(A) が解釈できず集計から除外:")
        print(missing[["file", "label", "level"]].to_string(index=False))
    df = df.dropna(subset=["D_mm", "A"])

    # 同じ D・A・f の条件で日付違いなど複数ファイルがある場合は平均値を使う
    agg = (df.groupby(["D_mm", "A", "f"])
             .agg(eta_mean=("eta", "mean"), ReM_g_mean=("ReM_g", "mean"),
                  ImM_g_mean=("ImM_g", "mean"), n=("eta", "size"))
             .reset_index())
    dup = agg[agg.n > 1]
    if len(dup):
        print(f"[info] 同一条件で複数ファイルがあり平均した件数: {len(dup)}")

    wb = Workbook()
    wb.remove(wb.active)
    write_d_f_table_sheet(wb, build_d_f_table(agg))
    write_by_frequency_sheet(wb, agg)
    ws_raw = wb.create_sheet("raw")
    ws_raw.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws_raw.append(list(row))

    out = os.path.join(OUTPUT_DIR, "eta_star_summary.xlsx")
    wb.save(out)
    print(f"保存: {out}")


if __name__ == "__main__":
    main()
