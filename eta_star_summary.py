# -*- coding: utf-8 -*-
"""
eta_star_summary.py
eta_star/ フォルダ内の eta_star_v5_*.xlsx から η*(および Re(M)/Im(M))を集計し、
社内で使っている配列形式(D×f の一覧表、周波数ごとの D/a_amp/η*/Re/Im 表)でまとめる。

入力 : ETASTAR_DIR (既定 "data_0927/eta_star" フォルダ) 内の eta_star_v5_<label>_<f>Hz_<level>_<date>[-N].xlsx
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
DATA_DIR = os.path.join(INPUT_DIR, "data_0927")     # 運用中のデータ一式(Rawdata/eta_star/output)
ETASTAR_DIR = os.path.join(DATA_DIR, "eta_star")    # eta_star_v5_*.xlsx はここに格納
OUTPUT_DIR = os.path.join(DATA_DIR, "output")
EMPTY_LABELS = ["box"]      # 空ボックスのラベル(この集計では除外)
LEVEL_TO_ACC = 10           # A[m/s^2] ≈ level × LEVEL_TO_ACC (プロジェクト内の既定の対応)
# ===============================================

FNAME_RE = re.compile(
    r"eta_star_v\d+_(?P<label>.+)_(?P<f>\d+(?:\.\d+)?)Hz_(?P<level>[^_.]+)(?:_(?P<date>\d+(?:-\d+)?))?\.xlsx$",
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
    is_box = label.lower() in EMPTY_LABELS  # 空ボックスはD(粒径)が無いので表では D="box" として別扱い

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

    return dict(file=os.path.basename(path), label=label, is_box=is_box,
                D_mm=D_um / 1000.0 if pd.notna(D_um) and not is_box else np.nan,
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


RUN_RE = re.compile(r"Hz_[^_.]+_(?P<run>\d+(?:-\d+)?)\.xlsx$", re.IGNORECASE)


def run_tag(fname):
    """ファイル名末尾の 日付[-繰り返し番号] (例: '0527', '0527-2')。日付なしのファイルは空文字。"""
    m = RUN_RE.search(fname)
    return m.group("run") if m else ""


def run_order(tag):
    """測定の並び順: 日付 → 繰り返し番号(接尾辞なし=1回目)"""
    date, _, rep = tag.partition("-")
    return (date, int(rep) if rep else 1)


def write_d_f_table_all_sheet(wb, df):
    """D_f_table と同じ形式で平均しない版: 行=(D[mm], A[m/s2], 測定=日付[-繰り返し番号])、列=f[Hz]、値=η*。
    同じ測定回(同じ日付・繰り返し番号)の値が1行に並ぶ。"""
    d = df.assign(run=df.file.map(run_tag))
    dup = d[d.duplicated(["D_mm", "A", "run", "f"], keep=False)]
    if len(dup):
        print("[warn] D_f_table_all: 同じ D/A/測定/f のファイルが複数あり平均した:")
        print(dup[["file"]].to_string(index=False))
    piv = d.pivot_table(index=["D_mm", "A", "run"], columns="f", values="eta")
    piv = piv.reindex(sorted(piv.columns), axis=1)
    piv = piv.iloc[sorted(range(len(piv)), key=lambda i: (piv.index[i][0], piv.index[i][1],
                                                           run_order(piv.index[i][2])))]

    ws = wb.create_sheet("D_f_table_all")
    freqs = list(piv.columns)
    n_f = len(freqs)
    ws.cell(row=1, column=1, value="D [mm]")
    ws.cell(row=1, column=2, value="A [m/s2]")
    ws.cell(row=1, column=3, value="測定")
    ws.cell(row=1, column=4, value="f [Hz]")
    if n_f > 1:
        ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=3 + n_f)
    for col in (1, 2, 3):
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
    for j, f_val in enumerate(freqs):
        ws.cell(row=2, column=4 + j, value=f_val)

    for i, ((D_mm, A, run), row) in enumerate(piv.iterrows(), start=3):
        ws.cell(row=i, column=1, value=round(D_mm, 4))
        ws.cell(row=i, column=2, value=A)
        ws.cell(row=i, column=3, value=run)
        for j, f_val in enumerate(freqs):
            v = row[f_val]
            if pd.notna(v):
                ws.cell(row=i, column=4 + j, value=round(float(v), 4))

    for col in range(1, 4 + n_f):
        ws.column_dimensions[get_column_letter(col)].width = 10


def write_by_frequency_sheet(wb, agg, box_agg):
    """添付画像2形式: f[Hz]ごとに D/a_amp/η*/Re/Im の表を横に並べる。
    各ブロックにはその f で測定がある (D, A) の行だけを詰めて書く(A=350 は 50Hz のみ、など
    周波数によって A の組が違うため、全ブロック共通の行にすると空欄行ができて散布図の線が途切れる)。
    各ブロックの末尾に、同じ f の空ボックスの値を D="box" として載せる(差し引きはしていない)。"""
    ws = wb.create_sheet("by_frequency")
    freqs = sorted(set(agg.f) | set(box_agg.f))

    total_cols = len(freqs) * 6 - 1  # 5列+区切り1列、最後は区切りなし
    ws.cell(row=1, column=1, value="η* vs 加速度振幅(周波数・粒径別)")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(total_cols, 1))

    for b, f_val in enumerate(freqs):
        c0 = b * 6 + 1  # このブロックの先頭列(D列)
        ws.cell(row=2, column=c0, value=f"f={f_val:g}Hz")
        ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + 4)
        for k, h in enumerate(["D", "a_amp", "η*", "Re", "Im"]):
            ws.cell(row=3, column=c0 + k, value=h)

        sub = agg[agg.f == f_val].sort_values(["D_mm", "A"]).assign(D_label=lambda d: d.D_mm.round(4))
        box = box_agg[box_agg.f == f_val].sort_values("A").assign(D_label="box")
        for i, r in enumerate(pd.concat([sub, box]).itertuples(index=False), start=4):
            ws.cell(row=i, column=c0 + 0, value=r.D_label)
            ws.cell(row=i, column=c0 + 1, value=r.A)
            ws.cell(row=i, column=c0 + 2, value=round(float(r.eta_mean), 4))
            ws.cell(row=i, column=c0 + 3, value=round(float(r.ReM_g_mean), 4))
            ws.cell(row=i, column=c0 + 4, value=round(float(r.ImM_g_mean), 4))

        for k in range(5):
            ws.column_dimensions[get_column_letter(c0 + k)].width = 9


def write_by_frequency_all_sheet(wb, df, box_df):
    """by_frequency と同じ並びで、平均せずにファイルごとの値を1行ずつ書く(日付違い・繰り返し測定も全部)。
    どの測定かが分かるよう、各ブロックの右端にファイル名の列を付ける。末尾に D="box" の空ボックス行。"""
    ws = wb.create_sheet("by_frequency_all")
    freqs = sorted(set(df.f) | set(box_df.f))
    width = 7  # D/a_amp/η*/Re/Im/file の6列 + 区切り1列

    ws.cell(row=1, column=1, value="η* vs 加速度振幅(周波数・粒径別、平均なし: ファイルごとの値)")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(freqs) * width - 1, 1))

    for b, f_val in enumerate(freqs):
        c0 = b * width + 1
        ws.cell(row=2, column=c0, value=f"f={f_val:g}Hz")
        ws.merge_cells(start_row=2, start_column=c0, end_row=2, end_column=c0 + 5)
        for k, h in enumerate(["D", "a_amp", "η*", "Re", "Im", "file"]):
            ws.cell(row=3, column=c0 + k, value=h)

        # 並び: D → A → 日付 → 繰り返し番号(接尾辞なし=1回目, -2, -3 の順)
        def ordered(d):
            order = d.file.str.replace(r"(_\d+)\.xlsx$", r"\1-1.xlsx", regex=True)
            return d.assign(_order=order).sort_values(["D_mm", "A", "_order"])
        sub = ordered(df[df.f == f_val]).assign(D_label=lambda d: d.D_mm.round(4))
        box = ordered(box_df[box_df.f == f_val]).assign(D_label="box")
        for i, r in enumerate(pd.concat([sub, box]).itertuples(index=False), start=4):
            ws.cell(row=i, column=c0 + 0, value=r.D_label)
            ws.cell(row=i, column=c0 + 1, value=r.A)
            if pd.notna(r.eta):
                ws.cell(row=i, column=c0 + 2, value=round(float(r.eta), 4))
            ws.cell(row=i, column=c0 + 3, value=round(float(r.ReM_g), 4))
            ws.cell(row=i, column=c0 + 4, value=round(float(r.ImM_g), 4))
            ws.cell(row=i, column=c0 + 5, value=r.file)

        for k in range(5):
            ws.column_dimensions[get_column_letter(c0 + k)].width = 9
        ws.column_dimensions[get_column_letter(c0 + 5)].width = 34


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    files = sorted(p for p in glob.glob(os.path.join(ETASTAR_DIR, "*.xlsx"))
                   if not os.path.basename(p).startswith("~$"))
    rows = [r for r in (read_one(p) for p in files) if r]
    if not rows:
        print("対象ファイルなし"); return
    df = pd.DataFrame(rows)

    box_df = df[df.is_box & df.A.notna()]
    missing = df[(~df.is_box & df.D_mm.isna()) | df.A.isna()]
    if len(missing):
        print("[warn] D または level(A) が解釈できず集計から除外:")
        print(missing[["file", "label", "level"]].to_string(index=False))
    raw_df = df.drop(index=missing.index)
    df = df[~df.is_box].dropna(subset=["D_mm", "A"])

    # 同じ D・A・f の条件で日付違いなど複数ファイルがある場合は平均値を使う
    agg = (df.groupby(["D_mm", "A", "f"])
             .agg(eta_mean=("eta", "mean"), ReM_g_mean=("ReM_g", "mean"),
                  ImM_g_mean=("ImM_g", "mean"), n=("eta", "size"))
             .reset_index())
    dup = agg[agg.n > 1]
    if len(dup):
        print(f"[info] 同一条件で複数ファイルがあり平均した件数: {len(dup)}")
    box_agg = (box_df.groupby(["A", "f"])
                     .agg(eta_mean=("eta", "mean"), ReM_g_mean=("ReM_g", "mean"),
                          ImM_g_mean=("ImM_g", "mean"), n=("eta", "size"))
                     .reset_index())

    wb = Workbook()
    wb.remove(wb.active)
    write_d_f_table_sheet(wb, build_d_f_table(agg))
    write_d_f_table_all_sheet(wb, df)
    write_by_frequency_sheet(wb, agg, box_agg)
    write_by_frequency_all_sheet(wb, df, box_df)
    ws_raw = wb.create_sheet("raw")
    ws_raw.append(list(raw_df.columns))
    for row in raw_df.itertuples(index=False):
        ws_raw.append([None if isinstance(v, float) and np.isnan(v) else v for v in row])

    out = os.path.join(OUTPUT_DIR, "eta_star_summary.xlsx")
    wb.save(out)
    print(f"保存: {out}")


if __name__ == "__main__":
    main()
