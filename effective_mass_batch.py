# -*- coding: utf-8 -*-
"""
effective_mass_batch.py  (v2)
eta_star_v5 形式の xlsx から粉体の複素有効質量 M_p を一括計算し、粒径別の図を作る。

入力 : INPUT_DIR 内の eta_star_v5_<label>_<f>Hz_<level>.xlsx
        例) eta_star_v5_30um_500Hz_20.xlsx / eta_star_v5_box_1000Hz_10.xlsx
        - EtaStar_Acc : B6=f[Hz], B7=m_p[g], B8=m_c[g], B38=η*(検算用)
        - TimeData    : A=時刻[s], B=力[N], D=加速度[m/s^2]
出力 : OUTPUT_DIR/effective_mass_summary.xlsx
       OUTPUT_DIR/effective_mass_by_size.png  (上段 Re, 中段 -Im, 下段 複素平面)

計算(シートと同一):
  X = Σ a·exp(-j2πft), Y = Σ F·exp(-j2πft), M = X*·Y/|X|^2,  η* = -2π·Im(M)/m_tot
  M_p = M − M_empty   (M_empty: 同じ f・level の空ボックス実測。無ければ下記の暫定値)

既存の集計ファイルから図だけ作り直す場合:
  python effective_mass_batch.py --from-summary 既存の effective_mass_summary.xlsx
"""
import re, glob, os, sys
import numpy as np
import pandas as pd
from openpyxl import load_workbook
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

# ===================== 設定 =====================
INPUT_DIR = r"C:\Users\USER\Desktop\calcREIM"
OUTPUT_DIR = "./output"
EMPTY_LABELS = ["box"]          # 空ボックスのラベル (m_p=0 も空扱い)
ACC_FACTOR = -1000 / 3.18       # D列が空のとき G列→加速度 換算

# 空ボックス実測が無い (f, level) の扱い:
#   Re: 空ボックス実測の全平均を使う / Im: 0 とする   → baseline 列に "provisional" と記録
FALLBACK_BASELINE = True

# 除外条件 (天面衝突=モード1)。各 dict の条件を全て満たす行を除外。
#   f: 周波数[Hz], level_min: このレベル以上, labels: 対象ラベル(省略時は全粒径)
EXCLUDE = [
    dict(f=50, level_min=20),   # 50 Hz の A≥200 は全粒径で天面衝突
]

LEVEL_ALIAS = {35: 40}          # 50 Hz の level 35 を 40 の系列として描く
LEVEL_TO_ACC = 10               # 凡例表示: level×10 ≈ A [m/s^2]
# ===============================================

FNAME_RE = re.compile(
    r"eta_star_v\d+_(?P<label>.+)_(?P<f>\d+(?:\.\d+)?)Hz_(?P<level>[^_.]+)\.xlsx$", re.IGNORECASE)


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def read_one(path):
    m = FNAME_RE.search(os.path.basename(path))
    if not m:
        print(f"[skip] ファイル名を解釈できない: {os.path.basename(path)}")
        return None
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["EtaStar_Acc"]
    f_sheet, m_p, m_c, eta_sheet = ws["B6"].value, ws["B7"].value or 0.0, ws["B8"].value or 0.0, ws["B38"].value
    t, F, a = [], [], []
    for row in wb["TimeData"].iter_rows(min_row=2, max_col=8, values_only=True):
        row = list(row) + [None] * (8 - len(row))
        ti, Fi, ai = row[0], row[1], row[3]
        if not num(ti):
            continue
        if not num(ai):
            if not num(row[6]):
                continue
            ai = row[6] * ACC_FACTOR
        if not num(Fi):
            Fi = row[7]
            if not num(Fi):
                continue
        t.append(ti); F.append(Fi); a.append(ai)
    wb.close()
    t, F, a = map(np.asarray, (t, F, a))
    f = float(f_sheet) if num(f_sheet) else float(m.group("f"))
    e = np.exp(-2j * np.pi * f * t)
    X, Y = np.sum(a * e), np.sum(F * e)
    M = np.conj(X) * Y / abs(X) ** 2
    m_tot = (m_p + m_c) / 1000.0
    label = m.group("label")
    return dict(file=os.path.basename(path), label=label,
                is_empty=(label.lower() in EMPTY_LABELS) or m_p == 0,
                f=f, level=m.group("level"), N=len(t), A_acc=2 * abs(X) / len(t),
                m_p_g=m_p, m_c_g=m_c, ReM_g=M.real * 1000, ImM_g=M.imag * 1000,
                eta_recalc=(-2 * np.pi * M.imag / m_tot) if m_tot > 0 else np.nan,
                eta_sheet=eta_sheet)


def size_um(label):
    d = re.search(r"(\d+(?:\.\d+)?)", str(label))
    return float(d.group(1)) if d else np.nan


def is_excluded(r):
    for c in EXCLUDE:
        if "f" in c and r.f != c["f"]:
            continue
        if "level_min" in c and r.level_num < c["level_min"]:
            continue
        if "labels" in c and r.label not in c["labels"]:
            continue
        return True
    return False


def process(df):
    df = df.copy()
    df["level_num"] = pd.to_numeric(df["level"], errors="coerce")
    df["is_empty"] = df["is_empty"].astype(bool)
    df["eta_check"] = df["eta_recalc"] - pd.to_numeric(df["eta_sheet"], errors="coerce")

    empty = df[df.is_empty].groupby(["f", "level_num"])[["ReM_g", "ImM_g"]].mean()
    re_fallback = empty["ReM_g"].mean() if len(empty) else np.nan
    p = df[~df.is_empty].copy()

    def base(r):
        k = (r.f, r.level_num)
        if k in empty.index:
            return pd.Series([empty.loc[k, "ReM_g"], empty.loc[k, "ImM_g"], "measured"])
        if FALLBACK_BASELINE and not np.isnan(re_fallback):
            return pd.Series([re_fallback, 0.0, "provisional"])
        return pd.Series([np.nan, np.nan, "missing"])

    p[["ReM_empty_g", "ImM_empty_g", "baseline"]] = p.apply(base, axis=1)
    p["ReMp_g"] = p.ReM_g - p.ReM_empty_g
    p["ImMp_g"] = p.ImM_g - p.ImM_empty_g
    p["ReMp_over_mp"] = p.ReMp_g / p.m_p_g
    p["negImMp_over_mp"] = -p.ImMp_g / p.m_p_g
    p["X_mm"] = p.A_acc / (2 * np.pi * p.f) ** 2 * 1000
    p["excluded"] = p.apply(is_excluded, axis=1)
    p["D_um"] = p.label.map(size_um)
    p["series"] = p.level_num.map(lambda v: LEVEL_ALIAS.get(v, v))
    return df, p, re_fallback


def plot_grid(p, path, re_fallback):
    q = p[~p.excluded & p.ReMp_g.notna()]
    Ds = sorted(q.D_um.dropna().unique())
    if not Ds:
        print("描画対象なし"); return
    lvs = sorted(q.series.dropna().unique())
    fig, ax = plt.subplots(3, len(Ds), figsize=(max(4.4 * len(Ds), 11), 12), squeeze=False)
    for j, D in enumerate(Ds):
        for lv in lvs:
            h = q[(q.D_um == D) & (q.series == lv)].sort_values("f")
            if h.empty:
                continue
            lab = f"A≈{lv * LEVEL_TO_ACC:g}"
            ax[0, j].plot(h.f, h.ReMp_over_mp, "o-", label=lab)
            ax[1, j].plot(h.f, h.negImMp_over_mp, "o-", label=lab)
            ax[2, j].plot(h.ReMp_over_mp, h.negImMp_over_mp, "o-", label=lab)
            for _, r in h.iterrows():
                ax[2, j].annotate(f"{r.f:g}", (r.ReMp_over_mp, r.negImMp_over_mp),
                                  fontsize=6, xytext=(2, 2), textcoords="offset points")
        ax[0, j].set_title(f"D={D:g} um")
        for i in (0, 1):
            ax[i, j].set_xscale("log"); ax[i, j].grid(alpha=.3); ax[i, j].set_xlim(40, 2000)
            ax[i, j].set_xlabel("f [Hz]")
        ax[0, j].set_ylim(-0.3, 1.3); ax[1, j].set_ylim(0, 0.9)
        ax[0, j].axhline(1, ls="--", c="gray", lw=.6); ax[0, j].axhline(0, c="gray", lw=.6)
        ax[2, j].set_xlim(-0.3, 1.3); ax[2, j].set_ylim(0, 0.9); ax[2, j].grid(alpha=.3)
        ax[2, j].axvline(1, ls="--", c="gray", lw=.6); ax[2, j].axvline(0, c="gray", lw=.6)
        ax[2, j].set_xlabel("Re(M_p)/m_p")
    ax[0, 0].set_ylabel("Re(M_p)/m_p"); ax[1, 0].set_ylabel("-Im(M_p)/m_p")
    ax[2, 0].set_ylabel("-Im(M_p)/m_p"); ax[0, 0].legend(fontsize=8)
    excl = "; ".join(", ".join(f"{k}={v}" for k, v in c.items()) for c in EXCLUDE) or "none"
    n_prov = (q.baseline == "provisional").sum()
    note = (f"Excluded: {excl}.  Baseline: measured empty box where available; "
            f"{n_prov} points use provisional Re={re_fallback:.3f} g, Im=0.") if n_prov else \
           f"Excluded: {excl}.  Baseline: measured empty box for all points."
    fig.suptitle("Powder effective mass by particle size (bottom row labels = f [Hz])\n" + note, fontsize=10)
    fig.tight_layout(); fig.savefig(path, dpi=120); plt.close(fig)
    print(f"保存: {path}")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    if len(sys.argv) >= 3 and sys.argv[1] == "--from-summary":
        df = pd.read_excel(sys.argv[2], sheet_name="all_files")
    else:
        files = sorted(glob.glob(os.path.join(INPUT_DIR, "*.xlsx")))
        df = pd.DataFrame([r for r in (read_one(x) for x in files) if r])
    if df.empty:
        print("対象ファイルなし"); return
    df, p, re_fb = process(df)

    bad = df[df.eta_check.abs() > 1e-6]
    if len(bad):
        print("[warn] シートのη*と再計算値が一致しないファイル:")
        print(bad[["file", "eta_recalc", "eta_sheet"]].to_string(index=False))
    miss = p[p.baseline == "missing"]
    if len(miss):
        print("[warn] 空ボックスが無く差し引けない条件:", len(miss))
    print(f"baseline: measured {(p.baseline=='measured').sum()} / provisional "
          f"{(p.baseline=='provisional').sum()} / 除外 {p.excluded.sum()}")

    out = os.path.join(OUTPUT_DIR, "effective_mass_summary.xlsx")
    with pd.ExcelWriter(out) as w:
        df.sort_values(["label", "level_num", "f"]).to_excel(w, sheet_name="all_files", index=False)
        p.sort_values(["D_um", "level_num", "f"]).to_excel(w, sheet_name="powder_Mp", index=False)
        q = p[~p.excluded]
        for col, name in [("ReMp_over_mp", "Re_pivot"), ("negImMp_over_mp", "negIm_pivot")]:
            q.pivot_table(index=["series", "f"], columns="D_um", values=col).round(3).to_excel(w, sheet_name=name)
    print(f"保存: {out}")
    plot_grid(p, os.path.join(OUTPUT_DIR, "effective_mass_by_size.png"), re_fb)


if __name__ == "__main__":
    main()
