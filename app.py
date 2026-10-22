# -*- coding: utf-8 -*-
"""
IPLM 2025 — Final (Target Sampel 33.88% per Jenis) — Versi Native Streamlit
"""

import io
import json
import math
import os
import re
import tempfile
import time
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from sklearn.preprocessing import PowerTransformer
import streamlit as st

# python-docx
DOCX_AVAILABLE = True
try:
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt
except Exception:
    DOCX_AVAILABLE = False
    Document = None

# huggingface client
HF_AVAILABLE = True
try:
    from huggingface_hub import InferenceClient
except Exception:
    HF_AVAILABLE = False
    InferenceClient = None

# ============================================================
# 1) KONFIGURASI
# ============================================================
st.set_page_config(page_title="IPLM 2025", layout="wide")

DATA_FILE = os.getenv(
    "DATA_FILE", "DATA CLEAN GABUNGAN SANGGAH-TIDAK SANGGAH ALL2026.xlsx"
)
POP_KAB = os.getenv("POP_KAB", "Data_populasi_Kab_kota_fixed.xlsx")
POP_PROV = os.getenv("POP_PROV", "Data_populasi_propinsi.xlsx")
POP_KHUSUS = os.getenv("POP_KHUSUS", "Data_populasi_perp_khusus.xlsx")

W_KEPATUHAN = float(os.getenv("W_KEPATUHAN", "0.30"))
W_KINERJA = float(os.getenv("W_KINERJA", "0.70"))
TARGET_RATIO = float(os.getenv("TARGET_RATIO", "0.3388"))

USE_LLM = True
LLM_MODEL_NAME = os.getenv(
    "LLM_MODEL_NAME", "deepseek-ai/DeepSeek-R1-Distill-Qwen-32B"
)
HF_TOKEN = (
    os.getenv("HF_SECRET")
    or os.getenv("HF_TOKEN")
    or os.getenv("HUGGINGFACEHUB_API_TOKEN")
    or os.getenv("HF_API_TOKEN")
)


# ============================================================
# 2) UTILITAS
# ============================================================
def _canon(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


def _disp_text(x):
    if pd.isna(x):
        return None
    t = str(x).strip().upper()
    return " ".join(t.split())


def pick_col(df, candidates):
    if df is None or df.empty:
        return None
    for c in candidates:
        if c in df.columns:
            return c
    can_map = {_canon(c): c for c in df.columns}
    for c in candidates:
        k = _canon(c)
        if k in can_map:
            return can_map[k]
    return None


def coerce_num(val):
    if pd.isna(val):
        return np.nan
    t = str(val).strip()
    if t == "" or t in {"-", "–", "—", "NA", "N/A", "null", "NULL"}:
        return np.nan
    t = t.replace("\u00a0", " ").replace("Rp", "").replace("%", "")
    t = re.sub(r"[^0-9,.\-]", "", t)
    if t.count(".") > 1 and t.count(",") == 1:
        t = t.replace(".", "").replace(",", ".")
    elif t.count(",") > 1 and t.count(".") == 1:
        t = t.replace(",", "")
    elif t.count(",") == 1 and t.count(".") == 0:
        t = t.replace(",", ".")
    else:
        t = t.replace(",", "")
    try:
        return float(t)
    except Exception:
        return np.nan


def minmax_norm(s: pd.Series) -> pd.Series:
    x = pd.to_numeric(s, errors="coerce").astype(float)
    mn, mx = x.min(skipna=True), x.max(skipna=True)
    if pd.isna(mn) or pd.isna(mx) or mx == mn:
        return pd.Series(0.0, index=s.index)
    return (x - mn) / (mx - mn)


def norm_kew(v):
    if pd.isna(v):
        return None
    t = str(v).strip().upper()
    if "KAB" in t or "KOTA" in t:
        return "KAB/KOTA"
    if "PROV" in t:
        return "PROVINSI"
    if "PUSAT" in t or "NASIONAL" in t:
        return "PUSAT"
    return t


def norm_prov_disp(s):
    if pd.isna(s):
        return None
    t = str(s).strip().upper().replace("\u00a0", " ")
    t = " ".join(t.split()).replace("PROPINSI", "PROVINSI")
    while t.startswith("PROVINSI PROVINSI "):
        t = t.replace("PROVINSI PROVINSI ", "PROVINSI ", 1)
    name = t[len("PROVINSI ") :].strip() if t.startswith("PROVINSI ") else t
    name = " ".join(name.split())
    return f"PROVINSI {name}" if name else None


def norm_prov_label(s):
    if pd.isna(s):
        return None
    t = str(s).strip().upper().replace("\u00a0", " ")
    t = (
        " ".join(t.split())
        .replace("PROPINSI", "PROVINSI")
        .replace("PROVINSI", "")
        .strip()
    )
    return re.sub(r"[^A-Z0-9]+", "", t)


def norm_kab_label(s):
    if pd.isna(s):
        return None
    t = str(s).upper()
    t = (
        t.replace("KABUPATEN", "KAB")
        .replace("KAB.", "KAB")
        .replace("KOTA ADMINISTRASI", "KOTA")
        .replace("KOTA ADM.", "KOTA")
        .replace("KOTA.", "KOTA")
    )
    return re.sub(r"[^A-Z0-9]+", "", " ".join(t.split()))


def safe_div(num, den):
    if den is None or pd.isna(den) or float(den) <= 0:
        return np.nan
    return float(num) / float(den)


def faktor_penyesuaian_total(n_total: float, target_total: float) -> float:
    if target_total is None or pd.isna(target_total) or float(target_total) <= 0:
        return 1.0
    if n_total is None or pd.isna(n_total) or float(n_total) < 0:
        n_total = 0.0
    return float(min(float(n_total) / float(target_total), 1.0))


# Indikator
koleksi_cols = [
    "JudulTercetak",
    "EksemplarTercetak",
    "JudulElektronik",
    "EksemplarElektronik",
    "TambahJudulTercetak",
    "TambahEksemplarTercetak",
    "TambahJudulElektronik",
    "TambahEksemplarElektronik",
    "KomitmenAnggaranKoleksi",
]
sdm_cols = [
    "TenagaKualifikasiIlmuPerpustakaan",
    "TenagaFungsionalProfesional",
    "TenagaPKB",
    "AnggaranTenaga",
]
pelayanan_cols = [
    "PesertaBudayaBaca",
    "PemustakaLuringDaring",
    "PemustakaFasilitasTIK",
    "PemanfaatanJudulTercetak",
    "PemanfaatanEksemplarTercetak",
    "PemanfaatanJudulElektronik",
    "PemanfaatanEksemplarElektronik",
]
pengelolaan_cols = [
    "KegiatanBudayaBaca",
    "KegiatanKerjasama",
    "VariasiLayanan",
    "Kebijakan",
    "AnggaranLayanan",
]
all_indicators = koleksi_cols + sdm_cols + pelayanan_cols + pengelolaan_cols

alias_map_raw = {
    "j_judul_koleksi_tercetak": "JudulTercetak",
    "j_eksemplar_koleksi_tercetak": "EksemplarTercetak",
    "j_judul_koleksi_digital": "JudulElektronik",
    "j_eksemplar_koleksi_digital": "EksemplarElektronik",
    "tambah_judul_koleksi_tercetak": "TambahJudulTercetak",
    "tambah_eksemplar_koleksi_tercetak": "TambahEksemplarTercetak",
    "tambah_judul_koleksi_digital": "TambahJudulElektronik",
    "tambah_eksemplar_koleksi_digital": "TambahEksemplarElektronik",
    "j_anggaran_koleksi": "KomitmenAnggaranKoleksi",
    "j_tenaga_ilmu_perpus": "TenagaKualifikasiIlmuPerpustakaan",
    "j_tenaga_nonilmu_perpus": "TenagaFungsionalProfesional",
    "j_tenaga_pkb": "TenagaPKB",
    "j_anggaran_diklat_perpus": "AnggaranTenaga",
    "j_peserta_budaya_baca": "PesertaBudayaBaca",
    "j_pemustaka_luring_daring": "PemustakaLuringDaring",
    "j_pemustaka_fasilitas_tik": "PemustakaFasilitasTIK",
    "j_judul_koleksi_tercetak_termanfaat": "PemanfaatanJudulTercetak",
    "j_eksemplar_koleksi_tercetak_termanfaat": "PemanfaatanEksemplarTercetak",
    "j_judul_koleksi_digital_termanfaat": "PemanfaatanJudulElektronik",
    "j_eksemplar_koleksi_digital_termanfaat": "PemanfaatanEksemplarElektronik",
    "j_kegiatan_budaya_baca_peningkatan_literasi": "KegiatanBudayaBaca",
    "j_kerjasama_pengembangan_perpus": "KegiatanKerjasama",
    "j_variasi_layanan": "VariasiLayanan",
    "j_kebijakan_prosedur_pelayanan": "Kebijakan",
    "j_anggaran_peningkatan_pelayanan": "AnggaranLayanan",
}
alias_map = {_canon(k): v for k, v in alias_map_raw.items()}


def _mean_norm_cols(row, cols):
    vals = [
        float(row[f"norm_{c}"])
        for c in cols
        if f"norm_{c}" in row.index and not pd.isna(row[f"norm_{c}"])
    ]
    return float(np.mean(vals)) if vals else 0.0


def prepare_global(df_src: pd.DataFrame) -> pd.DataFrame:
    if df_src is None or df_src.empty:
        return df_src
    df = df_src.copy()
    rename_map = {}
    for col in df.columns:
        c = _canon(col)
        if c in alias_map:
            rename_map[col] = alias_map[c]
        else:
            for tgt in all_indicators:
                if c == _canon(tgt):
                    rename_map[col] = tgt
                    break
    if rename_map:
        df = df.rename(columns=rename_map)

    available = [c for c in all_indicators if c in df.columns]
    for c in available:
        df[c] = df[c].apply(coerce_num)
        x = pd.to_numeric(df[c], errors="coerce").astype(float).values
        mask = ~np.isnan(x)
        transformed = np.full_like(x, np.nan, dtype=float)
        if mask.sum() > 1:
            pt = PowerTransformer(method="yeo-johnson", standardize=False)
            transformed[mask] = pt.fit_transform(x[mask].reshape(-1, 1)).ravel()
        else:
            transformed[mask] = x[mask]
        df[f"norm_{c}"] = minmax_norm(pd.Series(transformed, index=df.index))

    df["sub_koleksi"] = df.apply(
        lambda r: _mean_norm_cols(
            r, [c for c in koleksi_cols if c in available]
        ),
        axis=1,
    )
    df["sub_sdm"] = df.apply(
        lambda r: _mean_norm_cols(r, [c for c in sdm_cols if c in available]),
        axis=1,
    )
    df["sub_pelayanan"] = df.apply(
        lambda r: _mean_norm_cols(
            r, [c for c in pelayanan_cols if c in available]
        ),
        axis=1,
    )
    df["sub_pengelolaan"] = df.apply(
        lambda r: _mean_norm_cols(
            r, [c for c in pengelolaan_cols if c in available]
        ),
        axis=1,
    )
    df["dim_kepatuhan"] = df[["sub_koleksi", "sub_sdm"]].mean(axis=1)
    df["dim_kinerja"] = df[["sub_pelayanan", "sub_pengelolaan"]].mean(axis=1)
    df["Indeks_Dasar_0_100"] = 100 * (
        W_KEPATUHAN * df["dim_kepatuhan"] + W_KINERJA * df["dim_kinerja"]
    )

    for c in [
        "sub_koleksi",
        "sub_sdm",
        "sub_pelayanan",
        "sub_pengelolaan",
        "dim_kepatuhan",
        "dim_kinerja",
        "Indeks_Dasar_0_100",
    ]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


@st.cache_data(show_spinner="Memuat data dan menghitung normalisasi...")
def load_all_data():
    for p in [DATA_FILE, POP_KAB, POP_PROV, POP_KHUSUS]:
        if not Path(p).exists():
            return None, None, None, None, None, {}, f"File tidak ditemukan: {p}"

    fp = Path(DATA_FILE)
    xls = pd.ExcelFile(fp)
    frames = [pd.read_excel(fp, sheet_name=s) for s in xls.sheet_names]
    df_raw = pd.concat(frames, ignore_index=True, sort=False)

    prov_col = pick_col(df_raw, ["provinsi", "Provinsi", "PROVINSI"])
    kab_col = pick_col(
        df_raw,
        [
            "kab/kota",
            "Kab/Kota",
            "Kab_Kota",
            "KAB/KOTA",
            "kabupaten_kota",
            "Kabupaten/Kota",
            "kabupaten kota",
            "kota",
            "kab_kota",
        ],
    )
    kew_col = pick_col(
        df_raw, ["kewenangan", "jenis_kewenangan", "Kewenangan", "KEWENANGAN"]
    )
    jenis_col = pick_col(
        df_raw, ["jenis_perpustakaan", "Jenis Perpustakaan", "JENIS_PERPUSTAKAAN"]
    )
    nama_col = pick_col(
        df_raw,
        [
            "nm_perpustakaan",
            "nama_perpustakaan",
            "Nama Perpustakaan",
            "nm_instansi_lembaga",
            "nm_perpus",
        ],
    )

    val_map_jenis = {
        "PERPUSTAKAAN SEKOLAH": "sekolah",
        "SEKOLAH": "sekolah",
        "PERPUSTAKAAN UMUM": "umum",
        "UMUM": "umum",
        "PERPUSTAKAAN DAERAH": "umum",
        "PERPUSTAKAAN KHUSUS": "khusus",
        "KHUSUS": "khusus",
    }
    df_raw["KEW_NORM"] = df_raw[kew_col].apply(norm_kew)
    df_raw["_dataset"] = (
        df_raw[jenis_col]
        .astype(str)
        .str.strip()
        .str.upper()
        .map(val_map_jenis)
    )
    df_raw["PROV_DISP"] = df_raw[prov_col].apply(norm_prov_disp)
    df_raw["KAB_DISP"] = df_raw[kab_col].apply(_disp_text)
    df_raw["prov_key"] = df_raw["PROV_DISP"].apply(norm_prov_label)
    df_raw["kab_key"] = df_raw["KAB_DISP"].apply(norm_kab_label)

    kcols = (
        [prov_col, kab_col, kew_col, jenis_col, nama_col]
        if nama_col
        else [prov_col, kab_col, kew_col, jenis_col]
    )
    tmp = (
        df_raw[kcols].astype(str).fillna("").apply(lambda s: s.str.strip(), axis=0)
    )
    df_raw["_row_key"] = (
        tmp.apply(lambda r: "||".join(r.values.tolist()), axis=1).apply(_canon)
    )
    df_raw = df_raw.drop_duplicates(subset=["_row_key"], keep="first").copy()

    pk = pd.read_excel(POP_KAB)
    c_kab = pick_col(
        pk,
        [
            "KABUPATEN_KOTA",
            "Kab/Kota",
            "Kabupaten/Kota",
            "KAB/KOTA",
            "Kabupaten_Kota",
            "kab_kota",
            "kabupaten_kota",
        ],
    )
    c_prov = pick_col(pk, ["PROVINSI", "Provinsi", "provinsi"])
    pop_kab = pk.copy()
    pop_kab["Kab_Kota_Label"] = pk[c_kab].astype(str).str.strip()
    pop_kab["Provinsi_Label"] = (
        pk[c_prov].astype(str).str.strip() if c_prov else ""
    )
    pop_kab["kab_key"] = pop_kab["Kab_Kota_Label"].apply(norm_kab_label)
    pop_kab = pop_kab.groupby("kab_key", as_index=False).first()

    pp = pd.read_excel(POP_PROV)
    c_pr = pick_col(
        pp,
        [
            "Provinsi",
            "PROVINSI",
            "provinsi",
            "Propinsi",
            "PROPINSI",
            "propinsi",
        ],
    )
    pop_prov = pp.copy()
    pop_prov["Provinsi_Label"] = pp[c_pr].astype(str).str.strip()
    pop_prov["prov_key"] = pop_prov["Provinsi_Label"].apply(norm_prov_label)
    pop_prov = pop_prov.groupby("prov_key", as_index=False).first()

    # parse pop khusus
    df_pkh = pd.read_excel(POP_KHUSUS)
    c_mix = pick_col(
        df_pkh,
        [
            "Propinsi/Kab/kota",
            "Propinsi/Kab/Kota",
            "Provinsi/Kab/Kota",
            "Provinsi/Kab/kota",
            "Provinsi/Kabupaten/Kota",
            "Wilayah",
            "Nama Wilayah",
        ],
    )
    c_pop = pick_col(df_pkh, ["POP_KHUSUS", "Pop_Khusus", "pop_khusus"])
    mix = df_pkh[c_mix].astype(str).fillna("").str.strip()
    pval_s = df_pkh[c_pop].apply(coerce_num)
    rows_p = []
    c_cur_prov = None
    for m, pval in zip(mix.tolist(), pval_s.tolist()):
        mm = _disp_text(m) or ""
        if not mm:
            continue
        if mm.startswith("PROVINSI "):
            prov_n = mm.replace("PROVINSI", "").strip()
            c_cur_prov = prov_n
            rows_p.append({
                "LEVEL": "PROV",
                "Provinsi_Label": f"PROVINSI {prov_n}",
                "Kab_Kota_Label": None,
                "Pop_Total_Jenis": pval,
            })
        else:
            rows_p.append({
                "LEVEL": "KAB",
                "Provinsi_Label": (
                    f"PROVINSI {c_cur_prov}" if c_cur_prov else None
                ),
                "Kab_Kota_Label": mm,
                "Pop_Total_Jenis": pval,
            })
    pop_khusus = pd.DataFrame(rows_p)
    pop_khusus["Pop_Total_Jenis"] = (
        pd.to_numeric(pop_khusus["Pop_Total_Jenis"], errors="coerce").fillna(0.0)
    )
    pop_khusus["prov_key"] = pop_khusus["Provinsi_Label"].apply(norm_prov_label)
    pop_khusus["kab_key"] = pop_khusus["Kab_Kota_Label"].apply(norm_kab_label)

    df_all = prepare_global(df_raw)
    meta = dict(
        prov_col=prov_col,
        kab_col=kab_col,
        kew_col=kew_col,
        jenis_col=jenis_col,
        nama_col=nama_col,
    )
    return df_all, df_raw, pop_kab, pop_prov, pop_khusus, meta, "Sukses"


# ============================================================
# 3) KALKULASI & AGREGASI
# ============================================================
def build_faktor_wilayah_jenis(
    df_filtered, pop_kab, pop_prov, pop_khusus, kew_value
):
    if df_filtered is None or df_filtered.empty:
        return pd.DataFrame()
    kew_norm = str(kew_value or "").upper()
    df = df_filtered[
        df_filtered["_dataset"].isin(["sekolah", "umum", "khusus"])
    ].copy()
    if df.empty:
        return pd.DataFrame()

    jenis_list = ["sekolah", "umum", "khusus"]
    if "PROV" in kew_norm:
        key_col, label_col, label_name, mode = (
            "prov_key",
            "PROV_DISP",
            "Provinsi",
            "PROV",
        )
        base_pop = (
            pop_prov.copy()
            if (pop_prov is not None and not pop_prov.empty)
            else pd.DataFrame()
        )
        if not base_pop.empty and "prov_key" not in base_pop.columns:
            base_pop["prov_key"] = base_pop["Provinsi_Label"].apply(
                norm_prov_label
            )
        base_pop = (
            base_pop.set_index("prov_key")
            if (not base_pop.empty and "prov_key" in base_pop.columns)
            else pd.DataFrame().set_index(pd.Index([]))
        )
    else:
        key_col, label_col, label_name, mode = (
            "kab_key",
            "KAB_DISP",
            "Kab/Kota",
            "KAB",
        )
        base_pop = (
            pop_kab.copy()
            if (pop_kab is not None and not pop_kab.empty)
            else pd.DataFrame()
        )
        if not base_pop.empty and "kab_key" not in base_pop.columns:
            base_pop["kab_key"] = base_pop["Kab_Kota_Label"].apply(
                norm_kab_label
            )
        base_pop = (
            base_pop.set_index("kab_key")
            if (not base_pop.empty and "kab_key" in base_pop.columns)
            else pd.DataFrame().set_index(pd.Index([]))
        )

    base_keys = (
        df[[key_col, label_col]]
        .drop_duplicates()
        .rename(columns={key_col: "group_key", label_col: label_name})
    )
    full = (
        base_keys.assign(_tmp=1)
        .merge(pd.DataFrame({"Jenis": jenis_list, "_tmp": 1}), on="_tmp")
        .drop(columns="_tmp")
    )
    cnt = (
        df.groupby([key_col, label_col, "_dataset"], dropna=False)
        .size()
        .reset_index(name="n_jenis")
        .rename(
            columns={
                key_col: "group_key",
                label_col: label_name,
                "_dataset": "Jenis",
            }
        )
    )
    cnt["Jenis"] = cnt["Jenis"].astype(str).str.lower().str.strip()

    base_n = full.merge(cnt, on=["group_key", label_name, "Jenis"], how="left")
    base_n["n_jenis"] = (
        pd.to_numeric(base_n["n_jenis"], errors="coerce").fillna(0).astype(int)
    )
    base_n["target_total_33_88_jenis"] = 0.0
    base_n["pop_total_jenis"] = 0.0

    if not base_pop.empty:
        if mode == "KAB":
            pop_sekolah = pd.to_numeric(
                base_pop.get(
                    "jumlah_populasi_sekolah_base",
                    base_pop.get("jumlah_populasi_sekolah", 0),
                ),
                errors="coerce",
            ).fillna(0.0)
            pop_umum = pd.to_numeric(
                base_pop.get(
                    "jumlah_populasi_umum_base",
                    base_pop.get("jumlah_populasi_umum", 0),
                ),
                errors="coerce",
            ).fillna(0.0)
        else:
            sma = pd.to_numeric(
                base_pop.get("sma ", base_pop.get("sma", 0)), errors="coerce"
            ).fillna(0.0)
            smk = pd.to_numeric(base_pop.get("smk", 0), errors="coerce").fillna(
                0.0
            )
            slb = pd.to_numeric(base_pop.get("slb", 0), errors="coerce").fillna(
                0.0
            )
            pop_sekolah = sma + smk + slb
            pop_umum = pd.to_numeric(
                base_pop.get("perpus_umum_prop", 0), errors="coerce"
            ).fillna(0.0)

        tgt_sekolah = pop_sekolah * float(TARGET_RATIO)
        tgt_umum = pop_umum * float(TARGET_RATIO)

        m = base_n["Jenis"].eq("sekolah")
        base_n.loc[m, "pop_total_jenis"] = (
            base_n.loc[m, "group_key"].map(pop_sekolah).fillna(0.0).values
        )
        base_n.loc[m, "target_total_33_88_jenis"] = (
            base_n.loc[m, "group_key"].map(tgt_sekolah).fillna(0.0).values
        )

        m = base_n["Jenis"].eq("umum")
        base_n.loc[m, "pop_total_jenis"] = (
            base_n.loc[m, "group_key"].map(pop_umum).fillna(0.0).values
        )
        base_n.loc[m, "target_total_33_88_jenis"] = (
            base_n.loc[m, "group_key"].map(tgt_umum).fillna(0.0).values
        )

    if pop_khusus is not None and not pop_khusus.empty:
        pk = pop_khusus.copy()
        pk["Pop_Total_Jenis"] = pd.to_numeric(
            pk.get("Pop_Total_Jenis", 0), errors="coerce"
        ).fillna(0.0)
        lvl = "PROV" if mode == "PROV" else "KAB"
        kmap = (
            pk[pk["LEVEL"].astype(str).str.upper() == lvl]
            .groupby(f"{lvl.lower()}_key")
            .agg(pop=("Pop_Total_Jenis", "sum"))["pop"]
        )
        tgt_s = kmap * float(TARGET_RATIO)
        m = base_n["Jenis"].eq("khusus")
        base_n.loc[m, "pop_total_jenis"] = (
            base_n.loc[m, "group_key"].map(kmap).fillna(0.0).values
        )
        base_n.loc[m, "target_total_33_88_jenis"] = (
            base_n.loc[m, "group_key"].map(tgt_s).fillna(0.0).values
        )

    base_n["faktor_penyesuaian_jenis"] = [
        faktor_penyesuaian_total(n, t)
        for n, t in zip(
            pd.to_numeric(base_n["n_jenis"], errors="coerce")
            .fillna(0)
            .astype(float),
            pd.to_numeric(base_n["target_total_33_88_jenis"], errors="coerce")
            .fillna(0)
            .astype(float),
        )
    ]
    base_n["coverage_jenis_%"] = [
        (safe_div(n, p) * 100.0)
        if (p is not None and not pd.isna(p) and float(p) > 0)
        else 0.0
        for n, p in zip(
            pd.to_numeric(base_n["n_jenis"], errors="coerce")
            .fillna(0)
            .astype(float),
            pd.to_numeric(base_n["pop_total_jenis"], errors="coerce")
            .fillna(0)
            .astype(float),
        )
    ]
    base_n["gap_target33_88_jenis"] = [
        max(float(t) - float(n), 0.0)
        for n, t in zip(
            pd.to_numeric(base_n["n_jenis"], errors="coerce")
            .fillna(0)
            .astype(float),
            pd.to_numeric(base_n["target_total_33_88_jenis"], errors="coerce")
            .fillna(0)
            .astype(float),
        )
    ]
    base_n["target_total_33_88_jenis"] = (
        pd.to_numeric(base_n["target_total_33_88_jenis"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )
    base_n["pop_total_jenis"] = (
        pd.to_numeric(base_n["pop_total_jenis"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )
    base_n["coverage_jenis_%"] = (
        pd.to_numeric(base_n["coverage_jenis_%"], errors="coerce")
        .fillna(0.0)
        .round(2)
    )
    base_n["faktor_penyesuaian_jenis"] = (
        pd.to_numeric(base_n["faktor_penyesuaian_jenis"], errors="coerce")
        .fillna(1.0)
        .round(3)
    )
    base_n["gap_target33_88_jenis"] = (
        pd.to_numeric(base_n["gap_target33_88_jenis"], errors="coerce")
        .fillna(0)
        .round(0)
        .astype(int)
    )
    return base_n


def build_agg_wilayah_jenis(df_filtered, faktor_wilayah_jenis, kew_value):
    if df_filtered is None or df_filtered.empty:
        return pd.DataFrame()
    kew_norm = str(kew_value or "").upper()
    df = df_filtered[
        df_filtered["_dataset"].isin(["sekolah", "umum", "khusus"])
    ].copy()
    if df.empty:
        return pd.DataFrame()

    key_col, label_col, label_name = (
        ("prov_key", "PROV_DISP", "Provinsi")
        if "PROV" in kew_norm
        else ("kab_key", "KAB_DISP", "Kab/Kota")
    )
    jenis_list = ["sekolah", "umum", "khusus"]
    base_keys = (
        df[[key_col, label_col]]
        .drop_duplicates()
        .rename(columns={key_col: "group_key", label_col: label_name})
    )
    full = (
        base_keys.assign(_tmp=1)
        .merge(pd.DataFrame({"Jenis": jenis_list, "_tmp": 1}), on="_tmp")
        .drop(columns="_tmp")
    )

    agg_real = (
        df.groupby([key_col, label_col, "_dataset"], dropna=False)
        .agg(
            Jumlah=("Indeks_Dasar_0_100", "size"),
            Rata2_sub_koleksi=("sub_koleksi", "mean"),
            Rata2_sub_sdm=("sub_sdm", "mean"),
            Rata2_sub_pelayanan=("sub_pelayanan", "mean"),
            Rata2_sub_pengelolaan=("sub_pengelolaan", "mean"),
            Rata2_dim_kepatuhan=("dim_kepatuhan", "mean"),
            Rata2_dim_kinerja=("dim_kinerja", "mean"),
            Indeks_Dasar_Agregat_0_100=("Indeks_Dasar_0_100", "mean"),
        )
        .reset_index()
        .rename(
            columns={
                key_col: "group_key",
                label_col: label_name,
                "_dataset": "Jenis",
            }
        )
    )
    agg_real["Jenis"] = agg_real["Jenis"].astype(str).str.lower().str.strip()

    agg = full.merge(
        agg_real, on=["group_key", label_name, "Jenis"], how="left"
    )
    if faktor_wilayah_jenis is not None and not faktor_wilayah_jenis.empty:
        fw = faktor_wilayah_jenis.copy()
        fw["Jenis"] = fw["Jenis"].astype(str).str.lower().str.strip()
        keep = [
            "group_key",
            label_name,
            "Jenis",
            "faktor_penyesuaian_jenis",
            "target_total_33_88_jenis",
            "pop_total_jenis",
            "coverage_jenis_%",
            "gap_target33_88_jenis",
            "n_jenis",
        ]
        agg = agg.merge(
            fw[[c for c in keep if c in fw.columns]],
            on=["group_key", label_name, "Jenis"],
            how="left",
        )
    agg["faktor_penyesuaian_jenis"] = (
        pd.to_numeric(agg["faktor_penyesuaian_jenis"], errors="coerce")
        .fillna(1.0)
        .round(3)
    )
    agg["Indeks_Final_Agregat_0_100"] = (
        pd.to_numeric(agg["Indeks_Dasar_Agregat_0_100"], errors="coerce").fillna(
            0.0
        )
        * agg["faktor_penyesuaian_jenis"]
    ).round(2)
    return agg


def build_agg_wilayah_total_from_jenis(agg_jenis, kew_value):
    if agg_jenis is None or agg_jenis.empty:
        return pd.DataFrame()
    kew_norm = str(kew_value or "").upper()
    label_name = "Provinsi" if "PROV" in kew_norm else "Kab/Kota"

    out = (
        agg_jenis.groupby(["group_key", label_name], as_index=False)
        .agg(
            n_total=("Jumlah", "sum"),
            Rata2_sub_koleksi=("Rata2_sub_koleksi", "mean"),
            Rata2_sub_sdm=("Rata2_sub_sdm", "mean"),
            Rata2_sub_pelayanan=("Rata2_sub_pelayanan", "mean"),
            Rata2_sub_pengelolaan=("Rata2_sub_pengelolaan", "mean"),
            Rata2_dim_kepatuhan=("Rata2_dim_kepatuhan", "mean"),
            Rata2_dim_kinerja=("Rata2_dim_kinerja", "mean"),
            Indeks_Dasar_Agregat_0_100=("Indeks_Dasar_Agregat_0_100", "mean"),
            Indeks_Final_Wilayah_0_100=("Indeks_Final_Agregat_0_100", "mean"),
        )
        .round(
            {
                "Rata2_sub_koleksi": 3,
                "Rata2_sub_sdm": 3,
                "Rata2_sub_pelayanan": 3,
                "Rata2_sub_pengelolaan": 3,
                "Rata2_dim_kepatuhan": 3,
                "Rata2_dim_kinerja": 3,
                "Indeks_Dasar_Agregat_0_100": 2,
                "Indeks_Final_Wilayah_0_100": 2,
            }
        )
    )
    return out


def build_summary_per_jenis(agg_jenis, agg_total):
    jenis_list = ["sekolah", "umum", "khusus"]
    rows = []
    for j in jenis_list:
        sub = (
            agg_jenis[
                agg_jenis["Jenis"].astype(str).str.lower().str.strip() == j
            ]
            if not agg_jenis.empty
            else pd.DataFrame()
        )
        terkumpul = (
            int(sub["Jumlah"].sum())
            if not sub.empty and "Jumlah" in sub.columns
            else 0
        )
        tgt = (
            int(sub["target_total_33_88_jenis"].sum())
            if not sub.empty and "target_total_33_88_jenis" in sub.columns
            else 0
        )
        pop = (
            int(sub["pop_total_jenis"].sum())
            if not sub.empty and "pop_total_jenis" in sub.columns
            else 0
        )
        dasar = (
            float(sub["Indeks_Dasar_Agregat_0_100"].mean())
            if not sub.empty
            else 0.0
        )
        final = (
            float(sub["Indeks_Final_Agregat_0_100"].mean())
            if not sub.empty
            else 0.0
        )
        cov = (terkumpul / tgt * 100.0) if tgt > 0 else 0.0
        rows.append({
            "Jenis": j,
            "Jumlah_Wilayah": len(sub),
            "Total_Perpus": terkumpul,
            "Pop_Total_Jenis": pop,
            "Target33_88_Total_Jenis": tgt,
            "Terkumpul_Jenis": terkumpul,
            "Coverage_Target33_88_Jenis_%": round(cov, 2),
            "Indeks_Dasar_0_100": round(dasar, 2),
            "Indeks_Final_Disesuaikan_0_100": round(final, 2),
            "Penyesuaian_Poin": round(final - dasar, 2),
        })

    # Keseluruhan
    tot_terkumpul = sum(r["Terkumpul_Jenis"] for r in rows)
    tot_tgt = sum(r["Target33_88_Total_Jenis"] for r in rows)
    tot_pop = sum(r["Pop_Total_Jenis"] for r in rows)
    mean_dasar = np.mean([r["Indeks_Dasar_0_100"] for r in rows])
    mean_final = np.mean([r["Indeks_Final_Disesuaikan_0_100"] for r in rows])
    rows.append({
        "Jenis": "keseluruhan",
        "Jumlah_Wilayah": len(agg_total),
        "Total_Perpus": tot_terkumpul,
        "Pop_Total_Jenis": tot_pop,
        "Target33_88_Total_Jenis": tot_tgt,
        "Terkumpul_Jenis": tot_terkumpul,
        "Coverage_Target33_88_Jenis_%": (
            round((tot_terkumpul / tot_tgt * 100.0), 2) if tot_tgt > 0 else 0.0
        ),
        "Indeks_Dasar_0_100": round(mean_dasar, 2),
        "Indeks_Final_Disesuaikan_0_100": round(mean_final, 2),
        "Penyesuaian_Poin": round(mean_final - mean_dasar, 2),
    })
    return pd.DataFrame(rows)


def make_bell_curve(dfp, title, label_col="nm_perpustakaan"):
    fig = go.Figure()
    fig.update_layout(
        title=title,
        xaxis_title="Skor (0–100)",
        yaxis_title="Kepadatan",
        margin=dict(l=30, r=20, t=50, b=30),
    )
    if dfp is None or dfp.empty or "Indeks_Dasar_0_100" not in dfp.columns:
        fig.add_annotation(
            text="Tidak ada data",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
        )
        return fig

    x = (
        pd.to_numeric(dfp["Indeks_Dasar_0_100"], errors="coerce")
        .dropna()
        .values
    )
    if len(x) < 2:
        return fig

    mu, sigma = float(np.mean(x)), max(float(np.std(x, ddof=1)), 1e-3)
    xs = np.linspace(
        max(0.0, float(np.min(x)) - 5), min(100.0, float(np.max(x)) + 5), 200
    )
    pdf = (1.0 / (sigma * np.sqrt(2 * np.pi))) * np.exp(
        -0.5 * ((xs - mu) / sigma) ** 2
    )

    fig.add_trace(
        go.Scatter(x=xs, y=pdf, mode="lines", name="Kurva Normal (Fit)")
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=np.zeros_like(x),
            mode="markers",
            name="Data Riil",
            marker=dict(size=6, opacity=0.6),
        )
    )
    for q_val, q_lbl in [
        (np.percentile(x, 25), "Q1"),
        (np.median(x), "Median"),
        (np.percentile(x, 75), "Q3"),
    ]:
        fig.add_vline(
            x=float(q_val),
            line_dash="dash",
            annotation_text=f"{q_lbl}: {q_val:.1f}",
        )
    return fig


def to_excel_bytes(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet, index=False)
    return output.getvalue()


# ============================================================
# 4) TAMPILAN APLIKASI (STREAMLIT NATIVE UI)
# ============================================================
st.title(
    "📊 IPLM 2025 — Perhitungan Indeks Pembangunan Literasi Masyarakat"
)
st.caption("Target Sampel: 33.88% per Jenis | Skala Absolut Tanpa Relatif")

df_all, df_raw, pop_kab, pop_prov, pop_khusus, meta, status_msg = (
    load_all_data()
)

if df_all is None:
    st.error(
        f"Gagal memuat file dataset. Pastikan semua file Excel ada di repositori: {status_msg}"
    )
    st.stop()

# Sidebar Filters
st.sidebar.header("Filter Wilayah")
prov_list = ["(Semua)"] + sorted(
    [p for p in df_all["PROV_DISP"].dropna().unique() if p]
)
selected_prov = st.sidebar.selectbox("Provinsi", prov_list)

if selected_prov != "(Semua)":
    kab_list = ["(Semua)"] + sorted([
        k
        for k in df_all[df_all["PROV_DISP"] == selected_prov][
            "KAB_DISP"
        ].dropna().unique()
        if k
    ])
else:
    kab_list = ["(Semua)"] + sorted(
        [k for k in df_all["KAB_DISP"].dropna().unique() if k]
    )
selected_kab = st.sidebar.selectbox("Kabupaten / Kota", kab_list)

kew_list = ["(Semua)"] + sorted(
    [w for w in df_all["KEW_NORM"].dropna().unique() if w]
)
selected_kew = st.sidebar.selectbox("Kewenangan", kew_list, index=0)

# Filter Dataframe
df_filtered = df_all.copy()
if selected_prov != "(Semua)":
    df_filtered = df_filtered[df_filtered["PROV_DISP"] == selected_prov]
if selected_kab != "(Semua)":
    df_filtered = df_filtered[df_filtered["KAB_DISP"] == selected_kab]
if selected_kew != "(Semua)":
    df_filtered = df_filtered[df_filtered["KEW_NORM"] == selected_kew]

if df_filtered.empty:
    st.warning("Tidak ada data untuk kombinasi filter yang dipilih.")
    st.stop()

# Perhitungan
kew_calc = (
    selected_kew
    if selected_kew != "(Semua)"
    else ("KAB/KOTA" if selected_kab != "(Semua)" else "PROVINSI")
)
faktor_jenis = build_faktor_wilayah_jenis(
    df_filtered, pop_kab, pop_prov, pop_khusus, kew_calc
)
agg_jenis = build_agg_wilayah_jenis(df_filtered, faktor_jenis, kew_calc)
agg_total = build_agg_wilayah_total_from_jenis(agg_jenis, kew_calc)
summary_jenis = build_summary_per_jenis(agg_jenis, agg_total)

# KPI Metric Cards
row_keseluruhan = summary_jenis[summary_jenis["Jenis"] == "keseluruhan"].iloc[0]
c1, c2, c3, c4 = st.columns(4)
c1.metric(
    "IPLM Final (Target 33.88%)",
    f"{row_keseluruhan['Indeks_Final_Disesuaikan_0_100']:.2f}",
)
c2.metric(
    "Indeks Dasar (Sebelum Penyesuaian)",
    f"{row_keseluruhan['Indeks_Dasar_0_100']:.2f}",
)
c3.metric(
    "Coverage Target Sampel",
    f"{row_keseluruhan['Coverage_Target33_88_Jenis_%']:.1f}%",
)
c4.metric(
    "Total Responden / Perpus", f"{int(row_keseluruhan['Total_Perpus']):,}"
)

st.divider()

# Tab Layout
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Ringkasan & Agregat",
    "🔔 Kurva Distribusi (Bell Curve)",
    "📑 Kecukupan Sampel",
    "📥 Unduh Laporan",
])

with tab1:
    st.subheader("Ringkasan per Jenis Perpustakaan & Total")
    st.dataframe(summary_jenis, use_container_width=True)

    st.subheader("Agregat Wilayah")
    st.dataframe(agg_total, use_container_width=True)

with tab2:
    st.subheader("Distribusi Sebaran Indeks Dasar (Bell Curve)")
    c_bell1, c_bell2, c_bell3 = st.columns(3)
    with c_bell1:
        st.plotly_chart(
            make_bell_curve(
                df_filtered[df_filtered["_dataset"] == "umum"],
                "Perpustakaan Umum",
            ),
            use_container_width=True,
        )
    with c_bell2:
        st.plotly_chart(
            make_bell_curve(
                df_filtered[df_filtered["_dataset"] == "sekolah"],
                "Perpustakaan Sekolah",
            ),
            use_container_width=True,
        )
    with c_bell3:
        st.plotly_chart(
            make_bell_curve(
                df_filtered[df_filtered["_dataset"] == "khusus"],
                "Perpustakaan Khusus",
            ),
            use_container_width=True,
        )

with tab3:
    st.subheader("Evaluasi Pemenuhan Sampel 33.88%")
    st.dataframe(faktor_jenis, use_container_width=True)

with tab4:
    st.subheader("Download Hasil Olah Data")
    excel_data = to_excel_bytes({
        "Ringkasan": summary_jenis,
        "Agregat_Wilayah": agg_total,
        "Kecukupan_Sampel": faktor_jenis,
    })
    st.download_button(
        label="📥 Download Data Lengkap (.xlsx)",
        data=excel_data,
        file_name="Laporan_Kalkulasi_IPLM_2025.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
