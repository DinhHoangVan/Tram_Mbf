#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tra cứu thông tin trạm MobiFone - Streamlit
Chỉ nhập tên trạm → bấm nút → hiện:
  🔴🔋  BDTM23-HCMBCH03 ( 3G: 156, 4G: 337, 5G: 197 )  [UT1]
"""

import pickle
from pathlib import Path

import pandas as pd
import streamlit as st

# ================== CẤU HÌNH ==================
st.set_page_config(
    page_title="Tra cứu trạm MobiFone",
    page_icon="📡",
    layout="centered",
)

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "Ket_qua_Phu_luc"
if not DATA_DIR.is_dir():
    DATA_DIR = BASE / "data" if (BASE / "data").is_dir() else BASE

FILE_UU_TIEN = DATA_DIR / "Site_Uu_Tien_26_1_2026.xlsx"
FILE_AC_QUY = DATA_DIR / "MBF_Giam_tru_MLL_MASTER_DAY_DU.xlsx"
FILE_TEN = DATA_DIR / "Danh sach cell name old_new v1.xlsx"
FILE_THIEU = DATA_DIR / "Tram_thieu_Uu_tien.xlsx"
FILE_VLR = DATA_DIR / "VLR_1.xlsx"

CACHE_DIR = BASE / "cache_pkl"
CACHE_DIR.mkdir(exist_ok=True)
PKL_PATH = CACHE_DIR / "tram_data.pkl"

NO_BATTERY_KEYWORDS = [
    "trạm cran không có ắc quy",
    "tram cran không có ắc quy",
    "trạm outdoor không có ắc quy",
    "cran không có ắc quy",
    "smallcell",
    "trạm trong công ty không cho phép chạy máy phát điện",
    "trạm cran không có accu",
    "trạm cran nằm trong nhà xưởng, không có ắc quy",
    "trang bị mới bộ cảnh báo ngoài",
    "không trang bị",
    "không trang bị lý do trạm cran",
    "không trang bị do trạm dùng tủ enclosure",
    "trạm cran không có ắc quy",
    "không trang bị do trạm resort",
]

PRIORITY_ICON = {
    "UT1": "🔴",
    "UT2": "🟡",
    "UT3": "🟢",
}


def normalize(s) -> str:
    if pd.isna(s):
        return ""
    return str(s).strip().upper()


def is_no_battery(reason) -> bool:
    if not reason or pd.isna(reason):
        return False
    r = str(reason).lower().strip()
    return any(kw in r for kw in NO_BATTERY_KEYWORDS)


def build_from_excel():
    """Đọc Excel → dict tra cứu."""
    required = {
        "Site_Uu_Tien_26_1_2026.xlsx": FILE_UU_TIEN,
        "MBF_Giam_tru_MLL_MASTER_DAY_DU.xlsx": FILE_AC_QUY,
        "Danh sach cell name old_new v1.xlsx": FILE_TEN,
        "VLR_1.xlsx": FILE_VLR,
    }
    missing = [n for n, p in required.items() if not p.exists()]
    if missing:
        return None, missing

    # 1. Ưu tiên
    df_ut = pd.read_excel(FILE_UU_TIEN, sheet_name="Sheet1", usecols=["Sitename", "Ưu tiên"])
    df_ut["Sitename"] = df_ut["Sitename"].apply(normalize)
    df_ut = df_ut.drop_duplicates(subset=["Sitename"], keep="first")
    map_ut = dict(zip(df_ut["Sitename"], df_ut["Ưu tiên"]))

    # 2. Ắc quy
    df_aq = pd.read_excel(FILE_AC_QUY, usecols=["Tên trạm", "Nguyên nhân"])
    df_aq["Tên trạm"] = df_aq["Tên trạm"].apply(normalize)
    no_battery = {
        row["Tên trạm"]
        for _, row in df_aq.iterrows()
        if is_no_battery(row["Nguyên nhân"])
    }

    # 3. Tên cũ ↔ mới
    df_ten = pd.read_excel(
        FILE_TEN, sheet_name="DS_CHI_TIET", usecols=["Site name OLD", "Site name NEW"]
    )
    df_ten["Site name OLD"] = df_ten["Site name OLD"].apply(normalize)
    df_ten["Site name NEW"] = df_ten["Site name NEW"].apply(normalize)
    df_ten = df_ten.dropna(subset=["Site name OLD", "Site name NEW"])
    df_ten = df_ten[df_ten["Site name OLD"] != ""]
    df_ten = df_ten.drop_duplicates(subset=["Site name OLD"], keep="first")

    if FILE_THIEU.exists():
        df_th = pd.read_excel(FILE_THIEU)
        df_th["ten_cu"] = df_th["ten_cu"].apply(normalize)
        df_th["ten_moi"] = df_th["ten_moi"].apply(normalize)
        df_th = df_th.drop_duplicates(subset=["ten_cu"], keep="first")
        extra = df_th[~df_th["ten_cu"].isin(df_ten["Site name OLD"])]
        extra = extra.rename(columns={"ten_cu": "Site name OLD", "ten_moi": "Site name NEW"})
        df_ten = pd.concat(
            [df_ten, extra[["Site name OLD", "Site name NEW"]]], ignore_index=True
        )

    map_old_to_new = dict(zip(df_ten["Site name OLD"], df_ten["Site name NEW"]))
    map_new_to_old = dict(zip(df_ten["Site name NEW"], df_ten["Site name OLD"]))

    # 4. VLR (bản ghi mới nhất)
    df_vlr = pd.read_excel(
        FILE_VLR, usecols=["NGAY", "SITENAME", "VLR_3G", "VLR_4G", "VLR_SUBS_5G_SUPPORT"]
    )
    df_vlr["SITENAME"] = df_vlr["SITENAME"].apply(normalize)
    df_vlr["NGAY"] = pd.to_datetime(df_vlr["NGAY"], dayfirst=True, errors="coerce")
    df_vlr = df_vlr.sort_values("NGAY").groupby("SITENAME", as_index=False).last()
    map_vlr = {}
    for _, row in df_vlr.iterrows():
        map_vlr[row["SITENAME"]] = {
            "3G": int(row["VLR_3G"]) if pd.notna(row["VLR_3G"]) else 0,
            "4G": int(row["VLR_4G"]) if pd.notna(row["VLR_4G"]) else 0,
            "5G": int(row["VLR_SUBS_5G_SUPPORT"]) if pd.notna(row["VLR_SUBS_5G_SUPPORT"]) else 0,
        }

    data = {
        "map_ut": map_ut,
        "no_battery": no_battery,
        "map_old_to_new": map_old_to_new,
        "map_new_to_old": map_new_to_old,
        "map_vlr": map_vlr,
    }
    return data, []


def save_pkl(data, path=PKL_PATH):
    with open(path, "wb") as f:
        pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)


def load_pkl(path=PKL_PATH):
    if not path.exists():
        return None
    with open(path, "rb") as f:
        return pickle.load(f)


def excel_newer_than_pkl() -> bool:
    if not PKL_PATH.exists():
        return True
    t = PKL_PATH.stat().st_mtime
    for f in (FILE_UU_TIEN, FILE_AC_QUY, FILE_TEN, FILE_VLR, FILE_THIEU):
        if f.exists() and f.stat().st_mtime > t:
            return True
    return False


@st.cache_data(show_spinner="⏳ Đang tải dữ liệu...")
def load_data(force: bool = False):
    if not force and PKL_PATH.exists() and not excel_newer_than_pkl():
        data = load_pkl()
        if data is not None:
            return data, [], "pkl"

    data, missing = build_from_excel()
    if missing:
        return None, missing, "error"
    save_pkl(data)
    return data, [], "excel"


def lookup(query: str, data: dict) -> str:
    """
    Trả về 1 dòng:
      🔴🔋  BDTM23-HCMBCH03 ( 3G: 156, 4G: 337, 5G: 197 )  [UT1]
    """
    q = normalize(query)
    if not q:
        return ""

    map_ut = data["map_ut"]
    no_battery = data["no_battery"]
    map_old_to_new = data["map_old_to_new"]
    map_new_to_old = data["map_new_to_old"]
    map_vlr = data["map_vlr"]

    ten_cu, ten_moi = q, q
    if q in map_old_to_new:
        ten_cu, ten_moi = q, map_old_to_new[q]
    elif q in map_new_to_old:
        ten_cu, ten_moi = map_new_to_old[q], q

    ut = map_ut.get(ten_moi) or map_ut.get(ten_cu) or map_ut.get(q)
    icon_ut = PRIORITY_ICON.get(str(ut).upper(), "⚪") if ut else "⚪"
    ut_label = str(ut) if ut else "?"

    has_bat = not (ten_moi in no_battery or ten_cu in no_battery or q in no_battery)
    icon_bat = "🔋" if has_bat else "🪫"

    vlr = map_vlr.get(ten_moi) or map_vlr.get(ten_cu) or map_vlr.get(q)
    if vlr:
        vlr_str = f"3G: {vlr['3G']}, 4G: {vlr['4G']}, 5G: {vlr['5G']}"
    else:
        vlr_str = "3G: -, 4G: -, 5G: -"

    name = f"{ten_cu}-{ten_moi}" if ten_cu != ten_moi else ten_moi
    return f"{icon_ut}{icon_bat}  {name} ( {vlr_str} )  [{ut_label}]"


# ================== GIAO DIỆN ==================
st.title("📡 Tra cứu trạm MobiFone")
st.caption("Ký hiệu: 🔴=UT1  🟡=UT2  🟢=UT3  |  🔋=có ắc quy  🪫=không có ắc quy")

with st.sidebar:
    st.header("Cache")
    if PKL_PATH.exists():
        mb = PKL_PATH.stat().st_size / (1024 * 1024)
        st.success(f"pkl: {mb:.1f} MB")
    else:
        st.info("Chưa có pkl")
    if st.button("🔄 Rebuild từ Excel", use_container_width=True):
        st.cache_data.clear()
        force_rebuild = True
    else:
        force_rebuild = False

data, missing, source = load_data(force=force_rebuild)

if missing:
    st.error("Không tìm thấy file Excel:")
    for n in missing:
        st.write(f"- `{n}`")
    st.info(f"Thư mục: `{DATA_DIR}`")
    st.stop()

# --- Form: chỉ nhập tên trạm + nút ---
with st.form("tra_cuu"):
    ten_tram = st.text_input("Tên trạm", placeholder="Ví dụ: BDTM23")
    submitted = st.form_submit_button("Tra cứu", type="primary", use_container_width=True)

if submitted:
    if not ten_tram or not ten_tram.strip():
        st.warning("Vui lòng nhập tên trạm.")
    else:
        q = ten_tram.strip()
        line = lookup(q, data)
        st.markdown(f"**{normalize(q)}**")
        st.code(line, language=None)

st.divider()
st.caption(
    f"Nguồn: {'📦 pickle' if source == 'pkl' else '📊 Excel'}  ·  "
    f"Ưu tiên {len(data['map_ut']):,}  ·  "
    f"Không ắc quy {len(data['no_battery']):,}  ·  "
    f"Mapping {len(data['map_old_to_new']):,}  ·  "
    f"VLR {len(data['map_vlr']):,}"
)
