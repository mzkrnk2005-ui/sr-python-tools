"""
社会保険料計算モジュール（令和8年度）
ノートブック「社会保険料率(協会けんぽ)取得2026.ipynb」をモジュール化
"""

import io
import json
import math
import requests
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

# ============================================================
# 定数
# ============================================================

# 子ども・子育て支援金（2026年5月支払い分以降）
SHIEN_KIN_RATE = 0.0023

# ============================================================
# 雇用保険料率（令和8年度）
# 出典：厚生労働省公式PDFより自動取得（雇用保険料率_最新.json）
# ※毎年3月のGitHub Actions自動更新で最新値に更新される
# ============================================================

# 労働者負担
KOYO_RATES_WORKER_DEFAULT = {
    "一般":     0.005,   # 5/1,000
    "建設":     0.006,   # 6/1,000
    "農林水産": 0.006,   # 6/1,000
}

# 事業主負担（令和8年度）
# ※毎年変更される可能性があるため、必ず厚労省公式資料で確認すること
KOYO_RATES_EMPLOYER = {
    "一般":     0.0085,  # 8.5/1,000
    "建設":     0.0105,  # 10.5/1,000（二事業4.5/1,000を含む）
    "農林水産": 0.0095,  # 9.5/1,000
}

PREFS = [
    "北海道", "青森", "岩手", "宮城", "秋田", "山形", "福島",
    "茨城", "栃木", "群馬", "埼玉", "千葉", "東京", "神奈川",
    "新潟", "富山", "石川", "福井", "山梨", "長野", "岐阜",
    "静岡", "愛知", "三重", "滋賀", "京都", "大阪", "兵庫",
    "奈良", "和歌山", "鳥取", "島根", "岡山", "広島", "山口",
    "徳島", "香川", "愛媛", "高知", "福岡", "佐賀", "長崎",
    "熊本", "大分", "宮崎", "鹿児島", "沖縄",
]

# ============================================================
# 雇用保険労働者負担をJSONから読み込む
# （GitHub Actionsによる毎年3月の自動更新に対応）
# ============================================================

def load_koyo_rates_worker() -> dict:
    """
    雇用保険料率_最新.json から労働者負担の料率を読み込む。
    ファイルが存在しない場合はデフォルト値を返す。
    """
    json_path = Path(__file__).parent.parent / "雇用保険料率_最新.json"
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        return data.get("雇用保険料率_労働者負担", KOYO_RATES_WORKER_DEFAULT)
    except Exception:
        return KOYO_RATES_WORKER_DEFAULT


# 起動時に一度だけ読み込む
KOYO_RATES_WORKER = load_koyo_rates_worker()

# ============================================================
# 端数処理
# ============================================================

def round_shakai(yen: float) -> int:
    """健康保険・厚生年金：50銭以上切り上げ（四捨五入）"""
    return int(Decimal(str(yen)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def floor_shakai(yen: float) -> int:
    """雇用保険：切り捨て"""
    return math.floor(yen)


# ============================================================
# 協会けんぽ料率取得（健康保険・厚生年金）
# ============================================================

def fetch_rates() -> dict:
    """
    協会けんぽ公式サイトから令和8年度 都道府県別保険料率を取得。
    戻り値：{ 都道府県名: { 健康保険料率, 健康保険料率_介護, 厚生年金保険料率 } }
    """
    url = "https://www.kyoukaikenpo.or.jp/assets/r8ippan3.xlsx"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()

    excel_data = io.BytesIO(response.content)
    df_raw = pd.read_excel(excel_data, sheet_name=None)

    rates = {}
    for pref in PREFS:
        if pref not in df_raw:
            continue
        try:
            rate_row = df_raw[pref].iloc[6]
            rates[pref] = {
                "健康保険料率":      float(rate_row.iloc[5]),
                "健康保険料率_介護": float(rate_row.iloc[7]),
                "厚生年金保険料率":  float(rate_row.iloc[11]),
            }
        except Exception:
            pass

    return rates


# ============================================================
# 社会保険料計算（1名分）
# ============================================================

def calc_employee(row: dict, rates: dict, shien_kin: bool = True) -> dict:
    """
    1名分の社会保険料（従業員負担・事業主負担の両方）を計算して辞書で返す。

    Parameters
    ----------
    row       : 従業員データ（氏名, 年齢, 都道府県, 標準報酬月額, 業種）
    rates     : fetch_rates() の戻り値（協会けんぽ料率）
    shien_kin : True → 子ども・子育て支援金を加算（5月支払い分以降）
    """
    kyuyo  = int(row["標準報酬月額"])
    age    = int(row["年齢"])
    pref   = row["都道府県"]
    gyoshu = row.get("業種", "一般")

    pref_rates    = rates.get(pref, rates.get("東京", {}))
    koyo_worker   = KOYO_RATES_WORKER.get(gyoshu, KOYO_RATES_WORKER["一般"])
    koyo_employer = KOYO_RATES_EMPLOYER.get(gyoshu, KOYO_RATES_EMPLOYER["一般"])

    # 40歳以上は介護保険料率（健康保険料率_介護）を適用
    kenko_rate  = pref_rates["健康保険料率_介護"] if age >= 40 else pref_rates["健康保険料率"]
    nenkin_rate = pref_rates["厚生年金保険料率"]

    # ========== 従業員負担 ==========
    kenko_ko  = round_shakai(kyuyo * kenko_rate / 2)
    shien_ko  = round_shakai(kyuyo * SHIEN_KIN_RATE / 2) if shien_kin else 0
    nenkin_ko = round_shakai(kyuyo * nenkin_rate / 2)
    koyo_ko   = floor_shakai(kyuyo * koyo_worker)
    total_ko  = kenko_ko + shien_ko + nenkin_ko + koyo_ko

    # ========== 事業主負担 ==========
    kenko_sha  = round_shakai(kyuyo * kenko_rate / 2)
    shien_sha  = round_shakai(kyuyo * SHIEN_KIN_RATE / 2) if shien_kin else 0
    nenkin_sha = round_shakai(kyuyo * nenkin_rate / 2)
    koyo_sha   = floor_shakai(kyuyo * koyo_employer)
    total_sha  = kenko_sha + shien_sha + nenkin_sha + koyo_sha

    return {
        # 従業員負担
        "健康保険料(本人)":   kenko_ko,
        "子育て支援金(本人)": shien_ko,
        "厚生年金料(本人)":   nenkin_ko,
        "雇用保険料(本人)":   koyo_ko,
        "控除合計(本人)":     total_ko,
        "手取り概算":         kyuyo - total_ko,
        # 事業主負担
        "健康保険料(会社)":   kenko_sha,
        "子育て支援金(会社)": shien_sha,
        "厚生年金料(会社)":   nenkin_sha,
        "雇用保険料(会社)":   koyo_sha,
        "会社負担合計":        total_sha,
        # 合計人件費
        "人件費総額":          kyuyo + total_sha,
    }


# ============================================================
# DataFrame 一括計算
# ============================================================

def calc_dataframe(df: pd.DataFrame, rates: dict, shien_kin: bool = True) -> pd.DataFrame:
    """全従業員分を計算して元のDataFrameに列を追加して返す"""
    result_cols = df.apply(
        lambda r: pd.Series(calc_employee(r.to_dict(), rates, shien_kin)),
        axis=1,
    )
    return pd.concat([df.reset_index(drop=True), result_cols], axis=1)
