"""
雇用保険料率 自動更新スクリプト（令和8年度〜）
GitHub Actions から毎年3月1日に自動実行される。

処理の流れ：
  1. 厚労省ページから最新PDFのURLを自動検索
  2. PDFをダウンロードしてpdfplumberで料率を抽出
  3. 「雇用保険料率_最新.json」と「雇用保険料率_令和〇年度.json」に保存
  4. 終了コード 0 = 成功、1 = 失敗（GitHub Actionsがメール送信に使用）
"""

import io
import json
import re
import sys
import requests
from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path
import pdfplumber

# ============================================================
# 設定
# ============================================================
MHLW_URL = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000108634.html"
HEADERS  = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# このスクリプトと同じディレクトリ（リポジトリのルート）に保存
REPO_ROOT = Path(__file__).parent


# ============================================================
# 前年度の料率を読み込む（比較用）
# ============================================================
def load_prev_rates() -> dict:
    json_path = REPO_ROOT / "雇用保険料率_最新.json"
    try:
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        return data.get("雇用保険料率_労働者負担", {})
    except Exception:
        return {}


# ============================================================
# 厚労省ページからPDF URLを動的に取得
# ============================================================
def find_pdf_url() -> str:
    response = requests.get(MHLW_URL, headers=HEADERS, timeout=30)
    response.raise_for_status()
    soup = BeautifulSoup(response.content, "html.parser")

    for link in soup.find_all("a", href=True):
        href = link.get("href", "")
        text = link.get_text(strip=True)
        # 「雇用保険料率」を含むPDFリンクを探す
        if "雇用保険料率" in text and ".pdf" in href.lower():
            return href if href.startswith("http") else f"https://www.mhlw.go.jp{href}"

    raise ValueError("雇用保険料率のPDFリンクが見つかりませんでした。厚労省ページの構造が変わった可能性があります。")


# ============================================================
# PDFダウンロード
# ============================================================
def download_pdf(url: str) -> bytes:
    response = requests.get(url, headers=HEADERS, timeout=60)
    response.raise_for_status()
    return response.content


# ============================================================
# PDFから料率を抽出
# ============================================================
def extract_rates(pdf_bytes: bytes) -> dict:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(
            page.extract_text() or "" for page in pdf.pages
        )

    rates = {}

    # 方法①：正規表現で一括検索（一般・建設向き）
    text_oneline = " ".join(text.splitlines())
    for gyoshu, pattern in {
        "一般": r"一般の事業\s+(\d+(?:\.\d+)?)/1,000",
        "建設": r"建設の事業\s+(\d+(?:\.\d+)?)/1,000",
    }.items():
        m = re.search(pattern, text_oneline)
        if m:
            rates[gyoshu] = float(m.group(1)) / 1000

    # 方法②：行単位（農林水産は改行で分断されやすいため）
    target = None
    for line in text.splitlines():
        line = line.strip()
        if "農林水産" in line:
            target = "農林水産"
        if target == "農林水産" and re.match(r"^\d+/1,000", line):
            m = re.match(r"^(\d+(?:\.\d+)?)/1,000", line)
            if m and "農林水産" not in rates:
                rates["農林水産"] = float(m.group(1)) / 1000
                target = None

    if not rates:
        raise ValueError("料率を抽出できませんでした。PDFの書式が変わった可能性があります。")

    return rates


# ============================================================
# 年度の文字列を生成（例：令和8年度）
# ============================================================
def get_fiscal_year() -> str:
    now = datetime.now()
    # 1〜3月は前年度扱い
    year = now.year if now.month >= 4 else now.year - 1
    reiwa = year - 2018
    return f"令和{reiwa}年度"


# ============================================================
# JSONに保存
# ============================================================
def save_json(rates: dict, fiscal_year: str) -> None:
    data = {
        "年度":                    fiscal_year,
        "取得日":                  datetime.now().strftime("%Y/%m/%d"),
        "雇用保険料率_労働者負担": rates,
    }

    # ① 年度別ファイル（記録用）
    year_path = REPO_ROOT / f"雇用保険料率_{fiscal_year}.json"
    with open(year_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"   保存: {year_path.name}")

    # ② 最新ファイル（Streamlitアプリが常にここを読む）
    latest_path = REPO_ROOT / "雇用保険料率_最新.json"
    with open(latest_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"   保存: {latest_path.name}")


# ============================================================
# メイン
# ============================================================
def main():
    fiscal_year = get_fiscal_year()
    prev_rates  = load_prev_rates()

    print(f"=== {fiscal_year} 雇用保険料率 自動更新 ===\n")

    try:
        print("① 厚労省ページからPDF URLを検索中...")
        pdf_url = find_pdf_url()
        print(f"   URL: {pdf_url}\n")

        print("② PDFをダウンロード中...")
        pdf_bytes = download_pdf(pdf_url)
        print(f"   サイズ: {len(pdf_bytes):,} bytes\n")

        print("③ 料率を抽出中...")
        rates = extract_rates(pdf_bytes)

        print("\n=== 抽出結果（労働者負担） ===")
        for gyoshu in ["一般", "農林水産", "建設"]:
            if gyoshu not in rates:
                print(f"  ⚠️  {gyoshu}：抽出できませんでした")
                continue
            rate = rates[gyoshu]
            prev = prev_rates.get(gyoshu)
            if prev is not None:
                diff = (rate - prev) * 1000
                mark = "↓" if diff < 0 else "↑" if diff > 0 else "→"
                print(f"  ✅ {gyoshu}：{int(rate*1000)}/1,000"
                      f"（前年 {int(prev*1000)}/1,000 {mark} {abs(diff):.1f}）")
            else:
                print(f"  ✅ {gyoshu}：{int(rate*1000)}/1,000")

        print("\n④ JSONファイルに保存中...")
        save_json(rates, fiscal_year)

        print(f"\n✅ 完了：{fiscal_year}の雇用保険料率を更新しました。")
        sys.exit(0)

    except Exception as e:
        print(f"\n❌ エラー：{e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
