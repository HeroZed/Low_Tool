# -*- coding: utf-8 -*-
"""
moj_law_diff.py — 全國法規資料庫「一鍵比對新舊條文」原型工具

用法：
    pip install requests beautifulsoup4 openpyxl

    單一法規：
        python moj_law_diff.py N0060010
        python moj_law_diff.py N0060010 --out 教育訓練規則_鑑別草稿.xlsx

    一次比對多部法規（每部法規各佔輸出 Excel 的一個分頁）：
        python moj_law_diff.py N0060010 N0060001 N0060008 --out 全部比對結果.xlsx

    或是把公司要追蹤的法規代碼存成一份清單檔（例如 我的法規清單.txt，
    一行一個代碼，可以用 # 開頭加註解），之後每次都用同一份清單執行：
        python moj_law_diff.py --list 我的法規清單.txt --out 全部比對結果.xlsx

    只想看某段期間內的修正（適合每季／每半年定期追蹤，不用每次全部重看）：
        python moj_law_diff.py --list 我的法規清單.txt --since 2025-06-01 --out 本次追蹤.xlsx
        python moj_law_diff.py --list 我的法規清單.txt --since 2026-09-08 --until 2026-12-31 --out 追蹤.xlsx
    --since 是區間起點，--until 不給的話預設是「今天」。程式會自動抓出
    「起點前最後生效的版本」跟「終點當時生效的版本」來比對，就算這段期間
    修正了不只一次，也會一次抓出所有異動；如果這段期間根本沒有修正，
    會直接顯示「沒有查到任何修正紀錄」。

這支程式做的事（對應建置計畫的「Phase 2：API 串接原型」）：
    1. 用法規代碼 pcode（例如職業安全衛生教育訓練規則＝N0060010）向全國法規資料庫
       抓「現行條文」(LawAll.aspx)。
    2. 自動到該法規的「沿革」頁 (LawHistory.aspx) 找出「上一次修正」的日期，
       再抓那個舊版本的全文 (LawOldVer.aspx)。
    3. 逐條比對新舊條文文字，抓出哪些條文「新增」「刪除」「修正」。
    4. 把結果整理成一份 Excel 草稿，欄位對齊你現有「法規鑑別表」的格式
       （序號 / 法規條文內容(修正前) / 法規條文內容(修正後) / 鑑別日期 ...），
       有異動的文字會用粗體＋底線標示，符合你表頭原本的規則。

注意：
    - 這支程式必須在「有一般網路連線」的電腦上執行（你自己的電腦即可）；
      不能在本次對話的雲端沙盒環境裡執行，因為那裡的對外連線有白名單限制，
      連不到 law.moj.gov.tw。
    - 目前用的是全國法規資料庫「網站頁面」本身的結構（LawAll.aspx /
      LawHistory.aspx / LawOldVer.aspx），已經用職業安全衛生教育訓練規則
      實際測試過，比對結果跟官方沿革公告的異動條文完全一致
      （第1、17、18、20、40、43條修正，新增第42-4條）。
      如果法規資料庫改版導致網頁結構變動，這支程式的解析部分可能需要更新。
"""

import argparse
import datetime as dt
import difflib
import re
import sys
import warnings

import requests
from bs4 import BeautifulSoup

BASE = "https://law.moj.gov.tw/LawClass"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; RegComplianceTool/0.1; internal use)"
}

ARTICLE_NO_RE = re.compile(r"第\s*[\d\-]+\s*條")

_WARNED_INSECURE = False


def http_get(url: str):
    """對全國法規資料庫發出 GET 請求。

    部分 Windows 環境（較新版 Python／OpenSSL）在驗證這個網站的憑證鏈結時，
    會出現 'Missing Subject Key Identifier' 這類相容性錯誤，但這不代表連線
    不安全，只是本機的憑證驗證邏輯對這條鏈結比較嚴格。這裡先用正常方式驗證，
    只有在確定是這種憑證驗證錯誤時，才退回不驗證憑證的方式重試，並印出提醒。
    """
    global _WARNED_INSECURE
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
    except requests.exceptions.SSLError:
        if not _WARNED_INSECURE:
            print(
                "提醒：本機的憑證驗證對 law.moj.gov.tw 回報相容性錯誤"
                "（常見於新版 Python／Windows 環境），已自動改用不驗證憑證的"
                "方式重新連線。這裡連的是公開的政府法規網站，僅供內部讀取"
                "條文使用，風險低；如果之後要正式上線，建議改用 truststore"
                "套件修正，而不是長期關閉憑證驗證。",
                file=sys.stderr,
            )
            _WARNED_INSECURE = True
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            resp = requests.get(url, headers=HEADERS, timeout=20, verify=False)
    resp.raise_for_status()
    return resp


def normalize_article_no(raw: str) -> str:
    """把 '本條文有附件 第 3 條' 這種標籤正規化成 '第3條'，避免跟沒有附件標籤的
    舊版本條號對不起來。"""
    m = ARTICLE_NO_RE.search(raw)
    text = m.group(0) if m else raw
    return re.sub(r"\s+", "", text)


def fetch_articles(url: str) -> dict:
    """抓某個法規頁面（現行或歷史版本），回傳 {條號: 條文內容} 的有序 dict。"""
    resp = http_get(url)
    resp.encoding = resp.apparent_encoding or "utf-8"
    soup = BeautifulSoup(resp.text, "html.parser")

    articles = {}
    for row in soup.select(".law-reg .row"):
        no_el = row.select_one(".col-no")
        data_el = row.select_one(".col-data")
        if not no_el or not data_el:
            continue
        raw_no = no_el.get_text(strip=True)
        if "條" not in raw_no:
            continue  # 跳過「第 X 章」這類章節標題列
        no = normalize_article_no(raw_no)
        text = data_el.get_text("\n", strip=True)
        text = re.sub(r"\n{2,}", "\n", text)
        articles[no] = text
    return articles


def roc_to_yyyymmdd(s: str):
    """把「民國115年06月25日」這類字串轉成可直接比大小的 '20260625'。"""
    if not s:
        return None
    m = re.search(r"(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日", s)
    if not m:
        return None
    roc_year, month, day = (int(x) for x in m.groups())
    return f"{roc_year + 1911:04d}{month:02d}{day:02d}"


def fmt_date(d: str) -> str:
    """'20260625' -> '2026-06-25'，給人看的顯示格式。"""
    if not d or len(d) != 8:
        return d or ""
    return f"{d[0:4]}-{d[4:6]}-{d[6:8]}"


def list_versions(pcode: str) -> dict:
    """讀沿革頁，整理出這部法規『每一版從什麼時候開始生效』的清單（新到舊排序），
    包含目前現行版本。之後不管是要跟『上一版』比、還是跟『某個日期當時的版本』比，
    都是從這份清單裡挑兩個版本出來即可。"""
    url = f"{BASE}/LawHistory.aspx?pcode={pcode}"
    resp = http_get(url)
    resp.encoding = resp.apparent_encoding or "utf-8"
    soup = BeautifulSoup(resp.text, "html.parser")

    name_el = soup.find(string=re.compile("法規名稱"))
    law_name = None
    if name_el:
        row = name_el.find_parent("tr") or name_el.find_parent()
        if row:
            law_name = row.get_text(" ", strip=True).replace("法規名稱：", "").strip()
            # 法規名稱欄位右邊常附一個「EN」英文版連結，把它去掉
            law_name = re.sub(r"\s*EN\s*$", "", law_name).strip()

    # 已修正過的法規顯示「修正日期」；從沒修正過的只有「發布日期」
    amend_date_display = None
    for label in ("修正日期", "發布日期"):
        el = soup.find(string=re.compile(label))
        if not el:
            continue
        row = el.find_parent("tr") or el.find_parent()
        if row:
            amend_date_display = row.get_text(" ", strip=True).replace(f"{label}：", "").strip()
            break
    current_date = roc_to_yyyymmdd(amend_date_display)

    versions = []
    if current_date:
        versions.append({"date": current_date, "source": "current", "lnndate": None, "lser": None})

    seen_dates = {current_date} if current_date else set()
    for a in soup.find_all("a", href=re.compile(r"LawOldVer\.aspx\?pcode=")):
        href = a.get("href", "")
        m_date = re.search(r"lnndate=(\d{8})", href)
        if not m_date:
            continue
        d = m_date.group(1)
        if d in seen_dates:
            continue  # 同一個生效日期可能因為附表等原因重複列出，只留一筆
        seen_dates.add(d)
        m_lser = re.search(r"lser=(\d+)", href)
        versions.append(
            {"date": d, "source": "old", "lnndate": d, "lser": m_lser.group(1) if m_lser else "001"}
        )

    versions.sort(key=lambda v: v["date"], reverse=True)  # 新到舊
    return {"law_name": law_name, "amend_date_display": amend_date_display, "versions": versions}


def version_asof(versions_desc: list, target_date: str, strict_before: bool = False):
    """versions_desc 是新到舊排序的版本清單。找出『在 target_date 這一天，
    是哪一版正在生效』；strict_before=True 則是找『target_date 之前最後生效的版本』
    （用來當比對的起點，避免剛好卡在生效日當天算錯邊）。找不到就回傳 None。"""
    for v in versions_desc:
        if (v["date"] < target_date) if strict_before else (v["date"] <= target_date):
            return v
    return None


def inline_diff(old_text: str, new_text: str):
    """逐字比對，回傳 (old_runs, new_runs)。
    每個 run 是 (text, changed:bool)，changed=True 的片段對應
    「刪除文字」（在 old_runs 裡）或「增加文字」（在 new_runs 裡），
    也就是你表頭要求的粗體底線範圍。"""
    sm = difflib.SequenceMatcher(None, old_text, new_text, autojunk=False)
    old_runs, new_runs = [], []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            old_runs.append((old_text[i1:i2], False))
            new_runs.append((new_text[j1:j2], False))
        else:
            if i1 != i2:
                old_runs.append((old_text[i1:i2], True))
            if j1 != j2:
                new_runs.append((new_text[j1:j2], True))
    return old_runs, new_runs


def fetch_version_articles(pcode: str, version: dict) -> dict:
    if version["source"] == "current":
        return fetch_articles(f"{BASE}/LawAll.aspx?pcode={pcode}")
    return fetch_articles(
        f"{BASE}/LawOldVer.aspx?pcode={pcode}&lnndate={version['lnndate']}&lser={version['lser']}"
    )


def _empty_result(pcode, law_name, amend_date, period=None, old_v=None, new_v=None, note=None):
    return {
        "pcode": pcode,
        "law_name": law_name,
        "amend_date": amend_date,
        "period": period,
        "old_version_date": old_v["date"] if old_v else None,
        "new_version_date": new_v["date"] if new_v else None,
        "amendment_dates_in_range": None,
        "note": note,
        "changed_rows": [],
    }


def build_report(pcode: str, since: str = None, until: str = None):
    """比對某一部法規的新舊條文。

    - 不給 since：跟預設一樣，抓「現行版本」對「上一次修正前的版本」。
    - 給 since（可選 until，預設今天）：改成抓『since 這天之前最後生效的版本』
      對『until 這天當時生效的版本』，這樣可以抓出整個區間內、不管中間修正過
      幾次的所有變動，適合每季／每半年定期追蹤用。
    """
    info = list_versions(pcode)
    law_name = info["law_name"]
    if not law_name:
        raise RuntimeError(f"抓不到 pcode={pcode} 的法規名稱，請確認代碼是否正確。")

    versions = info["versions"]
    amend_date = info["amend_date_display"]

    if since is None:
        if len(versions) < 2:
            return _empty_result(pcode, law_name, amend_date)
        new_v, old_v, note = versions[0], versions[1], None
    else:
        until = until or dt.date.today().strftime("%Y%m%d")
        old_v = version_asof(versions, since, strict_before=True)
        new_v = version_asof(versions, until, strict_before=False)
        note = None
        if new_v is None:
            raise RuntimeError(
                f"{fmt_date(until)} 這個日期比這部法規最早的資料還早，查不到任何版本。"
            )
        if old_v is None:
            old_v = versions[-1]
            note = (
                f"這個資料庫對此法規最早只能追溯到 {fmt_date(old_v['date'])}，"
                f"更早之前的版本無法比對，已改用這個最早版本當基準。"
            )

    period = (since, until) if since else None

    if old_v["date"] == new_v["date"]:
        return _empty_result(pcode, law_name, amend_date, period, old_v, new_v, note)

    old = fetch_version_articles(pcode, old_v)
    current = fetch_version_articles(pcode, new_v)

    all_nos = list(dict.fromkeys(list(old.keys()) + list(current.keys())))

    rows = []
    for no in all_nos:
        old_text = old.get(no)
        new_text = current.get(no)
        if old_text == new_text:
            continue  # 條文沒變動，不需要重新鑑別
        if old_text is None:
            status = "新增"
            old_runs, new_runs = [], [(new_text, True)]
        elif new_text is None:
            status = "刪除"
            old_runs, new_runs = [(old_text, True)], []
        else:
            status = "修正"
            old_runs, new_runs = inline_diff(old_text, new_text)
        rows.append(
            {
                "article_no": no,
                "status": status,
                "old_runs": old_runs,
                "new_runs": new_runs,
            }
        )

    amendment_dates_in_range = None
    if since:
        amendment_dates_in_range = [
            v["date"] for v in reversed(versions) if old_v["date"] < v["date"] <= new_v["date"]
        ]

    return {
        "pcode": pcode,
        "law_name": law_name,
        "amend_date": amend_date,
        "period": period,
        "old_version_date": old_v["date"],
        "new_version_date": new_v["date"],
        "amendment_dates_in_range": amendment_dates_in_range,
        "note": note,
        "changed_rows": rows,
    }


def print_console_report(report: dict):
    print(f"法規名稱：{report['law_name']}（pcode={report['pcode']}）")
    print(f"現行修正日期：{report['amend_date']}")

    if report.get("note"):
        print(f"提醒：{report['note']}")

    if report["old_version_date"] is None:
        print("找不到更早的歷史版本，可能是第一次發布，或此法規尚無修正紀錄。")
        return

    if report["period"]:
        since, until = report["period"]
        print(f"查詢區間：{fmt_date(since)} ～ {fmt_date(until)}")
        print(
            f"比對基準：{fmt_date(report['old_version_date'])}（區間起點前的版本）"
            f" → {fmt_date(report['new_version_date'])}（區間終點當時的版本）"
        )
        dates = report.get("amendment_dates_in_range")
        if dates:
            print(f"這段期間實際發生的修正日期：{'、'.join(fmt_date(d) for d in dates)}")
        else:
            print("這段期間內沒有查到任何修正紀錄。")
    else:
        print(f"比對對象：上一版本（{fmt_date(report['old_version_date'])}）")

    print(f"共有 {len(report['changed_rows'])} 條條文異動：\n")
    for row in report["changed_rows"]:
        old_plain = "".join(t for t, _ in row["old_runs"])
        new_plain = "".join(t for t, _ in row["new_runs"])
        print(f"── {row['article_no']}（{row['status']}）")
        if old_plain:
            print(f"   修正前：{old_plain}")
        if new_plain:
            print(f"   修正後：{new_plain}")
        print()


def safe_sheet_name(name: str, used: set) -> str:
    """Excel 分頁名稱不能超過 31 字、不能有 \\/?*[]: 這些符號，
    同名時自動加編號區隔（例如同一份清單裡兩部法規剛好前 31 字重複）。"""
    cleaned = re.sub(r'[\\/?*\[\]:]', "", name).strip() or "鑑別草稿"
    base = cleaned[:31]
    candidate = base
    n = 2
    while candidate in used:
        suffix = f"({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def write_law_sheet(wb, report: dict, used_names: set):
    """把單一法規的比對結果，寫成 workbook 裡的一個分頁。"""
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.styles import Font, Alignment

    ws = wb.create_sheet(safe_sheet_name(report["law_name"], used_names))

    headers = [
        "序號",
        "法規名稱",
        "法規條文內容(修正前)",
        "法規條文內容(修正後)",
        "異動類型",
        "鑑別日期",
        "條文適用性\n(適用/參考/不適用)",
        "守規性之評估",
        "鑑別人員",
        "備註",
    ]

    header_row_idx = 1
    if report.get("period"):
        since, until = report["period"]
        info_bits = [f"查詢區間：{fmt_date(since)} ～ {fmt_date(until)}"]
        info_bits.append(
            f"比對版本：{fmt_date(report['old_version_date'])} → {fmt_date(report['new_version_date'])}"
        )
        dates = report.get("amendment_dates_in_range")
        if dates:
            info_bits.append("實際修正日期：" + "、".join(fmt_date(d) for d in dates))
        ws.append(["　｜　".join(info_bits)])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        ws.cell(row=1, column=1).font = Font(italic=True, size=10, color="808080")
        header_row_idx = 2

    ws.append(headers)
    for cell in ws[header_row_idx]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    bold_underline = InlineFont(b=True, u="single")
    article_label_font = InlineFont(b=True)

    def to_rich_text(article_no, runs):
        # 開頭一律先標明「第幾條」，再接條文內容，這樣只看 C／D 欄位本身
        # 也能知道這段是哪一條，不用另外對照序號欄。
        blocks = [TextBlock(article_label_font, f"{article_no}\n")]
        if not runs:
            blocks.append("（本次修正未變動此條）")
            return CellRichText(*blocks)
        for text, changed in runs:
            if not text:
                continue
            if changed:
                blocks.append(TextBlock(bold_underline, text))
            else:
                blocks.append(text)
        return CellRichText(*blocks)

    today = dt.date.today().isoformat()
    for i, row in enumerate(report["changed_rows"], start=1):
        ws.append(
            [
                i,
                report["law_name"],
                to_rich_text(row["article_no"], row["old_runs"]),
                to_rich_text(row["article_no"], row["new_runs"]),
                row["status"],
                today,
                "",
                "",
                "",
                "",
            ]
        )

    widths = [6, 20, 42, 42, 8, 12, 14, 20, 10, 16]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=header_row_idx, column=idx).column_letter].width = w
    for r in ws.iter_rows(min_row=header_row_idx + 1):
        for c in r:
            c.alignment = Alignment(wrap_text=True, vertical="top")


def export_excel_multi(reports: list, out_path: str):
    """把好幾部法規的比對結果，各佔一個分頁，存成同一份 Excel。"""
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)  # 先移除預設的空白分頁

    used_names = set()
    written = 0
    for report in reports:
        if not report["changed_rows"]:
            continue
        write_law_sheet(wb, report, used_names)
        written += 1

    if written == 0:
        print("這次檢查的法規都沒有異動條文，不需要輸出 Excel。")
        return

    wb.save(out_path)
    print(f"\n已輸出鑑別草稿：{out_path}（共 {written} 部法規有異動）")


def load_pcodes_from_file(path: str) -> list:
    """從文字檔讀取法規代碼清單，一行一個；'#' 開頭當作註解、空白行略過。
    方便你把公司要追蹤的所有法規代碼存成一份清單，之後每次都用同一份檔案。"""
    pcodes = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.split("#", 1)[0].strip()
            if line:
                pcodes.append(line)
    return pcodes


def date_arg(s: str) -> str:
    """把 2025-06-01 / 2025/06/01 / 20250601 都轉成統一的 '20250601'。"""
    digits = re.sub(r"\D", "", s)
    if len(digits) != 8:
        raise argparse.ArgumentTypeError(
            f"日期格式看不懂：{s!r}（請用 2025-06-01、2025/06/01 或 20250601 這種格式）"
        )
    return digits


def main():
    parser = argparse.ArgumentParser(description="全國法規資料庫新舊條文一鍵比對原型")
    parser.add_argument(
        "pcodes",
        nargs="*",
        help="一個或多個法規代碼，例如 N0060010 N0060001（可以一次列很多個）",
    )
    parser.add_argument(
        "--list",
        dest="list_file",
        help="法規代碼清單檔（純文字，一行一個代碼），跟 pcodes 可以合併使用",
    )
    parser.add_argument(
        "--since",
        type=date_arg,
        help="只看這個日期之後的修正（例如 2025-06-01）。不指定則跟預設一樣，"
        "只比對『現行版本』跟『上一版』。",
    )
    parser.add_argument(
        "--until",
        type=date_arg,
        help="區間的結束日期（例如 2026-09-08），要跟 --since 一起用；不指定則預設為今天。",
    )
    parser.add_argument("--out", help="輸出 Excel 檔路徑，例如 draft.xlsx（不指定則只印在畫面上）")
    args = parser.parse_args()

    if args.until and not args.since:
        parser.error("--until 要跟 --since 一起用（只給結束日期，不知道要從哪天開始比對）。")

    pcodes = list(args.pcodes)
    if args.list_file:
        pcodes.extend(load_pcodes_from_file(args.list_file))
    # 去重複，但保留原本的順序
    pcodes = list(dict.fromkeys(pcodes))

    if not pcodes:
        parser.error("請至少給一個法規代碼，或用 --list 指定清單檔。")

    reports = []
    for pcode in pcodes:
        try:
            report = build_report(pcode, since=args.since, until=args.until)
        except requests.RequestException as e:
            print(f"[{pcode}] 連線失敗：{e}", file=sys.stderr)
            continue
        except RuntimeError as e:
            print(f"[{pcode}] {e}", file=sys.stderr)
            continue
        print_console_report(report)
        print()
        reports.append(report)

    if args.out:
        export_excel_multi(reports, args.out)


if __name__ == "__main__":
    try:
        main()
    except requests.RequestException as e:
        print(f"連線失敗：{e}", file=sys.stderr)
        sys.exit(1)
