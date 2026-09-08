# -*- coding: utf-8 -*-
"""
law_tool_app.py — 法規鑑別追蹤工具

本機用法（雙擊 start_law_tool.bat，或手動執行）：
    pip install -r requirements.txt
    python law_tool_app.py

執行後會自動開啟瀏覽器，網址是 http://127.0.0.1:5050 。
資料存在同一個資料夾裡的 law_tool.db 檔案，關掉程式、重開電腦都不會不見。
這支程式必須在「有一般網路連線」的電腦上執行（你自己的電腦即可），不能在
雲端沙盒環境裡執行。

也可以部署到雲端主機（例如 Railway），變成一個固定網址、任何電腦開瀏覽器
就能用，不用再依賴某一台電腦。部署時用 gunicorn 啟動（見 Procfile），並且
一定要接上「永久儲存空間（Volume）」、把 LAW_TOOL_DB_PATH 指向那個路徑，
否則服務重啟時資料庫會被清空。詳細步驟見 README.md。

這是 moj_law_diff.py 的延伸：把「抓最新條文、比對舊版本、跑文字 diff」這套
已經驗證過的邏輯，包成一個可以長期使用的本機網頁系統，取代原本要一條一條
法規手動下指令的方式，並且把每一條的鑑別結果（適用性、守規性評估、鑑別人員、
是否結案）存起來，下次可以繼續編輯，也可以隨時匯出成 Excel。

功能：
    - 新增要追蹤的法規（給法規代碼 pcode，並設定「追蹤起始日」）
    - 「立即比對」：抓最新條文，跟上次確認到的日期之後比對，只列出新的異動
    - 在網頁上直接填寫條文適用性／守規性評估／鑑別人員／備註／是否結案
    - 匯出單一法規或全部法規的 Excel，欄位對齊你原本的法規鑑別表格式，
      有異動的文字自動用粗體＋底線標示
"""

import json
import os
import re
import sqlite3
import sys
import threading
import time
import warnings
import webbrowser
from datetime import date, datetime, timedelta

import requests
from bs4 import BeautifulSoup
from flask import Flask, flash, redirect, render_template_string, request, send_file, url_for
from markupsafe import Markup, escape

# ---------------------------------------------------------------------------
# 全國法規資料庫 抓取／比對引擎（跟 moj_law_diff.py 同一套邏輯）
# ---------------------------------------------------------------------------

BASE = "https://law.moj.gov.tw/LawClass"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; RegComplianceTool/0.2; internal use)"
}
ARTICLE_NO_RE = re.compile(r"第\s*[\d\-]+\s*條")
_WARNED_INSECURE = False


def http_get(url: str):
    """對全國法規資料庫發出 GET 請求；遇到某些 Windows 環境常見的憑證驗證
    相容性錯誤時，自動退回不驗證憑證的方式重試一次。"""
    global _WARNED_INSECURE
    try:
        resp = requests.get(url, headers=HEADERS, timeout=20)
    except requests.exceptions.SSLError:
        if not _WARNED_INSECURE:
            print(
                "提醒：本機的憑證驗證對 law.moj.gov.tw 回報相容性錯誤，"
                "已自動改用不驗證憑證的方式重新連線（僅讀取公開政府法規網站，風險低）。",
                file=sys.stderr,
            )
            _WARNED_INSECURE = True
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            resp = requests.get(url, headers=HEADERS, timeout=20, verify=False)
    resp.raise_for_status()
    return resp


def normalize_article_no(raw: str) -> str:
    m = ARTICLE_NO_RE.search(raw)
    text = m.group(0) if m else raw
    return re.sub(r"\s+", "", text)


def fetch_articles(url: str) -> dict:
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
            continue
        no = normalize_article_no(raw_no)
        text = data_el.get_text("\n", strip=True)
        text = re.sub(r"\n{2,}", "\n", text)
        articles[no] = text
    return articles


def roc_to_yyyymmdd(s: str):
    if not s:
        return None
    m = re.search(r"(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日", s)
    if not m:
        return None
    roc_year, month, day = (int(x) for x in m.groups())
    return f"{roc_year + 1911:04d}{month:02d}{day:02d}"


def fmt_date(d: str) -> str:
    if not d or len(d) != 8:
        return d or ""
    return f"{d[0:4]}-{d[4:6]}-{d[6:8]}"


def _shift_day(d: str, delta: int) -> str:
    dt_obj = date(int(d[0:4]), int(d[4:6]), int(d[6:8])) + timedelta(days=delta)
    return dt_obj.strftime("%Y%m%d")


def previous_day(d: str) -> str:
    return _shift_day(d, -1)


def next_day(d: str) -> str:
    return _shift_day(d, 1)


def list_versions(pcode: str) -> dict:
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
            law_name = re.sub(r"\s*EN\s*$", "", law_name).strip()

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
            continue
        seen_dates.add(d)
        m_lser = re.search(r"lser=(\d+)", href)
        versions.append(
            {"date": d, "source": "old", "lnndate": d, "lser": m_lser.group(1) if m_lser else "001"}
        )

    versions.sort(key=lambda v: v["date"], reverse=True)
    return {"law_name": law_name, "amend_date_display": amend_date_display, "versions": versions}


def search_laws_by_keyword(keyword: str, max_pages: int = 2) -> tuple:
    """用法規名稱關鍵字查詢全國法規資料庫（中央法規查詢的「法規名稱」比對，
    不是條文內容比對），回傳可能對應的法規清單。每筆包含 pcode、法規名稱、
    現行修正日期、是否已廢止。第二個回傳值代表結果是否被截斷（代表還有更多筆，
    建議使用者輸入更精確的關鍵字）。"""
    from urllib.parse import quote

    results = []
    truncated = False
    for page in range(1, max_pages + 1):
        url = (
            "https://law.moj.gov.tw/Law/LawSearchResult.aspx"
            f"?cur=Ln&ty=LAW&kw={quote(keyword)}&mo=1&page={page}"
        )
        resp = http_get(url)
        resp.encoding = resp.apparent_encoding or "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")
        panel = soup.select_one("#pnLaw")
        if not panel:
            break
        rows = panel.select("tbody tr")
        if not rows:
            break
        for tr in rows:
            a, pcode = None, None
            for cand in tr.find_all("a"):
                m = re.search(r"pcode=([^&]+)", cand.get("href", ""), re.IGNORECASE)
                if m:
                    a, pcode = cand, m.group(1)
                    break
            if not pcode:
                continue
            row_text = tr.get_text(" ", strip=True)
            date_m = re.search(r"民國\s*\d+\s*年\s*\d+\s*月\s*\d+\s*日", row_text)
            results.append(
                {
                    "pcode": pcode,
                    "law_name": a.get("title") or a.get_text(strip=True),
                    "amend_date_display": date_m.group(0) if date_m else "",
                    "repealed": tr.select_one(".label-fei") is not None,
                }
            )
        if len(rows) < 20:
            break
        if page == max_pages:
            truncated = True
    return results, truncated


def version_asof(versions_desc: list, target_date: str, strict_before: bool = False):
    for v in versions_desc:
        if (v["date"] < target_date) if strict_before else (v["date"] <= target_date):
            return v
    return None


def fetch_version_articles(pcode: str, version: dict) -> dict:
    if version["source"] == "current":
        return fetch_articles(f"{BASE}/LawAll.aspx?pcode={pcode}")
    return fetch_articles(
        f"{BASE}/LawOldVer.aspx?pcode={pcode}&lnndate={version['lnndate']}&lser={version['lser']}"
    )


def inline_diff(old_text: str, new_text: str):
    import difflib

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


def build_report(pcode: str, since: str = None, until: str = None):
    """比對某一部法規的新舊條文；given since，抓『since 之前最後生效的版本』
    對『until（預設今天）當時生效的版本』；沒給 since 就跟現行版本對上一版比。"""
    info = list_versions(pcode)
    law_name = info["law_name"]
    if not law_name:
        raise RuntimeError(f"抓不到 pcode={pcode} 的法規名稱，請確認代碼是否正確。")

    versions = info["versions"]
    amend_date = info["amend_date_display"]

    if since is None:
        if len(versions) < 2:
            return {
                "pcode": pcode, "law_name": law_name, "amend_date": amend_date,
                "old_version_date": None, "new_version_date": None,
                "amendment_dates_in_range": None, "note": None, "changed_rows": [],
            }
        new_v, old_v, note = versions[0], versions[1], None
        until = until or versions[0]["date"]
    else:
        until = until or date.today().strftime("%Y%m%d")
        old_v = version_asof(versions, since, strict_before=True)
        new_v = version_asof(versions, until, strict_before=False)
        note = None
        if new_v is None:
            raise RuntimeError(f"{fmt_date(until)} 這個日期比這部法規最早的資料還早，查不到任何版本。")
        if old_v is None:
            old_v = versions[-1]
            note = (
                f"這個資料庫對此法規最早只能追溯到 {fmt_date(old_v['date'])}，"
                f"更早之前的版本無法比對，已改用這個最早版本當基準。"
            )

    if old_v["date"] == new_v["date"]:
        return {
            "pcode": pcode, "law_name": law_name, "amend_date": amend_date,
            "old_version_date": old_v["date"], "new_version_date": new_v["date"],
            "amendment_dates_in_range": [], "note": note, "changed_rows": [],
        }

    old = fetch_version_articles(pcode, old_v)
    current = fetch_version_articles(pcode, new_v)
    all_nos = list(dict.fromkeys(list(old.keys()) + list(current.keys())))

    rows = []
    for no in all_nos:
        old_text, new_text = old.get(no), current.get(no)
        if old_text == new_text:
            continue
        if old_text is None:
            status, old_runs, new_runs = "新增", [], [(new_text, True)]
        elif new_text is None:
            status, old_runs, new_runs = "刪除", [(old_text, True)], []
        else:
            status = "修正"
            old_runs, new_runs = inline_diff(old_text, new_text)
        rows.append({"article_no": no, "status": status, "old_runs": old_runs, "new_runs": new_runs})

    amendment_dates_in_range = [v["date"] for v in reversed(versions) if old_v["date"] < v["date"] <= new_v["date"]]

    return {
        "pcode": pcode, "law_name": law_name, "amend_date": amend_date,
        "old_version_date": old_v["date"], "new_version_date": new_v["date"],
        "amendment_dates_in_range": amendment_dates_in_range, "note": note, "changed_rows": rows,
    }


# ---------------------------------------------------------------------------
# 資料庫
# ---------------------------------------------------------------------------

APP_DIR = os.path.dirname(os.path.abspath(__file__))
# 本機執行時，資料庫預設放在程式旁邊；部署到雲端主機（例如 Railway）時，
# 用環境變數 LAW_TOOL_DB_PATH 指到「永久儲存空間（Volume）」掛載的路徑，
# 這樣服務重啟、重新部署時，已追蹤的法規和鑑別紀錄才不會被清空。
DB_PATH = os.environ.get("LAW_TOOL_DB_PATH") or os.path.join(APP_DIR, "law_tool.db")
os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS laws (
            pcode TEXT PRIMARY KEY,
            law_name TEXT,
            amend_date_display TEXT,
            note TEXT DEFAULT '',
            last_checked_until TEXT,
            created_at TEXT
        );
        CREATE TABLE IF NOT EXISTS findings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pcode TEXT NOT NULL,
            article_no TEXT NOT NULL,
            status TEXT,
            old_text TEXT,
            new_text TEXT,
            old_runs_json TEXT,
            new_runs_json TEXT,
            identified_date TEXT,
            applicability TEXT DEFAULT '',
            compliance_eval TEXT DEFAULT '',
            identifier TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            closed INTEGER DEFAULT 0,
            UNIQUE(pcode, article_no)
        );
        """
    )
    conn.commit()
    conn.close()


def upsert_findings(pcode: str, changed_rows: list):
    """把這次比對到的異動條文寫進資料庫。已經存在、內容沒有再變的條文不動
    （保留你之前填好的評估結果）；內容又變了的條文，才重設評估欄位、
    提醒你重新鑑別。"""
    conn = get_db()
    today = date.today().isoformat()
    for row in changed_rows:
        old_text = "".join(t for t, _ in row["old_runs"])
        new_text = "".join(t for t, _ in row["new_runs"])
        existing = conn.execute(
            "SELECT new_text FROM findings WHERE pcode=? AND article_no=?",
            (pcode, row["article_no"]),
        ).fetchone()
        if existing and existing["new_text"] == new_text:
            continue  # 這條的最新內容跟資料庫裡記錄的一樣，不用重新鑑別
        if existing:
            conn.execute(
                """UPDATE findings SET status=?, old_text=?, new_text=?, old_runs_json=?,
                   new_runs_json=?, identified_date=?, applicability='', compliance_eval='',
                   closed=0 WHERE pcode=? AND article_no=?""",
                (
                    row["status"], old_text, new_text,
                    json.dumps(row["old_runs"], ensure_ascii=False),
                    json.dumps(row["new_runs"], ensure_ascii=False),
                    today, pcode, row["article_no"],
                ),
            )
        else:
            conn.execute(
                """INSERT INTO findings
                   (pcode, article_no, status, old_text, new_text, old_runs_json, new_runs_json,
                    identified_date)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (
                    pcode, row["article_no"], row["status"], old_text, new_text,
                    json.dumps(row["old_runs"], ensure_ascii=False),
                    json.dumps(row["new_runs"], ensure_ascii=False),
                    today,
                ),
            )
    conn.commit()
    conn.close()


def runs_to_html(runs_json: str) -> Markup:
    runs = json.loads(runs_json) if runs_json else []
    if not runs:
        return Markup('<span class="muted">（本次修正未變動此條）</span>')
    parts = []
    for text, changed in runs:
        t = str(escape(text)).replace("\n", "<br>")
        parts.append(f"<b class='chg'>{t}</b>" if changed else t)
    return Markup("".join(parts))


# ---------------------------------------------------------------------------
# Flask App
# ---------------------------------------------------------------------------

app = Flask(__name__)
# secret_key 只用來簽署 flash 訊息用的 session cookie，這個工具沒有登入機制、
# 不存放帳密，風險很低；部署到雲端主機時可用環境變數 SECRET_KEY 覆寫。
app.secret_key = os.environ.get("SECRET_KEY", "law-tool-local-only")

BASE_CSS = """
:root{
  --ink:#202a24; --paper:#f3f2ec; --surface:#ffffff; --line:#dcd9cd;
  --muted:#6e6a5c; --accent:#a8341f; --accent-2:#3f6b4a;
}
*{box-sizing:border-box;}
body{margin:0;background:var(--paper);color:var(--ink);
  font-family:"Noto Sans TC","Microsoft JhengHei",system-ui,sans-serif;line-height:1.6;}
header{background:var(--surface);border-bottom:1px solid var(--line);padding:16px 28px;
  display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;}
header h1{font-size:18px;margin:0;}
header h1 a{color:var(--ink);text-decoration:none;}
main{max-width:1100px;margin:0 auto;padding:24px 20px 80px;}
.flash{background:#fdf1ee;border:1px solid var(--accent);color:var(--accent);
  padding:10px 14px;border-radius:6px;margin-bottom:16px;font-size:14px;}
table{border-collapse:collapse;width:100%;background:var(--surface);
  border:1px solid var(--line);border-radius:8px;overflow:hidden;font-size:13.5px;}
th,td{padding:10px 12px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top;}
th{background:#faf9f4;font-weight:600;font-size:12.5px;color:var(--muted);}
tr:last-child td{border-bottom:none;}
.btn{display:inline-block;padding:7px 14px;border-radius:6px;border:1px solid var(--line);
  background:var(--surface);color:var(--ink);cursor:pointer;font-size:13.5px;text-decoration:none;}
.btn:hover{border-color:var(--ink);}
.btn-primary{background:var(--ink);color:var(--surface);border-color:var(--ink);}
.btn-danger{color:var(--accent);border-color:var(--accent);}
.card{background:var(--surface);border:1px solid var(--line);border-radius:8px;
  padding:20px;margin-bottom:20px;}
input[type=text],input[type=date],select{
  padding:6px 8px;border:1px solid var(--line);border-radius:5px;font-size:13.5px;
  font-family:inherit;width:100%;}
label{font-size:12.5px;color:var(--muted);display:block;margin-bottom:4px;}
.row{display:flex;gap:14px;flex-wrap:wrap;align-items:end;}
.row > div{flex:1;min-width:140px;}
.muted{color:var(--muted);}
.chg{text-decoration:underline;text-decoration-thickness:2px;}
.pill{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11.5px;font-weight:600;}
.pill-open{background:#fdf1ee;color:var(--accent);}
.pill-closed{background:#eef4ef;color:var(--accent-2);}
.diff-cell{max-width:280px;}
.top-actions{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap;}

/* 手機／窄螢幕 RWD：表格改成可以左右滑動，避免欄位被硬擠爆版；
   按鈕、輸入框加大，方便手指點按；整體邊距縮小，多留一點內容空間。 */
.table-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch;border-radius:8px;}
.table-scroll table{border-radius:0;}
@media (max-width:700px){
  main{padding:16px 14px 60px;}
  header{padding:12px 16px;}
  header h1{font-size:16px;}
  .card{padding:16px;}
  .table-scroll table{min-width:640px;}
  .table-scroll table.table-wide{min-width:980px;}
  .btn{padding:9px 14px;}
  input[type=text],input[type=date],select{font-size:16px;padding:8px;}
  .row{gap:10px;}
}
"""

LAYOUT = """
<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{ title or "法規鑑別追蹤工具" }}</title>
<style>{{ css|safe }}</style>
</head>
<body>
<header>
  <h1><a href="{{ url_for('dashboard') }}">法規鑑別追蹤工具</a></h1>
  <div class="muted" style="font-size:12.5px;">資料來源：全國法規資料庫（law.moj.gov.tw）</div>
</header>
<main>
  {% with messages = get_flashed_messages() %}
    {% for m in messages %}<div class="flash">{{ m }}</div>{% endfor %}
  {% endwith %}
  {{ body|safe }}
</main>
</body>
</html>
"""


def render(body_html, **ctx):
    body = render_template_string(body_html, **ctx)
    return render_template_string(LAYOUT, css=BASE_CSS, body=body, **ctx)


DASHBOARD_BODY = """
<div class="top-actions">
  <form method="post" action="{{ url_for('check_all') }}">
    <button class="btn btn-primary" type="submit">全部立即比對</button>
  </form>
  <a class="btn" href="{{ url_for('export_all') }}">匯出全部法規 Excel</a>
</div>

<div class="card">
  <h3 style="margin-top:0;">新增要追蹤的法規</h3>
  <form method="get" action="{{ url_for('search_laws') }}" style="margin-bottom:14px;">
    <div class="row">
      <div>
        <label>用法規名稱關鍵字搜尋（推薦，不用打完整全名，例如「教育訓練規則」）</label>
        <input type="text" name="kw" placeholder="輸入法規名稱關鍵字">
      </div>
      <div style="flex:0;">
        <button class="btn btn-primary" type="submit">搜尋法規</button>
      </div>
    </div>
  </form>
  <details>
    <summary class="muted" style="cursor:pointer;font-size:12.5px;">或者，直接輸入法規代碼（pcode）新增</summary>
    <form method="post" action="{{ url_for('add_law') }}" style="margin-top:12px;">
      <div class="row">
        <div>
          <label>法規代碼（pcode，可從 law.moj.gov.tw 該法規網址的 pcode= 後面找到）</label>
          <input type="text" name="pcode" placeholder="例如 N0060010" required>
        </div>
        <div>
          <label>追蹤起始日（只比對這天之後發生的修正）</label>
          <input type="date" name="start_date" value="{{ today }}" required>
        </div>
        <div style="flex:0;">
          <button class="btn btn-primary" type="submit">新增</button>
        </div>
      </div>
    </form>
  </details>
</div>

<div class="table-scroll">
<table>
  <thead><tr>
    <th>法規名稱</th><th>pcode</th><th>現行修正日期</th><th>已追蹤到</th>
    <th>待鑑別</th><th>操作</th>
  </tr></thead>
  <tbody>
    {% for law in laws %}
    <tr>
      <td><a href="{{ url_for('law_detail', pcode=law.pcode) }}">{{ law.law_name or law.pcode }}</a>
        {% if law.note %}<div class="muted" style="font-size:12px;">{{ law.note }}</div>{% endif %}
      </td>
      <td class="muted">{{ law.pcode }}</td>
      <td>{{ law.amend_date_display or "-" }}</td>
      <td>{{ law.last_checked_display }}</td>
      <td>
        {% if law.open_count %}
        <span class="pill pill-open">{{ law.open_count }} 筆未結案</span>
        {% else %}
        <span class="pill pill-closed">全部結案</span>
        {% endif %}
      </td>
      <td style="white-space:nowrap;">
        <form style="display:inline" method="post" action="{{ url_for('check_law', pcode=law.pcode) }}">
          <button class="btn" type="submit">立即比對</button>
        </form>
        <a class="btn" href="{{ url_for('law_detail', pcode=law.pcode) }}">查看</a>
        <form style="display:inline" method="post" action="{{ url_for('delete_law', pcode=law.pcode) }}"
              onsubmit="return confirm('確定要刪除這部法規跟它累積的鑑別紀錄嗎？');">
          <button class="btn btn-danger" type="submit">刪除</button>
        </form>
      </td>
    </tr>
    {% else %}
    <tr><td colspan="6" class="muted">還沒有追蹤任何法規，先在上面新增一部看看。</td></tr>
    {% endfor %}
  </tbody>
</table>
</div>
"""


@app.route("/")
def dashboard():
    conn = get_db()
    laws = conn.execute("SELECT * FROM laws ORDER BY created_at").fetchall()
    rows = []
    for law in laws:
        open_count = conn.execute(
            "SELECT COUNT(*) c FROM findings WHERE pcode=? AND closed=0", (law["pcode"],)
        ).fetchone()["c"]
        d = dict(law)
        d["open_count"] = open_count
        d["last_checked_display"] = fmt_date(law["last_checked_until"]) if law["last_checked_until"] else "尚未比對"
        rows.append(d)
    conn.close()
    return render(DASHBOARD_BODY, laws=rows, today=date.today().isoformat())


@app.route("/laws", methods=["POST"])
def add_law():
    pcode = request.form.get("pcode", "").strip()
    start_date = request.form.get("start_date", "").strip().replace("-", "")
    if not pcode:
        flash("請輸入法規代碼。")
        return redirect(url_for("dashboard"))
    try:
        info = list_versions(pcode)
    except requests.RequestException as e:
        flash(f"連線失敗：{e}")
        return redirect(url_for("dashboard"))
    if not info["law_name"]:
        flash(f"抓不到 pcode={pcode} 的法規名稱，請確認代碼是否正確。")
        return redirect(url_for("dashboard"))

    # last_checked_until 的定義統一是「已經確認到這一天為止，之後的檢查只看
    # 更新的部分」；所以「追蹤起始日」要換算成「起始日前一天」存起來，這樣第一次
    # 「立即比對」查詢時（since = next_day(last_checked_until)）才會剛好從
    # 使用者選的那一天開始算，不會漏掉剛好在起始日當天生效的修正。
    baseline = previous_day(start_date) if start_date else None

    conn = get_db()
    conn.execute(
        "INSERT OR REPLACE INTO laws (pcode, law_name, amend_date_display, note, last_checked_until, created_at) "
        "VALUES (?, ?, ?, COALESCE((SELECT note FROM laws WHERE pcode=?), ''), ?, "
        "COALESCE((SELECT created_at FROM laws WHERE pcode=?), ?))",
        (pcode, info["law_name"], info["amend_date_display"], pcode, baseline, pcode, datetime.now().isoformat()),
    )
    conn.commit()
    conn.close()
    flash(f"已新增「{info['law_name']}」，追蹤起始日設為 {fmt_date(start_date)}。")
    return redirect(url_for("dashboard"))


SEARCH_BODY = """
<p><a href="{{ url_for('dashboard') }}">&larr; 回法規清單</a></p>
<div class="card">
  <h3 style="margin-top:0;">用法規名稱關鍵字搜尋</h3>
  <form method="get" action="{{ url_for('search_laws') }}">
    <div class="row">
      <div>
        <label>法規名稱關鍵字（不用打完整全名，例如「教育訓練規則」）</label>
        <input type="text" name="kw" value="{{ keyword }}" placeholder="例如：教育訓練規則" autofocus>
      </div>
      <div style="flex:0;">
        <button class="btn btn-primary" type="submit">搜尋</button>
      </div>
    </div>
  </form>
</div>

{% if error %}
<div class="flash">{{ error }}</div>
{% elif keyword %}
  {% if results %}
  <div class="table-scroll">
  <table>
    <thead><tr><th>法規名稱</th><th>pcode</th><th>現行修正日期</th><th style="width:220px;">操作</th></tr></thead>
    <tbody>
      {% for r in results %}
      <tr>
        <td>{{ r.law_name }}{% if r.repealed %} <span class="pill pill-open">已廢止</span>{% endif %}</td>
        <td class="muted">{{ r.pcode }}</td>
        <td>{{ r.amend_date_display or "-" }}</td>
        <td>
          {% if r.pcode in tracked %}
          <a class="btn" href="{{ url_for('law_detail', pcode=r.pcode) }}">已在追蹤清單，查看</a>
          {% else %}
          <form method="post" action="{{ url_for('add_law') }}" class="row" style="gap:6px;flex-wrap:nowrap;">
            <input type="hidden" name="pcode" value="{{ r.pcode }}">
            <div style="min-width:130px;">
              <input type="date" name="start_date" value="{{ today }}" required>
            </div>
            <div style="flex:0;">
              <button class="btn btn-primary" type="submit">開始追蹤</button>
            </div>
          </form>
          {% endif %}
        </td>
      </tr>
      {% endfor %}
    </tbody>
  </table>
  </div>
  {% if truncated %}
  <p class="muted" style="margin-top:10px;">符合的結果較多，這裡只列出前面幾筆，如果沒看到你要的法規，
    請輸入更精確的關鍵字（例如加上「辦法」「規則」「標準」等法規名稱裡的字）。</p>
  {% endif %}
  {% else %}
  <p class="muted">找不到名稱包含「{{ keyword }}」的法規，換個關鍵字試試看
    （例如只打法規名稱裡最特別的幾個字，不用打「XX法」「XX辦法」這種通用字）。</p>
  {% endif %}
{% endif %}
"""


@app.route("/search")
def search_laws():
    keyword = request.args.get("kw", "").strip()
    results, truncated, error = [], False, None
    if keyword:
        try:
            results, truncated = search_laws_by_keyword(keyword)
        except requests.RequestException as e:
            error = f"連線失敗：{e}"
    conn = get_db()
    tracked = {r["pcode"] for r in conn.execute("SELECT pcode FROM laws").fetchall()}
    conn.close()
    return render(
        SEARCH_BODY, keyword=keyword, results=results, truncated=truncated,
        error=error, tracked=tracked, today=date.today().isoformat(),
    )


@app.route("/laws/<pcode>/delete", methods=["POST"])
def delete_law(pcode):
    conn = get_db()
    conn.execute("DELETE FROM findings WHERE pcode=?", (pcode,))
    conn.execute("DELETE FROM laws WHERE pcode=?", (pcode,))
    conn.commit()
    conn.close()
    flash("已刪除。")
    return redirect(url_for("dashboard"))


def _run_check(pcode: str):
    conn = get_db()
    law = conn.execute("SELECT * FROM laws WHERE pcode=?", (pcode,)).fetchone()
    conn.close()
    if not law:
        return f"[{pcode}] 找不到這部法規。"
    # last_checked_until 存的是「已經確認到這一天為止」，所以查詢時要從
    # 它的隔天開始算，才不會把已經看過、已經填好評估的那次異動又抓回來一次。
    since = next_day(law["last_checked_until"]) if law["last_checked_until"] else None
    try:
        report = build_report(pcode, since=since)
    except (requests.RequestException, RuntimeError) as e:
        return f"[{report_name(law)}] 比對失敗：{e}"

    upsert_findings(pcode, report["changed_rows"])

    new_baseline = report["new_version_date"] or law["last_checked_until"] or date.today().strftime("%Y%m%d")
    conn = get_db()
    conn.execute(
        "UPDATE laws SET last_checked_until=?, law_name=?, amend_date_display=? WHERE pcode=?",
        (new_baseline, report["law_name"], report["amend_date"], pcode),
    )
    conn.commit()
    conn.close()

    n = len(report["changed_rows"])
    if n == 0:
        return f"「{report['law_name']}」：這段期間沒有新的修正。"
    return f"「{report['law_name']}」：新增／更新了 {n} 條異動，已存入待鑑別清單。"


def report_name(law_row):
    return law_row["law_name"] or law_row["pcode"]


@app.route("/laws/<pcode>/check", methods=["POST"])
def check_law(pcode):
    msg = _run_check(pcode)
    flash(msg)
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/check_all", methods=["POST"])
def check_all():
    conn = get_db()
    pcodes = [r["pcode"] for r in conn.execute("SELECT pcode FROM laws").fetchall()]
    conn.close()
    msgs = [_run_check(p) for p in pcodes]
    flash(" ／ ".join(msgs) if msgs else "還沒有追蹤任何法規。")
    return redirect(url_for("dashboard"))


LAW_DETAIL_BODY = """
<p><a href="{{ url_for('dashboard') }}">&larr; 回法規清單</a></p>
<div class="card">
  <h2 style="margin:0 0 4px;">{{ law.law_name }}</h2>
  <div class="muted">pcode={{ law.pcode }}　｜　已追蹤到：{{ last_checked_display }}</div>
</div>

<div class="top-actions">
  <form method="post" action="{{ url_for('check_law', pcode=law.pcode) }}">
    <button class="btn btn-primary" type="submit">立即比對</button>
  </form>
  <a class="btn" href="{{ url_for('export_one', pcode=law.pcode) }}">匯出此法規 Excel</a>
</div>

<form method="post" action="{{ url_for('save_findings', pcode=law.pcode) }}">
<div class="table-scroll">
<table class="table-wide">
  <thead><tr>
    <th style="width:5%">條號</th>
    <th style="width:8%">異動</th>
    <th style="width:20%">修正前</th>
    <th style="width:20%">修正後</th>
    <th style="width:10%">鑑別日期</th>
    <th style="width:9%">條文適用性</th>
    <th style="width:14%">守規性之評估</th>
    <th style="width:8%">鑑別人員</th>
    <th style="width:10%">備註</th>
    <th style="width:6%">結案</th>
  </tr></thead>
  <tbody>
    {% for f in findings %}
    <tr>
      <td>{{ f.article_no }}</td>
      <td>{{ f.status }}</td>
      <td class="diff-cell">{{ f.old_html }}</td>
      <td class="diff-cell">{{ f.new_html }}</td>
      <td class="muted">{{ f.identified_date }}</td>
      <td>
        <select name="applicability_{{ f.id }}">
          {% for opt in ["", "適用", "參考", "不適用"] %}
          <option value="{{ opt }}" {{ "selected" if f.applicability==opt else "" }}>{{ opt or "（未選）" }}</option>
          {% endfor %}
        </select>
      </td>
      <td><input type="text" name="compliance_eval_{{ f.id }}" value="{{ f.compliance_eval }}"></td>
      <td><input type="text" name="identifier_{{ f.id }}" value="{{ f.identifier }}"></td>
      <td><input type="text" name="remark_{{ f.id }}" value="{{ f.remark }}"></td>
      <td style="text-align:center;">
        <input type="checkbox" name="closed_{{ f.id }}" {{ "checked" if f.closed else "" }}>
      </td>
    </tr>
    {% else %}
    <tr><td colspan="10" class="muted">目前沒有待鑑別的條文，按上面「立即比對」試試看。</td></tr>
    {% endfor %}
  </tbody>
</table>
</div>
{% if findings %}
<p style="margin-top:14px;"><button class="btn btn-primary" type="submit">儲存這一頁的填寫內容</button></p>
{% endif %}
</form>
"""


@app.route("/laws/<pcode>")
def law_detail(pcode):
    conn = get_db()
    law = conn.execute("SELECT * FROM laws WHERE pcode=?", (pcode,)).fetchone()
    if not law:
        conn.close()
        flash("找不到這部法規。")
        return redirect(url_for("dashboard"))
    findings = conn.execute(
        "SELECT * FROM findings WHERE pcode=? ORDER BY closed, identified_date DESC, id", (pcode,)
    ).fetchall()
    conn.close()

    rows = []
    for f in findings:
        d = dict(f)
        d["old_html"] = runs_to_html(f["old_runs_json"])
        d["new_html"] = runs_to_html(f["new_runs_json"])
        rows.append(d)

    last_checked_display = fmt_date(law["last_checked_until"]) if law["last_checked_until"] else "尚未比對"
    return render(LAW_DETAIL_BODY, law=law, findings=rows, last_checked_display=last_checked_display)


@app.route("/laws/<pcode>/save", methods=["POST"])
def save_findings(pcode):
    conn = get_db()
    ids = [r["id"] for r in conn.execute("SELECT id FROM findings WHERE pcode=?", (pcode,)).fetchall()]
    for fid in ids:
        applicability = request.form.get(f"applicability_{fid}", "")
        compliance_eval = request.form.get(f"compliance_eval_{fid}", "")
        identifier = request.form.get(f"identifier_{fid}", "")
        remark = request.form.get(f"remark_{fid}", "")
        closed = 1 if request.form.get(f"closed_{fid}") else 0
        conn.execute(
            """UPDATE findings SET applicability=?, compliance_eval=?, identifier=?, remark=?, closed=?
               WHERE id=?""",
            (applicability, compliance_eval, identifier, remark, closed, fid),
        )
    conn.commit()
    conn.close()
    flash("已儲存。")
    return redirect(url_for("law_detail", pcode=pcode))


# ---------------------------------------------------------------------------
# Excel 匯出
# ---------------------------------------------------------------------------


def safe_sheet_name(name: str, used: set) -> str:
    cleaned = re.sub(r"[\\/?*\[\]:]", "", name or "").strip() or "鑑別草稿"
    base = cleaned[:31]
    candidate = base
    n = 2
    while candidate in used:
        suffix = f"({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def write_law_sheet_from_db(wb, pcode: str, used_names: set):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    from openpyxl.styles import Alignment, Font

    conn = get_db()
    law = conn.execute("SELECT * FROM laws WHERE pcode=?", (pcode,)).fetchone()
    findings = conn.execute(
        "SELECT * FROM findings WHERE pcode=? ORDER BY id", (pcode,)
    ).fetchall()
    conn.close()
    if not law or not findings:
        return False

    ws = wb.create_sheet(safe_sheet_name(law["law_name"], used_names))
    headers = [
        "序號", "法規名稱", "法規條文內容(修正前)", "法規條文內容(修正後)", "異動類型",
        "鑑別日期", "條文適用性\n(適用/參考/不適用)", "守規性之評估", "鑑別人員",
        "是否結案", "備註",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    bold_underline = InlineFont(b=True, u="single")
    article_label_font = InlineFont(b=True)

    def to_rich_text(article_no, runs_json):
        runs = json.loads(runs_json) if runs_json else []
        blocks = [TextBlock(article_label_font, f"{article_no}\n")]
        if not runs:
            blocks.append("（本次修正未變動此條）")
            return CellRichText(*blocks)
        for text, changed in runs:
            if not text:
                continue
            blocks.append(TextBlock(bold_underline, text) if changed else text)
        return CellRichText(*blocks)

    for i, f in enumerate(findings, start=1):
        ws.append(
            [
                i, law["law_name"],
                to_rich_text(f["article_no"], f["old_runs_json"]),
                to_rich_text(f["article_no"], f["new_runs_json"]),
                f["status"], f["identified_date"], f["applicability"], f["compliance_eval"],
                f["identifier"], "已結案" if f["closed"] else "未結案", f["remark"],
            ]
        )

    widths = [6, 20, 40, 40, 8, 12, 14, 20, 10, 8, 16]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = w
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return True


@app.route("/export/<pcode>")
def export_one(pcode):
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    ok = write_law_sheet_from_db(wb, pcode, set())
    if not ok:
        flash("這部法規目前沒有可匯出的鑑別紀錄。")
        return redirect(url_for("law_detail", pcode=pcode))
    out_path = os.path.join(APP_DIR, f"{pcode}_鑑別結果.xlsx")
    wb.save(out_path)
    return send_file(out_path, as_attachment=True)


@app.route("/export/all")
def export_all():
    from openpyxl import Workbook

    conn = get_db()
    pcodes = [r["pcode"] for r in conn.execute("SELECT pcode FROM laws").fetchall()]
    conn.close()

    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    written = 0
    for pcode in pcodes:
        if write_law_sheet_from_db(wb, pcode, used):
            written += 1

    if written == 0:
        flash("目前沒有任何法規有可匯出的鑑別紀錄。")
        return redirect(url_for("dashboard"))

    out_path = os.path.join(APP_DIR, "全部法規_鑑別結果.xlsx")
    wb.save(out_path)
    return send_file(out_path, as_attachment=True)


# ---------------------------------------------------------------------------
# 啟動
# ---------------------------------------------------------------------------

# 不管是本機用 python law_tool_app.py 直接執行，還是部署到雲端主機用
# gunicorn 啟動（gunicorn 只會 import 這支檔案、不會執行下面的
# if __name__ == "__main__" 區塊），都要確保資料表存在，所以放在最外層。
init_db()


def _open_browser(url):
    time.sleep(1.0)
    webbrowser.open(url)


if __name__ == "__main__":
    # 本機雙擊 start_law_tool.bat／直接執行時走這條路：只監聽本機
    # 127.0.0.1，並自動開瀏覽器，跟以前的行為完全一樣。
    # （部署到雲端主機時，是用 gunicorn 啟動，不會跑到這裡，改由
    # gunicorn 監聽 0.0.0.0 上的 $PORT。）
    port = int(os.environ.get("PORT", 5050))
    url = f"http://127.0.0.1:{port}"
    threading.Thread(target=_open_browser, args=(url,), daemon=True).start()
    print(f"法規鑑別追蹤工具已啟動：{url}（要停止的話回到這個視窗按 Ctrl+C）")
    app.run(host="127.0.0.1", port=port, debug=False)
