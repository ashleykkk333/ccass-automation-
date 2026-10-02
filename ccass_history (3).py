"""
CCASS 12-month history builder for any HK stock.

Pulls daily CCASS holdings from HKEXnews' CCASS Shareholding Search
(https://www3.hkexnews.hk/sdw/search/searchsdw.aspx), which only keeps
the past 12 months, then builds a workbook in the same layout as the
01633 file: CCASS Holdings, Buyer/Seller Movement, and B1-B10 / S1-S10.

Setup:   pip install requests beautifulsoup4 pandas openpyxl
Run:     edit STOCK_CODE below, then:  python ccass_history.py
         (or override without editing:  python ccass_history.py --stock 700)

Each day is cached in ./ccass_cache/<stock>/ so an interrupted run resumes
where it stopped. HKEX's terms allow personal, non-commercial use only; keep
the delay between requests.
"""
# =====================================================================
#  EDIT THESE, THEN RUN THE SCRIPT
# =====================================================================
STOCK_CODE = "1985"     # any HK stock code, e.g. "700", "1633", "9988"
DAYS       = 365        # how far back to go (HKEX keeps max ~365 days)
TOP_N      = 10         # how many top buyers / sellers to show
# =====================================================================

import argparse, json, os, re, time
from datetime import date, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

URL = "https://www3.hkexnews.hk/sdw/search/searchsdw.aspx"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
F = Font(name="Arial", size=10)
FB = Font(name="Arial", size=10, bold=True)
GREEN = CellIsRule(operator="greaterThan", formula=["0"],
                   fill=PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE"),
                   font=Font(name="Arial", size=10, color="006100"))
RED = CellIsRule(operator="lessThan", formula=["0"],
                 fill=PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE"),
                 font=Font(name="Arial", size=10, color="9C0006"))


def highlight(ws, rng):
    """Green fill for buys (>0), red fill for sells (<0)."""
    ws.conditional_formatting.add(rng, GREEN)
    ws.conditional_formatting.add(rng, RED)


# ---------------------------------------------------------------- scraping
def _num(s):
    s = re.sub(r"[^\d.\-]", "", s or "")
    return float(s) if s not in ("", "-", ".") else 0.0


def _hidden_fields(session):
    soup = BeautifulSoup(session.get(URL, headers=HEADERS, timeout=30).text, "html.parser")
    return {i["name"]: i.get("value", "") for i in soup.select("input[type=hidden]") if i.get("name")}


def fetch_day(session, hidden, stock, d):
    data = dict(hidden)
    data.update({
        "__EVENTTARGET": "btnSearch", "__EVENTARGUMENT": "",
        "today": date.today().strftime("%Y%m%d"),
        "sortBy": "shareholding", "sortDirection": "desc", "alertMsg": "",
        "txtShareholdingDate": d.strftime("%Y/%m/%d"),
        "txtStockCode": stock, "txtStockName": "",
        "txtParticipantID": "", "txtParticipantName": "", "txtSelPartID": "",
    })
    r = session.post(URL, data=data, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return parse(r.text)


def _cell(td):
    body = td.select_one(".mobile-list-body")
    return (body or td).get_text(" ", strip=True)


def parse(html):
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for tr in soup.select("#pnlResultNormal table tbody tr") or soup.select("table tbody tr"):
        tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        pid, name, holding = _cell(tds[0]), _cell(tds[1]), _cell(tds[3])
        if not re.match(r"^[A-Z]\d{5}$", pid):   # skips header/summary rows
            continue
        rows.append({"id": pid, "name": name, "holding": _num(holding), "pct": _num(_cell(tds[4]))})

    text = soup.get_text(" ", strip=True)
    issued = None
    # Exact figure printed after "(last updated figure)"
    m = re.search(r"\(last updated figure\)\s*:?\s*([\d,]{5,})", text, re.I)
    if m:
        issued = _num(m.group(1))
    if not issued:
        # Backup: issued = holding / stake% for the largest holder (approximate)
        big = max((r for r in rows if r["pct"] > 0), key=lambda r: r["holding"], default=None)
        if big:
            issued = round(big["holding"] / big["pct"] * 100)
    return {"rows": rows, "issued": issued}


def scrape(stock, days, delay, weekdays_only):
    cache = os.path.join("ccass_cache", stock)
    os.makedirs(cache, exist_ok=True)
    end = date.today() - timedelta(days=1)          # HKEX publishes T-1
    dates = [end - timedelta(days=i) for i in range(days)]
    if weekdays_only:
        dates = [d for d in dates if d.weekday() < 5]

    s = requests.Session()
    hidden = _hidden_fields(s)
    out = {}
    for i, d in enumerate(sorted(dates)):
        f = os.path.join(cache, d.strftime("%Y-%m-%d") + ".json")
        if os.path.exists(f):
            out[d] = json.load(open(f))
            continue
        res = None
        for attempt in range(3):
            try:
                res = fetch_day(s, hidden, stock, d)
                break
            except Exception as e:
                print(f"  retry {d} ({e})")
                time.sleep(5)
                hidden = _hidden_fields(s)
        if res is None:
            continue
        if res["rows"]:
            json.dump(res, open(f, "w"))
            out[d] = res
        print(f"[{i+1}/{len(dates)}] {d}  participants={len(res['rows'])}")
        time.sleep(delay)
    return out


# ---------------------------------------------------------------- analysis
def build_frames(data):
    recs, names = [], {}
    for d, res in data.items():
        for r in res["rows"]:
            recs.append((d, r["id"], r["holding"]))
            names[r["id"]] = r["name"]
    df = pd.DataFrame(recs, columns=["date", "id", "holding"])
    hold = df.pivot_table(index="date", columns="id", values="holding", aggfunc="sum")
    hold = hold.sort_index().fillna(0)          # absent from CCASS list = 0
    return hold, names


def write_block(ws, r, c, text, bold=False):
    cell = ws.cell(row=r, column=c, value=text)
    cell.font = FB if bold else F
    return cell


def build_workbook(hold, names, issued, top, out):
    change = hold.diff().fillna(0)
    stake = hold / issued * 100
    net = hold.iloc[-1] - hold.iloc[0]
    buyers = net[net > 0].sort_values(ascending=False).head(top).index.tolist()
    sellers = net[net < 0].sort_values().head(top).index.tolist()
    dates_desc = list(hold.index[::-1])
    first, last = hold.index[0], hold.index[-1]

    wb = Workbook()

    # 1. Net change over the whole period
    ws = wb.active
    ws.title = "CCASS Holding Changes"
    write_block(ws, 1, 1, f"Net CCASS change {first:%Y/%m/%d} to {last:%Y/%m/%d}  "
                          f"(issued shares used: {issued:,.0f})", True)
    hdr = ["CCASS ID", "Name", f"Holding {first:%Y/%m/%d}", f"Holding {last:%Y/%m/%d}",
           "Net change", "Stake% start", "Stake% end", "Stake% change"]
    for j, h in enumerate(hdr, 1):
        write_block(ws, 3, j, h, True)
    for i, pid in enumerate(net.sort_values(ascending=False).index, 4):
        vals = [pid, names.get(pid, ""), hold[pid].iloc[0], hold[pid].iloc[-1], net[pid],
                round(stake[pid].iloc[0], 2), round(stake[pid].iloc[-1], 2),
                round(stake[pid].iloc[-1] - stake[pid].iloc[0], 2)]
        for j, v in enumerate(vals, 1):
            c = write_block(ws, i, j, v)
            if j in (3, 4, 5):
                c.number_format = "#,##0;(#,##0);-"
    n_last = 3 + len(net)
    highlight(ws, f"E4:E{n_last}")
    highlight(ws, f"H4:H{n_last}")

    # 2. Latest snapshot
    ws = wb.create_sheet("CCASS Holdings")
    write_block(ws, 1, 1, f"CCASS Holdings on {last:%Y/%m/%d}", True)
    tot = hold.iloc[-1].sum()
    for i, (lab, v) in enumerate([("Total in CCASS", tot), ("Securities not in CCASS", issued - tot),
                                  ("Issued securities", issued)], 3):
        write_block(ws, i, 1, lab)
        write_block(ws, i, 2, v).number_format = "#,##0"
        write_block(ws, i, 3, round(v / issued * 100, 2))
    for j, h in enumerate(["Row", "CCASS ID", "Name", "Holding", "Lastchange", "Stake%", "Cumul.Stake%"], 1):
        write_block(ws, 8, j, h, True)
    latest = hold.iloc[-1][hold.iloc[-1] > 0].sort_values(ascending=False)
    cum = 0
    for i, (pid, h) in enumerate(latest.items(), 1):
        nz = change[pid][change[pid] != 0]
        lastchg = nz.index[-1] if len(nz) else first
        pct = h / issued * 100
        cum += pct
        for j, v in enumerate([i, pid, names.get(pid, ""), h, lastchg.strftime("%Y/%m/%d"),
                               round(pct, 2), round(cum, 2)], 1):
            c = write_block(ws, 8 + i, j, v)
            if j == 4:
                c.number_format = "#,##0"

    # 3. Buyer / Seller movement grids
    for title, ids in [("Buyer Movement", buyers), ("Seller Movement", sellers)]:
        ws = wb.create_sheet(title)
        write_block(ws, 2, 1, "Holdingdate", True)
        for k, pid in enumerate(ids):
            c0 = 2 + k * 3
            write_block(ws, 1, c0, names.get(pid, pid), True)
            write_block(ws, 2, c0, "Change", True)
            write_block(ws, 2, c0 + 1, "Stake%", True)
            if k < len(ids) - 1:
                write_block(ws, 1, c0 + 2, "|")
        for r, d in enumerate(dates_desc, 3):
            write_block(ws, r, 1, d.strftime("%Y/%m/%d"))
            for k, pid in enumerate(ids):
                c0 = 2 + k * 3
                write_block(ws, r, c0, change.at[d, pid]).number_format = "#,##0;-#,##0;0"
                write_block(ws, r, c0 + 1, round(stake.at[d, pid], 2))
        last_row = 2 + len(dates_desc)
        for k in range(len(ids)):
            col = get_column_letter(2 + k * 3)
            highlight(ws, f"{col}3:{col}{last_row}")
        ws.freeze_panes = "B3"

    # 4. One sheet per top buyer / seller
    for prefix, label, ids in [("B", "Buyer", buyers), ("S", "Seller", sellers)]:
        for n, pid in enumerate(ids, 1):
            ws = wb.create_sheet(f"{prefix}{n}")
            write_block(ws, 1, 1, f"{label} {n}: {names.get(pid, pid)} ({pid})", True)
            write_block(ws, 2, 1, f"Net change over period: {net[pid]:,.0f}")
            for j, h in enumerate(["Date", "Change", "Stake%", "Holding"], 1):
                write_block(ws, 3, j, h, True)
            for r, d in enumerate(dates_desc, 4):
                write_block(ws, r, 1, d.strftime("%Y/%m/%d"))
                write_block(ws, r, 2, change.at[d, pid]).number_format = "#,##0;-#,##0;0"
                write_block(ws, r, 3, stake.at[d, pid] / 100).number_format = "0.00%"
                write_block(ws, r, 4, hold.at[d, pid]).number_format = "#,##0"
            highlight(ws, f"B4:B{3 + len(dates_desc)}")

    for ws in wb.worksheets:
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 14
    wb["CCASS Holdings"].column_dimensions["C"].width = 42
    wb["CCASS Holding Changes"].column_dimensions["B"].width = 42
    wb.save(out)
    print(f"Saved {out}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--stock", default=STOCK_CODE)
    ap.add_argument("--days", type=int, default=DAYS)
    ap.add_argument("--top", type=int, default=TOP_N)
    ap.add_argument("--delay", type=float, default=1.5, help="seconds between requests")
    ap.add_argument("--all-days", action="store_true", help="include weekends (default: weekdays only)")
    ap.add_argument("--issued", type=float, help="total issued shares, if it can't be read automatically")
    a, _ = ap.parse_known_args()   # ignores the extra "-f ..." arg Colab/Jupyter adds
    stock = str(a.stock).strip().upper().replace(".HK", "").zfill(5)
    if not stock.isdigit():
        raise SystemExit(f"Stock code '{a.stock}' doesn't look right - use digits only, e.g. 1985 or 700.")
    print(f"Running CCASS history for {stock}, last {a.days} days, top {a.top}")

    data = scrape(stock, a.days, a.delay, not a.all_days)
    if len(data) < 2:
        raise SystemExit("Not enough days scraped - check the stock code or whether HKEX changed its page.")
    hold, names = build_frames(data)

    if a.issued:
        issued = a.issued
    else:
        # Always take issued shares from a fresh fetch of the latest day (1 request)
        s = requests.Session()
        latest = max(data)
        res = fetch_day(s, _hidden_fields(s), stock, latest)
        issued = res["issued"]
        if not issued:
            raise SystemExit("Could not read issued shares. Rerun with --issued <number>, "
                             "taken from the HKEX CCASS search page for this stock.")
        print(f"Issued shares (from HKEX, {latest}): {issued:,.0f}")
    if hold.iloc[-1].sum() > issued:
        print("WARNING: total in CCASS exceeds issued shares - check the issued figure.")

    out = f"{stock}_HKEX_Historical_Analysis_{hold.index[0]:%Y-%m-%d}_to_{hold.index[-1]:%Y-%m-%d}.xlsx"
    build_workbook(hold, names, issued, a.top, out)
