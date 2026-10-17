#!/usr/bin/env python3
"""
Rebuild the data embedded in index.html from the Excel workbook.

    python scripts/build_dashboard.py                      # data/*.xlsx (exactly one) -> index.html (in place)
    python scripts/build_dashboard.py --check              # build in memory, compare with current index.html, do not write
    python scripts/build_dashboard.py --xlsx X --html Y    # custom paths

How it works
------------
index.html is both the template and the output.  It contains a number of
`const NAME = <data>;` statements.  For every NAME in BLOCKS below, this script
rebuilds the value from the workbook and swaps ONLY that statement.  All
layout / CSS / JS in index.html is left untouched, so the script is idempotent.

Needs:  pip install openpyxl
Cached formula values are used (data_only=True) -> save the workbook in Excel
before committing so formulas have cached results.
"""
import argparse
import json
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent

# ──────────────────────────────────────────────────────────────── helpers
def num(v, nd=None):
    """Cell -> number (or None). Strings like '' / '-' / '#N/A' -> None."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return round(v, nd) if nd is not None else v
    return None


def intish(v):
    """Whole floats -> int so JSON prints 75 not 75.0."""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def clean(v, nd=None):
    v = num(v, nd)
    return intish(v) if v is not None else None


def s(v):
    return None if v is None else str(v).strip()


MONTH_RE = re.compile(r"^\s*([A-Za-z]+'\d\d)")
MONTH_ANY_RE = re.compile(r"([A-Za-z]+'\d\d)")
ABBR = {"August": "Aug", "September": "Sep", "October": "Oct", "November": "Nov", "December": "Dec",
        "January": "Jan", "February": "Feb"}


def month_key(header, short=False):
    """'August'26 STT' -> "Aug'26"; "March'26 STT" stays "March'26" (matches the HTML's existing labels)."""
    m = (MONTH_ANY_RE.search if short else MONTH_RE.match)(header or "")
    if not m:
        return None
    k = m.group(1)
    name, yr = k.split("'")
    if short:                       # ByPartner sheet uses Jan'26, Feb'26, Mar'26 ... Sep'26
        return f"{name[:3]}'{yr}"
    if name in ("August", "September", "October", "November", "December"):
        name = ABBR[name]
    if name == "January":
        name = "Jan"
    if name == "February":
        name = "Feb"
    return f"{name}'{yr}"


def rows_of(wb, startswith):
    for ws in wb.worksheets:
        if ws.title.strip().lower().startswith(startswith.lower()):
            return [list(r) for r in ws.iter_rows(values_only=True)]
    raise KeyError(f"sheet starting with {startswith!r} not found; have {[w.title for w in wb.worksheets]}")


# ──────────────────────────────────────────────────────────────── block builders
def tva_month_groups(hdr):
    """Find (month, col_target, col_stt, col_ach, col_status) groups in the TVA sheet header row."""
    groups = []
    i = 0
    while i < len(hdr):
        h = hdr[i] or ""
        if h.replace("\n", " ").strip().lower().startswith("total target"):
            mk = month_key(h.split("For", 1)[-1].strip())
            groups.append((mk, i, i + 1, i + 2, i + 3))
            i += 4
        else:
            i += 1
    return groups


def build_tva(wb):
    """TSA_TVA_DATA : [{tsaId, months:[{month,target,stt,ach,status}]}]  (2 dp, ach in %)"""
    rows = rows_of(wb, "TSA Wise Target Vs")
    hdr = rows[1]
    groups = tva_month_groups(hdr)
    out = []
    for r in rows[2:]:
        tsa = s(r[4])
        if not tsa:
            continue
        months = []
        for mk, ct, cs, ca, cst in groups:
            t, st, a = num(r[ct]), num(r[cs]), num(r[ca])
            if t is None and st is None:
                continue
            months.append({"month": mk, "target": round(t, 2), "stt": round(st, 2),
                           "ach": round((a if a is not None else (st / t if t else 0)) * 100, 2),
                           "status": s(r[cst])})
        out.append({"tsaId": tsa, "months": months})
    return out


def build_data(wb):
    """DATA : channel-outlet rows joined with TVA rows (1 dp), + final CTG South total row."""
    ch = rows_of(wb, "TSA Wise By")
    tv = rows_of(wb, "TSA Wise Target Vs")
    groups = tva_month_groups(tv[1])
    by_loc = {s(r[3]): r for r in tv[2:] if s(r[3])}
    total_tv = next(r for r in tv[2:] if not s(r[3]) and s(r[1]) == "CTG South")
    keys = ["Silver", "Gold", "Platinum", "Partner", "RCC", "BCC", "TRC", "Progoti", "GT", "Total Outlet"]

    def months_of(r):
        out = []
        for mk, ct, cs, ca, cst in groups:
            t, st = num(r[ct]), num(r[cs])
            if t is None and st is None:
                continue
            ach = round((st / t) * 100, 1) if t else 0
            out.append({"month": mk, "target": round(t, 1), "stt": round(st, 1), "ach": intish(ach),
                        "status": "Target Achieved" if ach >= 100 else "Target Not Achieved"})
        return out

    out = []
    for r in ch[1:]:
        terr, loc, name = s(r[0]), s(r[1]), s(r[2])
        if not terr or terr == "Territory":
            break
        outlets = {k: (clean(r[3 + i]) or 0) for i, k in enumerate(keys)}
        if loc:
            t = by_loc[loc]
            out.append({"location": loc, "tsaId": s(t[4]), "tsaName": name, "territory": terr,
                        "months": months_of(t), "outlets": outlets})
        else:  # CTG South total row
            out.append({"location": 0, "tsaId": 0, "tsaName": 0, "territory": terr,
                        "months": months_of(total_tv), "outlets": outlets})
    return out


def build_dist(wb):
    """DIST_DATA : {BAT:{distributors,routeCount,sectionCount,outletCount,logisticsCount}, ...}"""
    rows = rows_of(wb, "Distribution Overview")
    out, cur = {}, None
    for r in rows:
        if r[2] in ("BAT", "AKTC", "JTI", "PMI") and r[1] != "Area" or (r[1] == "Area" and r[2] in ("BAT",)):
            cur = r[2]
        elif r[1] is None and r[2] in ("BAT", "AKTC", "JTI", "PMI"):
            cur = r[2]
        if r[1] == "CTG South" and cur:
            out[cur] = {"distributors": intish(r[2]), "routeCount": intish(r[3]), "sectionCount": intish(r[4]),
                        "outletCount": intish(r[5]), "logisticsCount": intish(r[6])}
            cur = None
    # BAT header is on the 'Area / BAT' row, others are on their own row above 'Area' – handled above.
    return out


def build_pl(wb):
    rows = rows_of(wb, "Price List")
    out = []
    company = None
    for r in rows[2:]:
        if r[2] is None:
            continue
        if r[1]:
            company = s(r[1])
        vals = [clean(r[i], 2) for i in range(3, 13)]
        out.append([company, s(r[2])] + vals)
    return out


def build_pw(wb):
    """PW_DATA : Partner winner numbers per TSA location."""
    rows = rows_of(wb, "TSA Wise Partner Winner")
    top, sub = rows[0], rows[1]
    groups = []  # (month, [(label,col)...])
    n = len(top)
    for i, h in enumerate(top):
        if h and "winner number" in str(h).lower():
            mk = month_key(h)
            j = i + 1
            while j < n and top[j] is None:
                j += 1
            cols = []
            for c in range(i, j):
                lab = (sub[c] or "").strip()
                lab = "Total" if lab.lower().startswith("total") else lab
                cols.append((lab, c))
            groups.append((mk, cols))
    out = []
    for r in rows[2:]:
        if not s(r[3]):
            continue
        d = {"location": s(r[2]), "tsa_id": s(r[3]), "partner": clean(r[4])}
        for mk, cols in groups:
            m = {}
            for lab, c in cols:
                v = clean(r[c])
                if v is not None:
                    m[lab] = v
            d[mk] = m
        out.append(d)
    return out


def build_ba(wb):
    rows = rows_of(wb, "CTG South Base")
    sections, cur = {}, None
    for r in rows:
        if r[0] in ("Area", "DH", "Territory", "Distribution Point") and r[1] == "B&H Family":
            cur = r[0]
            sections[cur] = {"brands": [str(x) for x in r[1:9]], "rows": []}
        elif cur and r[0]:
            sections[cur]["rows"].append((s(r[0]), [round(float(x or 0), 2) for x in r[1:9]]))
        elif cur and r[0] is None:
            cur = None
    return sections


def build_ylp(wb):
    rows = rows_of(wb, "YLP Point")
    T, L, R, D = [], [], [], []
    ix = {"T": {}, "L": {}, "R": {}}

    def idx(kind, lst, v):
        if v not in ix[kind]:
            ix[kind][v] = len(lst)
            lst.append(v)
        return ix[kind][v]

    for r in rows[4:]:
        if not r[1]:
            continue
        t = idx("T", T, s(r[2]))
        l = idx("L", L, s(r[8]) if r[8] is not None else "#N/A")
        rr = idx("R", R, s(r[9]) if r[9] is not None else "")
        vals = [clean(r[c], 2) or 0 for c in list(range(10, 25)) + [26]]
        D.append([t, l, rr, s(r[4])] + vals + [s(r[3])])
    return {"T": T, "L": L, "R": R, "D": D}


# ──────────────────────────────────────────────────────────────── KPI
TYPE_RENAME = {"CM Sup KPI": "MS KPI"}      # the dashboard calls this block "MS KPI"


def build_kpi(wb):
    """KPI_RECORDS_SOURCE : one record per (row, non-empty column)."""
    rows = rows_of(wb, "KPI")
    hdr = rows[0]
    out = []
    for r in rows[1:]:
        typ, mon = s(r[0]), s(r[1])
        if not typ or not mon:
            continue
        orig = typ
        typ = TYPE_RENAME.get(typ, typ)
        for c in range(2, len(hdr)):
            v = r[c]
            own = str(hdr[c] or "").startswith(orig + ":")      # column belongs to this KPI type
            if v is None or str(v).strip() == "":
                if not own:
                    continue
                v = "\u2014"                                      # keep an explicit dash for empty own cells
            elif isinstance(v, float) and 0 <= v <= 1 and not float(v).is_integer():
                v = f"{round(v * 100, 2):g}%"      # percentages stored as fractions
            elif isinstance(v, float) and v.is_integer():
                v = str(int(v))
            out.append({"type": typ, "month": mon, "field": hdr[c], "value": str(v)})
    return out


# ──────────────────────────────────────────────────────────────── ByPartner
def bp_groups(rows):
    """Parse the 2-row header of 'ByPartner Target Vs Achievement' into
    {month: {'volume': {metric: [(label,col)]}, 'memo': {...}, 'super': ...}}"""
    top, sub = rows[1], rows[2]
    n = len(top)
    starts = [c for c in range(n) if top[c] and c >= 13]
    out, last_month, last_scope = {}, None, "memo"
    for k, c in enumerate(starts):
        h = str(top[c])
        end = starts[k + 1] if k + 1 < len(starts) else n
        low = h.lower()
        if low.startswith("super over"):
            out.setdefault("_super", {})["cols"] = list(range(c, end))
            continue
        mk = month_key(h, short=True)
        if "remaining" in low:
            mk, scope, metric = last_month, "volume", "remaining"
        else:
            if "memo" in low:
                scope = "memo"
            elif "volume" in low:
                scope = "volume"
            else:                    # e.g. "Achievement (%) (February'26)" -> inherits the previous group's scope
                scope = last_scope if mk == last_month else "memo"
            if "%" in low:
                metric = "pct"
            elif "winner" in low:
                metric = "status"
            elif "achievement" in low:
                metric = "ach"
            elif "target" in low:
                metric = "target"
            else:
                continue
            last_month, last_scope = mk, scope
        cols = []
        for cc in range(c, end):
            lab = str(sub[cc] or "").strip()
            lab = "Total" if lab.lower().startswith("total") else lab
            cols.append((lab, cc))
        out.setdefault(mk, {}).setdefault(scope, {})[metric] = cols
    return out


def _r6(v):
    return None if v is None else round(v, 6)


def _letter(v):
    v = str(v or "").strip()
    return {"Winner": "W", "Failure": "F"}.get(v, "")


def build_bp(wb):
    rows = rows_of(wb, "ByPartner")
    G = bp_groups(rows)
    data = [r for r in rows[3:] if r[4]]
    months = [m for m in G if not m.startswith("_")]
    # brand order per month (memo target columns)
    brand_cfg = {m: [l for l, _ in G[m]["memo"]["target"]] for m in months if "memo" in G[m]}
    records, by_month, split_aug, super_over = [], {m: {} for m in brand_cfg}, {}, {}
    for r in data:
        code = s(r[4])
        base = {"Territory": s(r[1]), "TSA Location": s(r[11]), "Route / Section": s(r[3]), "Outlet Code": code,
                "Partner Name": s(r[5]), "Sub_Channel": s(r[9]), "Owner Name": s(r[6]), "TSA Name": s(r[10])}
        for m in months:
            g = G[m]

            def first(scope, metric, nd=None):
                cols = g.get(scope, {}).get(metric)
                if not cols:
                    return None
                return r[cols[0][1]]

            vt, va = first("volume", "target"), first("volume", "ach")
            vp, vs = first("volume", "pct"), first("volume", "status")
            rec = dict(base)
            rec["Month"] = m
            def vol(v):                      # number -> number, text such as '-' kept as is, empty -> ''
                return _r6(v) if num(v) is not None else (s(v) or "")
            rec["Volume Target"] = vol(vt)
            rec["Volume Achievement"] = vol(va)
            rec["Volume Achievement %"] = vol(vp)
            rec["Volume Winner Status"] = s(vs) or ""
            mem = g.get("memo")
            if mem:
                tg = [num(r[c]) for _, c in mem["target"]]
                ac = [num(r[c]) for _, c in mem["ach"]]
                pc = [num(r[c]) for _, c in mem["pct"]]
                st = [_letter(r[c]) for _, c in mem["status"]]
                # fallbacks when pct / status are formulas without cached values
                for i in range(len(tg)):
                    if i < len(pc) and pc[i] is None and tg[i] and ac[i] is not None:
                        pc[i] = ac[i] / tg[i]
                    if i < len(st) and not st[i] and i < len(pc) and pc[i] is not None:
                        st[i] = "W" if pc[i] >= 1 else "F"
                flat = [intish(_r6(x)) if x is not None else None for x in tg + ac + pc] + st
                if any(x not in (None, "") for x in flat):          # outlets with no memo data in that month are omitted
                    by_month[m][code] = flat
                if m == "Aug'26":
                    split_aug[code] = [intish(x) for x in (tg[0], tg[1], ac[0], ac[1])]
                    T, A = (tg[0] or 0) + (tg[1] or 0), (ac[0] or 0) + (ac[1] or 0)
                    rec.update({"Memo Target": intish(T), "Memo Achievement": intish(A),
                                "Memo Achievement %": intish(round(A / T, 6)) if T else "",
                                "Memo Winner Status": ("Winner" if T and A / T >= 1 else "Failure")})
                elif m == "Sep'26":
                    rec.update({"Memo Target": "", "Memo Achievement": "", "Memo Achievement %": "", "Memo Winner Status": ""})
                else:
                    rec.update({"Memo Target": intish(tg[0]) if tg[0] is not None else "",
                                "Memo Achievement": intish(ac[0]) if ac[0] is not None else "",
                                "Memo Achievement %": intish(_r6(pc[0])) if pc[0] is not None else "",
                                "Memo Winner Status": {"W": "Winner", "F": "Failure"}.get(st[0], "")})
            rem = g.get("volume", {}).get("remaining")
            if rem and m == "Aug'26":
                rec["Remaining"] = num(r[rem[0][1]])
            records.append(rec)
        sc = G.get("_super", {}).get("cols")
        if sc:
            t, stt, remn, ap = (num(r[c]) for c in sc[:4])
            if t is not None:
                if ap is None and t:
                    ap = (stt or 0) / t
                if remn is None:
                    remn = t - (stt or 0)
                super_over[code] = {"target": t, "stt": stt, "remaining": remn, "achPct": ap,
                                    "winnerStatus": "Winner" if ap is not None and ap >= 1 else "Failure"}
    return {"BP_RECORDS_SOURCE": records, "BP_MEMO_SPLIT_AUG": split_aug, "BP_SUPER_OVER": super_over,
            "BP_MEMO_BY_MONTH": by_month, "BP_MONTH_BRAND_CONFIG": brand_cfg}


# ──────────────────────────────────────────────────────────────── JS literal blocks
def js_json(v):
    return json.dumps(v, ensure_ascii=False).replace("</", "<\\/")


def js_dist(d):
    lines = ["const DIST_DATA = {"]
    for k in ("BAT", "AKTC", "JTI", "PMI"):
        if k in d:
            x = d[k]
            lines.append(f"  {k}: {{ distributors:{x['distributors']}, routeCount:{x['routeCount']}, "
                         f"sectionCount:{x['sectionCount']}, outletCount:{x['outletCount']}, "
                         f"logisticsCount:{x['logisticsCount']} }},")
    lines.append("};")
    return "\n".join(lines)


def js_ba(sec):
    names = sec["Area"]["brands"]

    def brands(vals):
        return js_json(dict(zip(names, vals)))

    a_name, a_vals = sec["Area"]["rows"][0]
    out = ["const BA_SHEET_DATA = {",
           f"  area: {{ name: {js_json(a_name)}, brands: {brands(a_vals)} }},",
           "  dh: [" + ",\n    ".join(f"{{ dh: {js_json(n)}, brands: {brands(v)} }}" for n, v in sec["DH"]["rows"]) + "],",
           "  dhPoint: [" + ",\n    ".join(f"{{ point: {js_json(n)}, brands: {brands(v)} }}" for n, v in sec["Distribution Point"]["rows"]) + "],",
           "  territory: [" + ",\n    ".join(f"{{ territory: {js_json(n)}, brands: {brands(v)} }}" for n, v in sec["Territory"]["rows"]) + "],",
           f"  BRANDS: {js_json(names)}",
           "};"]
    return "\n".join(out)


def js_pw(rows):
    return "const PW_DATA = [\n" + ",\n".join("  " + js_json(r) for r in rows) + "\n];"


# ──────────────────────────────────────────────────────────────── writer
JSON_BLOCKS = ["DATA", "TSA_TVA_DATA", "PL_ROWS", "YLP_RAW", "KPI_RECORDS_SOURCE", "BP_RECORDS_SOURCE",
               "BP_MEMO_SPLIT_AUG", "BP_SUPER_OVER", "BP_MEMO_BY_MONTH", "BP_MONTH_BRAND_CONFIG"]
DEC = json.JSONDecoder()


def replace_json_block(html, name, value):
    m = re.search(r"^const " + name + r" *= *", html, re.M)
    if not m:
        raise SystemExit(f"[ERROR] 'const {name} =' not found in index.html - was the HTML restructured?")
    try:
        old, end = DEC.raw_decode(html, m.end())
    except json.JSONDecodeError as e:
        raise SystemExit(f"[ERROR] existing {name} in index.html is not plain JSON: {e}")
    new_txt = js_json(value)
    return html[:m.end()] + new_txt + html[end:], old != value


def replace_literal_block(html, name, open_pat, new_text):
    m = re.search(r"^const " + name + r" *= *" + open_pat + r".*?^[\}\]];", html, re.M | re.S)
    if not m:
        raise SystemExit(f"[ERROR] block {name} not found in index.html")
    return html[:m.start()] + new_text + html[m.end():], html[m.start():m.end()] != new_text


def build_all(wb):
    bp = build_bp(wb)
    vals = {"DATA": build_data(wb), "TSA_TVA_DATA": build_tva(wb), "PL_ROWS": build_pl(wb),
            "YLP_RAW": build_ylp(wb), "KPI_RECORDS_SOURCE": build_kpi(wb), **bp}
    lit = {"DIST_DATA": ("\\{", js_dist(build_dist(wb))),
           "BA_SHEET_DATA": ("\\{", js_ba(build_ba(wb))),
           "PW_DATA": ("\\[", js_pw(build_pw(wb)))}
    return vals, lit


def sanity(vals, lit):
    must = {"DATA": 20, "TSA_TVA_DATA": 15, "PL_ROWS": 10, "YLP_RAW": 1, "KPI_RECORDS_SOURCE": 50,
            "BP_RECORDS_SOURCE": 1000, "BP_SUPER_OVER": 100, "BP_MEMO_BY_MONTH": 3}
    for k, n in must.items():
        size = len(vals[k]["D"]) if k == "YLP_RAW" else len(vals[k])
        if size < n:
            raise SystemExit(f"[ERROR] block {k} has only {size} items (expected >= {n}); refusing to overwrite index.html")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", default=None, help="default: the only .xlsx inside data/")
    ap.add_argument("--html", default=str(ROOT / "index.html"))
    ap.add_argument("--out", default=None, help="write here instead of overwriting --html")
    ap.add_argument("--check", action="store_true", help="dry run: report which blocks would change, write nothing")
    a = ap.parse_args()

    if a.xlsx is None:
        found = sorted((ROOT / "data").glob("*.xlsx"))
        if len(found) != 1:
            raise SystemExit(f"[ERROR] keep exactly ONE .xlsx in data/ (found {len(found)}: {[f.name for f in found]})")
        a.xlsx = str(found[0])
    print(f"Excel: {a.xlsx}")
    wb = openpyxl.load_workbook(a.xlsx, data_only=True)
    vals, lit = build_all(wb)
    sanity(vals, lit)
    html = Path(a.html).read_text(encoding="utf-8")
    report = []
    for name in JSON_BLOCKS:
        html, changed = replace_json_block(html, name, vals[name])
        report.append((name, changed))
    for name, (op, text) in lit.items():
        html, changed = replace_literal_block(html, name, op, text)
        report.append((name, changed))
    for name, changed in report:
        print(f"  {'UPDATED  ' if changed else 'unchanged'}  {name}")
    n = sum(c for _, c in report)
    if a.check:
        print(f"[check] {n} block(s) would change; nothing written.")
        return
    target = Path(a.out or a.html)
    if html != Path(a.html).read_text(encoding="utf-8") or a.out:
        target.write_text(html, encoding="utf-8")
        print(f"[done] wrote {target} ({n} block(s) changed)")
    else:
        print("[done] index.html already up to date")


if __name__ == "__main__":
    main()
