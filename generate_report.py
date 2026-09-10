#!/usr/bin/env python3
"""NeoTaste CS Monthly Report generator.

Builds the C-level monthly report (PPTX, 7 slides) from the NeoTaste CS KPI
sheet, in the exact layout of the established template. Values are computed
with the SAME definitions as the live dashboard (build_dashboard.extract) so
report and dashboard always agree. Slide 7 (Summary) is left blank for the
CS lead to fill in.

Usage:
    python3 generate_report.py <kpi.xlsx> <template.pptx> <out.pptx> <month> [year]
    # month: 1-12  (the reporting month)
"""
import sys
import calendar
import datetime
import importlib.util

import openpyxl
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData

# ---- load build_dashboard.extract (single source of truth for parsing) ----
_spec = importlib.util.spec_from_file_location(
    'bd', __file__.rsplit('/', 1)[0] + '/build_dashboard.py')
bd = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(bd)

MONTHS_EN = ['', 'January', 'February', 'March', 'April', 'May', 'June',
             'July', 'August', 'September', 'October', 'November', 'December']
MONTHS_DE = bd.MONTHS_DE

GREEN = RGBColor(0x5F, 0xE3, 0xA1)
RED = RGBColor(0xF2, 0x54, 0x5B)
AMBER = RGBColor(0xED, 0xA1, 0x00)
FLAT_FEE = 2000.0  # MoinAI monthly flat fee (€)

# good direction per KPI: True = higher is better
GOOD_UP = {
    'ticket_volume': False, 'mtfr': False, 'csat': True, 'automation': True,
    'savings_rdr': True, 'cpt': False, 'monthly_savings': True, 'share_conv': True,
}


# ----------------------------- helpers -------------------------------------
def n0(v):
    return '–' if v is None else f'{round(v):,}'


def eur0(v):
    return '–' if v is None else f'€{round(v):,}'


def pct0(v):
    return '–' if v is None else f'{round(v * 100)}%'


def delta_str(cur, prev):
    """Return ('↑ +7 %'|'↓ −7 %'|'–', color) — color by good/bad is set by caller."""
    if cur is None or prev in (None, 0):
        return '–', None
    ch = (cur - prev) / prev
    arrow = '↑' if ch >= 0 else '↓'
    sign = '+' if ch >= 0 else '−'
    return f'{arrow} {sign}{abs(round(ch * 100))} %', ch


def month_of_kw(kw, year):
    try:
        return datetime.date.fromisocalendar(year, kw, 4).month
    except ValueError:
        return None


def group(rows, year):
    g = {}
    for r in rows:
        m = month_of_kw(r['kw'], year)
        if m:
            g.setdefault(m, []).append(r)
    return g


def _sum(v):
    v = [x for x in v if x is not None]
    return sum(v) if v else None


def _avg(v):
    v = [x for x in v if x is not None]
    return sum(v) / len(v) if v else None


def savings_per_day(xlsx):
    """Monatliche Chatbot-Savings aus Monthly_Summary 'Savings (Ø/Tag)'.

    Der Tab hat eine rollierende, teils doppelte Monatsspalten-Struktur; wir
    nehmen je Monat die erste Fundstelle (von links). Rückgabe: {monat: €/Tag}.
    """
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ms = next((wb[s] for s in wb.sheetnames if 'monthly_summary' in s.lower()), None)
    if ms is None:
        return {}
    hdr = list(ms[3])
    srow = None
    for r in ms.iter_rows(min_row=4, max_row=40):
        lbl = str(r[0].value or '')
        if 'aving' in lbl and 'Tag' in lbl:
            srow = r
            break
    if srow is None:
        return {}
    per_day = {}
    for i, cell in enumerate(hdr):
        h = cell.value
        if isinstance(h, datetime.datetime) and i < len(srow):
            mon = h.month
            val = srow[i].value
            if mon not in per_day and isinstance(val, (int, float)):
                per_day[mon] = float(val)
    return per_day


def monthly_savings_of(per_day, m, year):
    """Ø/Tag × Tage im Monat (Monthly_Summary-Definition)."""
    if m not in per_day:
        return None
    return per_day[m] * calendar.monthrange(year, m)[1]


# KPI-Zeilen im Monthly_Summary → interne Feldnamen (Teilstring-Match auf Spalte A)
_MS_ROWS = [
    ('created tickets (gesamt)', 'created_hs'),
    ('chatbot conversations', 'chatbot_conv'),
    ('median first reply', 'mtfr'),
    ('csat happy', 'csat'),
    ('automation rate all tickets', 'auto_all'),
    ('net money contribution', 'net_contribution'),
    ('average handling time', 'aht'),
]


def read_monthly_summary(xlsx):
    """Key-Metrics je Monat aus dem Monthly_Summary-Tab.

    Der Tab hat pro Monat eine datierte Spalte; die '…-26'-Spalten sind die
    Ist-Werte und stehen rechts, daher gewinnt die am weitesten rechts
    stehende Spalte eines Monats. Rückgabe: {monat: {feld: wert}}.
    """
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ms = next((wb[s] for s in wb.sheetnames if 'monthly_summary' in s.lower()), None)
    if ms is None:
        return {}
    hdr = list(ms[3])
    # Spalten-Index → Monat (nur datierte Spalten)
    month_cols = [(i, h.value.month) for i, h in enumerate(hdr)
                  if isinstance(h.value, datetime.datetime)]
    # Zeilen einsammeln
    field_row = {}
    cpt_row = None
    for r in ms.iter_rows(min_row=4, max_row=40):
        name = str(r[0].value or '').strip()
        low = name.lower()
        for needle, field in _MS_ROWS:
            if needle in low:
                field_row[field] = r
        if low == 'cost per ticket':      # nur die Costs-Zeile, nicht '… in Chatbot'
            cpt_row = r
    out = {}
    for i, mon in month_cols:            # links→rechts: rechte Spalte überschreibt
        d = out.setdefault(mon, {})
        for field, row in field_row.items():
            v = row[i].value
            if isinstance(v, (int, float)):
                d[field] = float(v)
        if cpt_row is not None and isinstance(cpt_row[i].value, (int, float)):
            d['cpt'] = float(cpt_row[i].value)
    return out


def apply_monthly_summary(v, ms_month):
    """Key-Metrics eines Monats mit Monthly_Summary-Werten überschreiben."""
    if v is None or not ms_month:
        return
    hs, cb = ms_month.get('created_hs'), ms_month.get('chatbot_conv')
    if hs is not None:
        v['tickets_hs'] = hs
    if cb is not None:
        v['tickets_cb'] = cb
    if hs is not None or cb is not None:
        v['tickets_total'] = (hs or 0) + (cb or 0)
    for src, dst in (('mtfr', 'mtfr'), ('csat', 'csat'), ('auto_all', 'automation'),
                     ('net_contribution', 'savings_rdr'), ('cpt', 'cpt'), ('aht', 'aht')):
        if ms_month.get(src) is not None:
            v[dst] = ms_month[src]


# --------------------------- computation -----------------------------------
def compute_month(d, m, year):
    """All report values for month m (None if no data)."""
    ghs = group(d['hubspot'], year)
    gma = group(d['moinai'], year)
    grf = group(d['refunds'], year)
    gtm = group(d['team'], year)
    costs = {c['month']: c for c in d['costs']}
    if m not in ghs and MONTHS_DE[m] not in costs:
        return None
    hsr, mar, rfr, tmr = ghs.get(m, []), gma.get(m, []), grf.get(m, []), gtm.get(m, [])
    c = costs.get(MONTHS_DE[m])

    # weighted AHT over team messages
    wsum = wmsg = 0.0
    for r in tmr:
        for a in ('Eli', 'Jeanine', 'Vivien', 'Jan'):
            mm, ah = r[a]['messages'], r[a]['aht']
            if mm and ah:
                wsum += mm * ah
                wmsg += mm

    monthly_savings = _sum([r['savings'] for r in mar])
    conv_total = _sum([r['conv_total'] for r in mar])
    v = {
        'month_en': MONTHS_EN[m], 'month_de': MONTHS_DE[m], 'year': year,
        'tickets_hs': c['tickets_hs'] if c else None,
        'tickets_cb': c['tickets_cb'] if c else None,
        'tickets_total': ((c['tickets_hs'] or 0) + (c['tickets_cb'] or 0)) if c and (c['tickets_hs'] or c['tickets_cb']) else None,
        'mtfr': _avg([r['mtfr'] for r in hsr]),
        'csat': _avg([r['csat'] for r in hsr]),
        'automation': _avg([r['auto_all'] for r in mar]),
        'savings_rdr': _sum([r['net'] for r in rfr]),
        'cpt': c['cpt'] if c else None,
        'aht': (wsum / wmsg) if wmsg else None,
        'monthly_savings': monthly_savings,
        'share_conv': _avg([r['share_chatbot'] for r in mar]),
        'roi': (monthly_savings / FLAT_FEE) if monthly_savings is not None else None,
        'conv_total': conv_total,
        'cost_per_conv': (FLAT_FEE / conv_total) if conv_total else None,
        'weeks': [r['kw'] for r in hsr],
    }

    # team monthly rows
    team = []
    on_track = 0
    for a in ('Eli', 'Jeanine', 'Vivien', 'Jan'):
        msg = _sum([r[a]['messages'] for r in tmr])
        act = _sum([r[a]['active_hours'] for r in tmr])
        contract = _sum([r[a]['contract'] for r in tmr]) or 0
        proj = _sum([r[a]['projekt'] for r in tmr]) or 0
        meet = _sum([r[a]['meeting'] for r in tmr]) or 0
        off = _sum([r[a]['off'] for r in tmr]) or 0
        target = contract - proj - meet - off
        notes = [r[a]['note'] for r in tmr if r[a]['note']]
        note = max(set(notes), key=notes.count) if notes else '–'
        ok = (act is not None and target is not None and act >= target * 0.95)
        on_track += 1 if ok else 0
        team.append({
            'name': a, 'messages': msg, 'actual': act, 'target': target,
            'off': off, 'projekt': proj, 'note': note,
            'status': 'On Track' if ok else 'At Risk',
        })
    v['team'] = team
    v['team_messages'] = _sum([t['messages'] for t in team])
    v['team_actual'] = _sum([t['actual'] for t in team])
    v['team_off'] = _sum([t['off'] for t in team])
    v['team_ontrack'] = f'{on_track}/4 On Track'
    return v


def ytd_savings(d, m, year):
    gma = group(d['moinai'], year)
    tot = 0.0
    for mm in range(1, m + 1):
        s = _sum([r['savings'] for r in gma.get(mm, [])])
        tot += s or 0
    return tot


# ----------------------------- filling -------------------------------------
def set_text(shape, text, color=None):
    tf = shape.text_frame
    para = tf.paragraphs[0]
    if para.runs:
        para.runs[0].text = text
        for extra in para.runs[1:]:
            extra.text = ''
        if color is not None:
            para.runs[0].font.color.rgb = color
    else:
        para.text = text


def set_delta(shape, key, cur, prev):
    txt, ch = delta_str(cur, prev)
    if ch is None:
        set_text(shape, '–')
        return
    good = (ch >= 0) == GOOD_UP[key]
    set_text(shape, txt, GREEN if good else RED)


def _nice_ceil(x):
    import math
    if x <= 0:
        return x
    mag = 10 ** math.floor(math.log10(x))
    for m in (1, 1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8, 10):
        if m * mag >= x * 1.02:
            return round(m * mag, 4)
    return round(10 * mag, 4)


def _nice_floor(x):
    import math
    if x <= 0:
        return 0
    mag = 10 ** math.floor(math.log10(x))
    for m in (10, 8, 6, 5, 4, 3, 2.5, 2, 1.5, 1):
        if m * mag <= x:
            return round(m * mag, 4)
    return 0


def set_chart(shape, cats, series_name, values):
    cd = CategoryChartData()
    cd.categories = [str(c) for c in cats]
    cd.add_series(series_name, values)
    shape.chart.replace_data(cd)
    # Achse nur bei Bedarf erweitern, damit keine Werte abgeschnitten werden
    nums = [x for x in values if isinstance(x, (int, float))]
    if nums:
        va = shape.chart.value_axis
        try:
            if va.maximum_scale is not None and max(nums) > va.maximum_scale:
                va.maximum_scale = _nice_ceil(max(nums))
            if va.minimum_scale is not None and min(nums) < va.minimum_scale:
                va.minimum_scale = _nice_floor(min(nums))
        except Exception:
            pass


def fill(template, out, v, prev, d, m, per_day):
    p = Presentation(template)
    S = p.slides
    myr = f"{v['month_en']} {v['year']}"
    year = v['year']
    # Snapshot: only weeks/months up to and including the reporting month m
    wk = [r for r in d['hubspot'] if (month_of_kw(r['kw'], year) or 99) <= m]
    kws = [r['kw'] for r in wk]
    win = [mm for mm in range(m - 3, m + 1) if mm >= 1]  # rolling 4-month window
    costs_by = {c['month']: c for c in d['costs']}
    gma_all = group(d['moinai'], year)

    # Slide 1
    set_text(S[0].shapes[4], f"Overview · {myr}")

    # Slide 2 — Key Metrics
    s2 = S[1].shapes
    pm = prev['month_en'] if prev else ''
    set_text(s2[2], f"{myr} · Δ vs. {pm} {v['year']}" if prev else myr)
    set_text(s2[5], f"{n0(v['tickets_cb'])} CB   {n0(v['tickets_hs'])} HS")
    set_text(s2[6], n0(v['tickets_total']))
    set_delta(s2[7], 'ticket_volume', v['tickets_total'], prev and prev['tickets_total'])
    set_text(s2[11], '–' if v['mtfr'] is None else f"{v['mtfr']:.1f}h")
    set_delta(s2[12], 'mtfr', v['mtfr'], prev and prev['mtfr'])
    set_text(s2[16], pct0(v['csat']))
    set_delta(s2[17], 'csat', v['csat'], prev and prev['csat'])
    set_text(s2[21], pct0(v['automation']))
    set_delta(s2[22], 'automation', v['automation'], prev and prev['automation'])
    set_text(s2[26], eur0(v['savings_rdr']))
    set_delta(s2[27], 'savings_rdr', v['savings_rdr'], prev and prev['savings_rdr'])
    set_text(s2[31], '–' if v['cpt'] is None else f"€{v['cpt']:.2f}")
    set_delta(s2[32], 'cpt', v['cpt'], prev and prev['cpt'])
    set_text(s2[40], '–' if v['aht'] is None else f"{v['aht']:.2f} min.")
    set_text(s2[41], 'weighted Ø over all agents')

    # Slide 3 — Ticket Volume & Response Time
    s3 = S[2].shapes
    set_text(s3[4], n0(v['tickets_total']))
    set_delta(s3[5], 'ticket_volume', v['tickets_total'], prev and prev['tickets_total'])
    set_text(s3[8], '–' if v['mtfr'] is None else f"{v['mtfr']:.1f}h")
    set_delta(s3[9], 'mtfr', v['mtfr'], prev and prev['mtfr'])
    set_chart(s3[10], kws, 'Created Tickets', [r['created'] for r in wk])
    set_chart(s3[11], kws, 'Median First Reply (h)', [r['mtfr'] for r in wk])
    blk = v['weeks']
    blk_txt = f"KW {blk[0]}–{blk[-1]}" if blk else '–'
    set_text(s3[12], f"Weekly values KW {kws[0]}–{kws[-1]} · {v['month_en']} block = {blk_txt}")

    # Slide 4 — CSAT & Cost per Ticket
    s4 = S[3].shapes
    set_text(s4[4], pct0(v['csat']))
    set_delta(s4[5], 'csat', v['csat'], prev and prev['csat'])
    set_text(s4[8], '–' if v['cpt'] is None else f"€{v['cpt']:.2f}")
    set_delta(s4[9], 'cpt', v['cpt'], prev and prev['cpt'])
    set_chart(s4[10], kws, 'CSAT %', [None if r['csat'] is None else round(r['csat'] * 100) for r in wk])
    set_chart(s4[11], [MONTHS_EN[mm] for mm in win], 'Cost per Ticket (€)',
              [costs_by.get(MONTHS_DE[mm], {}).get('cpt') for mm in win])
    set_text(s4[12], f"CSAT weekly KW {kws[0]}–{kws[-1]} · Cost per Ticket from CST_Costs sheet (monthly)")

    # Slide 5 — Chatbot & Automation
    s5 = S[4].shapes
    set_text(s5[4], eur0(v['monthly_savings']))
    set_delta(s5[5], 'monthly_savings', v['monthly_savings'], prev and prev['monthly_savings'])
    set_text(s5[9], pct0(v['automation']))
    set_delta(s5[10], 'automation', v['automation'], prev and prev['automation'])
    set_text(s5[14], pct0(v['share_conv']))
    set_delta(s5[15], 'share_conv', v['share_conv'], prev and prev['share_conv'])
    set_text(s5[19], '–' if v['roi'] is None else f"{v['roi']:.1f}×")
    set_chart(s5[23], [MONTHS_EN[mm] for mm in win], 'Savings (€)',
              [monthly_savings_of(per_day, mm, year) for mm in win])
    set_text(s5[25], f"Chatbot in numbers · {v['month_en']}")
    set_text(s5[27], n0(v['conv_total']))
    set_text(s5[29], '–' if v['cost_per_conv'] is None else f"€{v['cost_per_conv']:.2f}")
    set_text(s5[31], eur0(FLAT_FEE))
    set_text(s5[33], eur0(v.get('ytd')))

    # Slide 6 — Team
    s6 = S[5].shapes
    set_text(s6[2], f"{myr} · Eli · Jeanine · Vivien · Jan")
    tbl = s6[3].table
    for i, t in enumerate(v['team']):
        row = i + 1
        set_cell(tbl.cell(row, 1), n0(t['messages']))
        set_cell(tbl.cell(row, 2), '–' if t['actual'] is None else f"{t['actual']:.1f} h")
        set_cell(tbl.cell(row, 3), '–' if t['target'] is None else f"{round(t['target'])} h")
        set_cell(tbl.cell(row, 4), f"{round(t['off'])} h")
        set_cell(tbl.cell(row, 5), f"{round(t['projekt'])} h" if t['projekt'] else '–')
        set_cell(tbl.cell(row, 6), t['note'])
        set_cell(tbl.cell(row, 7), t['status'], GREEN if t['status'] == 'On Track' else AMBER)
    set_text(s6[6], n0(v['team_messages']))
    set_text(s6[8], '–' if v['team_actual'] is None else f"{round(v['team_actual'])} h")
    set_text(s6[10], f"{round(v['team_off'])} h")
    set_text(s6[12], v['team_ontrack'])

    # Slide 7 left intentionally blank (CS lead commentary)
    p.save(out)


def set_cell(cell, text, color=None):
    """Set a table cell's text while preserving the existing run formatting."""
    para = cell.text_frame.paragraphs[0]
    if para.runs:
        para.runs[0].text = text
        for e in para.runs[1:]:
            e.text = ''
        if color is not None:
            para.runs[0].font.color.rgb = color
    else:
        para.text = text
        if color is not None and para.runs:
            para.runs[0].font.color.rgb = color


def main():
    xlsx, template, out, month = sys.argv[1:5]
    month = int(month)
    year = int(sys.argv[5]) if len(sys.argv) > 5 else datetime.date.today().year
    d = bd.extract(xlsx)
    v = compute_month(d, month, year)
    if v is None:
        raise SystemExit(f'Keine Daten für Monat {month}/{year}')
    prev = compute_month(d, month - 1, year) if month > 1 else None

    # Key-Metrics (Slide 2 + gleiche Headline-Zahlen auf Detail-Slides) aus
    # Monthly_Summary ziehen; Wochen-Charts bleiben als Trend aus den Rohdaten.
    ms = read_monthly_summary(xlsx)
    apply_monthly_summary(v, ms.get(month))
    if prev:
        apply_monthly_summary(prev, ms.get(month - 1))

    # Chatbot-Savings: Monthly_Summary-Definition (Ø/Tag × Tage) — wie bisher
    per_day = savings_per_day(xlsx)
    v['monthly_savings'] = monthly_savings_of(per_day, month, year)
    v['roi'] = (v['monthly_savings'] / FLAT_FEE) if v['monthly_savings'] is not None else None
    if prev:
        prev['monthly_savings'] = monthly_savings_of(per_day, month - 1, year)
    # YTD: kumuliert ab erstem Datenmonat (April) bis Berichtsmonat
    first = min((month_of_kw(r['kw'], year) for r in d['hubspot']
                 if month_of_kw(r['kw'], year)), default=1)
    v['ytd'] = sum(monthly_savings_of(per_day, mm, year) or 0
                   for mm in range(first, month + 1))
    fill(template, out, v, prev, d, month, per_day)
    print(f'OK: {out} — {v["month_en"]} {year} '
          f'(Tickets {n0(v["tickets_total"])}, CSAT {pct0(v["csat"])}, '
          f'CPT €{v["cpt"]:.2f})' if v['cpt'] else f'OK: {out}')


if __name__ == '__main__':
    main()
