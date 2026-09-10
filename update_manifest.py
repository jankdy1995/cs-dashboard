#!/usr/bin/env python3
"""reports_manifest.json aus einem Google-Drive-Ordner-Listing erzeugen.

Die Monatsautomatik listet den Report-Ordner über den Drive-Connector und
speichert die Datei-Liste (das "files"-Array aus search_files) als JSON.
Dieses Skript macht daraus das Manifest für die Dashboard-Download-Section:
nur PDFs, Monat aus dem Dateinamen, neueste zuerst, alle Monate.

Usage: python3 update_manifest.py <listing.json> <reports_manifest.json>
"""
import json
import re
import sys

MONTHS = {
    'january': (1, 'Januar'), 'januar': (1, 'Januar'),
    'february': (2, 'Februar'), 'februar': (2, 'Februar'),
    'march': (3, 'März'), 'maerz': (3, 'März'), 'märz': (3, 'März'),
    'april': (4, 'April'),
    'may': (5, 'Mai'), 'mai': (5, 'Mai'),
    'june': (6, 'Juni'), 'juni': (6, 'Juni'),
    'july': (7, 'Juli'), 'juli': (7, 'Juli'),
    'august': (8, 'August'),
    'september': (9, 'September'),
    'october': (10, 'Oktober'), 'oktober': (10, 'Oktober'),
    'november': (11, 'November'),
    'december': (12, 'Dezember'), 'dezember': (12, 'Dezember'),
}
_MONTH_RE = re.compile('(' + '|'.join(sorted(MONTHS, key=len, reverse=True)) +
                       r')[ _-]?(\d{2,4})', re.IGNORECASE)


def parse_month(title):
    """('2026-08', 'August 2026') aus einem Dateinamen, sonst None."""
    m = _MONTH_RE.search(title)
    if not m:
        return None
    num, de = MONTHS[m.group(1).lower()]
    yr = m.group(2)
    year = 2000 + int(yr) if len(yr) == 2 else int(yr)
    return f'{year:04d}-{num:02d}', f'{de} {year}'


def file_kind(f):
    """'pdf', 'pptx' oder None (z. B. Vorlage / anderes)."""
    ext = str(f.get('fileExtension', '')).lower()
    mt = str(f.get('mimeType', ''))
    title = str(f.get('title', '')).lower()
    if ext == 'pdf' or mt == 'application/pdf' or title.endswith('.pdf'):
        return 'pdf'
    if ext == 'pptx' or 'presentationml' in mt or title.endswith('.pptx'):
        return 'pptx'
    return None


def dl_link(f):
    fid = f.get('id')
    if fid:
        return f'https://drive.google.com/uc?export=download&id={fid}'
    return f.get('viewUrl') or f.get('webViewLink')


def build(files):
    """Ein Eintrag je Monat mit pdf- und/oder pptx-Link (neueste Datei je Typ)."""
    by_key = {}
    for f in files:
        kind = file_kind(f)
        if not kind:
            continue
        parsed = parse_month(f.get('title', ''))
        if not parsed:                      # z. B. Vorlage ohne Monat → überspringen
            continue
        key, label = parsed
        e = by_key.setdefault(key, {'label': label, 'key': key, '_mt': {}})
        mt = str(f.get('modifiedTime', ''))
        if kind not in e or mt >= e['_mt'].get(kind, ''):
            e[kind] = dl_link(f)
            e['_mt'][kind] = mt
    out = []
    for v in sorted(by_key.values(), key=lambda x: x['key'], reverse=True):
        item = {'label': v['label'], 'key': v['key']}
        if v.get('pdf'):
            item['pdf'] = v['pdf']
        if v.get('pptx'):
            item['pptx'] = v['pptx']
        out.append(item)
    return out


def main():
    listing = json.load(open(sys.argv[1], encoding='utf-8'))
    files = listing.get('files', listing) if isinstance(listing, dict) else listing
    manifest = build(files)
    with open(sys.argv[2], 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    print(f'{len(manifest)} Report(s) ins Manifest geschrieben: '
          + ', '.join(m['label'] for m in manifest))


if __name__ == '__main__':
    main()
