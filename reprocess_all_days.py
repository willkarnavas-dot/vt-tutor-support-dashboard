import csv, sys, json
sys.path.insert(0, '/home/claude')
from single_label_classifier import classify, ALL_CATS
from numbers_parser import Document

FILES = [
    ("Jan 1",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_04_42.csv"),
    ("Jan 2",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_04_54.csv"),
    ("Jan 3",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_04_59.csv"),
    ("Jan 4",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_05_10.csv"),
    ("Jan 5",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_05_19.csv"),
    ("Jan 6",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_05_38.csv"),
    ("Jan 7",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_05_48.csv"),
    ("Jan 8",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_06_01.csv"),
    ("Jan 9",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_06_11.csv"),
    ("Jan 10", "/mnt/user-data/uploads/evaluations_2026-02-19_22_06_25.csv"),
    ("Jan 11", "/mnt/user-data/uploads/evaluations_2026-02-19_22_23_14.csv"),
    ("Jan 12", "/mnt/user-data/uploads/evaluations_2026-02-19_22_23_24.csv"),
    ("Jan 13", "/mnt/user-data/uploads/evaluations_2026-02-19_22_23_39.csv"),
    ("Jan 14", "/mnt/user-data/uploads/evaluations_2026-02-19_22_23_49.csv"),
    ("Jan 15", "/mnt/user-data/uploads/evaluations_2026-02-19_22_23_59.csv"),
    ("Jan 16", "/mnt/user-data/uploads/evaluations_2026-02-19_22_24_07.csv"),
    ("Jan 17", "/mnt/user-data/uploads/evaluations_2026-02-19_22_24_27.csv"),
    ("Jan 18", "/mnt/user-data/uploads/evaluations_2026-02-19_22_24_47.csv"),
    ("Jan 19", "/mnt/user-data/uploads/evaluations_2026-02-19_22_24_56.csv"),
    ("Jan 20", "/mnt/user-data/uploads/evaluations_2026-02-19_22_25_06.csv"),
    ("Jan 21", "/mnt/user-data/uploads/evaluations_2026-02-19_22_25_39.csv"),
    ("Jan 22", "/mnt/user-data/uploads/evaluations_2026-02-19_22_31_51.csv"),
    ("Jan 23", "/mnt/user-data/uploads/evaluations_2026-02-19_22_32_01.csv"),
    ("Jan 24", "/mnt/user-data/uploads/evaluations_2026-02-19_22_32_07.csv"),
    ("Jan 25", "/mnt/user-data/uploads/evaluations_2026-02-19_22_32_17.csv"),
    ("Jan 26", "/mnt/user-data/uploads/evaluations_2026-02-19_22_32_28.csv"),
    ("Jan 27", "/mnt/user-data/uploads/evaluations_2026-02-19_22_32_46.csv"),
    ("Jan 28", "/mnt/user-data/uploads/evaluations_2026-02-19_22_32_58.csv"),
    ("Jan 29", "/mnt/user-data/uploads/evaluations_2026-02-19_22_33_08.csv"),
    ("Jan 30", "/mnt/user-data/uploads/evaluations_2026-02-19_22_33_15.csv"),
    ("Jan 31", "/mnt/user-data/uploads/evaluations_2026-02-19_22_33_25.csv"),
    ("Feb 1",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_33_33.csv"),
    ("Feb 2",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_33_53.csv"),
    ("Feb 3",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_34_03.csv"),
    ("Feb 4",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_34_14.csv"),
    ("Feb 5",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_41_40.csv"),
    ("Feb 6",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_42_07.csv"),
    ("Feb 7",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_42_19.csv"),
    ("Feb 8",  "/mnt/user-data/uploads/evaluations_2026-02-19_14_59_21.csv"),
    ("Feb 9",  "/mnt/user-data/uploads/evaluations_2026-02-19_22_42_48.csv"),
    ("Feb 10", "/mnt/user-data/uploads/evaluations_2026-02-19_22_43_10.csv"),
    ("Feb 11", "/mnt/user-data/uploads/evaluations_2026-02-19_22_43_25.csv"),
    ("Feb 12", ("numbers", "/mnt/user-data/uploads/evaluations_2026-02-19_23_05_20.numbers")),
    ("Feb 13", "/mnt/user-data/uploads/evaluations_2026-02-19_22_43_48.csv"),
    ("Feb 14", "/mnt/user-data/uploads/evaluations_2026-02-19_22_43_53.csv"),
    ("Feb 15", "/mnt/user-data/uploads/evaluations_2026-02-19_04_57_35.csv"),
    ("Feb 16", "/mnt/user-data/uploads/evaluations_2026-02-19_04_56_10.csv"),
    ("Feb 17", "/mnt/user-data/uploads/evaluations_2026-02-19_04_51_29.csv"),
    ("Feb 18", "/mnt/user-data/uploads/evaluations_2026-02-19_04_44_14.csv"),
    # Feb 19 missing
    ("Feb 20", "/mnt/user-data/uploads/evaluations_2026-02-23_16_22_12.csv"),
    ("Feb 21", "/mnt/user-data/uploads/evaluations_2026-02-23_16_24_05.csv"),
    ("Feb 22", "/mnt/user-data/uploads/evaluations_2026-02-23_16_24_17.csv"),
    ("Feb 23", ("numbers", "/mnt/user-data/uploads/evaluations_2026-02-26_15_09_35.numbers")),
    ("Feb 24", "/mnt/user-data/uploads/evaluations_2026-02-26_15_14_35.csv"),
    ("Feb 25", "/mnt/user-data/uploads/evaluations_2026-02-26_15_15_18.csv"),
    ("Feb 26", "/mnt/user-data/uploads/evaluations_2026-02-27_14_59_25.csv"),
    ("Feb 27", "/mnt/user-data/uploads/evaluations_2026-03-02_15_24_08.csv"),
    ("Feb 28", "/mnt/user-data/uploads/evaluations_2026-03-02_15_24_24.csv"),
    ("Mar 1",  "/mnt/user-data/uploads/evaluations_2026-03-02_15_24_35.csv"),
    ("Mar 2",  ("numbers", "/mnt/user-data/uploads/evaluations_2026-03-03_15_13_14.numbers")),
]

def load_file(val):
    if isinstance(val, tuple):
        doc = Document(val[1])
        sh=doc.sheets[0]; tb=sh.tables[0]; rn=list(tb.iter_rows())
        hd=[str(c.value) if c.value else '' for c in rn[0]]
        ci=hd.index('Conversation ID'); mi=hd.index('Customer Message')
        last={}
        for row in rn[1:]:
            cid=str(row[ci].value) if row[ci].value else ''
            t=str(row[mi].value) if row[mi].value else ''
            last[cid]=t
        return last
    else:
        if isinstance(val, tuple) and val[0] == 'numbers':
            from numbers_parser import Document as _Doc
            _doc = _Doc(val[1])
            _sh=_doc.sheets[0]; _tb=_sh.tables[0]; _rn=list(_tb.iter_rows())
            _hd=[str(c.value) if c.value else '' for c in _rn[0]]
            rows=[{_hd[i]:(str(_row[i].value) if _row[i].value else '') for i in range(len(_hd))} for _row in _rn[1:]]
        else:
            with open(val,'r',encoding='utf-8') as f:
                rows=list(csv.DictReader(f))
        last={}
        for r in rows:
            last[r.get('Conversation ID','')]=r.get('Customer Message','') or ''
        return last

days_out  = []
convs_out = []
cat_data  = {c: [] for c in ALL_CATS}
ytd       = {c: 0  for c in ALL_CATS}

for day_label, val in FILES:
    rows  = load_file(val)
    total = len(rows)
    days_out.append(day_label)
    convs_out.append(total)

    day_cats = {c: 0 for c in ALL_CATS}
    INVOICE_BUGS = {'No-Show Duration Error','Invoice Date Offset','Wrong Duration',
                    'Notes Cannot Be Added','Duplicate Invoices','Invoice Shows $0',
                    'Phantom/Unconfirmed Invoice','Student Removed Before Invoice',
                    'No-Show Invoice Process','Incentive/Rate Not Applied',
                    'Auto-Invoice Launch Confusion','Bot Stale Instructions','Invoice / Other'}
    for t in rows.values():
        # Classify on user lines only — prevents agent text from bleeding keywords
        ul = [l.replace('*USER*:','').strip() for l in t.split('\n') if l.startswith('*USER*:')]
        user_text = ' '.join(ul) if ul else t
        c = classify(user_text)
        # Fallback: if Other/General and full transcript available, try full text
        # but only for non-invoice categories to prevent agent bleed
        if c == 'Other / General' and len(t) > len(user_text) + 50:
            full_cat = classify(t.lower())
            if full_cat != 'Other / General' and full_cat not in INVOICE_BUGS:
                c = full_cat
        day_cats[c] += 1

    # Sanity check
    assert sum(day_cats.values()) == total, f"{day_label}: {sum(day_cats.values())} != {total}"

    for c in ALL_CATS:
        cat_data[c].append(day_cats[c])
        ytd[c] += day_cats[c]

    sys.stderr.write(f"  {day_label}: {total} convs, labels={sum(day_cats.values())} ✓\n")

result = {
    'days':        days_out,
    'convs':       convs_out,
    'cat_data':    cat_data,
    'ytd':         ytd,
    'total_convs': sum(convs_out),
    'n_days':      len(days_out),
}
json.dump(result, sys.stdout)
