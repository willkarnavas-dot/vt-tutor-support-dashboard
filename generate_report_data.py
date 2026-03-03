import csv, re, json, sys
from numbers_parser import Document


import subprocess, shutil

def push_dashboard_to_github():
    GH_TOKEN = "GITHUB_TOKEN_REMOVED"
    USERNAME = "willkarnavas-dot"
    REPO_URL = f"https://{GH_TOKEN}@github.com/{USERNAME}/vt-tutor-support-dashboard.git"
    REPO_DIR = "/home/claude/vt_dashboard_repo"
    DASHBOARD_SRC = "/mnt/user-data/outputs/invoicing_trends_dashboard.html"

    try:
        # Copy latest dashboard
        shutil.copy(DASHBOARD_SRC, f"{REPO_DIR}/index.html")
        # Commit and push
        subprocess.run(["git", "-C", REPO_DIR, "add", "index.html"], check=True, capture_output=True)
        result = subprocess.run(
            ["git", "-C", REPO_DIR, "commit", "-m", f"Dashboard update - {DATE_LABEL}"],
            capture_output=True, text=True
        )
        if "nothing to commit" in result.stdout:
            print("GitHub: dashboard unchanged, no push needed")
            return
        subprocess.run(
            ["git", "-C", REPO_DIR, "push", "origin", "master"],
            check=True, capture_output=True
        )
        print(f"GitHub: dashboard pushed successfully -> https://bucolic-paprenjak-a99728.netlify.app/")
    except Exception as e:
        print(f"GitHub push failed: {e}")

# ── CONFIG ────────────────────────────────────────────────────────────────────
TODAY_FILE = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/uploads/evaluations_2026-02-27_14_59_25.csv"
DATE_LABEL = sys.argv[2] if len(sys.argv) > 2 else "February 26, 2026"

# Prior 7 days for trending (hardcoded — updated daily by pipeline)
PREV_7_FILES = [
    "/mnt/user-data/uploads/evaluations_2026-02-19_04_57_35.csv",   # Feb 15
    "/mnt/user-data/uploads/evaluations_2026-02-19_04_56_10.csv",   # Feb 16
    "/mnt/user-data/uploads/evaluations_2026-02-19_04_51_29.csv",   # Feb 17
    "/mnt/user-data/uploads/evaluations_2026-02-19_04_44_14.csv",   # Feb 18
    "/mnt/user-data/uploads/evaluations_2026-02-23_16_22_12.csv",   # Feb 20
    "/mnt/user-data/uploads/evaluations_2026-02-23_16_24_05.csv",   # Feb 21
    "/mnt/user-data/uploads/evaluations_2026-02-23_16_24_17.csv",   # Feb 22
]
YESTERDAY_FILE = "/mnt/user-data/uploads/evaluations_2026-02-26_15_15_18.csv"  # Feb 25

email_pat = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')

# ── LOAD TODAY ────────────────────────────────────────────────────────────────
if TODAY_FILE.endswith('.numbers'):
    from numbers_parser import Document as _Doc
    _doc = _Doc(TODAY_FILE)
    _sh = _doc.sheets[0]; _tb = _sh.tables[0]; _rn = list(_tb.iter_rows())
    _hd = [str(c.value) if c.value else '' for c in _rn[0]]
    rows = []
    for _row in _rn[1:]:
        rows.append({_hd[i]: (str(_row[i].value) if _row[i].value else '') for i in range(len(_hd))})
else:
    with open(TODAY_FILE, 'r', encoding='utf-8') as f:
        rows = list(csv.DictReader(f))
today_last = {}
for r in rows: today_last[r.get('Conversation ID','')] = r
total_today = len(today_last)

# ── DISPOSITION ───────────────────────────────────────────────────────────────
status = {'resolved':0,'handed_off':0,'abandoned':0}
for r in today_last.values():
    s = (r.get('Status','') or '').strip().lower()
    if s in status: status[s] += 1

# ── PREV 7 DAY AVG ────────────────────────────────────────────────────────────
prev_totals = {'total':0,'resolved':0,'handed_off':0,'abandoned':0}
for fp in PREV_7_FILES:
    with open(fp,'r',encoding='utf-8') as f:
        prev_rows = list(csv.DictReader(f))
    last = {}
    for r in prev_rows: last[r.get('Conversation ID','')] = r
    prev_totals['total'] += len(last)
    for r in last.values():
        s = (r.get('Status','') or '').strip().lower()
        if s in prev_totals: prev_totals[s] += 1
n_prev = len(PREV_7_FILES)
prev_avg = prev_totals['total'] / n_prev
prev_res_pct = prev_totals['resolved']/prev_totals['total']*100
prev_ho_pct  = prev_totals['handed_off']/prev_totals['total']*100
prev_ab_pct  = prev_totals['abandoned']/prev_totals['total']*100

# ── YESTERDAY ────────────────────────────────────────────────────────────────
with open(YESTERDAY_FILE,'r',encoding='utf-8') as f:
    yrows = list(csv.DictReader(f))
yest_last = {}
for r in yrows: yest_last[r.get('Conversation ID','')] = r

# ── CATEGORY CLASSIFIER ───────────────────────────────────────────────────────
def categorize(t):
    if any(x in t for x in ['incentive','salary increment','rate increase','incremental','not applied','$1.00','$1 increase','per session increase','rate not','underpaid','pay rate','incorrect rate','wrong rate','rate is wrong']):
        return 'Incentive/Rate Not Applied'
    if any(x in t for x in ['activesupport','duration with nil','unable to create invoice']):
        return 'No-Show Duration Error'
    if any(x in t for x in ['no show','no-show']) and any(x in t for x in ['submit','invoic','paid','time and a half','how to','process','bill']):
        return 'No-Show Invoice Process'
    if any(x in t for x in ['session note','notes field','cannot add note','add note','session notes']):
        return 'Notes Cannot Be Added'
    if any(x in t for x in ['wrong date','date is wrong','date off','incorrect date','previous day','day before','date error','date on the invoice']):
        return 'Invoice Date Offset'
    if any(x in t for x in ['wrong duration','incorrect duration','duration is wrong','duration error','wrong length','recorded wrong']) and 'invoic' in t:
        return 'Wrong Duration'
    if any(x in t for x in ['duplicate','double invoice','invoiced twice','two invoice','charged twice']):
        return 'Duplicate Invoices'
    if any(x in t for x in ['$0','$0.00','shows zero','zero dollar','0.00 hours']):
        return 'Invoice Shows $0'
    if any(x in t for x in ['phantom','unconfirmed','automatically invoiced.*not show']):
        return 'Phantom/Unconfirmed Invoice'
    if any(x in t for x in ['student.*remov','remov.*student','student was removed']) and 'invoic' in t:
        return 'Student Removed Before Invoice'
    if any(x in t for x in ['reschedul','cancel session','schedule.*session','change.*session','move.*session','session time','time slot','scheduling','book a session']):
        return 'Scheduling'
    if any(x in t for x in ['when.*paid','when.*payout','disbursement','paypal','direct deposit','paycheck','inova','1099','tax form']):
        return 'Payment/Disbursement'
    if any(x in t for x in ["can't join","cannot join",'link broken','platform not working','kicked out','connectivity','rejoin','frozen','black screen']):
        return 'Session Link / Tech Issues'
    if any(x in t for x in ['opportunity','new student','accept.*opport','express interest']):
        return 'Opportunities / Matching'
    if any(x in t for x in ['profile','photo','bio','personal statement','profile score']):
        return 'Profile / Account'
    if any(x in t for x in ['remove.*student','drop.*student','student.*roster','roster']):
        return 'Student Roster'
    if any(x in t for x in ['vt4s','school session','group session','group class','polk','antelope valley','saginaw','geneva']):
        return 'VT4S / Schools'
    if any(x in t for x in ['substitute','sub session','sub room','subbing']):
        return 'Substitute Sessions'
    if any(x in t for x in ['instant','on-demand','instant session','instant tutoring']):
        return 'Instant Tutoring'
    return 'Other / General'

today_cats = {}
yest_cats  = {}
for r in today_last.values():
    t = (r.get('Customer Message','') or '').lower()
    c = categorize(t)
    today_cats[c] = today_cats.get(c,0) + 1
for r in yest_last.values():
    t = (r.get('Customer Message','') or '').lower()
    c = categorize(t)
    yest_cats[c] = yest_cats.get(c,0) + 1

# ── DISPOSITION BY CATEGORY ───────────────────────────────────────────────────
NAME_SHORT = {
    'Student Removed Before Invoice': 'Student Removed / No Invoice',
    'Phantom/Unconfirmed Invoice':    'Phantom Invoice',
}
cat_disp = {}
for cid, r in today_last.items():
    t   = (r.get('Customer Message','') or '').lower()
    sta = (r.get('Status','') or '').strip().lower()
    cat = NAME_SHORT.get(categorize(t), categorize(t))
    if cat not in cat_disp:
        cat_disp[cat] = {'resolved':0,'handed_off':0,'abandoned':0,'total':0}
    cat_disp[cat]['total'] += 1
    if sta in ('resolved','handed_off','abandoned'):
        cat_disp[cat][sta] += 1

disposition_by_cat = []
for cat, d in sorted(cat_disp.items(), key=lambda x: -x[1]['total']):
    t = d['total']
    if t == 0: continue
    disposition_by_cat.append({
        'category':     cat,
        'total':        t,
        'resolved':     d['resolved'],
        'resolvedPct':  f"{d['resolved']/t*100:.0f}%",
        'handedOff':    d['handed_off'],
        'handedOffPct': f"{d['handed_off']/t*100:.0f}%",
        'abandoned':    d['abandoned'],
        'abandonedPct': f"{d['abandoned']/t*100:.0f}%",
    })

# ── TOP DRIVERS ───────────────────────────────────────────────────────────────
top_drivers = sorted(today_cats.items(), key=lambda x: -x[1])[:10]
top_drivers_out = [{'name':n,'count':c,'pct':f"{c/total_today*100:.0f}%"} for n,c in top_drivers]

# ── TOP 3 CATEGORY EXAMPLES ────────────────────────────────────────────────────
# Pull 5 clean verbatim examples from the top 3 contact categories
top3_cats = [d['name'] for d in top_drivers_out[:3]]
cat_examples = {c: [] for c in top3_cats}

# Noise phrases to skip — too short or clearly misclassified
NOISE = [
    "it won't let me join", "it wont let me join", "cannot join",
    "i'd like to unmatch", "i would like to unmatch",
]

for cid, r in today_last.items():
    t   = r.get('Customer Message','') or ''
    ts  = r.get('Conversation Created At','')
    cat = categorize(t.lower())
    if cat not in cat_examples: continue
    if len(cat_examples[cat]) >= 5: continue

    ul = [l.replace('*USER*:', '').strip() for l in t.split('\n') if l.startswith('*USER*:')]
    if not ul: continue
    best = max(ul[:3], key=len, default='').strip()
    if len(best) < 35: continue
    if any(n in best.lower() for n in NOISE): continue

    time_str = ts[11:16] if len(ts) > 10 else ''
    em_match = email_pat.search(r.get('Customer Message','') or '')
    email_str = em_match.group() if em_match else ''
    if email_str and 'varsitytutors.com' not in email_str:
        cat_examples[cat].append(f"{time_str} — \"{best[:210]}\" ({email_str})")

top3_examples = [
    {'category': cat, 'examples': cat_examples.get(cat, [])}
    for cat in top3_cats
]


# ── TICKET DATA ───────────────────────────────────────────────────────────────
# Hardcoded 14-day and this-week from dashboard (updated daily)
TICKET_DATA_14D = {
    'MD-4':812,'MD-5':118,'MD-6':152,'MD-7':542,'MD-8':156,'MD-9':563,
    'MD-21':69,'MD-22':62,'MD-23':54,'MD-26':76,'MD-27':32,'MD-28':25,'MD-29':18,
}
TICKET_DATA_WK = {
    'MD-4':373,'MD-5':14,'MD-6':51,'MD-7':240,'MD-8':51,'MD-9':273,
    'MD-21':32,'MD-22':34,'MD-23':15,'MD-26':12,'MD-27':8,'MD-28':7,'MD-29':4,
}
TICKET_ISSUES = {
    'MD-4': 'Auto-Invoice Rate/Incentive Not Applied',
    'MD-5': 'Invoice Date Offset (One Day Prior)',
    'MD-6': 'Wrong Duration / Invoice Locked',
    'MD-7': 'Session Notes Unavailable',
    'MD-8': 'Duplicate Invoice Submission',
    'MD-9': 'No-Show Invoice Broken / Bot Misinforms',
    'MD-21': 'HTTP 409/422/400 Errors Creating Sessions',
    'MD-22': 'Cancel Session Option Missing',
    'MD-23': 'Edit/Reschedule Dropdown Non-Functional',
    'MD-26': 'Sessions Disappear from Calendar',
    'MD-27': 'Declined Sessions Stuck on Calendar',
    'MD-28': 'Time Slots Missing from Dropdown',
    'MD-29': 'Recurring Series Cannot Be Created/Edited',
}
TICKET_TODAY = {
    'MD-4':74,'MD-5':5,'MD-6':3,'MD-7':36,'MD-8':8,'MD-9':65,
    'MD-21':2,'MD-22':7,'MD-23':2,'MD-26':1,'MD-27':1,'MD-28':1,'MD-29':0,
}
TICKET_YESTERDAY = {
    'MD-4':80,'MD-5':3,'MD-6':16,'MD-7':73,'MD-8':13,'MD-9':59,
    'MD-21':3,'MD-22':7,'MD-23':2,'MD-26':2,'MD-27':1,'MD-28':1,'MD-29':0,
}

def build_tickets(keys):
    out = []
    for k in keys:
        out.append({
            'ticket': k,
            'issue': TICKET_ISSUES[k],
            'today': TICKET_TODAY.get(k,0),
            'yesterday': TICKET_YESTERDAY.get(k,0),
            'contacts14': TICKET_DATA_14D[k],
            'contactsWk': TICKET_DATA_WK[k],
            'status': 'Ideas',
        })
    return sorted(out, key=lambda x: -x['today'])

invoice_tickets = build_tickets(['MD-4','MD-5','MD-6','MD-7','MD-8','MD-9'])
sched_tickets   = build_tickets(['MD-21','MD-22','MD-23','MD-26','MD-27','MD-28','MD-29'])

# ── VERBATIM QUOTES — auto-pulled from today's transcripts ───────────────────
email_pat = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
TICKET_QUOTE_PATS = {
    'MD-4': [r'incentive', r'salary increment', r'rate increase', r'\$1 increase',
             r'rate not applied', r'wrong rate', r'underpaid', r'rate.*not.*correct'],
    'MD-7': [r'session note', r'notes field', r'cannot add note', r'add note', r'session notes'],
    'MD-9': [r'no.?show.*invoice', r'no.?show.*submit', r'unable to create invoice',
             r'activesupport', r'no.?show.*pay', r'invoice.*no.?show'],
}
quotes = []
for ticket, pats in TICKET_QUOTE_PATS.items():
    for cid, r in today_last.items():
        msg   = r.get('Customer Message','') or ''
        ul    = [l.replace('*USER*:','').strip() for l in msg.split('\n') if l.startswith('*USER*:')]
        if not ul: continue
        utext = ' '.join(ul).lower()
        best  = max(ul[:3], key=len, default='').strip()
        if len(best) < 30: continue
        if any(re.search(p, utext) for p in pats):
            ts  = r.get('Conversation Created At','') or ''
            pid = r.get('Platform ID','') or ''
            em  = email_pat.search(msg)
            em2 = email_pat.search(msg)
            quotes.append({
                'ticket': ticket,
                'text':   best[:210],
                'time':   f"{DATE_LABEL}, {ts[11:16]}",
                'email':  em2.group() if em2 else '',
            })
            break

# ── WATCH LIST ─────────────────────────────────────────────────────────────────
watch_list = sorted([
    {'issue':'Negative Balance / Clawback','contacts14':9,'contactsWk':3,'threshold':'< 20/wk — monitoring'},
    {'issue':'Substitute Session Flow Broken','contacts14':6,'contactsWk':2,'threshold':'< 20/wk — monitoring'},
    {'issue':'Session Auto-Created Incorrectly','contacts14':11,'contactsWk':3,'threshold':'< 20/wk — monitoring'},
    {'issue':'Conflict Error Despite No Conflict','contacts14':2,'contactsWk':1,'threshold':'< 20/wk — monitoring'},
], key=lambda x: -x['contacts14'])

# ── LEADERSHIP BULLETS ─────────────────────────────────────────────────────────
vol_diff = total_today - int(prev_avg)
vol_diff_str = f"+{vol_diff}" if vol_diff>0 else str(vol_diff)
res_diff = round(status['resolved']/total_today*100 - prev_res_pct, 1)
ho_diff  = round(status['handed_off']/total_today*100 - prev_ho_pct, 1)

leadership_bullets = [
    {'text': f"{total_today} total contacts today ({vol_diff_str} vs {int(prev_avg)}/day 7-day avg). Containment at {status['resolved']/total_today*100:.0f}% resolved, {status['handed_off']/total_today*100:.0f}% handed to agents.", 'bold': False},
    {'text': f"MD-4 Incentive/Rate bug remains the top driver at 74 contacts today. Every unresolved MD-4 case creates a manual Disbursements escalation.", 'bold': False},
    {'text': f"MD-9 No-Show Invoice is climbing, now the second-highest bug by 14-day volume (563). Tutors cannot submit no-show invoices without agent assistance.", 'bold': False},
    {'text': f"MD-7 Session Notes dropped to 36 today from 73 yesterday, the sharpest single-day decline this week.", 'bold': False},
    {'text': f"MD-5 Date Offset holding low at 5 contacts, down from the Feb 9-11 peak of 37-47/day. Potential stabilization.", 'bold': False, 'color': '1E8449'},
    {'text': f"No new issues crossed the 20-contact/week filing threshold. 4 items on watch list remain below threshold.", 'bold': False},
]


# Shorten long category names for disposition table
NAME_SHORT = {
    'Student Removed Before Invoice': 'Student Removed / No Invoice',
    'Phantom/Unconfirmed Invoice':    'Phantom Invoice',
    'Incentive/Rate Not Applied':     'Incentive/Rate Not Applied',
    'Session Link / Tech Issues':     'Session Link / Tech Issues',
    'Opportunities / Matching':       'Opportunities / Matching',
    'No-Show Invoice Process':        'No-Show Invoice Process',
    'Rate/Contract Questions':        'Rate / Contract Questions',
    'No-Show Duration Error':         'No-Show Duration Error',
    'Notes Cannot Be Added':          'Notes Cannot Be Added',
}


# ── WEEK LABEL ────────────────────────────────────────────────────────────────
from datetime import datetime as _dt
_parsed = _dt.strptime(DATE_LABEL, '%B %d, %Y')
_month_label = _parsed.strftime('%b %Y')
_days_since = (_parsed - _dt(2026, 1, 4)).days
_wk = 1 if _parsed < _dt(2026, 1, 4) else 2 + _days_since // 7
_week_label = f"Wk{_wk}"

# ── OUTPUT ─────────────────────────────────────────────────────────────────────
data = {
    'date': DATE_LABEL,
    'totalContacts': total_today,
    'totalVsAvg': f"{vol_diff_str} vs {int(prev_avg):.0f}/day (7d avg)",
    'resolvedCount': status['resolved'],
    'resolvedPct': f"{status['resolved']/total_today*100:.0f}",
    'resolvedVsPrev': f"{res_diff:+.1f}% vs 7d avg ({prev_res_pct:.0f}%)",
    'handedOffCount': status['handed_off'],
    'handedOffPct': f"{status['handed_off']/total_today*100:.0f}",
    'handedOffVsPrev': f"{ho_diff:+.1f}% vs 7d avg ({prev_ho_pct:.0f}%)",
    'abandonedCount': status['abandoned'],
    'abandonedPct': f"{status['abandoned']/total_today*100:.0f}",
    'abandonedVsPrev': f"7d avg: {prev_ab_pct:.0f}%",
    'topDrivers': top_drivers_out,
    'invoiceTickets': invoice_tickets,
    'schedTickets': sched_tickets,
    'quotes': quotes,
    'watchList': watch_list,
    'leadershipBullets': leadership_bullets,
    'top3Examples': top3_examples,
    'dispositionByCategory': disposition_by_cat,
    'weekLabel': _week_label,
    'dashboardUrl': 'https://bucolic-paprenjak-a99728.netlify.app/',
    'dataNote': f"Assembled CSV export | {total_today} conversations | Feb 26, 2026 | Will Karnavas / Program Management",
}

with open('/home/claude/report_data.json','w') as f:
    json.dump(data, f, indent=2)
print("Data written")
push_dashboard_to_github()


# ── AUTO-DEPLOY TO GITHUB (triggers Netlify) ──────────────────────────────────
import subprocess, shutil, os

GH_TOKEN  = "GITHUB_TOKEN_REMOVED"
REPO_DIR  = "/home/claude/vt_dashboard_repo"
REPO_URL  = f"https://{GH_TOKEN}@github.com/willkarnavas-dot/vt-tutor-support-dashboard.git"
DASHBOARD = "/mnt/user-data/outputs/invoicing_trends_dashboard.html"

try:
    # Ensure repo dir exists and is initialised
    if not os.path.exists(os.path.join(REPO_DIR, ".git")):
        subprocess.run(["git", "init"], cwd=REPO_DIR, check=True)
        subprocess.run(["git", "config", "user.email", "will.karnavas@varsitytutors.com"], cwd=REPO_DIR)
        subprocess.run(["git", "config", "user.name", "Will Karnavas"], cwd=REPO_DIR)
        subprocess.run(["git", "remote", "add", "origin", REPO_URL], cwd=REPO_DIR)
    else:
        # Update remote URL (token may have changed)
        subprocess.run(["git", "remote", "set-url", "origin", REPO_URL], cwd=REPO_DIR)

    # Copy latest dashboard
    shutil.copy(DASHBOARD, os.path.join(REPO_DIR, "index.html"))

    # Commit and push if changed
    subprocess.run(["git", "config", "user.email", "will.karnavas@varsitytutors.com"], cwd=REPO_DIR)
    subprocess.run(["git", "config", "user.name", "Will Karnavas"], cwd=REPO_DIR)
    subprocess.run(["git", "add", "index.html"], cwd=REPO_DIR, check=True)

    diff = subprocess.run(["git", "diff", "--cached", "--stat"], cwd=REPO_DIR, capture_output=True, text=True)
    if diff.stdout.strip():
        subprocess.run(["git", "commit", "-m", f"Dashboard update - {DATE_LABEL}"], cwd=REPO_DIR, check=True)
        subprocess.run(["git", "push", "origin", "master"], cwd=REPO_DIR, check=True)
        print("Dashboard pushed to GitHub — Netlify deploy triggered")
    else:
        print("Dashboard unchanged — no push needed")
except Exception as e:
    print(f"GitHub push warning: {e}")
