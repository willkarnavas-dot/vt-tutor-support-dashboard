"""
Single-label classifier — strict priority order.
Each conversation gets exactly ONE category.
"""
import re

PRIORITY_ORDER = [
    # ── SPECIFIC INVOICE BUGS (most specific first) ───────────────────────────
    ('No-Show Duration Error',
     lambda t: any(x in t for x in ['activesupport','duration with nil','unable to create invoice. error'])),

    ('Invoice Date Offset',
     lambda t: any(x in t for x in ['wrong date','date is wrong','date off','incorrect date',
                                     'previous day','day before','date on the invoice']) and 'invoic' in t),

    ('Wrong Duration',
     lambda t: any(x in t for x in ['wrong duration','incorrect duration','duration is wrong',
                                     'duration error','recorded wrong']) and 'invoic' in t),

    ('Notes Cannot Be Added',
     lambda t: any(x in t for x in ['session note','notes field','cannot add note',
                                     "can't add note",'add note to invoice',
                                     'session notes not available','session notes not accessible',
                                     'cannot.*edit.*invoice.*note']) and
               any(x in t for x in ['invoic','note','edit'])),

    ('Duplicate Invoices',
     lambda t: any(x in t for x in ['duplicate invoice','double invoice','invoiced twice',
                                     'two invoice','charged twice','billed twice',
                                     'invoice.*appears.*twice','double billed',
                                     'same session.*two invoice'])),

    ('Invoice Shows $0',
     lambda t: ('$0.00' in t or ('$0' in t and 'invoice' in t and 'amount' in t)
                or 'shows zero' in t or '0.00 hours' in t) and 'invoic' in t),

    ('Phantom/Unconfirmed Invoice',
     lambda t: ('phantom' in t or
                ('unconfirmed' in t and 'invoic' in t) or
                ('automatically invoiced' in t and ('not show' in t or 'no show' in t)))),

    ('Student Removed Before Invoice',
     lambda t: any(re.search(p, t) for p in [r'student.*remov', r'remov.*student',
                                              r'student was removed']) and 'invoic' in t),

    ('No-Show Invoice Process',
     lambda t: bool(re.search(r'no.?show', t)) and
               any(x in t for x in ['how.*invoice','invoice.*no.?show','no.?show.*invoice',
                                     'no.?show.*pay','time and a half','bill.*no.?show',
                                     'charge.*no.?show','submit.*no.?show',
                                     'should i.*charge','should i.*invoice',
                                     'what.*do.*no.?show','how.*handle.*no.?show',
                                     'no.?show.*what should','waited.*no.?show',
                                     'student.*not.*show.*invoic','student.*no show.*invoic'])),

    ('Incentive/Rate Not Applied',
     lambda t: any(x in t for x in ['incentive','salary increment','rate increase',
                                     'incremental','$1 increase','per session increase',
                                     'rate not applied','incorrect rate',
                                     'wrong rate','rate is wrong',
                                     '$1 dollar increment','dollar increment',
                                     'rate.*not.*correct','underpaid']) and
               any(x in t for x in ['invoic','pay','rate','earn','session'])),

    ('Auto-Invoice Launch Confusion',
     lambda t: any(x in t for x in ['opt out of automatic invoic','automatically invoiced now',
                                     'new automatic invoice','do not want.*auto.*invoic',
                                     'when did invoic.*change'])),

    ('Bot Stale Instructions',
     lambda t: any(x in t for x in ['edit option.*not integrated','not integrated.*yet',
                                     'bot.*told.*me.*invoice.*manual'])),

    ('Session Link / Tech Issues',
     lambda t: any(x in t for x in ["can't join session","cannot join session",
                                     'link broken','kicked out of session',
                                     'rejoin','frozen','black screen',
                                     'audio.*issue','video.*issue',
                                     'cannot.*connect.*session','session.*not.*loading',
                                     'session.*error','platform not working',
                                     'platform failed','session link.*not work',
                                     'website.*freez','site.*freez','not able.*enter.*session',
                                     'unable.*enter.*session','cannot.*get.*into.*session',
                                     'session.*ended.*early','session.*has ended',
                                     'session.*already ended','kicked.*out',
                                     'session.*not.*open','cannot.*access.*session',
                                     'can.*t.*access.*session','get into.*session',
                                     'get back into session','get back.*class',
                                     'cant open.*session','cannot open.*session',
                                     'error.*404','cookie too large',
                                     'session.*disappeared.*calendar','update failed',
                                     'glitch.*session','session.*glitch',
                                     'video.*not.*work','camera.*not.*work',
                                     'instant video.*not.*work'])),

    # ── BROAD INVOICING CATCH-ALL ─────────────────────────────────────────────
    # Catches all invoicing contacts not matched above
    ('Invoice / Other',
     lambda t: any(x in t for x in ['invoic','invoice']) and
               any(x in t for x in ['session','pay','earn','amount','hour','rate',
                                     'submit','bill','charge','disburs','correct',
                                     'help','issue','problem','error','wrong',
                                     'adjust','fix','change','edit','cannot','can\'t'])),

    # ── OPERATIONAL CATEGORIES ────────────────────────────────────────────────
    ('VT4S / Schools',
     lambda t: any(x in t for x in ['vt4s','school session','group session','group class',
                                     'polk county','antelope valley','saginaw','geneva',
                                     'broward','lastinger','newberry'])),

    ('Substitute Sessions',
     lambda t: any(x in t for x in ['substitute','sub session','sub room',
                                     'subbing','request.*sub','sub.*request'])),

    ('Student Roster',
     lambda t: any(re.search(p, t) for p in [r'remove.*student',r'drop.*student',
                                              r'unmatch.*student',r'student.*unmatch',
                                              r'reassign.*student',r'student.*roster']) or
               ('roster' in t and 'student' in t)),


    ('Instant Tutoring',
     lambda t: any(x in t for x in ['instant tutor','on-demand','instant session',
                                     'instant.*ping'])),

    ('Opportunities / Matching',
     lambda t: any(x in t for x in ['opportunit','express interest',
                                     'new.*assignment','matched.*student',
                                     'accept.*opport']) and
               not any(x in t for x in ['invoic','incentive','rate.*increase'])),

    ('Profile / Account',
     lambda t: any(x in t for x in ['my profile','profile photo','profile score',
                                     'my bio','account setting','profile.*update',
                                     'photo.*upload','background check'])),

    ('Rate/Contract Questions',
     lambda t: any(x in t for x in ['base rate','what is my rate','contract review',
                                     'hourly rate','rate.*question']) and
               not any(x in t for x in ['incentive','wrong rate','rate.*not.*applied',
                                         'rate.*increase','invoic'])),

    ('Payment/Disbursement',
     lambda t: any(x in t for x in ['disbursement','paypal','direct deposit',
                                     'paycheck','inova','1099','tax form',
                                     'when.*paid','when.*payout',
                                     'negative.*balance','negative.*disburse',
                                     'payment.*method','bank.*account',
                                     'tax.*document','tax.*statement','tax.*purpose',
                                     'tax.*2025','tax.*2026','w2','w-2',
                                     'address.*tax','verification.*form',
                                     'employment.*verification','lender.*verif',
                                     'milestone.*bonus.*not.*paid','milestone.*bonus.*paid',
                                     'how.*payment.*sent','how.*get.*paid',
                                     'payment.*sent','where.*payment'])),


    # ── MD-21: HTTP 409/422/400 Session Scheduling Error ─────────────────────
    ('MD-21 Session Error',
     lambda t: any(re.search(p, t) for p in [
         r'request failed.*409', r'error.*409', r'409.*error',
         r'request failed \(409\)', r'error code 409', r'getting.*409',
         r'still.*409', r'409.*scheduling', r'409.*session',
         r'request failed.*422', r'error.*422', r'422.*error',
         r'request failed.*400', r'error.*400.*session',
         r'failed.*409', r'409.*failed',
     ])),

    # ── MD-22: Cancel/Edit/Delete Session Option Missing ──────────────────────
    ('MD-22 Cancel Option Missing',
     lambda t: any(re.search(p, t) for p in [
         r'no cancel option', r'no.*option.*cancel', r'no option.*cancel',
         r'cancel.*option.*not.*there', r'cannot cancel', r"can't cancel",
         r'won.*t let.*cancel', r'will not.*let.*cancel', r'not.*let.*cancel',
         r'unable to cancel', r'it.*not.*cancel', r'system.*not.*cancel',
         r'system.*won.*t.*cancel', r'it will not allow.*cancel',
         r'no edit button', r'no.*edit.*option', r'edit.*not.*available',
         r'edit.*grayed', r'no.*delete.*option', r'cannot.*delete.*session',
         r"can't.*delete.*session", r'no.*remove.*option',
         r'declined.*can.*t.*delete', r'declined.*cannot.*delete',
         r'declined.*won.*t.*clear', r'declined.*stuck.*calendar',
         r'cancel.*series.*no option', r'no option.*cancel.*series',
         r'cant cancel.*series', r'cannot cancel.*series',
         r'not.*allowing.*cancel', r'i cannot.*cancel.*session',
     ]) and
     not any(re.search(p, t) for p in [r'409', r'422', r'400.*error'])),

    ('Scheduling',
     lambda t: any(x in t for x in ['reschedul','cancel session','session time',
                                     'time slot','book a session','schedule a session',
                                     'create.*session','scheduling issue',
                                     'cannot.*schedule','change.*session.*time',
                                     'move.*session.*time','move my session',
                                     'move.*schedule','cancel.*recurring',
                                     'cancel.*sessions','need to cancel',
                                     'cancel.*class','cancel.*appointment',
                                     'change.*appointment','edit.*session',
                                     'editing.*session','add.*session.*schedule',
                                     'schedule.*recurring','recurring.*session',
                                     'adjust.*session','session.*conflict',
                                     'cannot.*add.*session','add a session',
                                     'create a session','scheduling support',
                                     'help.*schedule','change.*time.*session',
                                     'move.*time','change my schedule',
                                     'remove.*session.*schedule']) and
               not any(x in t for x in ['invoic','incentive','rate.*increase',
                                         'pay.*session','earn.*session'])),

    # ── CATCH-ALL ─────────────────────────────────────────────────────────────
    ('Other / General', lambda t: True),
]

ALL_CATS = [cat for cat, _ in PRIORITY_ORDER]

def classify(text):
    t = (text or '').lower()
    for cat, fn in PRIORITY_ORDER:
        try:
            if fn(t):
                return cat
        except Exception:
            pass
    return 'Other / General'
