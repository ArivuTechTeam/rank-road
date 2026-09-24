import datetime as dt
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

# ---------------------------------------------------------------- syllabus
# (subject code, chapter no, chapter name, [(unit name, hours)...]) — ICAI New Scheme CA Foundation
SUBJECTS = {
    "Acc": ("Paper 1 – Accounting", "1F4E3D", "E2F0D9"),
    "Law": ("Paper 2 – Business Laws", "7F3F00", "FCE4D6"),
    "QA":  ("Paper 3 – Quantitative Aptitude", "1F3864", "DDEBF7"),
    "Eco": ("Paper 4 – Business Economics", "5B2C6F", "E4DFEC"),
}

CH = {
"Acc": [
 (1, "Theoretical Framework", [("U1 Meaning & Scope of Accounting", 2), ("U2 Accounting Concepts, Principles & Conventions", 2), ("U3 Capital & Revenue Expenditure/Receipts", 2), ("U4 Contingent Assets & Liabilities", 1), ("U5 Accounting Policies", 1), ("U6 Accounting as a Measurement Discipline", 1), ("U7 Accounting Standards (AS/Ind AS overview)", 2), ("PYQ practice – Ch 1 (theory Qs, T/F with reasons)", 2)]),
 (2, "Accounting Process", [("U1 Basic Accounting Procedures – Journal Entries", 3), ("U2 Ledgers", 2), ("U3 Trial Balance", 2), ("U4 Subsidiary Books", 2), ("U5 Cash Book", 2), ("U6 Rectification of Errors", 4), ("PYQ practice – Ch 2 (journal, rectification, suspense)", 3)]),
 (3, "Bank Reconciliation Statement", [("BRS – concept, all formats (from cash book / pass book / overdraft)", 4), ("PYQ practice – Ch 3", 2)]),
 (4, "Inventories", [("Inventory valuation – FIFO, weighted avg, adjusted selling price, stock on date after year-end", 4), ("PYQ practice – Ch 4", 2)]),
 (5, "Depreciation & Amortisation", [("Depreciation – SLM, WDV, change of method, disposal, asset account/provision account", 4), ("PYQ practice – Ch 5", 2)]),
 (6, "Bills of Exchange & Promissory Notes", [("Bills – drawing, discounting, endorsement, dishonour, renewal, accommodation bills", 4), ("PYQ practice – Ch 6", 2)]),
 (7, "Final Accounts of Sole Proprietors", [("U1 Trading & P&L Account, Balance Sheet – basics", 3), ("U2 Adjustments (closing stock, outstanding/prepaid, bad debts, provisions, depreciation, drawings, etc.)", 3), ("U3 Manufacturing Account + comprehensive problems", 2), ("PYQ practice – Ch 7", 3)]),
 (8, "Financial Statements of Not-for-Profit Organisations", [("NPO – Receipts & Payments, Income & Expenditure, Balance Sheet, subscription workings, capital fund", 4), ("PYQ practice – Ch 8", 2)]),
 (9, "Accounts from Incomplete Records", [("Single entry – statement of affairs method, conversion method, missing figures", 4), ("PYQ practice – Ch 9", 2)]),
 (10, "Partnership & LLP Accounts", [("U1 Introduction to Partnership – P&L Appropriation, capital accounts, guarantee", 2), ("U2 Treatment of Goodwill", 2), ("U3 Admission of a Partner (7-step master sequence)", 3), ("U4 Retirement of a Partner", 3), ("U5 Death of a Partner + LLP basics", 2), ("PYQ practice – Ch 10", 3)]),
 (11, "Company Accounts", [("U1 Introduction to Company Accounts", 1), ("U2 Issue, Forfeiture & Reissue of Shares", 3), ("U3 Bonus Shares & Rights Issue", 2), ("U4 Issue of Debentures", 2), ("U5 Redemption of Debentures", 2), ("U6 Redemption of Preference Shares", 2), ("PYQ practice – Ch 11", 3)]),
],
"Law": [
 (1, "Indian Regulatory Framework", [("Indian regulatory framework – sources of law, court structure, regulators", 2), ("PYQ practice – Ch 1", 1)]),
 (2, "Indian Contract Act, 1872", [("U1 Nature of Contracts (offer, acceptance, classification)", 3), ("U2 Consideration", 2), ("U3 Other Essential Elements (capacity, free consent, legality)", 3), ("U4 Performance of Contract", 3), ("U5 Breach of Contract & Remedies", 2), ("U6 Contingent & Quasi Contracts", 2), ("U7 Contract of Indemnity & Guarantee", 2), ("U8 Bailment & Pledge", 2), ("U9 Agency", 2), ("PYQ practice – Ch 2 (case-study Qs + landmark cases)", 4)]),
 (3, "Sale of Goods Act, 1930", [("U1 Formation of Contract of Sale", 2), ("U2 Conditions & Warranties", 2), ("U3 Transfer of Ownership & Delivery of Goods", 2), ("U4 Unpaid Seller", 2), ("PYQ practice – Ch 3", 2)]),
 (4, "Indian Partnership Act, 1932", [("U1 General Nature of Partnership", 2), ("U2 Relations of Partners", 3), ("U3 Registration & Dissolution of a Firm", 2), ("PYQ practice – Ch 4", 2)]),
 (5, "Limited Liability Partnership Act, 2008", [("LLP Act – nature, incorporation, partners, comparison with company/partnership", 3), ("PYQ practice – Ch 5", 1)]),
 (6, "Companies Act, 2013", [("U1 Meaning, Features & Corporate Veil", 2), ("U2 Classes of Companies", 2), ("U3 Incorporation, MOA, AOA; doctrines of ultra vires, constructive notice, indoor management", 3), ("PYQ practice – Ch 6", 2)]),
 (7, "Negotiable Instruments Act, 1881", [("U1 Meaning, Characteristics & Types of NI", 2), ("U2 Negotiation, Endorsement, Presentment", 2), ("U3 Dishonour, Crossing of Cheques, Sec 138 bouncing", 2), ("PYQ practice – Ch 7", 2)]),
],
"QA": [
 (1, "Ratio & Proportion, Indices, Logarithms", [("U1 Ratio & Proportion", 2), ("U2 Indices", 1.5), ("U3 Logarithms", 1.5), ("PYQ practice – Ch 1 (MCQ speed drill)", 2)]),
 (2, "Equations", [("U1 Linear & Simultaneous Equations", 2), ("U2 Quadratic & Cubic Equations", 2), ("PYQ practice – Ch 2", 1.5)]),
 (3, "Linear Inequalities", [("Linear inequalities & graphical solution", 2.5), ("PYQ practice – Ch 3", 1)]),
 (4, "Mathematics of Finance", [("U1 Simple & Compound Interest, effective rate", 2), ("U2 Annuities (present value, future value, sinking fund)", 3), ("U3 Leasing, NPV, CAGR, perpetuity, depreciation", 2), ("PYQ practice – Ch 4", 3)]),
 (5, "Permutations & Combinations", [("U1 Permutations", 2), ("U2 Combinations", 2), ("PYQ practice – Ch 5", 2)]),
 (6, "Sequence & Series (AP & GP)", [("AP, GP, sum of series, applications", 3), ("PYQ practice – Ch 6", 1.5)]),
 (7, "Sets, Relations, Functions; Limits & Continuity", [("U1 Sets & Relations", 2), ("U2 Functions", 1.5), ("U3 Basics of Limits & Continuity", 2), ("PYQ practice – Ch 7", 1.5)]),
 (8, "Differential & Integral Calculus", [("U1 Differentiation & applications (maxima/minima)", 3), ("U2 Integration & applications", 3), ("PYQ practice – Ch 8", 2)]),
 (9, "Number Series, Coding-Decoding, Odd Man Out", [("Number series, coding-decoding, odd man out", 2.5), ("PYQ practice – Ch 9", 1)]),
 (10, "Direction Tests", [("Direction tests", 1.5), ("PYQ practice – Ch 10", 1)]),
 (11, "Seating Arrangements", [("Linear & circular seating arrangements", 2.5), ("PYQ practice – Ch 11", 1)]),
 (12, "Blood Relations", [("Blood relations", 1.5), ("PYQ practice – Ch 12", 1)]),
 (13, "Statistical Description of Data & Sampling", [("U1 Collection, classification & presentation of data (tables, diagrams, frequency distribution)", 2), ("U2 Sampling – methods, sampling distribution", 2), ("PYQ practice – Ch 13", 1.5)]),
 (14, "Measures of Central Tendency & Dispersion", [("U1 Mean, Median, Mode, GM, HM, partition values", 3), ("U2 Range, MD, SD, QD, coefficient of variation", 3), ("PYQ practice – Ch 14", 2)]),
 (15, "Probability", [("Probability – classical, addition/multiplication theorems, conditional, expectation", 4), ("PYQ practice – Ch 15", 2)]),
 (16, "Theoretical Distributions", [("U1 Binomial & Poisson distributions", 2), ("U2 Normal distribution", 2), ("PYQ practice – Ch 16", 2)]),
 (17, "Correlation & Regression", [("U1 Correlation (Karl Pearson, Spearman, concurrent deviation)", 2), ("U2 Regression lines & coefficients", 2), ("PYQ practice – Ch 17", 2)]),
 (18, "Index Numbers", [("Index numbers – Laspeyres, Paasche, Fisher, tests, deflating, chain index", 3), ("PYQ practice – Ch 18", 1.5)]),
],
"Eco": [
 (1, "Nature & Scope of Business Economics", [("U1 Introduction to Business Economics", 2), ("U2 Basic Problems of an Economy & Role of Price Mechanism", 2), ("PYQ practice – Ch 1", 1.5)]),
 (2, "Theory of Demand & Supply", [("U1 Law of Demand & Elasticity of Demand", 3), ("U2 Demand Forecasting", 1.5), ("U3 Consumer Behaviour – utility analysis, indifference curves", 3), ("U4 Theory of Supply & elasticity of supply", 2), ("PYQ practice – Ch 2", 3)]),
 (3, "Theory of Production & Cost", [("U1 Theory of Production (laws of returns, isoquants)", 3), ("U2 Theory of Cost (short-run & long-run cost curves)", 3), ("PYQ practice – Ch 3", 2)]),
 (4, "Price Determination in Different Markets", [("U1 Meaning & Types of Markets", 2), ("U2 Determination of Prices (demand-supply equilibrium)", 1.5), ("U3 Perfect Competition & Monopoly", 3), ("U4 Monopolistic Competition & Oligopoly", 2), ("PYQ practice – Ch 4", 2.5)]),
 (5, "Business Cycles", [("Business cycles – phases, features, causes, indicators", 3), ("PYQ practice – Ch 5", 1.5)]),
 (6, "Determination of National Income", [("U1 National income concepts & measurement (GDP, GNP, NNP, methods)", 3), ("U2 Keynesian theory of income determination, multiplier", 3), ("PYQ practice – Ch 6", 2)]),
 (7, "Public Finance", [("U1 Fiscal Functions & Market Failure", 2), ("U2 Government Interventions to correct market failure", 1.5), ("U3 Fiscal Policy", 2), ("PYQ practice – Ch 7", 1.5)]),
 (8, "Money Market", [("U1 Concept of Money & Demand for Money", 2), ("U2 Supply of Money & Monetary Policy (RBI tools)", 3), ("PYQ practice – Ch 8", 1.5)]),
 (9, "International Trade", [("U1 Theories of International Trade", 2), ("U2 Trade Policy – instruments, tariffs, quotas, WTO", 2), ("U3 Exchange Rates, BOP, FDI", 2), ("PYQ practice – Ch 9", 2)]),
 (10, "Indian Economy", [("Indian economy – overview, sectors, reforms, current data", 3), ("PYQ practice – Ch 10", 1.5)]),
],
}

# ---------------------------------------------------------------- calendar
START = dt.date(2026, 10, 1)
P1_END = dt.date(2026, 11, 21)          # Study phase (SM + PYQ)
BUFFER_DAY = dt.date(2026, 11, 22)
MOCK1 = [dt.date(2026, 11, 23) + dt.timedelta(d) for d in range(4)]   # Nov 23 – 26
R1_START, R1_END = dt.date(2026, 11, 27), dt.date(2026, 12, 11)
MOCK2 = [dt.date(2026, 12, 12) + dt.timedelta(d) for d in range(4)]   # Dec 12 – 15
R2_START, R2_END = dt.date(2026, 12, 16), dt.date(2026, 12, 27)
MOCK3 = [dt.date(2026, 12, 28) + dt.timedelta(d) for d in range(4)]   # Dec 28 – 31
FINAL_START = dt.date(2027, 1, 1)
# Confirmed: ICAI announcement No. 13-CA(EXAM)/JANUARY/2027 dated 23-Sep-2026
EXAMS = {dt.date(2027, 1, 3): "Acc", dt.date(2027, 1, 5): "Law", dt.date(2027, 1, 7): "QA", dt.date(2027, 1, 9): "Eco"}
END = dt.date(2027, 1, 9)

def planned_hours(d):
    sun = d.weekday() == 6
    if d.month == 10:
        return 3.5 if sun else 7
    if d.month == 11:
        return 4.5 if sun else 9
    return 6 if sun else 12

def recap_hours(d):
    """Morning recap of the previous day's topics (active recall). Exam days: none (covered by the morning skim)."""
    if d in EXAMS:
        return 0
    if d.weekday() == 6:
        return 0.5
    return 0.5 if d.month == 10 else 0.75

CHNAME = {f"{s} Ch {no}:": f"{s} Ch {no} {name}" for s, chs in CH.items() for no, name, _ in chs}

def phase_of(d):
    if d <= P1_END:
        return "Phase 1 – Study (SM + PYQ)"
    if d == BUFFER_DAY:
        return "Buffer / Prep for Mock 1"
    if d in MOCK1:
        return "Mock Exam 1"
    if R1_START <= d <= R1_END:
        return "Revision 1"
    if d in MOCK2:
        return "Mock Exam 2"
    if R2_START <= d <= R2_END:
        return "Revision 2"
    if d in MOCK3:
        return "Mock Exam 3"
    if d in EXAMS:
        return "EXAM DAY"
    return "Final Revision"

# ---------------------------------------------------------------- allocation engine
PAIRS = [("Acc", "Law"), ("QA", "Eco"), ("Acc", "Eco"), ("QA", "Law")]

WEIGHT = {"Acc": 1.2, "Law": 1.35, "QA": 0.9, "Eco": 0.9}   # student priority: more time on Accounts & Law

# Exam importance per chapter: (approx marks in a 100-mark paper, priority). ArivuPro estimate from ICAI
# section-wise weightage and the June 2024 – Jan 2026 papers. Priority scales study time: High x1.25, Medium x1.0, Low x0.75.
IMPORTANCE = {
 ("Acc",1):("8–12 (theory + T/F)","Medium"), ("Acc",2):("10–15","High"), ("Acc",3):("5–10","Medium"), ("Acc",4):("5–10","Medium"),
 ("Acc",5):("5–10","Medium"), ("Acc",6):("5–10","Medium"), ("Acc",7):("10–15","High"), ("Acc",8):("10","High"),
 ("Acc",9):("10","Medium"), ("Acc",10):("10–15","High"), ("Acc",11):("10–15","High"),
 ("Law",1):("3–5","Low"), ("Law",2):("25–30","High"), ("Law",3):("10–15","Medium"), ("Law",4):("10–15","Medium"),
 ("Law",5):("5–8","Low"), ("Law",6):("15–20","High"), ("Law",7):("8–12","Medium"),
 ("QA",1):("4–6","Medium"), ("QA",2):("3–5","Medium"), ("QA",3):("2–4","Low"), ("QA",4):("8–10","High"), ("QA",5):("4–6","Medium"),
 ("QA",6):("3–5","Medium"), ("QA",7):("4–6","Medium"), ("QA",8):("5–7","Medium"), ("QA",9):("4–6 (LR – easy marks)","Medium"),
 ("QA",10):("4–5 (LR – easy marks)","Medium"), ("QA",11):("5–6 (LR – easy marks)","Medium"), ("QA",12):("4–5 (LR – easy marks)","Medium"),
 ("QA",13):("6–8","Medium"), ("QA",14):("8–10","High"), ("QA",15):("6–8","High"), ("QA",16):("6–8","Medium"), ("QA",17):("6–8","High"), ("QA",18):("4–6","Medium"),
 ("Eco",1):("5–8","Low"), ("Eco",2):("15–20","High"), ("Eco",3):("10–12","High"), ("Eco",4):("15–18","High"), ("Eco",5):("5–8","Low"),
 ("Eco",6):("8–10","Medium"), ("Eco",7):("8–10","Medium"), ("Eco",8):("8–10","Medium"), ("Eco",9):("8–10","Medium"), ("Eco",10):("5–8","Low"),
}
PRIO_FACTOR = {"High": 1.25, "Medium": 1.0, "Low": 0.75}
def wfac(s, no):
    return WEIGHT[s] * PRIO_FACTOR[IMPORTANCE[(s, no)][1]]

def build_units(phase):
    """Return {subj: [ {label, hours, chapter_key} ]}"""
    out = {}
    for s, chs in CH.items():
        lst = []
        order = chs if phase != "R2" else sorted(chs, key=lambda c: {"High": 0, "Medium": 1, "Low": 2}[IMPORTANCE[(s, c[0])][1]])
        for no, name, units in order:
            tot = sum(h for _, h in units) * wfac(s, no)
            pr = IMPORTANCE[(s, no)][1].upper()
            key = f"{s} Ch {no}:"
            name = f"{name} [{pr}]"
            if phase == 1:
                for u, h in units:
                    lst.append({"label": f"{key} {name} — {u}", "hours": h * wfac(s, no), "key": key, "ch": (s, no)})
            elif phase == "R1":
                h = round(tot * 0.42 * 2) / 2
                lst.append({"label": f"{key} {name} — REVISION 1: re-read concept notes, re-solve ICAI illustrations/TYK, PYQs (last 6 attempts), RTP", "hours": max(h, 1), "key": key, "ch": (s, no)})
            else:
                h = round(tot * 0.30 * 2) / 2
                lst.append({"label": f"{key} {name} — REVISION 2: one-page summary/formula sheet, MTP questions, timed MCQ drill, weak-area fix", "hours": max(h, 1), "key": key, "ch": (s, no)})
        out[s] = lst
    return out

def daterange(a, b):
    d = a
    while d <= b:
        yield d
        d += dt.timedelta(1)

def q(x):
    return round(x * 4) / 4

BLOCKS = [("Acc", "QA"), ("Law", "Eco")]   # study two subjects at a time; finish both, then move to the next pair

def allocate_block(units, days, a, b):
    """Two subjects a,b studied every day of `days`; hours split so both finish together."""
    total_avail = sum(h for _, h in days)
    total_need = sum(u["hours"] for s in (a, b) for u in units[s])
    scale = total_avail * 0.965 / total_need
    for s in (a, b):
        for u in units[s]:
            u["rem"] = max(0.25, q(u["hours"] * scale)); u["full"] = u["rem"]
    ptr = {a: 0, b: 0}
    remaining = {s: sum(u["rem"] for u in units[s]) for s in (a, b)}
    n = len(days)
    plan, last_day_of = {}, {}
    for i, (d, hrs) in enumerate(days):
        left_days = n - i
        fa, fb = remaining[a] / left_days, remaining[b] / left_days
        if fa + fb <= 0:
            plan[d] = [(a, hrs / 2, f"{a}: buffer – Error-Log review / MCQ drill / pending topics [{hrs/2:g}h]", []),
                       (b, hrs / 2, f"{b}: buffer – Error-Log review / MCQ drill / pending topics [{hrs/2:g}h]", [])]
            continue
        split_a = q(hrs * fa / (fa + fb))
        split_a = min(max(split_a, 0.5), hrs - 0.5)
        sessions = []
        for s, h in ((a, split_a), (b, hrs - split_a)):
            texts, keys, left = [], [], h
            while left > 0.01 and ptr[s] < len(units[s]):
                u = units[s][ptr[s]]
                take = min(u["rem"], left)
                u["rem"] -= take; left -= take; remaining[s] -= take
                tag = ""
                if u["rem"] > 0.01:
                    tag = " (start)" if abs(take + u["rem"] - u["full"]) < 0.01 else " (contd.)"
                elif abs(take - u["full"]) > 0.01:
                    tag = " (finish)"
                texts.append(f"{u['label']}{tag} [{take:g}h]")
                keys.append(u["key"])
                if u["rem"] <= 0.01:
                    ptr[s] += 1
            if left > 0.01:
                texts.append(f"{s}: buffer – Error-Log review / MCQ drill / pending topics [{left:g}h]")
            sessions.append([s, h, texts, keys])
            last_day_of[s] = d
        plan[d] = sessions
    for s in (a, b):   # carry-over safety net
        while ptr[s] < len(units[s]):
            u = units[s][ptr[s]]
            for sess in plan[last_day_of[s]]:
                if sess[0] == s:
                    sess[2].append(f"{u['label']} (carry-over – use buffer time) [{u['rem']:g}h]"); sess[3].append(u["key"])
            ptr[s] += 1
    return {d: [(s, h, "\n".join(t), k) if isinstance(t, list) else (s, h, t, k) for s, h, t, k in ss] for d, ss in plan.items()}

def allocate(units, days, pair_offset=0):
    """Split the days between the two subject-pairs in proportion to their workload, then fill each block."""
    need = [sum(u["hours"] for s in blk for u in units[s]) for blk in BLOCKS]
    total_avail = sum(h for _, h in days)
    target1 = total_avail * need[0] / sum(need)
    cum, cut = 0, 0
    for i, (d, h) in enumerate(days):
        cum += h
        if cum >= target1:
            cut = i + 1 if (cum - target1) <= target1 - (cum - h) else i
            break
    cut = max(1, min(cut, len(days) - 1))
    plan = {}
    plan.update(allocate_block(units, days[:cut], *BLOCKS[0]))
    plan.update(allocate_block(units, days[cut:], *BLOCKS[1]))
    return plan

# Phase 1 days (non-Sunday)
p1_days = [(d, planned_hours(d) - recap_hours(d)) for d in daterange(START, P1_END) if d.weekday() != 6]
plan1 = allocate(build_units(1), p1_days)
r1_days = [(d, planned_hours(d) - recap_hours(d)) for d in daterange(R1_START, R1_END) if d.weekday() != 6]
plan_r1 = allocate(build_units("R1"), r1_days, pair_offset=1)
r2_days = [(d, planned_hours(d) - recap_hours(d)) for d in daterange(R2_START, R2_END) if d.weekday() != 6]
plan_r2 = allocate(build_units("R2"), r2_days, pair_offset=2)

# chapters covered per week (for Sunday tests)
def week_chapters(plan, sunday):
    seen = []
    for k in range(6, 0, -1):
        d = sunday - dt.timedelta(k)
        for s, h, txt, keys in plan.get(d, []):
            for key in keys:
                if key not in seen:
                    seen.append(key)
    return seen

MOCK_ORDER = ["Acc", "Law", "QA", "Eco"]
PAPER_TIME = {"Acc": "3h descriptive – 100 marks; 1:45 pm reading time, paper 2:00–5:00 pm", "Law": "3h descriptive – 100 marks; 1:45 pm reading time, paper 2:00–5:00 pm", "QA": "2h MCQ – 100 marks (¼ negative); 2:00–4:00 pm, NO reading time", "Eco": "2h MCQ – 100 marks (¼ negative); 2:00–4:00 pm, NO reading time"}

rows = []  # dicts for Daily Plan
for d in daterange(START, END):
    hrs = planned_hours(d)
    ph = phase_of(d)
    sun = d.weekday() == 6
    s1 = s2 = ("", 0, "")
    extra = ""
    day_keys = []
    rh = recap_hours(d)
    if d in plan1 or d in plan_r1 or d in plan_r2:
        plan = plan1 if d in plan1 else (plan_r1 if d in plan_r1 else plan_r2)
        ss = plan[d]
        SR = "\n↳ Last 15 min = SESSION RECAP: close the book, write this session's key points / formats / formulas / sections on the chapter's one-page summary sheet + 5 flashcards"
        if len(ss) >= 1:
            s1 = (ss[0][0], ss[0][1], ss[0][2] + SR); day_keys += ss[0][3]
        if len(ss) >= 2:
            s2 = (ss[1][0], ss[1][1], ss[1][2] + SR); day_keys += ss[1][3]
    elif sun and (d <= P1_END or R1_START <= d <= R1_END or R2_START <= d <= R2_END):
        plan = plan1 if d <= P1_END else (plan_r1 if d <= R1_END else plan_r2)
        chs = week_chapters(plan, d)
        ch_txt = ", ".join(chs).replace(":", "") if chs else "chapters of this week"
        t_h = round((hrs - rh) * 0.5 * 2) / 2
        if d <= P1_END:
            s1 = ("Test", t_h, f"WEEKLY TEST (timed, closed book): {ch_txt}\nMix of ICAI-pattern descriptive Qs (Acc/Law) and 40 MCQs (QA/Eco) from this week's chapters")
            s2 = ("Review", hrs - rh - t_h, "Self-evaluate test against ICAI suggested answers → note errors in Error Log → clear pending/backlog topics of the week\nWEEKLY RECAP (last 45 min): read the one-page summary sheet of EVERY chapter covered Mon–Sat, re-solve every Error Log item of the week, then 15 min plan for next week")
        else:
            s1 = ("Test", t_h, f"WEEKLY REVISION TEST (timed): {ch_txt}\nUse ICAI RTP / MTP questions of these chapters")
            s2 = ("Review", hrs - rh - t_h, "Evaluate → update Error Log & one-page summary sheets → close pending revision items\nWEEKLY RECAP (last 45 min): flip through the summary sheets of every chapter revised Mon–Sat; re-solve this week's Error Log items; 15 min plan for next week")
    elif d == BUFFER_DAY:
        s1 = ("All", 2, "Prepare: formula sheets (QA), key-case list (Law), diagram list (Eco), format sheet (Acc) — everything you need for open-eye recall")
        s2 = ("Rest", 2, "Light: skim Error Log; sleep early. Mock 1 starts tomorrow in the real exam time-slot (2:00 pm)")
    elif ph.startswith("Mock"):
        n = 1 if d in MOCK1 else (2 if d in MOCK2 else 3)
        idx = (MOCK1 if n == 1 else MOCK2 if n == 2 else MOCK3).index(d)
        s = MOCK_ORDER[idx]
        src = {1: "ICAI MTP Series 1 (or a full-syllabus test paper)", 2: "ICAI MTP Series 2", 3: "Latest ICAI question paper / RTP-based full test"}[n]
        mock_h = 3 if s in ("Acc", "Law") else 2
        s1 = (s, mock_h + 2, f"MOCK {n} – {SUBJECTS[s][0]} ({PAPER_TIME[s]})\nSit at the real exam time, exam conditions, OMR/answer-sheet style. Source: {src}\nThen 2h: mark strictly vs ICAI answers, fill Mock Tracker, list every lost mark in Error Log")
        nxt = MOCK_ORDER[(idx + 1) % 4]
        s2 = (nxt if idx < 3 else "All", hrs - rh - mock_h - 2, (f"Morning: quick revision of {SUBJECTS[nxt][0]} for tomorrow's mock (summary sheets + Error Log)" if idx < 3 else f"Fix the weakest 3 topics found across Mock {n} (all 4 papers) – re-solve, re-read") )
    elif ph == "Final Revision":
        m = {dt.date(2027, 1, 1): (("Law", 6, "FINAL REVISION – Business Laws: all chapters from one-page summaries; key sections + landmark cases; 2 case-study Qs per chapter"), ("Acc", 6, "FINAL REVISION – Accounting: formats (Final A/cs, NPO, Partnership, Company); re-solve 1 PYQ per chapter Ch 7–11; Error Log")),
             dt.date(2027, 1, 2): (("Acc", 9, "DAY BEFORE PAPER 1 – Accounting only: theory (Ch 1), journal & rectification, BRS, Depreciation, Bills; then Ch 7–11 formats + working notes discipline"), ("Rest", 3, "Stop by 8 pm. Pack hall ticket, calculator, pens. Sleep 8h")),
             dt.date(2027, 1, 4): (("Law", 9, "DAY BEFORE PAPER 2 – Business Laws only: Contract Act (biggest weight), Companies Act, Sale of Goods, Partnership, LLP, NI Act — summaries + case list"), ("Rest", 3, "Stop by 8 pm")),
             dt.date(2027, 1, 6): (("QA", 9, "DAY BEFORE PAPER 3 – QA only: formula sheet all 18 chapters; 1 timed 100-MCQ paper (2h, no reading time); Statistics formulas last"), ("Rest", 3, "Stop by 8 pm")),
             dt.date(2027, 1, 8): (("Eco", 9, "DAY BEFORE PAPER 4 – Economics only: all diagrams, definitions, elasticity/multiplier numericals, Indian economy current facts; 100 MCQs timed"), ("Rest", 3, "Stop by 8 pm"))}
        s1, s2 = m[d]
        s2 = (s2[0], s2[1] - rh, s2[2])
        hrs = s1[1] + s2[1] + rh
    elif ph == "EXAM DAY":
        s = EXAMS[d]
        s1 = (s, 3, f"EXAM – {SUBJECTS[s][0]} ({PAPER_TIME[s]}). Morning: 1.5h light skim of summary sheet only; reach centre by 1:00 pm")
        s2 = ("Rest", 0, "No study after the paper. Do not discuss answers. Light revision of next paper in the evening only if fresh (max 1h)")
        hrs = 3
    rows.append({"date": d, "phase": ph, "hours": hrs, "s1": s1, "s2": s2, "keys": day_keys, "recap_h": rh})

# morning recap text = previous day's topics
for i, row_ in enumerate(rows):
    d = row_["date"]
    if row_["recap_h"] == 0:
        row_["recap"] = "— (exam day: morning skim only)"
        continue
    if i == 0:
        row_["recap"] = "First day – no recap. Use 30 min to set up: Error Log notebook, one-page summary-sheet folder (one page per chapter), flashcard box"
        continue
    prev = rows[i - 1]
    pk = []
    for k in prev["keys"]:
        if k not in pk:
            pk.append(k)
    if pk:
        row_["recap"] = "MORNING RECAP (active recall, closed book) of yesterday: " + "; ".join(CHNAME[k] for k in pk) + " → write the key formats/formulas/sections from memory, then check against the summary sheet; re-do 2 flashcards you got wrong"
    elif prev["phase"].startswith("Mock"):
        row_["recap"] = "MORNING RECAP: re-read yesterday's mock Error Log entries; 15 min formula/format sheet of TODAY's paper"
    elif prev["s1"][0] == "Test":
        row_["recap"] = "MORNING RECAP: Sunday test corrections + Error Log items of last week (closed book, then check)"
    elif prev["phase"] == "EXAM DAY":
        row_["recap"] = "MORNING RECAP: none from yesterday's exam – do not discuss the paper; 30 min summary sheets of today's subject"
    else:
        row_["recap"] = "MORNING RECAP: yesterday's summary sheets + Error Log (closed book, then check)"

# ---------------------------------------------------------------- workbook
wb = Workbook()
ARIAL = "Arial"
def F(bold=False, size=10, color="000000", italic=False):
    return Font(name=ARIAL, bold=bold, size=size, color=color, italic=italic)
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HDR_FILL = PatternFill("solid", fgColor="1F4E3D")
SUB_FILL = PatternFill("solid", fgColor="D9EAD3")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def hdr(ws, row, headers, fill=HDR_FILL, color="FFFFFF"):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = F(True, 10, color); cell.fill = fill; cell.alignment = CENTER; cell.border = BORDER

# ===== Sheet 1: Dashboard =====
ws = wb.active; ws.title = "Dashboard"
ws["A1"] = "CA Foundation – January 2027 Attempt | Rank-Oriented Study Timetable (Oct 2026 – Jan 2027)"
ws["A1"].font = F(True, 14, "1F4E3D")
ws["A2"] = "ArivuPro Academy | ICAI New Scheme (4 papers) | Prepared 23-Sep-2026 | Exam dates as per ICAI Announcement No. 13-CA(EXAM)/JANUARY/2027 dated 23-Sep-2026"
ws["A2"].font = F(False, 10, "595959", True)

ws["A4"] = "HOW TO USE"; ws["A4"].font = F(True, 11)
guide = [
 "1. 'Daily Timetable' is the master sheet: every date from 1-Oct-2026 to 9-Jan-2027 shows exactly 2 subjects, the chapter/unit to study and the hours for each. Subjects are taken TWO AT A TIME: Accounting + Quantitative Aptitude are completed first, then Business Laws + Business Economics — in Phase 1 and again inside each revision. Accounting and Law carry extra weight (≈20% and ≈35% more time than a neutral split), and within each paper HIGH-priority chapters (by exam marks) get 25% more time, LOW-priority 25% less — see 'Chapter Master'. Every task in the timetable is tagged [HIGH]/[MEDIUM]/[LOW].",
 "2. Each evening mark the 'Done?' column (Y/N) and enter 'Actual Hrs'. The Dashboard and Chapter Master update automatically.",
 "3. Sundays = weekly test (timed) on that week's chapters + evaluation + backlog clearing + WEEKLY RECAP, at reduced hours. Never skip the test – it is what converts study into marks.",
 "3a. Recap system built into every day: MORNING RECAP (first 30–45 min, closed-book recall of yesterday's topics – column G lists them), SESSION RECAP (last 15 min of each session – summary sheet + flashcards), WEEKLY RECAP (Sunday – all summary sheets + Error Log of the week). Roughly 12% of study time goes to recap; this is what makes Revision 1 fast.",
 "4. Three full mock cycles (4 papers each) are fixed: Mock 1 (23–26 Nov), Mock 2 (12–15 Dec), Mock 3 (28–31 Dec), each paper written at the real exam time. Enter scores in 'Mock & Test Tracker'.",
 "5. Two full revisions are built in: Revision 1 (27 Nov–11 Dec, ~42% of study time per chapter) and Revision 2 (16–27 Dec, ~30%). Final chapter-wise skim happens on 1–2 Jan and on the day before each paper.",
 "6. 'Rank Strategy' sheet gives the marks target, paper-wise time allocation and presentation rules that separate a rank from a pass.",
 "7. Yellow cells are for the student to fill in. Everything else is calculated – do not overwrite formulas.",
]
for i, g in enumerate(guide):
    c = ws.cell(row=5 + i, column=1, value=g); c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=5 + i, start_column=1, end_row=5 + i, end_column=8)
    ws.row_dimensions[5 + i].height = 30

r = 13
ws.cell(row=r, column=1, value="KEY INPUTS (exam dates CONFIRMED by ICAI)").font = F(True, 11)
r += 1
hdr(ws, r, ["Item", "Value", "Note"])
assump = [
 ("Exam date – Paper 1 Accounting", dt.date(2027, 1, 3), "CONFIRMED – ICAI Announcement No. 13-CA(EXAM)/JANUARY/2027 dated 23-Sep-2026. 2:00–5:00 pm (3 h), reading time 1:45–2:00 pm"),
 ("Exam date – Paper 2 Business Laws", dt.date(2027, 1, 5), "CONFIRMED – 2:00–5:00 pm (3 h), reading time 1:45–2:00 pm"),
 ("Exam date – Paper 3 Quantitative Aptitude", dt.date(2027, 1, 7), "CONFIRMED – 2:00–4:00 pm (2 h), NO advance reading time"),
 ("Exam date – Paper 4 Business Economics", dt.date(2027, 1, 9), "CONFIRMED – 2:00–4:00 pm (2 h), NO advance reading time"),
 ("Study hours – October (Mon–Sat / Sun)", "7 / 3.5", "Raised from 6 h because the exam is 11 days earlier than first planned; Sunday = test + review"),
 ("Study hours – November (Mon–Sat / Sun)", "9 / 4.5", "Raised from 8 h"),
 ("Study hours – December (Mon–Sat / Sun)", "12 / 6", "Top of the student's 10–12 h range"),
 ("Study hours – 1–2 January", "12", "Final revision; exam days = morning skim only"),
 ("Starting point on 1 Oct", "Classes completed; self-study of ICAI Study Material from Chapter 1", "Per student"),
 ("Source material", "ICAI Study Material (SM), ICAI Past Year Questions (PYQ, last 6 attempts), RTP, MTP", "All available on icai.org / BoS Knowledge Portal"),
]
for i, (a, b, c) in enumerate(assump):
    rr = r + 1 + i
    ws.cell(row=rr, column=1, value=a).font = F()
    cb = ws.cell(row=rr, column=2, value=b); cb.font = F(color="0000FF"); cb.fill = INPUT_FILL
    if isinstance(b, dt.date):
        cb.number_format = "DD-MMM-YYYY"
    ws.cell(row=rr, column=3, value=c).font = F(italic=True, color="595959")
    for cc in range(1, 4):
        ws.cell(row=rr, column=cc).border = BORDER; ws.cell(row=rr, column=cc).alignment = WRAP
r = r + 1 + len(assump) + 1

ws.cell(row=r, column=1, value="PHASE SUMMARY (calculated from Daily Timetable)").font = F(True, 11)
r += 1
hdr(ws, r, ["Phase", "From", "To", "Days", "Planned Hours", "Actual Hours", "Days Done", "% Days Done"])
phases = [
 ("Phase 1 – Study (SM + PYQ)", START, P1_END),
 ("Buffer / Prep for Mock 1", BUFFER_DAY, BUFFER_DAY),
 ("Mock Exam 1", MOCK1[0], MOCK1[-1]),
 ("Revision 1", R1_START, R1_END),
 ("Mock Exam 2", MOCK2[0], MOCK2[-1]),
 ("Revision 2", R2_START, R2_END),
 ("Mock Exam 3", MOCK3[0], MOCK3[-1]),
 ("Final Revision", FINAL_START, dt.date(2027, 1, 8)),
 ("EXAM DAY", dt.date(2027, 1, 3), dt.date(2027, 1, 9)),
]
ph_first = r + 1
for i, (name, a, b) in enumerate(phases):
    rr = r + 1 + i
    ws.cell(row=rr, column=1, value=name)
    ws.cell(row=rr, column=2, value=a).number_format = "DD-MMM-YY"
    ws.cell(row=rr, column=3, value=b).number_format = "DD-MMM-YY"
    ws.cell(row=rr, column=4, value=f'=COUNTIF(\'Daily Timetable\'!$D:$D,A{rr})')
    ws.cell(row=rr, column=5, value=f'=SUMIF(\'Daily Timetable\'!$D:$D,A{rr},\'Daily Timetable\'!$E:$E)')
    ws.cell(row=rr, column=6, value=f'=SUMIF(\'Daily Timetable\'!$D:$D,A{rr},\'Daily Timetable\'!$O:$O)')
    ws.cell(row=rr, column=7, value=f'=COUNTIFS(\'Daily Timetable\'!$D:$D,A{rr},\'Daily Timetable\'!$N:$N,"Y")')
    ws.cell(row=rr, column=8, value=f'=IF(D{rr}=0,0,G{rr}/D{rr})').number_format = "0%"
    for cc in range(1, 9):
        ws.cell(row=rr, column=cc).font = F(); ws.cell(row=rr, column=cc).border = BORDER
ph_last = r + len(phases)
rr = ph_last + 1
ws.cell(row=rr, column=1, value="TOTAL").font = F(True)
for col in "DEFG":
    ws[f"{col}{rr}"] = f"=SUM({col}{ph_first}:{col}{ph_last})"; ws[f"{col}{rr}"].font = F(True); ws[f"{col}{rr}"].border = BORDER
ws[f"H{rr}"] = f"=IF(D{rr}=0,0,G{rr}/D{rr})"; ws[f"H{rr}"].number_format = "0%"; ws[f"H{rr}"].font = F(True); ws[f"H{rr}"].border = BORDER
r = rr + 2

ws.cell(row=r, column=1, value="MONTHLY HOURS (calculated)").font = F(True, 11)
r += 1
hdr(ws, r, ["Month", "Planned Hours", "Actual Hours", "Avg Planned Hrs/Day", "Subject hours – Acc", "Subject hours – Law", "Subject hours – QA", "Subject hours – Eco"])
months = [("Oct-2026", 2026, 10), ("Nov-2026", 2026, 11), ("Dec-2026", 2026, 12), ("Jan-2027", 2027, 1)]  # Jan = 1–9 Jan only
for i, (lbl, y, m) in enumerate(months):
    rr = r + 1 + i
    ws.cell(row=rr, column=1, value=lbl)
    ws.cell(row=rr, column=2, value=f"=SUMIF('Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$E:$E)")
    ws.cell(row=rr, column=3, value=f"=SUMIF('Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$O:$O)")
    ws.cell(row=rr, column=4, value=f"=IF(COUNTIF('Daily Timetable'!$C:$C,A{rr})=0,0,B{rr}/COUNTIF('Daily Timetable'!$C:$C,A{rr}))").number_format = "0.0"
    for j, s in enumerate(["Acc", "Law", "QA", "Eco"]):
        ws.cell(row=rr, column=5 + j, value=f"=SUMIFS('Daily Timetable'!$I:$I,'Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$H:$H,\"{s}\")+SUMIFS('Daily Timetable'!$L:$L,'Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$K:$K,\"{s}\")")
    for cc in range(1, 9):
        ws.cell(row=rr, column=cc).font = F(); ws.cell(row=rr, column=cc).border = BORDER
r = r + len(months) + 2

ws.cell(row=r, column=1, value="SUGGESTED DAILY ROUTINE").font = F(True, 11)
r += 1
hdr(ws, r, ["Month", "Morning recap + Session 1 (Subject A)", "Session 2 (Subject B)", "Notes"])
routine = [
 ("October (7 h)", "6:00–6:30 am morning recap → 6:30–9:45 am Session 1 (last 15 min recap)", "5:30–9:00 pm Session 2 (last 15 min recap)", "Practical subject in the morning when fresh; theory in the evening."),
 ("November (9 h)", "6:00–6:45 am morning recap → 6:45–11:00 am Session 1", "3:30–8:00 pm Session 2", "Add a 30-min lunch-time MCQ drill using the 'buffer' minutes."),
 ("December (12 h)", "6:00–6:45 am morning recap → 7:00–9:30 am + 10:00 am–1:00 pm Session 1", "2:00–5:00 pm + 6:00–8:45 pm Session 2", "2:00–5:00 pm block mirrors the actual exam slot – always solve/mock in this window."),
 ("1–2 January (12 h)", "6:00–6:45 am morning recap → 7:00–9:30 am + 10:00 am–1:00 pm", "2:00–5:00 pm + 6:00–8:45 pm", "Sleep 7–8 h; no new topics after 27 Dec."),
]
for i, row_ in enumerate(routine):
    rr = r + 1 + i
    for cc, v in enumerate(row_, 1):
        c = ws.cell(row=rr, column=cc, value=v); c.font = F(); c.border = BORDER; c.alignment = WRAP
    ws.row_dimensions[rr].height = 32
for col, w in zip("ABCDEFGH", [44, 34, 30, 12, 16, 16, 16, 16]):
    ws.column_dimensions[col].width = w

# ===== Sheet 2: Daily Timetable =====
wt = wb.create_sheet("Daily Timetable")
headers = ["Date", "Day", "Month", "Phase", "Planned Hrs", "Recap Hrs", "Morning Recap – yesterday's topics (first thing in the day)", "Session 1 – Subject", "S1 Hrs", "Session 1 – What to study (chapter / unit / task) + session recap", "Session 2 – Subject", "S2 Hrs", "Session 2 – What to study (chapter / unit / task) + session recap", "Done? (Y/N)", "Actual Hrs", "Notes / Error log pointer"]
wt["A1"] = "DAILY TIMETABLE – 2 subjects per day | Morning recap → Session 1 (+recap) → Session 2 (+recap) | Sunday = weekly test + weekly recap | Yellow cells = fill in daily"; wt["A1"].font = F(True, 12, "1F4E3D")
hdr(wt, 2, headers)
dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True); wt.add_data_validation(dv)
SUBJ_FILL = {s: PatternFill("solid", fgColor=v[2]) for s, v in SUBJECTS.items()}
OTHER_FILL = {"Test": PatternFill("solid", fgColor="FFF2CC"), "Review": PatternFill("solid", fgColor="FFF2CC"), "Rest": PatternFill("solid", fgColor="EDEDED"), "All": PatternFill("solid", fgColor="FFF2CC")}
PHASE_FILL = {"Mock Exam 1": "F8CBAD", "Mock Exam 2": "F8CBAD", "Mock Exam 3": "F8CBAD", "EXAM DAY": "FF9999", "Revision 1": "C6E0B4", "Revision 2": "A9D08E", "Final Revision": "FFD966"}
first = 3
for i, row_ in enumerate(rows):
    rr = first + i
    d = row_["date"]
    wt.cell(row=rr, column=1, value=d).number_format = "DD-MMM-YY"
    wt.cell(row=rr, column=2, value=d.strftime("%a"))
    wt.cell(row=rr, column=3, value=d.strftime("%b-%Y"))
    wt.cell(row=rr, column=4, value=row_["phase"])
    wt.cell(row=rr, column=5, value=row_["hours"])
    s1, s2 = row_["s1"], row_["s2"]
    wt.cell(row=rr, column=6, value=row_["recap_h"]); wt.cell(row=rr, column=7, value=row_["recap"])
    wt.cell(row=rr, column=8, value=s1[0]); wt.cell(row=rr, column=9, value=s1[1]); wt.cell(row=rr, column=10, value=s1[2])
    wt.cell(row=rr, column=11, value=s2[0]); wt.cell(row=rr, column=12, value=s2[1]); wt.cell(row=rr, column=13, value=s2[2])
    for cc in range(1, 17):
        c = wt.cell(row=rr, column=cc); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
    for cc in (1, 2, 3, 5, 6, 9, 12, 14, 15):
        wt.cell(row=rr, column=cc).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
    wt.cell(row=rr, column=7).fill = PatternFill("solid", fgColor="F2F2F2"); wt.cell(row=rr, column=6).fill = PatternFill("solid", fgColor="F2F2F2")
    for subj_col, txt_col in ((8, 10), (11, 13)):
        s = wt.cell(row=rr, column=subj_col).value
        fill = SUBJ_FILL.get(s) or OTHER_FILL.get(s)
        if fill:
            wt.cell(row=rr, column=subj_col).fill = fill; wt.cell(row=rr, column=txt_col).fill = fill
        if s in SUBJECTS:
            wt.cell(row=rr, column=subj_col).font = F(True, 9, SUBJECTS[s][1])
    if row_["phase"] in PHASE_FILL:
        wt.cell(row=rr, column=4).fill = PatternFill("solid", fgColor=PHASE_FILL[row_["phase"]])
    if d.weekday() == 6:
        wt.cell(row=rr, column=2).fill = PatternFill("solid", fgColor="FFF2CC"); wt.cell(row=rr, column=2).font = F(True, 9)
    for cc in (14, 15, 16):
        wt.cell(row=rr, column=cc).fill = INPUT_FILL
    dv.add(wt.cell(row=rr, column=14))
    n_lines = max(str(s1[2]).count("\n"), str(s2[2]).count("\n")) + 1
    est = max(len(str(s1[2])), len(str(s2[2]))) / 70 + n_lines
    wt.row_dimensions[rr].height = max(45, min(150, 13 * est))
last = first + len(rows) - 1
# conditional formatting: done rows green tint on date column
wt.conditional_formatting.add(f"A{first}:A{last}", FormulaRule(formula=[f'$N{first}="Y"'], fill=PatternFill("solid", fgColor="92D050")))
wt.freeze_panes = "E3"
for col, w in zip("ABCDEFGHIJKLMNOP", [10, 5, 9, 20, 7, 6, 34, 9, 6, 60, 9, 6, 60, 7, 7, 22]):
    wt.column_dimensions[col].width = w
wt.auto_filter.ref = f"A2:P{last}"
wt.cell(row=last + 2, column=1, value="Legend: Acc = Paper 1 Accounting | Law = Paper 2 Business Laws | QA = Paper 3 Quantitative Aptitude | Eco = Paper 4 Business Economics. 'Uk' = Unit k of the ICAI Study Material chapter. PYQ = ICAI past-year questions (last 6 attempts). RTP/MTP = ICAI Revision Test Paper / Mock Test Paper. Hours in [ ] are the planned time for that item. Planned Hrs = Recap Hrs + S1 Hrs + S2 Hrs. Recap system: (1) MORNING RECAP – first 30–45 min of the day, closed-book active recall of yesterday's topics; (2) SESSION RECAP – last 15 min of every session, write the summary sheet + flashcards; (3) WEEKLY RECAP – Sunday, after the weekly test, re-read all summary sheets of the week and re-solve the week's Error Log.").font = F(italic=True, size=9)
wt.merge_cells(start_row=last + 2, start_column=1, end_row=last + 2, end_column=16)
wt.row_dimensions[last + 2].height = 45

# ===== Sheet 3: Chapter Master =====
wc = wb.create_sheet("Chapter Master")
wc["A1"] = "CHAPTER MASTER – syllabus coverage, hours and progress (calculated from Daily Timetable)"; wc["A1"].font = F(True, 12, "1F4E3D")
hdr(wc, 2, ["Paper", "Ch", "Chapter (ICAI Study Material)", "Units / sub-topics", "Approx marks in paper", "Exam priority", "Study hrs (Phase 1)", "Rev 1 hrs", "Rev 2 hrs", "Days scheduled (any phase)", "Study first date", "Study last date", "Rev 1 date", "Rev 2 date", "Self-rating after Rev 2 (1–5)", "Weak? (auto)"])
rr = 3
for s, chs in CH.items():
    for no, name, units in chs:
        tot = round(sum(h for _, h in units) * wfac(s, no) * 2) / 2
        key = f"{s} Ch {no}:"
        wc.cell(row=rr, column=1, value=SUBJECTS[s][0])
        wc.cell(row=rr, column=2, value=no)
        wc.cell(row=rr, column=3, value=name)
        wc.cell(row=rr, column=4, value="; ".join(u for u, _ in units if not u.startswith("PYQ")))
        wc.cell(row=rr, column=5, value=IMPORTANCE[(s, no)][0])
        pc = wc.cell(row=rr, column=6, value=IMPORTANCE[(s, no)][1])
        pc.fill = PatternFill("solid", fgColor={"High": "F4B084", "Medium": "FFE699", "Low": "E2EFDA"}[IMPORTANCE[(s, no)][1]])
        wc.cell(row=rr, column=7, value=tot)
        wc.cell(row=rr, column=8, value=f"=ROUND(G{rr}*0.42*2,0)/2")
        wc.cell(row=rr, column=9, value=f"=ROUND(G{rr}*0.3*2,0)/2")
        pat = f'"*{key}*"'
        wc.cell(row=rr, column=10, value=f"=COUNTIF('Daily Timetable'!$J:$J,{pat})+COUNTIF('Daily Timetable'!$M:$M,{pat})")
        p1_rows = f"'Daily Timetable'!$A${first}:$A${last}"
        cond = f"(ISNUMBER(SEARCH(\"{key}\",'Daily Timetable'!$J${first}:$J${last}))+ISNUMBER(SEARCH(\"{key}\",'Daily Timetable'!$M${first}:$M${last}))>0)"
        ph = f"'Daily Timetable'!$D${first}:$D${last}"
        minexpr = f'SUMPRODUCT(MIN({p1_rows}+({cond}*({ph}="Phase 1 – Study (SM + PYQ)")=0)*100000))'
        wc.cell(row=rr, column=11, value=f'=IF({minexpr}>100000,"",{minexpr})').number_format = "DD-MMM"
        wc.cell(row=rr, column=12, value=f'=IFERROR(1/(1/SUMPRODUCT(MAX({p1_rows}*{cond}*({ph}="Phase 1 – Study (SM + PYQ)")))),"")').number_format = "DD-MMM"
        wc.cell(row=rr, column=13, value=f'=IFERROR(1/(1/SUMPRODUCT(MAX({p1_rows}*{cond}*({ph}="Revision 1")))),"")').number_format = "DD-MMM"
        wc.cell(row=rr, column=14, value=f'=IFERROR(1/(1/SUMPRODUCT(MAX({p1_rows}*{cond}*({ph}="Revision 2")))),"")').number_format = "DD-MMM"
        c = wc.cell(row=rr, column=15); c.fill = INPUT_FILL
        wc.cell(row=rr, column=16, value=f'=IF(O{rr}="","",IF(O{rr}<=2,"WEAK – add to Final Revision",""))')
        for cc in range(1, 17):
            cell = wc.cell(row=rr, column=cc); cell.font = F(size=9); cell.border = BORDER; cell.alignment = WRAP
            if cc in (1, 3):
                cell.fill = SUBJ_FILL[s]
        rr += 1
tot_row = rr
wc.cell(row=tot_row, column=3, value="TOTAL").font = F(True)
for col in "GHIJ":
    wc[f"{col}{tot_row}"] = f"=SUM({col}3:{col}{rr-1})"; wc[f"{col}{tot_row}"].font = F(True); wc[f"{col}{tot_row}"].border = BORDER
wc.cell(row=tot_row + 2, column=1, value="Hours are planning estimates set by ArivuPro from ICAI chapter size and past-paper weight, with a subject weighting of ×1.2 on Accounting and ×1.35 on Business Laws (×0.9 on QA and Economics) and a chapter priority factor from exam importance (High ×1.25, Medium ×1.0, Low ×0.75). 'Approx marks' is ArivuPro's estimate from ICAI section-wise weightage and the June 2024 – Jan 2026 papers. Phase 1 and Revision 1 follow ICAI chapter order (concept dependencies); Revision 2 takes HIGH-priority chapters first; Phase 1 hours include unit reading, ICAI illustrations/TYK and past-year questions. Revision 1 ≈ 42% and Revision 2 ≈ 30% of Phase 1 time. Actual allocation in the timetable is scaled to the hours available in each phase.").font = F(italic=True, size=9)
wc.merge_cells(start_row=tot_row + 2, start_column=1, end_row=tot_row + 2, end_column=16)
wc.row_dimensions[tot_row + 2].height = 30
dv2 = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True); wc.add_data_validation(dv2); dv2.add(f"O3:O{rr-1}")
wc.freeze_panes = "D3"
for col, w in zip("ABCDEFGHIJKLMNOP", [26, 5, 36, 56, 14, 10, 9, 8, 8, 10, 10, 10, 10, 10, 12, 22]):
    wc.column_dimensions[col].width = w

# ===== Sheet 4: Mock & Test Tracker =====
wm = wb.create_sheet("Mock & Test Tracker")
wm["A1"] = "MOCK EXAMS & WEEKLY TESTS – score tracker (fill yellow cells)"; wm["A1"].font = F(True, 12, "1F4E3D")
hdr(wm, 2, ["Mock", "Date", "Paper", "Max Marks", "Marks Scored", "%", "Rank benchmark (85%)", "Gap to benchmark", "Time taken (min)", "Top 3 mistakes / topics to fix"])
rr = 3
mock_blocks = [(1, MOCK1), (2, MOCK2), (3, MOCK3)]
first_m = rr
for n, dates in mock_blocks:
    for d, s in zip(dates, MOCK_ORDER):
        wm.cell(row=rr, column=1, value=f"Mock {n}")
        wm.cell(row=rr, column=2, value=d).number_format = "DD-MMM-YY"
        wm.cell(row=rr, column=3, value=SUBJECTS[s][0])
        wm.cell(row=rr, column=4, value=100)
        e = wm.cell(row=rr, column=5); e.fill = INPUT_FILL
        wm.cell(row=rr, column=6, value=f'=IF(E{rr}="","",E{rr}/D{rr})').number_format = "0%"
        wm.cell(row=rr, column=7, value=f"=D{rr}*0.85")
        wm.cell(row=rr, column=8, value=f'=IF(E{rr}="","",E{rr}-G{rr})')
        wm.cell(row=rr, column=9).fill = INPUT_FILL
        wm.cell(row=rr, column=10).fill = INPUT_FILL
        for cc in range(1, 11):
            c = wm.cell(row=rr, column=cc); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
        wm.cell(row=rr, column=3).fill = SUBJ_FILL[s]
        rr += 1
    # aggregate row
    wm.cell(row=rr, column=1, value=f"Mock {n} – AGGREGATE").font = F(True)
    wm.cell(row=rr, column=4, value=f"=SUM(D{rr-4}:D{rr-1})")
    wm.cell(row=rr, column=5, value=f'=IF(COUNT(E{rr-4}:E{rr-1})=0,"",SUM(E{rr-4}:E{rr-1}))')
    wm.cell(row=rr, column=6, value=f'=IF(E{rr}="","",E{rr}/D{rr})').number_format = "0%"
    wm.cell(row=rr, column=7, value=f"=D{rr}*0.85")
    wm.cell(row=rr, column=8, value=f'=IF(E{rr}="","",E{rr}-G{rr})')
    wm.cell(row=rr, column=10, value=f'=IF(E{rr}="","",IF(F{rr}>=0.85,"On track for rank",IF(F{rr}>=0.7,"Strong pass – push weak paper to 80+",IF(F{rr}>=0.5,"Pass zone – double revision time on lowest paper","Below pass – re-plan with faculty"))))')
    for cc in range(1, 11):
        c = wm.cell(row=rr, column=cc); c.font = F(True, 9); c.border = BORDER; c.fill = SUB_FILL; c.alignment = WRAP
    rr += 2
rr += 1
wm.cell(row=rr, column=1, value="WEEKLY SUNDAY TESTS").font = F(True, 11); rr += 1
hdr(wm, rr, ["Week", "Test date", "Chapters covered (from timetable)", "Max Marks", "Marks Scored", "%", "Target (75%)", "Gap", "Time taken (min)", "Errors logged / action"]); rr += 1
wk = 1
for row_ in rows:
    if row_["s1"][0] == "Test":
        wm.cell(row=rr, column=1, value=f"Week {wk}")
        wm.cell(row=rr, column=2, value=row_["date"]).number_format = "DD-MMM-YY"
        wm.cell(row=rr, column=3, value=row_["s1"][2].split("\n")[0].replace("WEEKLY TEST (timed, closed book): ", "").replace("WEEKLY REVISION TEST (timed): ", ""))
        wm.cell(row=rr, column=4, value=100).fill = INPUT_FILL
        wm.cell(row=rr, column=5).fill = INPUT_FILL
        wm.cell(row=rr, column=6, value=f'=IF(E{rr}="","",E{rr}/D{rr})').number_format = "0%"
        wm.cell(row=rr, column=7, value=f"=D{rr}*0.75")
        wm.cell(row=rr, column=8, value=f'=IF(E{rr}="","",E{rr}-G{rr})')
        wm.cell(row=rr, column=9).fill = INPUT_FILL; wm.cell(row=rr, column=10).fill = INPUT_FILL
        for cc in range(1, 11):
            c = wm.cell(row=rr, column=cc); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
        wm.row_dimensions[rr].height = 30
        rr += 1; wk += 1
wm.cell(row=rr + 1, column=1, value="Weekly tests: set 100 marks (Acc/Law descriptive ~50% + QA/Eco MCQs ~50%) from the chapters listed, or use ArivuPro chapter tests. Max marks cell is editable if a different total is used.").font = F(italic=True, size=9)
wm.merge_cells(start_row=rr + 1, start_column=1, end_row=rr + 1, end_column=10)
for col, w in zip("ABCDEFGHIJ", [20, 11, 44, 9, 10, 7, 12, 10, 10, 40]):
    wm.column_dimensions[col].width = w
wm.freeze_panes = "A3"

# ===== Sheet 5: Rank Strategy =====
wr = wb.create_sheet("Rank Strategy")
wr["A1"] = "RANK STRATEGY – what a rank needs beyond a pass"; wr["A1"].font = F(True, 12, "1F4E3D")
hdr(wr, 2, ["Paper", "Pattern", "Marks target for rank", "Time plan in the hall", "Where ranks are won", "Where marks are lost"])
strat = [
 ("Paper 1 – Accounting (100)", "Descriptive, 3 h. Q1 compulsory (theory/T-F/short practicals) + 4 of 5 remaining; ~30% theory, ~70% practical", "80–90", "Use the 1:45–2:00 pm reading time to pick your 4 optional questions → attempt Q1 last (30 min) → 4 best practicals @ ~35 min each. Keep 10 min for checking.", "Full working notes; correct formats (Trading/P&L, Balance Sheet, R&P, I&E, Partners' Capital, Journal narration); Ch 7, 10, 11 carry the heaviest marks – master them.", "Missing narrations, no working notes, wrong dates in BRS/Bills, skipping T/F reasoning, poor presentation."),
 ("Paper 2 – Business Laws (100)", "Descriptive, 3 h. Q1 compulsory + 4 of 5; case-study style questions per chapter", "75–85", "Reading time 1:45–2:00 pm: mark the case questions you know best. Case questions: Facts → Provision (section + Act) → Analysis → Conclusion (4-step answer). ~35 min per question, theory questions 6–8 min per 4-marker.", "Quoting the section number and a landmark case; Contract Act (~30–35 marks) + Sale of Goods + Partnership are the bulk; structured headings.", "Writing stories without provision; missing conclusion; no distinction tables for 'distinguish' questions; running out of time on Q1."),
 ("Paper 3 – Quantitative Aptitude (100)", "MCQ, 2 h, 100 Qs; ¼ mark negative. Maths 40 + LR 20 + Stats 40", "80–90", "2:00–4:00 pm, no reading time – start instantly. Round 1 (60 min): all sure-shot questions. Round 2 (40 min): medium ones. Round 3 (20 min): guess only where 2 options eliminated; leave rest blank.", "Speed from formula recall (Time Value of Money, Statistics), LR accuracy (20 nearly free marks), calculator fluency (12-digit, all functions).", "Negative marking from blind guessing; slow LR; calculation slips in SD/regression/index numbers; not knowing calculator shortcuts."),
 ("Paper 4 – Business Economics (100)", "MCQ, 2 h, 100 Qs; ¼ mark negative", "80–90", "2:00–4:00 pm, no reading time. Same 3-round method. ~1 min per question; flag doubtful ones and return.", "Definitions + diagrams + numericals (elasticity, multiplier, national income); Indian economy current facts (budget, RBI rates, GDP).", "Confusing similar concepts (types of elasticity, cost curves, market forms); ignoring Ch 5–10 which together carry ~50 marks."),
]
for i, row_ in enumerate(strat):
    rr = 3 + i
    for cc, v in enumerate(row_, 1):
        c = wr.cell(row=rr, column=cc, value=v); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
    wr.row_dimensions[rr].height = 110
rr = 8
wr.cell(row=rr, column=1, value="AGGREGATE TARGET").font = F(True, 11); rr += 1
tgt = [("Pass criteria", "40 in each paper AND 50% aggregate (200/400)"), ("Rank zone (AIR 1–50 typically)", "Aggregate 320–360+ / 400 (80–90%). Target this plan: 340/400 = 85%."), ("Mock benchmark", "Mock 1 ≥ 60%, Mock 2 ≥ 75%, Mock 3 ≥ 85% aggregate. The Mock Tracker flags the gap automatically.")]
for a, b in tgt:
    wr.cell(row=rr, column=1, value=a).font = F(True, 9); wr.cell(row=rr, column=2, value=b).font = F(size=9)
    wr.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=6)
    for cc in range(1, 7): wr.cell(row=rr, column=cc).border = BORDER; wr.cell(row=rr, column=cc).alignment = WRAP
    rr += 1
rr += 1
wr.cell(row=rr, column=1, value="NON-NEGOTIABLE HABITS (rank discipline)").font = F(True, 11); rr += 1
habits = [
 "1. Error Log: one notebook. Every mistake in a PYQ, weekly test or mock goes in with the correct approach. Re-read it before every test; this is the single highest-ROI habit.",
 "1a. Recap discipline: morning recap (yesterday, closed book) → session recap (last 15 min, summary sheet + flashcards) → weekly recap (Sunday). Skipping recaps is the commonest reason a first reading is forgotten by December.",
 "2. One-page summary per chapter, built up session by session during the session recaps (formats for Acc, section/case list for Law, formula sheet for QA, diagram + definition list for Eco). Revision 2 and the day-before-exam skim use only these pages.",
 "3. Practical subjects are studied by WRITING, never by reading solutions. Solve every ICAI illustration with the book closed, then compare.",
 "4. PYQ = last 6 attempts (Jan 2026, Sep 2025, May 2025, Jan 2025, Sep 2024, June 2024) + RTP/MTP of the current attempt. Repeated questions are the highest-probability questions.",
 "5. Mocks are written at the real exam time – Acc/Law 1:45 pm reading + 2:00–5:00 pm on paper; QA/Eco 2:00–4:00 pm OMR-style with no reading time – strictly timed. Never pause a mock.",
 "6. Backlog rule: if a day's plan slips, it moves to that Sunday's review block – never to the next study day. Nothing pushes the mock dates.",
 "7. From 16 Dec no new source books, no new videos. Only ICAI material, summaries, Error Log and mocks.",
 "8. Health: 7 h sleep in Oct–Nov, 7.5 h in Dec–Jan; 20 min walk daily; screen off 30 min before sleep. A rank is a stamina event.",
]
for h in habits:
    c = wr.cell(row=rr, column=1, value=h); c.font = F(size=9); c.alignment = WRAP
    wr.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
    wr.row_dimensions[rr].height = 30
    rr += 1
for col, w in zip("ABCDEF", [30, 40, 14, 44, 44, 44]):
    wr.column_dimensions[col].width = w

for w in (ws, wt, wc, wm, wr):
    w.page_setup.orientation = "landscape"; w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0
    w.sheet_properties.pageSetUpPr.fitToPage = True
wt.print_title_rows = "2:2"
wb.save("../excel/CA_Foundation_Jan2027_Rank_Timetable.xlsx")

# diagnostics
import json
tot_by_month = defaultdict(float)
subj_by_phase = defaultdict(float)
for row_ in rows:
    tot_by_month[row_["date"].strftime("%b")] += row_["hours"]
    for s in (row_["s1"], row_["s2"]):
        subj_by_phase[(row_["phase"][:10], s[0])] += s[1]
print(dict(tot_by_month))
for k in sorted(subj_by_phase): print(k, round(subj_by_phase[k], 1))

# ---- export plan for web app
import json, re
KEYRE = re.compile(r"^(Acc|Law|QA|Eco) Ch (\d+):")
out_days = []
for row_ in rows:
    sess = []
    for sx in (row_["s1"], row_["s2"]):
        if not sx[0] and not sx[2]:
            continue
        items = []
        for line in str(sx[2]).split("\n"):
            line = line.strip()
            if not line:
                continue
            m = KEYRE.match(line)
            hm = re.search(r"\[([\d.]+)h\]\s*$", line)
            txt = re.sub(r"\s*\[[\d.]+h\]\s*$", "", line)
            items.append({"t": txt, "k": f"{m.group(1)}-{m.group(2)}" if m else None, "h": float(hm.group(1)) if hm else None})
        sess.append({"s": sx[0], "h": sx[1], "items": items})
    out_days.append({"d": row_["date"].isoformat(), "p": row_["phase"], "h": row_["hours"], "rh": row_["recap_h"], "recap": row_["recap"], "sess": sess})
chap = []
for s_, chs in CH.items():
    for no, name, units in chs:
        chap.append({"k": f"{s_}-{no}", "s": s_, "no": no, "n": name, "marks": IMPORTANCE[(s_, no)][0], "pr": IMPORTANCE[(s_, no)][1],
                     "hrs": round(sum(h for _, h in units) * wfac(s_, no) * 2) / 2})
json.dump({"days": out_days, "chapters": chap}, open("plan.json", "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
