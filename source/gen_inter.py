import sys, json, re, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from inter_data import SUBJECTS, PAPERS, PAPER_FILL, PAPER_TIME, CH, IMP

D = dt.date
def dr(a, b):
    d = a
    while d <= b:
        yield d; d += dt.timedelta(1)

# ------------------------------------------------------------------ plan configs
PLANS = {
 "g1": dict(
   title="CA Intermediate Group I", short="Inter Group I", file="CA_Inter_Group1_Jan2027_Timetable.xlsx",
   papers=["AA", "LAW", "TAX"], blocks=[("AA", "LAW"), ("IT", "GST")],
   hours={10: (8, 4), 11: (10, 5), 12: (12, 6), 1: (12, 6)},
   p1_end=D(2026, 11, 21), buffer=D(2026, 11, 22),
   mocks={1: [(D(2026, 11, 23), ["AA"]), (D(2026, 11, 24), ["LAW"]), (D(2026, 11, 25), ["TAX"])],
          2: [(D(2026, 12, 11), ["AA"]), (D(2026, 12, 12), ["LAW"]), (D(2026, 12, 13), ["TAX"])],
          3: [(D(2026, 12, 26), ["AA"]), (D(2026, 12, 27), ["LAW"]), (D(2026, 12, 28), ["TAX"])]},
   r1=(D(2026, 11, 26), D(2026, 12, 10)), r2=(D(2026, 12, 14), D(2026, 12, 25)),
   final={
     D(2026, 12, 29): [("AA", 6, "FINAL REVISION – Advanced Accounting: all AS from one-page summaries (AS number → key rule → one numerical each); formats for Schedule III, amalgamation, consolidation"), ("LAW", 6, "FINAL REVISION – Corporate & Other Laws: section list per chapter, time limits & numbers sheet; 2 case questions per chapter")],
     D(2026, 12, 30): [("IT", 6, "FINAL REVISION – Income Tax: PGBP & capital gains computations, deductions table, comprehensive TI computation (1 full question)"), ("GST", 6, "FINAL REVISION – GST: supply/exemption/ITC case questions; value & place of supply charts; payment & returns due dates")],
     D(2026, 12, 31): [("AA", 7, "FINAL REVISION – Advanced Accounting weak chapters (from Chapter Master ratings) + Mock 3 errors"), ("LAW", 5, "FINAL REVISION – Law weak chapters + General Clauses, Interpretation, FEMA summaries")],
     D(2027, 1, 1): [("AA", 9, "DAY BEFORE PAPER 1 – Advanced Accounting only: every AS summary, then formats (Schedule III, cash flow, amalgamation, AS 21, branch); 30 MCQs timed"), ("Rest", 3, "Stop by 8 pm. Pack admit card, calculator, pens. Sleep 8 h")],
     D(2027, 1, 3): [("LAW", 9, "DAY BEFORE PAPER 2 – Corporate & Other Laws only: section list + time limits; landmark provisions per chapter; 30 MCQs timed"), ("Rest", 3, "Stop by 8 pm")],
     D(2027, 1, 5): [("TAX", 9, "DAY BEFORE PAPER 3 – Taxation only: IT computation formats + GST charts; amendments list; 30 MCQs timed"), ("Rest", 3, "Stop by 8 pm")],
   },
   bench=70, bench_label="Exemption/distinction zone (70+)",
   note="Single-group candidates are not eligible for an All-India Rank (ICAI merit list is for candidates passing both groups together). This plan targets 70+ per paper: exemption in every paper and a distinction-level aggregate.",
   special=[("PGBP", "Heads conquered", "Income Tax Ch 3 – every head of income ticked", "IT-3"), ("AS", "Standards set", "Adv. Accounting Ch 5 – all assets-based AS ticked", "AA-5")],
 ),
 "g2": dict(
   title="CA Intermediate Group II", short="Inter Group II", file="CA_Inter_Group2_Jan2027_Timetable.xlsx",
   papers=["CMA", "AUD", "FMSM"], blocks=[("CMA", "AUD"), ("FM", "SM")],
   hours={10: (8, 4), 11: (10, 5), 12: (12, 6), 1: (12, 6)},
   p1_end=D(2026, 11, 25), buffer=D(2026, 11, 26),
   mocks={1: [(D(2026, 11, 27), ["CMA"]), (D(2026, 11, 28), ["AUD"]), (D(2026, 11, 29), ["FMSM"])],
          2: [(D(2026, 12, 16), ["CMA"]), (D(2026, 12, 17), ["AUD"]), (D(2026, 12, 18), ["FMSM"])],
          3: [(D(2027, 1, 2), ["CMA"]), (D(2027, 1, 3), ["AUD"]), (D(2027, 1, 4), ["FMSM"])]},
   r1=(D(2026, 11, 30), D(2026, 12, 15)), r2=(D(2026, 12, 19), D(2027, 1, 1)),
   final={
     D(2027, 1, 5): [("CMA", 6, "FINAL REVISION – Costing: variance formula sheet, process costing & marginal costing formats; 1 full question each from Ch 10, 13, 14"), ("AUD", 6, "FINAL REVISION – Audit: SA list (number → objective → key requirement), CARO/report formats, ethics clauses")],
     D(2027, 1, 6): [("FM", 6, "FINAL REVISION – FM: formula sheet (cost of capital, leverage, capital budgeting, WC); 1 numerical per chapter"), ("SM", 6, "FINAL REVISION – SM: chapter mind-maps, models (Porter, BCG, Ansoff, value chain) with one case application each")],
     D(2027, 1, 7): [("CMA", 9, "DAY BEFORE PAPER 4 – Cost & Management Accounting only: formula sheet all chapters; 30 MCQs timed; weak chapters from ratings"), ("Rest", 3, "Stop by 8 pm. Pack admit card, calculator, pens. Sleep 8 h")],
     D(2027, 1, 9): [("AUD", 9, "DAY BEFORE PAPER 5 – Auditing & Ethics only: SA summaries, reporting, bank & special-entity audit points; ethics clauses; 30 MCQs"), ("Rest", 3, "Stop by 8 pm")],
     D(2027, 1, 11): [("FMSM", 9, "DAY BEFORE PAPER 6 – FM formula sheet + SM mind-maps; 30 MCQs timed"), ("Rest", 3, "Stop by 8 pm")],
   },
   bench=70, bench_label="Exemption/distinction zone (70+)",
   note="Single-group candidates are not eligible for an All-India Rank (ICAI merit list is for candidates passing both groups together). This plan targets 70+ per paper: exemption in every paper and a distinction-level aggregate.",
   special=[("VAR", "Variance virtuoso", "Costing Ch 13 Standard Costing – every unit ticked", "CMA-13"), ("SA", "Evidence expert", "Audit Ch 4 Audit Evidence – every unit ticked", "AUD-4")],
 ),
 "both": dict(
   title="CA Intermediate Both Groups", short="Inter Both Groups", file="CA_Inter_BothGroups_Jan2027_Timetable.xlsx",
   papers=["AA", "LAW", "TAX", "CMA", "AUD", "FMSM"], blocks=[("AA", "LAW"), ("CMA", "AUD"), ("IT", "GST"), ("FM", "SM")],
   hours={10: (10, 5), 11: (11, 5.5), 12: (12, 6), 1: (12, 6)},
   p1_end=D(2026, 11, 21), buffer=D(2026, 11, 22),
   mocks={1: [(D(2026, 11, 23) + dt.timedelta(i), [p]) for i, p in enumerate(["AA", "LAW", "TAX", "CMA", "AUD", "FMSM"])],
          2: [(D(2026, 12, 13) + dt.timedelta(i), [p]) for i, p in enumerate(["AA", "LAW", "TAX", "CMA", "AUD", "FMSM"])],
          3: [(D(2026, 12, 27), ["AA", "CMA"]), (D(2026, 12, 28), ["LAW", "AUD"]), (D(2026, 12, 29), ["TAX", "FMSM"])]},
   r1=(D(2026, 11, 29), D(2026, 12, 12)), r2=(D(2026, 12, 19), D(2026, 12, 26)),
   final={
     D(2026, 12, 30): [("AA", 6, "FINAL REVISION – Advanced Accounting: AS summaries + formats (Schedule III, amalgamation, consolidation)"), ("LAW", 6, "FINAL REVISION – Corporate & Other Laws: section list, time limits, 2 case questions per chapter")],
     D(2026, 12, 31): [("TAX", 7, "FINAL REVISION – Taxation: IT computation formats + GST charts and amendments"), ("CMA", 5, "FINAL REVISION – Group II skim: Costing formula sheet (keep it warm for 8 Jan)")],
     D(2027, 1, 1): [("AA", 9, "DAY BEFORE PAPER 1 – Advanced Accounting only: every AS summary, formats, 30 MCQs timed"), ("Rest", 3, "Stop by 8 pm. Pack admit card, calculator, pens. Sleep 8 h")],
     D(2027, 1, 3): [("LAW", 9, "DAY BEFORE PAPER 2 – Corporate & Other Laws only: sections + time limits; 30 MCQs"), ("Rest", 3, "Stop by 8 pm")],
     D(2027, 1, 5): [("TAX", 9, "DAY BEFORE PAPER 3 – Taxation only: IT formats + GST charts; 30 MCQs"), ("Rest", 3, "Stop by 8 pm")],
     D(2027, 1, 7): [("CMA", 9, "DAY BEFORE PAPER 4 – Costing only: formula sheet all chapters; 30 MCQs timed"), ("Rest", 3, "Stop by 8 pm")],
     D(2027, 1, 9): [("AUD", 9, "DAY BEFORE PAPER 5 – Auditing & Ethics only: SA summaries, reporting, ethics; 30 MCQs"), ("Rest", 3, "Stop by 8 pm")],
     D(2027, 1, 11): [("FMSM", 9, "DAY BEFORE PAPER 6 – FM formula sheet + SM mind-maps; 30 MCQs"), ("Rest", 3, "Stop by 8 pm")],
   },
   exam_evening={D(2027, 1, 2): "Evening (after a 2 h break): 1.5 h Group II summary skim – Costing formula sheet", D(2027, 1, 4): "Evening (after a 2 h break): 1.5 h Group II summary skim – Audit SA list", D(2027, 1, 6): "Evening (after a 2 h break): 1.5 h Group II summary skim – FM formula sheet + SM models"},
   bench=70, bench_label="Rank zone (70+ per paper, ~420+/600)",
   note="AIR is awarded only to candidates who pass both groups in the same attempt. Recent Inter toppers score roughly 70–85% of 600; this plan targets 70+ in every paper (420/600) as the floor for a rank.",
   special=[("PGBP", "Heads conquered", "Income Tax Ch 3 – every head of income ticked", "IT-3"), ("VAR", "Variance virtuoso", "Costing Ch 13 Standard Costing – every unit ticked", "CMA-13")],
 ),
}

PRIO = {"High": 1.25, "Medium": 1.0, "Low": 0.75}
# relative weight per subject (≈ study effort per paper/section), used to normalise the unit-hour estimates
TARGET = {"AA": 100, "LAW": 80, "IT": 57, "GST": 45, "CMA": 100, "AUD": 80, "FM": 57, "SM": 38}
RAW = {s: sum(h * PRIO[IMP[s][no][1]] for no, _, us in CH[s] for _, h in us) for s in CH}
SW = {s: TARGET[s] / RAW[s] for s in CH}
def q(x): return round(x * 4) / 4

def build(pid):
    C = PLANS[pid]
    subs = [s for p in C["papers"] for s in PAPERS[p][1]]
    START = D(2026, 10, 1)
    EXAMS = {PAPERS[p][2]: p for p in C["papers"]}
    END = max(EXAMS)
    MOCKDAY = {d: (n, ps) for n, lst in C["mocks"].items() for d, ps in lst}
    FINAL = C["final"]

    def phase_of(d):
        if d <= C["p1_end"]: return "Phase 1 – Study (SM + PYQ)"
        if d == C["buffer"]: return "Buffer / Prep for Mock 1"
        if d in MOCKDAY: return f"Mock Exam {MOCKDAY[d][0]}"
        if C["r1"][0] <= d <= C["r1"][1]: return "Revision 1"
        if C["r2"][0] <= d <= C["r2"][1]: return "Revision 2"
        if d in EXAMS: return "EXAM DAY"
        return "Final Revision"

    def planned(d):
        wk, sun = C["hours"][d.month]
        if d in MOCKDAY or d in FINAL: return wk
        return sun if d.weekday() == 6 else wk

    def recap_h(d):
        if d in EXAMS: return 0
        if d.weekday() == 6: return 0.5
        return 0.5 if d.month == 10 else 0.75

    def wf(s, no): return PRIO[IMP[s][no][1]] * SW[s]

    def units(phase):
        out = {}
        for s in subs:
            chs = CH[s]
            if phase == "R2":
                chs = sorted(chs, key=lambda c: {"High": 0, "Medium": 1, "Low": 2}[IMP[s][c[0]][1]])
            lst = []
            for no, name, us in chs:
                tot = sum(h for _, h in us) * wf(s, no)
                key = f"{s} Ch {no}:"
                nm = f"{name} [{IMP[s][no][1].upper()}]"
                if phase == 1:
                    for u, h in us:
                        lst.append({"label": f"{key} {nm} — {u}", "hours": h * wf(s, no), "key": key})
                elif phase == "R1":
                    lst.append({"label": f"{key} {nm} — REVISION 1: re-read concept notes, re-solve ICAI illustrations, PYQs (last 6 attempts), RTP; 10 case-scenario MCQs", "hours": max(1, round(tot * .42 * 2) / 2), "key": key})
                else:
                    lst.append({"label": f"{key} {nm} — REVISION 2: one-page summary/formula sheet, MTP questions, timed MCQ drill, weak-area fix", "hours": max(1, round(tot * .30 * 2) / 2), "key": key})
            out[s] = lst
        return out

    def alloc_block(U, days, a, b):
        avail = sum(h for _, h in days); need = sum(u["hours"] for s in (a, b) for u in U[s])
        sc = avail * .965 / need
        for s in (a, b):
            for u in U[s]:
                u["rem"] = max(.25, q(u["hours"] * sc)); u["full"] = u["rem"]
        ptr = {a: 0, b: 0}; rem = {s: sum(u["rem"] for u in U[s]) for s in (a, b)}
        plan, lastd = {}, {}
        for i, (d, hrs) in enumerate(days):
            left = len(days) - i
            fa, fb = rem[a] / left, rem[b] / left
            split = q(hrs * fa / (fa + fb)) if fa + fb > 0 else hrs / 2
            split = min(max(split, .5), hrs - .5)
            ss = []
            for s, h in ((a, split), (b, hrs - split)):
                tx, ks, lf = [], [], h
                while lf > .01 and ptr[s] < len(U[s]):
                    u = U[s][ptr[s]]; take = min(u["rem"], lf)
                    u["rem"] -= take; lf -= take; rem[s] -= take
                    tag = ""
                    if u["rem"] > .01: tag = " (start)" if abs(take + u["rem"] - u["full"]) < .01 else " (contd.)"
                    elif abs(take - u["full"]) > .01: tag = " (finish)"
                    tx.append(f"{u['label']}{tag} [{take:g}h]"); ks.append(u["key"])
                    if u["rem"] <= .01: ptr[s] += 1
                if lf > .01: tx.append(f"{s}: buffer – Error-Log review / MCQ drill / pending topics [{lf:g}h]")
                ss.append([s, h, tx, ks]); lastd[s] = d
            plan[d] = ss
        for s in (a, b):
            while ptr[s] < len(U[s]):
                u = U[s][ptr[s]]
                for x in plan[lastd[s]]:
                    if x[0] == s: x[2].append(f"{u['label']} (carry-over – use buffer time) [{u['rem']:g}h]"); x[3].append(u["key"])
                ptr[s] += 1
        return {d: [(s, h, "\n".join(t), k) for s, h, t, k in ss] for d, ss in plan.items()}

    def allocate(U, days):
        need = [sum(u["hours"] for s in blk for u in U[s]) for blk in C["blocks"]]
        tot = sum(h for _, h in days); plan = {}; start = 0; cum = 0; target = 0
        for bi, blk in enumerate(C["blocks"]):
            target += tot * need[bi] / sum(need)
            if bi == len(C["blocks"]) - 1:
                end = len(days)
            else:
                end = start
                while end < len(days) and cum + days[end][1] <= target + days[end][1] / 2:
                    cum += days[end][1]; end += 1
                end = max(end, start + 1)
            plan.update(alloc_block(U, days[start:end], *blk)); start = end
        return plan

    def sdays(a, b): return [(d, planned(d) - recap_h(d)) for d in dr(a, b) if d.weekday() != 6 and d not in MOCKDAY]
    plan1 = allocate(units(1), sdays(START, C["p1_end"]))
    planr1 = allocate(units("R1"), sdays(*C["r1"]))
    planr2 = allocate(units("R2"), sdays(*C["r2"]))

    def week_keys(plan, sun):
        seen = []
        for k in range(6, 0, -1):
            for s, h, t, ks in plan.get(sun - dt.timedelta(k), []):
                for x in ks:
                    if x not in seen: seen.append(x)
        return seen

    SR = "\n↳ Last 15 min = SESSION RECAP: close the book, write this session's key points / formats / sections / formulas on the chapter's one-page summary sheet + 5 flashcards"
    rows = []
    mock_src = {1: "ICAI MTP Series 1 (or an ArivuPro full-syllabus test)", 2: "ICAI MTP Series 2", 3: "Latest ICAI question paper / RTP-based full test"}
    for d in dr(START, END):
        hrs = planned(d); ph = phase_of(d); rh = recap_h(d); keys = []
        s1 = s2 = ("", 0, "")
        src = plan1 if d in plan1 else planr1 if d in planr1 else planr2 if d in planr2 else None
        if src:
            ss = src[d]
            s1 = (ss[0][0], ss[0][1], ss[0][2] + SR); keys += ss[0][3]
            s2 = (ss[1][0], ss[1][1], ss[1][2] + SR); keys += ss[1][3]
        elif d.weekday() == 6 and d not in MOCKDAY and d not in EXAMS and d not in FINAL and d != C["buffer"]:
            pl = plan1 if d <= C["p1_end"] else planr1 if d <= C["r1"][1] else planr2
            ks = week_keys(pl, d); kt = ", ".join(ks).replace(":", "") or "chapters of this week"
            th = round((hrs - rh) * .5 * 2) / 2
            if d <= C["p1_end"]:
                s1 = ("Test", th, f"WEEKLY TEST (timed, closed book): {kt}\nICAI pattern: ~30% case-scenario MCQs + descriptive questions from this week's chapters")
                s2 = ("Review", hrs - rh - th, "Self-evaluate against ICAI suggested answers → Error Log → clear the week's backlog\nWEEKLY RECAP (last 45 min): read the one-page summary sheet of EVERY chapter covered Mon–Sat, re-solve the week's Error Log items, 15 min plan for next week")
            else:
                s1 = ("Test", th, f"WEEKLY REVISION TEST (timed): {kt}\nUse ICAI RTP / MTP questions of these chapters")
                s2 = ("Review", hrs - rh - th, "Evaluate → update Error Log & summary sheets → close pending revision items\nWEEKLY RECAP (last 45 min): summary sheets of every chapter revised Mon–Sat; re-solve this week's Error Log items; 15 min plan")
        elif d == C["buffer"]:
            s1 = ("All", 2, "Prepare for Mock 1: AS/section/SA lists, formula sheets and formats for every paper — everything for open-book-free recall")
            s2 = ("Rest", hrs - rh - 2, "Light: skim Error Log; sleep early. Mock 1 starts tomorrow at the real exam time (reading 1:45 pm, paper 2:00 pm)")
        elif d in MOCKDAY:
            n, ps = MOCKDAY[d]
            if len(ps) == 1:
                p = ps[0]
                s1 = (p, 5, f"MOCK {n} – {PAPERS[p][0]} ({PAPER_TIME})\nSit at 1:45 pm sharp, exam conditions. Source: {mock_src[n]}\nThen 2 h: mark strictly against ICAI answers, fill the Mock Tracker, log every lost mark in the Error Log")
                lst = C["mocks"][n]; idx = [x[0] for x in lst].index(d)
                if idx < len(lst) - 1:
                    nx = lst[idx + 1][1][0]
                    s2 = (nx, hrs - rh - 5, f"Morning: revise {PAPERS[nx][0]} for tomorrow's mock (summary sheets + Error Log)")
                else:
                    s2 = ("All", hrs - rh - 5, f"Morning: fix the weakest 3 topics found across Mock {n} – re-solve, re-read")
            else:
                a, b = ps
                h1 = round((hrs - rh) / 2 * 4) / 4
                s1 = (a, h1, f"MOCK {n} (morning, 9:45 am reading, 10:00 am–1:00 pm) – {PAPERS[a][0]}\nSource: {mock_src[n]}. Evaluate it tonight: 1.5 h")
                s2 = (b, hrs - rh - h1, f"MOCK {n} (real exam slot, 1:45 pm reading, 2:00–5:00 pm) – {PAPERS[b][0]}\nEvaluate tonight: 1.5 h; log every lost mark in the Error Log")
        elif d in FINAL:
            a, b = FINAL[d]
            s1 = a; s2 = (b[0], b[1] - rh, b[2]); hrs = s1[1] + s2[1] + rh
        elif d in EXAMS:
            p = EXAMS[d]
            s1 = (p, 3, f"EXAM – {PAPERS[p][0]} ({PAPER_TIME}). Morning: 1.5 h light skim of summary sheets only; reach the centre by 1:00 pm")
            ev = C.get("exam_evening", {}).get(d)
            if ev:
                s2 = ("All", 1.5, ev); hrs = 4.5
            else:
                s2 = ("Rest", 0, "No study after the paper. Do not discuss answers. Light look at the next paper only if fresh (max 1 h)"); hrs = 3
        rows.append({"date": d, "phase": ph, "hours": hrs, "s1": s1, "s2": s2, "keys": keys, "recap_h": rh})

    CHNAME = {f"{s} Ch {no}:": f"{s} Ch {no} {name}" for s in subs for no, name, _ in CH[s]}
    for i, r in enumerate(rows):
        if r["recap_h"] == 0: r["recap"] = "— (exam day: morning skim only)"; continue
        if i == 0: r["recap"] = "First day – no recap. Use 30 min to set up: Error Log notebook, one-page summary-sheet folder (one page per chapter), flashcard box"; continue
        pv = rows[i - 1]; pk = []
        for k in pv["keys"]:
            if k not in pk: pk.append(k)
        if pk: r["recap"] = "MORNING RECAP (active recall, closed book) of yesterday: " + "; ".join(CHNAME[k] for k in pk) + " → write the key provisions/formats/formulas from memory, then check against the summary sheet; re-do 2 flashcards you got wrong"
        elif pv["phase"].startswith("Mock"): r["recap"] = "MORNING RECAP: re-read yesterday's mock Error Log entries; 15 min formats/sections sheet of TODAY's paper"
        elif pv["s1"][0] == "Test": r["recap"] = "MORNING RECAP: Sunday test corrections + last week's Error Log items (closed book, then check)"
        elif pv["phase"] == "EXAM DAY": r["recap"] = "MORNING RECAP: nothing from yesterday's exam – do not discuss the paper; 30 min summary sheets of today's subject"
        else: r["recap"] = "MORNING RECAP: yesterday's summary sheets + Error Log (closed book, then check)"
    return C, subs, rows, EXAMS

# ------------------------------------------------------------------ workbook
def F(bold=False, size=10, color="000000", italic=False): return Font(name="Arial", bold=bold, size=size, color=color, italic=italic)
thin = Side(style="thin", color="BFBFBF"); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HDR = PatternFill("solid", fgColor="1F4E3D"); SUBF = PatternFill("solid", fgColor="D9EAD3"); INP = PatternFill("solid", fgColor="FFFF00")
WRAP = Alignment(wrap_text=True, vertical="top"); CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)
def hdr(ws, r, hs):
    for c, h in enumerate(hs, 1):
        x = ws.cell(row=r, column=c, value=h); x.font = F(True, 10, "FFFFFF"); x.fill = HDR; x.alignment = CEN; x.border = BORDER
def colL(n):
    s = ""
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

STRAT = {
 "AA": ("Descriptive + 30 MCQs, 3 h. Heavy on AS-based numericals and company accounts.", "Reading time: pick questions where the AS is clear to you. MCQs first (35–40 min), then descriptive by marks.", "Quote the AS number and paragraph logic; clean working notes; correct Schedule III and amalgamation formats.", "Treating AS as theory only; missing disclosure points; unbalanced consolidated balance sheets from skipped minority-interest workings."),
 "LAW": ("Descriptive + 30 case-scenario MCQs, 3 h. Company Law ~70, Other Laws ~30.", "MCQs 35 min; case questions in 4 steps: Provision → Facts → Analysis → Conclusion.", "Section numbers, time limits and numbers (days, %, members) quoted exactly; clear conclusions.", "Stories without provisions; confusing time limits; ignoring General Clauses/Interpretation/FEMA (easy 25–30 marks)."),
 "TAX": ("Income Tax 50 + GST 50; descriptive + MCQs, 3 h.", "Split the paper 90/90 min; within each, MCQs first then computation questions.", "Step-by-step computation with section references; correct regime choice (115BAC); GST case answers with provision + conclusion.", "Missing amendments; computation without working notes; mixing up exemptions vs zero-rating; ITC blocked-credit errors."),
 "CMA": ("Descriptive + 30 MCQs, 3 h. Largely numerical.", "MCQs 35 min; then highest-mark numericals first; keep formats standard.", "Variance reconciliation, process accounts with abnormal loss/gain, CVP decisions – presented with formulas stated.", "Arithmetic slips, wrong treatment of normal loss scrap, theory questions left blank."),
 "AUD": ("Descriptive + 30 case-scenario MCQs, 3 h. Fully theory.", "MCQs 35 min; case answers: relevant SA/provision → application to facts → conclusion.", "SA numbers with correct requirement; precise audit terminology; point-wise answers with headings.", "Generic answers without SA reference; vague reporting conclusions; skipping ethics schedules."),
 "FMSM": ("FM 50 (numerical) + SM 50 (theory/case), 3 h.", "FM first (95 min) while fresh, then SM (70 min) with 15 min buffer; MCQs inside each section first.", "FM: formulas stated, decision given clearly. SM: models applied to the case with headings (not definitions alone).", "Leaving SM answers too short; not concluding FM decisions (accept/reject); running out of time on SM."),
}

def workbook(pid, C, subs, rows, EXAMS):
    wb = Workbook()
    SF = {s: PatternFill("solid", fgColor=SUBJECTS[s][3]) for s in SUBJECTS}
    for p, (fc, fl) in PAPER_FILL.items(): SF[p] = PatternFill("solid", fgColor=fl)
    FC = {s: SUBJECTS[s][2] for s in SUBJECTS}; FC.update({p: v[0] for p, v in PAPER_FILL.items()})
    OTH = {"Test": "FFF2CC", "Review": "FFF2CC", "Rest": "EDEDED", "All": "FFF2CC"}
    PHF = {"Mock Exam 1": "F8CBAD", "Mock Exam 2": "F8CBAD", "Mock Exam 3": "F8CBAD", "EXAM DAY": "FF9999", "Revision 1": "C6E0B4", "Revision 2": "A9D08E", "Final Revision": "FFD966"}
    ex_dates = sorted(EXAMS)
    # Dashboard
    ws = wb.active; ws.title = "Dashboard"
    ws["A1"] = f"{C['title']} – January 2027 | Day-wise Study Timetable (Oct 2026 – Jan 2027)"; ws["A1"].font = F(True, 14, "1F4E3D")
    ws["A2"] = "ArivuPro Academy | ICAI New Scheme | Exam dates as per ICAI Announcement No. 13-CA(EXAM)/JANUARY/2027 dated 23-Sep-2026"; ws["A2"].font = F(italic=True, color="595959")
    blocks_txt = ", then ".join(" + ".join(SUBJECTS[s][0] for s in b) for b in C["blocks"])
    m = C["mocks"]; fr = lambda lst: f"{lst[0][0].strftime('%d %b')} – {lst[-1][0].strftime('%d %b')}"
    guide = [
     f"1. 'Daily Timetable' is the master sheet: every date from 1-Oct-2026 to {ex_dates[-1].strftime('%d-%b-%Y')} shows exactly 2 subjects, the chapter/unit to study and the hours for each. Subjects are taken TWO AT A TIME and finished before the next pair: {blocks_txt} — in Phase 1 and again inside each revision. HIGH-priority chapters (by exam marks) get 25% more time, LOW-priority 25% less; every task is tagged [HIGH]/[MEDIUM]/[LOW].",
     "2. Each evening mark 'Done?' (Y/N) and enter 'Actual Hrs'. Dashboard and Chapter Master update automatically.",
     "3. Recap system: MORNING RECAP (first 30–45 min, closed-book recall of yesterday's topics – column G), SESSION RECAP (last 15 min of each session – summary sheet + flashcards), WEEKLY RECAP (Sunday, after the weekly test).",
     f"4. Three full mock cycles at the real exam time: Mock 1 ({fr(m[1])}), Mock 2 ({fr(m[2])}), Mock 3 ({fr(m[3])}). Enter scores in 'Mock & Test Tracker'.",
     f"5. Revision 1 ({C['r1'][0].strftime('%d %b')} – {C['r1'][1].strftime('%d %b')}, ~42% of study time per chapter) and Revision 2 ({C['r2'][0].strftime('%d %b')} – {C['r2'][1].strftime('%d %b')}, ~30%, HIGH chapters first). The day before each paper is for that paper only.",
     "6. " + C["note"],
     "7. Yellow cells are for the student to fill in. Everything else is calculated – do not overwrite formulas.",
    ]
    for i, g in enumerate(guide):
        c = ws.cell(row=4 + i, column=1, value=g); c.font = F(); c.alignment = WRAP
        ws.merge_cells(start_row=4 + i, start_column=1, end_row=4 + i, end_column=10); ws.row_dimensions[4 + i].height = 42
    r = 12
    ws.cell(row=r, column=1, value="KEY INPUTS (exam dates CONFIRMED by ICAI)").font = F(True, 11); r += 1
    hdr(ws, r, ["Item", "Value", "Note"])
    inp = [(f"Exam date – {PAPERS[p][0]}", PAPERS[p][2], "CONFIRMED – 2:00–5:00 pm (3 h), reading time 1:45–2:00 pm") for p in C["papers"]]
    hh = C["hours"]
    inp += [("Study hours – October (Mon–Sat / Sun)", f"{hh[10][0]} / {hh[10][1]}", "Sunday = weekly test + review + weekly recap"),
            ("Study hours – November (Mon–Sat / Sun)", f"{hh[11][0]} / {hh[11][1]}", ""),
            ("Study hours – December & January (Mon–Sat / Sun)", f"{hh[12][0]} / {hh[12][1]}", "Mock and final-revision days run at full hours even on Sundays"),
            ("Starting point on 1 Oct", "Classes completed; self-study of ICAI Study Material from Chapter 1", "ArivuPro assumption"),
            ("Tax law applicable", "Income-tax Act, 1961 (+ CGST/IGST) as per ICAI's applicable Finance Act for Jan 2027", "ICAI: Income-tax Act, 2025 applies only from May 2027 exams onwards"),
            ("Source material", "ICAI Study Material, PYQs (last 6 attempts), RTP, MTP", "icai.org / BoS Knowledge Portal")]
    for i, (a, b, c) in enumerate(inp):
        rr = r + 1 + i
        ws.cell(row=rr, column=1, value=a); cb = ws.cell(row=rr, column=2, value=b); cb.font = F(color="0000FF"); cb.fill = INP
        if isinstance(b, dt.date): cb.number_format = "DD-MMM-YYYY"
        ws.cell(row=rr, column=3, value=c).font = F(italic=True, color="595959")
        for cc in range(1, 4): ws.cell(row=rr, column=cc).border = BORDER; ws.cell(row=rr, column=cc).alignment = WRAP
        ws.cell(row=rr, column=1).font = F()
    r = r + len(inp) + 2
    ws.cell(row=r, column=1, value="PHASE SUMMARY (calculated from Daily Timetable)").font = F(True, 11); r += 1
    hdr(ws, r, ["Phase", "Days", "Planned Hours", "Actual Hours", "Days Done", "% Days Done"])
    phs = []
    for x in rows:
        if x["phase"] not in phs: phs.append(x["phase"])
    f0 = r + 1
    for i, ph in enumerate(phs):
        rr = r + 1 + i
        ws.cell(row=rr, column=1, value=ph)
        ws.cell(row=rr, column=2, value=f"=COUNTIF('Daily Timetable'!$D:$D,A{rr})")
        ws.cell(row=rr, column=3, value=f"=SUMIF('Daily Timetable'!$D:$D,A{rr},'Daily Timetable'!$E:$E)")
        ws.cell(row=rr, column=4, value=f"=SUMIF('Daily Timetable'!$D:$D,A{rr},'Daily Timetable'!$O:$O)")
        ws.cell(row=rr, column=5, value=f"=COUNTIFS('Daily Timetable'!$D:$D,A{rr},'Daily Timetable'!$N:$N,\"Y\")")
        ws.cell(row=rr, column=6, value=f"=IF(B{rr}=0,0,E{rr}/B{rr})").number_format = "0%"
        for cc in range(1, 7): ws.cell(row=rr, column=cc).font = F(); ws.cell(row=rr, column=cc).border = BORDER
    rr = r + len(phs) + 1
    ws.cell(row=rr, column=1, value="TOTAL").font = F(True)
    for col in "BCDE": ws[f"{col}{rr}"] = f"=SUM({col}{f0}:{col}{rr-1})"; ws[f"{col}{rr}"].font = F(True); ws[f"{col}{rr}"].border = BORDER
    ws[f"F{rr}"] = f"=IF(B{rr}=0,0,E{rr}/B{rr})"; ws[f"F{rr}"].number_format = "0%"; ws[f"F{rr}"].border = BORDER
    r = rr + 2
    ws.cell(row=r, column=1, value="MONTHLY HOURS BY SUBJECT (calculated)").font = F(True, 11); r += 1
    hdr(ws, r, ["Month", "Planned Hours", "Actual Hours"] + [f"{s} hrs" for s in subs])
    for i, lbl in enumerate(["Oct-2026", "Nov-2026", "Dec-2026", "Jan-2027"]):
        rr = r + 1 + i
        ws.cell(row=rr, column=1, value=lbl)
        ws.cell(row=rr, column=2, value=f"=SUMIF('Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$E:$E)")
        ws.cell(row=rr, column=3, value=f"=SUMIF('Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$O:$O)")
        for j, s in enumerate(subs):
            ws.cell(row=rr, column=4 + j, value=f"=SUMIFS('Daily Timetable'!$I:$I,'Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$H:$H,\"{s}\")+SUMIFS('Daily Timetable'!$L:$L,'Daily Timetable'!$C:$C,A{rr},'Daily Timetable'!$K:$K,\"{s}\")")
        for cc in range(1, 4 + len(subs)): ws.cell(row=rr, column=cc).font = F(); ws.cell(row=rr, column=cc).border = BORDER
    ws.column_dimensions["A"].width = 46; ws.column_dimensions["B"].width = 30; ws.column_dimensions["C"].width = 40
    for j in range(4, 12): ws.column_dimensions[colL(j)].width = 12

    # Daily timetable
    wt = wb.create_sheet("Daily Timetable")
    wt["A1"] = "DAILY TIMETABLE – 2 subjects per day | Morning recap → Session 1 (+recap) → Session 2 (+recap) | Sunday = weekly test + weekly recap | Yellow = fill daily"; wt["A1"].font = F(True, 12, "1F4E3D")
    hdr(wt, 2, ["Date", "Day", "Month", "Phase", "Planned Hrs", "Recap Hrs", "Morning Recap – yesterday's topics", "Session 1 – Subject", "S1 Hrs", "Session 1 – What to study + session recap", "Session 2 – Subject", "S2 Hrs", "Session 2 – What to study + session recap", "Done? (Y/N)", "Actual Hrs", "Notes / Error log pointer"])
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True); wt.add_data_validation(dv)
    first = 3
    for i, x in enumerate(rows):
        rr = first + i; d = x["date"]; s1, s2 = x["s1"], x["s2"]
        vals = [d, d.strftime("%a"), d.strftime("%b-%Y"), x["phase"], x["hours"], x["recap_h"], x["recap"], s1[0], s1[1], s1[2], s2[0], s2[1], s2[2]]
        for cc, v in enumerate(vals, 1): wt.cell(row=rr, column=cc, value=v)
        wt.cell(row=rr, column=1).number_format = "DD-MMM-YY"
        for cc in range(1, 17):
            c = wt.cell(row=rr, column=cc); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
        for cc in (1, 2, 3, 5, 6, 9, 12, 14, 15): wt.cell(row=rr, column=cc).alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
        for cc in (6, 7): wt.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor="F2F2F2")
        for sc, tc in ((8, 10), (11, 13)):
            s = wt.cell(row=rr, column=sc).value
            fl = SF.get(s) or (PatternFill("solid", fgColor=OTH[s]) if s in OTH else None)
            if fl: wt.cell(row=rr, column=sc).fill = fl; wt.cell(row=rr, column=tc).fill = fl
            if s in FC: wt.cell(row=rr, column=sc).font = F(True, 9, FC[s])
        if x["phase"] in PHF: wt.cell(row=rr, column=4).fill = PatternFill("solid", fgColor=PHF[x["phase"]])
        if d.weekday() == 6: wt.cell(row=rr, column=2).fill = PatternFill("solid", fgColor="FFF2CC")
        for cc in (14, 15, 16): wt.cell(row=rr, column=cc).fill = INP
        dv.add(wt.cell(row=rr, column=14))
        est = max(len(str(s1[2])), len(str(s2[2]))) / 70 + max(str(s1[2]).count("\n"), str(s2[2]).count("\n")) + 1
        wt.row_dimensions[rr].height = max(45, min(160, 13 * est))
    last = first + len(rows) - 1
    wt.conditional_formatting.add(f"A{first}:A{last}", FormulaRule(formula=[f'$N{first}="Y"'], fill=PatternFill("solid", fgColor="92D050")))
    wt.freeze_panes = "E3"; wt.auto_filter.ref = f"A2:P{last}"
    for col, w in zip("ABCDEFGHIJKLMNOP", [10, 5, 9, 20, 7, 6, 34, 9, 6, 60, 9, 6, 60, 7, 7, 22]): wt.column_dimensions[col].width = w
    leg = "Legend: " + " | ".join(f"{s} = {SUBJECTS[s][1]}" for s in subs) + " | TAX = Paper 3 Taxation, FMSM = Paper 6 FM & SM (used on mock / final / exam days). PYQ = ICAI past-year questions; RTP/MTP = ICAI Revision Test Paper / Mock Test Paper. Planned Hrs = Recap Hrs + S1 Hrs + S2 Hrs."
    wt.cell(row=last + 2, column=1, value=leg).font = F(italic=True, size=9); wt.merge_cells(start_row=last + 2, start_column=1, end_row=last + 2, end_column=16); wt.row_dimensions[last + 2].height = 45

    # Chapter master
    wc = wb.create_sheet("Chapter Master")
    wc["A1"] = "CHAPTER MASTER – syllabus coverage, hours and progress (calculated from Daily Timetable)"; wc["A1"].font = F(True, 12, "1F4E3D")
    hdr(wc, 2, ["Paper / section", "Ch", "Chapter", "Units / sub-topics", "Approx marks", "Exam priority", "Study hrs (planned basis)", "Rev 1 hrs", "Rev 2 hrs", "Days scheduled", "Study first date", "Study last date", "Rev 1 date", "Rev 2 date", "Self-rating after Rev 2 (1–5)", "Weak? (auto)"])
    rr = 3
    for s in subs:
        for no, name, us in CH[s]:
            key = f"{s} Ch {no}:"
            vals = [SUBJECTS[s][1], no, name, "; ".join(u for u, _ in us if not u.startswith("PYQ")), IMP[s][no][0], IMP[s][no][1], round(sum(h for _, h in us) * PRIO[IMP[s][no][1]] * SW[s] * 2) / 2]
            for cc, v in enumerate(vals, 1): wc.cell(row=rr, column=cc, value=v)
            wc.cell(row=rr, column=6).fill = PatternFill("solid", fgColor={"High": "F4B084", "Medium": "FFE699", "Low": "E2EFDA"}[IMP[s][no][1]])
            wc.cell(row=rr, column=8, value=f"=ROUND(G{rr}*0.42*2,0)/2"); wc.cell(row=rr, column=9, value=f"=ROUND(G{rr}*0.3*2,0)/2")
            pat = f'"*{key}*"'
            wc.cell(row=rr, column=10, value=f"=COUNTIF('Daily Timetable'!$J:$J,{pat})+COUNTIF('Daily Timetable'!$M:$M,{pat})")
            A = f"'Daily Timetable'!$A${first}:$A${last}"; PH = f"'Daily Timetable'!$D${first}:$D${last}"
            cond = f"(ISNUMBER(SEARCH(\"{key}\",'Daily Timetable'!$J${first}:$J${last}))+ISNUMBER(SEARCH(\"{key}\",'Daily Timetable'!$M${first}:$M${last}))>0)"
            mn = f'SUMPRODUCT(MIN({A}+({cond}*({PH}="Phase 1 – Study (SM + PYQ)")=0)*100000))'
            wc.cell(row=rr, column=11, value=f'=IF({mn}>100000,"",{mn})').number_format = "DD-MMM"
            for col, phn in ((12, "Phase 1 – Study (SM + PYQ)"), (13, "Revision 1"), (14, "Revision 2")):
                wc.cell(row=rr, column=col, value=f'=IFERROR(1/(1/SUMPRODUCT(MAX({A}*{cond}*({PH}="{phn}")))),"")').number_format = "DD-MMM"
            wc.cell(row=rr, column=15).fill = INP
            wc.cell(row=rr, column=16, value=f'=IF(O{rr}="","",IF(O{rr}<=2,"WEAK – add to Final Revision",""))')
            for cc in range(1, 17):
                c = wc.cell(row=rr, column=cc); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
                if cc in (1, 3): c.fill = SF[s]
            rr += 1
    wc.cell(row=rr, column=3, value="TOTAL").font = F(True)
    for col in "GHIJ": wc[f"{col}{rr}"] = f"=SUM({col}3:{col}{rr-1})"; wc[f"{col}{rr}"].font = F(True)
    dv2 = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True); wc.add_data_validation(dv2); dv2.add(f"O3:O{rr-1}")
    wc.cell(row=rr + 2, column=1, value="Chapter grouping follows the ICAI Study Material content for the new scheme; verify chapter numbers against the latest ICAI SM edition. 'Approx marks' and hours are ArivuPro planning estimates (priority factor High ×1.25, Medium ×1.0, Low ×0.75); the timetable scales them to the hours available in each phase.").font = F(italic=True, size=9)
    wc.merge_cells(start_row=rr + 2, start_column=1, end_row=rr + 2, end_column=16); wc.row_dimensions[rr + 2].height = 30
    wc.freeze_panes = "D3"
    for col, w in zip("ABCDEFGHIJKLMNOP", [30, 5, 38, 56, 11, 10, 9, 8, 8, 9, 10, 10, 10, 10, 12, 22]): wc.column_dimensions[col].width = w

    # Mock tracker
    wm = wb.create_sheet("Mock & Test Tracker")
    wm["A1"] = "MOCK EXAMS & WEEKLY TESTS – score tracker (fill yellow cells)"; wm["A1"].font = F(True, 12, "1F4E3D")
    B = C["bench"]
    hdr(wm, 2, ["Mock", "Date", "Paper", "Max Marks", "Marks Scored", "%", f"Benchmark ({B}%)", "Gap", "Time taken (min)", "Top 3 mistakes / topics to fix"])
    rr = 3
    for n, lst in C["mocks"].items():
        st = rr
        for d, ps in lst:
            for p in ps:
                vals = [f"Mock {n}", d, PAPERS[p][0], 100]
                for cc, v in enumerate(vals, 1): wm.cell(row=rr, column=cc, value=v)
                wm.cell(row=rr, column=2).number_format = "DD-MMM-YY"
                wm.cell(row=rr, column=5).fill = INP; wm.cell(row=rr, column=9).fill = INP; wm.cell(row=rr, column=10).fill = INP
                wm.cell(row=rr, column=6, value=f'=IF(E{rr}="","",E{rr}/D{rr})').number_format = "0%"
                wm.cell(row=rr, column=7, value=f"=D{rr}*{B/100}")
                wm.cell(row=rr, column=8, value=f'=IF(E{rr}="","",E{rr}-G{rr})')
                for cc in range(1, 11): c = wm.cell(row=rr, column=cc); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
                wm.cell(row=rr, column=3).fill = SF.get(p, SUBF)
                rr += 1
        wm.cell(row=rr, column=1, value=f"Mock {n} – AGGREGATE")
        wm.cell(row=rr, column=4, value=f"=SUM(D{st}:D{rr-1})")
        wm.cell(row=rr, column=5, value=f'=IF(COUNT(E{st}:E{rr-1})=0,"",SUM(E{st}:E{rr-1}))')
        wm.cell(row=rr, column=6, value=f'=IF(E{rr}="","",E{rr}/D{rr})').number_format = "0%"
        wm.cell(row=rr, column=7, value=f"=D{rr}*{B/100}")
        wm.cell(row=rr, column=8, value=f'=IF(E{rr}="","",E{rr}-G{rr})')
        wm.cell(row=rr, column=10, value=f'=IF(E{rr}="","",IF(COUNTIF(E{st}:E{rr-1},"<40")>0,"A paper is below 40 – fix it first",IF(F{rr}>={B/100},"On target",IF(F{rr}>=0.6,"Close – push the weakest paper","Pass zone – re-plan revision time with faculty"))))')
        for cc in range(1, 11): c = wm.cell(row=rr, column=cc); c.font = F(True, 9); c.border = BORDER; c.fill = SUBF; c.alignment = WRAP
        rr += 2
    wm.cell(row=rr, column=1, value="WEEKLY SUNDAY TESTS").font = F(True, 11); rr += 1
    hdr(wm, rr, ["Week", "Test date", "Chapters covered", "Max Marks", "Marks Scored", "%", "Target (75%)", "Gap", "Time taken (min)", "Errors logged / action"]); rr += 1
    wk = 1
    for x in rows:
        if x["s1"][0] == "Test":
            wm.cell(row=rr, column=1, value=f"Week {wk}"); wm.cell(row=rr, column=2, value=x["date"]).number_format = "DD-MMM-YY"
            wm.cell(row=rr, column=3, value=x["s1"][2].split("\n")[0].split(": ", 1)[-1])
            wm.cell(row=rr, column=4, value=100).fill = INP; wm.cell(row=rr, column=5).fill = INP
            wm.cell(row=rr, column=6, value=f'=IF(E{rr}="","",E{rr}/D{rr})').number_format = "0%"
            wm.cell(row=rr, column=7, value=f"=D{rr}*0.75"); wm.cell(row=rr, column=8, value=f'=IF(E{rr}="","",E{rr}-G{rr})')
            wm.cell(row=rr, column=9).fill = INP; wm.cell(row=rr, column=10).fill = INP
            for cc in range(1, 11): c = wm.cell(row=rr, column=cc); c.font = F(size=9); c.border = BORDER; c.alignment = WRAP
            wm.row_dimensions[rr].height = 30; rr += 1; wk += 1
    for col, w in zip("ABCDEFGHIJ", [20, 11, 44, 9, 10, 7, 13, 9, 10, 40]): wm.column_dimensions[col].width = w
    wm.freeze_panes = "A3"

    # Strategy
    wr = wb.create_sheet("Exam Strategy")
    wr["A1"] = "EXAM STRATEGY – paper-wise"; wr["A1"].font = F(True, 12, "1F4E3D")
    hdr(wr, 2, ["Paper", "Pattern", "Target", "Time plan in the hall", "Where marks are won", "Where marks are lost"])
    for i, p in enumerate(C["papers"]):
        a, b, c, d = STRAT[p]
        for cc, v in enumerate([PAPERS[p][0], a, f"{B}+ / 100", b, c, d], 1):
            x = wr.cell(row=3 + i, column=cc, value=v); x.font = F(size=9); x.border = BORDER; x.alignment = WRAP
        wr.row_dimensions[3 + i].height = 90
    rr = 4 + len(C["papers"])
    tg = [("Pass criteria", "40 in each paper AND 50% aggregate in the group. Exemption: 60+ in a paper when the group is not cleared."),
          ("Target of this plan", C["bench_label"] + ". " + C["note"]),
          ("Mock benchmark", f"Mock 1 ≥ 55%, Mock 2 ≥ 65%, Mock 3 ≥ {B}% aggregate; no paper below 50 in Mock 3.")]
    for a, b in tg:
        wr.cell(row=rr, column=1, value=a).font = F(True, 9); wr.cell(row=rr, column=2, value=b).font = F(size=9)
        wr.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=6); wr.row_dimensions[rr].height = 36
        for cc in range(1, 7): wr.cell(row=rr, column=cc).border = BORDER; wr.cell(row=rr, column=cc).alignment = WRAP
        rr += 1
    rr += 1
    wr.cell(row=rr, column=1, value="NON-NEGOTIABLE HABITS").font = F(True, 11); rr += 1
    for h in ["1. Error Log: every mistake from PYQs, tests and mocks, with the correct approach. Re-read before every test.",
              "2. Recap discipline: morning recap → session recap → weekly recap. Skipping recaps is the commonest reason Phase 1 is forgotten by December.",
              "3. One-page summary per chapter, built during session recaps: AS/section/SA lists, formats, formula sheets. Revision 2 and the day-before skim use only these.",
              "4. Numericals are studied by WRITING. Solve every ICAI illustration with the book closed, then compare.",
              "5. 30% of every paper is case-scenario MCQs: do 10 MCQs per chapter in Revision 1 and timed MCQ sets in Revision 2.",
              "6. Mocks at the real time: 1:45 pm reading, 2:00–5:00 pm writing. Never pause a mock.",
              "7. Backlog goes into Sunday's review block — never pushes Monday's plan or a mock date.",
              "8. Sleep 7–7.5 h. A 6-paper attempt is a stamina event."]:
        c = wr.cell(row=rr, column=1, value=h); c.font = F(size=9); c.alignment = WRAP
        wr.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6); wr.row_dimensions[rr].height = 28; rr += 1
    for col, w in zip("ABCDEF", [34, 36, 12, 44, 44, 44]): wr.column_dimensions[col].width = w
    for w in wb.worksheets:
        w.page_setup.orientation = "landscape"; w.page_setup.fitToWidth = 1; w.page_setup.fitToHeight = 0; w.sheet_properties.pageSetUpPr.fitToPage = True
    wt.print_title_rows = "2:2"
    path = "../excel/" + C["file"]; wb.save(path); return path

def export(pid, C, subs, rows):
    KEYRE = re.compile(r"^([A-Z]+) Ch (\d+):")
    days = []
    for x in rows:
        sess = []
        for sx in (x["s1"], x["s2"]):
            if not sx[0] and not sx[2]: continue
            items = []
            for line in str(sx[2]).split("\n"):
                line = line.strip()
                if not line: continue
                mm = KEYRE.match(line); hm = re.search(r"\[([\d.]+)h\]\s*$", line)
                items.append({"t": re.sub(r"\s*\[[\d.]+h\]\s*$", "", line), "k": f"{mm.group(1)}-{mm.group(2)}" if mm else None, "h": float(hm.group(1)) if hm else None})
            sess.append({"s": sx[0], "h": sx[1], "items": items})
        days.append({"d": x["date"].isoformat(), "p": x["phase"], "h": x["hours"], "rh": x["recap_h"], "recap": x["recap"], "sess": sess})
    chap = [{"k": f"{s}-{no}", "s": s, "no": no, "n": name, "marks": IMP[s][no][0], "pr": IMP[s][no][1]} for s in subs for no, name, _ in CH[s]]
    cfg = {"id": pid, "title": C["title"], "short": C["short"], "bench": C["bench"], "benchLabel": C["bench_label"], "note": C["note"],
           "subjects": subs, "papers": [{"code": p, "name": PAPERS[p][0], "date": PAPERS[p][2].isoformat(), "subs": PAPERS[p][1]} for p in C["papers"]],
           "mocks": {str(n): [[d.isoformat(), ps] for d, ps in lst] for n, lst in C["mocks"].items()},
           "special": C["special"], "blocks": C["blocks"]}
    json.dump({"cfg": cfg, "days": days, "chapters": chap}, open(f"plan_{pid}.json", "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))

if __name__ == "__main__":
    for pid in sys.argv[1:] or ["g1", "g2", "both"]:
        C, subs, rows, EX = build(pid)
        p = workbook(pid, C, subs, rows, EX); export(pid, C, subs, rows)
        from collections import Counter
        hrs = Counter()
        for x in rows:
            for s in (x["s1"], x["s2"]):
                if x["phase"].startswith(("Phase", "Revision")): hrs[(x["phase"][:10], s[0])] += s[1]
        print(pid, p, len(rows), "days", sum(x["hours"] for x in rows), "h")
        print("  ", {k: v for k, v in sorted(hrs.items())})
