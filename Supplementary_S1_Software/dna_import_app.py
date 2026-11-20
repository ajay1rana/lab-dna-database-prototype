#!/usr/bin/env python3
"""
DNA Database Import App
=======================
Reads one or more Excel files (vertical "PI" profile sheets or horizontal tables, any kit,
any column order), maps case details and STR markers to the right fields, fills Kit / State /
District automatically from the rules stored INSIDE your database workbook, checks every
profile, and appends the accepted rows directly to the 'DNA Database' sheet.

All rules (header aliases, cell addresses, kit rules, location table, allele limits) are read
from the database workbook itself (sheets Import_Map, Kit_Rules, Settings, Loci_Reference), so
edits you make there are picked up by this app.

Run:   python dna_import_app.py            (opens the window)
CLI:   python dna_import_app.py --db DB.xlsx --files a.xlsx b.xlsx [--add] [--review-too]
                                [--set "Profile Category=Convicted" --set "Analyst=Priya"]
Requires: Python 3.9+ and openpyxl  (pip install openpyxl)
CLOSE the database in Excel before adding rows (Excel locks the file).
"""
import math, re, os, sys, shutil, datetime, argparse, threading, tempfile, zipfile, html, warnings, secrets, sqlite3
warnings.filterwarnings("ignore")
from openpyxl import load_workbook

DB_SHEET = "DNA Database"
TODAY = datetime.date.today()

# ----------------------------------------------------------------------------- helpers
def norm(x):
    return re.sub(r"[^a-z0-9]", "", str("" if x is None else x).lower())

def clean(v):
    if v is None: return ""
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def fmt_num(f):
    return int(f) if float(f).is_integer() else round(float(f), 2)

def parse_date(v):
    if isinstance(v, datetime.datetime): return v
    if isinstance(v, datetime.date): return datetime.datetime(v.year, v.month, v.day)
    s = clean(v)
    if not s: return None
    for f in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y", "%d.%m.%Y", "%d-%m-%y", "%d/%m/%y", "%Y/%m/%d"):
        try: return datetime.datetime.strptime(s, f)
        except ValueError: pass
    return None

INDIAN_STATES = ["Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chhattisgarh", "Goa", "Gujarat", "Haryana",
    "Himachal Pradesh", "Jharkhand", "Karnataka", "Kerala", "Madhya Pradesh", "Maharashtra", "Manipur", "Meghalaya", "Mizoram",
    "Nagaland", "Odisha", "Punjab", "Rajasthan", "Sikkim", "Tamil Nadu", "Telangana", "Tripura", "Uttar Pradesh", "Uttarakhand",
    "West Bengal", "Andaman and Nicobar Islands", "Chandigarh", "Dadra and Nagar Haveli and Daman and Diu", "Delhi (NCT)",
    "Jammu and Kashmir", "Ladakh", "Lakshadweep", "Puducherry"]
KERALA_DISTRICTS = ["Thiruvananthapuram", "Kollam", "Pathanamthitta", "Alappuzha", "Kottayam", "Idukki", "Ernakulam", "Thrissur",
    "Palakkad", "Malappuram", "Kozhikode", "Wayanad", "Kannur", "Kasaragod"]
TELANGANA_DISTRICTS = ["Adilabad", "Bhadradri Kothagudem", "Hanumakonda", "Hyderabad", "Jagtial", "Jangaon", "Jayashankar Bhupalpally",
    "Jogulamba Gadwal", "Kamareddy", "Karimnagar", "Khammam", "Kumuram Bheem Asifabad", "Mahabubabad", "Mahabubnagar", "Mancherial",
    "Medak", "Medchal-Malkajgiri", "Mulugu", "Nagarkurnool", "Nalgonda", "Narayanpet", "Nirmal", "Nizamabad", "Peddapalli",
    "Rajanna Sircilla", "Rangareddy", "Sangareddy", "Siddipet", "Suryapet", "Vikarabad", "Wanaparthy", "Warangal", "Yadadri Bhuvanagiri"]
PUDUCHERRY_DISTRICTS = ["Puducherry", "Karaikal", "Mahe", "Yanam"]
ANDAMAN_DISTRICTS = ["North and Middle Andaman", "South Andaman", "Nicobar"]
DISTRICT_STATE = {**{d.lower(): "Andaman and Nicobar Islands" for d in ANDAMAN_DISTRICTS}, **{d.lower(): "Puducherry" for d in PUDUCHERRY_DISTRICTS}, **{d.lower(): "Kerala" for d in KERALA_DISTRICTS}, **{d.lower(): "Telangana" for d in TELANGANA_DISTRICTS}}

def word_in(name, text):
    name = re.sub(r"\s*\(.*?\)", "", name).strip().lower()
    return bool(name) and re.search(r"(?<![a-z0-9])" + re.escape(name) + r"(?![a-z0-9])", text) is not None

DATE_HEADERS = {"Collection Date", "Date Received", "Analysis Date", "Review Date", "Entry Date", "Requisition Date", "FIR/Case Date"}
NUM_HEADERS = {"Y-indel", "No. of Contributors"}

# ----------------------------------------------------------------------------- rules from the workbook
class Rules:
    def __init__(self, wb):
        self.case = []        # dicts: header, aliases, default, addr
        self.loci = []        # dicts: name, names[]
        self.limits = {}      # locus -> (min,max)
        self.core = []
        self.kits = []        # (name, {marker: 'R'/'X'})
        self.markers = []
        self.locations = []   # (keyword, state, district)
        self.lists = {}
        self.mindup = 10
        self._read_map(wb["Import_Map"])
        self._read_kits(wb["Kit_Rules"])
        self._read_settings(wb["Settings"])
        self._read_limits(wb["Loci_Reference"])

    def _read_map(self, ws):
        rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=11, values_only=True))
        cs = next(i for i, r in enumerate(rows) if r[0] == "Case / sample field") + 1
        k = cs
        while k < len(rows) and rows[k][0]:
            r = rows[k]
            al = [norm(x) for x in r[1:7] if x not in (None, "")]
            if r[0] == "Lab Case No.":
                al += [a for a in ("reportnumber", "reportno", "labreportno", "labreportnumber") if a not in al]
            self.case.append({"header": r[0], "aliases": al,
                              "default": r[9], "addr": (str(r[10]).strip() if r[10] else None)})
            k += 1
        ls = next(i for i, r in enumerate(rows) if r[0] == "Marker") + 1
        k = ls
        while k < len(rows) and rows[k][0]:
            r = rows[k]
            self.loci.append({"name": r[0], "names": [norm(x) for x in r[1:4] if x not in (None, "")]})
            k += 1

    def _read_kits(self, ws):
        hdr = [c.value for c in ws[4]]
        self.markers = [h for h in hdr[1:] if h]
        for r in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
            if not r[0]: continue
            rules = {m: r[1 + i] for i, m in enumerate(self.markers) if r[1 + i] in ("R", "X")}
            if any(v == "R" for v in rules.values()): self.kits.append((r[0], rules))

    def _read_settings(self, ws):
        def col(letter, start=5):
            out = []
            r = start
            while ws[f"{letter}{r}"].value not in (None, ""):
                out.append(ws[f"{letter}{r}"].value); r += 1
            return out
        self.lists = {"Profile Category": col("A"), "Case Type": col("E"), "Sample Type": col("F"), "Profile Type": col("G"),
                      "Record Status": col("I"), "Kit/Method": col("J"), "Instrument": col("L"), "Analysis Software": col("M"),
                      "State": col("N")}
        for r in range(5, 45):
            kw, st_, di = ws[f"S{r}"].value, ws[f"T{r}"].value, ws[f"U{r}"].value
            if kw: self.locations.append((str(kw).lower(), st_, di))
        try: self.mindup = int(ws["Q5"].value)
        except Exception: pass
        self.kit_cover = {}
        r = 5
        while ws[f"J{r}"].value:
            try: self.kit_cover[ws[f"J{r}"].value] = int(ws[f"K{r}"].value)
            except (TypeError, ValueError): pass
            r += 1
        self.cat_index = {}; self.cat_min = {}
        for r in range(5, 40):
            if ws[f"A{r}"].value:
                self.cat_index[ws[f"A{r}"].value] = ws[f"B{r}"].value
                try: self.cat_min[ws[f"A{r}"].value] = int(ws[f"C{r}"].value)
                except (TypeError, ValueError): pass
        self.code_lab, self.code_keys = None, None
        try:
            if ws["P8"].value and "aboratory code" in str(ws["P8"].value):
                self.code_lab = str(ws["Q8"].value)
                self.code_keys = [int(ws[f"Q{r}"].value) for r in range(9, 15)]
        except Exception:
            self.code_lab, self.code_keys = None, None
        self.states = self.lists.get("State") or list(INDIAN_STATES)
        extra_d = []
        r = 5
        while ws[f"W{r}"].value not in (None, ""):
            extra_d.append(str(ws[f"W{r}"].value)); r += 1
        self.districts = list(dict.fromkeys(extra_d + KERALA_DISTRICTS + TELANGANA_DISTRICTS + PUDUCHERRY_DISTRICTS + ANDAMAN_DISTRICTS))

    def _read_limits(self, ws):
        r = 5
        while ws[f"B{r}"].value:
            name, mn, mx = ws[f"B{r}"].value, ws[f"G{r}"].value, ws[f"H{r}"].value
            if isinstance(mn, (int, float)) and isinstance(mx, (int, float)): self.limits[name] = (mn, mx)
            if ws[f"E{r}"].value == "Y": self.core.append(name)
            r += 1

# ----------------------------------------------------------------------------- allele parsing
def split_combined(s, amel):
    s = s.upper().strip()
    for a in (", ", "/", ";", "|", "-", "&", "+"): s = s.replace(a, ",")
    s = s.replace(" ", ",")
    while ",," in s: s = s.replace(",,", ",")
    if amel and "," not in s and len(s) == 2: return s[0], s[1], False
    parts = s.split(",")
    first = parts[0] if parts else ""
    second = parts[1] if len(parts) > 1 else ""
    extra = any(p for p in parts[2:])
    return first, second, extra

def parse_locus(locus, mode, v1, v2, homozygote):
    """returns (t1, t2, flag)"""
    amel = locus == "Amelogenin"
    t1 = t2 = ""; flag = ""
    if mode == "combined":
        t1, t2, extra = split_combined(clean(v1), amel)
        if t2 == "" and t1 != "" and homozygote: t2 = t1
        if extra: return t1, t2, "more than 2 alleles"
    elif mode == "2col":
        t1, t2 = clean(v1).upper(), clean(v2).upper()
    else:
        return "", "", ""
    if t1 == "" and t2 == "": return "", "", ""
    if t1 == "" or t2 == "": return t1, t2, "only one allele"
    if amel:
        if t1 not in ("X", "Y") or t2 not in ("X", "Y"): return t1, t2, "non-numeric allele"
    else:
        try: float(t1); float(t2)
        except ValueError: return t1, t2, "non-numeric allele"
    return t1, t2, ""

# ----------------------------------------------------------------------------- reading a source file
class Record:
    def __init__(self, source, row):
        self.source, self.row = source, row
        self.case = {}      # header -> raw value
        self.alleles = {}   # locus -> (t1,t2,flag)
        self.layout = ""
        self.warnings = []
        self.alltext = ""     # every text cell of the sheet/row (used to find state / district anywhere)

def first_idx(lst, names):
    for n in names:
        if n in lst: return lst.index(n)
    return None

LABELS = {"referencenumber", "date", "from", "casenumber", "us", "ps", "reportnumber", "numberofpagesofreport"}
ANCHORS = {"reportnumber": ("Lab Case No.", {"date": "_Report Date"}),
           "referencenumber": ("Requisition Ref. No.", {"date": "Requisition Date", "from": "Forwarding Authority"}),
           "casenumber": ("Case/FIR No.", {"date": "FIR/Case Date", "us": "Sections of Law (U/S)", "ps": "Police Station"})}

def vertical_case_block(ws):
    """Find 'Reference Number / Date / From' and 'Case Number / Date / U/S / PS' labels anywhere (right of the marker
    columns) and take the first non-label value below each label. Works even if the block moves or labels repeat."""
    maxr, maxc = min(ws.max_row, 80), min(ws.max_column, 20)
    def val(r, c): return ws.cell(row=r, column=c).value
    def below(r, c):
        for rr in range(r + 1, min(r + 12, ws.max_row) + 1):
            v = val(rr, c)
            if v in (None, ""): continue
            if isinstance(v, str) and norm(v) in LABELS: continue
            return v, ws.cell(row=rr, column=c)
        return None, None
    out, cells = {}, {}
    for r in range(1, maxr + 1):
        for c in range(3, maxc + 1):
            v = val(r, c)
            if not (isinstance(v, str) and norm(v) in ANCHORS): continue
            hdr, sibs = ANCHORS[norm(v)]
            if hdr in out: continue
            x, cell = below(r, c)
            if x is not None: out[hdr] = x; cells[hdr] = cell
            for cc in range(c + 1, maxc + 1):
                sv = val(r, cc)
                if isinstance(sv, str) and norm(sv) in sibs:
                    h2 = sibs[norm(sv)]
                    if h2 not in out:
                        x, cell = below(r, cc)
                        if x is not None: out[h2] = x; cells[h2] = cell
    return out, cells

def date_to_text(dt, fmt):
    fmt = (fmt or "").lower()
    if "mmm" in fmt and "d" not in fmt.replace("mmm", ""): return f"{dt.month}/{dt.year}"
    return f"{dt.day}/{dt.month}/{dt.year}"

def read_source(path, rules, homozygote=False):
    """Reads EVERY visible sheet of the workbook: a vertical sheet = one profile, a horizontal table = one profile per row."""
    wb = load_workbook(path, data_only=True)
    sheets = [w for w in wb.worksheets if w.sheet_state == "visible"]
    recs = []
    for ws in sheets:
        label = os.path.basename(path) + (f"[{ws.title}]" if len(sheets) > 1 else "")
        recs += parse_sheet(ws, label, rules, homozygote)
    return recs

def parse_sheet(ws, fname, rules, homozygote=False):
    maxr, maxc = ws.max_row, ws.max_column
    colA = [norm(ws.cell(row=r, column=1).value) for r in range(1, min(maxr, 150) + 1)]
    hits = sum(1 for L in rules.loci if first_idx(colA, L["names"]) is not None)
    recs = []
    if hits >= 5:                                   # vertical profile sheet
        rec = Record(fname, 1); rec.layout = "vertical"
        used_rows = set()
        for L in rules.loci:
            i = first_idx(colA, L["names"])
            if i is not None: used_rows.add(i)
            v = ws.cell(row=i + 1, column=2).value if i is not None else None
            rec.alleles[L["name"]] = parse_locus(L["name"], "combined" if i is not None else "none", v, None, homozygote)
        block, bcells = vertical_case_block(ws)
        rdate = block.pop("_Report Date", None)
        if block.get("Lab Case No."):                         # Report Number + its date -> Lab Case No.
            lab = clean(block["Lab Case No."])
            if rdate not in (None, ""):
                pdt = parse_date(rdate)
                lab += f" dt. {pdt:%d-%m-%Y}" if pdt else f" dt. {clean(rdate)}"
            block["Lab Case No."] = lab
        for c in rules.case:
            val = block.get(c["header"])
            if val is None and c["addr"]:
                try: val = ws[c["addr"]].value
                except Exception: val = None
            if val in (None, ""):
                i = first_idx(colA, c["aliases"])
                if i is not None: val = ws.cell(row=i + 1, column=2).value
            if isinstance(val, str) and norm(val) in LABELS: val = None      # a label, not a value
            if val not in (None, ""):
                cell = bcells.get(c["header"])
                if c["header"] == "Case/FIR No." and isinstance(val, (datetime.datetime, datetime.date)):
                    dt = val; txt = date_to_text(dt, cell.number_format if cell is not None else "")
                    rec.warnings.append(f"Case/FIR No. was turned into a date by Excel (shown as {dt:%b-%y}); stored as '{txt}' - verify against the original")
                    val = txt
                rec.case[c["header"]] = val
        # ---- extract everything else, from anywhere on the sheet
        pages = ""
        alias_map = {}
        for c in rules.case:
            for a in c["aliases"]: alias_map.setdefault(a, c["header"])
        known = set(alias_map) | {n for L in rules.loci for n in L["names"]}
        for r in range(1, min(ws.max_row, 150) + 1):
            for cc in range(1, min(ws.max_column, 30) + 1):
                cv = ws.cell(row=r, column=cc).value
                if isinstance(cv, str):
                    rec.alltext += " | " + cv
                    if norm(cv).startswith("numberofpages") and not pages:
                        pv = ws.cell(row=r, column=cc + 1).value
                        if pv in (None, ""): pv = ws.cell(row=r + 1, column=cc).value
                        if pv not in (None, ""): pages = clean(pv)
                    h = alias_map.get(norm(cv))
                    if h and h not in rec.case and norm(cv) not in LABELS:        # a labelled value anywhere: right, else below
                        nv = ws.cell(row=r, column=cc + 1).value
                        if nv in (None, "") or (isinstance(nv, str) and norm(nv) in alias_map):
                            nv = ws.cell(row=r + 1, column=cc).value
                        if nv not in (None, "") and not (isinstance(nv, str) and norm(nv) in alias_map): rec.case[h] = nv
        ystr, xstr, extras = [], [], []
        for r in range(1, min(ws.max_row, 150) + 1):
            a, b = ws.cell(row=r, column=1).value, ws.cell(row=r, column=2).value
            if a in (None, "") or b in (None, "") or (r - 1) in used_rows: continue
            n = norm(a)
            if n in known or n in LABELS: continue
            if n.startswith(("dys", "dyf", "ygata")): ystr.append(f"{clean(a)}={clean(b)}")
            elif n.startswith("dxs"): xstr.append(f"{clean(a)}={clean(b)}")
            else: extras.append(f"{clean(a)}: {clean(b)}")
        if pages: extras.insert(0, f"Report pages: {pages}")
        if ystr and "DYS/Y-STR" not in rec.case: rec.case["DYS/Y-STR"] = "; ".join(ystr)
        if xstr and "X-STR" not in rec.case: rec.case["X-STR"] = "; ".join(xstr)
        if extras:
            old = clean(rec.case.get("Remarks"))
            rec.case["Remarks"] = (old + "; " if old else "") + "From file: " + "; ".join(extras)
        recs.append(rec)
    else:                                           # horizontal table: headers in row 1
        hdr = [norm(ws.cell(row=1, column=c).value) for c in range(1, maxc + 1)]
        plan = {}
        for L in rules.loci:
            names = L["names"]
            comb = first_idx(hdr, names)
            comb2 = None
            if comb is not None:
                nm = hdr[comb]
                if nm in hdr[comb + 1:]: comb2 = comb + 1 + hdr[comb + 1:].index(nm)
            a1 = first_idx(hdr, [n + s for n in names for s in ("1", "allele1", "a1")] + ["allele1" + n for n in names] + ["a1" + n for n in names])
            a2 = first_idx(hdr, [n + s for n in names for s in ("2", "allele2", "a2")] + ["allele2" + n for n in names] + ["a2" + n for n in names])
            if a1 is not None and a2 is not None: plan[L["name"]] = ("2col", a1, a2)
            elif comb is not None and comb2 is not None: plan[L["name"]] = ("2col", comb, comb2)
            elif comb is not None: plan[L["name"]] = ("combined", comb, None)
            else: plan[L["name"]] = ("none", None, None)
        if all(m == "none" for m, _a, _b in plan.values()): return []      # no STR markers on this sheet
        cplan = {c["header"]: first_idx(hdr, c["aliases"]) for c in rules.case}
        for r in range(2, maxr + 1):
            row = [ws.cell(row=r, column=c).value for c in range(1, maxc + 1)]
            if not any(clean(v) for v in row): continue
            rec = Record(fname, r); rec.layout = "horizontal"
            for name, (mode, i1, i2) in plan.items():
                rec.alleles[name] = parse_locus(name, mode, row[i1] if i1 is not None else None,
                                                row[i2] if i2 is not None else None, homozygote)
            for h, i in cplan.items():
                if i is not None and clean(row[i]) != "": rec.case[h] = row[i]
            rec.alltext = " | ".join(str(v) for v in row if isinstance(v, str))
            ys = [f"{clean(ws.cell(row=1, column=i + 1).value)}={clean(row[i])}" for i, hname in enumerate(hdr)
                  if hname.startswith("dys") and hname not in ("dysystr",) and clean(row[i]) != ""]
            if ys and "DYS/Y-STR" not in rec.case: rec.case["DYS/Y-STR"] = "; ".join(ys)
            recs.append(rec)
    return recs


# ----------------------------------------------------------------------------- identification codes
CODE_ALPHABET = "0123456789ABCDEFGHJKMNPQRSTVWXYZ"          # Crockford base-32 (no I, L, O, U)
INDEX_CODES = {"Offender / Suspect": "OFS", "Forensic (casework)": "FOR", "Missing persons / UHR": "MPU",
               "Reference / Elimination": "REF", "Other": "OTH"}
INDEX_ARRAY = '{"OFS","FOR","MPU","REF","OTH"}'
GROUP_ARRAY = '{"Offender / Suspect","Forensic (casework)","Missing persons / UHR","Reference / Elimination","Other"}'

def code_serial(pid):
    t = str(pid).split("-", 1)
    return int(t[1]) if len(t) == 2 and t[1].strip().isdigit() else None

def make_codes(pid, idx_group, d_entry, d_analysis, keys, lab):
    """Returns (identifiable_code, deidentified_code). Must match the workbook formulas exactly."""
    n = code_serial(pid)
    if n is None or not keys: return "", ""
    A1, B1, A2, B2, A3, B3 = keys; M = 1 << 30
    p1 = (n * A1 + B1) % M
    p2 = ((p1 ^ (p1 >> 15)) * A2 + B2) % M
    p3 = ((p2 ^ (p2 >> 13)) * A3 + B3) % M
    v = [(p3 >> (5 * (5 - i))) & 31 for i in range(6)]
    chk = (1 * v[0] + 3 * v[1] + 5 * v[2] + 7 * v[3] + 9 * v[4] + 11 * v[5]) % 32
    A = CODE_ALPHABET
    did = "DI-" + "".join(A[x] for x in v[:4]) + "-" + A[v[4]] + A[v[5]] + A[chk]
    d = d_entry if isinstance(d_entry, (datetime.datetime, datetime.date)) else d_analysis
    yr = d.year if isinstance(d, (datetime.datetime, datetime.date)) else 0
    s = f"{yr:04d}{n:06d}"
    c = sum(int(s[i]) * (i + 2) for i in range(10)) % 11
    iden = f"{lab}-{INDEX_CODES.get(idx_group, 'UNK')}-{yr:04d}-{n:06d}-{'X' if c == 10 else c}"
    return iden, did

def new_code_keys():
    r = secrets.SystemRandom()
    odd = lambda hi: r.randrange(1 << 12, hi) | 1
    return [odd(1 << 30), r.randrange(1 << 30), odd(1 << 22), r.randrange(1 << 30), odd(1 << 22), r.randrange(1 << 30)]

# ----------------------------------------------------------------------------- the database
class Database:
    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(path)
        self.ws = self.wb[DB_SHEET]
        self.rules = Rules(self.wb)
        self.hmap = {self.ws.cell(row=2, column=c).value: c for c in range(1, self.ws.max_column + 1)
                     if self.ws.cell(row=2, column=c).value}
        t = self.ws.tables.get("DNADatabase")
        self.last_row = int(re.search(r"(\d+)$", t.ref).group(1)) if t else self.ws.max_row
        self.ids = set(); self.keys = {}; self.refsig = {}; self.refs = {}
        self.maxnum = 0
        for r in range(3, self.last_row + 1):
            pid = self.ws.cell(row=r, column=1).value
            if pid in (None, ""): continue
            pid = str(pid); self.ids.add(pid)
            m = re.match(r"DNA-(\d+)$", pid)
            if m: self.maxnum = max(self.maxnum, int(m.group(1)))
            k = self.key_from_db(r)
            if k: self.keys.setdefault(k, pid)
            fir = norm(self.ws.cell(row=r, column=self.hmap.get("Case/FIR No.", 1)).value) if "Case/FIR No." in self.hmap else ""
            ps = norm(self.ws.cell(row=r, column=self.hmap.get("Police Station", 1)).value) if "Police Station" in self.hmap else ""
            parts = []
            for L in self.rules.loci:
                a, b = (self.ws.cell(row=r, column=self.hmap.get(f"{L['name']} ({i})", 0) or 1).value for i in (1, 2))
                if a not in (None, "") and b not in (None, ""):
                    parts.append(f"{L['name']}:{fmt_num(a) if isinstance(a, (int, float)) else a},{fmt_num(b) if isinstance(b, (int, float)) else b}")
            sig = "|".join(parts)
            if sig and (fir or ps): self.refsig.setdefault((fir, ps, sig), pid)
            if fir or ps: self.refs.setdefault((fir, ps), pid)

    def key_from_db(self, r):
        parts, typed = [], 0
        for loc in self.rules.core:
            a, b = (self.ws.cell(row=r, column=self.hmap.get(f"{loc} ({i})", 0) or 1).value for i in (1, 2))
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                typed += 1; parts.append(f"{fmt_num(a)},{fmt_num(b)};")
            else: parts.append(",;")
        return "".join(parts) if typed >= self.rules.mindup else ""

    def first_empty_row(self):
        for r in range(3, self.last_row + 1):
            if self.ws.cell(row=r, column=1).value in (None, ""): return r
        return None

    def free_rows(self):
        return sum(1 for r in range(3, self.last_row + 1) if self.ws.cell(row=r, column=1).value in (None, ""))

# ----------------------------------------------------------------------------- evaluate records
def find_district(R, text):
    """District named in the text: state names are ignored (e.g. 'Nicobar' inside 'Andaman and Nicobar Islands'), a district that
    has the same name as a state (Puducherry) is used only if nothing more specific is found, otherwise the earliest mention wins."""
    t = re.sub(r"\s+", " ", text.lower().replace("&", " and "))
    t = re.sub(r"andaman\s+and\s+nicobar(\s+islands?)?", " ", t)
    state_names = {re.sub(r"\s*\(.*?\)", "", x).strip().lower() for x in R.states}
    hits = []
    for d in R.districts:
        m = re.search(r"(?<![a-z0-9])" + re.escape(d.lower()) + r"(?![a-z0-9])", t)
        if m: hits.append((d.lower() in state_names, m.start(), d))
    return min(hits)[2] if hits else None

def evaluate(db, records, overrides):
    R = db.rules
    out = []
    seen_ids, seen_keys, seen_ref = {}, {}, {}
    for rec in records:
        v = {}                                   # header -> final value
        auto = []
        # case fields: source -> app/GUI override -> workbook default
        for c in R.case:
            h = c["header"]
            val = rec.case.get(h)
            if val in (None, ""): val = overrides.get(h) or None
            if val in (None, ""):
                d = c["default"]
                if isinstance(d, str) and d.strip().upper() == "=TODAY()": d = datetime.datetime.combine(TODAY, datetime.time())
                if d not in (None, ""): val = d
            if val in (None, ""): continue
            if h in DATE_HEADERS:
                pd = parse_date(val); val = pd if pd else val
            elif h in NUM_HEADERS:
                try: val = fmt_num(float(clean(val)))
                except ValueError: val = clean(val)
            elif isinstance(val, str): val = val.strip()
            v[h] = val
        # alleles
        issues, typed_core, present = [], 0, {}
        for L in R.loci:
            name = L["name"]
            t1, t2, flag = rec.alleles.get(name, ("", "", ""))
            present[name] = bool(t1 and t2)
            if flag:
                issues.append(f"{name}: {flag}"); continue
            if not (t1 and t2): continue
            if name == "Amelogenin":
                a, b = sorted([t1, t2])
            else:
                a, b = sorted([float(t1), float(t2)])
                lo_hi = R.limits.get(name)
                if lo_hi and (a < lo_hi[0] or b > lo_hi[1]): issues.append(f"{name}: allele outside plausible range")
                a, b = fmt_num(a), fmt_num(b)
                if name in R.core: typed_core += 1
            v[f"{name} (1)"], v[f"{name} (2)"] = a, b
        present["Y-indel"] = "Y-indel" in v
        # inference: kit, state, district
        warns = list(rec.warnings)
        female = tuple(rec.alleles.get("Amelogenin", ("", "", ""))[:2]) == ("X", "X")
        matches = []
        for kit, rules in R.kits:
            ok = True
            for m, t in rules.items():
                if m == "Y-indel" and female: continue          # Y-indel is only expected in males
                pres = present.get(m, False)
                if (t == "R" and not pres) or (t == "X" and pres): ok = False; break
            if ok and typed_core == getattr(R, "kit_cover", {}).get(kit, typed_core): matches.append(kit)   # absent markers are only trusted for complete profiles
        if "Kit/Method" not in v:
            if len(matches) == 1: v["Kit/Method"] = matches[0]; auto.append("Kit/Method from markers")
            elif len(matches) > 1: warns.append("Kit could be " + " or ".join(matches) + " - choose Kit/Method")
        core_text = " ".join(str(v.get(h, "")) for h in ("Police Station", "Forwarding Authority", "Case/FIR No.", "Requisition Ref. No.")).lower()
        all_text = (rec.alltext or "").lower()
        if "State" not in v:
            st_found = None
            for kw, s_, d_ in R.locations:                                   # 1. your keyword table
                if s_ and kw in core_text: st_found = s_; break
            if not st_found:                                                 # 2. a state name in the case text, then anywhere in the file
                for txt in (core_text, all_text):
                    st_found = next((x for x in R.states if word_in(x, txt)), None)
                    if st_found: break
            if not st_found:                                                 # 3. a known district name implies its state
                for txt in (core_text, all_text):
                    st_found = next((DISTRICT_STATE[d.lower()] for d in R.districts if d.lower() in DISTRICT_STATE and word_in(d, txt)), None)
                    if st_found: break
            if st_found: v["State"] = st_found; auto.append("State from location")
        if "District" not in v:
            d_found = None
            for kw, s_, d_ in R.locations:
                if d_ and kw in core_text: d_found = d_; break
            if not d_found:
                for txt in (core_text, all_text):
                    d_found = find_district(R, txt)
                    if d_found: break
            if d_found: v["District"] = d_found; auto.append("District from location")
        # checks
        miss = [h for h in ("Profile Category", "Sample Type", "Kit/Method", "Analysis Date", "Analyst") if h not in v]
        if "Case/FIR No." not in v and "Person/Source ID" not in v: miss.append("FIR No. or Person ID")
        if miss: issues.append("Missing: " + ", ".join(miss))
        cat = v.get("Profile Category")
        if cat and cat not in R.lists["Profile Category"]: issues.append("Profile Category not in Settings list")
        mn_req = getattr(R, "cat_min", {}).get(cat)
        if mn_req and typed_core < mn_req:
            issues.append(f"Below the minimum of {mn_req} core loci for '{cat}' ({typed_core} typed)")
        pid = v.get("Profile ID")
        if pid:
            pid = str(pid)
            if pid in db.ids: issues.append("Profile ID already in database")
            if pid in seen_ids: issues.append("Duplicate Profile ID in this import")
            seen_ids[pid] = True
        sig = "|".join(f"{L['name']}:{v[L['name'] + ' (1)']},{v[L['name'] + ' (2)']}" for L in R.loci if f"{L['name']} (1)" in v)
        rfir, rps = norm(v.get("Case/FIR No.")), norm(v.get("Police Station"))
        if sig and (rfir or rps):
            refk = (rfir, rps, sig)
            if refk in db.refsig:
                issues.insert(0, f"DUPLICATE: this profile with the same case reference is already in the database ({db.refsig[refk]})")
            elif refk in seen_ref:
                issues.insert(0, "DUPLICATE: the same profile with the same case reference appears twice in this import")
            elif (rfir, rps) in db.refs:
                warns.append(f"Case reference already in the database under {db.refs[(rfir, rps)]} (different profile)")
            seen_ref[refk] = True
        key = ""
        if typed_core >= R.mindup:
            if getattr(db, "is_sqlite", False):
                key = genotype_key(R.core, {l: (a10(v[f"{l} (1)"]), a10(v[f"{l} (2)"])) for l in R.core if f"{l} (1)" in v})
            else:
                key = "".join(f"{v[f'{l} (1)']},{v[f'{l} (2)']};" if f"{l} (1)" in v else ",;" for l in R.core)
            if key in db.keys: issues.append(f"Identical profile already in database ({db.keys[key]})")
            if key in seen_keys: issues.append("Identical profile elsewhere in this import")
            seen_keys[key] = True
        hard = [i for i in issues if not i.startswith("Missing:")]
        status = ("Duplicate" if any(i.startswith("DUPLICATE") for i in issues) else
                  "Ready" if not issues else ("Incomplete" if not hard else "Review"))
        out.append({"rec": rec, "values": v, "issues": issues, "auto": auto, "typed": typed_core,
                    "warnings": warns, "status": status})
    return out

# ----------------------------------------------------------------------------- commit
def selected(results, include_incomplete=True, include_review=False):
    return [x for x in results if x["status"] == "Ready" or (include_incomplete and x["status"] == "Incomplete")
            or (include_review and x["status"] == "Review")]          # 'Duplicate' rows are never added


# ----------------------------------------------------------------------------- lossless xlsx patching
def _sheet_part(z, sheet_name):
    wbx = z.read("xl/workbook.xml").decode("utf-8")
    rid = None
    for m in re.finditer(r"<sheet\b[^>]*?/>", wbx):
        tag = m.group(0)
        nm = re.search(r'\bname="([^"]*)"', tag)
        if nm and html.unescape(nm.group(1)) == sheet_name:
            rid = re.search(r'\br:id="([^"]*)"', tag).group(1); break
    if rid is None: raise RuntimeError(f"Sheet '{sheet_name}' not found")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    for m in re.finditer(r"<Relationship\b[^>]*?/>", rels):
        tag = m.group(0)
        if re.search(r'\bId="%s"' % re.escape(rid), tag):
            t = re.search(r'\bTarget="([^"]*)"', tag).group(1)
            return t.lstrip("/") if t.startswith("/") else "xl/" + t
    raise RuntimeError("Sheet part not found")

class Formula(str):
    """A cell formula (without the leading '='). style_from = column index whose style the cell should copy."""
    style_from = None

class SharedStr(int):
    """Index of an existing shared string."""
    style_from = None

def formula(text, style_from=None):
    f = Formula(text); f.style_from = style_from; return f

def _cell_xml(ref, style, val):
    from openpyxl.utils.datetime import to_excel
    s = f' s="{style}"' if style else ""
    if val is None:
        return f'<c r="{ref}"{s}/>'
    if isinstance(val, SharedStr):
        return f'<c r="{ref}"{s} t="s"><v>{int(val)}</v></c>'
    if isinstance(val, Formula):
        return f'<c r="{ref}"{s}><f>{html.escape(str(val), quote=False)}</f></c>'
    if isinstance(val, (datetime.datetime, datetime.date)):
        return f'<c r="{ref}"{s}><v>{to_excel(val)}</v></c>'
    if isinstance(val, bool): return f'<c r="{ref}"{s} t="b"><v>{int(val)}</v></c>'
    if isinstance(val, (int, float)): return f'<c r="{ref}"{s}><v>{val!r}</v></c>'
    txt = html.escape(str(val), quote=False)
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{txt}</t></is></c>'

def patch_workbook(src, dst, sheet_name, updates, extra_sheets=None, transforms=None):
    """Writes cell values into existing rows of a sheet by editing the XML directly, so everything else in the file
    (dropdown lists, conditional formats, table, formulas, Excel extensions) stays byte-for-byte untouched.
    updates: {(row, col): value}. extra_sheets: {sheet_name: updates}. Forces Excel to recalculate on open."""
    from openpyxl.utils import get_column_letter, column_index_from_string
    jobs = {sheet_name: updates}
    jobs.update(extra_sheets or {})
    with zipfile.ZipFile(src) as zin:
        parts = {name: _sheet_part(zin, name) for name in jobs}
        newxml = {}
        for name, upd in jobs.items():
            xml = zin.read(parts[name]).decode("utf-8")
            byrow = {}
            for (r, c), val in upd.items(): byrow.setdefault(r, {})[c] = val
            pieces, pos, found = [], 0, set()
            for m in re.finditer(r'<row r="(\d+)"(?=[ >/])[^>]*?(?:/>|>.*?</row>)', xml, re.S):
                r = int(m.group(1))
                if r not in byrow: continue
                found.add(r)
                rowstr = m.group(0)
                open_tag = re.match(r"<row\b[^>]*?(?=/?>)", rowstr).group(0)
                existing = {}
                for cm in re.finditer(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>(.*?)</c>)' % r, rowstr, re.S):
                    existing[column_index_from_string(cm.group(1))] = (cm.group(0), cm.group(2))
                for c, val in byrow[r].items():
                    old = existing.get(c)
                    sm = re.search(r'\bs="(\d+)"', old[1]) if old else None
                    style = sm.group(1) if sm else None
                    sf = getattr(val, "style_from", None)
                    if sf and sf in existing:
                        m2 = re.search(r'\bs="(\d+)"', existing[sf][1])
                        if m2: style = m2.group(1)
                    existing[c] = (_cell_xml(f"{get_column_letter(c)}{r}", style, val), "")
                body = "".join(existing[c][0] for c in sorted(existing))
                pieces.append(xml[pos:m.start()]); pieces.append(open_tag + ">" + body + "</row>"); pos = m.end()
            pieces.append(xml[pos:]); xml = "".join(pieces)
            missing = set(byrow) - found
            if missing: raise RuntimeError(f"Row {min(missing)} not found in sheet '{name}' - the sheet layout is not as expected")
            newxml[parts[name]] = xml
        for key, fn in (transforms or {}).items():
            part = parts.get(key[6:]) if key.startswith("sheet:") else key
            data = newxml[part] if part in newxml else zin.read(part).decode("utf-8")
            newxml[part] = fn(data)
        wbx = newxml.get("xl/workbook.xml") or zin.read("xl/workbook.xml").decode("utf-8")
        mc = re.search(r"<calcPr\b[^>]*?/>", wbx)
        if mc:
            tag = mc.group(0)
            tag2 = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', tag) if "fullCalcOnLoad" in tag else tag.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
            wbx = wbx.replace(tag, tag2, 1)
        else:
            ins = '<calcPr fullCalcOnLoad="1"/>'
            wbx = wbx.replace("<extLst", ins + "<extLst", 1) if "<extLst" in wbx else wbx.replace("</workbook>", ins + "</workbook>", 1)
        newxml["xl/workbook.xml"] = wbx
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in newxml: data = newxml[item.filename].encode("utf-8")
                zout.writestr(item, data)


# ----------------------------------------------------------------------------- repair of the hidden Calc sheet
def calc_needs_repair(wb):
    """True if the helper sheet behind Search / Profile_Match / Dashboard was damaged (e.g. a row was deleted in the
    DNA Database sheet, leaving #REF! in Calc)."""
    if "Calc" not in wb.sheetnames: return False
    ws = wb["Calc"]
    for r in range(3, 40):
        for c in range(1, 9):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "#REF!" in v: return True
    for r in range(2000, 2004):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and "#REF!" in v: return True
    return False

def build_calc_updates(rules, hmap, first=3, last=2002):
    from openpyxl.utils import get_column_letter as GL
    def col(h): return GL(hmap[h])
    def idx(h): return f"INDEX('DNA Database'!${col(h)}:${col(h)},ROW())"
    fields = ["Profile ID", "Case/FIR No.", "Police Station", "Person/Source ID", "Exhibit No.", "Profile Category", "State"]
    core = rules.core
    upd = {}
    for r in range(first, last + 1):
        upd[(r, 1)] = Formula(f'IFERROR(VALUE(MID({idx("Profile ID")},5,10)),0)')
        upd[(r, 9)] = Formula(f'IF({idx("Profile ID")}="",0,1)')
        conds = ",".join(f'OR(Search!$B${4+i}="",ISNUMBER(SEARCH(Search!$B${4+i},{idx(h)})))' for i, h in enumerate(fields))
        upd[(r, 2)] = Formula(f'IF($I{r}=0,0,IF(COUNTA(Search!$B$4:$B$10)=0,0,IF(AND({conds}),1,0)))')
        upd[(r, 3)] = Formula(f"B{r}" if r == first else f"B{r}+C{r-1}")
        cmp_t, idn_t, shr_t = [], [], []
        for j, loc in enumerate(core):
            c1, c2 = 10 + 2 * j, 11 + 2 * j
            for k, c in ((1, c1), (2, c2)):
                upd[(r, c)] = Formula(f'IF({idx(f"{loc} ({k})")}="","",{idx(f"{loc} ({k})")})')
            m1, m2 = f"${GL(c1)}{r}", f"${GL(c2)}{r}"
            q1, q2 = f"Profile_Match!$B${9+j}", f"Profile_Match!$C${9+j}"
            allnum = f"AND(ISNUMBER({q1}),ISNUMBER({q2}),ISNUMBER({m1}),ISNUMBER({m2}))"
            cmp_t.append(f"IF({allnum},1,0)")
            idn_t.append(f"IF({allnum},IF(OR(AND({m1}={q1},{m2}={q2}),AND({m1}={q2},{m2}={q1})),1,0),0)")
            shr_t.append(f"IF({allnum},IF(OR({m1}={q1},{m1}={q2},{m2}={q1},{m2}={q2}),1,0),0)")
        upd[(r, 4)] = Formula(f"IF($I{r}=0,0," + "+".join(cmp_t) + ")")
        upd[(r, 5)] = Formula(f"IF($I{r}=0,0," + "+".join(idn_t) + ")")
        upd[(r, 6)] = Formula(f"IF($I{r}=0,0," + "+".join(shr_t) + ")")
        upd[(r, 7)] = Formula(f'IF(OR($I{r}=0,D{r}=0),0,IF(D{r}<Profile_Match!$B$5,0,IF(IF(Profile_Match!$B$4="High",D{r}-E{r},D{r}-F{r})<=Profile_Match!$B$6,1,0)))')
        upd[(r, 8)] = Formula(f'IF(G{r}=1,E{r}*1000000+F{r}*1000+(1-ROW()/100000),"")')
    return upd

def repair_database(path):
    db = Database(path)
    upd = build_calc_updates(db.rules, db.hmap)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="dna_repair_"); os.close(fd)
    try:
        bdir = backup_dir()
        shutil.copyfile(path, os.path.join(bdir, f"repair_{datetime.datetime.now():%Y%m%d_%H%M%S}_" + os.path.basename(path)))
        patch_workbook(path, tmp, "Calc", upd)
        load_workbook(tmp, read_only=True).close()
        shutil.copyfile(tmp, path)
    finally:
        try: os.remove(tmp)
        except OSError: pass


def _table_part(z, sheet_part):
    rels = "xl/worksheets/_rels/" + os.path.basename(sheet_part) + ".rels"
    t = z.read(rels).decode("utf-8")
    m = re.search(r'Target="([^"]*tables/[^"]*)"', t)
    p = m.group(1)
    return p.lstrip("/") if p.startswith("/") else "xl/" + os.path.normpath(os.path.join("worksheets", p)).replace("\\", "/").lstrip("./")

def add_code_columns(path, lab_code="CFSLH"):
    """Adds 'Identifiable Code' and 'De-identified Code' (plus 3 hidden helper columns) to the existing database table,
    with formulas for every row, and stores the laboratory code and secret code keys in the Settings sheet."""
    from openpyxl.utils import get_column_letter as GL
    db = Database(path)
    if "Identifiable Code" in db.hmap: return False
    ncols = max(db.hmap.values()); last = db.last_row
    c_id, c_di, c_n, c_p2, c_p3 = ncols + 1, ncols + 2, ncols + 3, ncols + 4, ncols + 5
    L_ = {k: GL(v) for k, v in (("id", c_id), ("di", c_di), ("n", c_n), ("p2", c_p2), ("p3", c_p3))}
    pid, idx = GL(db.hmap["Profile ID"]), GL(db.hmap["Index Group"])
    ed, ad = GL(db.hmap["Entry Date"]), GL(db.hmap["Analysis Date"])
    with zipfile.ZipFile(path) as z:
        sst = z.read("xl/sharedStrings.xml").decode("utf-8")
    base = len(re.findall(r"<si[ >]", sst))
    texts = ["Identifiable Code", "De-identified Code", "Code helper n", "Code helper p2", "Code helper p3", "DATABASE CODES (automatic)"]
    sidx = {t: base + i for i, t in enumerate(texts)}
    keys = new_code_keys()
    sett = {(8, 16): "Laboratory code (used in Identifiable Code)", (8, 17): lab_code}
    names = ["Code key 1 (secret)", "Code key 2 (secret)", "Code key 3 (secret)", "Code key 4 (secret)", "Code key 5 (secret)", "Code key 6 (secret)"]
    for i in range(6):
        sett[(9 + i, 16)] = names[i]; sett[(9 + i, 17)] = keys[i]
    sett[(15, 16)] = "Do NOT change the laboratory code or keys after codes have been issued - the codes would change."
    K = [f"Settings!$Q${9+i}" for i in range(6)]
    M = "1073741824"
    upd = {}
    # header band + headers
    band_style = None
    for col in (c_id, c_di, c_n, c_p2, c_p3):
        upd[(1, col)] = None if col != c_id else SharedStr(sidx["DATABASE CODES (automatic)"])
        upd[(2, col)] = SharedStr(sidx[texts[col - c_id]])
        upd[(2, col)].style_from = db.hmap["Index Group"]
    for col in (c_id, c_di, c_n, c_p2, c_p3):
        if col != c_id: upd[(1, col)] = None
    for r in range(3, last + 1):
        sf = db.hmap["Index Group"]
        A = f"${pid}{r}"
        n = f"${L_['n']}{r}"
        upd[(r, c_n)] = formula(f'IF({A}="","",IFERROR(VALUE(MID({A},FIND("-",{A})+1,12)),ROW()))', sf)
        p1 = f"MOD({n}*{K[0]}+{K[1]},{M})"
        upd[(r, c_p2)] = formula(f'IF({n}="","",MOD(_xlfn.BITXOR({p1},_xlfn.BITRSHIFT({p1},15))*{K[2]}+{K[3]},{M}))', sf)
        p2 = f"${L_['p2']}{r}"
        upd[(r, c_p3)] = formula(f'IF({n}="","",MOD(_xlfn.BITXOR({p2},_xlfn.BITRSHIFT({p2},13))*{K[4]}+{K[5]},{M}))', sf)
        p3 = f"${L_['p3']}{r}"
        v = [f"MOD(INT({p3}/32^{5-i}),32)" for i in range(6)]
        ch = [f'MID("{CODE_ALPHABET}",{v[i]}+1,1)' for i in range(6)]
        chk = f'MID("{CODE_ALPHABET}",MOD(1*{v[0]}+3*{v[1]}+5*{v[2]}+7*{v[3]}+9*{v[4]}+11*{v[5]},32)+1,1)'
        upd[(r, c_di)] = formula(f'IF({n}="","","DI-"&{ch[0]}&{ch[1]}&{ch[2]}&{ch[3]}&"-"&{ch[4]}&{ch[5]}&{chk})', sf)
        yr = f'TEXT(IF(ISNUMBER(${ed}{r}),YEAR(${ed}{r}),IF(ISNUMBER(${ad}{r}),YEAR(${ad}{r}),0)),"0000")'
        s10 = f'({yr}&TEXT({n},"000000"))'
        chk11 = "+".join(f"MID({s10},{i},1)*{i+1}" for i in range(1, 11))
        upd[(r, c_id)] = formula(
            f'IF({n}="","",Settings!$Q$8&"-"&IFERROR(INDEX({INDEX_ARRAY},MATCH(${idx}{r},{GROUP_ARRAY},0)),"UNK")&"-"&{yr}&"-"&TEXT({n},"000000")'
            f'&"-"&IF(MOD({chk11},11)=10,"X",MOD({chk11},11)))', sf)
    newn = c_p3
    old_ref = re.search(r'ref="(A2:[A-Z]+\d+)"', zipfile.ZipFile(path).read(_table_part(zipfile.ZipFile(path), _sheet_part(zipfile.ZipFile(path), DB_SHEET))).decode()).group(1)
    new_ref = f"A2:{GL(newn)}{last}"
    def sheet_fn(x):
        x = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:{GL(newn)}{last}"/>', x, 1)
        x = x.replace(f'spans="1:{ncols}"', f'spans="1:{newn}"')
        x = x.replace("</cols>", f'<col min="{c_id}" max="{c_di}" width="26" customWidth="1"/><col min="{c_n}" max="{c_p3}" width="14" hidden="1" customWidth="1"/></cols>', 1)
        mm = re.search(r'<mergeCells count="(\d+)">', x)
        x = x.replace(mm.group(0), f'<mergeCells count="{int(mm.group(1)) + 1}">', 1)
        x = x.replace("</mergeCells>", f'<mergeCell ref="{GL(c_id)}1:{GL(c_di)}1"/></mergeCells>', 1)
        return x
    def table_fn(x):
        x = x.replace(f'ref="{old_ref}"', f'ref="{new_ref}"')
        mx = max(int(i) for i in re.findall(r'<tableColumn id="(\d+)"', x))
        cnt = re.search(r'<tableColumns count="(\d+)"', x)
        x = x.replace(cnt.group(0), f'<tableColumns count="{int(cnt.group(1)) + 5}"', 1)
        add_cols = "".join(f'<tableColumn id="{mx+1+i}" name="{texts[i]}"/>' for i in range(5))
        return x.replace("</tableColumns>", add_cols + "</tableColumns>", 1)
    def sst_fn(x):
        m = re.search(r'<sst [^>]*count="(\d+)" uniqueCount="(\d+)"', x)
        x = x.replace(f'count="{m.group(1)}" uniqueCount="{m.group(2)}"', f'count="{int(m.group(1)) + 6}" uniqueCount="{int(m.group(2)) + 6}"', 1)
        return x.replace("</sst>", "".join(f"<si><t>{html.escape(t, quote=False)}</t></si>" for t in texts) + "</sst>", 1)
    with zipfile.ZipFile(path) as z:
        tpart = _table_part(z, _sheet_part(z, DB_SHEET))
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="dna_codes_"); os.close(fd)
    try:
        bdir = backup_dir()
        shutil.copyfile(path, os.path.join(bdir, f"before_codes_{datetime.datetime.now():%Y%m%d_%H%M%S}_" + os.path.basename(path)))
        patch_workbook(path, tmp, DB_SHEET, upd, {"Settings": sett},
                       {"sheet:" + DB_SHEET: sheet_fn, tpart: table_fn, "xl/sharedStrings.xml": sst_fn})
        chk_wb = load_workbook(tmp, read_only=True); chk_wb.close()
        shutil.copyfile(tmp, path)
    finally:
        try: os.remove(tmp)
        except OSError: pass
    return True


def export_deidentified(db, out_dir, only_approved=False, ids=None):
    """Writes (1) a de-identified profile file - de-identified code + STR profile only, no case or personal data - and
    (2) a separate CONFIDENTIAL linkage register (de-identified code <-> identifiable code / case references)."""
    from openpyxl import Workbook
    R = db.rules
    if not R.code_keys: raise RuntimeError("The database has no code columns yet - run 'Add code columns' first.")
    hm, ws = db.hmap, db.ws
    def val(r, h): return ws.cell(row=r, column=hm[h]).value if h in hm else None
    excl = {"Rejected", "Expunged / Removed", "On hold / Quarantined"}
    ok_status = {"Approved for upload", "Uploaded to DB"}
    heads = ["De-identified Code", "Index Group", "Profile Type", "No. of Contributors", "Kit/Method", "Sex (Amelogenin)"] + \
            [f"{L['name']} ({i})" for L in R.loci for i in (1, 2)] + ["Y-indel"]
    prof, link = [], []
    want = {str(i).strip().upper() for i in ids} if ids else None
    for r in range(3, db.last_row + 1):
        pid = val(r, "Profile ID")
        if pid in (None, ""): continue
        st_ = val(r, "Record Status")
        grp = R.cat_index.get(val(r, "Profile Category"))
        iden, did = make_codes(pid, grp, val(r, "Entry Date"), val(r, "Analysis Date"), R.code_keys, R.code_lab)
        if want is not None:
            if str(pid).upper() not in want and did.upper() not in want: continue
        elif st_ in excl or (only_approved and st_ not in ok_status): continue
        if not did: continue
        a1, a2 = val(r, "Amelogenin (1)"), val(r, "Amelogenin (2)")
        sex = "" if not (a1 and a2) else ("Female (XX)" if (a1, a2) == ("X", "X") else "Male (XY)" if {a1, a2} == {"X", "Y"} else "Check")
        row = [did, grp or "", val(r, "Profile Type"), val(r, "No. of Contributors"), val(r, "Kit/Method"), sex]
        for L in R.loci:
            row += [val(r, f"{L['name']} (1)"), val(r, f"{L['name']} (2)")]
        row.append(val(r, "Y-indel"))
        prof.append(row)
        link.append([did, iden, pid, val(r, "Lab Case No."), val(r, "Case/FIR No."), val(r, "Police Station"), val(r, "Person/Source ID"), val(r, "Exhibit No.")])
    if want is not None and not prof: raise RuntimeError("No matching profile found.")
    stamp = f"{datetime.datetime.now():%Y%m%d_%H%M}"
    os.makedirs(out_dir, exist_ok=True)
    p1 = os.path.join(out_dir, f"DNA_Deidentified_Export_{stamp}.xlsx")
    p2 = os.path.join(out_dir, f"Linkage_Register_CONFIDENTIAL_{stamp}.xlsx")
    wb = Workbook(); w1 = wb.active; w1.title = "De-identified profiles"
    w1.append(heads)
    for row in prof: w1.append(row)
    w1.freeze_panes = "B2"
    info = wb.create_sheet("Info")
    for t in (["Laboratory code", R.code_lab], ["Exported on", datetime.datetime.now().strftime("%d-%m-%Y %H:%M")], ["Profiles", len(prof)],
              ["Contents", "De-identified code and STR profile only. No case, person, place or date information."],
              ["Selection", "Profiles selected by the user" if want is not None else ("Approved/uploaded profiles only" if only_approved else "All profiles except rejected, removed or on hold")]):
        info.append(t)
    wb.save(p1)
    wb2 = Workbook(); w2 = wb2.active; w2.title = "Linkage register"
    w2.append(["CONFIDENTIAL - links de-identified codes to identifiable records. Store separately and restrict access."])
    w2.append(["De-identified Code", "Identifiable Code", "Profile ID", "Lab Case No.", "Case/FIR No.", "Police Station", "Person/Source ID", "Exhibit No."])
    for row in link: w2.append(row)
    wb2.save(p2)
    return p1, p2, len(prof)


def export_structured(db, out_dir, only_approved=False, ids=None, per_profile=False, formats=("json", "xml")):
    """Writes the de-identified profiles as structured JSON and XML (generic exchange schema) plus SHA-256 checksums.
    NOT an official CODIS CMF / Interpol / Pruem file: map to the receiving node's published schema before submitting."""
    import json, hashlib
    import xml.etree.ElementTree as ET
    R = db.rules
    if not R.code_keys: raise RuntimeError("The database has no code columns yet - run 'Add code columns' first.")
    hm, ws = db.hmap, db.ws
    def val(r, h): return ws.cell(row=r, column=hm[h]).value if h in hm else None
    excl = {"Rejected", "Expunged / Removed", "On hold / Quarantined"}
    ok_status = {"Approved for upload", "Uploaded to DB"}
    profiles = []
    want = {str(i).strip().upper() for i in ids} if ids else None
    for r in range(3, db.last_row + 1):
        pid = val(r, "Profile ID")
        if pid in (None, ""): continue
        st_ = val(r, "Record Status")
        grp = R.cat_index.get(val(r, "Profile Category"))
        _, did = make_codes(pid, grp, val(r, "Entry Date"), val(r, "Analysis Date"), R.code_keys, R.code_lab)
        if want is not None:
            if str(pid).upper() not in want and did.upper() not in want: continue      # explicitly chosen: export regardless of status
        elif st_ in excl or (only_approved and st_ not in ok_status): continue
        if not did: continue
        loci = []
        for L in R.loci:
            a, b = val(r, f"{L['name']} (1)"), val(r, f"{L['name']} (2)")
            if a in (None, "") or b in (None, ""): continue          # untyped loci are omitted, never sent as null
            loci.append({"locus": L["name"], "alleles": [clean(a), clean(b)]})
        p = {"specimen_id": did, "specimen_category": grp or "", "profile_type": val(r, "Profile Type") or "",
             "kit": val(r, "Kit/Method") or "", "loci": loci}
        if val(r, "Y-indel") not in (None, ""): p["y_indel"] = clean(val(r, "Y-indel"))
        profiles.append(p)
    if want is not None and not profiles: raise RuntimeError("No matching profile found for: " + ", ".join(sorted(want)))
    now = datetime.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"
    schema = "generic-dna-profile-exchange/1.0 (not an official CMF)"
    stamp = f"{datetime.datetime.now():%Y%m%d_%H%M}"
    os.makedirs(out_dir, exist_ok=True)

    def to_xml(plist):
        root = ET.Element("DnaProfileExchange", schema=schema)
        h = ET.SubElement(root, "Header")
        ET.SubElement(h, "SourceLaboratory").text = str(R.code_lab)
        ET.SubElement(h, "ExportedUtc").text = now
        ET.SubElement(h, "ProfileCount").text = str(len(plist))
        ps = ET.SubElement(root, "Profiles")
        for p in plist:
            e = ET.SubElement(ps, "Profile")
            for k, tag in (("specimen_id", "SpecimenId"), ("specimen_category", "SpecimenCategory"), ("profile_type", "ProfileType"), ("kit", "Kit")):
                ET.SubElement(e, tag).text = p[k]
            le = ET.SubElement(e, "Loci")
            for l in p["loci"]:
                x = ET.SubElement(le, "Locus", name=l["locus"])
                for a_ in l["alleles"]: ET.SubElement(x, "Allele").text = a_
            if "y_indel" in p: ET.SubElement(e, "YIndel").text = p["y_indel"]
        ET.indent(root)
        return ET.ElementTree(root)

    groups = [(p["specimen_id"], [p]) for p in profiles] if per_profile else [(f"dna_profiles_{stamp}", profiles)]
    written = []
    for base, plist in groups:
        meta = {"schema": schema, "source_laboratory": R.code_lab, "exported_utc": now, "profile_count": len(plist)}
        if "json" in formats:
            pj = os.path.join(out_dir, base + ".json")
            with open(pj, "w", encoding="utf-8") as f: json.dump({**meta, "profiles": plist}, f, indent=2, ensure_ascii=False)
            written.append(pj)
        if "xml" in formats:
            px = os.path.join(out_dir, base + ".xml")
            to_xml(plist).write(px, encoding="utf-8", xml_declaration=True)
            written.append(px)
    sums = os.path.join(out_dir, f"checksums_{stamp}.sha256")
    with open(sums, "w") as f:
        for p_ in written:
            f.write(f"{hashlib.sha256(open(p_, 'rb').read()).hexdigest()}  {os.path.basename(p_)}\n")
    return written, sums, len(profiles)

def list_profiles(db):
    """[(Profile ID, De-identified Code, Case/FIR No., Police Station, Profile Category)] for choosing a profile to export."""
    if getattr(db, "is_sqlite", False): return db.list_profiles()
    R, out = db.rules, []
    for r in range(3, db.last_row + 1):
        g = lambda h: db.ws.cell(row=r, column=db.hmap[h]).value if h in db.hmap else None
        pid = g("Profile ID")
        if pid in (None, ""): continue
        did = make_codes(pid, R.cat_index.get(g("Profile Category")), g("Entry Date"), g("Analysis Date"), R.code_keys, R.code_lab)[1] if R.code_keys else ""
        blob = " ".join(str(g(h) or "") for h in ("Profile ID", "Lab Case No.", "Case/FIR No.", "Police Station", "District", "State",
                                                "Profile Category", "Person/Source ID", "Exhibit No.", "Record Status", "Kit/Method")) + " " + str(did)
        out.append((str(pid), did, g("Case/FIR No.") or "", g("Police Station") or "", g("Profile Category") or "", g("Record Status") or "", blob.lower()))
    return out

def backup_dir():
    base = os.environ.get("DNA_IMPORT_BACKUP_DIR") or os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "DNAImportApp", "backups")
    os.makedirs(base, exist_ok=True)
    return base

def commit(db, results, include_review=False, backup=True, include_incomplete=True):
    """Adds the selected rows to the EXISTING database file. No new file is created next to it: the workbook is
    saved to a temporary file outside the database folder, checked, then written over the original in place.
    A rolling safety copy (last 5) is kept in the app's own folder, not in the database folder."""
    if getattr(db, "is_sqlite", False): return commit_sqlite(db, results, include_review, include_incomplete)
    rows = selected(results, include_incomplete, include_review)
    if not rows: return 0, None
    if len(rows) > db.free_rows(): raise RuntimeError("Not enough free rows in the database sheet.")
    bak = None
    if backup:
        bdir = backup_dir()
        stem = os.path.splitext(os.path.basename(db.path))[0]
        bak = os.path.join(bdir, f"{stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copyfile(db.path, bak)
        old = sorted(f for f in os.listdir(bdir) if f.startswith(stem + "_"))
        for f in old[:-5]:
            try: os.remove(os.path.join(bdir, f))
            except OSError: pass
    num = db.maxnum
    empties = [r for r in range(3, db.last_row + 1) if db.ws.cell(row=r, column=1).value in (None, "")][:len(rows)]
    updates = {}
    for x, r in zip(rows, empties):
        v = dict(x["values"])
        if not v.get("Profile ID"):
            num += 1; v["Profile ID"] = f"DNA-{num:06d}"
        for h, val in v.items():
            c = db.hmap.get(h)
            if c: updates[(r, c)] = val
        db.ids.add(str(v["Profile ID"])); x["assigned_id"] = v["Profile ID"]
        x["codes"] = make_codes(v["Profile ID"], db.rules.cat_index.get(v.get("Profile Category")), v.get("Entry Date"),
                                v.get("Analysis Date"), db.rules.code_keys, db.rules.code_lab) if db.rules.code_keys else ("", "")
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", prefix="dna_import_")      # system temp folder, not the database folder
    os.close(fd)
    try:
        patch_workbook(db.path, tmp, DB_SHEET, updates)                   # lossless edit of the existing file's XML
        chk = load_workbook(tmp, read_only=True)                          # verify the result opens and holds the new IDs
        cs = chk[DB_SHEET]
        for x, r in zip(rows, empties):
            if str(cs.cell(row=r, column=1).value) != str(x["assigned_id"]):
                raise RuntimeError("Verification failed - database left unchanged.")
        chk.close()
        shutil.copyfile(tmp, db.path)                                     # overwrite the existing database in place
    finally:
        try: os.remove(tmp)
        except OSError: pass
    return len(rows), bak


# ----------------------------------------------------------------------------- SQLite store (for lakhs of profiles)
SQL_SCHEMA = """
CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS locus(locus_id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, is_core INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS profile(
    profile_id INTEGER PRIMARY KEY,
    specimen_id TEXT NOT NULL UNIQUE,              -- the De-identified Code (no personal / case information)
    specimen_category TEXT, profile_type TEXT, kit TEXT, sex TEXT, y_indel TEXT,
    loci_typed INTEGER NOT NULL DEFAULT 0,
    genotype_key TEXT,                             -- all core loci in one string: exact whole-profile lookup
    status TEXT, updated TEXT);
CREATE TABLE IF NOT EXISTS allele(                 -- one row per profile and locus; alleles x10 as integers (9.3 -> 93), a1 <= a2
    profile_id INTEGER NOT NULL REFERENCES profile(profile_id) ON DELETE CASCADE,
    locus_id INTEGER NOT NULL REFERENCES locus(locus_id),
    a1 INTEGER NOT NULL, a2 INTEGER NOT NULL,
    PRIMARY KEY(profile_id, locus_id)) WITHOUT ROWID;
CREATE INDEX IF NOT EXISTS ix_allele_genotype ON allele(locus_id, a1, a2, profile_id);
CREATE INDEX IF NOT EXISTS ix_profile_key ON profile(genotype_key);
"""
LINK_SCHEMA = """
CREATE TABLE IF NOT EXISTS linkage(
    specimen_id TEXT PRIMARY KEY, identifiable_code TEXT, profile_id TEXT, lab_case_no TEXT, case_fir_no TEXT,
    police_station TEXT, person_source_id TEXT, exhibit_no TEXT, updated TEXT);
"""
SQL_DB_NAME, SQL_LINK_NAME = "dna_profiles.db", "dna_linkage_CONFIDENTIAL.db"

def a10(x):
    return int(round(float(x) * 10))

def genotype_key(core_loci, geno):
    """geno: {locus: (a1x10, a2x10)} -> 'D3S1358:150,160|vWA:-|...'"""
    return "|".join(f"{l}:{geno[l][0]},{geno[l][1]}" if l in geno else f"{l}:-" for l in core_loci)

def export_sqlite(db, out_dir, ids=None, only_approved=False, linkage=True):
    """Creates / updates dna_profiles.db (de-identified: profile + allele tables, indexed for fast matching) and, separately,
    dna_linkage_CONFIDENTIAL.db (de-identified code <-> identifiable details). Re-running updates existing profiles."""
    R = db.rules
    if not R.code_keys: raise RuntimeError("The database has no code columns yet - add them first.")
    hm, ws = db.hmap, db.ws
    def val(r, h): return ws.cell(row=r, column=hm[h]).value if h in hm else None
    want = {str(i).strip().upper() for i in ids} if ids else None
    excl = {"Rejected", "Expunged / Removed", "On hold / Quarantined"}
    ok_status = {"Approved for upload", "Uploaded to DB"}
    os.makedirs(out_dir, exist_ok=True)
    p_db, p_link = os.path.join(out_dir, SQL_DB_NAME), os.path.join(out_dir, SQL_LINK_NAME)
    con = sqlite3.connect(p_db)
    con.execute("PRAGMA journal_mode=WAL"); con.execute("PRAGMA foreign_keys=ON")
    con.executescript(SQL_SCHEMA)
    lcon = None
    if linkage:
        lcon = sqlite3.connect(p_link); lcon.executescript(LINK_SCHEMA)
    loci = [L["name"] for L in R.loci if L["name"] != "Amelogenin"]
    for i, name in enumerate(loci, 1):
        con.execute("INSERT OR IGNORE INTO locus(locus_id, name, is_core) VALUES (?,?,?)", (i, name, 1 if name in R.core else 0))
    lid = {name: i for i, name in enumerate(loci, 1)}
    now = datetime.datetime.now().isoformat(timespec="seconds")
    added = updated = 0
    for r in range(3, db.last_row + 1):
        pid = val(r, "Profile ID")
        if pid in (None, ""): continue
        st_ = val(r, "Record Status")
        grp = R.cat_index.get(val(r, "Profile Category"))
        iden, did = make_codes(pid, grp, val(r, "Entry Date"), val(r, "Analysis Date"), R.code_keys, R.code_lab)
        if not did: continue
        if want is not None:
            if str(pid).upper() not in want and did.upper() not in want: continue
        elif st_ in excl or (only_approved and st_ not in ok_status): continue
        geno = {}
        for name in loci:
            x, y = val(r, f"{name} (1)"), val(r, f"{name} (2)")
            if isinstance(x, (int, float)) and isinstance(y, (int, float)):
                geno[name] = tuple(sorted((a10(x), a10(y))))
        am1, am2 = val(r, "Amelogenin (1)"), val(r, "Amelogenin (2)")
        sex = "" if not (am1 and am2) else ("F" if (am1, am2) == ("X", "X") else "M" if {am1, am2} == {"X", "Y"} else "?")
        core_n = sum(1 for l in R.core if l in geno)
        row = (did, grp or "", val(r, "Profile Type") or "", val(r, "Kit/Method") or "", sex, clean(val(r, "Y-indel")), core_n,
               genotype_key(R.core, geno) if core_n else "", st_ or "", now)
        cur = con.execute("SELECT profile_id FROM profile WHERE specimen_id=?", (did,)).fetchone()
        if cur:
            pk = cur[0]
            con.execute("UPDATE profile SET specimen_category=?, profile_type=?, kit=?, sex=?, y_indel=?, loci_typed=?, genotype_key=?, status=?, updated=? WHERE profile_id=?",
                        row[1:] + (pk,))
            con.execute("DELETE FROM allele WHERE profile_id=?", (pk,)); updated += 1
        else:
            pk = con.execute("INSERT INTO profile(specimen_id, specimen_category, profile_type, kit, sex, y_indel, loci_typed, genotype_key, status, updated) VALUES (?,?,?,?,?,?,?,?,?,?)", row).lastrowid
            added += 1
        con.executemany("INSERT INTO allele(profile_id, locus_id, a1, a2) VALUES (?,?,?,?)", [(pk, lid[l], g[0], g[1]) for l, g in geno.items()])
        if lcon is not None:
            lcon.execute("INSERT OR REPLACE INTO linkage VALUES (?,?,?,?,?,?,?,?,?)",
                         (did, iden, str(pid), clean(val(r, "Lab Case No.")), clean(val(r, "Case/FIR No.")), clean(val(r, "Police Station")),
                          clean(val(r, "Person/Source ID")), clean(val(r, "Exhibit No.")), now))
    con.execute("INSERT OR REPLACE INTO meta VALUES ('schema_version','1')")
    con.execute("INSERT OR REPLACE INTO meta VALUES ('laboratory_code', ?)", (R.code_lab,))
    con.execute("INSERT OR REPLACE INTO meta VALUES ('allele_encoding','allele x 10 as integer (9.3 -> 93)')")
    con.execute("INSERT OR REPLACE INTO meta VALUES ('last_export', ?)", (now,))
    con.commit(); con.execute("PRAGMA optimize"); total = con.execute("SELECT COUNT(*) FROM profile").fetchone()[0]; con.close()
    if lcon is not None: lcon.commit(); lcon.close()
    return p_db, (p_link if linkage else None), added, updated, total

def parse_query(text):
    """'D3S1358 15,16' lines, or 'D3S1358=15,16;vWA=16,17' -> {locus: (a1x10, a2x10)}"""
    q = {}
    for part in re.split(r"[;\n]+", text or ""):
        part = part.strip()
        if not part: continue
        m = re.match(r"^\s*([A-Za-z0-9 ]+?)\s*[=:\t ]\s*([0-9.]+)\s*[,/ \t]\s*([0-9.]+)\s*$", part)
        if not m: raise ValueError(f"Cannot read '{part}'. Use e.g. D3S1358 15,16")
        name = m.group(1).strip()
        key = next((n for n in (L["name"] for L in ALL_LOCI_NAMES) if norm(n) == norm(name)), name)
        q[key] = tuple(sorted((a10(m.group(2)), a10(m.group(3)))))
    return q

ALL_LOCI_NAMES = [{"name": n} for n in ["D3S1358", "vWA", "D16S539", "CSF1PO", "TPOX", "D8S1179", "D21S11", "D18S51", "D5S818", "D13S317", "D7S820",
                                        "TH01", "D12S391", "D19S433", "FGA", "D2S1338", "D10S1248", "D22S1045", "D1S1656", "D2S441", "SE33", "Penta D", "Penta E", "D6S1043"]]

def json_dumps_q(q):
    return ";".join(f"{k}={v[0]},{v[1]}" for k, v in sorted(q.items()))

def sqlite_search(path, query, mode="high", min_loci=8, mismatches=0, limit=100):
    """Match a query profile against the SQLite store. mode 'high' = identical genotype at every compared locus (minus allowed
    mismatches); 'moderate' = at least one shared allele at every compared locus. Returns (rows, info)."""
    con = sqlite3.connect(path)
    lid = {n: i for i, n in con.execute("SELECT locus_id, name FROM locus")}
    core = [n for n, in con.execute("SELECT name FROM locus WHERE is_core=1 ORDER BY locus_id")]
    q = {}
    for name, g in query.items():
        key = next((n for n in lid if norm(n) == norm(name)), None)
        if key is None: raise ValueError(f"Unknown locus '{name}'")
        q[key] = g
    nq = len(q)
    if nq < min_loci: raise ValueError(f"Only {nq} locus/loci in the query - minimum loci compared is {min_loci}.")
    info = {"query_loci": nq, "path": "scan"}
    cand_sql = None
    # --- fastest path: whole core profile, no mismatches -> single indexed key lookup
    if mode == "high" and mismatches == 0 and all(l in q for l in core):
        key = genotype_key(core, q)
        ids = [r[0] for r in con.execute("SELECT profile_id FROM profile WHERE genotype_key=?", (key,))]
        info["path"] = "exact key lookup"
        cand = ids
    elif mode == "high":
        # pigeonhole: a profile identical at >= (min_loci - mismatches) of the nq loci must be identical at >= 1 of any
        # (nq - need + 1) loci -> use the rarest genotypes (cheap index counts) to build a small candidate set
        need = max(1, min_loci - mismatches); k = nq - need + 1
        counts = sorted(((con.execute("SELECT COUNT(*) FROM allele WHERE locus_id=? AND a1=? AND a2=?", (lid[l], g[0], g[1])).fetchone()[0], l)
                         for l, g in q.items()))
        pick = counts[:k]
        total = con.execute("SELECT COUNT(*) FROM profile").fetchone()[0] or 1
        if sum(c for c, _ in pick) < 0.4 * total:
            cs = set()
            for _, l in pick:
                cs.update(r[0] for r in con.execute("SELECT profile_id FROM allele WHERE locus_id=? AND a1=? AND a2=?", (lid[l], q[l][0], q[l][1])))
            cand = sorted(cs); info["path"] = f"index candidates (rarest {len(pick)} locus genotypes)"
        else:
            cand = None
    else:
        cand = None
    con.execute("CREATE TEMP TABLE q(locus_id INTEGER PRIMARY KEY, a INTEGER, b INTEGER)")
    con.executemany("INSERT INTO q VALUES (?,?,?)", [(lid[l], g[0], g[1]) for l, g in q.items()])
    if cand is not None:
        con.execute("CREATE TEMP TABLE cand(profile_id INTEGER PRIMARY KEY)")
        con.executemany("INSERT OR IGNORE INTO cand VALUES (?)", [(i,) for i in cand])
        src = "FROM cand c CROSS JOIN allele a ON a.profile_id = c.profile_id CROSS JOIN q ON q.locus_id = a.locus_id"   # CROSS JOIN fixes the join order: candidates first (the planner has no statistics on temp tables and otherwise scans the whole allele table)
        info["candidates"] = len(cand)
    else:
        src = "FROM allele a JOIN q ON q.locus_id = a.locus_id"
    miss = "comp - ident" if mode == "high" else "comp - shared"
    sql = f"""SELECT p.specimen_id, p.specimen_category, p.kit, comp, ident, shared FROM (
                SELECT a.profile_id AS pid, COUNT(*) AS comp,
                       SUM(a.a1 = q.a AND a.a2 = q.b) AS ident,
                       SUM(a.a1 IN (q.a, q.b) OR a.a2 IN (q.a, q.b)) AS shared
                {src} GROUP BY a.profile_id HAVING comp >= ? AND ({miss}) <= ?) t
              JOIN profile p ON p.profile_id = t.pid ORDER BY ident DESC, shared DESC, p.specimen_id LIMIT ?"""
    rows = con.execute(sql, (min_loci, mismatches, limit)).fetchall()
    try:
        if con.execute("SELECT 1 FROM sqlite_master WHERE name='search_log'").fetchone():
            import hashlib
            qh = hashlib.sha256(json_dumps_q(q).encode()).hexdigest()[:16]
            con.execute("INSERT INTO search_log(ts,user,mode,query_loci,min_loci,mismatches,results,query_hash) VALUES (?,?,?,?,?,?,?,?)",
                        (datetime.datetime.now().isoformat(timespec="seconds"), current_user(), mode, nq, min_loci, mismatches, len(rows), qh))
            con.commit()
    except sqlite3.Error:
        pass
    con.close()
    return rows, info

def linkage_lookup(link_path, specimen_ids):
    """Identifiers for de-identified codes: from a separate linkage database, or from the working database itself."""
    con = sqlite3.connect(link_path); out = {}
    has_view = con.execute("SELECT 1 FROM sqlite_master WHERE name='v_profile_full'").fetchone() is not None
    for sid in specimen_ids:
        if has_view:
            r = con.execute("SELECT identifiable_code, profile_ref, lab_case_no, fir_no, police_station FROM v_profile_full WHERE deidentified_code=?", (sid,)).fetchone()
        else:
            r = con.execute("SELECT identifiable_code, profile_id, lab_case_no, case_fir_no, police_station FROM linkage WHERE specimen_id=?", (sid,)).fetchone()
        if r: out[sid] = r
    con.close(); return out


# ----------------------------------------------------------------------------- FULL SQLite DATABASE (working database)
FULL_SCHEMA_VERSION = "2"
AUDITED_COLS = ["record_status", "specimen_category", "kit", "upload_status", "upload_date", "expunged_date", "expunge_reason",
                "technical_reviewer", "review_date", "administrative_reviewer", "sample_type", "person_source_id", "case_id",
                "analysis_date", "profile_type", "no_of_contributors"]
FULL_SCHEMA = """
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS meta(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS config(key TEXT PRIMARY KEY, value TEXT);          -- lab_code, code_key_1..6, min_dup_loci, audit_enabled, app_user
CREATE TABLE IF NOT EXISTS ref_category(ord INTEGER, category TEXT PRIMARY KEY, index_group TEXT, min_core_loci INTEGER);
CREATE TABLE IF NOT EXISTS ref_list(list_name TEXT NOT NULL, value TEXT NOT NULL, ord INTEGER, PRIMARY KEY(list_name, value));
CREATE TABLE IF NOT EXISTS ref_kit(ord INTEGER, kit TEXT PRIMARY KEY, core_loci_covered INTEGER);
CREATE TABLE IF NOT EXISTS ref_kit_rule(kit TEXT NOT NULL, marker TEXT NOT NULL, rule TEXT NOT NULL CHECK(rule IN ('R','X')), PRIMARY KEY(kit, marker));
CREATE TABLE IF NOT EXISTS ref_location(ord INTEGER, keyword TEXT, state TEXT, district TEXT);
CREATE TABLE IF NOT EXISTS ref_field(ord INTEGER, field TEXT PRIMARY KEY, default_value TEXT, vertical_address TEXT);
CREATE TABLE IF NOT EXISTS ref_alias(ord INTEGER, field TEXT NOT NULL, alias TEXT NOT NULL, PRIMARY KEY(field, alias));
CREATE TABLE IF NOT EXISTS ref_locus_alias(ord INTEGER, locus TEXT NOT NULL, alias TEXT NOT NULL, PRIMARY KEY(locus, alias));
CREATE TABLE IF NOT EXISTS locus(
    locus_id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, is_core INTEGER NOT NULL,
    chromosome TEXT, repeat_bp INTEGER, min_allele_x10 INTEGER, max_allele_x10 INTEGER, panel TEXT);
CREATE TABLE IF NOT EXISTS case_record(
    case_id INTEGER PRIMARY KEY,
    lab_case_no TEXT, report_no TEXT, report_date TEXT,
    requisition_ref_no TEXT, requisition_date TEXT, forwarding_authority TEXT,
    fir_no TEXT, fir_date TEXT, sections_of_law TEXT,
    police_station TEXT, district TEXT, state TEXT, case_type TEXT,
    fir_norm TEXT, ps_norm TEXT, created_at TEXT);
CREATE INDEX IF NOT EXISTS ix_case_ref ON case_record(fir_norm, ps_norm);
CREATE INDEX IF NOT EXISTS ix_case_lab ON case_record(lab_case_no);
CREATE INDEX IF NOT EXISTS ix_case_geo ON case_record(state, district);
CREATE TABLE IF NOT EXISTS profile(
    profile_id INTEGER PRIMARY KEY,
    profile_ref TEXT NOT NULL UNIQUE,              -- laboratory profile number, e.g. DNA-000123
    specimen_id TEXT NOT NULL UNIQUE,              -- De-identified Code (the only identifier that leaves the laboratory)
    identifiable_code TEXT UNIQUE,                 -- internal traceable code
    case_id INTEGER REFERENCES case_record(case_id),
    person_source_id TEXT, exhibit_no TEXT, lab_sample_id TEXT, sample_type TEXT,
    specimen_category TEXT, index_group TEXT,
    collection_date TEXT, received_date TEXT, analysis_date TEXT, batch_run_id TEXT,
    kit TEXT, instrument TEXT, analysis_software TEXT,
    analyst TEXT, technical_reviewer TEXT, review_date TEXT, administrative_reviewer TEXT, administrative_review_date TEXT,
    entered_by TEXT, entry_date TEXT,
    profile_type TEXT, no_of_contributors TEXT,
    sex TEXT, amel1 TEXT, amel2 TEXT, y_indel TEXT, ystr_text TEXT, xstr_text TEXT,
    loci_typed INTEGER NOT NULL DEFAULT 0,
    genotype_key TEXT,                             -- all core loci in one string: exact whole-profile lookup
    ref_sig TEXT,                                  -- every typed locus: duplicate-with-same-reference check
    record_status TEXT, remarks TEXT,
    upload_status TEXT DEFAULT 'Not uploaded', upload_date TEXT,
    expunged_date TEXT, expunge_reason TEXT, retention_until TEXT,
    created_at TEXT, updated_at TEXT);
CREATE INDEX IF NOT EXISTS ix_profile_key ON profile(genotype_key);
CREATE INDEX IF NOT EXISTS ix_profile_sig ON profile(ref_sig);
CREATE INDEX IF NOT EXISTS ix_profile_case ON profile(case_id);
CREATE INDEX IF NOT EXISTS ix_profile_status ON profile(record_status);
CREATE INDEX IF NOT EXISTS ix_profile_cat ON profile(specimen_category);
CREATE INDEX IF NOT EXISTS ix_profile_person ON profile(person_source_id);
CREATE INDEX IF NOT EXISTS ix_profile_adate ON profile(analysis_date);
CREATE TABLE IF NOT EXISTS allele(                 -- one row per profile and locus; alleles x10 as integers (9.3 -> 93), a1 <= a2
    profile_id INTEGER NOT NULL REFERENCES profile(profile_id) ON DELETE CASCADE,
    locus_id INTEGER NOT NULL REFERENCES locus(locus_id),
    a1 INTEGER NOT NULL, a2 INTEGER NOT NULL, CHECK(a1 <= a2),
    PRIMARY KEY(profile_id, locus_id)) WITHOUT ROWID;
CREATE INDEX IF NOT EXISTS ix_allele_genotype ON allele(locus_id, a1, a2, profile_id);
CREATE TABLE IF NOT EXISTS strmarker(              -- Y-STR / X-STR results, one row per marker
    profile_id INTEGER NOT NULL REFERENCES profile(profile_id) ON DELETE CASCADE,
    kind TEXT NOT NULL CHECK(kind IN ('Y','X')), marker TEXT NOT NULL, value TEXT,
    PRIMARY KEY(profile_id, kind, marker)) WITHOUT ROWID;
CREATE INDEX IF NOT EXISTS ix_str_marker ON strmarker(kind, marker, value);
CREATE TABLE IF NOT EXISTS audit_log(
    audit_id INTEGER PRIMARY KEY, ts TEXT, user TEXT, action TEXT, table_name TEXT, record_ref TEXT,
    field TEXT, old_value TEXT, new_value TEXT, note TEXT);
CREATE INDEX IF NOT EXISTS ix_audit_ref ON audit_log(record_ref);
CREATE INDEX IF NOT EXISTS ix_audit_ts ON audit_log(ts);
CREATE TABLE IF NOT EXISTS import_log(
    import_id INTEGER PRIMARY KEY, ts TEXT, user TEXT, source TEXT, rows_read INTEGER, rows_added INTEGER, rows_not_added INTEGER, note TEXT);
CREATE TABLE IF NOT EXISTS search_log(
    search_id INTEGER PRIMARY KEY, ts TEXT, user TEXT, mode TEXT, query_loci INTEGER, min_loci INTEGER, mismatches INTEGER, results INTEGER, query_hash TEXT);
CREATE VIEW IF NOT EXISTS v_profile_full AS
    SELECT p.profile_id, p.profile_ref, p.specimen_id AS deidentified_code, p.identifiable_code, p.specimen_category AS category, p.index_group,
           p.record_status, p.kit, p.sex, p.loci_typed, p.analysis_date, p.analyst, p.technical_reviewer, p.review_date,
           c.lab_case_no, c.report_no, c.report_date, c.requisition_ref_no, c.requisition_date, c.forwarding_authority,
           c.fir_no, c.fir_date, c.sections_of_law, c.police_station, c.district, c.state, c.case_type,
           p.person_source_id, p.exhibit_no, p.lab_sample_id, p.sample_type, p.upload_status, p.remarks
    FROM profile p LEFT JOIN case_record c ON c.case_id = p.case_id;
CREATE VIEW IF NOT EXISTS v_deidentified AS      -- safe to share: no case, person, place or date information
    SELECT specimen_id AS deidentified_code, specimen_category AS category, index_group, profile_type, kit, sex, loci_typed
    FROM profile WHERE record_status NOT IN ('Rejected','Expunged / Removed','On hold / Quarantined');
CREATE VIEW IF NOT EXISTS v_allele_named AS
    SELECT p.specimen_id AS deidentified_code, l.name AS locus, a.a1 / 10.0 AS allele_1, a.a2 / 10.0 AS allele_2
    FROM allele a JOIN profile p ON p.profile_id = a.profile_id JOIN locus l ON l.locus_id = a.locus_id;
"""

def full_triggers():
    on = "COALESCE((SELECT value FROM config WHERE key='audit_enabled'),'1')='1'"
    usr = "(SELECT value FROM config WHERE key='app_user')"
    t = [f"""CREATE TRIGGER IF NOT EXISTS trg_profile_ins AFTER INSERT ON profile WHEN {on} BEGIN
        INSERT INTO audit_log(ts,user,action,table_name,record_ref,field,new_value)
        VALUES (datetime('now','localtime'), {usr}, 'INSERT', 'profile', NEW.profile_ref, 'specimen_id', NEW.specimen_id); END;""",
         f"""CREATE TRIGGER IF NOT EXISTS trg_profile_del BEFORE DELETE ON profile WHEN {on} BEGIN
        INSERT INTO audit_log(ts,user,action,table_name,record_ref,field,old_value)
        VALUES (datetime('now','localtime'), {usr}, 'DELETE', 'profile', OLD.profile_ref, 'specimen_id', OLD.specimen_id); END;""",
         f"""CREATE TRIGGER IF NOT EXISTS trg_allele_upd AFTER UPDATE ON allele WHEN (OLD.a1 IS NOT NEW.a1 OR OLD.a2 IS NOT NEW.a2) AND {on} BEGIN
        INSERT INTO audit_log(ts,user,action,table_name,record_ref,field,old_value,new_value)
        VALUES (datetime('now','localtime'), {usr}, 'UPDATE', 'allele', (SELECT profile_ref FROM profile WHERE profile_id=NEW.profile_id),
                (SELECT name FROM locus WHERE locus_id=NEW.locus_id), OLD.a1||','||OLD.a2, NEW.a1||','||NEW.a2); END;""",
         f"""CREATE TRIGGER IF NOT EXISTS trg_allele_del AFTER DELETE ON allele WHEN {on} BEGIN
        INSERT INTO audit_log(ts,user,action,table_name,record_ref,field,old_value)
        VALUES (datetime('now','localtime'), {usr}, 'DELETE', 'allele', (SELECT profile_ref FROM profile WHERE profile_id=OLD.profile_id),
                (SELECT name FROM locus WHERE locus_id=OLD.locus_id), OLD.a1||','||OLD.a2); END;"""]
    for col in AUDITED_COLS:
        t.append(f"""CREATE TRIGGER IF NOT EXISTS trg_profile_upd_{col} AFTER UPDATE OF {col} ON profile
            WHEN OLD.{col} IS NOT NEW.{col} AND {on} BEGIN
            INSERT INTO audit_log(ts,user,action,table_name,record_ref,field,old_value,new_value)
            VALUES (datetime('now','localtime'), {usr}, 'UPDATE', 'profile', NEW.profile_ref, '{col}', OLD.{col}, NEW.{col}); END;""")
    return t

def init_full_db(con):
    con.executescript(FULL_SCHEMA)
    for tr in full_triggers(): con.execute(tr)
    ensure_ystr_schema(con)
    ensure_stats_schema(con)

def current_user():
    return os.environ.get("DNA_APP_USER") or os.environ.get("USERNAME") or os.environ.get("USER") or "unknown"

def to_iso(v):
    if isinstance(v, (datetime.datetime, datetime.date)): return f"{v:%Y-%m-%d}"
    return None if v in (None, "") else str(v)

def from_iso(s):
    if isinstance(s, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", s):
        return datetime.datetime.strptime(s, "%Y-%m-%d")
    return s

def split_lab_case(txt):
    m = re.match(r"^(?P<no>.+?)\s+dt\.\s*(?P<d>\d{1,2}-\d{1,2}-\d{4})\s*$", txt or "")
    if not m: return (txt or None), None
    d = datetime.datetime.strptime(m.group("d"), "%d-%m-%Y")
    return m.group("no"), f"{d:%Y-%m-%d}"

def parse_str_text(text):
    out = []
    for part in re.split(r"[;\n]+", text or ""):
        m = re.match(r"^\s*([A-Za-z0-9_./ -]+?)\s*[=:]\s*(.+?)\s*$", part)
        if m: out.append((m.group(1), m.group(2)))
    return out

def sig_from_values(loci, v):
    parts = []
    for L in loci:
        n = L["name"]; x, y = v.get(f"{n} (1)"), v.get(f"{n} (2)")
        if x in (None, "") or y in (None, ""): continue
        if n == "Amelogenin": a, b = sorted([str(x), str(y)])
        else:
            a, b = sorted([float(x), float(y)]); a, b = fmt_num(a), fmt_num(b)
        parts.append(f"{n}:{a},{b}")
    return "|".join(parts)

def get_or_create_case(con, v, now):
    fir, ps = norm(v.get("Case/FIR No.")), norm(v.get("Police Station"))
    lab = clean(v.get("Lab Case No.")) or None
    req = clean(v.get("Requisition Ref. No.")) or None
    if not any([fir, ps, lab, req]): return None
    row = con.execute("SELECT case_id FROM case_record WHERE fir_norm=? AND ps_norm=? AND lab_case_no IS ? AND requisition_ref_no IS ?",
                      (fir, ps, lab, req)).fetchone()                  # uses ix_case_ref (fast even with lakhs of cases)
    if row: return row[0]
    rno, rdate = split_lab_case(lab)
    return con.execute(
        "INSERT INTO case_record(lab_case_no,report_no,report_date,requisition_ref_no,requisition_date,forwarding_authority,fir_no,fir_date,"
        "sections_of_law,police_station,district,state,case_type,fir_norm,ps_norm,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (lab, rno, rdate, req, to_iso(v.get("Requisition Date")), clean(v.get("Forwarding Authority")) or None,
         clean(v.get("Case/FIR No.")) or None, to_iso(v.get("FIR/Case Date")), clean(v.get("Sections of Law (U/S)")) or None,
         clean(v.get("Police Station")) or None, clean(v.get("District")) or None, clean(v.get("State")) or None,
         clean(v.get("Case Type")) or None, fir, ps, now)).lastrowid

def insert_profile(con, R, v, now, entered_by=None, codes=None):
    """Inserts one profile (header->value dict) with case, alleles and Y/X-STR rows. Returns profile_id."""
    pref = str(v["Profile ID"])
    cat = v.get("Profile Category")
    grp = R.cat_index.get(cat)
    iden, did = codes or make_codes(pref, grp, v.get("Entry Date"), v.get("Analysis Date"), R.code_keys, R.code_lab)
    if not did: raise ValueError(f"Cannot create codes for {pref} (Profile ID must look like DNA-000001, and code keys must exist)")
    case_id = get_or_create_case(con, v, now)
    lid = {n: i for i, n in con.execute("SELECT locus_id, name FROM locus")}
    geno = {}
    for L in R.loci:
        n = L["name"]
        if n == "Amelogenin": continue
        x, y = v.get(f"{n} (1)"), v.get(f"{n} (2)")
        if isinstance(x, (int, float)) and isinstance(y, (int, float)):
            geno[n] = tuple(sorted((a10(x), a10(y))))
    am1, am2 = v.get("Amelogenin (1)"), v.get("Amelogenin (2)")
    sex = "" if not (am1 and am2) else ("F" if (am1, am2) == ("X", "X") else "M" if {am1, am2} == {"X", "Y"} else "?")
    core_n = sum(1 for l in R.core if l in geno)
    ystr = clean(v.get("DYS/Y-STR")); xstr = clean(v.get("X-STR"))
    cur = con.execute(
        """INSERT INTO profile(profile_ref, specimen_id, identifiable_code, case_id, person_source_id, exhibit_no, lab_sample_id, sample_type,
           specimen_category, index_group, collection_date, received_date, analysis_date, batch_run_id, kit, instrument, analysis_software,
           analyst, technical_reviewer, review_date, entered_by, entry_date, profile_type, no_of_contributors, sex, amel1, amel2, y_indel,
           ystr_text, xstr_text, loci_typed, genotype_key, ref_sig, record_status, remarks, created_at, updated_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (pref, did, iden, case_id, clean(v.get("Person/Source ID")) or None, clean(v.get("Exhibit No.")) or None,
         clean(v.get("Lab Sample ID")) or None, clean(v.get("Sample Type")) or None, cat, grp,
         to_iso(v.get("Collection Date")), to_iso(v.get("Date Received")), to_iso(v.get("Analysis Date")), clean(v.get("Batch/Run ID")) or None,
         clean(v.get("Kit/Method")) or None, clean(v.get("Instrument")) or None, clean(v.get("Analysis Software")) or None,
         clean(v.get("Analyst")) or None, clean(v.get("Technical Reviewer")) or None, to_iso(v.get("Review Date")),
         entered_by or clean(v.get("Data Entered By")) or None, to_iso(v.get("Entry Date")), clean(v.get("Profile Type")) or None,
         clean(v.get("No. of Contributors")) or None, sex, am1 or None, am2 or None, clean(v.get("Y-indel")) or None,
         ystr or None, xstr or None, core_n, genotype_key(R.core, geno) if core_n else None, sig_from_values(R.loci, v) or None,
         clean(v.get("Record Status")) or None, clean(v.get("Remarks")) or None, now, now))
    pk = cur.lastrowid
    con.executemany("INSERT INTO allele(profile_id, locus_id, a1, a2) VALUES (?,?,?,?)", [(pk, lid[l], g[0], g[1]) for l, g in geno.items()])
    yd = ystr_normalize(parse_str_text(ystr))
    for m, val_ in yd.items():
        con.execute("INSERT OR REPLACE INTO strmarker(profile_id, kind, marker, value) VALUES (?,?,?,?)", (pk, "Y", m, val_))
    if yd:
        con.execute("INSERT OR REPLACE INTO ystr_profile(profile_id, ystr_kit, loci_typed, haplotype_key) VALUES (?,?,?,?)", (pk, ystr_kit_guess(yd), len(yd), ystr_key(yd)))
    for m, val_ in parse_str_text(xstr):
        con.execute("INSERT OR REPLACE INTO strmarker(profile_id, kind, marker, value) VALUES (?,?,?,?)", (pk, "X", m, val_))
    return pk, iden, did

class SqlRules:
    """Same attributes as Rules, but read from the reference tables of the SQLite working database."""
    def __init__(self, con):
        cfg = dict(con.execute("SELECT key, value FROM config"))
        self.mindup = int(cfg.get("min_dup_loci") or 10)
        self.code_lab = cfg.get("lab_code")
        ks = [cfg.get(f"code_key_{i}") for i in range(1, 7)]
        self.code_keys = [int(k) for k in ks] if all(k not in (None, "") for k in ks) else None
        self.case = []
        for field, dflt, addr in con.execute("SELECT field, default_value, vertical_address FROM ref_field ORDER BY ord"):
            al = [norm(r[0]) for r in con.execute("SELECT alias FROM ref_alias WHERE field=? ORDER BY ord", (field,))]
            self.case.append({"header": field, "aliases": al, "default": dflt, "addr": addr})
        order, names = [], {}
        for loc, alias in con.execute("SELECT locus, alias FROM ref_locus_alias ORDER BY ord"):
            if loc not in names: names[loc] = []; order.append(loc)
            names[loc].append(norm(alias))
        self.loci = [{"name": n, "names": names[n]} for n in order]
        self.limits = {n: (mn / 10.0, mx / 10.0) for n, mn, mx in con.execute("SELECT name, min_allele_x10, max_allele_x10 FROM locus") if mn is not None}
        self.core = [n for n, in con.execute("SELECT name FROM locus WHERE is_core=1 ORDER BY locus_id")]
        kits = {}
        for kit, mk, rule in con.execute("SELECT kit, marker, rule FROM ref_kit_rule"): kits.setdefault(kit, {})[mk] = rule
        self.kits = [(k, kits[k]) for k, in con.execute("SELECT kit FROM ref_kit ORDER BY ord") if k in kits and "R" in kits[k].values()]
        self.markers = sorted({m for r in kits.values() for m in r})
        self.locations = [(str(k).lower(), s_, d_) for k, s_, d_ in con.execute("SELECT keyword, state, district FROM ref_location ORDER BY ord") if k]
        def lst(name): return [r[0] for r in con.execute("SELECT value FROM ref_list WHERE list_name=? ORDER BY ord", (name,))]
        self.lists = {"Profile Category": [r[0] for r in con.execute("SELECT category FROM ref_category ORDER BY ord")],
                      "Kit/Method": [r[0] for r in con.execute("SELECT kit FROM ref_kit ORDER BY ord")]}
        for n in ("Case Type", "Sample Type", "Profile Type", "Record Status", "Instrument", "Analysis Software", "State"): self.lists[n] = lst(n)
        self.cat_index = dict(con.execute("SELECT category, index_group FROM ref_category"))
        self.cat_min = {c: m for c, m in con.execute("SELECT category, min_core_loci FROM ref_category") if m is not None}
        self.kit_cover = {k: n for k, n in con.execute("SELECT kit, core_loci_covered FROM ref_kit") if n is not None}
        self.states = self.lists["State"] or list(INDIAN_STATES)
        self.districts = list(dict.fromkeys(lst("District") + KERALA_DISTRICTS + TELANGANA_DISTRICTS + PUDUCHERRY_DISTRICTS + ANDAMAN_DISTRICTS))

class SqlIds:
    def __init__(self, con): self.con = con
    def __contains__(self, pid): return self.con.execute("SELECT 1 FROM profile WHERE profile_ref=?", (str(pid),)).fetchone() is not None
    def __len__(self): return self.con.execute("SELECT COUNT(*) FROM profile").fetchone()[0]
    def add(self, pid): pass
class SqlKeys:
    def __init__(self, con): self.con = con
    def __contains__(self, key): return self.con.execute("SELECT 1 FROM profile WHERE genotype_key=? LIMIT 1", (key,)).fetchone() is not None
    def __getitem__(self, key): return self.con.execute("SELECT profile_ref FROM profile WHERE genotype_key=? LIMIT 1", (key,)).fetchone()[0]
class SqlRefSig:
    SQL = ("SELECT p.profile_ref FROM profile p JOIN case_record c ON c.case_id=p.case_id "
           "WHERE c.fir_norm=? AND c.ps_norm=? AND p.ref_sig=? LIMIT 1")
    def __init__(self, con): self.con = con
    def __contains__(self, k): return self.con.execute(self.SQL, (k[0], k[1], k[2])).fetchone() is not None
    def __getitem__(self, k): return self.con.execute(self.SQL, (k[0], k[1], k[2])).fetchone()[0]
class SqlRefs:
    SQL = ("SELECT p.profile_ref FROM profile p JOIN case_record c ON c.case_id=p.case_id WHERE c.fir_norm=? AND c.ps_norm=? LIMIT 1")
    def __init__(self, con): self.con = con
    def __contains__(self, k): return self.con.execute(self.SQL, (k[0], k[1])).fetchone() is not None
    def __getitem__(self, k): return self.con.execute(self.SQL, (k[0], k[1])).fetchone()[0]

class _Cell:
    __slots__ = ("value",)
    def __init__(self, v): self.value = v

class VirtualSheet:
    """Lets the workbook-oriented code (exports, lists, query builder) read the SQLite database row by row."""
    def __init__(self, store):
        self.s = store
        self.ids = [r[0] for r in store.con.execute("SELECT profile_id FROM profile ORDER BY profile_id")]
        self._cache = (None, None)
        self.max_column = len(store.headers)
    @property
    def max_row(self): return 2 + len(self.ids)
    def cell(self, row=None, column=None):
        if row == 2: return _Cell(self.s.headers[column - 1] if 0 < column <= len(self.s.headers) else None)
        if row is None or row < 3 or row - 3 >= len(self.ids) or not (0 < column <= len(self.s.headers)): return _Cell(None)
        pid = self.ids[row - 3]
        if self._cache[0] != pid: self._cache = (pid, self.s.profile_dict(pid))
        return _Cell(self._cache[1].get(self.s.headers[column - 1]))

SQL_BASE_HEADERS = ["Profile ID", "Lab Case No.", "Requisition Ref. No.", "Requisition Date", "Forwarding Authority", "Case/FIR No.", "FIR/Case Date",
                    "Sections of Law (U/S)", "Police Station", "District", "State", "Case Type", "Profile Category", "Person/Source ID", "Exhibit No.",
                    "Lab Sample ID", "Sample Type", "Collection Date", "Date Received", "Analysis Date", "Batch/Run ID", "Kit/Method", "Instrument",
                    "Analysis Software", "Analyst", "Technical Reviewer", "Review Date", "Data Entered By", "Entry Date", "Profile Type", "No. of Contributors"]

class SqlStore:
    is_sqlite = True
    def __init__(self, path):
        self.path = path
        # the window loads / reads / writes from different threads (never at the same time): allow that
        self.con = sqlite3.connect(path, check_same_thread=False, timeout=30)
        self.con.execute("PRAGMA foreign_keys=ON"); self.con.execute("PRAGMA journal_mode=WAL")
        if not self.con.execute("SELECT 1 FROM sqlite_master WHERE name='profile'").fetchone(): raise RuntimeError("This is not a DNA working database (no profile table).")
        ensure_ystr_schema(self.con); ensure_stats_schema(self.con)
        self.rules = SqlRules(self.con)
        self.ids, self.keys = SqlIds(self.con), SqlKeys(self.con)
        self.refsig, self.refs = SqlRefSig(self.con), SqlRefs(self.con)
        self.loc_ids = {n: i for i, n in self.con.execute("SELECT locus_id, name FROM locus")}
        h = list(SQL_BASE_HEADERS) + ["Amelogenin (1)", "Amelogenin (2)"]
        for L in self.rules.loci:
            if L["name"] != "Amelogenin": h += [f"{L['name']} (1)", f"{L['name']} (2)"]
        h += ["Y-indel", "DYS/Y-STR", "X-STR", "Record Status", "Remarks", "Index Group", "Sex (from Amelogenin)", "Core Loci Typed (of 20)",
              "Identifiable Code", "De-identified Code"]
        self.headers = h
        self.hmap = {x: i + 1 for i, x in enumerate(h)}
        self.ws = VirtualSheet(self)
        self.last_row = self.ws.max_row
    def free_rows(self): return 10 ** 9
    def maxnum(self):
        r = self.con.execute("SELECT MAX(CAST(SUBSTR(profile_ref,5) AS INTEGER)) FROM profile WHERE profile_ref LIKE 'DNA-%'").fetchone()[0]
        return r or 0
    def profile_dict(self, pid):
        cur = self.con.execute("""SELECT p.*, c.lab_case_no, c.requisition_ref_no, c.requisition_date, c.forwarding_authority, c.fir_no, c.fir_date,
               c.sections_of_law, c.police_station, c.district, c.state, c.case_type
               FROM profile p LEFT JOIN case_record c ON c.case_id = p.case_id WHERE p.profile_id=?""", (pid,))
        row = cur.fetchone()
        if row is None: return {}
        r = dict(zip([c[0] for c in cur.description], row))
        d = {"Profile ID": r["profile_ref"], "Lab Case No.": r["lab_case_no"], "Requisition Ref. No.": r["requisition_ref_no"],
             "Requisition Date": from_iso(r["requisition_date"]), "Forwarding Authority": r["forwarding_authority"], "Case/FIR No.": r["fir_no"],
             "FIR/Case Date": from_iso(r["fir_date"]), "Sections of Law (U/S)": r["sections_of_law"], "Police Station": r["police_station"],
             "District": r["district"], "State": r["state"], "Case Type": r["case_type"], "Profile Category": r["specimen_category"],
             "Person/Source ID": r["person_source_id"], "Exhibit No.": r["exhibit_no"], "Lab Sample ID": r["lab_sample_id"], "Sample Type": r["sample_type"],
             "Collection Date": from_iso(r["collection_date"]), "Date Received": from_iso(r["received_date"]), "Analysis Date": from_iso(r["analysis_date"]),
             "Batch/Run ID": r["batch_run_id"], "Kit/Method": r["kit"], "Instrument": r["instrument"], "Analysis Software": r["analysis_software"],
             "Analyst": r["analyst"], "Technical Reviewer": r["technical_reviewer"], "Review Date": from_iso(r["review_date"]),
             "Data Entered By": r["entered_by"], "Entry Date": from_iso(r["entry_date"]), "Profile Type": r["profile_type"],
             "No. of Contributors": (int(r["no_of_contributors"]) if str(r["no_of_contributors"] or "").isdigit() else r["no_of_contributors"]),
             "Amelogenin (1)": r["amel1"], "Amelogenin (2)": r["amel2"], "Y-indel": (int(r["y_indel"]) if str(r["y_indel"] or "").isdigit() else r["y_indel"]),
             "DYS/Y-STR": r["ystr_text"], "X-STR": r["xstr_text"], "Record Status": r["record_status"], "Remarks": r["remarks"],
             "Index Group": r["index_group"], "Core Loci Typed (of 20)": r["loci_typed"], "Identifiable Code": r["identifiable_code"],
             "De-identified Code": r["specimen_id"]}
        names = {i: n for n, i in self.loc_ids.items()}
        for lid, x, y in self.con.execute("SELECT locus_id, a1, a2 FROM allele WHERE profile_id=?", (pid,)):
            d[f"{names[lid]} (1)"], d[f"{names[lid]} (2)"] = fmt_num(x / 10.0), fmt_num(y / 10.0)
        return d
    def profile_dict_by_ref(self, ref):
        r = self.con.execute("SELECT profile_id FROM profile WHERE profile_ref=? COLLATE NOCASE", (str(ref).strip(),)).fetchone()
        return self.profile_dict(r[0]) if r else None
    def list_profiles(self, limit=20000):
        out = []
        for pref, sid, fir, ps, cat, st_, blob in self.con.execute(
                """SELECT p.profile_ref, p.specimen_id, c.fir_no, c.police_station, p.specimen_category, p.record_status,
                   LOWER(COALESCE(p.profile_ref,'')||' '||COALESCE(p.specimen_id,'')||' '||COALESCE(c.lab_case_no,'')||' '||COALESCE(c.fir_no,'')||' '||
                   COALESCE(c.police_station,'')||' '||COALESCE(c.district,'')||' '||COALESCE(c.state,'')||' '||COALESCE(p.specimen_category,'')||' '||
                   COALESCE(p.person_source_id,'')||' '||COALESCE(p.exhibit_no,'')||' '||COALESCE(p.record_status,'')||' '||COALESCE(p.kit,''))
                   FROM profile p LEFT JOIN case_record c ON c.case_id=p.case_id ORDER BY p.profile_id DESC LIMIT ?""", (limit,)):
            out.append((pref, sid, fir or "", ps or "", cat or "", st_ or "", blob))
        return out


# ----------------------------------------------------------------------------- Y-STR haplotype database
YSTR_LOCI = [("DYS19", 1, 1), ("DYS385", 2, 1), ("DYS389I", 1, 1), ("DYS389II", 1, 1), ("DYS390", 1, 1), ("DYS391", 1, 1), ("DYS392", 1, 1),
             ("DYS393", 1, 1), ("DYS437", 1, 0), ("DYS438", 1, 0), ("DYS439", 1, 0), ("DYS448", 1, 0), ("DYS456", 1, 0), ("DYS458", 1, 0),
             ("DYS635", 1, 0), ("YGATAH4", 1, 0), ("DYS449", 1, 0), ("DYS460", 1, 0), ("DYS481", 1, 0), ("DYS518", 1, 0), ("DYS533", 1, 0),
             ("DYS549", 1, 0), ("DYS570", 1, 0), ("DYS576", 1, 0), ("DYS627", 1, 0), ("DYS643", 1, 0), ("DYF387S1", 2, 0)]
_YF = ["DYS456", "DYS389I", "DYS390", "DYS389II", "DYS458", "DYS19", "DYS385", "DYS393", "DYS391", "DYS439", "DYS635", "DYS392", "YGATAH4",
       "DYS437", "DYS438", "DYS448"]
YSTR_KITS = {"Yfiler Plus": _YF + ["DYS449", "DYS460", "DYS481", "DYS518", "DYS533", "DYS549", "DYS570", "DYS576", "DYS627", "DYF387S1"],
             "PowerPlex Y23": ["DYS19", "DYS385", "DYS389I", "DYS389II", "DYS390", "DYS391", "DYS392", "DYS393", "DYS437", "DYS438", "DYS439",
                               "DYS448", "DYS456", "DYS458", "DYS481", "DYS533", "DYS570", "DYS576", "DYS635", "DYS643", "YGATAH4"],
             "Yfiler": _YF,
             "Minimal haplotype (YHRD)": ["DYS19", "DYS385", "DYS389I", "DYS389II", "DYS390", "DYS391", "DYS392", "DYS393"]}
YSTR_ORDER = {m: i for i, (m, _c, _x) in enumerate(YSTR_LOCI)}
YSTR_SCHEMA = """
CREATE TABLE IF NOT EXISTS ref_ystr_locus(ord INTEGER, marker TEXT PRIMARY KEY, copies INTEGER NOT NULL DEFAULT 1, minimal_haplotype INTEGER NOT NULL DEFAULT 0);
CREATE TABLE IF NOT EXISTS ref_ystr_kit(kit TEXT NOT NULL, marker TEXT NOT NULL, PRIMARY KEY(kit, marker));
CREATE TABLE IF NOT EXISTS ystr_profile(
    profile_id INTEGER PRIMARY KEY REFERENCES profile(profile_id) ON DELETE CASCADE,
    ystr_kit TEXT, loci_typed INTEGER NOT NULL, haplotype_key TEXT);          -- haplotype_key: 'DYS19=14|DYS385=11,14|...' in standard order
CREATE INDEX IF NOT EXISTS ix_ystr_key ON ystr_profile(haplotype_key);
CREATE VIEW IF NOT EXISTS v_ystr_haplotype AS
    SELECT p.specimen_id AS deidentified_code, p.profile_ref, y.ystr_kit, y.loci_typed, s.marker, s.value
    FROM profile p JOIN ystr_profile y ON y.profile_id = p.profile_id JOIN strmarker s ON s.profile_id = p.profile_id AND s.kind = 'Y';
"""

def ystr_marker(name):
    s = re.sub(r"[^A-Z0-9]", "", str(name).upper())
    s = re.sub(r"^DYS0+(?=\d)", "DYS", s)
    if s in ("GATAH4", "YGATAH4"): return "YGATAH4", None
    m = re.match(r"^(DYS385|DYF387S1)(AB|A|B)?$", s)
    if m: return m.group(1), m.group(2)
    m = re.match(r"^DYS389(I|II|1|2)$", s)
    if m: return ("DYS389I" if m.group(1) in ("I", "1") else "DYS389II"), None
    return s, None

def _ystr_val(txt):
    toks = [t for t in re.split(r"[,/ ;]+", str(txt).strip()) if t]
    out = []
    for t in toks:
        try: out.append(fmt_num(float(t)))
        except ValueError: out.append(t.upper())
    return out

def ystr_normalize(pairs):
    """[(name, value text)] -> {standard marker: normalised value}; multi-copy loci (DYS385, DYF387S1) as sorted 'a,b'."""
    got = {}
    for name, val in pairs:
        marker, copy = ystr_marker(name)
        vals = [v for v in _ystr_val(val) if v not in ("-", "NR", "NA", "")]
        if vals: got.setdefault(marker, []).extend(vals)
    out = {}
    for marker, vals in got.items():
        multi = dict((m, c) for m, c, _x in YSTR_LOCI).get(marker, 1) > 1
        if multi or len(vals) > 1:
            try: vals = sorted(vals, key=float)
            except ValueError: vals = sorted(vals)
            out[marker] = ",".join(str(x) for x in vals)
        else:
            out[marker] = str(vals[0])
    return out

def ystr_kit_guess(markers):
    typed = set(markers)
    best, cov = None, 0.0
    for kit, ms in YSTR_KITS.items():
        if typed <= set(ms):
            c = len(typed) / len(ms)
            if c > cov: best, cov = kit, c
    return best if best and cov >= 0.5 else (f"Partial ({len(typed)} loci)" if typed else None)

def ystr_key(d):
    return "|".join(f"{m}={d[m]}" for m in sorted(d, key=lambda x: YSTR_ORDER.get(x, 999)))

def ensure_ystr_schema(con):
    con.executescript(YSTR_SCHEMA)
    if con.execute("SELECT COUNT(*) FROM ref_ystr_locus").fetchone()[0] == 0:
        con.executemany("INSERT INTO ref_ystr_locus VALUES (?,?,?,?)", [(i, m, c, x) for i, (m, c, x) in enumerate(YSTR_LOCI)])
        for kit, ms in YSTR_KITS.items(): con.executemany("INSERT OR IGNORE INTO ref_ystr_kit VALUES (?,?)", [(kit, m) for m in ms])
    # back-fill for profiles that were loaded before the Y-STR tables existed
    todo = con.execute("SELECT profile_id, ystr_text FROM profile WHERE ystr_text IS NOT NULL AND ystr_text<>'' "
                       "AND profile_id NOT IN (SELECT profile_id FROM ystr_profile)").fetchall()
    for pk, txt in todo:
        d = ystr_normalize(parse_str_text(txt))
        con.execute("DELETE FROM strmarker WHERE profile_id=? AND kind='Y'", (pk,))
        for m, v in d.items(): con.execute("INSERT OR REPLACE INTO strmarker VALUES (?,?,?,?)", (pk, "Y", m, v))
        if d: con.execute("INSERT OR REPLACE INTO ystr_profile VALUES (?,?,?,?)", (pk, ystr_kit_guess(d), len(d), ystr_key(d)))
    con.commit()

def parse_haplotype(text):
    pairs = []
    for line in re.split(r"[;\n]+", text or ""):
        line = line.strip()
        if not line: continue
        m = re.match(r"^([A-Za-z0-9_./ -]+?)\s*[=:\t ]\s*(\d[\d.,/ ]*)$", line)
        if not m: raise ValueError(f"Cannot read '{line}'. Use e.g. DYS391 10  or  DYS385 11,14")
        pairs.append((m.group(1), m.group(2)))
    d = ystr_normalize(pairs)
    if not d: raise ValueError("No Y-STR values found in the query.")
    return d

def ystr_search(path, query, mismatches=0, min_loci=1, limit=100):
    """Y-STR haplotype search. query: {marker: value}. Returns (rows, info); each row = (specimen_id, ystr_kit, compared, identical).
    info['exact'] = number of profiles identical at every queried locus (the haplotype count used to report a Y-STR match)."""
    con = sqlite3.connect(path)
    ensure_ystr_schema(con)
    nq = len(query)
    if nq < min_loci: raise ValueError(f"Only {nq} locus/loci in the query - minimum loci compared is {min_loci}.")
    con.execute("CREATE TEMP TABLE q(marker TEXT PRIMARY KEY, value TEXT)")
    con.executemany("INSERT INTO q VALUES (?,?)", list(query.items()))
    where, args = "", []
    if mismatches == 0:   # exact: start from the rarest queried marker value (index lookup)
        best = min(((con.execute("SELECT COUNT(*) FROM strmarker WHERE kind='Y' AND marker=? AND value=?", (m, v)).fetchone()[0], m, v)
                    for m, v in query.items()))
        where = "AND s.profile_id IN (SELECT profile_id FROM strmarker WHERE kind='Y' AND marker=? AND value=?)"; args = [best[1], best[2]]
    sql = f"""SELECT p.specimen_id, y.ystr_kit, t.comp, t.ident FROM (
                SELECT s.profile_id AS pid, COUNT(*) AS comp, SUM(s.value = q.value) AS ident
                FROM strmarker s JOIN q ON q.marker = s.marker WHERE s.kind='Y' {where}
                GROUP BY s.profile_id HAVING comp >= ? AND comp - ident <= ?) t
              JOIN profile p ON p.profile_id = t.pid LEFT JOIN ystr_profile y ON y.profile_id = t.pid
              ORDER BY t.ident DESC, t.comp DESC, p.specimen_id"""
    allrows = con.execute(sql, args + [min_loci, mismatches]).fetchall()
    exact = sum(1 for r in allrows if r[3] == nq and r[2] == nq)
    try:
        if con.execute("SELECT 1 FROM sqlite_master WHERE name='search_log'").fetchone():
            con.execute("INSERT INTO search_log(ts,user,mode,query_loci,min_loci,mismatches,results,query_hash) VALUES (?,?,?,?,?,?,?,?)",
                        (datetime.datetime.now().isoformat(timespec="seconds"), current_user(), "Y-STR", nq, min_loci, mismatches, len(allrows), ""))
            con.commit()
    except sqlite3.Error: pass
    con.close()
    return allrows[:limit], {"query_loci": nq, "exact": exact, "total": len(allrows)}

def export_ystr(path, out_dir, ids=None):
    """De-identified Y-STR haplotype table (one row per profile, one column per standard Y-STR locus) as Excel."""
    from openpyxl import Workbook
    con = sqlite3.connect(path); ensure_ystr_schema(con)
    want = {str(i).strip().upper() for i in ids} if ids else None
    rows = []
    for pk, sid, pref, kit, n in con.execute("SELECT p.profile_id, p.specimen_id, p.profile_ref, y.ystr_kit, y.loci_typed FROM profile p "
                                             "JOIN ystr_profile y ON y.profile_id=p.profile_id WHERE p.record_status NOT IN ('Rejected','Expunged / Removed','On hold / Quarantined') "
                                             "OR p.record_status IS NULL ORDER BY p.profile_id"):
        if want is not None and sid.upper() not in want and str(pref).upper() not in want: continue
        d = dict(con.execute("SELECT marker, value FROM strmarker WHERE profile_id=? AND kind='Y'", (pk,)).fetchall())
        rows.append([sid, kit, n] + [d.get(m) for m, _c, _x in YSTR_LOCI] + [", ".join(f"{k}={v}" for k, v in d.items() if k not in YSTR_ORDER)])
    con.close()
    if not rows: raise RuntimeError("No Y-STR profiles found.")
    os.makedirs(out_dir, exist_ok=True)
    p = os.path.join(out_dir, f"YSTR_Haplotypes_{datetime.datetime.now():%Y%m%d_%H%M}.xlsx")
    wb = Workbook(); ws = wb.active; ws.title = "Y-STR haplotypes"
    ws.append(["De-identified Code", "Y-STR kit", "Loci typed"] + [m + (" (a,b)" if c > 1 else "") for m, c, _x in YSTR_LOCI] + ["Other markers"])
    for r in rows: ws.append(r)
    ws.freeze_panes = "D2"
    info = wb.create_sheet("Info")
    for t in (["Contents", "De-identified Code and Y-STR haplotype only. No case, person, place or date information."],
              ["Multi-copy loci", "DYS385 and DYF387S1 are written as 'a,b' (sorted)."],
              ["Exported on", datetime.datetime.now().strftime("%d-%m-%Y %H:%M")], ["Profiles", len(rows)]): info.append(t)
    wb.save(p)
    return p, len(rows)

# ----------------------------------------------------------------------------- browsing the database
def browse_rows(store, text="", offset=0, limit=500):
    """(total, rows) for the browser window: rows = (profile_ref, de-id code, FIR, police station, state, category, kit, loci typed, status, analysis date)."""
    words = [w.lower() for w in (text or "").split()]
    if getattr(store, "is_sqlite", False):
        expr = ("LOWER(COALESCE(p.profile_ref,'')||' '||COALESCE(p.specimen_id,'')||' '||COALESCE(c.lab_case_no,'')||' '||COALESCE(c.fir_no,'')||' '||"
                "COALESCE(c.police_station,'')||' '||COALESCE(c.district,'')||' '||COALESCE(c.state,'')||' '||COALESCE(p.specimen_category,'')||' '||"
                "COALESCE(p.person_source_id,'')||' '||COALESCE(p.exhibit_no,'')||' '||COALESCE(p.record_status,'')||' '||COALESCE(p.kit,'')||' '||"
                "COALESCE(p.sample_type,'')||' '||COALESCE(p.analyst,'')||' '||COALESCE(c.sections_of_law,'')||' '||COALESCE(c.forwarding_authority,''))")
        cond = " AND ".join([f"{expr} LIKE ?"] * len(words)) or "1=1"
        args = [f"%{w}%" for w in words]
        frm = "FROM profile p LEFT JOIN case_record c ON c.case_id=p.case_id"
        total = store.con.execute(f"SELECT COUNT(*) {frm} WHERE {cond}", args).fetchone()[0]
        rows = store.con.execute(
            f"SELECT p.profile_ref, p.specimen_id, c.fir_no, c.police_station, c.state, p.specimen_category, p.kit, p.loci_typed, p.record_status, p.analysis_date "
            f"{frm} WHERE {cond} ORDER BY p.profile_id DESC LIMIT ? OFFSET ?", args + [limit, offset]).fetchall()
        return total, rows
    allr = [r for r in list_profiles(store) if all(w in r[6] for w in words)]
    out = []
    for r in allr[offset:offset + limit]:
        out.append((r[0], r[1], r[2], r[3], "", r[4], "", "", r[5], ""))
    return len(allr), out

def profile_detail(store, ref):
    """Everything known about one profile: list of (section, field, value)."""
    if getattr(store, "is_sqlite", False):
        d = store.profile_dict_by_ref(ref)
    else:
        d = None
        for r in range(3, store.last_row + 1):
            if str(store.ws.cell(row=r, column=1).value).upper() == str(ref).upper():
                d = {h: store.ws.cell(row=r, column=c).value for h, c in store.hmap.items()}; break
    if d is None: raise ValueError(f"Profile {ref} not found")
    sections = [("Identification", ["Profile ID", "De-identified Code", "Identifiable Code", "Profile Category", "Index Group", "Record Status"]),
                ("Case", ["Lab Case No.", "Requisition Ref. No.", "Requisition Date", "Forwarding Authority", "Case/FIR No.", "FIR/Case Date",
                          "Sections of Law (U/S)", "Police Station", "District", "State", "Case Type"]),
                ("Sample", ["Person/Source ID", "Exhibit No.", "Lab Sample ID", "Sample Type", "Collection Date", "Date Received"]),
                ("Laboratory", ["Analysis Date", "Batch/Run ID", "Kit/Method", "Instrument", "Analysis Software", "Analyst", "Technical Reviewer",
                                "Review Date", "Data Entered By", "Entry Date", "Profile Type", "No. of Contributors", "Core Loci Typed (of 20)", "Remarks"])]
    out = []
    def fmt(v):
        if isinstance(v, (datetime.datetime, datetime.date)): return f"{v:%d-%m-%Y}"
        return "" if v is None else str(v)
    for title, fields in sections:
        for f in fields:
            if f in d: out.append((title, f, fmt(d[f])))
    names = [L["name"] for L in store.rules.loci]
    for n in names:
        a, b = d.get(f"{n} (1)"), d.get(f"{n} (2)")
        if a not in (None, "") or b not in (None, ""): out.append(("STR alleles", n, f"{fmt(a)}, {fmt(b)}"))
    if d.get("Y-indel") not in (None, ""): out.append(("STR alleles", "Y-indel", fmt(d["Y-indel"])))
    if getattr(store, "is_sqlite", False):
        pk = store.con.execute("SELECT profile_id FROM profile WHERE profile_ref=?", (d["Profile ID"],)).fetchone()[0]
        for m, v in store.con.execute("SELECT marker, value FROM strmarker WHERE profile_id=? AND kind='Y' ORDER BY marker", (pk,)):
            out.append(("Y-STR", m, v))
        for m, v in store.con.execute("SELECT marker, value FROM strmarker WHERE profile_id=? AND kind='X' ORDER BY marker", (pk,)):
            out.append(("X-STR", m, v))
        for ts, usr, act, fld, ov, nv in store.con.execute("SELECT ts, user, action, field, old_value, new_value FROM audit_log WHERE record_ref=? ORDER BY audit_id DESC LIMIT 50", (d["Profile ID"],)):
            out.append(("History", f"{ts}  {usr or ''}", f"{act} {fld or ''}: {ov or ''} -> {nv or ''}"))
    else:
        if d.get("DYS/Y-STR"): out.append(("Y-STR", "text", str(d["DYS/Y-STR"])))
    return out



# ----------------------------------------------------------------------------- forensic match statistics
STATS_SCHEMA = """
CREATE TABLE IF NOT EXISTS population(pop_id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, source TEXT, n_individuals INTEGER, created TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS allele_freq(
    pop_id INTEGER NOT NULL REFERENCES population(pop_id) ON DELETE CASCADE,
    locus_id INTEGER NOT NULL REFERENCES locus(locus_id),
    allele INTEGER NOT NULL,                        -- allele x 10
    count INTEGER NOT NULL, n_alleles INTEGER NOT NULL, freq REAL NOT NULL,
    PRIMARY KEY(pop_id, locus_id, allele)) WITHOUT ROWID;
"""
DEFAULT_THETA = 0.01
KINSHIP = {"parent-child": (0.0, 1.0, 0.0), "full-siblings": (0.25, 0.5, 0.25), "half-siblings": (0.5, 0.5, 0.0),
           "uncle-nephew": (0.5, 0.5, 0.0), "grandparent-grandchild": (0.5, 0.5, 0.0), "first-cousins": (0.75, 0.25, 0.0),
           "unrelated": (1.0, 0.0, 0.0)}

def ensure_stats_schema(con):
    con.executescript(STATS_SCHEMA)
    for k, v in (("stats_theta", str(DEFAULT_THETA)), ("stats_min_freq_rule", "5/2N"), ("stats_mutation_rate", "0.001")):
        con.execute("INSERT OR IGNORE INTO config VALUES (?,?)", (k, v))
    con.commit()

def build_allele_frequencies(path, name, all_profiles=False, source="internal database", note=""):
    """Allele frequencies from this database. Default population = people (offender / suspect / reference / elimination categories),
    one profile per person (Person/Source ID), so crime-scene and duplicate profiles do not distort the frequencies."""
    con = sqlite3.connect(path); ensure_stats_schema(con)
    if all_profiles:
        sel = "SELECT profile_id FROM profile"
    else:
        sel = ("SELECT MIN(p.profile_id) FROM profile p JOIN ref_category c ON c.category = p.specimen_category "
               "WHERE c.index_group IN ('Offender / Suspect','Reference / Elimination') "
               "GROUP BY COALESCE(p.person_source_id, 'P' || p.profile_id)")
    con.execute("DELETE FROM population WHERE name=?", (name,))
    n = con.execute(f"SELECT COUNT(*) FROM ({sel})").fetchone()[0]
    if n == 0: raise ValueError("No profiles qualify for the population (set categories, or use all profiles).")
    pid = con.execute("INSERT INTO population(name, source, n_individuals, created, note) VALUES (?,?,?,?,?)",
                      (name, source, n, datetime.datetime.now().isoformat(timespec="seconds"), note or ("all profiles" if all_profiles else "one profile per person; offender/suspect/reference/elimination"))).lastrowid
    con.execute(f"""INSERT INTO allele_freq(pop_id, locus_id, allele, count, n_alleles, freq)
        SELECT ?, t.locus_id, t.allele, t.c, tot.n, 1.0 * t.c / tot.n FROM
          (SELECT locus_id, allele, COUNT(*) AS c FROM
              (SELECT locus_id, a1 AS allele FROM allele WHERE profile_id IN ({sel})
               UNION ALL SELECT locus_id, a2 FROM allele WHERE profile_id IN ({sel}))
           GROUP BY locus_id, allele) t
          JOIN (SELECT locus_id, 2 * COUNT(*) AS n FROM allele WHERE profile_id IN ({sel}) GROUP BY locus_id) tot ON tot.locus_id = t.locus_id""", (pid,))
    con.commit()
    k = con.execute("SELECT COUNT(*) FROM allele_freq WHERE pop_id=?", (pid,)).fetchone()[0]
    con.close()
    return n, k

def import_allele_frequencies(path, name, csv_path, n_individuals, source="published"):
    """Published frequency table (CSV columns: locus, allele, frequency) -> population. Alleles may be like 9.3."""
    import csv
    con = sqlite3.connect(path); ensure_stats_schema(con)
    lid = {n: i for i, n in con.execute("SELECT locus_id, name FROM locus")}
    con.execute("DELETE FROM population WHERE name=?", (name,))
    pid = con.execute("INSERT INTO population(name, source, n_individuals, created, note) VALUES (?,?,?,?,?)",
                      (name, source, n_individuals, datetime.datetime.now().isoformat(timespec="seconds"), os.path.basename(csv_path))).lastrowid
    rows = 0
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            loc = next((n for n in lid if norm(n) == norm(r.get("locus") or r.get("Locus") or "")), None)
            if loc is None: continue
            fr = float(r.get("frequency") or r.get("Frequency")); al = a10(r.get("allele") or r.get("Allele"))
            con.execute("INSERT OR REPLACE INTO allele_freq VALUES (?,?,?,?,?,?)", (pid, lid[loc], al, round(fr * 2 * n_individuals), 2 * n_individuals, fr)); rows += 1
    con.commit(); con.close()
    return rows

class FreqTable:
    def __init__(self, path, population):
        con = sqlite3.connect(path)
        row = con.execute("SELECT pop_id, n_individuals FROM population WHERE name=?", (population,)).fetchone()
        if not row: raise ValueError(f"Population '{population}' not found - build it first.")
        self.pop_id, self.n = row
        self.theta = float(dict(con.execute("SELECT key, value FROM config")).get("stats_theta", DEFAULT_THETA))
        names = dict(con.execute("SELECT locus_id, name FROM locus"))
        self.f = {}; self.nall = {}
        for lid, al, fr, na in con.execute("SELECT locus_id, allele, freq, n_alleles FROM allele_freq WHERE pop_id=?", (self.pop_id,)):
            self.f[(names[lid], al)] = fr; self.nall[names[lid]] = na
        con.close()
    def p(self, locus, allele):
        na = self.nall.get(locus)
        if not na: raise ValueError(f"No frequencies for locus {locus}")
        return max(self.f.get((locus, allele), 0.0), 5.0 / na)          # minimum allele frequency 5 / 2N (NRC II)

def match_probability(ft, locus, g, theta=None):
    """Random match probability of genotype g=(a,b) at one locus, NRC II recommendation 4.10 (theta = FST sub-population correction)."""
    th = ft.theta if theta is None else theta
    a, b = g; pa, pb = ft.p(locus, a), ft.p(locus, b)
    den = (1 + th) * (1 + 2 * th)
    if a == b: return (2 * th + (1 - th) * pa) * (3 * th + (1 - th) * pa) / den
    return 2 * (th + (1 - th) * pa) * (th + (1 - th) * pb) / den

def rmp(ft, geno, theta=None):
    per = {l: match_probability(ft, l, g, theta) for l, g in geno.items()}
    prod = 1.0
    for v in per.values(): prod *= v
    return prod, per

def kinship_lr_locus(ft, g1, g2, relation, locus, mu=0.001):
    """LR = P(g1,g2 | related) / P(g1,g2 | unrelated) for one locus, from the IBD coefficients (k0,k1,k2)."""
    k0, k1, k2 = KINSHIP[relation]
    (x, y), (a, b) = g1, g2
    pa, pb = ft.p(locus, a), ft.p(locus, b)
    pg2 = pa * pa if a == b else 2 * pa * pb
    if a == b: p1 = 0.5 * (x == a) * pa + 0.5 * (y == a) * pa
    else: p1 = 0.5 * ((x == a) * pb + (y == a) * pb + (x == b) * pa + (y == b) * pa)
    p2 = 1.0 if sorted(g1) == sorted(g2) else 0.0
    lr = k0 + k1 * p1 / pg2 + k2 * p2 / pg2
    if relation == "parent-child" and lr == 0: lr = mu       # simple mutation allowance: a locus mismatch is not an exclusion
    return lr

def kinship_lr(ft, geno1, geno2, relation, mu=0.001):
    per = {l: kinship_lr_locus(ft, geno1[l], geno2[l], relation, l, mu) for l in geno1 if l in geno2}
    prod = 1.0
    for v in per.values(): prod *= v
    return prod, per

def _gammainc_upper_reg(a, x):
    if x <= 0: return 1.0
    if x < a + 1:                       # series for the lower incomplete gamma
        term = s = 1.0 / a; n = a
        for _ in range(500):
            n += 1; term *= x / n; s += term
            if abs(term) < abs(s) * 1e-14: break
        return 1.0 - s * math.exp(-x + a * math.log(x) - math.lgamma(a))
    b = x + 1 - a; c = 1e300; d = 1 / b; h = d     # continued fraction
    for i in range(1, 500):
        an = -i * (i - a); b += 2; d = an * d + b; d = 1e-300 if abs(d) < 1e-300 else d
        c = b + an / c; c = 1e-300 if abs(c) < 1e-300 else c
        d = 1 / d; de = d * c; h *= de
        if abs(de - 1) < 1e-14: break
    return math.exp(-x + a * math.log(x) - math.lgamma(a)) * h

def chi2_sf(x, df):
    return _gammainc_upper_reg(df / 2.0, x / 2.0)

def hwe_permutation_p(gens, perms=2000, seed=12345):
    """Monte-Carlo (allele-permutation) test of Hardy-Weinberg proportions, conditional on the observed allele counts.
    Statistic: sum over genotype classes of (O - E)^2 / E. Needs numpy; returns None if numpy is not installed."""
    try:
        import numpy as np
    except ImportError:
        return None
    als = sorted({x for a1_, a2_, _c in gens for x in (a1_, a2_)}); idx = {x: i for i, x in enumerate(als)}; k = len(als)
    alleles = []
    for a1_, a2_, c in gens: alleles += [idx[a1_], idx[a2_]] * c        # genotype pairs, in order
    arr = np.array(alleles); N = len(arr) // 2
    cnt = np.bincount(arr, minlength=k); p = cnt / (2.0 * N)
    E = np.zeros((k, k))
    for i in range(k):
        for j in range(i, k): E[i, j] = N * (p[i] ** 2 if i == j else 2 * p[i] * p[j])
    mask = E > 0
    def stat(a_):
        g = a_.reshape(-1, 2); lo = g.min(axis=1); hi = g.max(axis=1)
        O = np.bincount(lo * k + hi, minlength=k * k).reshape(k, k)
        return float((((O - E) ** 2)[mask] / E[mask]).sum())
    obs = stat(arr); rng = np.random.default_rng(seed); ge = 0
    for _ in range(perms):
        rng.shuffle(arr); ge += stat(arr) >= obs - 1e-9
    return (ge + 1.0) / (perms + 1.0)

def locus_statistics(path, all_profiles=False):
    """Per-locus forensic parameters from the database: N, alleles, Ho, He, PIC, PD (power of discrimination), HWE chi-square p."""
    con = sqlite3.connect(path)
    if all_profiles: sel = "SELECT profile_id FROM profile"
    else:
        sel = ("SELECT MIN(p.profile_id) FROM profile p JOIN ref_category c ON c.category = p.specimen_category "
               "WHERE c.index_group IN ('Offender / Suspect','Reference / Elimination') GROUP BY COALESCE(p.person_source_id, 'P' || p.profile_id)")
    out = []
    for lid, name in con.execute("SELECT locus_id, name FROM locus ORDER BY locus_id").fetchall():
        gens = con.execute(f"SELECT a1, a2, COUNT(*) FROM allele WHERE locus_id=? AND profile_id IN ({sel}) GROUP BY a1, a2", (lid,)).fetchall()
        N = sum(c for _a, _b, c in gens)
        if N < 5: continue
        cnt = {}
        for a1_, a2_, c in gens:
            cnt[a1_] = cnt.get(a1_, 0) + c; cnt[a2_] = cnt.get(a2_, 0) + c
        p = {al: c / (2.0 * N) for al, c in cnt.items()}
        het = sum(c for a1_, a2_, c in gens if a1_ != a2_)
        ho = het / N
        he = (2.0 * N / (2.0 * N - 1)) * (1 - sum(v * v for v in p.values()))
        als = sorted(p)
        pic = 1 - sum(v * v for v in p.values()) - sum(2 * p[als[i]] ** 2 * p[als[j]] ** 2 for i in range(len(als)) for j in range(i + 1, len(als)))
        pd = 1 - sum((p[als[i]] ** 2) ** 2 for i in range(len(als))) - sum((2 * p[als[i]] * p[als[j]]) ** 2 for i in range(len(als)) for j in range(i + 1, len(als)))
        obs = {(a1_, a2_): c for a1_, a2_, c in gens}
        pooled_o = pooled_e = 0.0; chi = 0.0; classes = 0
        for i in range(len(als)):
            for j in range(i, len(als)):
                e = N * (p[als[i]] ** 2 if i == j else 2 * p[als[i]] * p[als[j]])
                o = obs.get((als[i], als[j]), 0)
                if e < 5: pooled_o += o; pooled_e += e
                else: chi += (o - e) ** 2 / e; classes += 1
        if pooled_e >= 1e-9:
            if pooled_e >= 1: chi += (pooled_o - pooled_e) ** 2 / pooled_e; classes += 1
        df = classes - 1 - (len(als) - 1)
        hwe_chi = chi2_sf(chi, df) if df >= 1 else None
        hwe_mc = hwe_permutation_p(gens) if N <= 20000 else None
        out.append({"locus": name, "N": N, "alleles": len(als), "Ho": ho, "He": he, "PIC": pic, "PD": pd,
                    "HWE_p": hwe_mc if hwe_mc is not None else hwe_chi, "HWE_p_chi2": hwe_chi, "chi2": chi, "df": df})
    con.close()
    return out

def ystr_frequency(path, query):
    """Counting-method Y-STR haplotype frequency. N = profiles typed at every queried locus; x = those identical.
    Reports x/N, the (x+1)/(N+1) estimate and the one-sided 95% upper confidence limit (exact binomial)."""
    con = sqlite3.connect(path); ensure_ystr_schema(con)
    markers = list(query)
    ph = ",".join("?" * len(markers))
    N = con.execute(f"SELECT COUNT(*) FROM (SELECT profile_id FROM strmarker WHERE kind='Y' AND marker IN ({ph}) GROUP BY profile_id HAVING COUNT(*)=?)",
                    markers + [len(markers)]).fetchone()[0]
    q = ["kind='Y'", "(" + " OR ".join(["(marker=? AND value=?)"] * len(markers)) + ")"]
    args = []
    for m, v in query.items(): args += [m, v]
    x = con.execute(f"SELECT COUNT(*) FROM (SELECT profile_id FROM strmarker WHERE {' AND '.join(q)} GROUP BY profile_id HAVING COUNT(*)=?)", args + [len(markers)]).fetchone()[0]
    con.close()
    return {"N": N, "x": x, "freq": (x / N if N else None), "estimate": ((x + 1) / (N + 1) if N else None), "upper95": binom_upper(x, N) if N else None}

def binom_upper(x, n, conf=0.95):
    """One-sided exact (Clopper-Pearson) upper bound for a proportion: smallest p with P(X <= x | p) <= 1 - conf."""
    if x >= n: return 1.0
    alpha = 1 - conf
    def cdf(p):                      # P(X <= x) for Binomial(n, p), computed in log space
        s = 0.0
        for k in range(x + 1):
            s += math.exp(math.lgamma(n + 1) - math.lgamma(k + 1) - math.lgamma(n - k + 1) + k * math.log(p) + (n - k) * math.log1p(-p))
        return s
    lo, hi = 0.0, 1.0
    for _ in range(80):
        mid = (lo + hi) / 2
        if cdf(mid) > alpha: lo = mid
        else: hi = mid
    return hi

def ystr_diversity(path):
    """Haplotype diversity and discrimination capacity over profiles typed at the 'minimal haplotype' set."""
    con = sqlite3.connect(path)
    rows = con.execute("SELECT haplotype_key FROM ystr_profile WHERE haplotype_key IS NOT NULL").fetchall()
    con.close()
    n = len(rows)
    if n < 2: return {"N": n}
    cnt = {}
    for (k,) in rows: cnt[k] = cnt.get(k, 0) + 1
    return {"N": n, "distinct": len(cnt), "HD": n / (n - 1) * (1 - sum((c / n) ** 2 for c in cnt.values())), "DC": len(cnt) / n,
            "max_count": max(cnt.values()), "singletons": sum(1 for c in cnt.values() if c == 1)}

def profile_genotype(path, profile_ref):
    con = sqlite3.connect(path)
    r = con.execute("SELECT profile_id FROM profile WHERE profile_ref=? COLLATE NOCASE OR specimen_id=? COLLATE NOCASE", (profile_ref, profile_ref)).fetchone()
    if not r: raise ValueError(f"Profile {profile_ref} not found")
    g = {n: (a, b) for n, a, b in con.execute("SELECT l.name, a.a1, a.a2 FROM allele a JOIN locus l USING(locus_id) WHERE a.profile_id=? ORDER BY l.locus_id", (r[0],))}
    con.close(); return g

def match_statistics(path, hits, query, population):
    """For each search hit add the random match probability and LR over the loci where hit and query are identical."""
    ft = FreqTable(path, population)
    out = {}
    for sid in hits:
        g = profile_genotype(path, sid)
        same = {l: g[l] for l in query if l in g and tuple(sorted(g[l])) == tuple(sorted(query[l]))}
        if same:
            p, _ = rmp(ft, same); out[sid] = (p, 1.0 / p, len(same))
    return out


def new_empty_database(dst, template, lab_code=None):
    """Creates a NEW empty SQLite working database with the full schema, copying the reference tables (categories, lists, kit rules,
    locations, header aliases, locus limits) from a template database, with fresh secret code keys for the new laboratory."""
    if os.path.exists(dst): raise FileExistsError(f"{dst} already exists - the tool never overwrites a database.")
    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    con = sqlite3.connect(dst); init_full_db(con)
    con.execute("ATTACH DATABASE ? AS src", (template,))
    for t in ("ref_category", "ref_list", "ref_kit", "ref_kit_rule", "ref_location", "ref_field", "ref_alias", "ref_locus_alias", "locus"):
        con.execute(f"DELETE FROM main.{t}"); con.execute(f"INSERT INTO main.{t} SELECT * FROM src.{t}")
    cfg = dict(con.execute("SELECT key, value FROM src.config"))
    con.execute("DELETE FROM main.ref_ystr_locus"); con.execute("INSERT INTO main.ref_ystr_locus SELECT * FROM src.ref_ystr_locus")
    con.execute("DELETE FROM main.ref_ystr_kit"); con.execute("INSERT INTO main.ref_ystr_kit SELECT * FROM src.ref_ystr_kit")
    keys = new_code_keys()
    new = {"lab_code": lab_code or cfg.get("lab_code", "LAB"), "min_dup_loci": cfg.get("min_dup_loci", "10"), "audit_enabled": "1", "app_user": current_user(),
           "core_loci_count": cfg.get("core_loci_count", "20")}
    for i, k in enumerate(keys, 1): new[f"code_key_{i}"] = str(k)
    for k, v in new.items(): con.execute("INSERT OR REPLACE INTO config VALUES (?,?)", (k, v))
    now = datetime.datetime.now().isoformat(timespec="seconds")
    for k, v in (("schema_version", FULL_SCHEMA_VERSION), ("allele_encoding", "allele x 10 as integer (9.3 -> 93), a1 <= a2"), ("created", now), ("template", os.path.basename(template))):
        con.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", (k, v))
    con.commit(); con.execute("DETACH DATABASE src"); con.close()
    return dst

def open_database(path):
    return SqlStore(path) if str(path).lower().endswith((".db", ".sqlite", ".sqlite3")) else Database(path)

def commit_sqlite(store, results, include_review=False, include_incomplete=True, user=None, source=""):
    rows = selected(results, include_incomplete, include_review)
    if not rows: return 0, None
    con, R = store.con, store.rules
    if not R.code_keys: raise RuntimeError("The database has no code keys (config table).")
    user = user or current_user()
    now = datetime.datetime.now().isoformat(timespec="seconds")
    num = store.maxnum()
    con.execute("INSERT OR REPLACE INTO config VALUES ('app_user', ?)", (user,))
    try:
        for x in rows:
            v = dict(x["values"])
            if not v.get("Profile ID"):
                num += 1; v["Profile ID"] = f"DNA-{num:06d}"
            pk, iden, did = insert_profile(con, R, v, now, entered_by=user)
            x["assigned_id"] = v["Profile ID"]; x["codes"] = (iden, did)
        not_added = len(results) - len(rows)
        con.execute("INSERT INTO import_log(ts,user,source,rows_read,rows_added,rows_not_added,note) VALUES (?,?,?,?,?,?,?)",
                    (now, user, source, len(results), len(rows), not_added, None))
        con.commit()
    except Exception:
        con.rollback(); raise
    return len(rows), None

def _nv(x):
    if x in (None, ""): return None
    if isinstance(x, datetime.datetime): return x.date()
    if isinstance(x, datetime.date): return x
    if isinstance(x, (int, float)) and not isinstance(x, bool): return fmt_num(x)
    s = str(x).strip()
    return s if s else None

def convert_workbook_to_sqlite(xlsx, out_db, lab_code="CFSLH", progress=None):
    """Creates a NEW SQLite working database from the Excel workbook: every profile with all its details, plus the reference
    tables (lists, aliases, kit rules, locations, locus limits, code keys). Verifies the result field by field."""
    if os.path.exists(out_db): raise FileExistsError(f"{out_db} already exists - choose a new file name (the converter never overwrites).")
    wdb = Database(xlsx); R = wdb.rules
    wb = wdb.wb
    keys, lab = R.code_keys, R.code_lab
    if not keys: keys, lab = new_code_keys(), lab_code
    R.code_keys, R.code_lab = keys, lab
    os.makedirs(os.path.dirname(os.path.abspath(out_db)), exist_ok=True)
    con = sqlite3.connect(out_db)
    init_full_db(con)
    now = datetime.datetime.now().isoformat(timespec="seconds")
    st = wb["Settings"]
    with con:
        cfg = {"lab_code": lab, "min_dup_loci": str(R.mindup), "audit_enabled": "0", "app_user": current_user(),
               "core_loci_count": str(len(R.core)), "converted_from": os.path.basename(xlsx), "converted_on": now}
        for i, k in enumerate(keys, 1): cfg[f"code_key_{i}"] = str(k)
        con.executemany("INSERT OR REPLACE INTO config VALUES (?,?)", list(cfg.items()))
        con.executemany("INSERT OR REPLACE INTO meta VALUES (?,?)", [("schema_version", FULL_SCHEMA_VERSION), ("allele_encoding", "allele x 10 as integer (9.3 -> 93), a1 <= a2"),
                                                                    ("created", now), ("source_workbook", os.path.basename(xlsx))])
        r = 5; i = 0
        while st[f"A{r}"].value:
            con.execute("INSERT INTO ref_category VALUES (?,?,?,?)", (i, st[f"A{r}"].value, st[f"B{r}"].value, st[f"C{r}"].value)); r += 1; i += 1
        for name, col in (("Case Type", "E"), ("Sample Type", "F"), ("Profile Type", "G"), ("No. of Contributors", "H"), ("Record Status", "I"),
                          ("Instrument", "L"), ("Analysis Software", "M"), ("State", "N"), ("District", "W")):
            r = 5; i = 0
            while st[f"{col}{r}"].value not in (None, ""):
                con.execute("INSERT OR IGNORE INTO ref_list VALUES (?,?,?)", (name, str(st[f"{col}{r}"].value), i)); r += 1; i += 1
        r = 5; i = 0
        while st[f"J{r}"].value:
            con.execute("INSERT INTO ref_kit VALUES (?,?,?)", (i, st[f"J{r}"].value, st[f"K{r}"].value)); r += 1; i += 1
        for kit, rules in R.kits:
            for mk, rule in rules.items(): con.execute("INSERT OR REPLACE INTO ref_kit_rule VALUES (?,?,?)", (kit, mk, rule))
        for i, (kw, s_, d_) in enumerate(R.locations): con.execute("INSERT INTO ref_location VALUES (?,?,?,?)", (i, kw, s_, d_))
        for i, c in enumerate(R.case):
            con.execute("INSERT INTO ref_field VALUES (?,?,?,?)", (i, c["header"], None if c["default"] in (None, "") else str(c["default"]), c["addr"]))
            for j, al in enumerate(c["aliases"]): con.execute("INSERT OR IGNORE INTO ref_alias VALUES (?,?,?)", (j, c["header"], al))
        for i, L in enumerate(R.loci):
            for j, al in enumerate(L["names"]): con.execute("INSERT OR IGNORE INTO ref_locus_alias VALUES (?,?,?)", (i * 10 + j, L["name"], al))
        lr = wb["Loci_Reference"]; r = 5; lid = 0
        while lr[f"B{r}"].value:
            mn, mx = lr[f"G{r}"].value, lr[f"H{r}"].value
            if isinstance(mn, (int, float)) and isinstance(mx, (int, float)):
                lid += 1
                con.execute("INSERT INTO locus VALUES (?,?,?,?,?,?,?,?)", (lid, lr[f"B{r}"].value, 1 if lr[f"E{r}"].value == "Y" else 0,
                            str(lr[f"C{r}"].value), lr[f"D{r}"].value, int(round(mn * 10)), int(round(mx * 10)),
                            "CODIS 20 core" if lr[f"E{r}"].value == "Y" else "Supplementary"))
            r += 1
    heads = [h for h in wdb.hmap if h not in ("Index Group", "Sex (from Amelogenin)", "Core Loci Typed (of 20)", "Min Loci Required", "Profile Completeness",
                                               "Entry Check", "Duplicate Check", "Upload Readiness", "Profile Key (helper)", "Identifiable Code", "De-identified Code",
                                               "Code helper n", "Code helper p2", "Code helper p3")]
    problems, n_ok = [], 0
    con.execute("BEGIN")
    for r in range(3, wdb.last_row + 1):
        if wdb.ws.cell(row=r, column=1).value in (None, ""): continue
        v = {h: wdb.ws.cell(row=r, column=wdb.hmap[h]).value for h in heads}
        try:
            insert_profile(con, R, v, now, entered_by=None); n_ok += 1
        except Exception as e:
            problems.append((v.get("Profile ID"), str(e)))
    con.execute("INSERT OR REPLACE INTO config VALUES ('audit_enabled','1')")
    con.execute("INSERT INTO audit_log(ts,user,action,table_name,record_ref,note) VALUES (?,?,?,?,?,?)",
                (now, current_user(), "CONVERT", "database", None, f"Converted {n_ok} profile(s) from {os.path.basename(xlsx)}"))
    con.commit(); con.execute("ANALYZE"); con.close()
    # ---- verification: read everything back and compare with the workbook, field by field
    store = SqlStore(out_db)
    checked = 0; mismatches = []
    allele_loci = [L["name"] for L in R.loci]
    plain = [h for h in heads if h in store.hmap and not (h.endswith("(1)") or h.endswith("(2)"))]
    for r in range(3, wdb.last_row + 1):
        pid = wdb.ws.cell(row=r, column=1).value
        if pid in (None, ""): continue
        d = store.profile_dict_by_ref(pid)
        if d is None: mismatches.append((pid, "missing in SQLite")); continue
        for h in plain:
            checked += 1
            a_, b_ = _nv(wdb.ws.cell(row=r, column=wdb.hmap[h]).value), _nv(d.get(h))
            if a_ != b_: mismatches.append((pid, f"{h}: workbook={a_!r} sqlite={b_!r}"))
        for n in allele_loci:
            if f"{n} (1)" not in wdb.hmap: continue
            checked += 1
            w1, w2 = _nv(wdb.ws.cell(row=r, column=wdb.hmap[f"{n} (1)"]).value), _nv(wdb.ws.cell(row=r, column=wdb.hmap[f"{n} (2)"]).value)
            s1, s2 = _nv(d.get(f"{n} (1)")), _nv(d.get(f"{n} (2)"))
            wp = tuple(sorted(map(str, (w1, w2)))) if (w1 is not None or w2 is not None) else None
            sp = tuple(sorted(map(str, (s1, s2)))) if (s1 is not None or s2 is not None) else None
            if wp != sp and n != "Amelogenin": mismatches.append((pid, f"{n}: workbook={wp} sqlite={sp}"))
            elif wp != sp and n == "Amelogenin": mismatches.append((pid, f"{n}: workbook={wp} sqlite={sp}"))
    stats = {"profiles": n_ok, "fields_checked": checked, "mismatches": mismatches, "problems": problems, "lab_code": lab, "path": out_db}
    store.con.close()
    return stats

def profile_query_text(db, profile_id):
    if getattr(db, "is_sqlite", False):
        d = db.profile_dict_by_ref(profile_id)
        if d is None: raise ValueError(f"Profile {profile_id} not found")
        return "\n".join(f"{L['name']} {d[L['name'] + ' (1)']},{d[L['name'] + ' (2)']}" for L in db.rules.loci
                         if L["name"] != "Amelogenin" and isinstance(d.get(f"{L['name']} (1)"), (int, float)))
    for r in range(3, db.last_row + 1):
        if str(db.ws.cell(row=r, column=1).value).upper() == str(profile_id).strip().upper():
            lines = []
            for L in db.rules.loci:
                if L["name"] == "Amelogenin": continue
                x = db.ws.cell(row=r, column=db.hmap.get(f"{L['name']} (1)", 1)).value
                y = db.ws.cell(row=r, column=db.hmap.get(f"{L['name']} (2)", 1)).value
                if isinstance(x, (int, float)) and isinstance(y, (int, float)): lines.append(f"{L['name']} {fmt_num(x)},{fmt_num(y)}")
            return "\n".join(lines)
    raise ValueError(f"Profile {profile_id} not found in the workbook")

def default_export_dir():
    base = os.path.join(os.path.expanduser("~"), "Documents", "DNA_Database", "Exports")
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:
        base = os.path.join(os.path.expanduser("~"), "DNA_Exports"); os.makedirs(base, exist_ok=True)
    return base

def open_export_dialog(tk, ttk, filedialog, messagebox, root, d, mode):
    """One export window for every export type: choose profiles from the list, choose the format and the folder."""
    if not d.rules.code_keys:
        messagebox.showinfo("Export", "Add the code columns first (the app offers this when the database is loaded)."); return
    rows_all = list_profiles(d)
    if not rows_all:
        messagebox.showinfo("Export", "The database has no profiles."); return
    top = tk.Toplevel(root); top.title("Export profiles"); top.geometry("940x620")
    ok_status = ("Approved for upload", "Uploaded to DB")
    shown = []
    kind = tk.StringVar(value=mode)
    per_file = tk.BooleanVar(value=(mode in ("xml", "json", "both")))
    only_appr = tk.BooleanVar(value=False)
    open_after = tk.BooleanVar(value=True)
    folder = tk.StringVar(value=default_export_dir())

    ttk.Label(top, text="1. Choose the profile(s) to export - type to search (Profile ID, code, FIR/case no., police station, district, category, status...)").pack(anchor="w", padx=10, pady=(8, 2))
    search = tk.StringVar(value="")
    sfr = ttk.Frame(top); sfr.pack(fill="x", padx=10)
    ttk.Label(sfr, text="Search:").pack(side="left")
    sentry = ttk.Entry(sfr, textvariable=search, width=60); sentry.pack(side="left", padx=6)
    ttk.Button(sfr, text="Clear search", command=lambda: search.set("")).pack(side="left")
    count = tk.StringVar(value="")
    ttk.Label(sfr, textvariable=count).pack(side="left", padx=14)
    lb = tk.Listbox(top, selectmode="extended", exportselection=False, font=("Consolas", 10), height=14)
    lb.pack(fill="both", expand=True, padx=10, pady=(4, 0))

    def refill(*_):
        words = (search.get() or "").lower().split()
        shown.clear(); lb.delete(0, "end")
        for r in rows_all:
            if only_appr.get() and r[5] not in ok_status: continue
            if words and not all(w in r[6] for w in words): continue
            shown.append(r)
            lb.insert("end", f"{r[0]:<12} {r[1]:<13} {str(r[2]):<16} {str(r[3]):<18} {str(r[4]):<22} {r[5]}")
        lb.select_set(0, "end")                      # everything that matches the search is selected; click to narrow down
        count.set(f"{len(shown)} of {len(rows_all)} profile(s) shown and selected")
    search.trace_add("write", refill)

    bar = ttk.Frame(top); bar.pack(fill="x", padx=10, pady=4)
    ttk.Button(bar, text="Select all shown", command=lambda: lb.select_set(0, "end")).pack(side="left")
    ttk.Button(bar, text="Clear selection", command=lambda: lb.selection_clear(0, "end")).pack(side="left", padx=6)
    ttk.Checkbutton(bar, text="list approved / uploaded profiles only", variable=only_appr, command=refill).pack(side="left", padx=12)

    ttk.Label(top, text="2. Format").pack(anchor="w", padx=10, pady=(6, 0))
    fr = ttk.Frame(top); fr.pack(fill="x", padx=10)
    for v, t in (("deid", "De-identified Excel + linkage register"), ("xml", "XML"), ("json", "JSON"), ("both", "XML + JSON"), ("sqlite", "SQLite database (fast search)")):
        ttk.Radiobutton(fr, text=t, variable=kind, value=v).pack(side="left", padx=6)
    ttk.Checkbutton(top, text="XML / JSON: one file per profile (named by De-identified Code)", variable=per_file).pack(anchor="w", padx=14)

    ttk.Label(top, text="3. Save in folder (created automatically if it does not exist)").pack(anchor="w", padx=10, pady=(8, 0))
    fr2 = ttk.Frame(top); fr2.pack(fill="x", padx=10)
    ttk.Entry(fr2, textvariable=folder, width=95).pack(side="left", fill="x", expand=True)
    def browse():
        p = filedialog.askdirectory(title="Folder to save the export in", initialdir=folder.get() or default_export_dir(), mustexist=False)
        if p: folder.set(os.path.normpath(p))
    ttk.Button(fr2, text="Browse...", command=browse).pack(side="left", padx=6)
    ttk.Checkbutton(top, text="open the folder after exporting", variable=open_after).pack(anchor="w", padx=14, pady=4)

    def go():
        ids = [shown[i][0] for i in lb.curselection()]
        if not ids:
            messagebox.showinfo("Export", "Select at least one profile.", parent=top); return
        out = os.path.abspath(os.path.expanduser((folder.get() or "").strip().strip('"')))
        try:
            os.makedirs(out, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Export", f"The folder could not be created:\n{out}\n\n{e}\n\nChoose another folder.", parent=top); return
        try:
            if kind.get() == "sqlite":
                p1, p2, added, upd, total = export_sqlite(d, out, ids, False, True); files = [p1, p2]; n = added + upd
                note = f"SQLite: {added} added, {upd} updated, {total} profile(s) in the database now. Re-exporting updates the same files.\nKeep the CONFIDENTIAL linkage database separate from the profile database."
            elif kind.get() == "deid":
                p1, p2, n = export_deidentified(d, out, False, ids); files = [p1, p2]; note = "The linkage register is CONFIDENTIAL - keep it separate."
            else:
                fm = ("json", "xml") if kind.get() == "both" else (kind.get(),)
                files, sums, n = export_structured(d, out, False, ids, per_file.get(), fm); files = files + [sums]; note = "Generic exchange schema (not an official CMF)."
            missing = [f for f in files if not os.path.isfile(f)]
            if missing: raise RuntimeError("These files were not written:\n" + "\n".join(missing))
        except Exception as e:
            messagebox.showerror("Export", f"Export failed:\n{e}", parent=top); return
        messagebox.showinfo("Export", f"{n} profile(s) exported to:\n{out}\n\n" + "\n".join(os.path.basename(f) for f in files) + "\n\n" + note, parent=top)
        if open_after.get() and hasattr(os, "startfile"):
            try: os.startfile(out)
            except OSError: pass
        top.destroy()
    ttk.Button(top, text="Export", command=go).pack(side="right", padx=10, pady=8)
    refill()
    try: sentry.focus_set()
    except Exception: pass
    return top


def open_sqlite_search(tk, ttk, filedialog, messagebox, root, d=None):
    """Match a query STR profile against an exported SQLite database."""
    import time
    top = tk.Toplevel(root); top.title("SQLite match search"); top.geometry("1000x680")
    default_db = d.path if getattr(d, "is_sqlite", False) else os.path.join(default_export_dir(), SQL_DB_NAME)
    dbvar = tk.StringVar(value=default_db if os.path.isfile(default_db) else "")
    mode = tk.StringVar(value="High (identical genotypes)")
    minl, mism, lim = tk.StringVar(value="8"), tk.StringVar(value="0"), tk.StringVar(value="50")
    pidv = tk.StringVar(value="")
    look = tk.BooleanVar(value=False)
    status = tk.StringVar(value="")
    f1 = ttk.Frame(top); f1.pack(fill="x", padx=10, pady=6)
    ttk.Label(f1, text="SQLite database:").pack(side="left")
    ttk.Entry(f1, textvariable=dbvar, width=80).pack(side="left", padx=6)
    ttk.Button(f1, text="Browse...", command=lambda: dbvar.set(os.path.normpath(filedialog.askopenfilename(filetypes=[("SQLite", "*.db *.sqlite")]) or dbvar.get()))).pack(side="left")
    ttk.Label(top, text="Query profile - one locus per line, e.g.  D3S1358 15,16   (paste from Excel works too)").pack(anchor="w", padx=10)
    txt = tk.Text(top, height=11, width=60, font=("Consolas", 10)); txt.pack(fill="x", padx=10)
    f2 = ttk.Frame(top); f2.pack(fill="x", padx=10, pady=4)
    ttk.Label(f2, text="or take it from workbook profile:").pack(side="left")
    ttk.Entry(f2, textvariable=pidv, width=14).pack(side="left", padx=6)
    def fill():
        if d is None: messagebox.showinfo("Search", "Load the workbook database in the main window first.", parent=top); return
        try:
            t = profile_query_text(d, pidv.get())
            txt.delete("1.0", "end"); txt.insert("1.0", t)
        except Exception as e:
            messagebox.showerror("Search", str(e), parent=top)
    ttk.Button(f2, text="Fill", command=fill).pack(side="left")
    f3 = ttk.Frame(top); f3.pack(fill="x", padx=10, pady=4)
    ttk.Label(f3, text="Stringency").pack(side="left")
    ttk.Combobox(f3, textvariable=mode, values=["High (identical genotypes)", "Moderate (>=1 shared allele)"], width=28, state="readonly").pack(side="left", padx=6)
    for lab, var, w in (("Min loci compared", minl, 4), ("Mismatches allowed", mism, 4), ("Max results", lim, 6)):
        ttk.Label(f3, text=lab).pack(side="left", padx=(12, 2)); ttk.Entry(f3, textvariable=var, width=w).pack(side="left")
    ttk.Checkbutton(top, text="also look up identifiers (needs dna_linkage_CONFIDENTIAL.db next to the database) - restricted", variable=look).pack(anchor="w", padx=10)
    cols = ("spec", "cat", "kit", "comp", "ident", "shared", "cls", "ident_code", "fir", "ps")
    tv = ttk.Treeview(top, columns=cols, show="headings", height=12)
    for c, w, t in zip(cols, (110, 140, 130, 70, 70, 70, 120, 190, 100, 120),
                       ("De-identified Code", "Category", "Kit", "Compared", "Identical", "Shared", "Match class", "Identifiable Code", "FIR/Case", "Police Station")):
        tv.heading(c, text=t); tv.column(c, width=w, anchor="w")
    tv.pack(fill="both", expand=True, padx=10, pady=4)
    ttk.Label(top, textvariable=status).pack(fill="x", padx=10)
    def run():
        try:
            path = (dbvar.get() or "").strip().strip('"')
            if not os.path.isfile(path): raise ValueError("Choose the SQLite database file (dna_profiles.db).")
            q = parse_query(txt.get("1.0", "end"))
            t0 = time.time()
            rows, info = sqlite_search(path, q, "high" if mode.get().startswith("High") else "moderate", int(minl.get()), int(mism.get()), int(lim.get()))
            el = time.time() - t0
            link = {}
            if look.get():
                lp = os.path.join(os.path.dirname(path), SQL_LINK_NAME)
                _c = sqlite3.connect(path); _full = _c.execute("SELECT 1 FROM sqlite_master WHERE name='v_profile_full'").fetchone(); _c.close()
                if _full: link = linkage_lookup(path, [r[0] for r in rows])
                elif os.path.isfile(lp): link = linkage_lookup(lp, [r[0] for r in rows])
                else: messagebox.showinfo("Search", "No linkage database next to the profile database - showing de-identified results only.", parent=top)
            tv.delete(*tv.get_children())
            for sid, cat, kit, comp, ident, shared in rows:
                cls = "High" if ident == comp else ("Moderate" if shared == comp else "Partial")
                lk = link.get(sid, ("", "", "", "", ""))
                tv.insert("", "end", values=(sid, cat, kit, comp, ident, shared, cls, lk[0], lk[3], lk[4]))
            status.set(f"{len(rows)} result(s) in {el*1000:.0f} ms - {info['path']}" + (f", {info['candidates']} candidate(s) checked" if "candidates" in info else "") +
                       ". Hits are investigative leads requiring verification.")
        except Exception as e:
            messagebox.showerror("Search", str(e), parent=top)
    ttk.Button(top, text="Search", command=run).pack(side="right", padx=10, pady=8)
    return top


def open_detail(tk, ttk, messagebox, root, d, ref):
    try:
        info = profile_detail(d, ref)
    except Exception as e:
        messagebox.showerror("Profile", str(e)); return
    top = tk.Toplevel(root); top.title(f"Profile {ref}"); top.geometry("820x700")
    tv = ttk.Treeview(top, columns=("sec", "field", "val"), show="headings")
    for c, w, t in (("sec", 110, "Section"), ("field", 230, "Field"), ("val", 440, "Value")):
        tv.heading(c, text=t); tv.column(c, width=w, anchor="w")
    for sec, f, v in info: tv.insert("", "end", values=(sec, f, v))
    tv.pack(fill="both", expand=True, padx=8, pady=8)
    return top

def open_browser(tk, ttk, filedialog, messagebox, root, d):
    """All profiles of the loaded database (SQLite or Excel): type to search, page through, double-click to open one."""
    PAGE = 500
    top = tk.Toplevel(root); top.title("Browse profiles"); top.geometry("1250x680")
    search = tk.StringVar(value=""); status = tk.StringVar(value=""); page = [0]
    ttk.Label(top, text="Search - any words: Profile ID, code, FIR / case no., police station, district, state, category, kit, person ID, sections of law, analyst, status...").pack(anchor="w", padx=10, pady=(8, 2))
    fr = ttk.Frame(top); fr.pack(fill="x", padx=10)
    ttk.Entry(fr, textvariable=search, width=70).pack(side="left")
    ttk.Button(fr, text="Clear", command=lambda: search.set("")).pack(side="left", padx=6)
    cols = ("ref", "code", "fir", "ps", "state", "cat", "kit", "loci", "status", "adate")
    tv = ttk.Treeview(top, columns=cols, show="headings", height=22)
    for c, w, t in zip(cols, (95, 105, 100, 120, 130, 130, 130, 50, 140, 90),
                       ("Profile ID", "De-identified Code", "FIR / Case", "Police Station", "State", "Category", "Kit", "Loci", "Status", "Analysis date")):
        tv.heading(c, text=t); tv.column(c, width=w, anchor="w")
    tv.pack(fill="both", expand=True, padx=10, pady=6)
    ttk.Label(top, textvariable=status).pack(anchor="w", padx=10)
    def refresh(*_):
        total, rows = browse_rows(d, search.get(), page[0] * PAGE, PAGE)
        tv.delete(*tv.get_children())
        for r in rows: tv.insert("", "end", values=tuple("" if x is None else x for x in r))
        lo = page[0] * PAGE + 1 if total else 0
        status.set(f"{lo:,} - {page[0] * PAGE + len(rows):,} of {total:,} profile(s). Double-click a row (or select it and click Open) to see everything about it.")
        refresh.total = total
    def newsearch(*_):
        page[0] = 0; refresh()
    def nxt():
        if (page[0] + 1) * PAGE < getattr(refresh, "total", 0): page[0] += 1; refresh()
    def prv():
        if page[0] > 0: page[0] -= 1; refresh()
    def opn(*_):
        sel = tv.selection()
        if sel: open_detail(tk, ttk, messagebox, root, d, tv.item(sel[0])["values"][0])
    search.trace_add("write", newsearch)
    bf = ttk.Frame(top); bf.pack(fill="x", padx=10, pady=6)
    ttk.Button(bf, text="< Previous", command=prv).pack(side="left"); ttk.Button(bf, text="Next >", command=nxt).pack(side="left", padx=6)
    ttk.Button(bf, text="Open profile", command=opn).pack(side="left", padx=12)
    tv.bind("<Double-1>", opn)
    refresh()
    return top

def open_ystr_search(tk, ttk, filedialog, messagebox, root, d=None):
    top = tk.Toplevel(root); top.title("Y-STR haplotype search"); top.geometry("900x620")
    default_db = d.path if getattr(d, "is_sqlite", False) else os.path.join(default_export_dir(), SQL_DB_NAME)
    dbvar = tk.StringVar(value=default_db if os.path.isfile(default_db) else "")
    mism, minl, lim = tk.StringVar(value="0"), tk.StringVar(value="1"), tk.StringVar(value="100")
    status = tk.StringVar(value=""); look = tk.BooleanVar(value=False)
    f1 = ttk.Frame(top); f1.pack(fill="x", padx=10, pady=6)
    ttk.Label(f1, text="SQLite database:").pack(side="left")
    ttk.Entry(f1, textvariable=dbvar, width=70).pack(side="left", padx=6)
    ttk.Button(f1, text="Browse...", command=lambda: dbvar.set(os.path.normpath(filedialog.askopenfilename(filetypes=[("SQLite", "*.db *.sqlite")]) or dbvar.get()))).pack(side="left")
    ttk.Label(top, text="Query haplotype - one locus per line, e.g.  DYS391 10   DYS385 11,14   (multi-copy loci: both values)").pack(anchor="w", padx=10)
    txt = tk.Text(top, height=10, width=50, font=("Consolas", 10)); txt.pack(fill="x", padx=10)
    f3 = ttk.Frame(top); f3.pack(fill="x", padx=10, pady=4)
    for lab, var, w in (("Mismatching loci allowed", mism, 4), ("Min loci compared", minl, 4), ("Max results", lim, 6)):
        ttk.Label(f3, text=lab).pack(side="left", padx=(0, 2)); ttk.Entry(f3, textvariable=var, width=w).pack(side="left", padx=(0, 12))
    ttk.Checkbutton(top, text="also look up identifiers (restricted)", variable=look).pack(anchor="w", padx=10)
    cols = ("code", "kit", "comp", "ident", "ident_code", "fir", "ps")
    tv = ttk.Treeview(top, columns=cols, show="headings", height=12)
    for c, w, t in zip(cols, (120, 150, 80, 80, 200, 110, 130), ("De-identified Code", "Y-STR kit", "Compared", "Identical", "Identifiable Code", "FIR/Case", "Police Station")):
        tv.heading(c, text=t); tv.column(c, width=w, anchor="w")
    tv.pack(fill="both", expand=True, padx=10, pady=4)
    ttk.Label(top, textvariable=status).pack(fill="x", padx=10)
    def run():
        try:
            path = (dbvar.get() or "").strip().strip('"')
            if not os.path.isfile(path): raise ValueError("Choose the SQLite database file.")
            q = parse_haplotype(txt.get("1.0", "end"))
            rows, info = ystr_search(path, q, int(mism.get()), int(minl.get()), int(lim.get()))
            link = linkage_lookup(path, [r[0] for r in rows]) if look.get() else {}
            tv.delete(*tv.get_children())
            for sid, kit, comp, ident in rows:
                lk = link.get(sid, ("", "", "", "", ""))
                tv.insert("", "end", values=(sid, kit or "", comp, ident, lk[0], lk[3], lk[4]))
            status.set(f"{info['total']} profile(s) matched; {info['exact']} identical at all {info['query_loci']} queried loci (haplotype count). "
                       "Y-STR matches are lineage evidence, not individual identification.")
        except Exception as e:
            messagebox.showerror("Y-STR search", str(e), parent=top)
    ttk.Button(top, text="Search", command=run).pack(side="right", padx=10, pady=8)
    return top

def open_minloci_dialog(tk, ttk, messagebox, root, d):
    if not getattr(d, "is_sqlite", False):
        messagebox.showinfo("Minimum loci", "For the Excel workbook, edit the 'Min core loci (policy)' column on the Settings sheet."); return
    top = tk.Toplevel(root); top.title("Minimum core loci per category"); top.geometry("520x520")
    ttk.Label(top, text="Profiles with fewer typed core loci than the minimum are marked 'Review' on import.").pack(anchor="w", padx=10, pady=8)
    rows = d.con.execute("SELECT category, index_group, min_core_loci FROM ref_category ORDER BY ord").fetchall()
    vars_ = {}
    fr = ttk.Frame(top); fr.pack(fill="both", expand=True, padx=10)
    for i, (cat, grp, mn) in enumerate(rows):
        ttk.Label(fr, text=f"{cat}   ({grp})").grid(row=i, column=0, sticky="w", pady=2)
        v = tk.StringVar(value=str(mn)); vars_[cat] = v
        ttk.Entry(fr, textvariable=v, width=6).grid(row=i, column=1, padx=10)
    def save():
        try:
            for cat, v in vars_.items():
                n = int(v.get())
                if not 0 <= n <= 20: raise ValueError(f"{cat}: enter a number between 0 and 20")
                d.con.execute("UPDATE ref_category SET min_core_loci=? WHERE category=?", (n, cat))
            d.con.commit(); d.rules.cat_min = {c: int(v.get()) for c, v in vars_.items()}
            messagebox.showinfo("Minimum loci", "Saved.", parent=top); top.destroy()
        except Exception as e:
            messagebox.showerror("Minimum loci", str(e), parent=top)
    ttk.Button(top, text="Save", command=save).pack(side="right", padx=10, pady=8)
    return top

# ----------------------------------------------------------------------------- CLI
def run_cli(args):
    if args.build_freq or args.import_freq or args.locus_stats or args.rmp or args.kinship or args.ystr_freq or args.ystr_diversity:
        path = args.db
        if args.import_freq:
            n = import_allele_frequencies(path, args.population, args.import_freq, args.n_individuals or 100)
            print(f"Imported {n} allele frequencies into population '{args.population}'"); return
        if args.build_freq:
            n, k = build_allele_frequencies(path, args.population or "internal", args.all_profiles)
            print(f"Population '{args.population or 'internal'}': {n} individual(s), {k} allele frequencies stored"); return
        if args.locus_stats:
            print(f"{'Locus':<10}{'N':>6}{'Alleles':>8}{'Ho':>8}{'He':>8}{'PIC':>8}{'PD':>9}{'HWE p':>9}   (HWE: permutation test if numpy is installed, else chi-square)")
            for r in locus_statistics(path, args.all_profiles):
                print(f"{r['locus']:<10}{r['N']:>6}{r['alleles']:>8}{r['Ho']:>8.3f}{r['He']:>8.3f}{r['PIC']:>8.3f}{r['PD']:>9.4f}" + (f"{r['HWE_p']:>9.3f}" if r['HWE_p'] is not None else f"{'n/a':>9}"))
            return
        if args.ystr_diversity:
            print(ystr_diversity(path)); return
        if args.ystr_freq:
            r = ystr_frequency(path, parse_haplotype(args.haplotype))
            print(f"Haplotype seen {r['x']} time(s) in {r['N']} profile(s) typed at all queried loci")
            if r["N"]: print(f"  x/N = {r['freq']:.5f}   (x+1)/(N+1) = {r['estimate']:.5f}   95% upper limit = {r['upper95']:.5f}  (1 in {1/r['upper95']:,.0f})")
            return
        ft = FreqTable(path, args.population or "internal")
        ids = [i.strip() for x in (args.profile or []) for i in x.split(",") if i.strip()]
        if args.rmp:
            g = profile_genotype(path, ids[0])
            p, per = rmp(ft, g, args.theta)
            print(f"Profile {ids[0]}: {len(g)} loci, theta = {args.theta if args.theta is not None else ft.theta}, population '{args.population or 'internal'}' (N = {ft.n})")
            for l, v in per.items(): print(f"  {l:<10} {g[l][0]/10:g},{g[l][1]/10:g}   P = {v:.5g}")
            print(f"Combined random match probability = {p:.3e}   (1 in {1/p:,.0f});  LR (same source vs unrelated) = {1/p:.3e}")
            return
        if args.kinship:
            g1, g2 = profile_genotype(path, ids[0]), profile_genotype(path, args.profile2)
            lr, per = kinship_lr(ft, g1, g2, args.relation or "parent-child")
            print(f"{ids[0]} vs {args.profile2} - hypothesis '{args.relation or 'parent-child'}' vs unrelated, {len(per)} loci")
            for l, v in per.items(): print(f"  {l:<10} LR = {v:.4g}")
            print(f"Combined LR = {lr:.4e}   (log10 = {math.log10(lr) if lr > 0 else float('-inf'):.2f})"); return
    if args.ystr_search:
        q = parse_haplotype(args.haplotype)
        rows, info = ystr_search(args.ystr_search, q, args.mismatches, args.min_loci or 1, args.limit)
        print(f"{info['total']} profile(s) matched; {info['exact']} identical at all {info['query_loci']} queried loci (haplotype count)")
        for sid, kit, comp, ident in rows: print(f"  {sid}  {str(kit):<22} compared {comp:>2}  identical {ident:>2}")
        return
    if args.export_ystr:
        p, n = export_ystr(args.db, args.export_ystr, [i.strip() for x in (args.profile or []) for i in x.split(",") if i.strip()] or None)
        print(f"Exported {n} Y-STR haplotype(s): {p}"); return
    if args.list or args.show:
        db = open_database(args.db)
        if args.show:
            for sec, f, v in profile_detail(db, args.show): print(f"{sec:<12} {f:<28} {v}")
            return
        total, rows = browse_rows(db, args.find or "", 0, args.limit)
        print(f"{total} profile(s)" + (f" matching '{args.find}'" if args.find else "") + f" - showing {len(rows)}")
        for r in rows: print("  " + "  ".join(str(x if x is not None else "") for x in r[:9]))
        return
    if args.set_min_loci:
        db = open_database(args.db)
        for item in args.set_min_loci:
            cat, n = item.rsplit("=", 1)
            cur = db.con.execute("UPDATE ref_category SET min_core_loci=? WHERE category=?", (int(n), cat.strip()))
            print(f"{cat.strip()}: minimum core loci = {int(n)}" if cur.rowcount else f"Category '{cat.strip()}' not found")
        db.con.commit(); return
    if args.new_sqlite:
        new_empty_database(args.new_sqlite, args.template, args.lab_code if args.lab_code != "CFSLH" else None)
        print(f"Created {args.new_sqlite} (reference tables copied from {args.template}; new secret code keys generated)"); return
    if args.convert_to_sqlite:
        st = convert_workbook_to_sqlite(args.convert_to_sqlite, args.out, args.lab_code)
        print(f"Converted {st['profiles']} profile(s) -> {st['path']}")
        print(f"Verified {st['fields_checked']} field(s) against the workbook: {len(st['mismatches'])} mismatch(es)")
        for m in st["mismatches"][:20]: print("   MISMATCH", m)
        for p in st["problems"]: print("   NOT CONVERTED", p)
        print(f"Laboratory code: {st['lab_code']}  (codes and code keys are stored in the config table - keep the database access-controlled)")
        return
    if args.sqlite_search:
        q_text = args.query
        if args.query_profile:
            q_text = profile_query_text(open_database(args.db), args.query_profile)
        import time
        t0 = time.time()
        rows, info = sqlite_search(args.sqlite_search, parse_query(q_text), "high" if args.mode == "high" else "moderate",
                                   args.min_loci or 8, args.mismatches, args.limit)
        stats_ = match_statistics(args.sqlite_search, [r[0] for r in rows], parse_query(q_text), args.population) if args.population else {}
        link = linkage_lookup(args.linkage, [r[0] for r in rows]) if args.linkage and os.path.isfile(args.linkage) else {}
        print(f"{len(rows)} result(s) in {(time.time() - t0) * 1000:.0f} ms [{info['path']}]")
        for sid, cat, kit, comp, ident, shared in rows:
            extra = ("  " + " | ".join(str(x) for x in link[sid])) if sid in link else ""
            if sid in stats_: extra += f"   RMP {stats_[sid][0]:.2e} over {stats_[sid][2]} identical loci (LR {stats_[sid][1]:.2e})"
            print(f"  {sid}  {kit:<20} compared {comp:>2}  identical {ident:>2}  shared {shared:>2}{extra}")
        return
    if args.export_sqlite:
        ids = [i.strip() for x in (args.profile or []) for i in x.split(",") if i.strip()] or None
        p1, p2, added, upd, total = export_sqlite(open_database(args.db), args.export_sqlite, ids, args.only_approved, not args.no_linkage)
        print(f"SQLite: {added} added, {upd} updated, {total} profile(s) in total\n  {p1}" + (f"\n  {p2}  (CONFIDENTIAL)" if p2 else "")); return
    if args.repair:
        repair_database(args.db); print("Repaired the Calc helper sheet in", args.db); return
    if args.add_codes:
        print("Code columns added." if add_code_columns(args.db, args.lab_code) else "Code columns already present."); return
    if args.export_xml:
        ids = [i.strip() for x in (args.profile or []) for i in x.split(",") if i.strip()] or None
        fmts = ("json", "xml") if args.format == "both" else (args.format,)
        if ids and not args.format_given: fmts = ("xml",)
        files, sums, n = export_structured(open_database(args.db), args.export_xml, args.only_approved, ids, bool(ids) or args.per_profile, fmts)
        print(f"Exported {n} profile(s):"); [print("  " + f) for f in files]; print("  " + sums); return
    if args.export:
        ids = [i.strip() for x in (args.profile or []) for i in x.split(",") if i.strip()] or None
        p1, p2, n = export_deidentified(open_database(args.db), args.export, args.only_approved, ids)
        print(f"Exported {n} profile(s):\n  {p1}\n  {p2}  (CONFIDENTIAL linkage register)"); return
    db = open_database(args.db)
    if hasattr(db, "wb") and calc_needs_repair(db.wb): print("WARNING: the hidden Calc sheet is damaged (a row was deleted?). Run with --repair to fix searches/dashboard.")
    ov = dict(s.split("=", 1) for s in (args.set or []))
    recs = []
    for f in args.files: recs += read_source(f, db.rules, args.homozygote)
    res = evaluate(db, recs, ov)
    for x in res:
        v = x["values"]
        print(f"[{x['status']:6}] {x['rec'].source} row {x['rec'].row} ({x['rec'].layout}) FIR={v.get('Case/FIR No.')} PS={v.get('Police Station')} "
              f"Kit={v.get('Kit/Method')} State={v.get('State')} District={v.get('District')} loci={x['typed']}/20")
        if x["auto"]: print("          auto-filled:", "; ".join(x["auto"]))
        for i in x["issues"]: print("          -", i)
        for w in x["warnings"]: print("          ! check:", w)
    if args.add:
        n, bak = commit(db, res, args.review_too, True, not args.no_incomplete)
        print(f"Added {n} row(s) to {args.db}. " + (f"Safety copy (outside the database folder): {bak}" if bak else "(SQLite: transaction committed and audit-logged)"))
        for x in res:
            if x.get("codes") and x["codes"][1] and "assigned_id" in x: print(f"  {x['assigned_id']}: {x['codes'][0]}  /  {x['codes'][1]}")

# ----------------------------------------------------------------------------- GUI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    cfg_path = os.path.join(os.path.expanduser("~"), ".dna_import_app.json")
    import json
    cfg = {}
    try: cfg = json.load(open(cfg_path))
    except Exception: pass

    root = tk.Tk(); root.title("DNA Database Import"); root.geometry("1280x720")
    state = {"db": None, "files": [], "records": [], "results": []}
    dbvar = tk.StringVar(value=cfg.get("db", ""))
    status = tk.StringVar(value="Choose the database workbook, then add Excel files.")
    review_too = tk.BooleanVar(value=False)
    incomplete = tk.BooleanVar(value=True)
    homo = tk.BooleanVar(value=False)

    top = ttk.Frame(root, padding=8); top.pack(fill="x")
    ttk.Label(top, text="Database (Excel workbook or SQLite .db):").grid(row=0, column=0, sticky="w")
    ttk.Entry(top, textvariable=dbvar, width=90).grid(row=0, column=1, padx=6)
    ttk.Button(top, text="Browse...", command=lambda: pick_db()).grid(row=0, column=2)
    ttk.Button(top, text="Load / reload", command=lambda: load_db()).grid(row=0, column=3, padx=6)

    dfr = ttk.LabelFrame(root, text="Defaults for fields missing in the imported files (optional)", padding=8); dfr.pack(fill="x", padx=8)
    fields = ["Kit/Method", "Profile Category", "Sample Type", "Case Type", "Analyst", "Analysis Date", "Instrument", "Analysis Software", "Record Status"]
    dvars = {}
    for i, f in enumerate(fields):
        ttk.Label(dfr, text=f + (" (STR kit used)" if f == "Kit/Method" else "")).grid(row=i // 5 * 2, column=i % 5, sticky="w", padx=6)
        v = tk.StringVar(); dvars[f] = v
        cb = ttk.Combobox(dfr, textvariable=v, width=24); cb.grid(row=i // 5 * 2 + 1, column=i % 5, padx=6, pady=(0, 4), sticky="w")
        dvars[f + "_cb"] = cb
    ttk.Checkbutton(dfr, text="Single allele in one cell = homozygote", variable=homo).grid(row=3, column=0, columnspan=2, padx=10, sticky="w")

    bar = ttk.Frame(root, padding=8); bar.pack(fill="x")
    ttk.Button(bar, text="Add Excel files...", command=lambda: add_files()).pack(side="left")
    ttk.Button(bar, text="Clear list", command=lambda: clear_files()).pack(side="left", padx=6)
    ttk.Button(bar, text="Re-check", command=lambda: recheck()).pack(side="left")
    ttk.Checkbutton(bar, text="add incomplete rows (missing details, flagged in the database)", variable=incomplete).pack(side="left", padx=10)
    ttk.Checkbutton(bar, text="also add rows needing review", variable=review_too).pack(side="left", padx=6)
    ttk.Button(bar, text="Add to database", command=lambda: do_commit()).pack(side="right")
    ttk.Button(bar, text="Convert Excel to SQLite...", command=lambda: do_convert()).pack(side="right", padx=4)
    ttk.Button(bar, text="SQLite match search...", command=lambda: open_sqlite_search(tk, ttk, filedialog, messagebox, root, state["db"])).pack(side="right", padx=4)
    ttk.Button(bar, text="Export XML / JSON...", command=lambda: open_export_dialog(tk, ttk, filedialog, messagebox, root, state["db"], "xml") if state["db"] else messagebox.showinfo("Export", "Load the database first.")).pack(side="right", padx=4)
    ttk.Button(bar, text="Export de-identified...", command=lambda: open_export_dialog(tk, ttk, filedialog, messagebox, root, state["db"], "deid") if state["db"] else messagebox.showinfo("Export", "Load the database first.")).pack(side="right", padx=4)

    bar2 = ttk.Frame(root, padding=(8, 0)); bar2.pack(fill="x")
    ttk.Button(bar2, text="Browse all profiles...", command=lambda: open_browser(tk, ttk, filedialog, messagebox, root, state["db"]) if state["db"] else messagebox.showinfo("Browse", "Load the database first.")).pack(side="left")
    ttk.Button(bar2, text="Y-STR search...", command=lambda: open_ystr_search(tk, ttk, filedialog, messagebox, root, state["db"])).pack(side="left", padx=6)
    def do_export_ystr():
        d = state["db"]
        if not getattr(d, "is_sqlite", False): messagebox.showinfo("Y-STR", "Y-STR export needs the SQLite database (.db)."); return
        out = filedialog.askdirectory(title="Folder for the Y-STR haplotype export", initialdir=default_export_dir())
        if not out: return
        try:
            p, n = export_ystr(d.path, out); messagebox.showinfo("Y-STR", f"{n} haplotype(s) exported:\n{p}")
        except Exception as e:
            messagebox.showerror("Y-STR", str(e))
    ttk.Button(bar2, text="Export Y-STR haplotypes...", command=do_export_ystr).pack(side="left")
    ttk.Button(bar2, text="Minimum loci per category...", command=lambda: open_minloci_dialog(tk, ttk, messagebox, root, state["db"]) if state["db"] else None).pack(side="left", padx=6)

    cols = ("status", "file", "fir", "ps", "state", "district", "kit", "loci", "issues")
    tv = ttk.Treeview(root, columns=cols, show="headings", height=16)
    for c, w, t in zip(cols, (70, 160, 110, 110, 90, 100, 120, 55, 560),
                       ("Status", "File", "FIR / Case No.", "Police Station", "State", "District", "Kit", "Loci", "Issues / auto-filled")):
        tv.heading(c, text=t); tv.column(c, width=w, anchor="w")
    tv.tag_configure("Ready", background="#C6E0B4"); tv.tag_configure("Incomplete", background="#DDEBF7"); tv.tag_configure("Review", background="#FFE699"); tv.tag_configure("Duplicate", background="#F4B183")
    tv.pack(fill="both", expand=True, padx=8, pady=4)
    ttk.Label(root, textvariable=status, padding=8).pack(fill="x")

    def busy(msg):
        status.set(msg); root.update_idletasks()

    def pick_db():
        p = filedialog.askopenfilename(filetypes=[("Database", "*.xlsx *.db *.sqlite"), ("Excel", "*.xlsx"), ("SQLite", "*.db *.sqlite")])
        if p: dbvar.set(p); load_db()

    def load_db():
        p = dbvar.get().strip()
        if not os.path.isfile(p): messagebox.showerror("Database", "Database file not found."); return
        busy("Loading database (can take up to a minute)...")
        def work():
            try:
                d = open_database(p)
                root.after(0, lambda: loaded(d, p))
            except Exception as e:
                root.after(0, lambda: messagebox.showerror("Database", f"Could not open the database:\n{e}"))
        threading.Thread(target=work, daemon=True).start()

    def loaded(d, p):
        state["db"] = d
        cfg["db"] = p
        try: json.dump(cfg, open(cfg_path, "w"))
        except Exception: pass
        for f in fields:
            lst = d.rules.lists.get(f)
            if lst: dvars[f + "_cb"]["values"] = lst
        busy(f"Database loaded: {len(d.ids)} profiles, {d.free_rows()} free rows. Add Excel files to import.")
        if getattr(d, "is_sqlite", False):
            busy(f"SQLite database loaded: {len(d.ids):,} profiles. Add Excel files to import - they are written straight into the database.")
            recheck(rescan=True); return
        if "Identifiable Code" not in d.hmap:
            if messagebox.askyesno("Identification codes",
                    "This database has no Identifiable Code / De-identified Code columns yet.\n\nAdd them now? "
                    "(Close the workbook in Excel first. Your profiles are not changed.)"):
                from tkinter import simpledialog
                lab = simpledialog.askstring("Laboratory code", "Laboratory code used in the Identifiable Code (e.g. CFSLH):", initialvalue="CFSLH") or "CFSLH"
                try:
                    add_code_columns(p, re.sub(r"[^A-Za-z0-9]", "", lab).upper() or "LAB")
                    messagebox.showinfo("Identification codes", "Code columns added. Open the workbook in Excel - codes are generated automatically.")
                    root.after(100, load_db); return
                except Exception as e:
                    messagebox.showerror("Identification codes", f"Could not add the columns:\n{e}")
        if calc_needs_repair(d.wb):
            if messagebox.askyesno("Database repair needed",
                    "The database's hidden helper sheet is damaged (usually because a row was deleted in the DNA Database sheet). "
                    "Search, Profile_Match and the Dashboard will show errors.\n\nRepair it now? (Close the workbook in Excel first. "
                    "Your profiles are not changed.)"):
                try:
                    repair_database(p); messagebox.showinfo("Repair", "Repaired. Open the workbook in Excel - formulas recalculate on opening.")
                except Exception as e:
                    messagebox.showerror("Repair", f"Repair failed:\n{e}")
        recheck(rescan=True)

    def add_files():
        if not state["db"]: messagebox.showinfo("Import", "Load the database first."); return
        ps = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if ps: state["files"] += [p for p in ps if p not in state["files"]]; recheck(rescan=True)

    def clear_files():
        state["files"] = []; state["records"] = []; state["results"] = []; tv.delete(*tv.get_children())
        busy("List cleared.")

    def overrides():
        return {f: dvars[f].get().strip() for f in fields if dvars[f].get().strip()}

    def recheck(rescan=False):
        d = state["db"]
        if not d: return
        try:
            if rescan:
                recs = []
                for f in state["files"]:
                    busy(f"Reading {os.path.basename(f)} ...")
                    recs += read_source(f, d.rules, homo.get())
                state["records"] = recs
            state["results"] = evaluate(d, state["records"], overrides())
        except Exception as e:
            messagebox.showerror("Import", f"Could not read a file:\n{e}"); return
        tv.delete(*tv.get_children())
        for x in state["results"]:
            v = x["values"]
            note = "; ".join(x["issues"] + (["Auto-filled: " + ", ".join(x["auto"])] if x["auto"] else [])
                             + ["CHECK: " + w for w in x["warnings"]])
            tv.insert("", "end", tags=(x["status"],), values=(x["status"], x["rec"].source, v.get("Case/FIR No.", ""), v.get("Police Station", ""),
                      v.get("State", ""), v.get("District", ""), v.get("Kit/Method", ""), f"{x['typed']}/20", note))
        cnt = {k: sum(1 for x in state["results"] if x["status"] == k) for k in ("Ready", "Incomplete", "Review", "Duplicate")}
        busy(f"{len(state['results'])} profile(s): {cnt['Ready']} ready, {cnt['Incomplete']} incomplete (missing details), {cnt['Review']} need review, {cnt['Duplicate']} DUPLICATE (never added). "
             "Set the defaults above to fill missing details, then Add to database.")

    def do_commit():
        d = state["db"]
        if not d or not state["results"]: return
        rows = selected(state["results"], incomplete.get(), review_too.get())
        if not rows:
            lines = []
            for x in state["results"][:8]:
                lines.append(f"- {x['rec'].source}: {x['status']} - " + "; ".join(x["issues"]))
            messagebox.showinfo("Nothing to add",
                "No rows can be added with the current settings.\n\n" + "\n".join(lines) +
                "\n\nFill the Defaults above (Profile Category, Sample Type, Analyst, Analysis Date), then click Re-check - "
                "or tick 'add incomplete rows' / 'also add rows needing review'.")
            return
        n = len(rows)
        if not messagebox.askyesno("Add to database", f"Add {n} row(s) to the database?\nClose the workbook in Excel first. The rows are added to the existing file (no new file is created)."): return
        busy("Writing to the database...")
        def work():
            try:
                cnt, bak = commit(d, state["results"], review_too.get(), True, incomplete.get())
                root.after(0, lambda: done(cnt, bak))
            except PermissionError:
                root.after(0, lambda: messagebox.showerror("Import", "The database is open in Excel. Close it and try again."))
            except Exception as e:
                root.after(0, lambda: messagebox.showerror("Import", f"Import failed:\n{e}"))
        threading.Thread(target=work, daemon=True).start()

    def do_convert():
        from tkinter import simpledialog
        src = filedialog.askopenfilename(title="Excel database to convert", filetypes=[("Excel", "*.xlsx")])
        if not src: return
        dst = filedialog.asksaveasfilename(title="New SQLite database file", defaultextension=".db", initialdir=default_export_dir(),
                                           initialfile="dna_database.db", filetypes=[("SQLite", "*.db")])
        if not dst: return
        lab = simpledialog.askstring("Laboratory code", "Laboratory code (only used if the workbook has none):", initialvalue="CFSLH") or "CFSLH"
        busy("Converting...")
        def work():
            try:
                st = convert_workbook_to_sqlite(src, dst, re.sub(r"[^A-Za-z0-9]", "", lab).upper() or "LAB")
                root.after(0, lambda: messagebox.showinfo("Convert", f"{st['profiles']} profile(s) converted.\n{st['fields_checked']} fields verified, {len(st['mismatches'])} mismatch(es).\n\n{dst}\n\nBrowse to this file to keep building the database."))
                root.after(0, lambda: (dbvar.set(dst), load_db()))
            except Exception as e:
                root.after(0, lambda: messagebox.showerror("Convert", str(e)))
        threading.Thread(target=work, daemon=True).start()

    def done(cnt, bak):
        lines = [f"{x['assigned_id']}:  {x['codes'][0]}  /  {x['codes'][1]}" for x in state["results"] if x.get("codes") and x["codes"][1] and "assigned_id" in x]
        messagebox.showinfo("Import", f"Added {cnt} profile(s) to the existing database file.\n" + (f"Safety copy kept in the app folder: {bak}\n" if bak else "(SQLite: transactional write, logged in the audit tables)\n")
                            + ("\nCodes (Identifiable / De-identified):\n" + "\n".join(lines) + "\n" if lines else "")
                            + "\nOpen the workbook in Excel - formulas recalculate on opening.")
        state["files"] = []; state["records"] = []; state["results"] = []; tv.delete(*tv.get_children())
        load_db()

    for f in fields[:0]: pass
    if dbvar.get() and os.path.isfile(dbvar.get()): root.after(200, load_db)
    root.mainloop()

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--db"); ap.add_argument("--files", nargs="*"); ap.add_argument("--add", action="store_true"); ap.add_argument("--repair", action="store_true")
    ap.add_argument("--add-codes", action="store_true"); ap.add_argument("--lab-code", default="CFSLH")
    ap.add_argument("--export"); ap.add_argument("--new-sqlite"); ap.add_argument("--template"); ap.add_argument("--population"); ap.add_argument("--theta", type=float, default=None); ap.add_argument("--all-profiles", action="store_true"); ap.add_argument("--build-freq", action="store_true"); ap.add_argument("--import-freq"); ap.add_argument("--n-individuals", type=int); ap.add_argument("--locus-stats", action="store_true"); ap.add_argument("--rmp", action="store_true"); ap.add_argument("--kinship", action="store_true"); ap.add_argument("--profile2"); ap.add_argument("--relation", choices=list(KINSHIP)); ap.add_argument("--ystr-freq", action="store_true"); ap.add_argument("--ystr-diversity", action="store_true"); ap.add_argument("--ystr-search"); ap.add_argument("--haplotype"); ap.add_argument("--export-ystr"); ap.add_argument("--list", action="store_true"); ap.add_argument("--find"); ap.add_argument("--show"); ap.add_argument("--set-min-loci", action="append"); ap.add_argument("--convert-to-sqlite"); ap.add_argument("--out"); ap.add_argument("--user"); ap.add_argument("--export-sqlite"); ap.add_argument("--no-linkage", action="store_true"); ap.add_argument("--sqlite-search"); ap.add_argument("--query"); ap.add_argument("--query-profile"); ap.add_argument("--mode", choices=["high", "moderate"], default="high"); ap.add_argument("--min-loci", type=int, default=None); ap.add_argument("--mismatches", type=int, default=0); ap.add_argument("--limit", type=int, default=50); ap.add_argument("--linkage"); ap.add_argument("--export-xml"); ap.add_argument("--profile", action="append"); ap.add_argument("--per-profile", action="store_true"); ap.add_argument("--format", choices=["xml", "json", "both"], default=None); ap.add_argument("--only-approved", action="store_true")
    ap.add_argument("--review-too", action="store_true"); ap.add_argument("--set", action="append")
    ap.add_argument("--homozygote", action="store_true"); ap.add_argument("--no-incomplete", action="store_true")
    a = ap.parse_args()
    a.format_given = a.format is not None
    a.format = a.format or "both"
    if a.user: os.environ["DNA_APP_USER"] = a.user
    if a.new_sqlite or a.convert_to_sqlite or a.sqlite_search or a.ystr_search or (a.db and (a.build_freq or a.import_freq or a.locus_stats or a.rmp or a.kinship or a.ystr_freq or a.ystr_diversity or a.files or a.repair or a.add_codes or a.export or a.export_xml or a.export_sqlite or a.export_ystr or a.list or a.show or a.set_min_loci)): run_cli(a)
    else: run_gui()
