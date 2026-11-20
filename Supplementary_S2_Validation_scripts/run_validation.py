#!/usr/bin/env python3
"""Validation harness for the laboratory DNA database prototype. Usage: python run_validation.py --data synth --out results [--perf 1000,10000,100000,300000]
Environment: DNA_APP_DIR = folder containing dna_import_app.py, TEMPLATE_DB = database to copy reference tables from."""
import argparse, csv, glob, json, os, random, sqlite3, sys, time, statistics, shutil, itertools, datetime
sys.path.insert(0, os.environ.get("DNA_APP_DIR", "."))
import dna_import_app as A

ap = argparse.ArgumentParser(); ap.add_argument("--data", default="synth"); ap.add_argument("--out", default="results")
ap.add_argument("--perf", default="1000,10000,100000"); ap.add_argument("--template", default=os.environ.get("TEMPLATE_DB", "Forensic_DNA_Database.db")); a = ap.parse_args()
os.makedirs(a.out, exist_ok=True); R = {}
def rd(name): return list(csv.DictReader(open(os.path.join(a.data, name))))
def fmtv(x): return str(int(x)) if float(x).is_integer() else str(x)
truth = {r["profile_ref"]: r for r in rd("truth_profiles.csv")}
indiv = {r["person"]: r for r in rd("truth_individuals.csv")}
rng = random.Random(7)
OV = {"Profile Category": "Convicted", "Sample Type": "Blood", "Analyst": "Analyst A", "Analysis Date": "05-11-2026"}

def fresh(name):
    p = os.path.join(a.out, name)
    for s in ("", "-wal", "-shm"):
        if os.path.exists(p + s): os.remove(p + s)
    A.new_empty_database(p, a.template, "LABX"); return p
def run_import(store, files, ov):
    recs = []
    for f in files: recs += A.read_source(f, store.rules)
    res = A.evaluate(store, recs, ov); return res
# ------------------------------------------------------------------ 1. import + QC
db = fresh("validation.db"); store = A.SqlStore(db)
t0 = time.time(); hfiles = sorted(glob.glob(os.path.join(a.data, "import_horizontal_*.xlsx")))
res_h = run_import(store, hfiles, {k: v for k, v in OV.items() if k != "Profile Category"})
res_v = run_import(store, [os.path.join(a.data, "import_vertical.xlsx")], {"Profile Category": "Convicted", "Sample Type": "Blood", "Analyst": "Analyst A", "Analysis Date": "05-11-2026"})
def ref_of(x): return str(x["values"].get("Lab Sample ID") or x["values"].get("Sample Name") or x["values"].get("Profile ID") or "")
from collections import Counter
R["import_status_horizontal"] = Counter(x["status"] for x in res_h); R["import_status_vertical"] = Counter(x["status"] for x in res_v)
R["import_n_records"] = {"horizontal": len(res_h), "vertical": len(res_v)}
# identity of each parsed record: use sample name stored in Remarks? fall back on FIR+report order -> match by genotype
def geno_of(v):
    g = {}
    for l in A.CORE if hasattr(A, "CORE") else []: pass
    return g
loci_names = [L["name"] for L in store.rules.loci if L["name"] != "Amelogenin"]
def vgeno(v):
    g = {}
    for n in loci_names:
        x, y = v.get(f"{n} (1)"), v.get(f"{n} (2)")
        if isinstance(x, (int, float)) and isinstance(y, (int, float)): g[n] = tuple(sorted((float(x), float(y))))
    return g
def tgeno(ref):
    t = truth[ref]; ind = indiv[t["person"]]; typed = t["typed_loci"].split("|"); g = {}
    for l in typed:
        if l in loci_names: g[l] = tuple(sorted((float(ind[f"{l}_1"]), float(ind[f"{l}_2"]))))
    return g
# accuracy of extraction: records are identified by the Lab Sample ID carried through the import
kitsize = {}
for t_ in truth.values(): kitsize[t_['kit']] = max(kitsize.get(t_['kit'], 0), len(t_['typed_loci'].split('|')))
kit_full = [0, 0]; kit_part = [0, 0]
fields = Counter(); ok = Counter(); allele_total = allele_ok = 0; matched = 0; unmatched = 0; status_vs_truth = Counter()
all_res = res_h + res_v
for x in all_res:
    v = x["values"]; ref = str(v.get("Lab Sample ID") or "")
    if ref not in truth: unmatched += 1; continue
    t = truth[ref]; matched += 1
    for lab, got, exp in (("police_station", v.get("Police Station"), t["police_station"]), ("district", v.get("District"), t["district"]), ("state", v.get("State"), t["state"]), ("kit", v.get("Kit/Method"), t["kit"]), ("lab_case_no", v.get("Lab Case No."), t["report"]), ("fir", v.get("Case/FIR No."), t["fir"])):
        fields[lab] += 1; ok[lab] += (str(got) == exp)
    full_ = len(t['typed_loci'].split('|')) == kitsize[t['kit']]; z = kit_full if full_ else kit_part; z[1] += 1; z[0] += (str(v.get('Kit/Method')) == t['kit'])
    tg, vg = tgeno(ref), vgeno(v)
    for l, g in tg.items():
        allele_total += 1; allele_ok += (vg.get(l) == g)
    complete = int(t["n_loci_typed"]) >= int(t["min_required"])
    status_vs_truth[(("complete" if complete else "partial"), x["status"])] += 1
R["extraction"] = {"records_matched_to_truth": matched, "unmatched": unmatched, "allele_genotypes_compared": allele_total, "allele_genotypes_correct": allele_ok,
                   "kit_full_profiles": kit_full, "kit_partial_profiles": kit_part, "field_accuracy": {k: [ok[k], fields[k]] for k in fields}, "status_vs_truth": {f"{a_}|{b_}": n_ for (a_, b_), n_ in status_vs_truth.items()}}
n_added = 0
store2 = store
n, _ = A.commit_sqlite(store, all_res, include_review=False, include_incomplete=True, user="validator", source="validation") if True else (0, 0)
R["import_added"] = n; R["import_seconds"] = round(time.time() - t0, 2)
R["import_status_total"] = dict(Counter(x["status"] for x in all_res))
# ------------------------------------------------------------------ 2. duplicates
store = A.SqlStore(db)
dres = run_import(store, [os.path.join(a.data, "duplicates_horizontal.xlsx")], {"Analyst": "Analyst A", "Analysis Date": "05-11-2026", "Sample Type": "Blood"})
dj = json.load(open(os.path.join(a.data, "truth_duplicates.json")))
R["duplicates"] = {"records": len(dres), "status": dict(Counter(x["status"] for x in dres)),
                   "planted_exact": len(dj["exact_duplicates"]), "planted_same_person_other_case": len(dj["same_person_other_case"])}
# per-class status
cls = {}
for x in dres:
    v = x["values"]; r_ = str(v.get("Lab Sample ID") or "")
    t = truth.get(r_)
    cls.setdefault(t["planted"] if t else "?", Counter())[x["status"]] += 1
R["duplicates"]["by_class"] = {k: dict(v) for k, v in cls.items()}
R["duplicates"]["flagged_identical_profile"] = {k: sum(1 for x in dres if truth.get(str(x["values"].get("Lab Sample ID")), {}).get("planted") == k and any("dentical profile" in i or "DUPLICATE" in i for i in x["issues"])) for k in ("exact_duplicate", "same_person_other_case")}
R["duplicates"]["originals_not_in_database"] = sum(1 for k_ in dj["exact_duplicates"] if True) - R["duplicates"]["flagged_identical_profile"]["exact_duplicate"]
# ------------------------------------------------------------------ 3. error handling
eres = run_import(store, [os.path.join(a.data, "errors_horizontal.xlsx")], {"Analysis Date": "05-11-2026", "Sample Type": "Blood", "Profile Category": "Convicted"})
terr = {r["sample"]: r for r in rd("truth_errors.csv")}
conf = {}
for x in eres:
    v = x["values"]; r_ = str(v.get("Lab Sample ID") or "")
    t = terr.get(r_)
    if t: conf.setdefault(t["error_type"], Counter())[x["status"]] += 1
R["errors"] = {"records": len(eres), "by_type": {k: dict(v) for k, v in conf.items()}, "expected": {t["error_type"]: t["expected_status"] for t in terr.values()}}
# ------------------------------------------------------------------ 4. pseudonymisation
con = sqlite3.connect(db); cfg = dict(con.execute("SELECT key,value FROM config"))
keys = [int(cfg[f"code_key_{i}"]) for i in range(1, 7)]
did = [r[0] for r in con.execute("SELECT specimen_id FROM profile")]; iden = [r[0] for r in con.execute("SELECT identifiable_code FROM profile")]
nprof = len(did); ps = {}
ps["profiles"] = nprof; ps["unique_deid"] = len(set(did)); ps["unique_iden"] = len(set(iden))
# permutation injectivity over 1..200000 and check character detection
codes = {}
for n_ in range(1, 200001):
    codes[A.make_codes(f"DNA-{n_:06d}", "Reference / Elimination", datetime.date(2026, 1, 1), None, keys, "LABX")[1]] = n_
ps["injective_serials_tested"] = 200000; ps["injective_unique_codes"] = len(codes)
ALPH = A.CODE_ALPHABET
def chk_ok(c):
    s = c[3:].replace("-", ""); v = [ALPH.index(ch) for ch in s]; return (1*v[0]+3*v[1]+5*v[2]+7*v[3]+9*v[4]+11*v[5]) % 32 == v[6]
sample = list(codes)[:20000]; det_sub = tot_sub = det_tr = tot_tr = 0
for c in sample:
    body = c[3:].replace("-", "")
    for pos in range(7):
        for ch in ALPH:
            if ch == body[pos]: continue
            b2 = body[:pos] + ch + body[pos+1:]; tot_sub += 1; det_sub += (not chk_ok("DI-" + b2[:4] + "-" + b2[4:]))
    for pos in range(6):
        if body[pos] == body[pos+1]: continue
        b2 = body[:pos] + body[pos+1] + body[pos] + body[pos+2:]; tot_tr += 1; det_tr += (not chk_ok("DI-" + b2[:4] + "-" + b2[4:]))
        if len(sample) and pos == 5 and tot_tr > 200000: break
    if tot_sub > 3_000_000: break
ps["single_substitution_detected"] = [det_sub, tot_sub]; ps["adjacent_transposition_detected"] = [det_tr, tot_tr]
# serial recoverability: does the code reveal the serial order? correlation of sequential serials with code
import statistics as st
firsts = [codes_k for codes_k in list(codes)[:2000]]
sn = list(range(1, 2001)); c_num = [sum(ALPH.index(ch) * 32 ** (6 - i) for i, ch in enumerate(A.make_codes(f"DNA-{n_:06d}", "x", datetime.date(2026,1,1), None, keys, "L")[1][3:].replace("-", "")[:4])) for n_ in sn]
mx = st.mean(sn); my = st.mean(c_num); r_ = sum((x - mx) * (y - my) for x, y in zip(sn, c_num)) / (len(sn) * st.pstdev(sn) * st.pstdev(c_num))
ps["serial_vs_code_pearson_r"] = round(r_, 4)
ps["same_serial_different_key_equal"] = sum(1 for n_ in range(1, 2001) if A.make_codes(f"DNA-{n_:06d}", "x", datetime.date(2026,1,1), None, [x + 1 for x in keys], "L")[1] == A.make_codes(f"DNA-{n_:06d}", "x", datetime.date(2026,1,1), None, keys, "L")[1])
v_cols = [r[1] for r in con.execute("PRAGMA table_info(v_deidentified)")]; ps["v_deidentified_columns"] = v_cols
R["pseudonymisation"] = ps
con.close()
# ------------------------------------------------------------------ 5. matching (SQLite)
store = A.SqlStore(db); con = sqlite3.connect(db)
ref2sid = {}
for ref, sid in con.execute("SELECT profile_ref, specimen_id FROM profile"): ref2sid[ref] = sid
pid_by_sid = {s: i for i, s in con.execute("SELECT profile_id, specimen_id FROM profile")}
rows_all = con.execute("SELECT p.profile_id, p.specimen_id, p.person_source_id, p.loci_typed, p.kit, p.specimen_category FROM profile p").fetchall()
# person of each stored profile via genotype key -> truth person (genotype unique). Build map from stored genotype to truth person
def stored_geno(pid):
    return {n: (a / 10, b / 10) for n, a, b in con.execute("SELECT l.name,a.a1,a.a2 FROM allele a JOIN locus l USING(locus_id) WHERE a.profile_id=?", (pid,))}
person_of = {}
gkey2person = {}
for pid, sid, *_ in rows_all:
    g = stored_geno(pid); full = tuple(sorted(g.items()))
    gkey2person[full] = None
truth_people = {}
for pr, t in truth.items():
    truth_people.setdefault(t["person"], []).append(pr)
# stored profile -> person: match on the complete genotype of typed loci against individuals
ind_geno = {p: {l: tuple(sorted((float(r[f"{l}_1"]), float(r[f"{l}_2"])))) for l in loci_names} for p, r in indiv.items()}
core_names = [n for n, in con.execute("SELECT name FROM locus WHERE is_core=1 ORDER BY locus_id")]
idx = {}
for p, g in ind_geno.items(): idx.setdefault(tuple(g[l] for l in core_names[:6]), []).append(p)
def person_for(pid):
    g = stored_geno(pid); key = tuple(g.get(l) for l in core_names[:6])
    if None in key: return None
    c = idx.get(key, [])
    c = [p for p in c if all(ind_geno[p][l] == v for l, v in g.items())]
    return c[0] if len(c) == 1 else (c[0] if c else None)
pid_person = {pid: person_for(pid) for pid, *_ in rows_all}
people_to_pids = {}
for pid, p in pid_person.items():
    if p is not None: people_to_pids.setdefault(p, []).append(pid)
sid_of = {pid: sid for pid, sid, *_ in rows_all}
def q_of(pid): return {n: v for n, v in stored_geno(pid).items()}
def qtext(g): return {n: v for n, v in g.items()}
def do_search(g, **kw):
    t = time.time(); rows, info = A.sqlite_search(db, {n: (int(round(v[0]*10)), int(round(v[1]*10))) for n, v in g.items()}, **kw); return rows, info, (time.time() - t) * 1000
# exact self-search + search for same-person dup profiles
ms = {"exact": {"queries": 0, "self_found": 0, "all_same_person_found": 0, "false_positive_rows": 0, "ms": []},
      "pigeon": {}}
test = rng.sample([pid for pid, *_ in rows_all if (A.SqlStore and True)], min(300, len(rows_all)))
for pid in test:
    g = q_of(pid)
    if len(g) < 8: continue
    try: rows, info, ms_ = do_search(g, mode="high", min_loci=8, mismatches=0)
    except ValueError: continue
    hit_sids = {r[0] for r in rows}; same = {sid_of[p_] for p_ in people_to_pids.get(pid_person[pid], []) if p_ != pid} if pid_person[pid] is not None else set()
    ms["exact"]["queries"] += 1; ms["exact"]["self_found"] += (sid_of[pid] in hit_sids)
    ms["exact"]["all_same_person_found"] += all(s in hit_sids for s in same)
    ms["exact"]["false_positive_rows"] += sum(1 for r in rows if r[0] != sid_of[pid] and pid_person[pid_by_sid[r[0]]] != pid_person[pid])
    ms["exact"]["ms"].append(ms_)
e = ms["exact"]; e["median_ms"] = statistics.median(e["ms"]); e["p95_ms"] = sorted(e["ms"])[int(.95 * len(e["ms"])) - 1]; del e["ms"]
R["match_exact"] = e
# mismatches: introduce k errors (allele changes) at k loci; search with k mismatches allowed
mm = {}
for k in (0, 1, 2, 3):
    found = tot = 0; times = []
    for pid in rng.sample([p for p, *_ in rows_all], 150):
        g = q_of(pid)
        if len(g) < 12: continue
        ls = rng.sample(list(g), k)
        for l in ls: g[l] = (g[l][0] + 1.0, g[l][1])
        try: rows, info, ms_ = do_search(g, mode="high", min_loci=8, mismatches=k)
        except ValueError: continue
        tot += 1; found += (sid_of[pid] in {r[0] for r in rows}); times.append(ms_)
    mm[k] = {"queries": tot, "recovered": found, "median_ms": statistics.median(times) if times else None}
R["match_mismatch"] = mm
# partial profiles (drop-out) min loci 8
pp = {}
for nl in (8, 10, 12, 15):
    found = tot = 0; nres = []
    for pid in rng.sample([p for p, *_ in rows_all], 150):
        g = q_of(pid)
        if len(g) < 20: continue
        keep = rng.sample(list(g), nl); gq = {l: g[l] for l in keep}
        rows, info, _ = do_search(gq, mode="high", min_loci=8, mismatches=0)
        tot += 1; found += (sid_of[pid] in {r[0] for r in rows}); nres.append(len(rows))
    pp[nl] = {"queries": tot, "recovered": found, "mean_results": statistics.mean(nres) if nres else None}
R["match_partial"] = pp
# expected vs observed adventitious matches (using true population frequencies) for partial queries
freq = {}
for r in rd("population_frequencies.csv"): freq.setdefault(r["locus"], {})[float(r["allele"])] = float(r["frequency"])
def gp(l, g):
    a, b = g; pa, pb = freq[l].get(a, 1e-4), freq[l].get(b, 1e-4); return pa * pa if a == b else 2 * pa * pb
adv = []
for nl in (8, 10, 12):
    exp_tot = obs_tot = 0; nq = 0
    people = list(ind_geno)
    N = len(set(p for p in pid_person.values() if p))
    for pid in rng.sample([p for p, *_ in rows_all], 150):
        g = q_of(pid)
        if len(g) < 20: continue
        keep = rng.sample(list(g), nl); gq = {l: g[l] for l in keep}
        rows, info, _ = do_search(gq, mode="high", min_loci=8, mismatches=0)
        pr = 1.0
        for l in keep: pr *= gp(l, g[l])
        ownp = pid_person[pid]; own = people_to_pids.get(ownp, [pid])
        # eligible others = stored profiles of different persons typed at all queried loci
        elig = 0
        for p2, *_x in rows_all:
            pass
        obs = sum(1 for r in rows if pid_person[pid_by_sid[r[0]]] != ownp and pid_person[pid_by_sid[r[0]]] is not None)
        eligible = con.execute("SELECT COUNT(DISTINCT p) FROM (SELECT profile_id p FROM allele WHERE locus_id IN (SELECT locus_id FROM locus WHERE name IN (%s)) GROUP BY profile_id HAVING COUNT(*)=%d)" % (",".join("?" * nl), nl), keep).fetchone()[0]
        exp_tot += pr * max(eligible - len(own), 0); obs_tot += obs; nq += 1
    adv.append({"loci": nl, "queries": nq, "expected_adventitious": round(exp_tot, 5), "observed": obs_tot})
R["match_adventitious"] = adv
# moderate stringency: query = child, find parent (shares >=1 allele at every locus)
rel = rd("truth_relations.csv"); mod = {"parent-child": [0, 0], "full-siblings": [0, 0]}; mod_other = []
for r in rel:
    if r["relation"] not in mod: continue
    pa, pb = people_to_pids.get(r["person_a"]), people_to_pids.get(r["person_b"])
    if not pa or not pb: continue
    g = q_of(pa[0])
    if len(g) < 20: continue
    rows, info, _ = do_search(g, mode="moderate", min_loci=8)
    mod[r["relation"]][1] += 1; mod[r["relation"]][0] += (sid_of[pb[0]] in {x[0] for x in rows}); mod_other.append(len(rows))
R["match_moderate"] = {"found_true_relative": mod, "mean_other_results": statistics.mean(mod_other) if mod_other else None}
# ------------------------------------------------------------------ 6. statistics
pop_all = "validation_all"; n_pop, k_pop = A.build_allele_frequencies(db, "internal", True)
ft = A.FreqTable(db, "internal"); st_rows = A.locus_statistics(db, True)
dev = []; 
for r in st_rows:
    l = r["locus"]
    if l not in freq: continue
    pt = freq[l]; he_true = 1 - sum(p * p for p in pt.values())
    dev.append((l, r["He"], he_true, r["Ho"], r["HWE_p"], r["N"]))
R["stats_locus"] = {"n_loci": len(dev), "max_abs_He_error": max(abs(x[1] - x[2]) for x in dev), "mean_abs_He_error": statistics.mean(abs(x[1] - x[2]) for x in dev),
                    "hwe_p_below_0.05": sum(1 for x in dev if x[4] is not None and x[4] < 0.05), "hwe_p_values": [x[4] for x in dev], "N": [x[5] for x in dev][0],
                    "table": [{"locus": r["locus"], "N": r["N"], "alleles": r["alleles"], "Ho": r["Ho"], "He": r["He"], "PIC": r["PIC"], "PD": r["PD"], "HWE_p": r["HWE_p"]} for r in st_rows]}
# RMP vs analytical (true frequencies, theta 0) on 20 random full profiles, log10 comparison (theta=0, same freq source)
logdiff = []; rmps = []
for pid in rng.sample([p for p, *_ in rows_all], 200):
    g = stored_geno(pid)
    if len(g) < 20: continue
    gx = {l: (int(round(v[0] * 10)), int(round(v[1] * 10))) for l, v in g.items()}
    p_app, _ = A.rmp(ft, gx, 0.0); p_true = 1.0
    for l, v in g.items(): p_true *= gp(l, v)
    logdiff.append(abs(__import__("math").log10(p_app) - __import__("math").log10(p_true))); rmps.append(p_app)
    p_th, _ = A.rmp(ft, gx, 0.01)
R["stats_rmp"] = {"profiles": len(logdiff), "median_abs_log10_diff_vs_true_freq": statistics.median(logdiff), "max_abs_log10_diff": max(logdiff),
                  "median_log10_rmp": statistics.median([__import__("math").log10(x) for x in rmps]), "min_rmp": min(rmps), "max_rmp": max(rmps)}
# theta effect
import math
g0 = None
for pid, *_ in rows_all:
    g0 = stored_geno(pid)
    if len(g0) >= 20: break
gx0 = {l: (int(round(v[0] * 10)), int(round(v[1] * 10))) for l, v in g0.items()}
R["stats_theta_effect"] = {str(t): math.log10(A.rmp(ft, gx0, t)[0]) for t in (0.0, 0.01, 0.03)}
# kinship LR on true pairs
kin = {"parent-child": [], "full-siblings": [], "unrelated": []}
for r in rel:
    pa, pb = people_to_pids.get(r["person_a"]), people_to_pids.get(r["person_b"])
    if not pa or not pb: continue
    g1, g2 = stored_geno(pa[0]), stored_geno(pb[0])
    if len(g1) < 20 or len(g2) < 20: continue
    g1 = {l: (int(round(v[0]*10)), int(round(v[1]*10))) for l, v in g1.items()}; g2 = {l: (int(round(v[0]*10)), int(round(v[1]*10))) for l, v in g2.items()}
    for hyp in ("parent-child", "full-siblings"):
        lr, _ = A.kinship_lr(ft, g1, g2, hyp); kin.setdefault(f"{r['relation']}|{hyp}", []).append(math.log10(max(lr, 1e-300)))
R["kinship"] = {k: {"n": len(v), "median_log10LR": statistics.median(v), "q05": sorted(v)[int(.05 * len(v))], "q95": sorted(v)[min(int(.95 * len(v)), len(v)-1)],
                    "frac_LR_gt_1": sum(1 for x in v if x > 0) / len(v), "values": v} for k, v in kin.items() if "|" in k and v}
# ------------------------------------------------------------------ 7. Y-STR
yt = rd("ystr_truth.csv"); cons = sqlite3.connect(db)
ys = {}
ys["ystr_profiles"] = cons.execute("SELECT COUNT(*) FROM ystr_profile").fetchone()[0]
ys["kit_distribution"] = dict(cons.execute("SELECT ystr_kit, COUNT(*) FROM ystr_profile GROUP BY 1").fetchall())
ys["diversity"] = str(A.ystr_diversity(db))
# exact haplotype searches: query = stored haplotype; compare hit count with truth count
rows_y = cons.execute("SELECT p.profile_id FROM ystr_profile y JOIN profile p USING(profile_id)").fetchall()
cnt_ok = tot = 0; ms_y = []; fr = []
for (pid,) in rng.sample(rows_y, min(len(rows_y), 200)):
    q = dict(cons.execute("SELECT marker, value FROM strmarker WHERE kind='Y' AND profile_id=?", (pid,)).fetchall())
    if not q: continue
    t = time.time(); rws, info = A.ystr_search(db, q, 0, 1); ms_y.append((time.time() - t) * 1000)
    # brute-force count
    pids = None
    for m, v in q.items():
        s = {r[0] for r in cons.execute("SELECT profile_id FROM strmarker WHERE kind='Y' AND marker=? AND value=?", (m, v))}
        pids = s if pids is None else pids & s
    tot += 1; cnt_ok += (info["exact"] == len(pids))
    f = A.ystr_frequency(db, q); fr.append(f)
ys["exact_count_agrees_with_bruteforce"] = [cnt_ok, tot]; ys["median_ms"] = statistics.median(ms_y) if ms_y else None
ys["median_N"] = statistics.median(f["N"] for f in fr) if fr else None; ys["median_x"] = statistics.median(f["x"] for f in fr) if fr else None
ys["median_upper95"] = statistics.median(f["upper95"] for f in fr) if fr else None
# coverage of the 95% upper bound vs true population frequency (simulate haplotype draws)
import numpy as np
w = np.array([float(r["true_frequency"]) for r in yt]); w = w / w.sum(); nrng = np.random.default_rng(3)
cover = []
for n_ in (50, 200, 1000):
    c = 0; B = 400
    for _ in range(B):
        hap = nrng.choice(len(w), p=w); sample_ = nrng.choice(len(w), size=n_, p=w); x = int((sample_ == hap).sum())
        c += (A.binom_upper(x, n_) >= w[hap])
    cover.append({"N": n_, "coverage": c / B})
ys["upper_limit_coverage"] = cover
# exact binomial upper limit check
ys["binom_upper_checks"] = {"x0_N100": A.binom_upper(0, 100), "x0_N100_expected": 1 - 0.05 ** (1 / 100), "x1_N20": A.binom_upper(1, 20)}
R["ystr"] = ys
# ------------------------------------------------------------------ 8. integrity & audit
ig = {}
c = sqlite3.connect(db); c.execute("PRAGMA foreign_keys=ON")
ig["integrity_check"] = c.execute("PRAGMA integrity_check").fetchone()[0]; ig["fk_violations"] = len(c.execute("PRAGMA foreign_key_check").fetchall())
ig["audit_rows"] = c.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0]; ig["import_log_rows"] = c.execute("SELECT COUNT(*) FROM import_log").fetchone()[0]
pid1 = c.execute("SELECT profile_id FROM profile LIMIT 1").fetchone()[0]
n0 = c.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0]
c.execute("UPDATE profile SET record_status='Approved' WHERE profile_id=?", (pid1,)); c.execute("UPDATE profile SET kit='Edited kit' WHERE profile_id=?", (pid1,))
c.execute("UPDATE allele SET a1=a1 WHERE profile_id=?", (pid1,))
a1 = c.execute("SELECT locus_id,a1,a2 FROM allele WHERE profile_id=? LIMIT 1", (pid1,)).fetchone()
c.execute("UPDATE allele SET a1=? WHERE profile_id=? AND locus_id=?", (a1[1] + 10 if a1[1] + 10 <= a1[2] else a1[1], pid1, a1[0])); c.commit()
ig["audit_rows_generated_by_3_edits"] = c.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0] - n0
try: c.execute("UPDATE allele SET a1=999, a2=1 WHERE profile_id=? AND locus_id=?", (pid1, a1[0])); c.commit(); ig["a1_gt_a2_rejected"] = False
except sqlite3.Error as e: c.rollback(); ig["a1_gt_a2_rejected"] = True
try: c.execute("INSERT INTO allele VALUES (999999,1,100,100)"); c.commit(); ig["orphan_allele_rejected"] = False
except sqlite3.Error: c.rollback(); ig["orphan_allele_rejected"] = True
# genotype key consistency: recompute from allele table
bad = 0; chk = 0
core_ids = [r for r in c.execute("SELECT locus_id,name FROM locus WHERE is_core=1 ORDER BY locus_id")]
for pid, key in c.execute("SELECT profile_id, genotype_key FROM profile WHERE genotype_key IS NOT NULL LIMIT 1500"):
    g = {n: (x, y) for n, x, y in c.execute("SELECT l.name,a.a1,a.a2 FROM allele a JOIN locus l USING(locus_id) WHERE a.profile_id=?", (pid,))}
    chk += 1; bad += (A.genotype_key([n for _, n in core_ids], {k: v for k, v in g.items() if k in [n for _, n in core_ids]}) != key)
ig["genotype_key_checked"] = chk; ig["genotype_key_inconsistent"] = bad
R["integrity"] = ig
R["db_profiles"] = c.execute("SELECT COUNT(*) FROM profile").fetchone()[0]; R["db_alleles"] = c.execute("SELECT COUNT(*) FROM allele").fetchone()[0]
c.close()
# ------------------------------------------------------------------ 9. performance (bulk synthetic loader)
def bulk(path, n, seed=1):
    r = random.Random(seed)
    p = path
    for s in ("", "-wal", "-shm"):
        if os.path.exists(p + s): os.remove(p + s)
    A.new_empty_database(p, a.template, "PERF"); con = sqlite3.connect(p); con.execute("PRAGMA journal_mode=OFF"); con.execute("PRAGMA synchronous=OFF")
    cores = con.execute("SELECT locus_id,name FROM locus WHERE is_core=1 ORDER BY locus_id").fetchall(); core_n = [n_ for _, n_ in cores]
    f2 = {n_: (list(freq[n_]), list(freq[n_].values())) for n_ in core_n if n_ in freq}
    cfgp = dict(con.execute("SELECT key,value FROM config")); ks = [int(cfgp[f"code_key_{i}"]) for i in range(1, 7)]
    t = time.time(); con.execute("BEGIN")
    drop = con.execute("DROP INDEX IF EXISTS ix_allele_genotype"); con.execute("DROP INDEX IF EXISTS ix_profile_key")
    for i in range(1, n + 1):
        g = {}
        for lid, nm in cores:
            al, w_ = f2[nm]; x, y = r.choices(al, w_, k=2); g[nm] = tuple(sorted((int(round(x * 10)), int(round(y * 10)))))
        key = A.genotype_key(core_n, g); di = A.make_codes(f"DNA-{i:06d}", "Reference / Elimination", datetime.date(2026,1,1), None, ks, "PERF")
        cur = con.execute("INSERT INTO profile(profile_ref,specimen_id,identifiable_code,specimen_category,index_group,kit,loci_typed,genotype_key,created_at) VALUES (?,?,?,?,?,?,20,?,?)",
                          (f"DNA-{i:06d}", di[1], di[0], "Convicted", "Offender / Suspect", "GlobalFiler", key, "2026"))
        con.executemany("INSERT INTO allele VALUES (?,?,?,?)", [(cur.lastrowid, lid, g[nm][0], g[nm][1]) for lid, nm in cores])
    con.commit(); con.execute("CREATE INDEX ix_allele_genotype ON allele(locus_id,a1,a2,profile_id)"); con.execute("CREATE INDEX ix_profile_key ON profile(genotype_key)")
    con.execute("ANALYZE"); con.commit(); con.close(); return time.time() - t
perf = {}
for n in [int(x) for x in a.perf.split(",") if x]:
    pth = os.path.join(a.out, f"perf_{n}.db"); load_s = bulk(pth, n); size = os.path.getsize(pth) / 1e6
    con = sqlite3.connect(pth); core_n = [x for _, x in con.execute("SELECT locus_id,name FROM locus WHERE is_core=1 ORDER BY locus_id")]
    pids = [r[0] for r in con.execute("SELECT profile_id FROM profile ORDER BY RANDOM() LIMIT 60")]
    def gq(pid): return {nm: (x, y) for nm, x, y in con.execute("SELECT l.name,a.a1,a.a2 FROM allele a JOIN locus l USING(locus_id) WHERE a.profile_id=?", (pid,))}
    res = {"load_s": round(load_s, 1), "size_MB": round(size, 1), "profiles": n}
    def timeit(fn, reps=len(pids)):
        ts = []
        for p in pids[:reps]: t = time.perf_counter(); fn(p); ts.append((time.perf_counter() - t) * 1000)
        return {"median_ms": statistics.median(ts), "p95_ms": sorted(ts)[int(.95 * len(ts)) - 1], "max_ms": max(ts)}
    res["exact_20_loci"] = timeit(lambda p: A.sqlite_search(pth, gq(p), "high", 8, 0, 100))
    res["partial_10_loci"] = timeit(lambda p: A.sqlite_search(pth, dict(list(gq(p).items())[:10]), "high", 8, 0, 100))
    res["partial_8_loci"] = timeit(lambda p: A.sqlite_search(pth, dict(list(gq(p).items())[:8]), "high", 8, 0, 100))
    res["mismatch_2_of_20"] = timeit(lambda p: A.sqlite_search(pth, gq(p), "high", 8, 2, 100), 30)
    res["moderate_20_loci"] = timeit(lambda p: A.sqlite_search(pth, gq(p), "moderate", 8, 0, 100), 15)
    t = time.perf_counter(); A.build_allele_frequencies(pth, "perf", True); res["build_frequencies_s"] = round(time.perf_counter() - t, 2)
    t = time.perf_counter(); A.locus_statistics(pth, True); res["locus_statistics_s"] = round(time.perf_counter() - t, 2)
    perf[n] = res; con.close(); print("perf", n, json.dumps(res)[:200], flush=True)
    if n >= 100000 and os.path.exists(pth) and a.perf.split(",")[-1] != str(n): pass
R["performance"] = perf
json.dump(R, open(os.path.join(a.out, "results.json"), "w"), indent=1, default=str)
print(json.dumps({k: v for k, v in R.items() if k not in ("stats_locus", "kinship", "performance")}, indent=1, default=str)[:6000])
