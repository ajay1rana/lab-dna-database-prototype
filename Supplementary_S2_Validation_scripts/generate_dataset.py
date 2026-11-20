#!/usr/bin/env python3
"""Synthetic dataset generator for validating the laboratory DNA database prototype. Everything is simulated (no real persons/cases). Fixed seed = reproducible.
Outputs (in --out): truth_individuals.csv, truth_profiles.csv, truth_relations.csv, truth_errors.csv, truth_duplicates.json, population_frequencies.csv, ystr_truth.csv,
import_horizontal_*.xlsx, import_vertical.xlsx, duplicates_horizontal.xlsx, errors_horizontal.xlsx"""
import argparse, csv, json, os, random
import numpy as np
from openpyxl import Workbook

CORE = ["D3S1358","vWA","D16S539","CSF1PO","TPOX","D8S1179","D21S11","D18S51","D5S818","D13S317","D7S820","TH01","D12S391","D19S433","FGA","D2S1338","D10S1248","D22S1045","D1S1656","D2S441"]
SUPP = ["SE33","Penta D","Penta E","D6S1043"]
RANGES = {"D3S1358":(12,20),"vWA":(11,24),"D16S539":(5,15),"CSF1PO":(6,15),"TPOX":(6,13),"D8S1179":(7,19),"D21S11":(24,38),"D18S51":(9,27),"D5S818":(7,16),"D13S317":(7,15),"D7S820":(6,15),"TH01":(4,11),"D12S391":(15,27),"D19S433":(9,17),"FGA":(17,30),"D2S1338":(15,28),"D10S1248":(10,18),"D22S1045":(8,19),"D1S1656":(9,20),"D2S441":(8,16),"SE33":(14,32),"Penta D":(5,15),"Penta E":(5,20),"D6S1043":(10,21)}
MICRO = {"TH01":[9.3],"D21S11":[29.2,30.2,31.2,32.2,33.2],"FGA":[21.2,22.2,23.2,24.2,26.2],"D12S391":[17.3,18.3],"D19S433":[12.2,13.2,14.2,15.2],"D1S1656":[15.3,16.3,17.3,18.3],"SE33":[17.2,20.2,24.2,25.2,26.2,27.2,28.2,29.2]}
YSTR = {"DYS19":(13,17),"DYS385":(10,17),"DYS389I":(11,15),"DYS389II":(27,31),"DYS390":(21,26),"DYS391":(9,12),"DYS392":(11,15),"DYS393":(12,15),"DYS437":(13,16),"DYS438":(9,13),"DYS439":(10,14),"DYS448":(17,21),"DYS456":(14,18),"DYS458":(14,19),"DYS635":(20,25),"YGATAH4":(10,13)}
KITS = {"VeriFiler Plus":(CORE+["D6S1043","Penta D","Penta E"],True,[]),"GlobalFiler":(CORE+["SE33"],True,[]),
        "PowerPlex Fusion 6C":(CORE+["SE33","Penta D","Penta E"],False,["DYS391","DYS576","DYS570"]),"PowerPlex Fusion":(CORE+["Penta D","Penta E"],False,[]),
        "Identifiler Plus":(["D8S1179","D21S11","D7S820","CSF1PO","D3S1358","TH01","D13S317","D16S539","D2S1338","D19S433","vWA","TPOX","D18S51","D5S818","FGA"],False,[])}
PLACES = [("Kadthal","Rangareddy","Telangana","Dy. Commissioner of Police, Shamshabad Zone, Cyberabad, TS"),("Shadnagar","Rangareddy","Telangana","Commissioner of Police, Cyberabad"),
          ("Kanjirappally","Kottayam","Kerala","Judicial First Class Magistrate, Kanjirappally, Kottayam, Kerala"),("Yanam","Yanam","Puducherry","Principal Magistrate, Juvenile Justice Board, Yanam"),
          ("Baratang","North and Middle Andaman","Andaman and Nicobar Islands","Superintendent of Police, Andaman and Nicobar Islands"),("Chevella","Rangareddy","Telangana","Commissioner of Police, Cyberabad, Telangana"),
          ("Shamshabad","Rangareddy","Telangana","Commissioner of Police, Cyberabad"),("Nandigama","Rangareddy","Telangana","Commissioner of Police, Cyberabad, TS"),
          ("Amangal","Rangareddy","Telangana","Commissioner of Police, Cyberabad, TS"),("Kothur","Rangareddy","Telangana","Commissioner of Police, Cyberabad, TS")]
CATS = [("Convicted",1),("Accused",1),("Reference",1),("Crime Scene",0),("Unknown",0),("Missing Person",0),("Staff Elimination",1)]

def alleles_for(l):
    lo, hi = RANGES[l]; return [float(a) for a in range(lo, hi+1)] + [float(a) for a in MICRO.get(l, [])]
def fmt(a): return str(int(a)) if float(a).is_integer() else str(a)
def make_population(rng):
    f = {}
    for l in CORE+SUPP:
        al = alleles_for(l); w = rng.dirichlet(np.ones(len(al))*1.2)**1.6; w = np.maximum(w/w.sum(), 0.0015); f[l] = (al, w/w.sum())
    return f
def draw_genotype(rng, f, loci): return {l: tuple(sorted(rng.choice(f[l][0], size=2, p=f[l][1]))) for l in loci}
def child_of(rng, g1, g2, mu=0.002):
    out = {}
    for l in g1:
        a = float(rng.choice(g1[l])); b = float(rng.choice(g2[l]))
        if rng.random() < mu: a += rng.choice([-1.0, 1.0])
        if rng.random() < mu: b += rng.choice([-1.0, 1.0])
        out[l] = tuple(sorted((a, b)))
    return out
def make_ystr_population(rng, n_hap=400):
    haps, seen = [], set()
    while len(haps) < n_hap:
        h = {m: int(rng.integers(lo, hi+1)) for m, (lo, hi) in YSTR.items() if m != "DYS385"}
        x, y = sorted((int(rng.integers(10, 18)), int(rng.integers(10, 18)))); h["DYS385"] = f"{x},{y}"
        k = json.dumps(h, sort_keys=True)
        if k not in seen: seen.add(k); haps.append(h)
    w = 1.0/np.arange(1, n_hap+1)**0.9; return haps, w/w.sum()

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--out", default="synth"); ap.add_argument("--seed", type=int, default=20260101)
    ap.add_argument("--n", type=int, default=1500); ap.add_argument("--pairs", type=int, default=100); ap.add_argument("--vertical", type=int, default=60); a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True); rng = np.random.default_rng(a.seed); prng = random.Random(a.seed)
    freqs = make_population(rng); loci = CORE+SUPP
    with open(os.path.join(a.out, "population_frequencies.csv"), "w", newline="") as f:
        w = csv.writer(f); w.writerow(["locus","allele","frequency"])
        for l, (al, p) in freqs.items():
            for x, y in zip(al, p): w.writerow([l, fmt(x), f"{y:.6f}"])
    haps, hw = make_ystr_population(rng)
    with open(os.path.join(a.out, "ystr_truth.csv"), "w", newline="") as f:
        w = csv.writer(f); w.writerow(["hap_id","true_frequency"]+list(YSTR))
        for i, (h, p) in enumerate(zip(haps, hw)): w.writerow([i, f"{p:.8f}"]+[h[m] for m in YSTR])
    persons = []
    def new_person(geno=None, sex=None):
        sx = sex or ("M" if rng.random() < 0.5 else "F")
        p = {"pid": f"P{len(persons)+1:06d}", "geno": geno or draw_genotype(rng, freqs, loci), "sex": sx, "hap": int(rng.choice(len(haps), p=hw)) if sx == "M" else None}
        persons.append(p); return p
    for _ in range(a.n): new_person()
    relations = []
    for _ in range(a.pairs):
        par = new_person(); mate = new_person(); ch = new_person(child_of(rng, par["geno"], mate["geno"])); relations.append((par["pid"], ch["pid"], "parent-child"))
    for _ in range(a.pairs):
        f_, m_ = new_person(sex="M"), new_person(sex="F"); s1, s2 = new_person(child_of(rng, f_["geno"], m_["geno"])), new_person(child_of(rng, f_["geno"], m_["geno"])); relations.append((s1["pid"], s2["pid"], "full-siblings"))
    relations += [(persons[int(i)]["pid"], persons[int(j)]["pid"], "unrelated") for i, j in zip(rng.integers(0, a.n, a.pairs), rng.integers(0, a.n, a.pairs)) if i != j]
    with open(os.path.join(a.out, "truth_individuals.csv"), "w", newline="") as f:
        w = csv.writer(f); w.writerow(["person","sex","hap_id"]+[f"{l}_{k}" for l in loci for k in (1,2)])
        for p in persons: w.writerow([p["pid"], p["sex"], p["hap"] if p["hap"] is not None else ""]+[fmt(x) for l in loci for x in p["geno"][l]])
    with open(os.path.join(a.out, "truth_relations.csv"), "w", newline="") as f:
        w = csv.writer(f); w.writerow(["person_a","person_b","relation"]); w.writerows(relations)
    profiles = []; kit_names = list(KITS)
    def make_profile(person, kit=None, cat=None, place=None, error=None, ref=None):
        kit = kit or prng.choice(kit_names); place = place or prng.choice(PLACES); markers, yindel, ys = KITS[kit]; drop = 0
        if cat is None:
            cat = prng.choice(CATS)[0] if sum(1 for m in markers if m in CORE) == 20 else prng.choice(["Crime Scene","Unknown","Missing Person"])
            if cat in ("Crime Scene","Unknown","Missing Person"): drop = prng.choice([0,0,2,5,8])
        typed = list(markers)
        if drop: typed = prng.sample(typed, max(len(typed)-drop, 1))
        pr = {"ref": ref or f"DNA-{len(profiles)+1:06d}", "person": person["pid"], "kit": kit, "category": cat, "typed": typed, "yindel": (2 if person["sex"]=="M" else None) if yindel else None,
              "sex": person["sex"], "place": place, "fir": f"{prng.randint(1,999)}/{prng.choice([2021,2022,2023,2024])}", "error": error,
              "ystr": (dict(haps[person["hap"]]) if (person["hap"] is not None and prng.random() < 0.6) else {}),
              "report": f"CFSL(H)/{prng.randint(100,999)}/DNA/{prng.randint(1,99)}/{prng.choice([2022,2023])}", "geno": person["geno"]}
        profiles.append(pr); return pr
    for p in persons: make_profile(p)
    dups_exact, dups_other = [], []
    for pr in prng.sample(profiles[:a.n], 60):
        d = dict(pr); d["ref"] = f"DNA-{len(profiles)+1:06d}"; d["error"] = "exact_duplicate"; profiles.append(d); dups_exact.append(d["ref"])
    for pr in prng.sample(profiles[:a.n], 60):
        d = dict(pr); d["ref"] = f"DNA-{len(profiles)+1:06d}"; d["fir"] = f"{prng.randint(1000,1999)}/2025"; d["place"] = prng.choice(PLACES); d["error"] = "same_person_other_case"; profiles.append(d); dups_other.append(d["ref"])
    json.dump({"exact_duplicates": dups_exact, "same_person_other_case": dups_other, "seed": a.seed}, open(os.path.join(a.out, "truth_duplicates.json"), "w"))
    base = [p for p in profiles if p["error"] is None]; vert = base[:a.vertical]; horiz = base[a.vertical:]
    chunks = [horiz[i:i+450] for i in range(0, len(horiz), 450)]
    def hv(l, s): return {0:(f"{l} Allele 1",f"{l} Allele 2"),1:(f"{l}_1",f"{l}_2"),2:(f"{l} (1)",f"{l} (2)"),3:(l,)}[s]
    def write_h(path, chunk, mutate=None):
        wb = Workbook(); ws = wb.active; ws.title = "Profiles"; style = {l: prng.choice([0,1,2,3]) for l in loci}
        heads = ["Sample Name","Case/FIR No.","Police Station","Forwarding Authority","Report Number","Profile Category","Sample Type","Analyst","Analysis Date","Amelogenin (1)","Amelogenin (2)","Y-indel"]
        if mutate: heads.append("Kit/Method")
        for l in loci: heads += list(hv(l, style[l]))
        yst = sorted({m for p in chunk for m in p["ystr"]}); heads += yst
        order = list(range(len(heads))); prng.shuffle(order); ws.append([heads[i] for i in order])
        for p in chunk:
            row = {"Sample Name": p["ref"], "Case/FIR No.": p["fir"], "Police Station": p["place"][0], "Forwarding Authority": p["place"][3], "Report Number": p["report"], "Profile Category": p["category"],
                   "Sample Type": "Blood", "Analyst": "Analyst A", "Analysis Date": "05-11-2026", "Amelogenin (1)": "X", "Amelogenin (2)": "X" if p["sex"]=="F" else "Y", "Y-indel": p["yindel"]}
            for l in loci:
                if l in p["typed"]:
                    g = p["geno"][l]; s = style[l]
                    if s == 3: row[l] = f"{fmt(g[0])},{fmt(g[1])}"
                    else: h = hv(l, s); row[h[0]], row[h[1]] = fmt(g[0]), fmt(g[1])
            for m in yst:
                if m in p["ystr"] and p["ystr"][m] is not None: row[m] = p["ystr"][m]
            if mutate: row["Kit/Method"] = p["kit"]; mutate(row, p, style)
            ws.append([row.get(heads[i]) for i in order])
        wb.save(path)
    for ci, ch in enumerate(chunks): write_h(os.path.join(a.out, f"import_horizontal_{ci+1}.xlsx"), ch)
    write_h(os.path.join(a.out, "duplicates_horizontal.xlsx"), [p for p in profiles if p["error"] in ("exact_duplicate","same_person_other_case")])
    err_types = ["non_numeric","tri_allelic","single_allele","out_of_range","below_min_loci","missing_mandatory","bad_category","swapped_order","clean_control"]
    err_rows = []; k = 0
    for et in err_types:
        for _ in range(12):
            person = new_person(); k += 1
            pr = make_profile(person, kit="VeriFiler Plus", cat="Convicted", place=prng.choice(PLACES), error="err_"+et, ref=f"ERR-{k:04d}"); pr["fir"] = f"E{k}/2026"; err_rows.append((pr, et))
    emap = {q["ref"]: e for q, e in err_rows}
    def mut(row, p, style):
        et = emap[p["ref"]]; l = prng.choice(CORE); g = p["geno"][l]; s = style[l]; h = hv(l, s)
        if et == "non_numeric": row[h[0]] = "OL"
        if et == "tri_allelic":
            if s == 3: row[l] = f"{fmt(g[0])},{fmt(g[1])},{fmt(g[1]+1)}"
            else: row[h[0]] = None; row[h[1]] = None
        if et == "single_allele": row[h[-1]] = f"{fmt(g[0])}" if s == 3 else None
        if et == "out_of_range":
            if s == 3: row[l] = f"99,{fmt(g[1])}"
            else: row[h[0]] = "99"
        if et == "below_min_loci":
            for ll in CORE[12:]:
                for hh in hv(ll, style[ll]): row[hh] = None
        if et == "missing_mandatory": row["Analyst"] = None
        if et == "bad_category": row["Profile Category"] = "Convict"
        if et == "swapped_order":
            if s == 3: row[l] = f"{fmt(max(g))},{fmt(min(g))}"
            else: row[h[0]], row[h[1]] = fmt(max(g)), fmt(min(g))
    write_h(os.path.join(a.out, "errors_horizontal.xlsx"), [r[0] for r in err_rows], mut)
    exp = {"non_numeric":"Review","tri_allelic":"Review","single_allele":"Review","out_of_range":"Review","below_min_loci":"Review","missing_mandatory":"Incomplete","bad_category":"Review","swapped_order":"Ready","clean_control":"Ready"}
    with open(os.path.join(a.out, "truth_errors.csv"), "w", newline="") as f:
        w = csv.writer(f); w.writerow(["sample","error_type","expected_status"])
        for pr, et in err_rows: w.writerow([pr["ref"], et, exp[et]])
    wb = Workbook(); wb.remove(wb.active)
    for p in vert:
        ws = wb.create_sheet(p["ref"])
        rows = [("Amelogenin", "X, "+("X" if p["sex"]=="F" else "Y"))]+[(l, f"{fmt(p['geno'][l][0])}, {fmt(p['geno'][l][1])}") for l in loci if l in p["typed"]]
        if p["yindel"]: rows.append(("Yindel", p["yindel"]))
        rows += [(m, v) for m, v in p["ystr"].items() if v is not None]; prng.shuffle(rows)
        for i, (kk, v) in enumerate(rows, 1): ws.cell(row=i, column=1, value=kk); ws.cell(row=i, column=2, value=v)
        ws["D1"]="Report Number"; ws["E1"]="DATE"; ws["D3"]=p["report"]; ws["E3"]="26-07-2022"
        ws["D5"]="Case Number"; ws["E5"]="Date"; ws["F5"]="U/S"; ws["G5"]="PS"; ws["D7"]=p["fir"]; ws["E7"]="10-01-2021"; ws["F7"]="376 IPC"; ws["G7"]=p["place"][0]
        ws["D9"]="Reference Number"; ws["E9"]="Date"; ws["F9"]="From"; ws["D11"]="R/"+p["fir"]; ws["E11"]="11-05-2022"; ws["F11"]=p["place"][3]
        ws["D13"]="Lab Sample ID"; ws["E13"]=p["ref"]
    wb.save(os.path.join(a.out, "import_vertical.xlsx"))
    vset = {p["ref"] for p in vert}
    with open(os.path.join(a.out, "truth_profiles.csv"), "w", newline="") as f:
        w = csv.writer(f); w.writerow(["profile_ref","person","sex","kit","category","fir","police_station","district","state","report","n_loci_typed","min_required","typed_loci","yindel","ystr","planted","layout"])
        for p in profiles:
            w.writerow([p["ref"],p["person"],p["sex"],p["kit"],p["category"],p["fir"],p["place"][0],p["place"][1],p["place"][2],p["report"],len([l for l in p["typed"] if l in CORE]),
                        (10 if p["category"] in ("Crime Scene","Unknown","Missing Person","Other") else 20),"|".join(p["typed"]),p["yindel"] or "",json.dumps(p["ystr"]),p["error"] or "","vertical" if p["ref"] in vset else "horizontal"])
    print(f"individuals {len(persons)}, profiles {len(profiles)}, vertical {len(vert)}, horizontal files {len(chunks)}")
if __name__ == "__main__": main()
