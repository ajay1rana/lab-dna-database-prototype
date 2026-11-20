#!/usr/bin/env python3
"""Figures for the manuscript from results/results.json -> figures/*.png"""
import json, os, sys, math
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
R = json.load(open(sys.argv[1] if len(sys.argv) > 1 else "results/results.json")); os.makedirs("figures", exist_ok=True)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.spines.top": False, "axes.spines.right": False})
C = ["#1f4e79", "#c0504d", "#2e8b57", "#d98c00", "#6a5acd"]
def box(ax, x, y, w, h, t, fc="#eaf1f8", ec="#1f4e79", fs=8):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.08", fc=fc, ec=ec, lw=1.1)); ax.text(x + w / 2, y + h / 2, t, ha="center", va="center", fontsize=fs)
def arrow(ax, x1, y1, x2, y2): ax.annotate("", xy=(x2, y2), xytext=(x1, y1), arrowprops=dict(arrowstyle="->", color="#444", lw=1.1))
# Fig 1 architecture
fig, ax = plt.subplots(figsize=(7.2, 4.2)); ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
box(ax, .2, 4.6, 2.2, .9, "Excel source files\n(vertical PI sheets,\nhorizontal tables)"); box(ax, 3.0, 4.6, 2.0, .9, "Reader / parser\n(any kit, any\ncolumn order)")
box(ax, 5.6, 4.6, 2.0, .9, "Quality control\n(range, zygosity, min\nloci, duplicates)", fc="#fff4e0", ec="#d98c00"); box(ax, 8.0, 4.6, 1.8, .9, "Preview\n(Ready / Incomplete /\nReview / Duplicate)", fc="#fff4e0", ec="#d98c00")
arrow(ax, 2.4, 5.05, 3.0, 5.05); arrow(ax, 5.0, 5.05, 5.6, 5.05); arrow(ax, 7.6, 5.05, 8.0, 5.05)
box(ax, 2.6, 2.7, 4.8, 1.3, "SQLite working database (single file, WAL)\nprofile · case_record · allele · strmarker · ystr_profile\naudit_log · import_log · search_log · config · ref_* · population/allele_freq", fc="#e6f4ea", ec="#2e8b57")
arrow(ax, 8.9, 4.6, 6.6, 4.0)
box(ax, .2, .6, 2.2, 1.1, "Pseudonymisation\nDI-XXXX-XXX\nLAB-IDX-YYYY-NNNNNN-C", fc="#f0e8f8", ec="#6a5acd"); box(ax, 2.8, .6, 2.2, 1.1, "Search\nexact · pigeonhole\nmismatch · moderate\nY-STR haplotype", fc="#f0e8f8", ec="#6a5acd")
box(ax, 5.4, .6, 2.1, 1.1, "Match statistics\nRMP (θ) · kinship LR\nHo/He/PIC/HWE\nY-STR counting", fc="#f0e8f8", ec="#6a5acd"); box(ax, 7.9, .6, 1.9, 1.1, "Exports\nde-identified XML/JSON\nSQLite · Y-STR", fc="#f0e8f8", ec="#6a5acd")
for x in (1.3, 3.9, 6.45, 8.85): arrow(ax, 5.0, 2.7, x, 1.7)
ax.text(5, 0.15, "Laboratory-level prototype: no network service, no personal data in de-identified views", ha="center", fontsize=7.5, style="italic", color="#444")
fig.savefig("figures/fig1_architecture.png", dpi=220, bbox_inches="tight"); plt.close(fig)
# Fig 2 data model
fig, ax = plt.subplots(figsize=(7.2, 4.4)); ax.set_xlim(0, 10); ax.set_ylim(0, 6.2); ax.axis("off")
box(ax, 3.7, 3.9, 2.6, 1.9, "profile\nprofile_id PK · profile_ref\nspecimen_id (DI code)\nidentifiable_code\ncategory · kit · sex\ngenotype_key · ref_sig", fc="#e6f4ea", ec="#2e8b57", fs=7.5)
box(ax, .2, 4.1, 2.6, 1.5, "case_record\nlab case no. + date\nFIR / requisition\npolice station\ndistrict · state", fs=7.5); box(ax, 7.2, 4.4, 2.6, 1.2, "audit_log\nimport_log · search_log\n(trigger-written)", fc="#fff4e0", ec="#d98c00", fs=7.5)
box(ax, .2, 1.8, 2.6, 1.3, "allele\n(profile_id, locus_id) PK\na1 ≤ a2 (×10 integers)\nindex (locus, a1, a2)", fs=7.5); box(ax, 3.7, 1.8, 2.6, 1.3, "strmarker\nkind Y / X · marker · value\nystr_profile\nhaplotype_key", fs=7.5)
box(ax, 7.2, 1.8, 2.6, 1.3, "population · allele_freq\nθ in config\nstatistics tables", fc="#f0e8f8", ec="#6a5acd", fs=7.5); box(ax, 1.6, .1, 3.0, 1.1, "locus · ref_category · ref_kit\nref_kit_rule · ref_location\nref_alias (rules as data)", fc="#f4f4f4", ec="#777", fs=7.5)
box(ax, 5.4, .1, 3.0, 1.1, "views: v_profile_full\nv_deidentified (shareable)\nv_allele_named · v_ystr_haplotype", fc="#f4f4f4", ec="#777", fs=7.5)
arrow(ax, 2.8, 4.85, 3.7, 4.85); arrow(ax, 1.5, 4.1, 1.5, 3.1); arrow(ax, 4.9, 3.9, 4.9, 3.1); arrow(ax, 6.3, 4.8, 7.2, 4.9); arrow(ax, 6.2, 3.9, 7.7, 3.1)
fig.savefig("figures/fig2_data_model.png", dpi=220, bbox_inches="tight"); plt.close(fig)
# Fig 3 performance
P = R["performance"]; ns = sorted(int(k) for k in P)
fig, axs = plt.subplots(1, 2, figsize=(7.4, 3.2))
for key, lab, c in (("exact_20_loci", "Exact, 20 loci", C[0]), ("partial_10_loci", "Partial, 10 loci", C[2]), ("partial_8_loci", "Partial, 8 loci", C[3]), ("mismatch_2_of_20", "≤2 mismatches", C[1]), ("moderate_20_loci", "Moderate stringency", C[4])):
    axs[0].plot(ns, [P[str(n)][key]["median_ms"] for n in ns], "o-", color=c, label=lab, lw=1.5, ms=4)
axs[0].set_xscale("log"); axs[0].set_yscale("log"); axs[0].set_xlabel("Profiles in database"); axs[0].set_ylabel("Median search time (ms)"); axs[0].legend(fontsize=7, frameon=False); axs[0].set_title("a  Search latency", loc="left", fontsize=9)
axs[1].bar(range(len(ns)), [P[str(n)]["size_MB"] for n in ns], color=C[0], width=.55); axs[1].set_xticks(range(len(ns))); axs[1].set_xticklabels([f"{n:,}" for n in ns], fontsize=7)
for i, n in enumerate(ns): axs[1].text(i, P[str(n)]["size_MB"], f"{P[str(n)]['size_MB']:.0f}", ha="center", va="bottom", fontsize=7)
axs[1].set_ylabel("Database file size (MB)"); axs[1].set_xlabel("Profiles in database"); axs[1].set_title("b  Storage", loc="left", fontsize=9)
fig.tight_layout(); fig.savefig("figures/fig3_performance.png", dpi=220); plt.close(fig)
# Fig 4 kinship LR
K = R["kinship"]; fig, axs = plt.subplots(1, 2, figsize=(7.4, 3.1))
for ax, hyp, title in ((axs[0], "parent-child", "a  H1: parent–child vs unrelated"), (axs[1], "full-siblings", "b  H1: full siblings vs unrelated")):
    data, labs = [], []
    for rel, nm in (("parent-child", "parent–child"), ("full-siblings", "siblings")):
        k = f"{rel}|{hyp}"
        if k in K: data.append(K[k]["values"]); labs.append(f"{nm} pairs\n(n={K[k]['n']})")
    k = f"unrelated|{hyp}"
    if k in K: data.append(K[k]["values"]); labs.append(f"unrelated pairs\n(n={K[k]['n']})")
    bp = ax.boxplot(data, tick_labels=labs, showfliers=True, patch_artist=True, flierprops=dict(markersize=2))
    for p_, c in zip(bp["boxes"], C): p_.set_facecolor(c); p_.set_alpha(.5)
    ax.axhline(0, color="#444", lw=.8, ls="--"); ax.set_ylabel("log10 LR"); ax.set_title(title, loc="left", fontsize=9); ax.tick_params(axis="x", labelsize=7)
fig.tight_layout(); fig.savefig("figures/fig4_kinship_lr.png", dpi=220); plt.close(fig)
# Fig 5 locus statistics
T = R["stats_locus"]["table"]; fig, axs = plt.subplots(1, 2, figsize=(7.4, 3.1)); names = [t["locus"] for t in T]
x = range(len(T)); axs[0].bar(x, [t["PD"] for t in T], color=C[0]); axs[0].set_xticks(list(x)); axs[0].set_xticklabels(names, rotation=90, fontsize=6); axs[0].set_ylim(.7, 1); axs[0].set_ylabel("Power of discrimination"); axs[0].set_title("a  Per-locus power of discrimination", loc="left", fontsize=9)
axs[1].plot([0, 1], [0, 1], color="#999", lw=1); axs[1].scatter([t["He"] for t in T], [t["Ho"] for t in T], color=C[1], s=18); axs[1].set_xlabel("Expected heterozygosity"); axs[1].set_ylabel("Observed heterozygosity"); axs[1].set_title("b  Ho vs He (synthetic data)", loc="left", fontsize=9)
fig.tight_layout(); fig.savefig("figures/fig5_locus_statistics.png", dpi=220); plt.close(fig)
print("figures written")
