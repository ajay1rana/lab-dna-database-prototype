# DNA working database (SQLite) - data dictionary

File: `Forensic_DNA_Database.db` - converted from your Excel workbook (7 profiles, 427 fields verified, 0 mismatches).
Open it with the app (Browse -> the `.db` file) or any SQLite tool (DB Browser for SQLite, Python, R, DBeaver).

## Core tables
| Table | One row per | Holds |
|---|---|---|
| `profile` | DNA profile | Profile number (`profile_ref`, e.g. DNA-000004), **De-identified Code** (`specimen_id`), **Identifiable Code**, category, index group, sample/exhibit/person codes, collection / received / analysis dates, batch, kit, instrument, software, analyst, technical reviewer + date, administrative reviewer + date, entered-by / entry date, profile type, number of contributors, sex + Amelogenin, Y-indel, Y-STR / X-STR text, loci typed, record status, remarks, **upload status/date, expungement date/reason, retention-until** |
| `case_record` | case / requisition | Lab case no. (report no. + date), requisition reference / date / forwarding authority, FIR/case no. + date, sections of law, police station, district, state, case type. Many profiles can belong to one case. |
| `allele` | profile x locus | `a1`, `a2` = allele x 10 as integers (9.3 -> 93), `a1 <= a2` enforced. Indexed on (locus, a1, a2) for fast matching. All 20 CODIS core loci + SE33, Penta D, Penta E, D6S1043. |
| `strmarker` | profile x Y/X-STR marker | e.g. DYS391 = 10 (parsed from the Y-STR text; the original text is kept in `profile.ystr_text`). |
| `locus` | locus | name, core flag, chromosome, repeat length, plausible allele range. |

## Audit and control tables
| Table | Purpose |
|---|---|
| `audit_log` | Automatic (triggers): every profile insert/delete, every change to status, category, kit, reviewer, review date, upload status, expungement, sample type, person ID, case link, analysis date, profile type, contributors, and every allele edit - with user, time, old and new value. |
| `import_log` | One row per import: who, when, source, rows read / added / not added. |
| `search_log` | One row per match search: who, when, mode, loci, mismatches, number of results, a hash of the query. |
| `config` | Laboratory code, **secret code keys**, minimum loci for duplicate check, audit switch, current user. Protect this file (access control / disk encryption). |

## Reference tables (rules the app reads)
`ref_category` (category -> index group, minimum core loci), `ref_list` (case types, sample types, profile types, statuses, instruments, software, states, districts), `ref_kit`, `ref_kit_rule` (marker rules used to detect the kit), `ref_location` (keyword -> state / district), `ref_field` + `ref_alias` (header spellings and default values for importing), `ref_locus_alias`. Edit these tables to change the app's behaviour; nothing is hard-wired except the built-in state / district lists.

## Views
- `v_profile_full` - profile + case details, one row per profile (for reports).
- `v_deidentified` - **safe to share**: De-identified Code, category, index group, profile type, kit, sex, loci typed (no case, person, place or date).
- `v_allele_named` - alleles with locus names and decimal values.

## Built-in safeguards
Unique profile number, De-identified Code and Identifiable Code; foreign keys; `a1 <= a2`; duplicate profile + same case reference blocked by the app; transactions (an import is all-or-nothing); WAL journaling.

## Useful queries
```sql
SELECT * FROM v_profile_full WHERE state = 'Kerala';
SELECT specimen_category, COUNT(*) FROM profile GROUP BY 1;
SELECT * FROM audit_log WHERE record_ref = 'DNA-000004' ORDER BY audit_id;
SELECT deidentified_code, locus, allele_1, allele_2 FROM v_allele_named WHERE deidentified_code = 'DI-FZX7-7KY';
```

## Growing the schema
SQLite allows `ALTER TABLE profile ADD COLUMN ...` at any time (e.g. quantification, extraction method, consent/legal basis). Existing data and the app keep working.

## Y-STR database (added)
| Table / view | Purpose |
|---|---|
| `strmarker` (kind = 'Y') | One row per Y-STR marker per profile; marker names are standardised (DYS019 -> DYS19, DYS385a + DYS385b -> DYS385 "11,14", Y-GATA-H4 -> YGATAH4). |
| `ystr_profile` | One row per profile with Y-STR data: detected Y-STR kit (Yfiler Plus, PowerPlex Y23, Yfiler, Minimal haplotype (YHRD), or "Partial (n loci)"), number of loci, `haplotype_key` (all values in standard order, for exact whole-haplotype lookup). |
| `ref_ystr_locus` | The 27 standard Y-STR loci (copies: DYS385 and DYF387S1 = 2) and whether each is in the minimal haplotype. |
| `ref_ystr_kit` | Which loci each standard Y-STR kit covers (verify against the manufacturer's documentation). |
| `v_ystr_haplotype` | De-identified code + marker + value for every Y-STR result. |

Y-STR matches are lineage (paternal line) evidence: report them as haplotype counts, not individual identification.

## Minimum loci per category
`ref_category.min_core_loci` (defaults: 20 for offender / suspect / reference / elimination / victim / relative, 10 for forensic casework, missing persons, unidentified remains, other).
Profiles with fewer typed core loci are marked "Review" on import. Change it in the app ("Minimum loci per category...") or with
`UPDATE ref_category SET min_core_loci = 8 WHERE category = 'Crime Scene';`

## Match statistics (added)
| Table | Purpose |
|---|---|
| `population` | A named allele-frequency set (internal database, or published frequencies imported from CSV) with N individuals and source. |
| `allele_freq` | Allele counts and frequencies per population and locus. Minimum allele frequency 5/2N is applied when a rare allele is not in the table. |
| `config.stats_theta` | Co-ancestry (theta) correction used by the random match probability (default 0.01 - set the value your laboratory has validated). |

Command line: `--build-freq`, `--import-freq`, `--locus-stats`, `--rmp`, `--kinship`, `--ystr-freq`, `--ystr-diversity`, and `--population` with `--sqlite-search`.
Reported: random match probability (NRC II 4.10, theta-corrected), likelihood ratios for parent-child / full-sibling / half-sibling / other kinship hypotheses, Ho / He / PIC / power of discrimination, Hardy-Weinberg tests, Y-STR counting-method frequency with an exact 95% upper limit, haplotype diversity.
Statistics from the laboratory's own database are descriptive; for casework reporting use a validated population database.
