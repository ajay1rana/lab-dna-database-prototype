Reproduce the validation (put dna_import_app.py in the folder named by DNA_APP_DIR; TEMPLATE_DB = any database created by the app):
  python generate_dataset.py --out synth            (seed 20260101)
  DNA_APP_DIR=../S1_Software TEMPLATE_DB=template.db python run_validation.py --data synth --out results --perf 1000,10000,100000,300000
  python make_figures.py results/results.json
Requires numpy, openpyxl, matplotlib. The 300,000-profile benchmark takes several minutes.
