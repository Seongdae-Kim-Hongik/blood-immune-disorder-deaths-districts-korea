#!/bin/bash
# run_all_bldd.sh — BLDD BJH manuscript build (numbers → refs → figures → manuscript → supporting information → cover letter → checks)
set -e; cd "$(dirname "$0")"
python3 numbers_bldd.py >/dev/null
python3 build_refs_bldd.py >/dev/null
python3 figs_bldd.py
python3 build_bldd.py
python3 esm_bldd.py
python3 cover_bldd.py "$@"
python3 sanitize_docx.py --no-anon ../out/Manuscript_BJH_BLDD.docx ../out/Supporting_Information_BJH_BLDD.docx ../out/Cover_Letter_BJH_BLDD.docx ../out/STROBE_Checklist_BJH_BLDD.docx
python3 check_bldd.py
