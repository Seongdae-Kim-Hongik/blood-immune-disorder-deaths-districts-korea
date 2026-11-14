# cover_bldd.py — BLDD BJH 커버레터(2026-11-14, cover_ijd_sknd 이식). 영국식, 1쪽, 단독 저자 비인칭, 김성대 명의. 자료 중복 고지 금지(저자 결정), 제외 심사자는 포털 칸에만. IRB 번호 포함.
#  Zenodo는 공개(open) 게시, 커버레터에는 공개 DOI만. --draft면 DOI 자리표시자 허용.
import os, re, sys, datetime, subprocess
from docx import Document
from docx.shared import Pt, Cm
import content_bldd as C
from build_bldd import TITLE as DTITLE, ZENODO, ADDRESS
from british import brit
HERE = os.path.dirname(os.path.abspath(__file__)); OUT = os.path.join(os.path.dirname(HERE), "out"); DRAFT = "--draft" in sys.argv
N = C.N; W = dict(l.split("=") for l in open(os.path.join(OUT, "wordcount_bjh.txt")).read().split()); A0 = C.A0
if not DRAFT: assert ZENODO.startswith("10.5281/zenodo."), "공개 DOI 없음(투고 전 게이트)"
_d = datetime.date.today(); cl = [f"{_d.day} {_d:%B} {_d.year}", "The Editor", "British Journal of Haematology", "",
      "Dear Editor,",
      f"Please find enclosed the manuscript entitled \"{DTITLE}\" by Seongdae Kim, for consideration as an Original Paper in the British Journal of Haematology.",
      f"Deaths from anaemias and other blood and immune disorders occur mostly in older people, but their geography within countries is rarely described. This study covers all {N['cases_total']:,} deaths with an underlying cause in ICD-10 D50–D89 in the {N['n_districts']} districts of the Republic of Korea over 2011–2025, of which {N['anaemia_share']:.0f}% were anaemias and {N['pct65']:.0f}% occurred at 65 years or older in the national table. "
      f"After standardisation for sex and age, the district geography was clustered: all {N['n_busan']} districts of Busan were residual hot spots, and Busan had the highest residual rate ratios in each of three periods. None of eight district characteristics, including the standardised ratio of suspected anaemia at national health screening, had a credible association.",
      "These findings are relevant to haematologists and health planners because they document a regional excess that measured anaemia burden, deprivation and care did not account for, point to death-certification and multiple-cause studies in the high-rate area, and show why single-year district figures are unreliable for such rare causes of death. "
      "The study used only published aggregate statistics; the Institutional Review Board of Hongik University approved it (7002340-202511-HR-012) and its continuing review (7002340-202611-HR-040). The manuscript has not been published and is not under consideration elsewhere.",
      f"The main text contains {int(W['body_words'])} words, with a summary of {int(W['summary_words'])} words, two tables, three figures and {W['refs']} references. Supporting information and the completed STROBE checklist are provided as separate files. The analysis code is openly available on Zenodo (https://doi.org/{ZENODO}).",
      "The author used Claude (Anthropic) and ChatGPT (OpenAI) only for translation and English-language editing of the manuscript and takes full responsibility for the content.",
      "Thank you for considering this manuscript.", "", "Yours sincerely,", "",
      "Seongdae Kim, PhD (corresponding author)", ADDRESS, "Telephone: +82-2-320-1114. Email: kimsd@hongik.ac.kr"]
d = Document(); st = d.styles["Normal"]; st.font.name = "Times New Roman"; st.font.size = Pt(11); st.paragraph_format.space_after = Pt(4); st.paragraph_format.line_spacing = 1.0
for s in d.sections: s.page_width, s.page_height = Cm(21), Cm(29.7); s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Cm(2.54)
gap = False
for t in cl:
    if t == "": gap = True; continue
    p = d.add_paragraph(brit(t))
    if gap: p.paragraph_format.space_before = Pt(10); gap = False
d.core_properties.author = "Seongdae Kim"; d.core_properties.last_modified_by = "Seongdae Kim"; d.core_properties.title = "Cover Letter"; d.core_properties.comments = ""
fp = os.path.join(OUT, "Cover_Letter_BJH_BLDD.docx"); d.save(fp)
subprocess.run(["/opt/homebrew/bin/soffice", "--headless", "-env:UserInstallation=file:///tmp/lo_bldd_cl", "--convert-to", "pdf", "--outdir", OUT, fp], capture_output=True)
txt = subprocess.run(["pdftotext", fp.replace(".docx", ".pdf"), "-"], capture_output=True, text=True).stdout
assert not re.search(r"\b(we|our|us|I)\b", txt) and "—" not in txt and re.search(r"7002340-?\s*202611-?\s*HR-040", txt) and not re.search(r"(?i)funding|competing|overlap|previously submitted|token|private link", txt), "문체·IRB·금지 항목"
pages = subprocess.run(["pdfinfo", fp.replace(".docx", ".pdf")], capture_output=True, text=True).stdout.split("Pages:")[1].split()[0]
assert pages == "1", pages; print("cover pages", pages, "draft" if DRAFT else "")
