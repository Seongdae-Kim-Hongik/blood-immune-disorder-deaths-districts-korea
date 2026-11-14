# build_bldd.py — BLDD British Journal of Haematology Original Paper 원고 조립(2026-11-14, build_ijd_sknd.py 이식).
#  BJH 규정(Wiley Author Guidelines 13652141, 11/14 확인): Summary ≤200, 본문 ≤3,000, 그림+표 ≤7, 참고문헌 ≤60, running title ≤60자, Vancouver, 줄번호,
#   제목면(제목·running title·저자·ORCID·기관·교신·단어/그림/표/참고문헌 수·키워드), 데이터·연구비·이해충돌·윤리·환자 동의·재사용 허가·임상시험 등록 진술, 저자 기여는 Acknowledgements 안. 단일 익명. 소속 = 홍익대 단독.
import os, re, json, subprocess
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from jkms_engine import JKMS
from british import brit
import content_bldd as C
HERE = os.path.dirname(os.path.abspath(__file__)); MS = os.path.dirname(HERE); OUT = os.path.join(MS, "out"); os.makedirs(OUT, exist_ok=True)
N = C.N; REFS = json.load(open(os.path.join(HERE, "refs_bldd.json"), encoding="utf-8"))
DRAFT = os.path.join(MS, "draft", "01_manuscript.md")
_zf = os.path.join(OUT, "zenodo_doi.txt"); ZENODO = open(_zf).read().strip() if os.path.exists(_zf) else "[ZENODO DOI]"
TITLE = C.TITLE; ORCID = "0000-0002-1525-0091"
AFFIL = "Department of Biohealth Convergence Open Sharing System, Hongik University, Seoul, Republic of Korea"
ADDRESS = "Department of Biohealth Convergence Open Sharing System, Hongik University, 94 Wausan-ro, Mapo-gu, Seoul 04066, Republic of Korea"
AI_STATEMENT = ("The author used Claude (Anthropic) and ChatGPT (OpenAI) only for translation and English-language editing of the manuscript. "
                "They were not used to design the study, analyse data, interpret results or generate references. The author reviewed and edited the content and takes full responsibility for it.")
AI_HEAD = "Use of artificial intelligence tools"
PHONE = "+82-2-320-1114"
KEYWORDS = C.KEYWORDS
FUNDING = "None."
COI = "The author declares no conflicts of interest."
ETHICS = "The Institutional Review Board of Hongik University approved the study (7002340-202511-HR-012; continuing review 7002340-202611-HR-040) and waived informed consent because only published aggregate statistics were used."
CONSENT = "Not applicable; only published aggregate statistics were used and no patient was involved."
PERMISSION = "Not applicable; no material was reproduced from other sources."
TRIAL = "Not applicable."
AUTHORSHIP = "S.K. designed the study, assembled and analysed the data and wrote the paper (sole author)."
ACK = "None."
def data_statement(): return (f"All source data are publicly available aggregate statistics from Statistics Korea, the Ministry of the Interior and Safety, the National Health Insurance Service health screening statistics and the other national sources cited in the Methods. "
                             f"The analysis code that generates every number, table and figure is openly available on Zenodo (https://doi.org/{ZENODO}). The processed district-year data are available from the author on reasonable request.")

def read_draft():
    md = open(DRAFT, encoding="utf-8").read(); md = re.sub(r"<!--.*?-->", "", md, flags=re.S); S = {}; cur = None
    for line in md.split("\n"):
        m = re.match(r"^## (.+)$", line)
        if m: cur = m.group(1).strip(); S[cur] = []; continue
        if cur is not None: S[cur].append(line)
    return {k: [p.strip() for p in "\n".join(v).split("\n\n") if p.strip()] for k, v in S.items()}
def pandoc_cites(t): return re.sub(r"\s*\[(@[^\]]+)\]", lambda m: "{{" + ",".join(k.strip().lstrip("@") for k in m.group(1).split(";")) + "}}", t)
def all_keys():
    src = open(os.path.join(HERE, "content_bldd.py"), encoding="utf-8").read() + open(DRAFT, encoding="utf-8").read()
    src = re.sub(r"<!--.*?-->", "", src, flags=re.S); src = re.sub(r"(?m)^\s*#.*$", "", src)
    return sorted(k for k in {k.strip().lstrip("@") for g in re.findall(r"\[(@[^\]]+)\]", src) for k in g.split(";")} if re.fullmatch(r"(web:)?[a-z][a-z0-9]*", k))
KEYS = all_keys()

def ref(k):   # Vancouver: 저자 4명 이하 전부, 넘으면 3명 + et al. MEDLINE 약어. doi 표기.
    r = REFS[k]
    if r["type"] == "web": return f"{r['authors']}. {r['title']} [Internet]. {r['year']}. Available at: {r['url']} (last accessed {r['accessed']})."
    au = r["authors"]; a = ", ".join(au) if len(au) <= 4 else ", ".join(au[:3]) + " et al"
    t = r["title"].rstrip("."); t = t + ("" if t.endswith(("?", "!")) else ".")
    s = f"{a}. {t} {r['journal']} {r['year']}"
    if r.get("vol"): s += f"; {r['vol']}"
    if r.get("pages"): s += f":{r['pages'].replace('-', '–')}"
    s += "."
    if r.get("doi"): s += f" https://doi.org/{r['doi']}"
    elif r.get("pmid"): s += f" PMID: {r['pmid']}."
    return s

class TX(JKMS):
    raw = False
    def _emit(self, p, text, sz=12, bold=False, italic=False):
        if not self.raw:
            text = brit(text); text = re.sub(r"(^|[\s(])-(?=\d)", "\\1\u2212", text)
        return super()._emit(p, text, sz, bold, italic)

TAB1_COV = [("ext__anaemia_suspect_sir", "Suspected anaemia at screening, log standardised ratio"), ("ext__kidney_suspect_share", "Suspected kidney disease at screening, %"), ("basic_livelihood_rate", "Basic livelihood recipients, %"),
            ("tertiary_educated_share", "Tertiary-educated adults, %"), ("ext__unmet_need_rate", "Unmet medical need, %"), ("ltc_hospitals_per_100k_65", "Long-term care hospitals per 100,000 aged ≥ 65 years"),
            ("im_specialists_per_100k", "Internal-medicine specialists, log(1 + per 100,000)"), ("ill_defined_death_share", "Ill-defined causes of death, %"), ("fiscal_independence", "Fiscal autonomy, %"), ("pop80_share", "Population aged ≥ 80 years, %")]
WC = ["W1_2011-15", "W2_2016-19", "W3_2020-25", "ALL"]; WH = ["2011–2015", "2016–2019", "2020–2025", "2011–2025"]
def table1():
    T = N["T1"]; A = N["asr"]
    def mean_asr(a, b): return sum(A[(y, "계")] for y in range(a, b + 1)) / (b - a + 1)
    rows = [["__H__", "Deaths"], ["  District-years, no."] + [f"{T[c]['n']:,}" for c in WC], ["  Deaths (D50–D89), no."] + [f"{T[c]['cases']:,}" for c in WC],
            ["  District-years with no death, no."] + [f"{T[c]['zeros']:,}" for c in WC],
            ["  Age-standardised rate per 100,000*"] + [f"{mean_asr(a, b):.2f}" for a, b in ((2011, 2015), (2016, 2019), (2020, 2025), (2011, 2025))],
            ["  District SMR, median (P10–P90)†"] + [f"{T[c]['med']:.2f} ({T[c]['p10']:.2f}–{T[c]['p90']:.2f})" for c in WC], ["__H__", "District characteristics, mean (SD)"]]
    for v, lab in TAB1_COV:
        mx = max(abs(T[c]['cov'][v][0]) for c in WC); dp = 0 if mx >= 100 else (2 if mx < 1 else 1)
        rows.append([f"  {lab}"] + [f"{T[c]['cov'][v][0]:.{dp}f} ({T[c]['cov'][v][1]:.{dp}f})" for c in WC])
    foot = ("P10–P90, 10th–90th percentile; SD, standard deviation; SMR, standardised mortality ratio. *Mean of annual national rates, 2020 resident population as standard. †Observed over sex- and age-standardised expected deaths summed over the period. "
            "Characteristics are district-year means; the other forced controls are in Table S1. Values not yet published for 2025 carried their 2024 values.")
    return (f"Table 1 Blood and immune disorder deaths and district characteristics, {N['n_districts']} districts, Republic of Korea, 2011–2025", ["Characteristic"] + WH, rows, foot)

GROUPS = [("Anaemia and kidney disease burden", [C.ANE, C.KID]), ("Deprivation and access", [C.BLS, C.TE, C.UNM]), ("Care and certification", [C.LTC, C.IMS, C.ILL]),
          ("Forced controls", ["urban_population_share", "physicians_per_1000", "sex_ratio", "fiscal_independence", "pop80_share"])]
LABEL = {C.ANE: "Suspected anaemia, standardised ratio (+)", C.KID: "Suspected kidney disease (+)", C.BLS: "Basic livelihood recipients (+)", C.TE: "Tertiary-educated adults (−)", C.UNM: "Unmet medical need (+)",
         C.LTC: "Long-term care hospitals (+)", C.IMS: "Internal-medicine specialists (−)", C.ILL: "Ill-defined causes of death (−)",
         "urban_population_share": "Urban population share", "physicians_per_1000": "Physicians per 1,000", "sex_ratio": "Sex ratio", "fiscal_independence": "Fiscal autonomy", "pop80_share": "Population aged ≥ 80 years"}
def f3(t): return f"{t[0]:.3f} ({t[1]:.3f}–{t[2]:.3f})"
def table2():
    rows = []; assert sum(len(v) for g, v in GROUPS[:-1]) == N["n_hyp"] and set(LABEL) == set(N["main"])
    for g, vs in GROUPS:
        rows.append(["__H__", g])
        for v in vs:
            m = N["main"][v]; mv = N["mv"].get(v); k = N["s43"][(C.BEST_KH, v)]; y24 = N["sens"]["v2y24"][v]; sc = N["sens"]["v2sc"][v]; tr = N["nco"][("TRANS", v)]
            rows.append([f"  {LABEL[v]}", f3((m["IRR"], m["lo"], m["hi"])), f"{100*mv['cred_frac']:.0f}" if mv else "Forced", f3(k), f3(y24), f3(sc), f3(tr)])
    foot = (f"CrI, credible interval; RR, rate ratio per 1 standard deviation; expected direction in parentheses. Principal model: negative binomial with a Besag–York–Mollié 2 district effect and random-walk and independent year effects, all characteristics and five forced controls (n = {N['n_rows']:,} district-years). "
            "*Share of the 200 covariate-subset models containing the characteristic in which the CrI excluded 1. †Knorr-Held type IV space-time interaction added, the best-fitting interaction model. "
            "‡Crude expected deaths (population × national rate) instead of sex- and age-standardised expected deaths. §Transport-injury deaths as a comparison outcome with the same model.")
    return (f"Table 2 District characteristics and blood and immune disorder deaths, {N['n_districts']} districts, Republic of Korea, 2011–2025",
            ["Characteristic", "Principal RR (95% CrI)", "Subsets credible, %*", "Type IV RR (95% CrI)†", "2011–2024 only", "Crude expected deaths‡", "Transport-injury deaths§"], rows, foot)

FIGCAP = {
 1: ("Deaths from blood and immune disorders (D50–D89), Republic of Korea, 2011–2025. (A) Deaths by year and sex. (B) Age-standardised rates (2020 population) by sex. (C) Age distribution of deaths by sex.",
     "Three panels. Stacked bars rise from about 570 to 940 deaths a year, more in women. Lines show age-standardised rates falling slightly, higher in men. Stacked bars show most deaths at 75 years or older, especially in women."),
 2: ("District geography of blood and immune disorder deaths, 229 districts, 2011–2025. (A) Standardised mortality ratio, 15 years pooled. (B) Residual spatial rate ratio from the principal model. (C) Residual hot spots and cold spots (posterior probability of a rate ratio above 1 > 0.95 or < 0.05). Thick lines are province boundaries.",
     "Three maps of South Korea. Higher ratios and hot spots lie in Busan and the surrounding south-east; lower ratios and cold spots lie mainly in Seoul and Gyeonggi-do in the north-west."),
 3: ("Spatial rate ratios refitted separately in (A) 2011–2015, (B) 2016–2019 and (C) 2020–2025, 229 districts. Hot spots are outlined.",
     "Three maps of South Korea. In every period the highest ratios and outlined hot spots are in Busan and the south-east; the south-west is low in 2011–2015."),
}
SUPPINFO = ("Appendix S1 Supplementary methods and additional results. Table S1 Source, unit, observed years and expected direction of each district characteristic. "
            "Table S2 Likelihood comparison. Table S3 Rate ratios across alternative specifications. Table S4 Residual hot spots and cold spots. Appendix S2 STROBE checklist.")
def foot_row(J, text, sz=9, ls=13):
    from docx.enum.text import WD_LINE_SPACING
    t = J.doc.tables[-1]; r = t.add_row(); m = r.cells[0].merge(r.cells[-1]); m.text = ""; p = m.paragraphs[0]
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY; p.paragraph_format.line_spacing = Pt(ls); p.paragraph_format.space_before = Pt(4)
    c = J.counting; J.counting = False; J._emit(p, text, sz); J.counting = c
def put_table(J, t, fs=9):
    c = J.counting; J.counting = False; J.table_fit(t[0], t[1], t[2], "", fs_max=fs, title_bold=True); foot_row(J, t[3], 8); J.counting = c

def bullets(J, D, head):
    J.para(f"<b>{head}</b>", before=6)
    for s in re.split(r"(?<=\.)\s+(?=[A-Z])", D[head][0]): J.para(f"• {s}")

def build(bw=0, aw=0):
    D = read_draft(); J = TX(cite="sup")
    J.para(f"<b>{TITLE}</b>")
    J.para(f"<b>Running title:</b> {C.RUNNING}", before=6); J.para("<b>Article type:</b> Original Paper")
    J.para("Seongdae Kim", before=6); J.para(f"ORCID: https://orcid.org/{ORCID}"); J.para(AFFIL)
    J.para(f"<b>Institution where the research was conducted:</b> Hongik University, Seoul, Republic of Korea", before=6)
    J.para(f"<b>Corresponding author:</b> Seongdae Kim, {ADDRESS}. Telephone: {PHONE}. Email: kimsd@hongik.ac.kr", before=6)
    J.para(f"<b>Word count:</b> main text {bw}; summary {aw}. <b>Tables:</b> 2. <b>Figures:</b> 3. <b>References:</b> {len(KEYS)}.", before=6)
    J.para(f"<b>Conflict of interest disclosure:</b> {COI}", before=6); J.para(f"<b>Funding sources:</b> {FUNDING}")
    J.para(f"<b>Data availability:</b> {data_statement()}"); J.para(f"<b>Ethics approval statement:</b> {ETHICS}"); J.para(f"<b>Patient consent statement:</b> {CONSENT}")
    J.para(f"<b>Permission to reproduce material from other sources:</b> {PERMISSION}"); J.para(f"<b>Clinical trial registration:</b> {TRIAL}")
    J.para(f"<b>Keywords:</b> {'; '.join(KEYWORDS)}", before=6)
    J.page_break(); J.head("Summary", 1)
    for p in D["Abstract"]: J.para(p)
    J.page_break(); J.counting = True
    J.head("Introduction", 1)
    for p in D["Introduction"]: J.para(pandoc_cites(p), indent=True)
    J.head("Methods", 1)
    for x in C.METHODS:
        J.head(x[0], 2)
        for p in x[1:]: J.para(pandoc_cites(p), indent=True)
    J.head("Results", 1)
    for x in C.RESULTS:
        J.head(x[0], 2)
        for p in x[1:]: J.para(pandoc_cites(p), indent=True)
    J.head("Discussion", 1)
    for p in D["Discussion"]:
        if p.startswith("### "): J.head(p[4:].strip(), 2); continue
        J.para(pandoc_cites(p), indent=True)
    J.counting = False
    J.head("Acknowledgements", 1); J.para(ACK)
    J.head("Author contributions", 1); J.para(AUTHORSHIP)
    J.head(AI_HEAD, 1); J.para(AI_STATEMENT)
    J.head("Funding", 1); J.para("No funding was received for this study.")
    J.head("Conflict of interest", 1); J.para(COI)
    J.head("Data availability statement", 1); J.para(data_statement())
    J.head("Ethics approval statement", 1); J.para(ETHICS)
    assert sorted(J.cites) == KEYS, (set(J.cites) ^ set(KEYS))
    J.page_break(); J.head("References", 1); J.raw = True; rl = []
    for i, k in enumerate(J.cites, 1): s = f"{i}. " + ref(k); rl.append(s); J.para(s)
    J.raw = False
    J.head("Supporting information", 1); J.para(SUPPINFO)
    J.page_break(); J.head("Figure legends", 1)
    for k in sorted(FIGCAP): J.para(f"<b>Figure {k}</b> {FIGCAP[k][0]}"); J.para(f"Alt text: {FIGCAP[k][1]}", after=6)
    J.page_break(); put_table(J, table1())
    J.new_page("landscape"); put_table(J, table2(), fs=8)
    return J, rl, D

def docx_bw(path):   # 본문 단어 수: Introduction ~ Discussion 끝(Acknowledgments 직전)
    from docx import Document
    t = "\n".join(x.text for x in Document(path).paragraphs)
    i = t.find("\nIntroduction\n"); j = re.search(r"\nAcknowledge?ments\n", t).start(); assert 0 < i < j; return len(t[i:j].split())
def fix_orphans(path, heads, max_iter=8):
    from docx import Document
    fixed = set()
    for _ in range(max_iter):
        pdf = path.replace(".docx", ".pdf")
        subprocess.run(["/opt/homebrew/bin/soffice", "--headless", "-env:UserInstallation=file:///tmp/lo_bldd_bjh", "--convert-to", "pdf", "--outdir", os.path.dirname(path), path], check=True, capture_output=True)
        pages = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout.split("\f"); orphan = None
        for pg in pages:
            ls = [re.sub(r"^\d+\s+", "", l.strip()) for l in pg.split("\n") if l.strip()]; ls = [l for l in ls if not re.fullmatch(r"\d+", l)]
            if ls and ls[-1] in heads:
                k = len(ls) - 1
                while k > 0 and ls[k - 1] in heads: k -= 1
                orphan = ls[k]; break
        if orphan is None: return
        if orphan in fixed: print("orphan persists, stop:", orphan); return
        fixed.add(orphan); d = Document(path)
        for p in d.paragraphs:
            if p.text.strip() == orphan:
                np_ = OxmlElement("w:p"); r = OxmlElement("w:r"); br = OxmlElement("w:br"); br.set(qn("w:type"), "page"); r.append(br); np_.append(r); p._p.addprevious(np_); break
        d.save(path); print("orphan fixed:", orphan)

if __name__ == "__main__":
    p = os.path.join(OUT, "Manuscript_BJH_BLDD.docx"); D = read_draft()
    aw = sum(len(x.split()) for x in D["Abstract"])
    J, rl, D = build(); J.doc_title = "Manuscript"; J.save(p); bw = docx_bw(p)
    J, rl, D = build(bw, aw); J.doc_title = "Manuscript"; J.save(p); assert docx_bw(p) == bw
    heads = {"Summary", "Introduction", "Methods", "Results", "Discussion", "References", "Figure legends", "Acknowledgements", AI_HEAD, "Funding", "Conflict of interest", "Data availability statement", "Ethics approval statement", "Supporting information", "Author contributions"} | {x[0] for x in C.METHODS + C.RESULTS}
    fix_orphans(p, heads)
    open(os.path.join(OUT, "wordcount_bjh.txt"), "w").write(f"body_words={bw}\nsummary_words={aw}\nrefs={len(KEYS)}\ntables=2\nfigures=3\nrunning_title_chars={len(C.RUNNING)}\n")
    open(os.path.join(OUT, "reference_list.txt"), "w").write("\n".join(rl) + "\n")
    print(f"body {bw} | summary {aw} | refs {len(KEYS)}")
    assert bw <= 3000 and aw <= 200 and len(KEYS) <= 60 and len(C.RUNNING) <= 60
