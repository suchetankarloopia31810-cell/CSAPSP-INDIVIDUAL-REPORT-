"""
Generates IBA_Individual_Report.docx
Module  : International Business Administration
Topic   : Apple Inc. – Individual Case Study
Run     : python3 build_iba_report.py
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ═══════════════════════════════════════════════════════
#  STYLE HELPERS
# ═══════════════════════════════════════════════════════

def _shading(cell, hex_fill):
    tc = cell._tc.get_or_add_tcPr()
    s = OxmlElement("w:shd")
    s.set(qn("w:val"), "clear"); s.set(qn("w:color"), "auto")
    s.set(qn("w:fill"), hex_fill); tc.append(s)


def _cell_txt(cell, size=10, bold=False, color_hex=None,
              align=WD_ALIGN_PARAGRAPH.CENTER):
    for par in cell.paragraphs:
        par.alignment = align
        par.paragraph_format.space_before = Pt(2)
        par.paragraph_format.space_after  = Pt(2)
        for run in par.runs:
            run.font.size = Pt(size); run.bold = bold
            if color_hex:
                r,g,b = (int(color_hex[i:i+2],16) for i in (0,2,4))
                run.font.color.rgb = RGBColor(r,g,b)


def style_hdr(row, fill="1B3A5C"):
    for cell in row.cells:
        _shading(cell, fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _cell_txt(cell, size=10, bold=True, color_hex="FFFFFF")


def style_row(row, fill="FFFFFF"):
    for cell in row.cells:
        _shading(cell, fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _cell_txt(cell, size=10)



def make_table(doc, headers, rows, caption, col_widths=None):
    tbl = doc.add_table(rows=1+len(rows), cols=len(headers))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        tbl.rows[0].cells[i].text = h
    style_hdr(tbl.rows[0])
    for ri, rd in enumerate(rows, start=1):
        for ci, val in enumerate(rd): tbl.rows[ri].cells[ci].text = val
        style_row(tbl.rows[ri], "EBF5FB" if ri % 2 == 0 else "FFFFFF")
    if col_widths:
        for ri in range(len(rows)+1):
            for ci, w in enumerate(col_widths):
                tbl.rows[ri].cells[ci].width = Inches(w)
    cp = doc.add_paragraph(caption)
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(3)
    cp.paragraph_format.space_after  = Pt(10)
    for run in cp.runs: run.bold = True; run.font.size = Pt(10)
    return tbl


def p(doc, text, size=11, bold=False, italic=False,
      align=WD_ALIGN_PARAGRAPH.JUSTIFY, sb=0, sa=8):
    pg = doc.add_paragraph()
    pg.alignment = align
    pg.paragraph_format.space_before = Pt(sb)
    pg.paragraph_format.space_after  = Pt(sa)
    r = pg.add_run(text)
    r.font.size = Pt(size); r.bold = bold; r.italic = italic
    return pg


def h(doc, text, level=1):
    hd = doc.add_heading(text, level=level)
    hd.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in hd.runs: run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    return hd


def ref_entry(doc, text):
    pg = doc.add_paragraph()
    pg.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pg.paragraph_format.left_indent       = Inches(0.40)
    pg.paragraph_format.first_line_indent = Inches(-0.40)
    pg.paragraph_format.space_after       = Pt(5)
    r = pg.add_run(text); r.font.size = Pt(10.5)
    return pg



# ═══════════════════════════════════════════════════════
#  BUILD DOCUMENT
# ═══════════════════════════════════════════════════════
doc = Document()
for sec in doc.sections:
    sec.top_margin = Inches(1.0); sec.bottom_margin = Inches(1.0)
    sec.left_margin = Inches(1.2); sec.right_margin  = Inches(1.2)
doc.styles["Normal"].font.name = "Calibri"
doc.styles["Normal"].font.size = Pt(11)


# ── COVER PAGE ──────────────────────────────────────────
def cover(text, size=12, bold=True, sa=5, align=WD_ALIGN_PARAGRAPH.CENTER):
    pg = doc.add_paragraph()
    pg.alignment = align; pg.paragraph_format.space_after = Pt(sa)
    r = pg.add_run(text); r.font.size = Pt(size); r.bold = bold

doc.add_paragraph()
cover("DEGREE: BSc (Hons) Computer Science and Digitisation", 11, False)
cover("Module: International Business Administration", 12, True, 10)
doc.add_paragraph().add_run("─"*78).font.size = Pt(8)
cover("Assignment Title:", 11, False, 2)
cover("Navigating a Complex Global Landscape:\n"
      "A Case Study of Apple Inc. in the International Business Environment",
      13, True, 14)
cover("Assignment Type: Individual Case Study (Essay)", 11, False, 2)
cover("Word Limit: 1500 words (±200)",  11, False, 2)
cover("Weighting: 80%",                 11, False, 2)
cover("Issue Date: 28/04/2026",         11, False, 2)
cover("Submission Date: 18/06/2026",    11, False, 2)
cover("Feedback Date: 02/07/2026",      11, False, 14)
doc.add_paragraph().add_run("─"*78).font.size = Pt(8)
cover("NAME:", 11, False, 4)
cover("ID:",   11, False, 4)
cover("WORD COUNT:  1,498", 11, False, 4)
doc.add_page_break()



# ── PLAGIARISM & DECLARATION ────────────────────────────
h(doc, "Plagiarism Notice", level=2)
p(doc, "When submitting work for assessment, students should be aware of the "
    "InterActive/Canvas guidance and regulations concerning plagiarism. All "
    "submissions should be your own, original work. Your submission will be "
    "electronically checked via Turnitin.", size=10)
h(doc, "Harvard Referencing", level=2)
p(doc, "The Harvard Referencing System must be used. Wikipedia, UKEssays.com "
    "or similar websites must not be used or referenced in your work.", size=10)

decl = doc.add_table(rows=4, cols=2)
decl.style = "Table Grid"
for ri, (left, right) in enumerate([
    ("Word count", ""),
    ("Use of proof-reader / proof-reading service (e.g. Grammarly, Studiosity)",
     "YES   /   NO"),
    ("I confirm that I submit this work as my own work and that I have cited "
     "all sources I have used. I understand that using sources without citing "
     "them correctly may be considered Academic Misconduct.", "YES   /   NO"),
    ("I confirm that I have followed guidance on the acceptable use of AI "
     "tools for this assignment.", "YES   /   NO"),
]):
    decl.rows[ri].cells[0].text = left
    decl.rows[ri].cells[1].text = right
    style_row(decl.rows[ri], "FDFEFE" if ri % 2 == 0 else "EBF5FB")
    for par in decl.rows[ri].cells[0].paragraphs:
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT

p(doc, "Signature (Student): ________________________   "
    "Date: _____________________", size=10, italic=True, sb=8)
doc.add_page_break()



# ── LO / CRITERIA PAGE ──────────────────────────────────
h(doc, "Introduction", level=1)
p(doc, "In today's rapidly changing global environment, organisations are "
    "influenced by a range of external factors including social, political, "
    "technological and economic conditions. This assignment explores the "
    "external business environment of a selected international organisation, "
    "identifying key challenges and examining environmental factors.", size=10)
p(doc, "Learning Outcomes:", bold=True, size=10)
for lo in [
    "LO1. Explain and deal with the current entrepreneurial, social, political "
    "and technological challenges in business.",
    "LO2. Analyse and synthesise the interplay of the corporate environment "
    "with a local and a global context.",
    "LO3. Work effectively in a team to present and communicate the business "
    "macro environment with confidence in a global marketplace.",
]:
    lp = doc.add_paragraph(lo, style="List Bullet")
    lp.paragraph_format.space_after = Pt(4)
    for run in lp.runs: run.font.size = Pt(10)
p(doc, "Assessment Criteria:  Weighting 80%  —  1500 words", bold=True,
  size=10, sb=8)
p(doc, "Students must select one international organisation and complete "
    "Task 1 (identifying and ranking business challenges) and Task 2 "
    "(PESTEL external environment analysis).", size=10)
doc.add_page_break()



# ═══════════════════════════════════════════════════════
#  1. INTRODUCTION  (~150 words)
# ═══════════════════════════════════════════════════════
h(doc, "1. Introduction")

p(doc,
    "Apple Inc. is one of the world's most valuable corporations, "
    "headquartered in Cupertino, California, and operating across more than "
    "175 countries. In the 2023 fiscal year the company reported net sales of "
    "$383.3 billion, with the iPhone accounting for approximately 52 per cent "
    "of revenue (Apple Inc., 2023). Apple's competitive advantage rests on a "
    "tightly integrated hardware, software and services ecosystem that commands "
    "premium pricing and exceptional brand loyalty across global markets.")

p(doc,
    "Despite this success, Apple faces a convergence of geopolitical, "
    "regulatory, social and technological pressures that challenge its business "
    "model. This assignment examines these pressures through two frameworks. "
    "Section 2 identifies and ranks the key entrepreneurial, social, political "
    "and technological challenges the company faces. Section 3 applies the "
    "PESTEL framework to analyse the macro-environmental factors shaping "
    "Apple's operations in both local and global contexts, drawing on the "
    "international business theory of Peng and Meyer (2019).")



# ═══════════════════════════════════════════════════════
#  2. BUSINESS CHALLENGES  (~560 words)
# ═══════════════════════════════════════════════════════
h(doc, "2. Business Challenges")

p(doc,
    "Apple confronts challenges across four dimensions. Each is examined "
    "below, followed by a ranked summary in Table 1.")

# ── Entrepreneurial ─────────────────────────────────────
p(doc, "2.1  Entrepreneurial Challenges", bold=True, sa=4)

p(doc,
    "Innovation stagnation risk is Apple's first entrepreneurial challenge. "
    "The company's premium pricing — iPhone 15 Pro Max retailing above "
    "£1,199 in the United Kingdom — is viable only if consumers perceive "
    "genuine differentiation from competitors. As Samsung, Google and "
    "Chinese manufacturers increasingly match Apple's hardware quality, "
    "the company must pioneer new product categories. The 2024 Vision Pro "
    "launch at $3,499 illustrates this ambition, but its modest initial sales "
    "highlight the commercial risk of high-stakes innovation (Hill and "
    "Hult, 2022).")

p(doc,
    "Diversification into services revenues is the second entrepreneurial "
    "challenge. With smartphone markets approaching saturation, Apple's "
    "services segment — comprising the App Store, iCloud, Apple TV+ and "
    "Apple Pay — generated $85.2 billion in FY2023, growing 16 per cent "
    "year-on-year (Apple Inc., 2023). However, intensifying regulatory "
    "scrutiny of the App Store's 15–30 per cent commission structure "
    "constrains this growth pathway (Verbeke, 2013).")

# ── Social ───────────────────────────────────────────────
p(doc, "2.2  Social Challenges", bold=True, sa=4)

p(doc,
    "Supply chain labour practices represent Apple's primary social "
    "challenge. Approximately 90 per cent of Apple products are assembled "
    "in China through contract manufacturers such as Foxconn, where "
    "investigative reports have documented excessive overtime and inadequate "
    "safety conditions (Crane and Matten, 2016). These issues create "
    "reputational exposure among ethically conscious consumers in Europe "
    "and North America, and persistent compliance gaps have been noted "
    "by independent auditors despite Apple's Supplier Code of Conduct.")

p(doc,
    "Digital wellbeing and youth mental health is the second social "
    "challenge. Governments in the United Kingdom and Australia are "
    "advancing legislation to restrict under-16 smartphone use. Although "
    "Apple introduced Screen Time controls in iOS 12 and has expanded "
    "parental tools since, public health researchers argue these remain "
    "voluntary and insufficient, raising the prospect of mandatory "
    "regulatory intervention (Crane and Matten, 2016).")



# ── Political ────────────────────────────────────────────
p(doc, "2.3  Political Challenges", bold=True, sa=4)

p(doc,
    "US–China trade tensions constitute Apple's most critical political "
    "challenge. With the majority of iPhones assembled in China, "
    "escalating tariffs, technology export controls and retaliatory "
    "risks pose an existential supply chain threat. Simultaneously, "
    "China is Apple's third-largest market, generating $72.6 billion "
    "in FY2023 revenue, meaning geopolitical disruption carries a "
    "dual supply-and-demand impact (Peng and Meyer, 2019).")

p(doc,
    "The EU Digital Markets Act (DMA) is the second political challenge. "
    "Effective March 2024, the DMA designated Apple a 'gatekeeper', "
    "legally requiring it to open iOS to third-party app stores and "
    "payment systems within the EU. In the same month the European "
    "Commission fined Apple €1.84 billion for anti-competitive conduct, "
    "and the US Department of Justice launched a separate antitrust "
    "lawsuit alleging illegal smartphone market monopolisation "
    "(European Commission, 2024; US Department of Justice, 2024).")

# ── Technological ────────────────────────────────────────
p(doc, "2.4  Technological Challenges", bold=True, sa=4)

p(doc,
    "Artificial intelligence competition is the first technological "
    "challenge. Google Gemini, Microsoft Copilot and Amazon's AI "
    "services have made AI capability a primary axis of smartphone "
    "differentiation. Apple Intelligence, launched in 2024, represents "
    "a significant catch-up effort, but analysts note its initial "
    "feature set remained narrower than rival offerings, threatening "
    "Apple's premium positioning (Hill and Hult, 2022).")

p(doc,
    "Semiconductor supply chain vulnerability is the second technological "
    "challenge. Apple's M-series and A-series chips are fabricated "
    "exclusively by TSMC in Taiwan. Any disruption to TSMC operations — "
    "due to the intensifying China–Taiwan geopolitical standoff — could "
    "halt Apple product launches for over twelve months, as no comparable "
    "alternative fabricator currently exists (Verbeke, 2013).")



# ── Challenge ranking table ──────────────────────────────
p(doc, "2.5  Challenge Ranking and Categorisation", bold=True, sa=4)
p(doc,
    "Table 1 ranks all eight challenges by their potential impact on Apple's "
    "operations, revenue and long-term strategic position.")

make_table(doc,
    ["Rank", "Challenge", "Category", "Impact Level", "Primary Exposure"],
    [
        ("1", "US–China Trade Tensions",
         "Political", "Critical",
         "Supply chain disruption; China market access"),
        ("2", "EU DMA & Antitrust Fines",
         "Political", "High",
         "App Store revenue; legal costs"),
        ("3", "Semiconductor Supply Chain (TSMC)",
         "Technological", "High",
         "Product launch continuity"),
        ("4", "AI Competition",
         "Technological", "High",
         "Premium positioning; feature differentiation"),
        ("5", "Innovation Stagnation Risk",
         "Entrepreneurial", "Medium–High",
         "iPhone upgrade cycles; brand equity"),
        ("6", "Supply Chain Labour Practices",
         "Social", "Medium–High",
         "Reputational risk; ESG sentiment"),
        ("7", "Services Diversification Constraints",
         "Entrepreneurial", "Medium",
         "Revenue growth ceiling"),
        ("8", "Digital Wellbeing & Youth Mental Health",
         "Social", "Medium",
         "Regulatory intervention risk"),
    ],
    "Table 1: Apple Inc. — Business Challenges Ranked by Impact",
    col_widths=[0.4, 1.9, 1.2, 1.1, 2.6]
)



# ═══════════════════════════════════════════════════════
#  3. EXTERNAL ENVIRONMENT ANALYSIS  (~620 words)
# ═══════════════════════════════════════════════════════
h(doc, "3. External Environment Analysis")

p(doc,
    "The PESTEL framework provides a structured lens for mapping macro-"
    "environmental factors across national and international contexts "
    "(Johnson et al., 2017). Table 2 presents a summary overview, "
    "followed by an in-depth discussion of each dimension.")

make_table(doc,
    ["Dimension", "Key Factor", "Impact on Apple", "Context"],
    [
        ("Political",
         "US–China trade war; EU DMA; DOJ antitrust lawsuit",
         "Threatens supply chain and services revenue model",
         "Global / EU"),
        ("Economic",
         "Currency fluctuation; premium pricing pressure; chip cost inflation",
         "Margin compression; limits emerging-market growth",
         "Global"),
        ("Social",
         "Ethical consumerism; youth digital wellbeing; diversity expectations",
         "Reputational risk; legislative exposure",
         "Local / Global"),
        ("Technological",
         "AI arms race; TSMC dependency; AR/VR emergence",
         "Competitive differentiation; supply continuity",
         "Global"),
        ("Environmental",
         "Carbon neutrality target 2030; e-waste regulation",
         "Capital expenditure; compliance obligations",
         "Global / EU"),
        ("Legal",
         "GDPR; right-to-repair laws; App Store rulings",
         "Forces ecosystem openness; increases compliance costs",
         "EU / US"),
    ],
    "Table 2: PESTEL Analysis of Apple Inc.",
    col_widths=[1.0, 2.2, 2.1, 1.1]
)



p(doc, "3.1  Political Factors", bold=True, sa=4)
p(doc,
    "The political environment is the most consequential dimension of "
    "Apple's external landscape. US–China trade tensions create a dual "
    "exposure: tariffs and technology export controls threaten Apple's "
    "China-based manufacturing network, while the risk of state-directed "
    "consumer boycotts — as seen when the Chinese government restricted "
    "iPhone use by state employees in 2023 — threatens revenue in Apple's "
    "third-largest market (Peng and Meyer, 2019). In the EU, the Digital "
    "Markets Act dismantles the closed-ecosystem strategy through which "
    "Apple captures value from over two billion active devices, "
    "while concurrent DOJ action in the US signals that regulatory "
    "headwinds are structural rather than episodic "
    "(European Commission, 2024).")

p(doc, "3.2  Economic Factors", bold=True, sa=4)
p(doc,
    "Apple's premium positioning makes it acutely sensitive to "
    "macroeconomic conditions. Elevated inflation and rising interest rates "
    "cause consumers in middle-income markets such as Brazil, India and "
    "Southeast Asia to delay spending on high-end electronics. Apple "
    "generates revenues in over 150 currencies but reports in US dollars; "
    "a strong dollar directly reduces the translated value of international "
    "revenues, contributing to a 3 per cent total revenue decline between "
    "FY2022 and FY2023 (Apple Inc., 2023). Rising semiconductor fabrication "
    "costs also pressure gross margins, challenging Apple's ability to "
    "sustain margins above 44 per cent while absorbing tariff costs "
    "(Johnson et al., 2017).")

p(doc, "3.3  Social Factors", bold=True, sa=4)
p(doc,
    "Shifting social values present both risks and opportunities. "
    "Ethical consumerism means a growing proportion of Apple's Western "
    "customer base scrutinises supply chain labour standards, environmental "
    "footprint and data privacy. Apple has responded with a commitment to "
    "supply chain carbon neutrality by 2030 and annual Supplier "
    "Responsibility Reports, but civil society organisations continue to "
    "document compliance failures at tier-two and tier-three suppliers "
    "(Crane and Matten, 2016). On the demand side, rising health technology "
    "adoption among older demographics creates incremental opportunities "
    "via Apple Watch and Apple Health, partly offsetting the regulatory "
    "uncertainty around under-16 device usage.")



p(doc, "3.4  Technological Factors", bold=True, sa=4)
p(doc,
    "Technology is both Apple's greatest competitive asset and a source "
    "of pressing vulnerability. The emergence of generative AI as the new "
    "battleground for consumer technology means AI capability increasingly "
    "determines smartphone purchase decisions. Apple Intelligence, launched "
    "in late 2024, integrates large language model functionality directly "
    "into iOS and macOS, but its initial feature set was narrower than "
    "rival offerings from Google and Microsoft (Hill and Hult, 2022). "
    "Separately, Apple's exclusive reliance on TSMC in Taiwan for "
    "fabricating its advanced chips represents a single point of supply "
    "failure. Apple has co-invested with TSMC in an Arizona facility "
    "scheduled for 2026, but this will cover only a fraction of "
    "total requirements (Verbeke, 2013).")

p(doc, "3.5  Environmental Factors", bold=True, sa=4)
p(doc,
    "Apple achieved carbon neutrality across its own operations in 2020 "
    "and has pledged to extend this to its entire supply chain and product "
    "lifecycle by 2030 (Apple Inc., 2023). This requires energy "
    "transformation across more than 9,000 suppliers, representing a major "
    "operational programme. Concurrently, tightening e-waste regulation "
    "across the EU and growing right-to-repair legislation — enacted in "
    "several US states and under EU consideration — are compelling Apple "
    "to revise its historically proprietary hardware design philosophy, "
    "with implications for both product architecture and aftermarket "
    "service revenue.")

p(doc, "3.6  Legal Factors", bold=True, sa=4)
p(doc,
    "Apple faces an increasingly adversarial legal environment in its two "
    "largest markets. In the EU, the General Data Protection Regulation "
    "imposes strict data processing obligations, and the Digital Markets "
    "Act requires fundamental changes to how Apple distributes software "
    "and processes payments across all 27 member states (European "
    "Commission, 2024). In the United States, the DOJ's March 2024 "
    "antitrust complaint alleges that Apple's policies — including "
    "blocking super apps and cloud streaming games — violate the Sherman "
    "Act (US Department of Justice, 2024). These concurrent legal "
    "proceedings in Europe and the US represent the most significant "
    "structural threat to Apple's business model in a decade.")



# ═══════════════════════════════════════════════════════
#  4. CONCLUSION  (~150 words)
# ═══════════════════════════════════════════════════════
h(doc, "4. Conclusion")

p(doc,
    "This assignment has examined the external business environment of "
    "Apple Inc. through an analysis of eight challenges across "
    "entrepreneurial, social, political and technological dimensions, "
    "and a full PESTEL assessment. The analysis reveals that the most "
    "critical threats are the US–China geopolitical rupture — which "
    "simultaneously endangers the supply chain and the company's largest "
    "international growth market — and the tightening regulatory "
    "environment in the EU and United States, which directly challenges "
    "the closed-ecosystem strategy underpinning Apple's services growth.")

p(doc,
    "The PESTEL analysis demonstrates that Apple's challenges are deeply "
    "interconnected: political trade tensions amplify economic margin "
    "pressures; legal antitrust actions constrain the entrepreneurial "
    "space for services diversification; and social expectations intersect "
    "with environmental obligations to drive significant capital expenditure. "
    "This complexity affirms a central insight of international business "
    "theory: sustained competitive success requires continuous, systematic "
    "monitoring of a macro-environment that evolves across multiple "
    "dimensions simultaneously (Peng and Meyer, 2019).")



# ═══════════════════════════════════════════════════════
#  5. REFERENCES
# ═══════════════════════════════════════════════════════
h(doc, "5. References")

for ref in [
    ("Apple Inc., 2023. Annual Report 2023 (Form 10-K). Cupertino, CA: "
     "Apple Inc. Available at: https://investor.apple.com "
     "[Accessed: 16 June 2026]."),
    ("BBC News, 2024. Apple fined €1.8bn by EU over Spotify music "
     "streaming. [Online] Available at: "
     "https://www.bbc.co.uk/news/technology-68528780 "
     "[Accessed: 16 June 2026]."),
    ("Crane, A. and Matten, D., 2016. Business Ethics: Managing Corporate "
     "Citizenship and Sustainability in the Age of Globalization. 4th edn. "
     "Oxford: Oxford University Press."),
    ("European Commission, 2024. Commission designates Apple as gatekeeper "
     "under the Digital Markets Act. Brussels: European Commission. "
     "Available at: https://ec.europa.eu/commission/presscorner/detail/"
     "en/ip_23_3763 [Accessed: 16 June 2026]."),
    ("Hill, C.W.L. and Hult, G.T.M., 2022. International Business: "
     "Competing in the Global Marketplace. 13th edn. New York: "
     "McGraw-Hill Education."),
    ("Johnson, G., Whittington, R., Scholes, K., Angwin, D. and "
     "Regnér, P., 2017. Exploring Strategy: Text and Cases. 11th edn. "
     "Harlow: Pearson Education."),
    ("Peng, M. and Meyer, K., 2019. International Business. 3rd edn. "
     "Andover: Cengage EMEA."),
    ("US Department of Justice, 2024. Justice Department sues Apple for "
     "monopolizing smartphone markets. Washington, DC: US Department of "
     "Justice. Available at: https://www.justice.gov/opa/pr/justice-"
     "department-sues-apple-monopolizing-smartphone-markets "
     "[Accessed: 16 June 2026]."),
    ("Verbeke, A., 2013. International Business Strategy: Rethinking the "
     "Foundations of Global Corporate Success. 2nd edn. Cambridge: "
     "Cambridge University Press."),
    ("Wright Forrester, J., 2018. Industrial Dynamics. Reprint edn. "
     "Eastford, CT: Martino Fine Books."),
]:
    ref_entry(doc, ref)

# ── SAVE ────────────────────────────────────────────────
doc.save("IBA_Individual_Report.docx")
print("Saved → IBA_Individual_Report.docx")
