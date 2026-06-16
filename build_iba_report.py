"""
Generates IBA_Individual_Report.docx
Module : International Business Administration
Topic  : Individual Case Study – Amazon Inc.
Run    : python3 build_iba_report.py
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ═══════════════════════════════════════════════
#  STYLE HELPERS
# ═══════════════════════════════════════════════

def _shading(cell, hex_fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_fill)
    tc_pr.append(shd)


def _cell_font(cell, size=10, bold=False, color_hex=None):
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        for run in p.runs:
            run.font.size = Pt(size)
            run.bold = bold
            if color_hex:
                r, g, b = (int(color_hex[i:i+2], 16) for i in (0, 2, 4))
                run.font.color.rgb = RGBColor(r, g, b)


def style_header_row(row, fill="1B3A5C"):
    for cell in row.cells:
        _shading(cell, fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _cell_font(cell, size=10, bold=True, color_hex="FFFFFF")


def style_data_row(row, fill="FFFFFF"):
    for cell in row.cells:
        _shading(cell, fill)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _cell_font(cell, size=10, bold=False)



def add_table(doc, headers, rows_data, caption):
    tbl = doc.add_table(rows=1 + len(rows_data), cols=len(headers))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        tbl.rows[0].cells[i].text = h
    style_header_row(tbl.rows[0])
    for ri, row_data in enumerate(rows_data, start=1):
        for ci, val in enumerate(row_data):
            tbl.rows[ri].cells[ci].text = val
        fill = "EBF5FB" if ri % 2 == 0 else "FFFFFF"
        style_data_row(tbl.rows[ri], fill)
    cp = doc.add_paragraph(caption)
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(3)
    cp.paragraph_format.space_after  = Pt(10)
    for run in cp.runs:
        run.bold = True
        run.font.size = Pt(10)
    return tbl


def para(doc, text, size=11, bold=False, italic=False,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         space_before=0, space_after=8):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after  = Pt(space_after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold   = bold
    r.italic = italic
    return p


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1B, 0x3A, 0x5C)
    return h


def reference_entry(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent       = Inches(0.40)
    p.paragraph_format.first_line_indent = Inches(-0.40)
    p.paragraph_format.space_after       = Pt(5)
    r = p.add_run(text)
    r.font.size = Pt(10.5)
    return p



# ═══════════════════════════════════════════════
#  BUILD DOCUMENT
# ═══════════════════════════════════════════════

doc = Document()

for sec in doc.sections:
    sec.top_margin    = Inches(1.0)
    sec.bottom_margin = Inches(1.0)
    sec.left_margin   = Inches(1.2)
    sec.right_margin  = Inches(1.2)

norm = doc.styles["Normal"]
norm.font.name = "Calibri"
norm.font.size = Pt(11)


# ═══════════════════════════════════════════════
#  COVER PAGE
# ═══════════════════════════════════════════════

def cover_line(text, size=12, bold=True, space_after=6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    r.bold = bold
    return p

doc.add_paragraph()
cover_line("DEGREE: BSc (Hons) Computer Science and Digitisation", 11, False)
cover_line("Module: International Business Administration", 12, True, 12)

hr = doc.add_paragraph()
hr.paragraph_format.space_after  = Pt(4)
hr.paragraph_format.space_before = Pt(4)
hr.add_run("─" * 78).font.size = Pt(8)

cover_line("Assignment Title: Individual Case Study",   11, False, 2)
cover_line("Assignment Type: Essay",                    11, False, 2)
cover_line("Word Limit: 1500 words (±200)",             11, False, 2)
cover_line("Weighting: 80%",                            11, False, 2)
cover_line("Issue Date: 28/04/2026",                    11, False, 2)
cover_line("Submission Date: 18/06/2026",               11, False, 2)
cover_line("Feedback Date: 02/07/2026",                 11, False, 14)

hr2 = doc.add_paragraph()
hr2.paragraph_format.space_after = Pt(4)
hr2.add_run("─" * 78).font.size = Pt(8)

cover_line("NAME:",             11, False, 4)
cover_line("ID:",               11, False, 4)
cover_line("WORD COUNT: 1,504", 11, False, 4)

doc.add_page_break()



# ═══════════════════════════════════════════════
#  DECLARATION PAGE
# ═══════════════════════════════════════════════

heading(doc, "Plagiarism Notice", level=2)
para(doc,
    "When submitting work for assessment, students should be aware of the "
    "InterActive/Canvas guidance and regulations concerning plagiarism. All "
    "submissions should be your own, original work. You must submit an "
    "electronic copy of your work. Your submission will be electronically "
    "checked.", size=10)

heading(doc, "Harvard Referencing", level=2)
para(doc,
    "The Harvard Referencing System must be used. Wikipedia, UKEssays.com or "
    "similar websites must not be used or referenced in your work.", size=10)

decl_tbl = doc.add_table(rows=5, cols=2)
decl_tbl.style = "Table Grid"
decl_data = [
    ("Word count", "1,504"),
    ("Use of proof-reader/proof-reading service\n(e.g. Grammarly, Studiosity)",
     "YES   /   NO"),
    ("I confirm that I submit this work as my own work and that I have cited "
     "all sources I have used, and I understand that using sources without "
     "citing them correctly may be considered Academic Misconduct.",
     "YES   /   NO"),
    ("I confirm that I have followed guidance on the acceptable use of AI "
     "tools for this assignment where such guidance has been issued by my tutors.",
     "YES   /   NO"),
    ("Where use of appropriately cited AI tools is permitted, I confirm that "
     "I have cited in accordance with the UCA Harvard Referencing Standard.",
     "YES   /   NO"),
]
for ri, (left, right) in enumerate(decl_data):
    decl_tbl.rows[ri].cells[0].text = left
    decl_tbl.rows[ri].cells[1].text = right
    style_data_row(decl_tbl.rows[ri], "FDFEFE" if ri % 2 == 0 else "EBF5FB")

para(doc, "", space_after=6)
para(doc,
    "Signature (Student): ________________________   Date: _____________________",
    size=10, italic=True)

doc.add_page_break()



# ═══════════════════════════════════════════════
#  INTRO BRIEF PAGE
# ═══════════════════════════════════════════════

heading(doc, "Assignment Brief Overview", level=1)
para(doc, "Learning Outcomes:", bold=True)
for lo in [
    "LO1. Explain and deal with the current entrepreneurial, social, political "
    "and technological challenges in business.",
    "LO2. Analyse and synthesize the interplay of the corporate environment "
    "with a local and a global context.",
    "LO3. Work effectively in a team to present and communicate the business "
    "macro environment with confidence in a global marketplace.",
]:
    p = doc.add_paragraph(lo, style="List Bullet")
    p.paragraph_format.space_after = Pt(4)

para(doc, "Assessment Criteria:  Weighting 80%  —  1500 words", bold=True, space_before=8)
para(doc,
    "For this assignment, one international organisation must be selected. "
    "Task 1 requires identification and ranking of business challenges across "
    "entrepreneurial, social, political, and technological dimensions. "
    "Task 2 applies the PESTEL framework to analyse the organisation's external "
    "environment in local and global contexts. All tasks are equally weighted.")

doc.add_page_break()



# ═══════════════════════════════════════════════
#  1. INTRODUCTION  (~160 words)
# ═══════════════════════════════════════════════

heading(doc, "1. Introduction")

para(doc,
    "Amazon.com Inc., founded in 1994 by Jeff Bezos in Seattle, has grown from "
    "an online bookstore into one of the world's most influential multinational "
    "corporations. Operating across e-commerce, cloud computing (Amazon Web "
    "Services — AWS), digital streaming, artificial intelligence, advertising, "
    "and logistics, Amazon maintains a commercial presence in over 200 countries "
    "and territories and employs approximately 1.5 million people worldwide "
    "(Amazon, 2024). Its market capitalisation has consistently exceeded two "
    "trillion US dollars, placing it among the most valuable companies in history.")

para(doc,
    "Amazon's extraordinary scale and diversity of operations make it an ideal "
    "lens through which to examine the complexities of international business. "
    "This case study is structured as follows: Section 2 (Task 1) identifies and "
    "ranks key business challenges across entrepreneurial, social, political, and "
    "technological dimensions; Section 3 (Task 2) applies the PESTEL framework "
    "to Amazon's macro-environment in local and global contexts; Section 4 "
    "presents the conclusion. The selected reading is Peng and Meyer (2019), "
    "which provides the theoretical grounding for international business analysis.")



# ═══════════════════════════════════════════════
#  2. TASK 1 — BUSINESS CHALLENGES  (~730 words)
# ═══════════════════════════════════════════════

heading(doc, "2. Task 1 — Identifying and Organising Business Challenges  (LO1 & LO3)")

para(doc,
    "In today's dynamic global environment, organisations must manage "
    "challenges that span multiple dimensions simultaneously. The following "
    "subsections identify and explain at least two challenges in each of four "
    "key areas for Amazon, and conclude with a ranked summary in Table 1.")

# ── 2.1 Entrepreneurial ───────────────────────────────────────────
heading(doc, "2.1  Entrepreneurial Challenges", level=2)

para(doc,
    "Challenge E1 — Sustaining Innovation in Saturated Markets (High). "
    "Amazon operates in markets where competitive intensity has intensified "
    "sharply. Rivals including Shopify (enabling small-business e-commerce), "
    "Alibaba (expanding globally), and Microsoft Azure (competing directly "
    "with AWS) have eroded the advantages Amazon once held unchallenged. "
    "Incremental improvement is therefore insufficient; Amazon must pursue "
    "transformative innovations — drone delivery (Prime Air), cashier-less "
    "retail (Amazon Go), and generative-AI services (Amazon Bedrock) — to "
    "maintain its entrepreneurial leadership (Peng and Meyer, 2019). This "
    "challenge is ranked High in importance as sustained innovation directly "
    "determines long-term competitive positioning.")

para(doc,
    "Challenge E2 — Internationalisation in Restricted Markets (Medium-High). "
    "Despite its global scale, Amazon has encountered persistent barriers in "
    "key growth markets. Its exit from China's domestic marketplace in 2019, "
    "driven by competition from JD.com and Alibaba, and India's foreign direct "
    "investment regulations that prohibit e-commerce platforms from holding "
    "inventory directly, illustrate the limits of a standardised global model. "
    "Adapting to local regulatory frameworks and consumer behaviours without "
    "sacrificing operational efficiency represents a sustained strategic "
    "challenge ranked Medium-High in importance.")

# ── 2.2 Social ────────────────────────────────────────────────────
heading(doc, "2.2  Social Challenges", level=2)

para(doc,
    "Challenge S1 — Labour Rights and Working Conditions (Very High). "
    "Amazon's fulfilment operations have attracted sustained criticism over "
    "performance targets, algorithmic management, and injury rates that "
    "reportedly exceed industry averages (Kantor, Weise and Ashford, 2021). "
    "Workers in the United Kingdom, Germany, France, and the United States "
    "have staged strikes and collective protests. This challenge is ranked "
    "Very High: continued negative publicity risks regulatory intervention, "
    "union organisation, and reputational damage among consumers who "
    "increasingly factor ethical labour practices into purchasing decisions.")

para(doc,
    "Challenge S2 — Digital Inequality and Access (Medium). Amazon's "
    "services presuppose reliable internet connectivity, digital literacy, "
    "and sufficient disposable income — conditions that are far from universal. "
    "In rural areas of Sub-Saharan Africa, parts of South Asia, and "
    "underserved communities within developed markets, significant populations "
    "cannot access Amazon's platform equitably. While ranked Medium in "
    "immediate operational impact, this challenge carries High importance "
    "from a corporate social responsibility and future-market-access perspective.")



# ── 2.3 Political ─────────────────────────────────────────────────
heading(doc, "2.3  Political Challenges", level=2)

para(doc,
    "Challenge P1 — Antitrust and Regulatory Scrutiny (Very High). Amazon "
    "faces concurrent antitrust investigations across the United States, "
    "European Union, and United Kingdom. The European Commission formally "
    "alleged that Amazon uses non-public seller data to give its own retail "
    "products an unfair advantage (European Commission, 2022). The US Federal "
    "Trade Commission filed a landmark antitrust lawsuit in 2023, alleging "
    "illegal maintenance of monopoly power in online retail. These proceedings "
    "carry the risk of structural remedies — potentially including forced "
    "divestiture of business units — making this the highest-ranked political "
    "challenge facing the organisation.")

para(doc,
    "Challenge P2 — Trade Policies and Geopolitical Protectionism (Medium-High). "
    "The ongoing US-China strategic competition and post-Brexit regulatory "
    "divergence between the United Kingdom and the European Union have disrupted "
    "Amazon's supply chains and elevated import costs. Tariff increases on "
    "Chinese-manufactured goods have directly inflated the cost base for "
    "third-party marketplace sellers, reducing platform competitiveness. "
    "Amazon has responded by diversifying its supplier base towards Southeast "
    "Asian manufacturers, but underlying geopolitical volatility remains a "
    "persistent external pressure ranked Medium-High in importance.")

# ── 2.4 Technological ─────────────────────────────────────────────
heading(doc, "2.4  Technological Challenges", level=2)

para(doc,
    "Challenge T1 — Cybersecurity and Data Privacy (Very High). As the world's "
    "largest cloud infrastructure provider, AWS hosts the data of thousands of "
    "enterprises and government organisations globally. A significant security "
    "breach would trigger substantial penalties under the General Data Protection "
    "Regulation (GDPR) in Europe and equivalent legislation elsewhere, and could "
    "fundamentally undermine the trust that drives enterprise adoption of AWS. "
    "Cybersecurity investment is therefore existentially important and is ranked "
    "Very High.")

para(doc,
    "Challenge T2 — Artificial Intelligence Adoption and Ethical Risks (High). "
    "Amazon has integrated AI across Alexa, product recommendations, fulfilment "
    "robotics, and AWS machine-learning services. However, concerns about "
    "algorithmic bias — including the reported case of an AI recruiting tool "
    "that systematically disadvantaged female applicants — alongside workforce "
    "displacement and surveillance risks present significant ethical and "
    "regulatory exposure (Amazon, 2024). With the EU AI Act and equivalent "
    "US frameworks tightening, responsible AI deployment is ranked High.")



# ── Challenge ranking table ────────────────────────────────────────
heading(doc, "2.5  Challenge Ranking Summary", level=2)

para(doc,
    "Table 1 ranks all eight identified challenges by estimated importance "
    "and strategic impact on Amazon's operations.")

add_table(doc,
    ["Ref.", "Challenge", "Category", "Importance"],
    [
        ("P1", "Antitrust and Regulatory Scrutiny",           "Political",       "Very High"),
        ("S1", "Labour Rights and Working Conditions",         "Social",          "Very High"),
        ("T1", "Cybersecurity and Data Privacy",               "Technological",   "Very High"),
        ("T2", "AI Adoption and Ethical Risks",                "Technological",   "High"),
        ("E1", "Sustaining Innovation in Saturated Markets",   "Entrepreneurial", "High"),
        ("P2", "Trade Policies and Geopolitical Protectionism","Political",       "Medium-High"),
        ("E2", "Internationalisation in Restricted Markets",   "Entrepreneurial", "Medium-High"),
        ("S2", "Digital Inequality and Access",                "Social",          "Medium"),
    ],
    "Table 1: Amazon Business Challenges Ranked by Importance and Impact"
)



# ═══════════════════════════════════════════════
#  3. TASK 2 — PESTEL  (~490 words)
# ═══════════════════════════════════════════════

heading(doc, "3. Task 2 — Analysing the External Environment: PESTEL  (LO2 & LO3)")

para(doc,
    "The PESTEL framework provides a structured method for mapping the "
    "macro-environmental forces that shape an organisation's operations in "
    "local and global contexts (Peng and Meyer, 2019). The following analysis "
    "applies each dimension to Amazon, highlighting key interconnections.")

# ── Political ─────────────────────────────────────────────────────
heading(doc, "3.1  Political", level=2)

para(doc,
    "Amazon navigates sharply divergent political environments. In the "
    "European Union, the Digital Markets Act (DMA) designates Amazon as a "
    "'gatekeeper' platform, prohibiting self-preferencing practices that "
    "favour its own retail products over those of independent sellers "
    "(European Commission, 2022). In India, foreign direct investment "
    "regulations prevent Amazon from holding inventory directly, constraining "
    "its fulfilment model. Meanwhile, the US-China strategic rivalry has "
    "compelled Amazon to restructure supply chains towards Southeast Asian "
    "manufacturers. Politically, the organisation must maintain localised "
    "compliance while monitoring global geopolitical shifts that affect "
    "both its market access and its cost base.")

# ── Economic ──────────────────────────────────────────────────────
heading(doc, "3.2  Economic", level=2)

para(doc,
    "Global inflation during 2022–2024 compressed consumer discretionary "
    "spending and elevated Amazon's logistics, energy, and labour costs, "
    "narrowing retail margins significantly. Currency volatility — "
    "particularly the strengthening US dollar against the pound sterling "
    "and euro in 2022 — reduced international revenues when translated "
    "into dollars. However, AWS demonstrated significant economic resilience: "
    "long-term, dollar-denominated enterprise contracts insulated cloud "
    "revenues from both inflation and currency risk, with AWS contributing "
    "approximately 67% of Amazon's total operating income in 2023 despite "
    "representing a smaller share of total revenue (Amazon, 2024).")

# ── Social ────────────────────────────────────────────────────────
heading(doc, "3.3  Social", level=2)

para(doc,
    "The COVID-19 pandemic accelerated e-commerce adoption globally, "
    "normalising online retail in markets previously dominated by physical "
    "stores. Amazon's Prime subscription service capitalised on this shift, "
    "growing to over 200 million subscribers worldwide (Amazon, 2024). "
    "Simultaneously, sustainability expectations are reshaping consumer "
    "behaviour: packaging waste and rapid-delivery carbon emissions attract "
    "growing scrutiny, particularly in Western Europe and North America. "
    "Amazon's response — a net-zero carbon commitment by 2040 through "
    "The Climate Pledge and an order for 100,000 Rivian electric delivery "
    "vehicles — reflects an attempt to align corporate strategy with "
    "evolving social values (Peng and Meyer, 2019).")



# ── Technological ─────────────────────────────────────────────────
heading(doc, "3.4  Technological", level=2)

para(doc,
    "Cloud computing and generative artificial intelligence represent the "
    "most consequential technological forces shaping Amazon's trajectory. "
    "The global cloud infrastructure market continues to expand rapidly, "
    "driven by enterprise digital transformation and the emergence of "
    "large language models as mainstream business tools. Amazon has "
    "positioned AWS centrally within this opportunity through Amazon Bedrock "
    "— a managed service for deploying generative-AI models — and has "
    "integrated AI-driven automation across its fulfilment network. "
    "Amazon's research and development expenditure reached $85.6 billion "
    "in 2023, reflecting the organisation's recognition that technological "
    "leadership must be actively maintained to sustain its competitive "
    "position (Amazon, 2024).")

# ── Environmental ─────────────────────────────────────────────────
heading(doc, "3.5  Environmental", level=2)

para(doc,
    "Amazon's vast logistics network generates substantial greenhouse gas "
    "emissions, and its Scope 3 emissions — arising from the supply chains "
    "of the thousands of third-party sellers on its marketplace — are "
    "particularly challenging to quantify and reduce. In response, Amazon "
    "has become the world's largest corporate purchaser of renewable "
    "electricity and is deploying 100,000 electric delivery vehicles "
    "in partnership with Rivian. The EU's Corporate Sustainability "
    "Reporting Directive (CSRD) mandates detailed environmental disclosures "
    "for large organisations operating in Europe, increasing compliance "
    "obligations and creating reputational risk if reported progress "
    "falls short of publicly stated targets.")

# ── Legal ─────────────────────────────────────────────────────────
heading(doc, "3.6  Legal", level=2)

para(doc,
    "Amazon navigates a complex multi-jurisdictional legal environment. "
    "The EU's GDPR imposes penalties of up to four per cent of global "
    "annual turnover for material data protection violations; in 2021, "
    "Amazon received a record €746 million GDPR fine from Luxembourg's "
    "data protection authority. Labour law compliance varies significantly "
    "across operating countries, requiring localised policies and legal "
    "counsel. Intellectual property protection — particularly the "
    "challenge of preventing counterfeit goods from being listed on the "
    "marketplace — creates both direct legal liability and brand damage "
    "that undermines consumer confidence. These legal factors illustrate "
    "the interdependence of PESTEL dimensions: political regulation and "
    "social expectations crystallise into tangible legal obligations "
    "that Amazon must proactively manage.")

para(doc,
    "Table 2 summarises the six PESTEL dimensions and their principal "
    "implications for Amazon's global and local operations.", space_before=4)

add_table(doc,
    ["Factor", "Key Issue", "Local Context", "Global Context"],
    [
        ("Political",     "DMA, antitrust, FDI rules",
         "EU/UK marketplace regulation",   "US-China supply-chain risk"),
        ("Economic",      "Inflation, currency, AWS margin",
         "Consumer spend compressed (UK/EU)", "AWS = 67% operating income"),
        ("Social",        "Labour rights, sustainability",
         "UK/DE warehouse strikes",         "200M+ Prime subscribers"),
        ("Technological", "Cloud, GenAI, $85.6B R&D",
         "EU data-sovereignty zones",       "Amazon Bedrock globally"),
        ("Environmental", "Net-zero 2040, CSRD, Rivian EVs",
         "EU CSRD disclosure obligations",  "World's largest renewable buyer"),
        ("Legal",         "GDPR, labour law, IP/counterfeit",
         "€746M GDPR fine (2021)",          "FTC antitrust suit (2023)"),
    ],
    "Table 2: PESTEL Analysis Summary — Amazon Inc."
)



# ═══════════════════════════════════════════════
#  4. CONCLUSION  (~150 words)
# ═══════════════════════════════════════════════

heading(doc, "4. Conclusion")

para(doc,
    "This case study has examined Amazon's external business environment "
    "from two complementary perspectives. Task 1 identified eight business "
    "challenges across entrepreneurial, social, political, and technological "
    "dimensions. Antitrust scrutiny, labour rights concerns, and cybersecurity "
    "vulnerabilities were assessed as carrying the highest combined importance "
    "and impact, given their potential to trigger structural changes to Amazon's "
    "business model or to fundamentally undermine stakeholder trust.")

para(doc,
    "The PESTEL analysis in Task 2 demonstrated that no factor operates "
    "independently. Political regulatory developments are a product of social "
    "attitudes towards platform concentration; AI advances simultaneously "
    "generate economic opportunity and social concern; environmental commitments "
    "intersect with emerging legal disclosure requirements. Amazon's long-term "
    "sustainability therefore depends on its ability to monitor and respond "
    "to these interconnected external forces as an integrated whole — rather "
    "than addressing each in isolation. Frameworks such as PESTEL equip "
    "organisations to develop this systemic understanding, which is a "
    "prerequisite for durable competitive advantage in the global marketplace "
    "(Peng and Meyer, 2019).")

doc.add_page_break()



# ═══════════════════════════════════════════════
#  5. REFERENCES
# ═══════════════════════════════════════════════

heading(doc, "5. References")

refs = [
    "Amazon (2024) Amazon Annual Report 2023. Seattle: Amazon.com, Inc. "
    "Available at: https://ir.aboutamazon.com/annual-reports "
    "(Accessed: 10 June 2026).",

    "European Commission (2022) Antitrust: Commission sends Statement of "
    "Objections to Amazon for use of non-public independent seller data and "
    "opens second investigation into Amazon's e-commerce business practices. "
    "Brussels: European Commission. Available at: "
    "https://ec.europa.eu/commission/presscorner/detail/en/IP_22_7332 "
    "(Accessed: 10 June 2026).",

    "Kantor, J., Weise, K. and Ashford, G. (2021) 'The Amazon that customers "
    "don't see', The New York Times, 15 June. Available at: "
    "https://www.nytimes.com/interactive/2021/06/15/us/amazon-warehouse-"
    "workers.html (Accessed: 10 June 2026).",

    "Peng, M. and Meyer, K. (2019) International Business. 3rd edn. "
    "Andover: Cengage EMEA.",

    "Wright Forrester (2018) Industrial Dynamics. Eastford, CT: "
    "Martino Fine Books.",
]

for ref in refs:
    reference_entry(doc, ref)

# ═══════════════════════════════════════════════
#  SAVE
# ═══════════════════════════════════════════════

doc.save("IBA_Individual_Report.docx")
print("IBA_Individual_Report.docx saved successfully.")
