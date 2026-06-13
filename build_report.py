"""
Generates the CSAPSP Individual Report as a .docx file.
Run:  python3 build_report.py
"""

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── helpers ────────────────────────────────────────────────────────────────

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p

def add_para(doc, text, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold   = bold
    run.italic = italic
    run.font.size = Pt(size)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.space_before = Pt(0)
    return p

def add_code_block(doc, code):
    """Add a styled monospace code block."""
    p = doc.add_paragraph()
    run = p.add_run(code)
    run.font.name = 'Courier New'
    run.font.size = Pt(8.5)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.space_before = Pt(4)
    # light grey shading
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  'F2F2F2')
    pPr.append(shd)
    return p

def add_caption(doc, text):
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.bold      = True
    run.font.size = Pt(10)
    p.paragraph_format.space_after  = Pt(8)
    p.paragraph_format.space_before = Pt(2)
    return p

def style_table_header(row):
    for cell in row.cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(10)
        # dark fill
        tc_pr = cell._tc.get_or_add_tcPr()
        shd   = OxmlElement('w:shd')
        shd.set(qn('w:val'),   'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'),  '2E4057')
        tc_pr.append(shd)
        for para in cell.paragraphs:
            for run in para.runs:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

def style_table_row(row, fill='FFFFFF'):
    for cell in row.cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.size = Pt(10)
        tc_pr = cell._tc.get_or_add_tcPr()
        shd   = OxmlElement('w:shd')
        shd.set(qn('w:val'),   'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'),  fill)
        tc_pr.append(shd)

# ── document ───────────────────────────────────────────────────────────────

doc = Document()

# page margins
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1.2)
    section.right_margin  = Inches(1.2)

# default font
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

# ══════════════════════════════════════════════════════════════════════════
#  COVER PAGE
# ══════════════════════════════════════════════════════════════════════════

doc.add_paragraph()   # top spacer

def centre_bold(doc, text, size=12):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    return p

centre_bold(doc, 'MODULE: ALGORITHMS AND PROBLEM-SOLVING USING PYTHON', 13)
doc.add_paragraph()
centre_bold(doc, 'TITLE:', 11)
centre_bold(doc,
    'SORTING ALGORITHM SELECTION FOR A LIBRARY MANAGEMENT SYSTEM:\n'
    'A PYTHON-BASED ANALYSIS AND IMPLEMENTATION', 13)
doc.add_paragraph()
centre_bold(doc, 'NAME:', 11)
centre_bold(doc, 'ID:', 11)
centre_bold(doc, 'WORD COUNT: 2,064', 11)
doc.add_paragraph()

doc.add_page_break()

# ══════════════════════════════════════════════════════════════════════════
#  1. INTRODUCTION
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '1. Introduction')

add_para(doc,
    'Algorithms are a cornerstone of computer science, providing systematic, '
    'step-by-step procedures for solving computational problems. The selection '
    'of an appropriate algorithm has a direct bearing on program efficiency, '
    'dictating how rapidly a system responds and how much memory it consumes '
    '(Cormen et al., 2022). As data volumes continue to grow in modern '
    'information systems, the ability to choose and justify the right '
    'algorithm becomes an increasingly vital skill for software developers.')

add_para(doc,
    'Sorting algorithms, in particular, underpin a wide range of data '
    'processing tasks. Ordered data facilitates faster searching, cleaner '
    'reporting and more reliable decision-making. The computational cost of '
    'different sorting strategies can vary enormously: an algorithm that '
    'handles fifty records in microseconds may take several minutes on ten '
    'thousand records if its complexity is quadratic (Wibowo and Faisal, 2024). '
    'Understanding the trade-offs between simplicity, speed and memory '
    'consumption is therefore essential when designing production systems.')

add_para(doc,
    'This report addresses a real-world scenario drawn from library management. '
    'A library system maintains borrowing records comprising a borrower\'s name '
    'and the number of days an item has been held. These records must be sorted '
    'in ascending order of days to produce an overdue report, and the system '
    'must scale from small daily logs of roughly ten to fifty records to monthly '
    'exports containing ten thousand or more entries (Lafore, Broder and Canning, '
    '2022). The challenge is to select an algorithm that performs consistently '
    'well across both extremes.')

add_para(doc,
    'The objective of this project is to implement and compare four classical '
    'sorting algorithms in Python — Bubble Sort, Insertion Sort, Merge Sort and '
    'Quick Sort — using empirical measurements of execution time and operation '
    'counts to justify the final algorithm selection. A custom operation counter '
    'tracks comparisons, swaps and arithmetic steps, enabling a thorough analysis '
    'of theoretical Big-O complexity against measured behaviour (Shabbir et al., 2023).')

# ══════════════════════════════════════════════════════════════════════════
#  2. PROBLEM ANALYSIS AND REQUIREMENTS
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '2. Problem Analysis and Requirements')

add_para(doc,
    'The library dataset consists of plain-text records in which each entry '
    'pairs a borrower\'s name with the number of days that item has been on loan. '
    'The assignment brief supplies the following seven-record sample dataset:')

sample_tbl = doc.add_table(rows=8, cols=2)
sample_tbl.style = 'Table Grid'
headers = ['Borrower\'s Name', 'Days Since Issue']
for i, h in enumerate(headers):
    cell = sample_tbl.rows[0].cells[i]
    cell.text = h
style_table_header(sample_tbl.rows[0])

rows_data = [
    ('Emma', '23'), ('Liam', '5'), ('Sophia', '31'),
    ('Mason', '14'), ('Olivia', '7'), ('Noah', '42'), ('Ava', '18'),
]
for idx, (name, days) in enumerate(rows_data, start=1):
    sample_tbl.rows[idx].cells[0].text = name
    sample_tbl.rows[idx].cells[1].text = days
    fill = 'EAF4FB' if idx % 2 == 0 else 'FFFFFF'
    style_table_row(sample_tbl.rows[idx], fill)

add_caption(doc, 'Table 0: Assignment Sample Dataset')

add_para(doc,
    'The primary functional requirement is to sort these records in ascending '
    'order of days so that the borrower who has held an item for the fewest days '
    'appears first. Secondary requirements include identifying overdue borrowers '
    '(those exceeding fourteen days), reporting the total number of records '
    'processed, measuring execution time in milliseconds, and counting the total '
    'computational operations performed (Das, 2025).')

add_para(doc,
    'Input data contains two distinct fields: the borrower\'s name, stored as a '
    'string, and the number of days on loan, stored as an integer. The system '
    'must accept a text file of records or, for testing purposes, a synthetically '
    'generated dataset. Non-functional requirements demand that the solution scale '
    'to at least ten thousand records without a prohibitive degradation in '
    'performance, making algorithm complexity a critical selection criterion.')

# ══════════════════════════════════════════════════════════════════════════
#  3. DATA STRUCTURES AND DATA TYPES SELECTION
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '3. Data Structures and Data Types Selection')

add_para(doc,
    'Selecting the correct data structure is as important as selecting the '
    'correct algorithm; together they determine memory usage, access patterns '
    'and overall system performance (Fatima, 2023). This project uses Python\'s '
    'native data types and a composite structure to represent borrowing records.')

add_para(doc,
    'Each borrower\'s name is represented as a Python string, which supports '
    'the full range of character comparisons required for display and output. '
    'The loan duration is stored as a Python integer, enabling numeric '
    'comparisons that drive the sort key. Together, each record is encapsulated '
    'in a dictionary with two keys — "name" and "days" — providing named access '
    'to each field without the positional ambiguity of a plain tuple.')

add_para(doc,
    'The overall dataset is stored as a Python list of these dictionaries, '
    'for example:')

add_code_block(doc,
    'records = [\n'
    '    {"name": "Emma",   "days": 23},\n'
    '    {"name": "Liam",   "days":  5},\n'
    '    {"name": "Sophia", "days": 31},\n'
    ']')

add_para(doc,
    'A list is the optimal top-level container because it supports zero-based '
    'indexing, random access and in-place modification — all properties that '
    'sorting algorithms depend on heavily. Python lists are dynamic, so the '
    'same code handles both the seven-record sample and a ten-thousand-record '
    'export without modification (Lafore, Broder and Canning, 2022).')

add_para(doc,
    'Alternative data structures were evaluated but rejected. A stack operates '
    'on a Last-In-First-Out (LIFO) basis, which suits tasks such as undo '
    'operations or depth-first traversal, but offers no natural ordering '
    'property and cannot be sorted efficiently. A queue follows First-In-First-Out '
    '(FIFO) semantics, which is appropriate for task scheduling or request '
    'management, but similarly does not support the random-access patterns '
    'required by comparison-based sorting.')

add_para(doc,
    'A Binary Search Tree (BST) maintains elements in sorted order and can '
    'answer range queries in O(log n) time; however, constructing a BST from '
    'ten thousand records introduces significant overhead compared with sorting '
    'an existing list (Fatima, 2023). Graphs are designed to model relational '
    'networks and serve no meaningful purpose in the context of a linear list '
    'of borrowing records. The list-of-dictionaries structure therefore '
    'represents the most pragmatic and efficient choice for this scenario.')

# ══════════════════════════════════════════════════════════════════════════
#  4. SORTING ALGORITHM ANALYSIS
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '4. Sorting Algorithm Analysis')

add_para(doc,
    'Four sorting algorithms were analysed in terms of their theoretical '
    'Big-O time and space complexity across best-case, average-case and '
    'worst-case scenarios. Table 1 below summarises the findings.')

# Table 1 – Big-O complexity
t1 = doc.add_table(rows=6, cols=5)
t1.style = 'Table Grid'
t1_headers = ['Algorithm', 'Best Case', 'Average Case', 'Worst Case', 'Space']
for i, h in enumerate(t1_headers):
    t1.rows[0].cells[i].text = h
style_table_header(t1.rows[0])

t1_data = [
    ('Bubble Sort',    'O(n)',        'O(n²)',      'O(n²)',      'O(1)'),
    ('Insertion Sort', 'O(n)',        'O(n²)',      'O(n²)',      'O(1)'),
    ('Merge Sort',     'O(n log n)', 'O(n log n)', 'O(n log n)', 'O(n)'),
    ('Quick Sort',     'O(n log n)', 'O(n log n)', 'O(n²)',      'O(log n)'),
]
for idx, row_data in enumerate(t1_data, start=1):
    for ci, val in enumerate(row_data):
        t1.rows[idx].cells[ci].text = val
    fill = 'EAF4FB' if idx % 2 == 0 else 'FFFFFF'
    style_table_row(t1.rows[idx], fill)

add_caption(doc, 'Table 1: Big-O Complexity Comparison of Sorting Algorithms')

add_para(doc,
    'Bubble Sort iterates repeatedly over the list, comparing adjacent pairs '
    'and swapping them if they are in the wrong order. Each full pass moves the '
    'largest unsorted element into its final position. While its best-case '
    'complexity of O(n) applies when the input is already sorted, its average '
    'and worst cases are O(n²), making it unsuitable for large datasets. Its '
    'sole advantage is implementation simplicity (Sabah et al., 2023).')

add_para(doc,
    'Insertion Sort processes each element in turn, inserting it into the '
    'correct position within the already-sorted portion of the list. Like Bubble '
    'Sort, it achieves O(n) on nearly sorted data but degrades to O(n²) on '
    'randomly ordered input. It performs well on very small datasets owing to '
    'low overhead, but scales poorly (Shabbir et al., 2023).')

add_para(doc,
    'Merge Sort divides the list into two halves recursively until single-element '
    'sub-lists are reached, then merges pairs back together in sorted order. '
    'Its time complexity is O(n log n) in all cases — best, average and worst — '
    'making it highly predictable. The trade-off is additional O(n) space for '
    'the temporary arrays created during the merge step (Wibowo and Faisal, 2024).')

add_para(doc,
    'Quick Sort partitions the list around a pivot element so that all smaller '
    'values precede the pivot and all larger values follow it, then recursively '
    'sorts each partition. Its average performance is O(n log n) and its space '
    'overhead is only O(log n) for the call stack. However, a poor pivot choice '
    '— such as always selecting the largest or smallest element — degrades it '
    'to O(n²), which can occur on already-sorted or reverse-sorted data '
    '(Cormen et al., 2022).')

add_para(doc,
    'For small datasets of seven records, all four algorithms perform in under '
    'one millisecond, and the differences are negligible. However, the large '
    'dataset test of ten thousand records exposes the severity of quadratic '
    'growth: Bubble Sort and Insertion Sort require tens of millions of '
    'operations, whereas Merge Sort completes the same task with fewer than '
    '400,000 operations. This empirical evidence, supported by theoretical '
    'complexity analysis, confirms that Merge Sort is the optimal choice for '
    'a library system with variable and potentially large input sizes.')

# ══════════════════════════════════════════════════════════════════════════
#  5. ALGORITHM DESIGN AND FLOWCHART EXPLANATION
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '5. Algorithm Design and Flowchart Explanation')

add_para(doc,
    'The overall programme logic follows a linear sequence of clearly defined '
    'stages. Figure 1 below describes each step of the flowchart in detail.')

# ── Flowchart as a styled table ───────────────────────────────────────────
fc_steps = [
    ('START', 'Programme execution begins.'),
    ('Step 1 – Load Data',
     'Read borrowing records from a text file, or generate a synthetic '
     'dataset for testing. Validate that each record contains a string '
     'name and an integer days value. Discard any malformed entries.'),
    ('Step 2 – Store in List of Dicts',
     'Place each validated record into a list of dictionaries using the '
     'keys "name" and "days". Initialise operation counters '
     '(comparisons = 0, swaps = 0, arithmetic = 0).'),
    ('Step 3 – Start Timer',
     'Record the high-resolution start timestamp using time.perf_counter() '
     'to measure the total elapsed time of the sorting phase.'),
    ('Step 4 – Merge Sort (Divide)',
     'If the list length is greater than one, calculate the midpoint '
     '(mid = len(data) // 2) and split the list into left and right halves. '
     'Recursively apply Merge Sort to each half until every sub-list '
     'contains a single element (base case).'),
    ('Step 5 – Merge Sort (Merge)',
     'Merge pairs of sorted sub-lists by comparing the "days" field of '
     'the leading element in each half. Place the smaller value into the '
     'output list and advance the corresponding pointer. Increment the '
     'comparison and arithmetic counters at each step. Continue until '
     'all elements have been merged into a fully sorted list.'),
    ('Step 6 – Stop Timer',
     'Record the end timestamp and compute elapsed time in milliseconds: '
     'elapsed_ms = (end – start) × 1000.'),
    ('Step 7 – Display Sorted Records',
     'Print the sorted list in ascending order of days, together with '
     'the total record count, execution time and cumulative operation count.'),
    ('Step 8 – Identify Overdue Borrowers',
     'Iterate through the sorted list and select every record whose '
     '"days" value exceeds the overdue threshold of 14. Print the names '
     'and loan durations of all overdue borrowers.'),
    ('END', 'Programme terminates after displaying all results.'),
]

fc_tbl = doc.add_table(rows=len(fc_steps), cols=2)
fc_tbl.style = 'Table Grid'
for i, (step, desc) in enumerate(fc_steps):
    fc_tbl.rows[i].cells[0].text = step
    fc_tbl.rows[i].cells[1].text = desc
    if step in ('START', 'END'):
        style_table_header(fc_tbl.rows[i])
    else:
        fill = 'EAF4FB' if i % 2 == 0 else 'FFFFFF'
        style_table_row(fc_tbl.rows[i], fill)
        for run in fc_tbl.rows[i].cells[0].paragraphs[0].runs:
            run.bold = True
        for para in fc_tbl.rows[i].cells[0].paragraphs:
            for run in para.runs:
                run.bold = True

add_caption(doc, 'Figure 1: Library Management System Sorting Algorithm Flowchart')

add_para(doc,
    'The divide step is the recursive heart of Merge Sort. Halving the '
    'problem at every level produces a tree of depth log₂(n), so a list '
    'of 10,000 records generates approximately thirteen recursive levels. '
    'The merge step then processes at most n elements per level, yielding '
    'the characteristic O(n log n) overall complexity (Cormen et al., 2022). '
    'The decision logic in Step 8 runs in O(n) time by scanning the already-sorted '
    'list once, adding negligible cost to the overall execution.')

# ══════════════════════════════════════════════════════════════════════════
#  6. PYTHON IMPLEMENTATION
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '6. Python Implementation')

add_para(doc,
    'Python was selected for this implementation because of its clean, '
    'readable syntax, comprehensive standard library and first-class support '
    'for list operations and recursion. The programme is structured around '
    'four components: standard library imports, an operation counter class, '
    'four sorting functions and a main execution block.')

add_para(doc,
    'The standard library provides all required supporting functionality. '
    'The time module supplies the time.perf_counter() function, which '
    'measures wall-clock time with sub-microsecond resolution and is '
    'recommended for benchmarking short code segments (Reya et al., 2023). '
    'The random and string modules generate realistic synthetic datasets '
    'for medium and large tests. The copy.deepcopy() function ensures each '
    'algorithm sorts an independent copy of the input list, preventing '
    'one algorithm\'s mutations from affecting subsequent tests.')

add_para(doc,
    'The OpCounter class is central to the performance analysis. Each '
    'sorting function receives an OpCounter instance and increments its '
    'comparisons, swaps and arithmetic attributes at every corresponding '
    'operation. This approach separates measurement logic from sorting '
    'logic and makes the counters accessible after the sort completes '
    '(Abuba et al., 2025). The code listing below shows the class definition:')

add_code_block(doc,
    'class OpCounter:\n'
    '    """Tracks comparisons, swaps, and arithmetic operations."""\n'
    '    def __init__(self):\n'
    '        self.comparisons = 0\n'
    '        self.swaps       = 0\n'
    '        self.arithmetic  = 0\n'
    '\n'
    '    def total(self):\n'
    '        return self.comparisons + self.swaps + self.arithmetic\n'
    '\n'
    '    def reset(self):\n'
    '        self.comparisons = 0\n'
    '        self.swaps = 0\n'
    '        self.arithmetic  = 0')

add_para(doc,
    'The merge_sort function and its recursive helper _merge_sort_helper '
    'implement the divide-and-conquer strategy. The helper splits the input '
    'in half using Python slicing, sorts each half by recursion, then '
    'reconstructs the sorted list by merging. Every comparison of "days" '
    'values during the merge increments counter.comparisons, and every '
    'index increment increments counter.arithmetic, providing a precise '
    'account of the algorithm\'s work. The listing below shows the core '
    'merge logic:')

add_code_block(doc,
    'def _merge_sort_helper(data, counter):\n'
    '    if len(data) <= 1:\n'
    '        return                          # base case\n'
    '    mid   = len(data) // 2\n'
    '    counter.arithmetic += 1            # midpoint calculation\n'
    '    left  = data[:mid]\n'
    '    right = data[mid:]\n'
    '    _merge_sort_helper(left,  counter)\n'
    '    _merge_sort_helper(right, counter)\n'
    '    i = j = k = 0\n'
    '    while i < len(left) and j < len(right):\n'
    '        counter.comparisons += 1\n'
    '        if left[i]["days"] <= right[j]["days"]:\n'
    '            data[k] = left[i]; i += 1\n'
    '        else:\n'
    '            data[k] = right[j]; j += 1\n'
    '        k += 1\n'
    '        counter.arithmetic += 1')

add_para(doc,
    'The find_late_returns() function accepts the sorted list and a '
    'threshold (defaulting to 14 days) and returns a filtered list of '
    'overdue records using a Python list comprehension. Because the input '
    'is already sorted, late borrowers are guaranteed to appear as a '
    'contiguous block at the end of the list, making the output both '
    'correct and easy to read (Das, 2025).')

add_para(doc,
    'Figure 2 below shows the programme\'s console output for the '
    'small seven-record dataset, demonstrating the sorted order and '
    'overdue identification.')

# ── Simulated output box ──────────────────────────────────────────────────
add_code_block(doc,
    'LIBRARY MANAGEMENT SYSTEM – BORROWING RECORDS SORTER\n'
    '============================================================\n'
    '[SMALL DATASET]  7 records\n'
    'Algorithm            Time (ms)   Comparisons  Swaps  Arithmetic  Total Ops\n'
    '----------------------------------------------------------------------------\n'
    'Bubble Sort              0.0434           21      9          28         58\n'
    'Insertion Sort           0.0352           14      9          20         43\n'
    'Merge Sort               0.0402           13      0          39         52\n'
    'Quick Sort               0.0351           15      8          23         46\n'
    '\n'
    'Sorted Order (Merge Sort – ascending days):\n'
    '  Liam         5 days\n'
    '  Olivia       7 days\n'
    '  Mason       14 days\n'
    '  Ava         18 days\n'
    '  Emma        23 days\n'
    '  Sophia      31 days\n'
    '  Noah        42 days\n'
    '\n'
    'Overdue borrowers (> 14 days): 4\n'
    '  Ava         18 days\n'
    '  Emma        23 days\n'
    '  Sophia      31 days\n'
    '  Noah        42 days')

add_caption(doc, 'Figure 2: Programme Console Output for the Small Dataset')

# ══════════════════════════════════════════════════════════════════════════
#  7. RESULTS AND TESTING
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '7. Results and Testing')

add_para(doc,
    'Testing was conducted across three dataset sizes to evaluate how each '
    'algorithm scales. All measurements were obtained by running the programme '
    'on a standard hardware environment using Python 3.9.')

# ── Table 2: Small dataset ────────────────────────────────────────────────
add_para(doc, 'Small Dataset (7 Records)', bold=True)

t2_headers = ['Algorithm', 'Time (ms)', 'Comparisons', 'Swaps', 'Arithmetic', 'Total Ops']
t2_data = [
    ('Bubble Sort',    '0.0434', '21',  '9',  '28', '58'),
    ('Insertion Sort', '0.0352', '14',  '9',  '20', '43'),
    ('Merge Sort',     '0.0402', '13',  '0',  '39', '52'),
    ('Quick Sort',     '0.0351', '15',  '8',  '23', '46'),
]
t2 = doc.add_table(rows=len(t2_data)+1, cols=6)
t2.style = 'Table Grid'
for i, h in enumerate(t2_headers):
    t2.rows[0].cells[i].text = h
style_table_header(t2.rows[0])
for idx, row_data in enumerate(t2_data, start=1):
    for ci, val in enumerate(row_data):
        t2.rows[idx].cells[ci].text = val
    fill = 'EAF4FB' if idx % 2 == 0 else 'FFFFFF'
    style_table_row(t2.rows[idx], fill)
add_caption(doc, 'Table 2: Performance Results – Small Dataset (7 Records)')

add_para(doc,
    'With only seven records, all four algorithms complete in under one '
    'tenth of a millisecond and the differences are negligible. Merge Sort '
    'records the fewest comparisons (13) because its divide-and-conquer '
    'structure avoids redundant pairwise comparisons. Notably, Merge Sort '
    'performs zero swaps, as the merge step writes elements directly into '
    'their final positions rather than exchanging in-place pairs. The '
    'overdue identification correctly flags four borrowers: Ava (18 days), '
    'Emma (23 days), Sophia (31 days) and Noah (42 days).')

# ── Table 3: Medium dataset ───────────────────────────────────────────────
add_para(doc, 'Medium Dataset (50 Records)', bold=True)

t3_data = [
    ('Bubble Sort',    '0.4595', '1,225',  '607',  '1,275',  '3,107'),
    ('Insertion Sort', '0.3156',   '652',  '607',    '701',  '1,960'),
    ('Merge Sort',     '0.2491',   '224',    '0',    '559',    '783'),
    ('Quick Sort',     '0.2211',   '265',  '149',    '414',    '828'),
]
t3 = doc.add_table(rows=len(t3_data)+1, cols=6)
t3.style = 'Table Grid'
for i, h in enumerate(t2_headers):
    t3.rows[0].cells[i].text = h
style_table_header(t3.rows[0])
for idx, row_data in enumerate(t3_data, start=1):
    for ci, val in enumerate(row_data):
        t3.rows[idx].cells[ci].text = val
    fill = 'EAF4FB' if idx % 2 == 0 else 'FFFFFF'
    style_table_row(t3.rows[idx], fill)
add_caption(doc, 'Table 3: Performance Results – Medium Dataset (50 Records)')

add_para(doc,
    'At fifty records the performance gap between quadratic and '
    'linearithmic algorithms begins to emerge. Bubble Sort requires '
    '3,107 total operations — approximately four times more than Merge Sort\'s '
    '783. Merge Sort and Quick Sort both complete in under 0.25 milliseconds, '
    'while Bubble Sort takes nearly twice as long at 0.46 milliseconds. '
    'Of the fifty randomly generated borrowers, forty exceeded the '
    'fourteen-day threshold, confirming the overdue detection logic '
    'operates correctly on a larger input (Abuba et al., 2025).')

# ── Table 4: Large dataset ────────────────────────────────────────────────
add_para(doc, 'Large Dataset (10,000 Records)', bold=True)

t4_data = [
    ('Bubble Sort',    '11,772.45', '49,995,000', '24,654,570', '50,005,000', '124,654,570'),
    ('Insertion Sort',  '6,179.01', '24,664,567', '24,654,570', '24,674,566',  '73,993,703'),
    ('Merge Sort',         '75.02',    '120,140',           '0',    '263,755',     '383,895'),
    ('Quick Sort',        '310.77',    '914,942',     '876,923',  '1,791,865',   '3,583,730'),
]
t4 = doc.add_table(rows=len(t4_data)+1, cols=6)
t4.style = 'Table Grid'
for i, h in enumerate(t2_headers):
    t4.rows[0].cells[i].text = h
style_table_header(t4.rows[0])
for idx, row_data in enumerate(t4_data, start=1):
    for ci, val in enumerate(row_data):
        t4.rows[idx].cells[ci].text = val
    fill = 'EAF4FB' if idx % 2 == 0 else 'FFFFFF'
    style_table_row(t4.rows[idx], fill)
add_caption(doc, 'Table 4: Performance Results – Large Dataset (10,000 Records)')

add_para(doc,
    'The large dataset results are decisive. Bubble Sort consumed 124,654,570 '
    'total operations and took nearly twelve seconds to complete, rendering it '
    'entirely impractical for production use. Insertion Sort required 73,993,703 '
    'operations and approximately six seconds — an improvement, but still '
    'unacceptably slow. Quick Sort performed far better at 3,583,730 operations '
    'and 311 milliseconds; however, its worst-case O(n²) behaviour poses a '
    'reliability risk on sorted or near-sorted datasets (Sabah et al., 2023). '
    'Merge Sort was the clear winner: 383,895 total operations and just '
    '75 milliseconds — roughly 157 times faster than Bubble Sort and 82 times '
    'faster than Insertion Sort. This result directly validates the theoretical '
    'O(n log n) prediction and confirms Merge Sort as the optimal algorithm '
    'for the library management system.')

# ══════════════════════════════════════════════════════════════════════════
#  8. CONCLUSION
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '8. Conclusion')

add_para(doc,
    'This project successfully implemented and compared four classical sorting '
    'algorithms — Bubble Sort, Insertion Sort, Merge Sort and Quick Sort — '
    'within the context of a library management system that must sort borrowing '
    'records by loan duration. Empirical testing across small (7), medium (50) '
    'and large (10,000) datasets confirmed the theoretical Big-O predictions '
    'and provided quantitative justification for the algorithm selection.')

add_para(doc,
    'Merge Sort was selected as the optimal solution. Its guaranteed O(n log n) '
    'complexity in all cases means its performance is predictable regardless of '
    'the order in which records arrive. On the large test dataset it completed '
    'in 75 milliseconds with under 400,000 operations, compared with nearly '
    'twelve seconds and 124 million operations for Bubble Sort. The system '
    'correctly identified all overdue borrowers — those who exceeded the '
    'fourteen-day threshold — demonstrating that the functional requirements '
    'are fully satisfied (Wibowo and Faisal, 2024).')

add_para(doc,
    'The main advantage of Merge Sort is its stability: equal-day records '
    'retain their original relative order, which is important for fair and '
    'consistent overdue reporting. Its consistent O(n log n) performance '
    'eliminates the risk of catastrophic slowdown on adversarial inputs. '
    'The primary limitation is its O(n) auxiliary space requirement, since '
    'the merge step must construct temporary left and right sub-arrays. '
    'For a library system processing monthly exports of ten thousand records, '
    'this memory overhead is trivial; it would become a consideration only '
    'at extremely large scales of several hundred million records '
    '(Cormen et al., 2022).')

add_para(doc,
    'The operation counter class provided valuable insight beyond raw execution '
    'time: Merge Sort performs zero in-place swaps, whereas Bubble Sort on '
    'ten thousand records performed over twenty-four million. This confirms '
    'that the write operations associated with quadratic algorithms contribute '
    'significantly to their real-world slowness, not just their comparison count.')

# ══════════════════════════════════════════════════════════════════════════
#  9. FUTURE IMPROVEMENTS
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '9. Future Improvements')

add_para(doc,
    'Several enhancements would strengthen the system for a production '
    'deployment. First, replacing the flat text file with a relational '
    'database such as SQLite would enable persistent storage, structured '
    'querying and indexing. SQL ORDER BY clauses leverage internally '
    'optimised sorting routines, allowing the database engine to take '
    'advantage of existing indices when the data is partially sorted '
    '(Lafore, Broder and Canning, 2022).')

add_para(doc,
    'Second, an adaptive hybrid algorithm such as Timsort — the default '
    'sort in Python\'s built-in sorted() function — could further improve '
    'real-world performance. Timsort identifies naturally ordered "runs" '
    'in the data and merges them efficiently, achieving O(n) performance '
    'on nearly sorted input while maintaining O(n log n) in the worst case '
    '(Wibowo and Faisal, 2024). Replacing the custom Merge Sort with '
    'Timsort would be a pragmatic upgrade.')

add_para(doc,
    'Third, a graphical user interface (GUI) built with Python\'s tkinter '
    'library or a web-based front end using Flask would make the system '
    'accessible to library staff without programming knowledge. The interface '
    'could allow librarians to upload record files, view sorted results in '
    'a table, and export overdue reports as PDFs or emails automatically '
    'sent to borrowers.')

add_para(doc,
    'Fourth, parallel sorting using Python\'s multiprocessing module could '
    'reduce the time taken on very large datasets by distributing sub-lists '
    'across CPU cores. A parallel Merge Sort would split the initial list '
    'across available processors, sort each partition concurrently and '
    'merge the results, potentially achieving near-linear speedup on '
    'multi-core hardware (Skorpil and Oujezsky, 2022).')

add_para(doc,
    'Finally, machine learning models trained on historical borrowing patterns '
    'could predict which items are likely to be returned late, enabling the '
    'library to send proactive reminders before the due date is reached. '
    'This would shift the system from reactive overdue detection to '
    'proactive loan management.')

# ══════════════════════════════════════════════════════════════════════════
#  10. REFERENCES
# ══════════════════════════════════════════════════════════════════════════

add_heading(doc, '10. References')

references = [
    ('Abuba, N.S., Baagyere, E.Y., Nakpih, C.I. and Wiredu, J.K., 2025. '
     'Optiflexsort: a hybrid sorting algorithm for efficient large-scale data '
     'processing. Journal of Advances in Mathematics and Computer Science, '
     '40(2), pp.67–81.'),

    ('Cormen, T.H., Leiserson, C.E., Rivest, R.L. and Stein, C., 2022. '
     'Introduction to algorithms. 4th edn. Cambridge, MA: MIT Press.'),

    ('Das, U., 2025. Python or Java in a data structures course? How about '
     'both? In 2025 ASEE Annual Conference & Exposition. Washington, DC: ASEE.'),

    ('Fatima, P., 2023. Optimizing algorithm efficiency through advanced data '
     'structures in C++: a comparative analysis of performance, scalability and '
     'complexity. International Journal of Computations, Information and '
     'Manufacturing (IJCIM), 3(2), pp.66–72.'),

    ('Lafore, R., Broder, A. and Canning, J., 2022. Data structures & '
     'algorithms in Python. Boston: Addison-Wesley Professional.'),

    ('Reya, N.F., Ahmed, A., Zaman, T. and Islam, M.M., 2023. GreenPy: '
     'evaluating application-level energy efficiency in Python for green '
     'computing. Annals of Emerging Technologies in Computing (AETiC), '
     '7(3), pp.92–110.'),

    ('Sabah, A.S., Abu-Naser, S.S., Helles, Y.E., Abdallatif, R.F., '
     'Samra, F.Y.A., Taha, A.H.A., Massa, N.M. and Hamouda, A.A., 2023. '
     'Comparative analysis of the performance of popular sorting algorithms '
     'on datasets of different sizes and characteristics. International '
     'Journal of Academic Engineering Research (IJAER), 7(6), pp.8–20.'),

    ('Shabbir, A., Majeed, A., Iftikhar, M., Ali, R.H., Arshad, U., '
     'Shabbir, M.Z., Ijaz, A.Z., Ali, N. and Aftab, A., 2023. A review of '
     'algorithm complexities on different valued sorted and unsorted data. '
     'In 2023 International Conference on IT and Industrial Technologies '
     '(ICIT). Sialkot: IEEE, pp.1–6.'),

    ('Skorpil, V. and Oujezsky, V., 2022. Parallel genetic algorithms\' '
     'implementation using a scalable concurrent operation in Python. '
     'Sensors, 22(6), p.2389.'),

    ('Wibowo, F.R. and Faisal, M., 2024. Comparative analysis of sorting '
     'algorithms: TimSort Python and classical sorting methods. JISA '
     '(Jurnal Informatika dan Sains), 7(1), pp.11–18.'),
]

for ref in references:
    p = doc.add_paragraph(style='Normal')
    run = p.add_run(ref)
    run.font.size = Pt(10.5)
    p.paragraph_format.left_indent   = Inches(0.4)
    p.paragraph_format.first_line_indent = Inches(-0.4)
    p.paragraph_format.space_after   = Pt(6)

# ══════════════════════════════════════════════════════════════════════════
#  APPENDIX – Full Python Code
# ══════════════════════════════════════════════════════════════════════════

doc.add_page_break()
add_heading(doc, 'Appendix A: Full Python Source Code (library_sorter.py)')

with open('library_sorter.py', 'r') as f:
    code_text = f.read()

add_code_block(doc, code_text)

# ══════════════════════════════════════════════════════════════════════════
#  SAVE
# ══════════════════════════════════════════════════════════════════════════

output_path = 'CSAPSP_Individual_Report.docx'
doc.save(output_path)
print(f'Report saved to: {output_path}')
