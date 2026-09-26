import os
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    """Set the background color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Set inner margins (padding) for a table cell (in twips)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_paragraph_with_spacing(doc, text, style='Normal', space_before=0, space_after=6, line_spacing=1.15, align=WD_ALIGN_PARAGRAPH.LEFT):
    """Add a paragraph with controlled spacing and alignment."""
    p = doc.add_paragraph(style=style)
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        p.add_run(text)
    return p

# Initialize Document
doc = Document()

# Set standard margins (1 inch on all sides)
for section in doc.sections:
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

# Document Font Styling
style = doc.styles['Normal']
style.font.name = 'Arial'
style.font.size = Pt(10)
style.font.color.rgb = RGBColor(0x2C, 0x3E, 0x50) # Charcoal

# Title & Headers
p_school = add_paragraph_with_spacing(doc, "Tai Po Old Market Public School", space_before=10, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
p_school.runs[0].font.size = Pt(14)
p_school.runs[0].font.bold = True
p_school.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66) # Navy Blue

p_app = add_paragraph_with_spacing(doc, "Secondary 1 Admission Application (2027-2028)", space_before=0, space_after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
p_app.runs[0].font.size = Pt(12)
p_app.runs[0].font.bold = True
p_app.runs[0].font.color.rgb = RGBColor(0x00, 0x80, 0x80) # Teal

p_title = add_paragraph_with_spacing(doc, "Student Personal Portfolio (Grades 4 to 6)", space_before=0, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER)
p_title.runs[0].font.size = Pt(12)
p_title.runs[0].font.bold = True
p_title.runs[0].font.color.rgb = RGBColor(0x00, 0x80, 0x80)

# ================= TABLE 0: PERSONAL PROFILE =================
p_sec0 = add_paragraph_with_spacing(doc, "Student Personal Profile (個人基本資料)", space_before=10, space_after=6)
p_sec0.runs[0].font.size = Pt(11)
p_sec0.runs[0].font.bold = True
p_sec0.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

t0_data = [
    [("Chinese Name:", True), ("黃熹澄", False), ("English Name:", True), ("Wong Hei Ching, Phoebe", False)],
    [("Gender:", True), ("Female", False), ("Date of Birth:", True), ("January 9, 2015", False)],
    [("Current School:", True), ("Tai Po Old Market Public School", False), ("Student No. (STRN):", True), ("S121112", False)],
    [("Class (Seat No.):", True), ("6A (31)", False), ("Contact Mobile:", True), ("+852 6855 3258", False)],
    [("Home Address:", True), ("Flat I, 28/F, Hong Shin Court, Sun Hing Garden, Tai Po, N.T.", False), ("", False), ("", False)],
    [("Hobbies & Interests:", True), ("Painting, STEAM, Singing, Basketball, Reading", False), ("", False), ("", False)]
]

table0 = doc.add_table(rows=6, cols=4)
table0.alignment = WD_TABLE_ALIGNMENT.CENTER
for r_idx, row in enumerate(table0.rows):
    data_row = t0_data[r_idx]
    
    # Custom merge for Address & Interests span
    if r_idx in [4, 5]:
        cell_lbl = row.cells[0]
        cell_val = row.cells[1]
        # Merge cell 1, 2, and 3
        cell_val.merge(row.cells[2])
        cell_val.merge(row.cells[3])
        
        # Populate
        cell_lbl.text = data_row[0][0]
        cell_val.text = data_row[1][0]
        
        # Styling
        cell_lbl.paragraphs[0].runs[0].font.bold = True
        set_cell_background(cell_lbl, "F1F3F5")
        set_cell_margins(cell_lbl, top=80, bottom=80, left=100, right=100)
        set_cell_margins(cell_val, top=80, bottom=80, left=100, right=100)
    else:
        for c_idx in range(4):
            cell = row.cells[c_idx]
            label_text, is_bold = data_row[c_idx]
            cell.text = label_text
            if is_bold:
                cell.paragraphs[0].runs[0].font.bold = True
                set_cell_background(cell, "F1F3F5")
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)

# ================= TABLE 1: ACADEMIC RESULTS =================
p_sec1 = add_paragraph_with_spacing(doc, "甲. Academic Performance (學業成績)", space_before=15, space_after=6)
p_sec1.runs[0].font.size = Pt(11)
p_sec1.runs[0].font.bold = True
p_sec1.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

table1 = doc.add_table(rows=8, cols=4)
table1.alignment = WD_TABLE_ALIGNMENT.CENTER

headers1 = ["Grade", "Term", "Subject Scores (Marks)", "Class/Level Rankings & Academic Honours"]
for c_idx, text in enumerate(headers1):
    cell = table1.rows[0].cells[c_idx]
    cell.text = text
    cell.paragraphs[0].runs[0].font.bold = True
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, "003366")
    set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

t1_rows = [
    ["Grade 5 (P5)", "Term 1 Assessment", "Chinese: 91.4 | English: 91.4\nMathematics: 92.0 | General Studies: 98.0", "Class Rank: 7th | Level Rank: 11th\n• Learning Attitude Award\n• First Class Academic Achievement Award"],
    ["Grade 5 (P5)", "Term 2 Assessment", "Chinese: 94.6 | English: 92.7\nMathematics: 86.0 | General Studies: 98.0", "Class Rank: 10th | Level Rank: 22nd\n• First Class Academic Achievement Award"],
    ["Grade 5 (P5)", "Term 3 Assessment", "Chinese: 94.8 | English: 92.3\nMathematics: 96.0 | General Studies: 99.0", "Class Rank: 5th | Level Rank: 12th\n• First Class Academic Achievement Award\n• Full Year: Academic Honour Award, Science & Tech Scholarship, Outstanding Art Award, Global Flight Scholarship, Service Award"],
    ["Grade 4 (P4)", "Term 1 Mid-Term", "Chinese: 96.8 | English: 96.4\nMathematics: 96.0 | General Studies: 100", "Class Rank: 3rd | Level Rank: 6th\n• Learning Attitude Award\n• First Class Academic Achievement Award"],
    ["Grade 4 (P4)", "Term 1 Final", "Chinese: 90.7 | English: 98.8\nMathematics: 95.0 | General Studies: 100", "Class Rank: 3rd | Level Rank: 7th\n• First Class Academic Achievement Award"],
    ["Grade 4 (P4)", "Term 2 Mid-Term", "Chinese: 91.9 | English: 96.0\nMathematics: 97.0 | General Studies: 98.0", "Class Rank: 7th | Level Rank: 10th\n• First Class Academic Achievement Award"],
    ["Grade 4 (P4)", "Term 2 Final", "Chinese: 92.1 | English: 95.3\nMathematics: 96.0 | General Studies: 97.0", "Class Rank: 5th | Level Rank: 10th\n• First Class Academic Achievement Award\n• Full Year: Academic Honour Award, Science & Tech Scholarship, Outstanding Art Award"]
]

# Populate Table 1
for r_idx, r_data in enumerate(t1_rows):
    row_cells = table1.rows[r_idx+1].cells
    for c_idx in range(4):
        row_cells[c_idx].text = r_data[c_idx]
        set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=100, right=100)
        # Highlight Grade 5 row background
        if r_idx < 3:
            set_cell_background(row_cells[c_idx], "FDFCF7")

# Merge cells for P5 Grade spans and P4 Grade spans
table1.rows[1].cells[0].merge(table1.rows[2].cells[0]).merge(table1.rows[3].cells[0])
table1.rows[4].cells[0].merge(table1.rows[5].cells[0]).merge(table1.rows[6].cells[0]).merge(table1.rows[7].cells[0])

# Re-apply text centering vertical alignments after merge
for row_idx in [1, 4]:
    table1.rows[row_idx].cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# ================= TABLE 2: NON-ACADEMIC ACHIEVEMENT =================
p_sec2 = add_paragraph_with_spacing(doc, "乙. Non-Academic Achievements & Awards (學業以外的表現)", space_before=15, space_after=6)
p_sec2.runs[0].font.size = Pt(11)
p_sec2.runs[0].font.bold = True
p_sec2.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

t2_data = [
    # GRADE 5 (P5)
    ["P5", "STEAM — FIRST LEGO League (FLL) Asia Open Championship — 2nd Runner-up (3rd Place)", "Pending"],
    ["P5", "STEAM — HK Bio-Tech STEAM Education Elite Awards — Hong Kong Outstanding STEAM Student Award", "P5_002"],
    ["P5", "STEAM — Now TV STEM Award 2025 — Silver Award", "P5_003"],
    ["P5", "STEAM — HK Inter-School STEAM Gifted Cup (IT) — Grand Final Gold Award (2025-2026)", "P5_004"],
    ["P5", "STEAM — HK Inter-School STEAM Gifted Cup (IT) — Semi-Final Gold Award", "P5_005"],
    ["P5", "STEAM — HK Inter-School STEAM Gifted Cup (IT) — Preliminary Gold Award", "P5_006"],
    ["P5", "STEAM — Queen's College 10th \"Little Inventor\" Competition — Merit Award", "P5_007"],
    ["P5", "STEAM — Tai Po Old Market Public School — \"Passing the Torch\" Global Flight Scholarship", "P5_008"],
    ["P5", "STEAM — Hong Kong Tech Challenge Junior 2026 — Certificate of Participation", "P5_009"],
    ["P5", "STEAM — Tai Po Old Market Public School — Science & Technology Incentive Scholarship", "P5_010"],
    ["P5", "Visual Arts — S.A.M. Project 2025 (HK Triple Potentials Selection) — Visual Arts Potential Student Award", "P5_011"],
    ["P5", "Visual Arts — Creative Fair 2026 (IT & Sports Foundation) — MTF Artistic Scholarship", "P5_012"],
    ["P5", "Visual Arts — Creative Fair 2026 — Pioneer \"Designer Chair\" AI Master Design Award", "P5_013"],
    ["P5", "Visual Arts — Creative Fair 2026 — Pioneer XR Movie Award", "P5_014"],
    ["P5", "Visual Arts — Joy of Travel World HK Children Art & Drawing Contest 2025 — Gold Award", "P5_015"],
    ["P5", "Visual Arts — \"With One Heart\" Red Packet Envelope Design Contest 2026 — Gold Award", "P5_016"],
    ["P5", "Visual Arts — South Asia Books \"HK Primary 100-Storey School Creative Design Contest\" — Outstanding Winner", "P5_017"],
    ["P5", "Visual Arts — HK School Art Arena 3rd Chinese Hard Pen Calligraphy Contest — Silver Award", "P5_018"],
    ["P5", "Visual Arts — Tai Po District Festive Lantern Design Competition 2025 — Merit Award", "P5_019"],
    ["P5", "Visual Arts — Tai Po Old Market Public School — Outstanding Artistic Achievement Award", "P5_020"],
    ["P5", "Visual Arts — \"Creative & Playful in Old Market\" Visual Arts Exhibition — Outstanding Artwork Award", "P5_021"],
    ["P5", "Visual Arts — \"Creative & Playful in Old Market\" Oscar Ceremony 2026 — Best Costume Award", "P5_022"],
    ["P5", "Visual Arts — \"Creative & Playful in Old Market\" Oscar 2026 — Best Craftsmanship & Design Award", "P5_023"],
    ["P5", "Visual Arts — \"Creative & Playful in Old Market\" Oscar 2026 — Best Music Film Award", "P5_024"],
    ["P5", "Visual Arts — School Spring Dream Market 2026 — Most Outstanding Planning Grand Award", "P5_025"],
    ["P5", "Visual Arts — School Creative Designer Award", "P5_026"],
    ["P5", "Visual Arts — 14th World Children Art Exhibition 2025 — Certificate of Participation", "P5_027"],
    ["P5", "Music — 23rd Hong Kong Primary School English Folk Song Singing Contest — 1st Runner-up (Silver)", "P5_028"],
    ["P5", "Music — PTA Annual Talent Gala Night — Outstanding Performance Award", "P5_029"],
    ["P5", "Speech — \"I Love Hong Kong 2026\" Chinese Virtues Inter-School Speech Grand Final — Primary Merit Award", "P5_030"],
    ["P5", "Speech — HK Schools Music & Speech Association 77th Speech Festival — Solo Verse Speaking — Merit", "P5_031"],
    ["P5", "Speech — HK Schools Music & Speech Association 77th Speech Festival — Solo Cantonese Prose — Excellent", "P5_032"],
    ["P5", "Speech — NET Section, CDI, EDB \"Once Upon a Book\" Storytelling Competition — Certificate of Participation", "P5_033"],
    ["P5", "Sports — Tai Po Sam Yuk Secondary School 2025 \"Anti-Drug Cup\" Basketball — 4th Place", "P5_034"],
    ["P5", "Sports — HK Education Bureau — Sport ACT Award Scheme (Gold Medal)", "P5_035"],
    ["P5", "Sports — HK School Physical Fitness Award Scheme — Bronze Award", "P5_036"],
    ["P5", "Leadership & Services — Tai Po District Good Student Award Scheme — Outstanding Performance Award", "P5_037"],
    ["P5", "Leadership & Services — Tai Po Police District — Outstanding Grape Ambassador (Anti-Scam Envoy)", "P5_038"],
    ["P5", "Leadership & Services — Tai Po Police District — Grape Ambassador Appointment Certificate", "P5_039"],
    ["P5", "Leadership & Services — School Service Award (2025-2026)", "P5_040"],
    ["P5", "Leadership & Services — Outstanding Prefect Election (Outstanding Prefect Medal)", "P5_041"],
    ["P5", "Leadership & Services — Prefect Merit Recognition (One School Merit recorded)", "P5_042"],
    ["P5", "Leadership & Services — HK Guide Dogs Association New Territories Flag Day — Volunteer Appreciation", "P5_043"],
    ["P5", "Leadership & Services — \"Let Children Stand Straight\" All HK Flag Day — Volunteer Appreciation", "P5_044"],
    ["P5", "Academic Subjects — HK Scholar Arena 2026 — Success Recognition Award", "P5_045"],
    ["P5", "Academic Subjects — HK Scholar Arena 2026 — English Silver Award", "P5_046"],
    ["P5", "Academic Subjects — HK Scholar Arena 2026 — Chinese Silver Award", "P5_047"],
    ["P5", "Academic Subjects — HK Scholar Arena 2026 — Mathematics Bronze Award", "P5_048"],
    ["P5", "Academic Subjects — School Reading Incentive Scheme (Gold Award)", "P5_049"],
    ["P5", "Academic Subjects — School Good Children Incentive Scheme — Outstanding Five-Star Badge", "P5_052"],
    
    # GRADE 4 (P4)
    ["P4", "STEAM — 2025 FIRST LEGO League (FLL) Asia Open Championship — 2nd Runner-up (3rd Place)", "P4_001"],
    ["P4", "STEAM — FIRST LEGO League (FLL) Hong Kong Tournament 2024-2025 — Overall 2nd Runner-up", "P4_002"],
    ["P4", "STEAM — FIRST LEGO League (FLL) Hong Kong Tournament 2024-2025 — Team Model Award", "P4_003"],
    ["P4", "STEAM — HK Inter-School STEAM Gifted Cup (IT) — Grand Final \"Academic Champion\" 1st Place", "P4_004"],
    ["P4", "STEAM — HK Inter-School STEAM Gifted Cup (IT) — Preliminary Round Gold Award", "P4_005"],
    ["P4", "STEAM — Universal Robotics Challenge (URC) Robot Tournament — Junior Division 1st Class Award", "P4_006"],
    ["P4", "STEAM — Inter-Primary School STEAM Tournament (6-Legged Robot Combat) — 1st Class Award", "P4_007"],
    ["P4", "STEAM — World Scratch Challenge (WSC) Asia Pacific Challenge — Silver Award", "P4_008"],
    ["P4", "STEAM — School Representative Scholarship for Overseas Exchange (2024-2025)", "P4_009"],
    ["P4", "STEAM — London BETT Show 2025 — Marty Robot Innovation & Tech Excellence Award (Prister)", "P4_010"],
    ["P4", "STEAM — 2nd Global Marty Robot Challenge — Best Presentation Award (Prister)", "P4_011"],
    ["P4", "STEAM — Robofest Hong Kong (RoboParade Junior 2025) — Gold Award", "P4_012"],
    ["P4", "STEAM — Robofest Hong Kong (BottleSumo Junior 2025) — Gold Award", "P4_013"],
    ["P4", "STEAM — School Science & Technology Incentive Scholarship", "P4_014"],
    ["P4", "STEAM — HK Association for Computer Education — IT Challenge Scheme (Silver Badge)", "P4_015"],
    ["P4", "STEAM — AI-Assisted Film Virtual Production Competition — Excellent Award", "P4_016"],
    ["P4", "STEAM — Now TV STEM Award 2024 — Silver Award", "P4_017"],
    ["P4", "STEAM — HK Robotic Academy — Robotic Elite Training Program (Excellent Grade)", "P4_018"],
    ["P4", "STEAM — Carmel Pak U Secondary School \"Science Master Challenge 2024\" — Participation Certificate", "P4_019"],
    ["P4", "STEAM — Prister Global Classroom (Class of 2025) — Participation Certificate", "P4_020"],
    ["P4", "Visual Arts — HK Youth & Children Contemporary Drawing Contest 2025 — Gold Award", "P4_021"],
    ["P4", "Visual Arts — HK Diocesan Children Art Contest — Champion", "P4_022"],
    ["P4", "Visual Arts — Tai Po Police District Anti-Scam Wish Plaque Design Contest (Senior Primary) — 1st Runner-up", "P4_023"],
    ["P4", "Visual Arts — School Outstanding Artistic Achievement Award", "P4_024"],
    ["P4", "Visual Arts — PTA Parents-Children Eco-Friendly Household Product Design Contest — Silver Award", "P4_025"],
    ["P4", "Visual Arts — \"Campuses of Love\" Visual Arts & Tech Exhibition — Outstanding Artwork Award", "P4_026"],
    ["P4", "Visual Arts — School Traditional Puppetry Heritage Award", "P4_027"],
    ["P4", "Visual Arts — \"Campuses of Love\" Visual Arts & Tech Exhibition — AI Virtual Singer Award", "P4_028"],
    ["P4", "Visual Arts — World Children Art 13th World Children Painting Grand Prix 2024 — Excellent Award", "P4_029"],
    ["P4", "Visual Arts — HK Polytechnic University 1st Intangible Heritage Cheongsam Design Contest — Finalist", "P4_030"],
    ["P4", "Visual Arts — IT & Sports Foundation — AI Singer Course (Certificate of Completion)", "P4_031"],
    ["P4", "Visual Arts — Hong Kong Projection Mapping Festival — Certificate of Participation", "P4_032"],
    ["P4", "Music — 22nd English Folk Song Contest (Primary Division) — Old Market Voice (Preliminary) — Winner", "P4_033"],
    ["P4", "Music — 22nd English Folk Song Contest — Old Market Voice (Final) — Best Movement Award & Overall 4th", "P4_034"],
    ["P4", "Music — 22nd Hong Kong Primary School English Folk Song Group Contest — Group Merit Award", "P4_035"],
    ["P4", "Sports — School Sports Federation Tai Po District Primary Inter-School Basketball — 2nd Runner-up", "P4_036"],
    ["P4", "Sports — Carmel Pak U Secondary School Tai Po Primary Basketball Tournament — 1st Runner-up", "P4_037"],
    ["P4", "Sports — HK Education Bureau — Sport ACT Award Scheme (Gold Medal)", "P4_038"],
    ["P4", "Sports — \"Active School, Active Life\" Road to Paris Olympics Challenge & MVPA60 — Gold Medal", "P4_039"],
    ["P4", "Sports — 9th Heep Yunn School United Primary School Basketball Invitation — Participation Certificate", "P4_040"],
    ["P4", "Sports — All Hong Kong Primary School Basketball Tournament — Participation Certificate", "P4_041"],
    ["P4", "Sports — \"Five Virtues in Harmony\" Martial Arts Campus Scheme — National Studies Certificate", "P4_042"],
    ["P4", "Academic Subjects — Hong Kong School Elite Competition 2025 (English Group P4) — Bronze Award", "P4_043"],
    ["P4", "Academic Subjects — Putonghua Proficiency Test (SPC) — Intermediate Level Grade A Certificate", "P4_044"],
    ["P4", "Speech — HK Schools Music & Speech Association 76th Speech Festival — Solo Verse Speaking — 1st Runner-up", "P4_045"],
    ["P4", "Speech — HK Schools Music & Speech Association 76th Speech Festival — Solo Cantonese Prose — 2nd Runner-up", "P4_046"],
    ["P4", "Leadership & Services — HK Education Bureau Community Youth Club (CYC) — Junior Primary Bronze Badge", "P4_047"],
    ["P4", "Leadership & Services — Tai Po Police District — Grape Ambassador Appointment Certificate", "P4_048"],
    ["P4", "Leadership & Services — Tai Po Police District Grape Ambassador Scheme 2024-2025 — Appreciation", "P4_049"],
    ["P4", "Leadership & Services — Tai Po District Student Award Scheme — Outstanding Performance Award", "P4_050"],
    ["P4", "Leadership & Services — School Academic Excellence Honour Award (2024-2025)", "P4_051"],
    ["P4", "Leadership & Services — School Prefect Merit Recognition (One School Merit recorded)", "P4_052"],
    ["P4", "Leadership & Services — School Reading Incentive Scheme (Bronze Award)", "P4_053"],
    ["P4", "Leadership & Services — School Good Children Incentive Scheme — Outstanding Four-Star Badge", "P4_054"]
]

table2 = doc.add_table(rows=len(t2_data)+1, cols=3)
table2.alignment = WD_TABLE_ALIGNMENT.CENTER

headers2 = ["Grade", "Award / Achievement", "Remarks / Ref Code"]
for c_idx, text in enumerate(headers2):
    cell = table2.rows[0].cells[c_idx]
    cell.text = text
    cell.paragraphs[0].runs[0].font.bold = True
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, "003366")
    set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

for r_idx, r_data in enumerate(t2_data):
    row_cells = table2.rows[r_idx+1].cells
    for c_idx in range(3):
        row_cells[c_idx].text = r_data[c_idx]
        set_cell_margins(row_cells[c_idx], top=85, bottom=85, left=100, right=100)
        if r_data[0] == "P5":
            set_cell_background(row_cells[c_idx], "FDFCF7")

# Merge cells for P5 and P4 spans in table 2
p5_row_count = sum(1 for d in t2_data if d[0] == "P5")
p4_row_count = len(t2_data) - p5_row_count

# Perform vertical merges for Column 0 (Grade)
# P5 range (row 1 to p5_row_count)
p5_top_cell = table2.rows[1].cells[0]
for idx in range(2, p5_row_count+1):
    p5_top_cell.merge(table2.rows[idx].cells[0])
p5_top_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# P4 range (row p5_row_count+1 to end)
p4_top_cell = table2.rows[p5_row_count+1].cells[0]
for idx in range(p5_row_count+2, len(t2_data)+1):
    p4_top_cell.merge(table2.rows[idx].cells[0])
p4_top_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# ================= TABLE 3: EXTRACURRICULAR ACTIVITIES =================
p_sec3 = add_paragraph_with_spacing(doc, "丙. Extracurricular Activities (課外活動)", space_before=15, space_after=6)
p_sec3.runs[0].font.size = Pt(11)
p_sec3.runs[0].font.bold = True
p_sec3.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

table3 = doc.add_table(rows=3, cols=2)
table3.alignment = WD_TABLE_ALIGNMENT.CENTER

headers3 = ["Grade", "Activities / Teams / Courses"]
for c_idx, text in enumerate(headers3):
    cell = table3.rows[0].cells[c_idx]
    cell.text = text
    cell.paragraphs[0].runs[0].font.bold = True
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, "003366")
    set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

t3_rows = [
    ["Grade 5 (P5)", "• School Robotics Programming Team (Active Member)\n• School Innovation & Technology (STEAM) Team (Active Member)\n• AI-Assisted Virtual Chair Design Course (AIVID)\n• AI-Assisted XR Music Video & Film Production Course\n• International Collaborative Exchange Projects\n• School English Folk Song Choir (\"Old Market Voice\")\n• School Girls' Basketball Team (Active Squad Member)\n• Community Youth Club (CYC) | Junior Police Call (JPC)\n• Student Ambassador for the Constitution and Basic Law"],
    ["Grade 4 (P4)", "• School Robotics Programming Team (Active Member)\n• Robotic Programming Seed Program\n• AI-Assisted Robotic Arm Training Course\n• VPSTEM Micro-Film Virtual Production Training Class\n• AI-Assisted Virtual Singing Program\n• School English Folk Song Choir (\"Old Market Voice\")\n• School Girls' Basketball Team (One School Merit recorded)\n• Community Youth Club (CYC)"]
]

for r_idx, r_data in enumerate(t3_rows):
    row_cells = table3.rows[r_idx+1].cells
    for c_idx in range(2):
        row_cells[c_idx].text = r_data[c_idx]
        set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=100, right=100)
        if r_idx == 0:
            set_cell_background(row_cells[c_idx], "FDFCF7")

# ================= TABLE 4: SCHOOL & SOCIAL SERVICES =================
p_sec4 = add_paragraph_with_spacing(doc, "丁. School & Social Services (校內及社會服務)", space_before=15, space_after=6)
p_sec4.runs[0].font.size = Pt(11)
p_sec4.runs[0].font.bold = True
p_sec4.runs[0].font.color.rgb = RGBColor(0x00, 0x33, 0x66)

table4 = doc.add_table(rows=15, cols=3)
table4.alignment = WD_TABLE_ALIGNMENT.CENTER

headers4 = ["Grade", "Service / Leadership Role", "Remarks / Verification Source"]
for c_idx, text in enumerate(headers4):
    cell = table4.rows[0].cells[c_idx]
    cell.text = text
    cell.paragraphs[0].runs[0].font.bold = True
    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
    set_cell_background(cell, "003366")
    set_cell_margins(cell, top=100, bottom=100, left=120, right=120)

t4_rows = [
    ["P5", "School Prefect (Selected as Outstanding Prefect)", "Code: P5_041"],
    ["P5", "Prefect Duty Excellence (One School Merit recorded)", "Academic Report Card"],
    ["P5", "Student Council Class Representative", "Academic Report Card"],
    ["P5", "Class Monitor", "Academic Report Card"],
    ["P5", "Tai Po Police District — Outstanding Grape Ambassador (Anti-Scam Envoy)", "Code: P5_038"],
    ["P5", "Tai Po Police District — Grape Ambassador (Anti-Scam Envoy)", "Academic Report Card"],
    ["P5", "Junior Prefect Service Alliance", "Academic Report Card"],
    ["P5", "Hong Kong Guide Dogs Association New Territories Flag Day — Volunteer Appreciation", "Code: P5_043"],
    ["P5", "Let Children Stand Straight All Hong Kong Flag Day — Volunteer Appreciation", "Code: P5_044"],
    ["P4", "School Prefect Duty Excellence (One School Merit recorded)", "Academic Report Card"],
    ["P4", "Student Council Class Representative", "Academic Report Card"],
    ["P4", "Class Monitor", "Academic Report Card"],
    ["P4", "Junior Prefect Service Alliance", "Academic Report Card"],
    ["P4", "Tai Po Police District — Grape Ambassador (Anti-Scam Envoy)", "Academic Report Card"]
]

for r_idx, r_data in enumerate(t4_rows):
    row_cells = table4.rows[r_idx+1].cells
    for c_idx in range(3):
        row_cells[c_idx].text = r_data[c_idx]
        set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=100, right=100)
        if r_idx < 9:
            set_cell_background(row_cells[c_idx], "FDFCF7")

# Merge cells for P5 and P4 spans in table 4
p5_services_count = 9
p4_services_count = 5

p5_top_cell_s = table4.rows[1].cells[0]
for idx in range(2, p5_services_count+1):
    p5_top_cell_s.merge(table4.rows[idx].cells[0])
p5_top_cell_s.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

p4_top_cell_s = table4.rows[p5_services_count+1].cells[0]
for idx in range(p5_services_count+2, len(t4_rows)+1):
    p4_top_cell_s.merge(table4.rows[idx].cells[0])
p4_top_cell_s.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# Ensure output path is safe and exists
out_dir = "履歷表 6A31黃熹澄"
if not os.path.exists(out_dir):
    os.makedirs(out_dir)

out_file_path = os.path.join(out_dir, "6A31Wong Hei Ching English Version.docx")
doc.save(out_file_path)

print(f"Success! Beautifully styled English resume document saved to '{out_file_path}'")
