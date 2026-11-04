import os
import sys
import subprocess
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    """Set cell margins in twips (20 twips = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_background(cell, hex_color="F7F8FA"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_borders(cell, color="000000", sz="4", val="single"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def make_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.RIGHT):
    p.alignment = align
    pPr = p._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    pPr.append(bidi)

def make_run_rtl(run, font_name="Arial", size_pt=9.5, bold=False, color_rgb=(0,0,0)):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rPr.append(rtl)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:cs'), font_name)

def populate_cell_paragraphs(cell, text, size_pt=8, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False):
    """
    Splits multi-line text into individual Word paragraphs to guarantee 100% strict alignment
    on every single line without layout drift or unwanted centering.
    """
    lines = [l.strip() for l in (text or '').splitlines() if l.strip()]
    if not lines:
        p = cell.paragraphs[0]
        p.text = ""
        make_paragraph_rtl(p, align=align)
        return

    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        make_paragraph_rtl(p, align=align)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(line)
        make_run_rtl(r, font_name="Arial", size_pt=size_pt, bold=bold)

def build_header(doc, emblem_path):
    tbl = doc.add_table(rows=1, cols=3)
    tbl.autofit = False
    
    # Table visual RTL
    tblPr = tbl._tbl.tblPr
    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))

    # Column widths: Right (Arabic) 4.2 in, Center (Emblem) 2.4 in, Left (English) 4.2 in
    col_widths = [Inches(4.2), Inches(2.4), Inches(4.2)]
    for i, w in enumerate(col_widths):
        tbl.columns[i].width = w

    # Cell 0: Arabic Header (Rightmost)
    c_ar = tbl.cell(0, 0)
    p_ar = c_ar.paragraphs[0]
    make_paragraph_rtl(p_ar, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p_ar.paragraph_format.space_before = Pt(0)
    p_ar.paragraph_format.space_after = Pt(1)
    p_ar.paragraph_format.line_spacing = 1.15
    r_ar1 = p_ar.add_run("دولة فلسطين\n")
    make_run_rtl(r_ar1, font_name="Arial", size_pt=11, bold=True)
    r_ar2 = p_ar.add_run("وزارة التربية والتعليم العالي")
    make_run_rtl(r_ar2, font_name="Arial", size_pt=11, bold=True)

    # Cell 1: Official Palestine Eagle Emblem in Center
    c_mid = tbl.cell(0, 1)
    p_mid = c_mid.paragraphs[0]
    p_mid.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_mid.paragraph_format.space_before = Pt(0)
    p_mid.paragraph_format.space_after = Pt(0)
    if os.path.exists(emblem_path):
        p_mid.add_run().add_picture(emblem_path, height=Inches(0.62))

    # Cell 2: English Header (Leftmost)
    c_en = tbl.cell(0, 2)
    p_en = c_en.paragraphs[0]
    p_en.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_en.paragraph_format.space_before = Pt(0)
    p_en.paragraph_format.space_after = Pt(1)
    p_en.paragraph_format.line_spacing = 1.15
    r_en1 = p_en.add_run("State of Palestine\n")
    r_en1.font.name = "Arial"
    r_en1.font.size = Pt(9.5)
    r_en1.font.bold = True
    r_en2 = p_en.add_run("Ministry of Education & Higher\nEducation")
    r_en2.font.name = "Arial"
    r_en2.font.size = Pt(9.5)
    r_en2.font.bold = True

    # Remove borders from header
    for row in tbl.rows:
        for c in row.cells:
            set_cell_margins(c, top=0, bottom=0, left=20, right=20)
            tcPr = c._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'  <w:top w:val="none"/>'
                f'  <w:left w:val="none"/>'
                f'  <w:bottom w:val="none"/>'
                f'  <w:right w:val="none"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)

    # Official Sharp Double Line underneath (matching official Palestinian document strictly)
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(2)
    p_line.paragraph_format.space_after = Pt(4)
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/>'
        f'</w:pBdr>'
    )
    p_line._p.get_or_add_pPr().append(pBdr)

def generate_palestine_lesson_docx(data, output_path, emblem_path="palestine_emblem_trans.png"):
    doc = docx.Document()
    
    # Page setup: A4 Landscape with compact margins to guarantee single-page fit
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Inches(0.3)
    section.bottom_margin = Inches(0.28)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

    # 1. Header (Strictly matching Ministry emblem & texts)
    build_header(doc, emblem_path)

    # 2. Document Title
    p_title = doc.add_paragraph()
    make_paragraph_rtl(p_title, align=WD_ALIGN_PARAGRAPH.CENTER)
    p_title.paragraph_format.space_before = Pt(1)
    p_title.paragraph_format.space_after = Pt(3)
    r_t = p_title.add_run("نموذج إعداد الدروس اليومية")
    make_run_rtl(r_t, font_name="Arial", size_pt=12.5, bold=True)

    # 3. Subheader Table: الفصل الدراسي (أقصى اليمين)، الصف، المبحث، المعلم (أقصى اليسار)
    semester_val = data.get('semester', 'الفصل الأول') or 'الفصل الأول'
    grade_val = data.get('grade', '..................')
    subj_val = data.get('subject', '..................')
    teacher_val = data.get('teacher_name', '').strip()

    sub_tbl = doc.add_table(rows=1, cols=4)
    sub_tbl.autofit = False
    sub_tblPr = sub_tbl._tbl.tblPr
    sub_tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    sub_tblPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))

    # Total width: 10.89 inches (distributed so teacher name has ample space and never wraps)
    sub_widths = [Inches(2.5), Inches(2.4), Inches(2.5), Inches(3.49)]
    for i, w in enumerate(sub_widths):
        sub_tbl.columns[i].width = w

    # Cell 0 (أقصى اليمين): الفصل الدراسي
    c_sem = sub_tbl.cell(0, 0)
    p_sem = c_sem.paragraphs[0]
    make_paragraph_rtl(p_sem, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p_sem.paragraph_format.space_before = Pt(0)
    p_sem.paragraph_format.space_after = Pt(3)
    r_sem = p_sem.add_run(f"الفصل الدراسي: {semester_val}")
    make_run_rtl(r_sem, font_name="Arial", size_pt=10, bold=True)

    # Cell 1: الصف
    c_grd = sub_tbl.cell(0, 1)
    p_grd = c_grd.paragraphs[0]
    make_paragraph_rtl(p_grd, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p_grd.paragraph_format.space_before = Pt(0)
    p_grd.paragraph_format.space_after = Pt(3)
    r_grd = p_grd.add_run(f"الصف: {grade_val}")
    make_run_rtl(r_grd, font_name="Arial", size_pt=10, bold=True)

    # Cell 2: المبحث
    c_sub = sub_tbl.cell(0, 2)
    p_sub = c_sub.paragraphs[0]
    make_paragraph_rtl(p_sub, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(3)
    r_sub = p_sub.add_run(f"المبحث: {subj_val}")
    make_run_rtl(r_sub, font_name="Arial", size_pt=10, bold=True)

    # Cell 3 (أقصى اليسار): المعلم/ة
    c_tch = sub_tbl.cell(0, 3)
    p_tch = c_tch.paragraphs[0]
    make_paragraph_rtl(p_tch, align=WD_ALIGN_PARAGRAPH.RIGHT)
    p_tch.paragraph_format.space_before = Pt(0)
    p_tch.paragraph_format.space_after = Pt(3)
    teacher_label = f"المعلم/ـة: {teacher_val}" if teacher_val else "المعلم/ـة: ...................."
    r_tch = p_tch.add_run(teacher_label)
    make_run_rtl(r_tch, font_name="Arial", size_pt=10, bold=True)

    # Remove borders from subheader table
    for row in sub_tbl.rows:
        for c in row.cells:
            set_cell_margins(c, top=0, bottom=0, left=10, right=10)
            tcPr = c._tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'  <w:top w:val="none"/>'
                f'  <w:left w:val="none"/>'
                f'  <w:bottom w:val="none"/>'
                f'  <w:right w:val="none"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)

    # 4. Main Table (6 columns) - Strictly RTL
    headers = [
        "اليوم/التاريخ",
        "عنوان الدرس",
        "نتاج التعلم",
        "الإجراءات والأنشطة",
        "التقويم",
        "ملاحظات"
    ]
    
    col_widths = [
        Inches(1.2),  # اليوم/التاريخ
        Inches(1.5),  # عنوان الدرس
        Inches(2.2),  # نتاج التعلم
        Inches(3.2),  # الإجراءات والأنشطة
        Inches(1.6),  # التقويم
        Inches(1.19)  # ملاحظات
    ]

    lessons = data.get('lessons', [])
    num_rows = 1 + max(len(lessons), 1)
    tbl = doc.add_table(rows=num_rows, cols=6)
    tbl.autofit = False

    # Force Table visual RTL & Center on Page
    tblPr = tbl._tbl.tblPr
    tblPr.append(parse_xml(f'<w:bidiVisual {nsdecls("w")}/>'))
    tblPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))

    # Header Row
    hdr_cells = tbl.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].width = col_widths[i]
        set_cell_margins(hdr_cells[i], top=50, bottom=50, left=50, right=50)
        set_cell_background(hdr_cells[i], "EBEFF2")
        set_cell_borders(hdr_cells[i], color="000000", sz="5", val="single")
        hdr_cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        
        p = hdr_cells[i].paragraphs[0]
        make_paragraph_rtl(p, align=WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        run = p.add_run(title)
        make_run_rtl(run, font_name="Arial", size_pt=9.5, bold=True)

    # Populate Data Rows
    for row_idx, lesson in enumerate(lessons, start=1):
        row = tbl.rows[row_idx]
        row_cells = row.cells
        for col_idx in range(6):
            row_cells[col_idx].width = col_widths[col_idx]
            set_cell_margins(row_cells[col_idx], top=40, bottom=40, left=50, right=50)
            set_cell_borders(row_cells[col_idx], color="000000", sz="4", val="single")
        
        # Col 0: اليوم/التاريخ - توسيط عمودي وأفقي تام
        row_cells[0].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        populate_cell_paragraphs(row_cells[0], lesson.get('day_date', ''), size_pt=8.5, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

        # Col 1: عنوان الدرس - توسيط عمودي وأفقي تام
        row_cells[1].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        populate_cell_paragraphs(row_cells[1], lesson.get('title', ''), size_pt=8.5, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)

        # Col 2: نتاج التعلم - محاذاة يمين صارمة سطر بسطر
        populate_cell_paragraphs(row_cells[2], lesson.get('outcomes', ''), size_pt=8, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False)

        # Col 3: الإجراءات والأنشطة - محاذاة يمين صارمة سطر بسطر
        populate_cell_paragraphs(row_cells[3], lesson.get('procedures', ''), size_pt=8, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False)

        # Col 4: التقويم - محاذاة يمين صارمة سطر بسطر
        populate_cell_paragraphs(row_cells[4], lesson.get('assessment', ''), size_pt=8, align=WD_ALIGN_PARAGRAPH.RIGHT, bold=False)

        # Col 5: ملاحظات (فارغة دائماً لملاحظات وتأملات المعلم اليدوية بالقلم)
        p5 = row_cells[5].paragraphs[0]
        p5.text = ""
        make_paragraph_rtl(p5, align=WD_ALIGN_PARAGRAPH.RIGHT)

    # Save document
    doc.save(output_path)
    return output_path

def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Universal Word-to-PDF converter:
    1. On Windows: Uses Microsoft Word COM for 100% native fidelity.
    2. On Linux / Cloud: Uses LibreOffice headless.
    """
    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path)

    # 1. Try Windows Word COM
    if sys.platform == "win32":
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch('Word.Application')
            word.Visible = False
            try:
                doc = word.Documents.Open(abs_docx)
                doc.ExportAsFixedFormat(abs_pdf, 17) # 17 = wdExportFormatPDF
                doc.Close(False)
            finally:
                word.Quit()
                pythoncom.CoUninitialize()
            if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                return abs_pdf
        except Exception as e:
            print("Windows Word COM failed or not available:", e)

    # 2. Try LibreOffice (Standard on Linux servers / Docker / VPS)
    for cmd in ["libreoffice", "soffice"]:
        try:
            out_dir = os.path.dirname(abs_pdf)
            res = subprocess.run(
                [cmd, "--headless", "--convert-to", "pdf", "--outdir", out_dir, abs_docx],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=30
            )
            # Libreoffice saves as same basename with .pdf
            gen_pdf = os.path.splitext(abs_docx)[0] + ".pdf"
            if os.path.exists(gen_pdf) and gen_pdf != abs_pdf:
                os.replace(gen_pdf, abs_pdf)
            if os.path.exists(abs_pdf) and os.path.getsize(abs_pdf) > 0:
                return abs_pdf
        except Exception as e:
            pass

    return None
