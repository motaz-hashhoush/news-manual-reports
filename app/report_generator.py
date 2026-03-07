import os
from collections import Counter
from datetime import datetime

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

from app.config import UPLOAD_DIR

REPORTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)

ARABIC_MONTHS = {
    1: "يناير", 2: "فبراير", 3: "مارس", 4: "أبريل",
    5: "مايو", 6: "يونيو", 7: "يوليو", 8: "أغسطس",
    9: "سبتمبر", 10: "أكتوبر", 11: "نوفمبر", 12: "ديسمبر",
}

DIST_NARRATIVE_LABELS = {
    "ضيف": "مادة ضيف",
    "مراسل": "مادة مراسلين",
    "تقرير": "تقارير",
    "فيلر": "فيلر",
    "مذيع": "مادة مذيع",
    "عاجل": "مادة عاجلة",
    "مسؤول": "مادة لمسؤولين كبار",
    "وول": "وول",
    "تحليل": "تحليل",
}

DIST_PCT_LABELS = {
    "مذيع": "أخبار المذيعين",
    "ضيف": "الضيوف",
    "مراسل": "المراسلين",
    "عاجل": "الأخبار العاجلة",
    "فيلر": "الفيلرز",
    "مسؤول": "المسؤولين الكبار",
    "تقرير": "التقارير",
    "وول": "الوول",
    "تحليل": "التحليل",
}


def _set_cell_shading(cell, color_hex: str):
    shading = cell._element.get_or_add_tcPr()
    shading_elm = shading.makeelement(
        qn("w:shd"),
        {
            qn("w:val"): "clear",
            qn("w:color"): "auto",
            qn("w:fill"): color_hex,
        },
    )
    shading.append(shading_elm)


def _style_header_row(row, bg_color="1F4E79"):
    for cell in row.cells:
        _set_cell_shading(cell, bg_color)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.bold = True
                run.font.size = Pt(10)


def _set_cell_rtl(cell):
    """Set RTL direction on a table cell so Arabic text flows correctly."""
    for paragraph in cell.paragraphs:
        pPr = paragraph._element.get_or_add_pPr()
        bidi = pPr.makeelement(qn("w:bidi"), {})
        pPr.append(bidi)


def _set_col_widths(table, widths_cm):
    """Set explicit column widths on a table."""
    table.autofit = False
    for row in table.rows:
        for idx, width in enumerate(widths_cm):
            if idx < len(row.cells):
                row.cells[idx].width = Cm(width)


def _format_arabic_dt(dt):
    hour = dt.hour
    if hour >= 12:
        period = "مساءً"
        display_hour = hour - 12 if hour > 12 else 12
    else:
        period = "صباحًا"
        display_hour = hour if hour > 0 else 12
    minute = dt.minute
    month = ARABIC_MONTHS.get(dt.month, str(dt.month))
    time_str = str(display_hour)
    if minute:
        time_str += f":{minute:02d}"
    return f"{time_str} {period} يوم {dt.day} {month} {dt.year}"


def _add_rtl_paragraph(doc, text, bold=False, size=12, alignment=WD_ALIGN_PARAGRAPH.RIGHT):
    p = doc.add_paragraph()
    p.alignment = alignment
    pPr = p._element.get_or_add_pPr()
    bidi = pPr.makeelement(qn("w:bidi"), {})
    pPr.append(bidi)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    return p


def generate_docx_report(report_session, entries, breaking_news_count: int = 0) -> str:
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)

    _add_rtl_paragraph(
        doc,
        "تقرير التغطية الإخبارية – قناة الشرق للأخبار",
        bold=True,
        size=18,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _add_rtl_paragraph(
        doc,
        report_session.name,
        bold=True,
        size=14,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )
    _add_rtl_paragraph(
        doc,
        f"تاريخ التقرير: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        size=11,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    doc.add_paragraph("")

    total_entries = len(entries)
    dist_counter = Counter(e.distribution for e in entries if e.distribution)
    type_counter = Counter(e.entry_type for e in entries if e.entry_type)
    program_counter = Counter(e.program for e in entries if e.program)
    published = sum(1 for e in entries if e.publish_link and e.publish_link.strip() and "لم ينشر" not in e.publish_link)
    not_published = total_entries - published

    # --- Narrative summary ---
    if report_session.start_at and report_session.deadline_at:
        period_from = _format_arabic_dt(report_session.start_at)
        period_to = _format_arabic_dt(report_session.deadline_at)
        period_text = f"من {period_from} حتى {period_to}"
    else:
        period_text = "خلال فترة الجلسة"

    desc_context = report_session.description or "التطورات الإخبارية"

    _add_rtl_paragraph(
        doc,
        f"يرصد هذا التقرير التغطية الإخبارية التي بثّتها الشرق للأخبار خلال الفترة الزمنية "
        f"الممتدة {period_text}، "
        f"وذلك في سياق متابعة التطورات المرتبطة بـ {desc_context}.",
        size=11,
    )

    _add_rtl_paragraph(
        doc,
        "ويهدف التقرير إلى تقديم قراءة تحليلية شاملة للمحتوى الإعلامي الذي ظهر على الشاشة، "
        "من خلال توثيق المواد التي تم بثّها وتحليل طبيعة التغطية من حيث الموضوعات المطروحة، "
        "مصادر المعلومات، أشكال التقديم التلفزيوني، إضافة إلى متابعة انتقال المحتوى من البث "
        "التلفزيوني إلى المنصات الرقمية التابعة للقناة.",
        size=11,
    )

    dist_parts = []
    for dist, count in dist_counter.most_common():
        label = DIST_NARRATIVE_LABELS.get(dist, dist)
        dist_parts.append(f"{count} {label}")
    dist_text = "، ".join(dist_parts)

    stats_para = f"خلال الفترة المرصودة تم تسجيل {total_entries} مادة إعلامية ظهرت على الشاشة"
    if dist_text:
        stats_para += f"، توزعت بين {dist_text}."
    else:
        stats_para += "."
    _add_rtl_paragraph(doc, stats_para, size=11)

    _add_rtl_paragraph(
        doc,
        f"كما أظهر الرصد أن {published} مادة من إجمالي المواد التي ظهرت على الشاشة تم نشرها "
        f"لاحقًا عبر المنصات الرقمية التابعة للقناة، في حين بقيت {not_published} مادة ضمن البث "
        "التلفزيوني دون نشر رقمي حتى وقت إعداد التقرير.",
        size=11,
    )

    _add_rtl_paragraph(
        doc,
        "وتوزعت أشكال التقديم الإعلامي خلال الفترة المرصودة بين تقديم المذيع داخل الاستوديو، "
        "ومداخلات المراسلين من مواقع مختلفة، واستضافة خبراء ومحللين، إضافة إلى استخدام مواد "
        "تفسيرية ومرئية تهدف إلى تبسيط سياق الأحداث وتقديم خلفية تحليلية للمشاهد.",
        size=11,
    )

    sorted_dist = dist_counter.most_common()
    if sorted_dist and total_entries:
        top_dist, top_count = sorted_dist[0]
        top_pct = top_count / total_entries * 100
        top_label = DIST_PCT_LABELS.get(top_dist, top_dist)

        pct_text = (
            f"وتشير بيانات الرصد إلى أن {top_label} شكّلت النسبة الأكبر من إجمالي المواد "
            f"المعروضة، بنسبة {top_pct:.2f}%"
        )

        if len(sorted_dist) > 1:
            rest_parts = []
            for dist, count in sorted_dist[1:]:
                pct = count / total_entries * 100
                label = DIST_PCT_LABELS.get(dist, dist)
                rest_parts.append(f"{label} بنسبة {pct:.2f}%")
            pct_text += "، يليه " + " و".join(rest_parts)

        pct_text += "."
        _add_rtl_paragraph(doc, pct_text, size=11)

    doc.add_paragraph("")
    doc.add_page_break()

    # --- Table 1: Program breakdown ---
    _add_rtl_paragraph(doc, "بطاقة التغطية", bold=True, size=13)
    table1 = doc.add_table(rows=1 + len(program_counter), cols=2)
    table1.style = "Table Grid"
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = table1.rows[0]
    hdr.cells[0].text = "نوع التغطية (اسم البرنامج)"
    hdr.cells[1].text = "العدد"
    _style_header_row(hdr)

    for i, (program, count) in enumerate(program_counter.most_common(), 1):
        table1.rows[i].cells[0].text = program
        table1.rows[i].cells[1].text = str(count)
    for row in table1.rows:
        for cell in row.cells:
            _set_cell_rtl(cell)

    doc.add_paragraph("")

    # --- Table 2: Publication summary ---
    _add_rtl_paragraph(doc, "ملخص النشر", bold=True, size=13)
    table2 = doc.add_table(rows=4, cols=2)
    table2.style = "Table Grid"
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr2 = table2.rows[0]
    hdr2.cells[0].text = "ملخص النشر"
    hdr2.cells[1].text = ""
    _style_header_row(hdr2)

    table2.rows[1].cells[0].text = "إجمالي مواد البث"
    table2.rows[1].cells[1].text = str(total_entries)
    table2.rows[2].cells[0].text = "منشور على السوشيال ميديا"
    table2.rows[2].cells[1].text = str(published)
    table2.rows[3].cells[0].text = "لم يُنشر بعد"
    table2.rows[3].cells[1].text = str(not_published)
    for row in table2.rows:
        for cell in row.cells:
            _set_cell_rtl(cell)

    doc.add_paragraph("")

    # --- Table 3: Distribution stats ---
    _add_rtl_paragraph(doc, "توزيع أشكال التقديم", bold=True, size=13)
    dist_items = dist_counter.most_common()
    table3 = doc.add_table(rows=1 + len(dist_items), cols=3)
    table3.style = "Table Grid"
    table3.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr3 = table3.rows[0]
    hdr3.cells[0].text = "شكل التقديم/المصدر"
    hdr3.cells[1].text = "العدد"
    hdr3.cells[2].text = "النسبة %"
    _style_header_row(hdr3)

    dist_label_map = {
        "ضيف": "عدد الضيوف",
        "مراسل": "عدد المراسلين",
        "تقرير": "عدد التقارير",
        "فيلر": "عدد الفيلرز",
        "مذيع": "أخبار مذيعين",
        "عاجل": "عدد العواجل",
        "مسؤول": "مسؤولون كبار",
        "وول": "وول",
        "تحليل": "تحليل",
    }

    for i, (dist, count) in enumerate(dist_items, 1):
        pct = f"{(count / total_entries * 100):.2f}%" if total_entries else "0%"
        table3.rows[i].cells[0].text = dist_label_map.get(dist, dist)
        table3.rows[i].cells[1].text = str(count)
        table3.rows[i].cells[2].text = pct
    for row in table3.rows:
        for cell in row.cells:
            _set_cell_rtl(cell)

    doc.add_paragraph("")

    # --- Table 4: Type breakdown ---
    _add_rtl_paragraph(doc, "التصنيف الموضوعي", bold=True, size=13)
    type_items = type_counter.most_common()
    table4 = doc.add_table(rows=1 + len(type_items), cols=3)
    table4.style = "Table Grid"
    table4.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr4 = table4.rows[0]
    hdr4.cells[0].text = "التصنيف الموضوعي"
    hdr4.cells[1].text = "العدد"
    hdr4.cells[2].text = "النسبة"
    _style_header_row(hdr4)

    for i, (etype, count) in enumerate(type_items, 1):
        pct = f"{(count / total_entries * 100):.2f}%" if total_entries else "0%"
        table4.rows[i].cells[0].text = etype
        table4.rows[i].cells[1].text = str(count)
        table4.rows[i].cells[2].text = pct
    for row in table4.rows:
        for cell in row.cells:
            _set_cell_rtl(cell)

    doc.add_paragraph("")

    # --- Screenshots section ---
    screenshot_entries = [e for e in entries if e.screenshot_path]
    if screenshot_entries:
        _add_rtl_paragraph(doc, "نماذج (سكرين شوت)", bold=True, size=13)
        for sc_entry in screenshot_entries:
            abs_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                sc_entry.screenshot_path.lstrip("/"),
            )
            if os.path.exists(abs_path):
                try:
                    doc.add_picture(abs_path, width=Inches(5))
                    last_p = doc.paragraphs[-1]
                    last_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception:
                    _add_rtl_paragraph(doc, f"[صورة: {sc_entry.title}]")

        doc.add_paragraph("")

    # --- Table 5: Detailed broadcast log (landscape page) ---
    section = doc.add_section(start_type=2)  # new page
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    _add_rtl_paragraph(doc, "سجل البث التفصيلي", bold=True, size=14, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph("")

    headers = [
        "التوقيت", "العنوان", "البرنامج", "التوزيع",
        "النوع", "الضيف/المراسل", "رابط النشر",
    ]
    col_widths = [2.8, 8.5, 3.0, 2.2, 2.0, 4.0, 4.2]

    table5 = doc.add_table(rows=1 + len(entries), cols=len(headers))
    table5.style = "Table Grid"
    table5.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_col_widths(table5, col_widths)

    hdr5 = table5.rows[0]
    for j, h in enumerate(headers):
        hdr5.cells[j].text = h
        _set_cell_rtl(hdr5.cells[j])
    _style_header_row(hdr5)

    for i, entry in enumerate(entries, 1):
        row = table5.rows[i]

        link = entry.publish_link or ""
        if link.strip() and "لم ينشر" not in link:
            link_display = "نُشر"
        elif "لم ينشر" in link:
            link_display = "لم يُنشر"
        else:
            link_display = ""

        row.cells[0].text = entry.monitoring_time or ""
        row.cells[1].text = entry.title or ""
        row.cells[2].text = entry.program or ""
        row.cells[3].text = entry.distribution or ""
        row.cells[4].text = entry.entry_type or ""
        row.cells[5].text = entry.guest_reporter_name or ""
        row.cells[6].text = link_display

        for cell in row.cells:
            _set_cell_rtl(cell)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for run in paragraph.runs:
                    run.font.size = Pt(8)

    doc.add_paragraph("")

    # revert to portrait for remaining content
    section2 = doc.add_section(start_type=2)
    section2.orientation = WD_ORIENT.PORTRAIT
    section2.page_width = Cm(21)
    section2.page_height = Cm(29.7)
    section2.left_margin = Cm(2.54)
    section2.right_margin = Cm(2.54)

    # --- Summary stats ---
    _add_rtl_paragraph(doc, "ملخص إحصائي سريع", bold=True, size=13)
    stats_table = doc.add_table(rows=4, cols=2)
    stats_table.style = "Table Grid"
    stats_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_stats = stats_table.rows[0]
    hdr_stats.cells[0].text = "المؤشر"
    hdr_stats.cells[1].text = "العدد"
    _style_header_row(hdr_stats)

    stats_table.rows[1].cells[0].text = "إجمالي المواد"
    stats_table.rows[1].cells[1].text = str(total_entries)
    stats_table.rows[2].cells[0].text = "عدد العواجل"
    stats_table.rows[2].cells[1].text = str(dist_counter.get("عاجل", 0))
    stats_table.rows[3].cells[0].text = "عدد الأخبار العاجلة"
    stats_table.rows[3].cells[1].text = str(breaking_news_count)
    for row in stats_table.rows:
        for cell in row.cells:
            _set_cell_rtl(cell)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in report_session.name).strip()
    filename = f"{safe_name}_{timestamp}.docx"
    file_path = os.path.join(REPORTS_DIR, filename)
    doc.save(file_path)

    return file_path
