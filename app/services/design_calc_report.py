"""Builds the downloadable PDF report for the 24in 600# Design Calculation
page - takes the exact input/output values the browser already computed
(sent up from the page's own recalc()) and lays them out to match the
company's own printed calculation sheet (DP103207), i.e. an Excel-style
export: plain borderless rows, gray shading only on manually-entered
cells, a logo/title header and "Page X of Y" footer on every page, and a
green "design is safe" sentence under each check (red if it isn't). No
formulas are re-evaluated here; this only formats what's already on
screen, so the PDF always matches what the user was looking at."""
import io
import os
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

from app.services.app_settings import get_report_logo_path

GREEN = colors.HexColor("#1F7A3D")
RED = colors.HexColor("#C00000")
YELLOW_BG = colors.HexColor("#FFFF00")
GRAY_BG = colors.HexColor("#EDEDED")

_styles = getSampleStyleSheet()
SECTION_STYLE = ParagraphStyle("Section", parent=_styles["Normal"], fontSize=10.5, fontName="Helvetica-Bold", spaceBefore=12, spaceAfter=6, textColor=colors.black)
SUBSECTION_STYLE = ParagraphStyle("Subsection", parent=_styles["Normal"], fontSize=9.5, fontName="Helvetica-Bold", spaceBefore=8, spaceAfter=4, textColor=colors.black)
SAFE_STYLE = ParagraphStyle("Safe", parent=_styles["Normal"], fontSize=8.3, textColor=GREEN, spaceBefore=2, spaceAfter=8)
UNSAFE_STYLE = ParagraphStyle("Unsafe", parent=_styles["Normal"], fontSize=8.3, textColor=RED, spaceBefore=2, spaceAfter=8)
CAPTION_STYLE = ParagraphStyle("Caption", parent=_styles["Normal"], fontSize=8, textColor=colors.grey, alignment=1, spaceBefore=3, spaceAfter=8)
CELL_STYLE = ParagraphStyle("CellX", parent=_styles["Normal"], fontSize=8.3, leading=10.5)
CELL_BOLD = ParagraphStyle("CellB", parent=CELL_STYLE, fontName="Helvetica-Bold")
NOTE_CELL_STYLE = ParagraphStyle("NoteCell", parent=_styles["Normal"], fontSize=7.3, leading=9, textColor=colors.HexColor("#555555"))

PAGE_W, PAGE_H = A4
MARGIN_L = 20 * mm
MARGIN_R = 15 * mm
MARGIN_TOP = 26 * mm
MARGIN_BOTTOM = 16 * mm
CONTENT_W = PAGE_W - MARGIN_L - MARGIN_R

REPORT_TITLE = "Series 10: 24inch 600# Design Calculation Results"


def _num(data, key):
    try:
        return float(data.get(key, ""))
    except (TypeError, ValueError):
        return None


def _fmt(data, key, decimals=3, default="-"):
    v = _num(data, key)
    if v is None:
        raw = data.get(key)
        return str(raw) if raw not in (None, "") else default
    return f"{v:,.{decimals}f}"


def _txt(data, key, default="-"):
    raw = data.get(key)
    return str(raw) if raw not in (None, "") else default


def _pass_fail(data, value_key, compare_key, op=">="):
    a, b = _num(data, value_key), _num(data, compare_key)
    if a is None or b is None:
        return None
    if op == ">=":
        return a >= b
    if op == "<=":
        return a <= b
    if op == "<":
        return a < b
    if op == ">":
        return a > b
    return None


def _cell(text, style=CELL_STYLE, color=None, bold=False):
    if color:
        text = f'<font color="{color.hexval()}"><b>{text}</b></font>' if bold else f'<font color="{color.hexval()}">{text}</font>'
    elif bold:
        style = CELL_BOLD
    return Paragraph(str(text), style)


class _HeaderFooterCanvas(pdfcanvas.Canvas):
    """Collects pages so the footer can print 'Page X of Y' (needs a
    second pass, same trick reportlab's own cookbook uses)."""

    def __init__(self, *args, **kwargs):
        pdfcanvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_states = []

    def showPage(self):
        self._saved_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_states)
        for state in self._saved_states:
            self.__dict__.update(state)
            self._draw_header_footer(total)
            pdfcanvas.Canvas.showPage(self)
        pdfcanvas.Canvas.save(self)

    def _draw_header_footer(self, total_pages):
        c = self
        page_num = self._pageNumber

        # Header: logo top-left, title + ref/date top-right
        if self._logo_path and os.path.exists(self._logo_path):
            try:
                c.drawImage(self._logo_path, MARGIN_L, PAGE_H - 20 * mm, width=22 * mm, height=7.3 * mm,
                            preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        c.setFont("Helvetica-Bold", 12)
        c.setFillColor(colors.black)
        c.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 14 * mm, REPORT_TITLE)
        c.setFont("Helvetica", 8.5)
        c.setFillColor(colors.HexColor("#444444"))
        c.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 19 * mm, self._ref_line)
        c.drawRightString(PAGE_W - MARGIN_R, PAGE_H - 23 * mm, self._date_line)
        c.setStrokeColor(colors.HexColor("#cccccc"))
        c.line(MARGIN_L, PAGE_H - 24.5 * mm, PAGE_W - MARGIN_R, PAGE_H - 24.5 * mm)

        # Footer
        c.setFont("Helvetica", 7.5)
        c.setFillColor(colors.HexColor("#777777"))
        c.drawString(MARGIN_L, 10 * mm, self._doc_no_line)
        c.drawRightString(PAGE_W - MARGIN_R, 10 * mm, f"Page {page_num} of {total_pages}")


def _make_canvas_factory(ref_line, date_line, doc_no_line, logo_path):
    def factory(*args, **kwargs):
        c = _HeaderFooterCanvas(*args, **kwargs)
        c._ref_line = ref_line
        c._date_line = date_line
        c._doc_no_line = doc_no_line
        c._logo_path = logo_path
        return c
    return factory


def _rows_table(rows, col_widths=None):
    """rows: list of (label, symbol, value_text, unit, note, is_input, color, bold).
    Renders as a borderless, Excel-like listing - gray background only on
    the Value cell of input rows, no grid lines anywhere."""
    if col_widths is None:
        col_widths = [0.40 * CONTENT_W, 0.08 * CONTENT_W, 0.16 * CONTENT_W, 0.10 * CONTENT_W, 0.26 * CONTENT_W]

    data = []
    style_cmds = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]
    for i, (label, symbol, value, unit, note, is_input, color, bold) in enumerate(rows):
        data.append([
            _cell(label),
            _cell(symbol, style=NOTE_CELL_STYLE),
            _cell(value, color=color, bold=bold),
            _cell(unit, style=NOTE_CELL_STYLE),
            _cell(note, style=NOTE_CELL_STYLE),
        ])
        if is_input:
            style_cmds.append(("BACKGROUND", (2, i), (2, i), GRAY_BG))

    t = Table(data, colWidths=col_widths, hAlign="LEFT")
    t.setStyle(TableStyle(style_cmds))
    return t


class _Report:
    def __init__(self, data):
        self.data = data
        self.rows = []

    def num(self, key):
        return _num(self.data, key)

    def fmt(self, key, decimals=3):
        return _fmt(self.data, key, decimals)

    def txt(self, key):
        return _txt(self.data, key)

    def ok(self, key, compare_key, op=">="):
        return _pass_fail(self.data, key, compare_key, op)

    def row(self, label, symbol="", unit="", key=None, decimals=3, note="", is_input=False,
             compare_key=None, op=">=", text_value=None):
        value = text_value if text_value is not None else (
            self.fmt(key, decimals) if key else ""
        )
        color, bold = None, False
        if compare_key and key:
            result = self.ok(key, compare_key, op)
            if result is True:
                color, bold = GREEN, True
            elif result is False:
                color, bold = RED, True
        self.rows.append((label, symbol, value, unit, note, is_input, color, bold))

    def flush(self, col_widths=None):
        t = _rows_table(self.rows, col_widths)
        self.rows = []
        return t

    def safe_note(self, ok, safe_text, unsafe_text=None):
        if ok is None:
            return None
        if ok:
            return Paragraph(safe_text, SAFE_STYLE)
        return Paragraph(unsafe_text or ("NOT SAFE: " + safe_text.replace("safe", "").strip()), UNSAFE_STYLE)


def _applicability_row(applicable, not_applicable_text):
    """Matches the source sheet's own convention: gray 'Yes' when this
    subsection applies to the current bonnet type, yellow-highlighted
    'Not applicable...' when it doesn't."""
    cell = _cell("Yes") if applicable else _cell(not_applicable_text)
    tbl = Table([[_cell("Applicability"), cell]], colWidths=[0.4 * CONTENT_W, 0.4 * CONTENT_W])
    style = [("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    style.append(("BACKGROUND", (1, 0), (1, 0), GRAY_BG if applicable else YELLOW_BG))
    tbl.setStyle(TableStyle(style))
    return tbl


def _figure(path_name, caption, width_mm, height_mm=None):
    static_dir = os.path.join(os.path.dirname(__file__), "..", "..", "static", "images", "design_calc")
    path = os.path.normpath(os.path.join(static_dir, path_name))
    if not os.path.exists(path):
        return None
    img = Image(path, width=width_mm * mm, height=(height_mm or width_mm * 0.7) * mm)
    img.hAlign = "CENTER"
    return img


def build_24in_600_report(data: dict) -> io.BytesIO:
    """data: flat {field_id: value} dict, as collected from every
    input/select with an id inside the page's <section class="content">."""
    r = _Report(data)
    bonnet_type = (data.get("bonnet_type") or "bell").lower()
    ref_no = "DP103207"
    ref_iss = "Iss 1"
    ref_date = date.today().strftime("%d %b %Y")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM, leftMargin=MARGIN_L, rightMargin=MARGIN_R,
        title="24in 600# Design Calculation Report",
    )
    story = []

    # ================= 1. General Information =================
    story.append(Paragraph("1. General Information", SECTION_STYLE))
    r.row("Work Order", text_value=r.txt("gi_work_order"), is_input=True)
    r.row("Serial No", text_value=r.txt("gi_serial_no"), is_input=True)
    r.row("Customer", text_value=r.txt("gi_customer"), is_input=True)
    story.append(r.flush())
    r.row("Prepared by", text_value=r.txt("gi_prepared_by"), is_input=True)
    r.row("Checked by", text_value=r.txt("gi_checked_by"), is_input=True)
    story.append(r.flush())

    # ================= 2. Valve Specification =================
    story.append(Paragraph("2. Valve Specification", SECTION_STYLE))
    r.row("Valve Size", text_value=r.txt("vs_size"), unit="in", is_input=True)
    r.row("ASME Class", "Pc", text_value=r.txt("vs_class"), is_input=True)
    r.row("Design Press. @ 38C", "P", unit="bar", key="vs_pressure", decimals=2, is_input=True)
    story.append(r.flush())

    # ================= 3. Material Properties =================
    story.append(Paragraph("3. Material Properties", SECTION_STYLE))
    mat_headers = ["Component", "Body", "Bonnet", "Stud/Bolts"]
    mat_rows = [
        ["Material", r.txt("mat_body_material"), r.txt("mat_bonnet_material"), r.txt("mat_bolt_material")],
        ["Allow. Stress @ 38C (Mpa)", r.fmt("mat_body_stress", 0), r.fmt("mat_bonnet_stress", 0), r.fmt("mat_bolt_stress", 0)],
    ]
    mtab = [[_cell(h, bold=True) for h in mat_headers]] + [[_cell(c) for c in row] for row in mat_rows]
    mstyle = TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (1, 1), (3, 2), GRAY_BG),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, colors.HexColor("#999999")),
    ])
    mt = Table(mtab, colWidths=[0.28 * CONTENT_W, 0.24 * CONTENT_W, 0.24 * CONTENT_W, 0.24 * CONTENT_W])
    mt.setStyle(mstyle)
    story.append(mt)
    story.append(Spacer(1, 10))

    # ================= 4. Part Code Details =================
    story.append(Paragraph("4. Part Code Details", SECTION_STYLE))
    r.row("Body M/c", text_value=r.txt("pc_body_mc"), is_input=True)
    r.row("Body Casting", text_value=r.txt("pc_body_casting"), is_input=True)
    r.row("Bonnet M/c", text_value=r.txt("pc_bonnet_mc"), is_input=True)
    r.row("Bonnet Casting", text_value=r.txt("pc_bonnet_casting"), is_input=True)
    story.append(r.flush())

    # ================= 5. PED Details =================
    story.append(Paragraph("5. PED Details (If PED Applicable)", SECTION_STYLE))
    r.row("Category", text_value=r.txt("ped_category"), is_input=True)
    r.row("Fluid State", text_value=r.txt("ped_fluid_state"), is_input=True)
    r.row("Fluid Group", text_value=r.txt("ped_fluid_group"), is_input=True)
    story.append(r.flush())

    fig = _figure("fig1_ped_chart.png", "", 70)
    if fig:
        story.append(fig)
        story.append(Paragraph("Fig.1 Table-6. Piping referred to in Article 4(1)(c)(i) of PED 2014/68/EU", CAPTION_STYLE))

    story.append(PageBreak())

    # ================= 6. Body Wall Thickness =================
    story.append(Paragraph("6. Body Design: Wall thickness as per ASME B16.34", SECTION_STYLE))
    fig = _figure("fig2_body_cross_section.png", "", 130, 56)
    if fig:
        story.append(fig)
        story.append(Paragraph("Fig.2 Cross-section of Typical Globe Straight Body", CAPTION_STYLE))

    story.append(Paragraph("6.1 Neck Region", SUBSECTION_STYLE))
    r.row("Largest applicable dia", "d1", "mm", "wt_d1", 2, is_input=True)
    r.row("Thickness", "t1", "mm", "wt_t1", 2, note="from ASME B16.34 Table 3A", is_input=True)
    r.row("Corrosion Allowance", "Cr", "mm", "wt_cr1", 1, is_input=True)
    r.row("Required wall thickness", "T1", "mm", "wt_T1", 2, note="T1=t1+Cr")
    r.row("Provided wall thickness", "tp1", "mm", "wt_tp1", 2, is_input=True, compare_key="wt_T1", op=">=")
    story.append(r.flush())
    story.append(r.safe_note(r.ok("wt_tp1", "wt_T1"),
                              "As the provided wall thickness is greater than required wall thickness, the design is safe"))

    story.append(Paragraph("6.2 Top Gallery", SUBSECTION_STYLE))
    r.row("Largest applicable dia", "d2", "mm", "wt_d2", 2, is_input=True)
    r.row("Dia d2' for corresponding d2", "d2'", "mm", "wt_d2p", 2, note="Consider d2'=d2/1.5 only If d2 > 1.5*d1, else d2")
    r.row("Thickness for d2 or d2'", "t2", "mm", "wt_t2", 2, note="from ASME B16.34 Table 3A", is_input=True)
    r.row("Corrosion Allowance", "Cr", "mm", "wt_cr2", 2, is_input=True)
    r.row("Required wall thickness", "T2", "mm", "wt_T2", 2, note="T2=t2+Cr")
    r.row("Provided wall thickness", "tp2", "mm", "wt_tp2", 2, is_input=True, compare_key="wt_T2", op=">=")
    story.append(r.flush())
    story.append(r.safe_note(r.ok("wt_tp2", "wt_T2"),
                              "As the provided wall thickness is greater than required wall thickness, the design is safe"))

    story.append(Paragraph("6.3 Bottom Gallery", SUBSECTION_STYLE))
    r.row("Largest applicable dia", "d3", "mm", "wt_d3", 2, is_input=True)
    r.row("Dia d3' for corresponding d3", "d3'", "mm", "wt_d3p", 2, note="Consider d3'=d3/1.5 only If d3 > 1.5*d1, else d3")
    r.row("Thickness for d3 or d3'", "t3", "mm", "wt_t3", 2, note="from ASME B16.34 Table 3A", is_input=True)
    r.row("Corrosion Allowance", "Cr", "mm", "wt_cr3", 2, is_input=True)
    r.row("Required wall thickness", "T3", "mm", "wt_T3", 2, note="T3=t3+Cr")
    r.row("Provided wall thickness", "tp3", "mm", "wt_tp3", 2, is_input=True, compare_key="wt_T3", op=">=")
    story.append(r.flush())
    story.append(r.safe_note(r.ok("wt_tp3", "wt_T3"),
                              "As the provided wall thickness is greater than required wall thickness, the design is safe"))

    story.append(PageBreak())

    # ================= 7. Bolting Calculation =================
    story.append(Paragraph("7. Bolting Calculation as per ASME Sec VIII Div 1, Mandatory Appendix 2", SECTION_STYLE))
    r.row("Gasket type", text_value=r.txt("b_gasket_type"), is_input=True)
    r.row("Body-gasket ID", "D1", "mm", "b_D1", 2, is_input=True)
    r.row("Body-gasket OD", "D2", "mm", "b_D2", 2, is_input=True)
    r.row("Seat-gasket ID", "D3", "mm", "b_D3", 2, is_input=True)
    r.row("Seat-gasket OD", "D4", "mm", "b_D4", 2, is_input=True)
    r.row("Gasket factor", "m", text_value=r.fmt("b_m", 1), is_input=True)
    r.row("Gasket factor", "y", "bar", "b_y", 0, is_input=True)
    r.row("Actuator Size", "A", text_value=r.fmt("b_A", 0), is_input=True)
    r.row("Body-gasket width", "N", "mm", "b_N", 2, note="(D2-D1)/2")
    r.row("Basic body-gasket seating width", "bo", "mm", "b_bo", 2, note="N/2")
    r.row("Effective body-gasket width", "b", "mm", "b_b", 2, note="If bo <= 6 consider 'bo' or '2.5*(bo)^0.5'")
    r.row("Body-gasket load reaction diameter", "G", "mm", "b_G", 2, note="D2 - (2*b)")
    r.row("Seat-gasket width", "Ns", "mm", "b_Ns", 2, note="(D4-D3)/2")
    r.row("Basic seat-gasket seating width", "bos", "mm", "b_bos", 2, note="Ns/2")
    r.row("Effective seat-gasket width", "bs", "mm", "b_bs", 2, note="If bos <= 6 consider 'bos' or '2.5*(bos)^0.5'")
    r.row("Seat-gasket load reaction diameter", "Gs", "mm", "b_Gs", 2, note="D4 - (2*bs)")
    story.append(r.flush())

    fig = _figure("fig3_gasket_location.png", "", 150, 61.5)
    if fig:
        story.append(fig)
        story.append(Paragraph("Fig.3 Location of Gasket Load Reaction Diameter", CAPTION_STYLE))

    r.row("Total hydrostatic end force", "H", "N", "b_H", 1, note="0.7854*G^2*P")
    r.row("Total joint contact surface compressive load", "Hp", "N", "b_Hp", 1, note="2b*3.14*G*m*P")
    r.row("Min bolt load at operating condition", "Wm1", "N", "b_Wm1", 1, note="H + Hp")
    r.row("Max. possible Actuator Thrust", "Hact", "N", "b_Hact", 1, is_input=True)
    r.row("Operating load", "AA", "N", "b_AA", 1, note="H + Hp + Hact")
    r.row("Body-gasket seating load", "Wm2", "N", "b_Wm2", 1, note="3.14*b*G*y")
    r.row("Seat-gasket seating load", "Wm2s", "N", "b_Wm2s", 1, note="3.14*bs*Gs*y")
    r.row("Assembly load", "BB", "N", "b_BB", 1)
    r.row("Max Load, AA or BB", "", "N", "b_MaxLoad", 1)
    story.append(r.flush())

    story.append(PageBreak())

    # ---- Flange Drilling Details ----
    story.append(Paragraph("Flange Drilling Details", SUBSECTION_STYLE))
    r.row("Minimum bolt area (Operating load)", "Am1", "mm2", "dr_Am1", 0, note="Operating load/Allowable stress of bolt")
    r.row("Minimum bolt area (Assembly load)", "Am2", "mm2", "dr_Am2", 0, note="Assembly load/Allowable stress of bolt")
    r.row("Minimum Bolt Area (ASME B16.4)", "Am3", "mm2", "dr_Am3", 0, note="(0.7854*D2^2*Pc)/(Allow.Stress of bolt*65.26)")
    r.row("Required bolt area, Max (Am1, Am2, Am3)", "A", "mm2", "dr_Areq_mm2", 0)
    r.row("Nominal bolt diameter", "D", "in", "dr_D", 3, is_input=True)
    r.row("Number of Bolts", "NB", text_value=r.fmt("dr_NB", 0), is_input=True)
    r.row("TPI", "n", text_value=r.fmt("dr_TPI", 0), is_input=True)
    r.row("Required bolt area", "A", "in2", "dr_Areq_in2", 2)
    r.row("Tensile stress area", "At", "in2", "dr_At", 2, note="NB*0.7854*[d-(0.9743/n)]^2",
          compare_key="dr_Areq_in2", op=">=")
    story.append(r.flush())
    story.append(r.safe_note(r.ok("dr_At", "dr_Areq_in2"),
                              "As the provided tensile stress area of the bolt is greater than required bolt area, the design is safe"))

    # ---- Hydrostatic Test Load ----
    story.append(Paragraph("Hydrostatic Test Load", SUBSECTION_STYLE))
    r.row("Hydrostatic test pressure", "Hy", "bar", "hy_Hy", 0)
    r.row("Bolt stress at Hydrostatic test pressure", "Sh", "bar", "hy_Sh", 0)
    r.row("Allow. bolt stress at Hydrostatic test pressure", "Sah", "bar", "hy_Sah", 0, compare_key="hy_Sh", op=">=")
    story.append(r.flush())
    story.append(r.safe_note(r.ok("hy_Sah", "hy_Sh"),
                              "As the allowable bolt stress at hydro test pressure is higher than the bolt stress generated at hydro test pressure, the design is safe."))

    # ---- PCD calculation ----
    story.append(Paragraph("PCD calculation", SUBSECTION_STYLE))
    r.row("Bolt spacing", "Ls", "mm", "pcd_Ls", 0, is_input=True)
    r.row("PCD1", "P1", "mm", "pcd_P1", 0, note="(NB*Ls)/3.14")
    r.row("Bonnet ID", "Bid", "mm", "pcd_Bid", 2, is_input=True)
    r.row("Wall thickness", "tp4", "mm", "pcd_tp4", 0, is_input=True)
    r.row("Radial spacing", "Rw", "mm", "pcd_Rw", 0, is_input=True)
    r.row("PCD2", "P2", "mm", "pcd_P2", 0, note="Bid+(2*tp4)+(2*Rw)")
    r.row("Body gasket housing inside diameter", "Gid", "mm", "pcd_Gid", 0, is_input=True)
    r.row("Diameter to consider as per B16.34, 6.1.3b", "", "mm", "pcd_dia1634", 0)
    r.row("Body gallery wall thickness", "Tg", "mm", "pcd_Tg", 2, is_input=True)
    r.row("PCD3", "P3", "mm", "pcd_P3", 0, note="Gid+(0.5*Tg)+(D*25.4)")
    r.row("Yoke Mounting OD", "Y", "mm", "pcd_Y", 0, is_input=True)
    r.row("Clearance", "C1", "mm", "pcd_C1", 2, is_input=True)
    r.row("Square drive socket Dia", "Ds", "mm", "pcd_Ds", 1, is_input=True)
    r.row("PCD 4", "P4", "mm", "pcd_P4", 0, note="Y+(2*C1)+Ds")
    r.row("Req. PCD: Max PCD (P1, P2, P3, P4)", "", "mm", "pcd_ReqPCD", 0)
    r.row("Provided PCD", "PCD", "mm", "pcd_PCD", 2, is_input=True, compare_key="pcd_ReqPCD", op=">=")
    r.row("Required Flange OD", "", "mm", "pcd_ReqFOD", 0, note="PCD+(2*D*25.4)")
    r.row("Provided Flange OD", "FOD", "mm", "pcd_FOD", 2, is_input=True, compare_key="pcd_ReqFOD", op=">=")
    story.append(r.flush())
    pcd_ok = r.ok("pcd_PCD", "pcd_ReqPCD")
    fod_ok = r.ok("pcd_FOD", "pcd_ReqFOD")
    both_ok = None if (pcd_ok is None or fod_ok is None) else (pcd_ok and fod_ok)
    story.append(r.safe_note(both_ok,
                              "As the provided PCD & Flange OD is greater than the required PCD & Flange OD the design is safe."))

    story.append(PageBreak())

    # ================= 8. Bonnet Design =================
    story.append(Paragraph("8. Bonnet Design", SECTION_STYLE))

    # ---- 8.1 Bell bonnet wall thickness ----
    story.append(Paragraph("8.1. Wall thickness calculation (applicable only to bell bonnet, Fig-5) as per ASME Sec VIII Div 1, UG-34", SUBSECTION_STYLE))
    applicable_81 = bonnet_type == "bell"
    story.append(_applicability_row(applicable_81, "Not applicable (Flat Bonnet selected)"))
    story.append(Spacer(1, 4))

    fig4 = _figure("fig4_flat_bonnet.png", "", 62)
    fig5 = _figure("fig5_bell_bonnet.png", "", 62)
    if fig4 and fig5:
        cap_tbl = Table([[fig4, fig5]], colWidths=[CONTENT_W / 2, CONTENT_W / 2])
        cap_tbl.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        story.append(cap_tbl)
        cap_tbl2 = Table([[Paragraph("Fig-4: Flat Bonnet", CAPTION_STYLE), Paragraph("Fig-5: Bell Bonnet", CAPTION_STYLE)]],
                          colWidths=[CONTENT_W / 2, CONTENT_W / 2])
        story.append(cap_tbl2)

    r.row("Diameter of the internal wet area", "d4", "mm", "bw_d4", 2, is_input=True)
    r.row("Attachment Factor UG-34", "C", text_value=r.fmt("bw_C", 2), is_input=True)
    r.row("Pressure Class", "P", "psi", "bw_Ppsi", 2, note="P(bar)*14.5")
    r.row("Allowable Stress (YS*0.67)", "S", "psi", "bw_S", 0, is_input=True)
    r.row("Joint Efficeincy", "E", text_value=r.fmt("bw_E", 0), is_input=True)
    r.row("Required Min Bonnet Wall Thickness", "t", "mm", "bw_treq", 3, note="t = d4*sqrt(C*P/(S*E))")
    r.row("Provided Bonnet Wall Thickness", "t", "mm", "bw_tp", 1, is_input=True, compare_key="bw_treq", op=">=")
    story.append(r.flush())
    story.append(r.safe_note(r.ok("bw_tp", "bw_treq") if applicable_81 else None,
                              "As the provided wall thickness is greater than required wall thickness, the design is safe"))

    fig_d4t = _figure("fig_d4t_diagram.png", "", 60, 21)
    if fig_d4t:
        story.append(fig_d4t)
        story.append(Paragraph("d4 and t as used above", CAPTION_STYLE))

    story.append(PageBreak())

    # ---- 8.2 Flat bonnet flange ----
    story.append(Paragraph("8.2. Flange thickness calculation for Flat Bonnet as per ASME Sec VIII Div 1, UG-34", SUBSECTION_STYLE))
    applicable_82 = bonnet_type == "flat"
    size_txt, cls_txt = r.txt("vs_size"), r.txt("vs_class")
    story.append(_applicability_row(applicable_82, f'Not applicable for {size_txt}" - {cls_txt}# Design'))
    story.append(Spacer(1, 6))

    r.row("Casting Quality Factor", "Qf", text_value=r.fmt("ff_Qf", 2), is_input=True)
    r.row("Allow.Stress of Bonnet w/ Qf", "Sq", "Mpa", "ff_Sq", 1)
    r.row("Joint Efficiency", "E", text_value=r.fmt("ff_E", 2), note="ASME Sec VIII Div 1, Table UW-12", is_input=True)
    r.row("Attachment factor", "C", text_value=r.fmt("ff_C", 2), is_input=True)
    r.row("Lever arm", "hG", "mm", "ff_hG", 1)
    r.row("Flange thickness ( for AA)", "to", "mm", "ff_to", 1, note="G*sqrt(((C*P)/(Sq*E))+((1.9*AA*hG)/(Sq*E*G^3)))")
    r.row("Flange thickness (for BB)", "ta", "mm", "ff_ta", 1, note="G*sqrt((1.9*BB*hG)/(Sq*E*G^3))")
    r.row("Required Flange Thick.", "t", "mm", "ff_treq", 1, note="Max of (to, ta)")
    r.row("Provided Flange Thick.", "td", "mm", "ff_td", 1, is_input=True, compare_key="ff_treq", op=">=")
    story.append(r.flush())
    story.append(r.safe_note(r.ok("ff_td", "ff_treq") if applicable_82 else None,
                              "As the provided flange thickness is greater than required flange thickness, the design is safe"))

    fig6 = _figure("fig6_attachment_factor.png", "", 55, 31)
    if fig6:
        story.append(fig6)
        story.append(Paragraph("Fig-6: Attachment factor for unstayed flat head (flat bonnet)", CAPTION_STYLE))

    story.append(PageBreak())

    # ---- 8.3 Bell bonnet flange (Appendix 2) ----
    story.append(Paragraph("8.3. Flange thickness calculation for Bell Bonnet as per ASME Sec VIII Div 1, Mandatory Appendix-2", SUBSECTION_STYLE))
    applicable_83 = bonnet_type == "bell"
    story.append(_applicability_row(applicable_83, "Not applicable (Flat Bonnet selected)"))
    story.append(Spacer(1, 6))

    r.row("Casting Quality Factor", "Qf", text_value=r.fmt("bf_Qf", 2), is_input=True)
    r.row("Allow.Stress of Bonnet w/ Qf", "Sq", "Mpa", "bf_Sq", 0)
    r.row("Hydrostatic end force on inside of flange", "HD", "N", "bf_HD", 0, note="0.7854*Bid^2*P")
    r.row("Hydrostatic end force", "Ht", "N", "bf_Ht", 0, note="H-HD")
    r.row("Load", "Wm1", "N", "b_Wm1", 0)
    r.row("Gasket load", "HG", "N", "bf_HG", 0, note="Wm1-H")
    r.row("Radial dist. from gasket load reaction to the bolt circle", "hG", "mm", "bf_hG", 1, note="(PCD-G)/2")
    r.row("Radial dist. from bolt circle, to the circle on which HD acts", "hD", "mm", "bf_hD", 3, note="(PCD-Bid-tp4)/2")
    r.row("Radial dist. from bolt circle to the circle on which HT acts", "hT", "mm", "bf_hT", 2, note="((PCD-Bid)/2+hG)/2")
    r.row("Bending moment", "MG", "N.mm", "bf_MG", 0, note="HG*hG")
    r.row("Bending moment", "MD", "N.mm", "bf_MD", 0, note="HD*hD")
    r.row("Bending moment", "MT", "N.mm", "bf_MT", 0, note="HT*hT")
    r.row("Bending moment", "M1", "N.mm", "bf_M1", 0, note="MG+MD+MT")
    r.row("Bending moment", "M2", "N.mm", "bf_M2", 0, note="BB*hG")
    r.row("Max Bending moment", "Mo", "N.mm", "bf_Mo", 0, note="if (Am1 > Am2, M1, M2)")
    r.row("Factor", "K", text_value=r.fmt("bf_K", 2), note="FOD/Bid")
    r.row("Factor", "ho", text_value=r.fmt("bf_ho", 2), note="sqrt (Bid*tp4)")
    r.row("Factor", "F", text_value=r.fmt("bf_F", 5), note="ASME Sec VIII Div 1 Fig 2-7.2", is_input=True)
    r.row("Factor", "e", text_value=r.fmt("bf_e", 3), note="F/ho")
    r.row("Factor", "f", text_value=r.fmt("bf_f", 0), note="ASME Sec VIII Div 1 Fig 2-7.6", is_input=True)
    r.row("Factor", "T", text_value=r.fmt("bf_T", 2), note="ASME Sec VIII Div 1 Fig 2-7.1")
    r.row("Factor", "U", text_value=r.fmt("bf_U", 2), note="ASME Sec VIII Div 1 Fig 2-7.1")
    r.row("Factor", "V", text_value=r.fmt("bf_V", 6), note="ASME Sec VIII Div 1 Fig 2-7.3", is_input=True)
    r.row("Factor", "Q", text_value=r.fmt("bf_Q", 2))
    r.row("Factor", "Y", text_value=r.fmt("bf_Y", 2), note="ASME Sec VIII Div 1 Fig 2-7.1")
    r.row("Factor", "Z", text_value=r.fmt("bf_Z", 2), note="ASME Sec VIII Div 1 Fig 2-7.1")
    r.row("Factor", "dprime", text_value=r.fmt("bf_dprime", 0), note="(U/V)*ho*tp4^2")
    r.row("Required Bonnet Flange Thickness", "Tf", "mm", "bf_Tf", 2, is_input=True)
    r.row("Factor", "L", "mm", "bf_L", 1, note="(1+(Tf*e))/T)+(Tf^3/dprime)")
    r.row("Radial stress", "Sr", "Mpa", "bf_Sr", 0, note="(((1.33Tf*e)+1)Mo)/(L*Tf^2*Bid)")
    r.row("Tangential stress", "St", "Mpa", "bf_St", 0, note="((Y*Mo)/(Tf^2*Bid))-(Z*Sr)")
    r.row("Combined stress", "Srh", "Mpa", "bf_Srh", 0, note="0.5*(Sh+Sr)")
    r.row("Combined stress", "Sth", "Mpa", "bf_Sth", 0, note="0.5*(Sh+St)")
    r.row("MAX(Sr, St, Srh, Sth)", "", "Mpa", "bf_MaxStress", 0, compare_key="bf_Sq", op="<")
    r.row("Longitudinal stress", "Sh", "Mpa", "bf_Sh", 1, note="(f*Mo)/(L*tp4^2*Bid)", compare_key="bf_ShAllow", op="<")
    r.row("Sh.allowable", "", "Mpa", "bf_ShAllow", 1, note="1.5*Sq")
    r.row("Provided Bonnet Flange Thickness", "Tp", "mm", "bf_Tp", 2, is_input=True, compare_key="bf_Tf", op=">=")
    story.append(r.flush())

    tp_ok = r.ok("bf_Tp", "bf_Tf")
    story.append(r.safe_note(tp_ok if applicable_83 else None,
                              "As the provided flange thickness is greater than required flange thickness, the design is safe"))

    doc_no_line = f"{ref_no} {ref_iss.replace(' ', '')}"
    logo_path = get_report_logo_path()
    canvas_factory = _make_canvas_factory(f"Ref No: {ref_no}, {ref_iss}", ref_date, doc_no_line, logo_path)
    doc.build(story, canvasmaker=canvas_factory)
    buffer.seek(0)
    return buffer
