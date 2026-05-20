"""
Generator PPTX prezentacije za predmet "Napredne tehnike racunarske inteligencije".

Tema: LLM-driven pipeline za automatizovan Intelligent Tutoring System
Tim: Uros Petraskovic, Luka Saric, Stefan Lazarevic

Dizajn principi:
- Sve pozicije se izvode iz LAYOUT konstanti (responsive, simetricno, bez hardkodovanja).
- Helperi za naslov, sadrzaj, tabele, citate, image placeholder-e, kod.
- Minimalisticki: belina, jasna hijerarhija, akcentna boja samo za fokus.
- Image placeholder = svetlosivi okvir sa labelom (mesto da tim ubaci sliku).
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from copy import deepcopy
from lxml import etree


# =============================================================================
# DIZAJN TOKENI
# =============================================================================

# Boje (akademski, minimalisticki ton)
C_PRIMARY      = RGBColor(0x1B, 0x36, 0x5D)   # tamno teget (naslovi, akcenti)
C_ACCENT       = RGBColor(0xC2, 0x57, 0x00)   # topla bakarna (highlights, brojevi)
C_TEXT         = RGBColor(0x1F, 0x29, 0x37)   # skoro crna
C_MUTED        = RGBColor(0x6B, 0x72, 0x80)   # siva za sekundarni tekst
C_LIGHT_LINE   = RGBColor(0xE5, 0xE7, 0xEB)   # svetle linije i borderi
C_CARD_BG      = RGBColor(0xF5, 0xF6, 0xF8)   # svetla pozadina kartica
C_IMG_BG       = RGBColor(0xEE, 0xEF, 0xF2)   # placeholder pozadina
C_WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
C_TABLE_HEADER = RGBColor(0x1B, 0x36, 0x5D)
C_TABLE_ALT    = RGBColor(0xFA, 0xFB, 0xFC)

# Tipografija
FONT_FAMILY        = "Calibri"
FONT_FAMILY_MONO   = "Consolas"

FS_TITLE_BIG       = 40   # naslovna strana
FS_TITLE           = 30   # naslov slajda
FS_SECTION_LABEL   = 14   # "SEGMENT B" sitno gore
FS_SECTION_TITLE   = 44   # veliki naslov section dividera
FS_SECTION_SUB     = 18   # podnaslov section dividera
FS_SUBTITLE        = 18
FS_BODY            = 17
FS_BODY_SM         = 15
FS_CAPTION         = 11
FS_FOOTER          = 9
FS_CODE            = 13
FS_QUOTE           = 22

# Layout (16:9 widescreen)
SLIDE_W            = Inches(13.333)
SLIDE_H            = Inches(7.5)

# Spoljne margine
MARGIN_X           = Inches(0.6)
MARGIN_TOP         = Inches(0.45)
MARGIN_BOTTOM      = Inches(0.45)

# Header
HEADER_H           = Inches(0.95)   # visina naslovnog bloka
ACCENT_LINE_H      = Emu(20000)     # tanka akcentna linija ispod naslova
GAP_AFTER_HEADER   = Inches(0.18)

# Footer
FOOTER_H           = Inches(0.3)

# Sirina i visina kontent regiona
CONTENT_X          = MARGIN_X
CONTENT_W          = SLIDE_W - 2 * MARGIN_X
CONTENT_Y          = MARGIN_TOP + HEADER_H + GAP_AFTER_HEADER
CONTENT_H          = SLIDE_H - CONTENT_Y - MARGIN_BOTTOM - FOOTER_H

# Konstantne reference za prezentaciju
PROJECT_FOOTER     = "LLM-driven ITS pipeline - Napredne tehnike racunarske inteligencije"


# =============================================================================
# LOW-LEVEL HELPER-I
# =============================================================================

def set_run(run, *, size=FS_BODY, bold=False, italic=False,
            color=C_TEXT, font=FONT_FAMILY):
    """Postavlja stil run-a (jedinog tekst-fragmenta)."""
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_text_box(slide, x, y, w, h, *, text="", size=FS_BODY, bold=False,
                 italic=False, color=C_TEXT, font=FONT_FAMILY,
                 align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, word_wrap=True):
    """Dodaje textbox sa jednim run-om."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    tf.word_wrap = word_wrap
    tf.vertical_anchor = anchor

    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, italic=italic, color=color, font=font)
    return tb


def add_filled_rect(slide, x, y, w, h, *, fill=C_WHITE, line=None, line_w=0.75):
    """Pravougaonik sa fill bojom i opcionim border-om."""
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    rect.fill.solid()
    rect.fill.fore_color.rgb = fill
    if line is None:
        rect.line.fill.background()
    else:
        rect.line.color.rgb = line
        rect.line.width = Pt(line_w)
    rect.shadow.inherit = False
    return rect


def add_horizontal_line(slide, x, y, w, *, color=C_PRIMARY, height=ACCENT_LINE_H):
    """Tanka horizontalna linija (kao akcent)."""
    return add_filled_rect(slide, x, y, w, height, fill=color, line=None)


def set_solid_bg(slide, color):
    """Postavlja punu pozadinsku boju slajda."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


# =============================================================================
# STANDARDNI BLOKOVI (HEADER, FOOTER, KARTICE)
# =============================================================================

def add_header(slide, *, title, subtitle=None, page_label=None):
    """
    Dosledan zaglavlje bloka:
      - mali "page label" (npr. SEGMENT B - ONTOLOGIJE)
      - veliki naslov
      - opcioni podnaslov
      - tanka akcentna linija
    """
    y = MARGIN_TOP

    if page_label:
        add_text_box(
            slide, MARGIN_X, y, CONTENT_W, Inches(0.22),
            text=page_label.upper(), size=FS_SECTION_LABEL,
            bold=True, color=C_ACCENT,
        )
        title_y = y + Inches(0.27)
    else:
        title_y = y

    add_text_box(
        slide, MARGIN_X, title_y, CONTENT_W, Inches(0.55),
        text=title, size=FS_TITLE, bold=True, color=C_PRIMARY,
    )

    if subtitle:
        add_text_box(
            slide, MARGIN_X, title_y + Inches(0.55), CONTENT_W, Inches(0.3),
            text=subtitle, size=FS_SUBTITLE, italic=True, color=C_MUTED,
        )

    # Akcentna linija na dnu zaglavlja
    line_y = MARGIN_TOP + HEADER_H - Inches(0.05)
    add_horizontal_line(slide, MARGIN_X, line_y, CONTENT_W,
                        color=C_LIGHT_LINE, height=Emu(15000))


def add_footer(slide, *, page_number=None, total=None):
    """Footer iskljucen - korisnik ne zeli ni labelu ni broj slajda."""
    return


def add_card(slide, x, y, w, h, *, title=None, body_lines=None,
             title_color=C_PRIMARY, accent_top=True):
    """Kartica sa svetlom pozadinom, opcionim naslovom i bulletima."""
    card = add_filled_rect(slide, x, y, w, h, fill=C_CARD_BG,
                           line=C_LIGHT_LINE, line_w=0.5)

    pad_x = Inches(0.22)
    pad_y = Inches(0.16)
    inner_x = x + pad_x
    inner_w = w - 2 * pad_x

    cur_y = y + pad_y

    if accent_top:
        add_horizontal_line(slide, x, y, w, color=C_PRIMARY,
                            height=Emu(25000))

    if title:
        add_text_box(
            slide, inner_x, cur_y, inner_w, Inches(0.3),
            text=title, size=FS_SUBTITLE, bold=True, color=title_color,
        )
        cur_y += Inches(0.36)

    if body_lines:
        remaining_h = (y + h) - cur_y - pad_y
        add_bullet_list(slide, inner_x, cur_y, inner_w, remaining_h,
                        body_lines, size=FS_BODY_SM)

    return card


def add_image_placeholder(slide, x, y, w, h, *, label="Slika"):
    """
    Vizuelni placeholder za buduce ubacivanje slike.
    Pravi svetlo-sivi okvir sa labelom u centru.
    """
    add_filled_rect(slide, x, y, w, h, fill=C_IMG_BG,
                    line=C_LIGHT_LINE, line_w=0.75)

    # ikonica simulirana karakterom + tekst
    icon_h = Inches(0.4)
    icon_tb = add_text_box(
        slide, x, y + h / 2 - icon_h, w, icon_h,
        text="[ Slika ]", size=20, bold=True, color=C_MUTED,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.BOTTOM,
    )

    add_text_box(
        slide, x, y + h / 2 + Inches(0.05), w, Inches(0.4),
        text=label, size=FS_CAPTION, italic=True, color=C_MUTED,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP,
    )


# =============================================================================
# BULLET LISTE I PARAGRAFI
# =============================================================================

def add_bullet_list(slide, x, y, w, h, items, *,
                    size=FS_BODY, color=C_TEXT, line_spacing=1.25,
                    space_after=4, bullet_char="•"):
    """
    Bullet lista. items moze biti:
      - string: obican bullet
      - (text, "strong"): section header (uppercase, accent, bez crtice)
      - (text, "sub"): indented sub-bullet
    Prazan string "" pravi mali vertikalni razmak izmedju grupa.
    """
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)
    tf.word_wrap = True

    first = True
    for item in items:
        kind = None
        text = item
        if isinstance(item, tuple):
            text = item[0]
            kind = item[1] if len(item) > 1 else None

        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = line_spacing

        if text == "":
            # Spacer izmedju grupa
            p.space_after = Pt(6)
            p.space_before = Pt(0)
            run = p.add_run()
            run.text = " "
            set_run(run, size=4, color=color)
            continue

        if kind == "strong":
            # Section header: mala uppercase labela u akcentnoj boji
            p.space_before = Pt(4)
            p.space_after = Pt(3)
            run = p.add_run()
            run.text = text.upper()
            set_run(run, size=11, bold=True, color=C_ACCENT)
        elif kind == "sub":
            p.space_after = Pt(space_after)
            run = p.add_run()
            run.text = "      ◦  " + text
            set_run(run, size=size - 1, color=color)
        else:
            p.space_after = Pt(space_after)
            run = p.add_run()
            run.text = f"{bullet_char}  " + text
            set_run(run, size=size, color=color)

    return tb


def add_two_column_bullets(slide, x, y, w, h, left_title, left_items,
                            right_title, right_items, *, gap=Inches(0.35)):
    """Dve simetricne kolone sa naslovima i bulletima."""
    col_w = (w - gap) / 2

    # Levo
    add_text_box(slide, x, y, col_w, Inches(0.32),
                 text=left_title, size=FS_SUBTITLE, bold=True, color=C_PRIMARY)
    add_horizontal_line(slide, x, y + Inches(0.36), Inches(0.6),
                        color=C_ACCENT, height=Emu(20000))
    add_bullet_list(slide, x, y + Inches(0.5), col_w, h - Inches(0.5),
                    left_items, size=FS_BODY)

    # Desno
    rx = x + col_w + gap
    add_text_box(slide, rx, y, col_w, Inches(0.32),
                 text=right_title, size=FS_SUBTITLE, bold=True, color=C_PRIMARY)
    add_horizontal_line(slide, rx, y + Inches(0.36), Inches(0.6),
                        color=C_ACCENT, height=Emu(20000))
    add_bullet_list(slide, rx, y + Inches(0.5), col_w, h - Inches(0.5),
                    right_items, size=FS_BODY)


# =============================================================================
# TABELE
# =============================================================================

def add_styled_table(slide, x, y, w, h, headers, rows, *,
                     col_widths=None, header_color=C_TABLE_HEADER,
                     header_text=C_WHITE, alt=True):
    """
    Stilizovana tabela: header trake u akcentnoj boji, alternativne redove
    u svetloj boji, tanke linije.
    """
    n_cols = len(headers)
    n_rows = len(rows) + 1
    table_shape = slide.shapes.add_table(n_rows, n_cols, x, y, w, h)
    table = table_shape.table

    # Sirine kolona
    if col_widths:
        total = sum(col_widths)
        for i, ratio in enumerate(col_widths):
            table.columns[i].width = int(w * ratio / total)

    # Header
    for i, htext in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_color
        cell.margin_left = cell.margin_right = Inches(0.1)
        cell.margin_top = cell.margin_bottom = Inches(0.06)
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = htext
        set_run(run, size=FS_BODY_SM, bold=True, color=header_text)

    # Rows
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_TABLE_ALT if (alt and r % 2 == 0) else C_WHITE
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.margin_top = cell.margin_bottom = Inches(0.05)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            run = p.add_run()
            run.text = str(val)
            set_run(run, size=FS_BODY_SM, color=C_TEXT)

    return table


# =============================================================================
# CODE BLOK
# =============================================================================

def add_code_block(slide, x, y, w, h, code_text, *, caption=None):
    """Tamniji blok sa monospace tekstom (za code/SPARQL/JSON snippete)."""
    bg = add_filled_rect(slide, x, y, w, h, fill=RGBColor(0x1F, 0x29, 0x37))
    pad = Inches(0.18)
    tb = slide.shapes.add_textbox(x + pad, y + pad, w - 2 * pad, h - 2 * pad)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0)
    tf.margin_top = tf.margin_bottom = Inches(0)

    for i, line in enumerate(code_text.splitlines()):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.space_after = Pt(2)
        run = p.add_run()
        run.text = line if line else " "
        set_run(run, size=FS_CODE, color=RGBColor(0xE5, 0xE7, 0xEB),
                font=FONT_FAMILY_MONO)

    if caption:
        add_text_box(slide, x, y + h + Inches(0.05), w, Inches(0.25),
                     text=caption, size=FS_CAPTION, italic=True, color=C_MUTED)


# =============================================================================
# SHEME SLAJDOVA (TEMPLATES)
# =============================================================================

def new_slide(prs):
    """Prazan layout (bez placeholdera) i bela pozadina."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_solid_bg(slide, C_WHITE)
    return slide


def make_title_slide(prs, *, title, subtitle, team, course, date_str):
    slide = new_slide(prs)
    set_solid_bg(slide, C_WHITE)

    # Tanka leva traka (akcent)
    add_filled_rect(slide, Inches(0), Inches(0), Inches(0.18), SLIDE_H,
                    fill=C_PRIMARY)

    # Course label (gore)
    add_text_box(
        slide, Inches(0.85), Inches(0.7), SLIDE_W - Inches(1.6), Inches(0.3),
        text=course.upper(), size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
    )

    # Naslov
    add_text_box(
        slide, Inches(0.85), Inches(1.5), SLIDE_W - Inches(1.6), Inches(2.0),
        text=title, size=FS_TITLE_BIG, bold=True, color=C_PRIMARY,
    )

    # Akcentna linija
    add_horizontal_line(slide, Inches(0.85), Inches(3.7), Inches(1.2),
                        color=C_ACCENT, height=Emu(30000))

    # Podnaslov
    add_text_box(
        slide, Inches(0.85), Inches(3.9), SLIDE_W - Inches(1.6), Inches(1.2),
        text=subtitle, size=20, italic=True, color=C_MUTED,
    )

    # Tim
    add_text_box(
        slide, Inches(0.85), Inches(6.0), SLIDE_W - Inches(1.6), Inches(0.4),
        text="Tim", size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
    )
    add_text_box(
        slide, Inches(0.85), Inches(6.3), SLIDE_W - Inches(1.6), Inches(0.4),
        text=team, size=FS_BODY, color=C_TEXT,
    )

    # Datum
    add_text_box(
        slide, Inches(0.85), Inches(6.8), SLIDE_W - Inches(1.6), Inches(0.3),
        text=date_str, size=FS_CAPTION, color=C_MUTED,
    )

    return slide


def make_section_divider(prs, *, segment_label, title, subtitle, time_str):
    slide = new_slide(prs)
    set_solid_bg(slide, C_PRIMARY)

    # Akcentna pruga
    add_filled_rect(slide, Inches(0), SLIDE_H - Inches(0.4), SLIDE_W, Inches(0.4),
                    fill=C_ACCENT)

    add_text_box(
        slide, Inches(0.85), Inches(2.4), SLIDE_W - Inches(1.7), Inches(0.4),
        text=segment_label.upper(), size=18, bold=True,
        color=RGBColor(0xF5, 0xC0, 0x60),
    )

    add_text_box(
        slide, Inches(0.85), Inches(2.85), SLIDE_W - Inches(1.7), Inches(1.6),
        text=title, size=FS_SECTION_TITLE, bold=True, color=C_WHITE,
    )

    add_horizontal_line(slide, Inches(0.85), Inches(4.55), Inches(1.5),
                        color=C_ACCENT, height=Emu(40000))

    add_text_box(
        slide, Inches(0.85), Inches(4.75), SLIDE_W - Inches(1.7), Inches(1.0),
        text=subtitle, size=FS_SECTION_SUB, italic=True,
        color=RGBColor(0xD1, 0xD5, 0xDB),
    )

    add_text_box(
        slide, Inches(0.85), SLIDE_H - Inches(1.0), SLIDE_W - Inches(1.7), Inches(0.4),
        text=time_str, size=FS_CAPTION, color=RGBColor(0xC4, 0xC9, 0xD4),
    )

    return slide


def make_content_slide(prs, *, page_label, title, subtitle=None,
                       body_func=None, page_no=None, total_pages=None):
    slide = new_slide(prs)
    add_header(slide, title=title, subtitle=subtitle, page_label=page_label)
    if body_func:
        body_func(slide)
    add_footer(slide, page_number=page_no, total=total_pages)
    return slide


def make_quote_slide(prs, *, page_label, quote, attribution=None,
                     page_no=None, total_pages=None):
    slide = new_slide(prs)

    # Tanak akcent gore
    add_horizontal_line(slide, Inches(0), Inches(0), SLIDE_W,
                        color=C_PRIMARY, height=Emu(40000))

    # Sitan label
    add_text_box(
        slide, MARGIN_X, Inches(0.9), CONTENT_W, Inches(0.3),
        text=page_label.upper(), size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
    )

    # Veliki citat
    add_text_box(
        slide, Inches(1.5), Inches(2.4), SLIDE_W - Inches(3.0), Inches(3.0),
        text=f"“{quote}”", size=FS_QUOTE, italic=True, color=C_PRIMARY,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
    )

    if attribution:
        add_text_box(
            slide, Inches(1.5), Inches(5.5), SLIDE_W - Inches(3.0), Inches(0.4),
            text=f"- {attribution}", size=FS_BODY, color=C_MUTED,
            align=PP_ALIGN.CENTER,
        )

    add_footer(slide, page_number=page_no, total=total_pages)
    return slide


def make_thank_you_slide(prs, *, team):
    slide = new_slide(prs)
    set_solid_bg(slide, C_PRIMARY)

    add_text_box(
        slide, Inches(0.85), Inches(2.8), SLIDE_W - Inches(1.7), Inches(1.4),
        text="Hvala na paznji.", size=54, bold=True, color=C_WHITE,
        align=PP_ALIGN.CENTER,
    )

    add_horizontal_line(slide, SLIDE_W / 2 - Inches(0.6), Inches(4.2), Inches(1.2),
                        color=C_ACCENT, height=Emu(40000))

    add_text_box(
        slide, Inches(0.85), Inches(4.5), SLIDE_W - Inches(1.7), Inches(0.6),
        text="Pitanja i diskusija", size=24, italic=True,
        color=RGBColor(0xE5, 0xE7, 0xEB), align=PP_ALIGN.CENTER,
    )

    add_text_box(
        slide, Inches(0.85), Inches(6.0), SLIDE_W - Inches(1.7), Inches(0.4),
        text=team, size=FS_BODY, color=RGBColor(0xC4, 0xC9, 0xD4),
        align=PP_ALIGN.CENTER,
    )

    add_filled_rect(slide, Inches(0), SLIDE_H - Inches(0.3), SLIDE_W, Inches(0.3),
                    fill=C_ACCENT)
    return slide


# =============================================================================
# SADRZAJ - SLAJDOVI
# =============================================================================
# Svaki body_* helper prima 'slide' i koristi CONTENT_X / CONTENT_Y /
# CONTENT_W / CONTENT_H za simetrican raspored.

def body_agenda(slide):
    items = [
        ("Uvod i motivacija",            "10 min", "Sta i zasto - od ishoda ucenja do ITS-a"),
        ("LLM prompting i taksonomije",  "25 min", "Bloom, SOLO, PS4 prompting, hijerarhija zadataka"),
        ("Ontologije i formalni grounding","25 min", "OWL, RDF, SPARQL i anti-halucinacioni mehanizmi"),
        ("Evaluacija i ogranicenja",     "25 min", "Kvalitet po nivoima, robustnost, human-in-the-loop"),
        ("Sinteza, buduci rad, zakljucak","10-15 min","Integrisana arhitektura i naredni koraci"),
    ]

    n = len(items)
    gap = Inches(0.12)
    row_h = (CONTENT_H - (n - 1) * gap) / n

    for i, (title, time, sub) in enumerate(items):
        y = CONTENT_Y + i * (row_h + gap)

        # leva traka
        add_filled_rect(slide, CONTENT_X, y, Inches(0.12), row_h,
                        fill=C_ACCENT)

        # broj
        add_text_box(
            slide, CONTENT_X + Inches(0.25), y, Inches(0.7), row_h,
            text=f"0{i}.", size=24, bold=True, color=C_PRIMARY,
            anchor=MSO_ANCHOR.MIDDLE,
        )

        # naslov
        add_text_box(
            slide, CONTENT_X + Inches(1.05), y + Inches(0.05),
            CONTENT_W - Inches(3.0), Inches(0.45),
            text=title, size=20, bold=True, color=C_PRIMARY,
        )

        # podnaslov
        add_text_box(
            slide, CONTENT_X + Inches(1.05), y + Inches(0.5),
            CONTENT_W - Inches(3.0), row_h - Inches(0.5),
            text=sub, size=FS_BODY_SM, color=C_MUTED,
        )

        # vreme desno
        add_text_box(
            slide, CONTENT_X + CONTENT_W - Inches(1.9), y,
            Inches(1.9), row_h,
            text=time, size=FS_BODY, bold=True, color=C_ACCENT,
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE,
        )


def body_team(slide):
    members = [
        ("Uros Petraskovic",
         "SOLO Quiz Generator",
         "Generisanje pitanja po SOLO taksonomiji uz ontoloski grounding i source-line citate."),
        ("Luka Saric",
         "Course Structure Generator",
         "Automatska struktura kursa po Bloomovoj taksonomiji - moduli, koncepti, ishodi."),
        ("Stefan Lazarevic",
         "Computer Use Video Tutorials",
         "Video demonstracije praktickih zadataka uz OWL validaciju izvrsnog plana."),
    ]

    gap = Inches(0.3)
    col_w = (CONTENT_W - 2 * gap) / 3

    for i, (name, role, desc) in enumerate(members):
        x = CONTENT_X + i * (col_w + gap)

        # kartica
        add_filled_rect(slide, x, CONTENT_Y, col_w, CONTENT_H - Inches(0.3),
                        fill=C_CARD_BG, line=C_LIGHT_LINE, line_w=0.75)

        # akcent na vrhu kartice
        add_filled_rect(slide, x, CONTENT_Y, col_w, Inches(0.08), fill=C_ACCENT)

        pad = Inches(0.25)
        inner_x = x + pad
        inner_w = col_w - 2 * pad
        cur_y = CONTENT_Y + Inches(0.3)

        # avatar krug placeholder
        avatar_size = Inches(1.0)
        avatar_x = x + (col_w - avatar_size) / 2
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, avatar_x, cur_y,
                                        avatar_size, avatar_size)
        circle.fill.solid()
        circle.fill.fore_color.rgb = C_PRIMARY
        circle.line.fill.background()
        # inicijal
        ini = name.split()[0][0] + name.split()[1][0]
        tb = circle.text_frame
        tb.margin_left = tb.margin_right = Inches(0)
        tb.margin_top = tb.margin_bottom = Inches(0)
        p = tb.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = ini
        set_run(run, size=28, bold=True, color=C_WHITE)

        cur_y += avatar_size + Inches(0.2)

        add_text_box(
            slide, inner_x, cur_y, inner_w, Inches(0.4),
            text=name, size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
            align=PP_ALIGN.CENTER,
        )
        cur_y += Inches(0.4)

        add_text_box(
            slide, inner_x, cur_y, inner_w, Inches(0.35),
            text=role, size=FS_BODY_SM, italic=True, color=C_ACCENT,
            align=PP_ALIGN.CENTER,
        )
        cur_y += Inches(0.5)

        add_text_box(
            slide, inner_x, cur_y, inner_w, CONTENT_H - (cur_y - CONTENT_Y) - Inches(0.3),
            text=desc, size=FS_BODY_SM, color=C_TEXT,
            align=PP_ALIGN.CENTER,
        )


def body_predmet_context(slide):
    # Leva polovina - tekst
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y, half_w, CONTENT_H,
        [
            ("Predmet: Napredne tehnike racunarske inteligencije", "strong"),
            "Master studije, FTN Univerzitet u Novom Sadu",
            "Predavaci: dr Prokic i dr Kovacevic",
            "",
            ("Tema u preseku dve oblasti", "strong"),
            "AI in (Computing) Education",
            "Intelligent Tutoring Systems",
            "",
            ("Format polaganja", "strong"),
            "Snimljena prezentacija 1h - 1h30min",
            "Demonstracija integrisanog ITS pipeline-a",
        ],
        size=FS_BODY,
    )

    # Desna polovina - vizuelni placeholder
    right_x = CONTENT_X + half_w + Inches(0.4)
    add_image_placeholder(
        slide, right_x, CONTENT_Y, half_w, CONTENT_H,
        label="Logo FTN-a / vizual predmeta",
    )


def body_problem_stats(slide):
    """3 brojke + sazet opis problema."""
    # Gornji red - 3 velike brojke
    stats_h = Inches(2.1)
    gap = Inches(0.3)
    col_w = (CONTENT_W - 2 * gap) / 3

    stats = [
        ("> 40%",
         "radnog vremena nastavnika",
         "trosi se na pripremu materijala (planiranje, primeri, video)"),
        ("4 nivoa",
         "kognitivne slozenosti",
         "moraju biti pokriveni pitanjima (SOLO taksonomija)"),
        ("3 izvora",
         "ulaza za sistem",
         "ishodi ucenja, PDF materijali, prirodno-jezicke instrukcije"),
    ]

    for i, (num, lbl, sub) in enumerate(stats):
        x = CONTENT_X + i * (col_w + gap)
        add_filled_rect(slide, x, CONTENT_Y, col_w, stats_h,
                        fill=C_WHITE, line=C_LIGHT_LINE)
        add_filled_rect(slide, x, CONTENT_Y, Inches(0.08), stats_h,
                        fill=C_ACCENT)

        add_text_box(
            slide, x + Inches(0.3), CONTENT_Y + Inches(0.2),
            col_w - Inches(0.5), Inches(0.9),
            text=num, size=38, bold=True, color=C_ACCENT,
        )
        add_text_box(
            slide, x + Inches(0.3), CONTENT_Y + Inches(1.05),
            col_w - Inches(0.5), Inches(0.35),
            text=lbl, size=FS_BODY, bold=True, color=C_PRIMARY,
        )
        add_text_box(
            slide, x + Inches(0.3), CONTENT_Y + Inches(1.4),
            col_w - Inches(0.5), Inches(0.7),
            text=sub, size=FS_BODY_SM, color=C_MUTED,
        )

    # Donji blok - poruka
    y2 = CONTENT_Y + stats_h + Inches(0.4)
    add_text_box(
        slide, CONTENT_X, y2, CONTENT_W, Inches(0.4),
        text="Sta nastavnik radi rucno?", size=FS_SUBTITLE, bold=True,
        color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, y2 + Inches(0.45), CONTENT_W,
        CONTENT_H - (y2 + Inches(0.45) - CONTENT_Y),
        [
            "Planira strukturu kursa - module, koncepte, ishode ucenja",
            "Pise objasnjenja, primere, definicije - usaglasava ih sa nivoom studenata",
            "Priprema praktiche demonstracije, cesto u formi video tutorijala",
            "Sastavlja pitanja na razlicitim kognitivnim nivoima - i tu nastaje uska grla",
        ],
        size=FS_BODY,
    )


def body_llm_kao_odgovor(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="LLM je otvorio vrata", size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.45), half_w, CONTENT_H - Inches(0.5),
        [
            "Generisanje sadrzaja iz prirodno-jezickog uputstva",
            "Razlaganje kompleksnih zadataka",
            "Brza iteracija i prilagodjavanje materijala",
            "Multimodalnost: tekst, kod, vizija",
        ],
        size=FS_BODY,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="...ali otvorio i nove probleme", size=FS_SUBTITLE, bold=True,
        color=C_ACCENT,
    )
    add_bullet_list(
        slide, right_x, CONTENT_Y + Inches(0.45), half_w, CONTENT_H - Inches(0.5),
        [
            "Halucinacije - izmisljene cinjenice",
            "Genericnost izlaza, malo veze sa konkretnim materijalom",
            "Pitanja na pogresnom kognitivnom nivou",
            "Nedostatak formalne strukture i ponovljivosti",
            "Tesko proveriti gde se sta oslonilo na izvor",
        ],
        size=FS_BODY,
    )


def body_vizija_pipeline(slide):
    # Pipeline kao 5 koraka
    steps = [
        ("Ulaz",          "Ishodi ucenja\nPDF materijal\nNL instrukcije"),
        ("Struktura",     "Bloom-driven\nmoduli i ishodi"),
        ("Ontologija",    "OWL/RDF\nkonceptualne relacije"),
        ("Generisanje",   "SOLO pitanja\nvideo demoi"),
        ("Izlaz",         "Strukturisan\nITS pipeline"),
    ]
    n = len(steps)
    gap = Inches(0.18)
    box_h = Inches(2.4)
    box_w = (CONTENT_W - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.2)

    for i, (title, sub) in enumerate(steps):
        x = CONTENT_X + i * (box_w + gap)
        # kartica
        add_filled_rect(slide, x, y, box_w, box_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        add_filled_rect(slide, x, y, box_w, Inches(0.1), fill=C_PRIMARY)

        # broj
        add_text_box(
            slide, x, y + Inches(0.2), box_w, Inches(0.4),
            text=f"0{i+1}", size=22, bold=True, color=C_ACCENT,
            align=PP_ALIGN.CENTER,
        )
        # naslov
        add_text_box(
            slide, x, y + Inches(0.7), box_w, Inches(0.4),
            text=title, size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
            align=PP_ALIGN.CENTER,
        )
        # podsadrzaj
        add_text_box(
            slide, x + Inches(0.1), y + Inches(1.2), box_w - Inches(0.2),
            box_h - Inches(1.3),
            text=sub, size=FS_BODY_SM, color=C_TEXT,
            align=PP_ALIGN.CENTER,
        )

        # strelica (osim za poslednji)
        if i < n - 1:
            arrow_x = x + box_w + Inches(0.02)
            arrow_y = y + box_h / 2 - Inches(0.1)
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW, arrow_x, arrow_y,
                gap - Inches(0.04), Inches(0.2),
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = C_ACCENT
            arrow.line.fill.background()

    # Ispod - kljucna poruka
    y2 = y + box_h + Inches(0.3)
    add_filled_rect(slide, CONTENT_X, y2, CONTENT_W, CONTENT_H - (y2 - CONTENT_Y) - Inches(0.1),
                    fill=C_CARD_BG, line=None)
    add_filled_rect(slide, CONTENT_X, y2, Inches(0.08),
                    CONTENT_H - (y2 - CONTENT_Y) - Inches(0.1), fill=C_ACCENT)
    add_text_box(
        slide, CONTENT_X + Inches(0.3), y2 + Inches(0.15),
        CONTENT_W - Inches(0.5), Inches(0.4),
        text="Vizija", size=FS_BODY_SM, bold=True, color=C_ACCENT,
    )
    add_text_box(
        slide, CONTENT_X + Inches(0.3), y2 + Inches(0.5),
        CONTENT_W - Inches(0.5), Inches(0.8),
        text="Od ishoda ucenja, automatski, do snimljene video demonstracije i SOLO kviza - bez rucne intervencije.",
        size=FS_BODY, color=C_TEXT,
    )


def body_tri_komponente(slide):
    cols = [
        ("Luka",
         "Struktura kursa",
         "Bloomova taksonomija",
         ["Hijerarhija: kurs -> modul -> koncept -> ishod -> aktivnost",
          "JSON serijalizacija sa rekurzivnom strukturom",
          "Bloomovi nivoi mapirani na tipove aktivnosti",
          "State-of-generation kontekst za koherentnost"]),
        ("Stefan",
         "Video demonstracije",
         "OWL ontologija + Computer Use",
         ["NL instrukcija -> Task / Step / Action",
          "OWL validacija plana pre izvrsenja",
          "SPARQL upiti vode izvrsavanje",
          "Vision (Qwen 2.5 VL) za UI detekciju",
          "FFmpeg snima MP4 tutorijal"]),
        ("Uros",
         "SOLO pitanja",
         "Ontology-grounded MCQ",
         ["4 SOLO nivoa: U / M / R / EA",
          "PS4 prompt template",
          "source_line citat iz PDF-a",
          "Ontology anchor iz ConceptRelationship",
          "Two-pass generisanje za EA"]),
    ]

    gap = Inches(0.25)
    col_w = (CONTENT_W - 2 * gap) / 3

    for i, (member, title, framework, points) in enumerate(cols):
        x = CONTENT_X + i * (col_w + gap)

        add_filled_rect(slide, x, CONTENT_Y, col_w, CONTENT_H - Inches(0.1),
                        fill=C_WHITE, line=C_LIGHT_LINE)
        # gornja akcent traka
        add_filled_rect(slide, x, CONTENT_Y, col_w, Inches(0.08), fill=C_ACCENT)

        pad = Inches(0.2)
        cur_y = CONTENT_Y + Inches(0.25)

        add_text_box(
            slide, x + pad, cur_y, col_w - 2 * pad, Inches(0.3),
            text=member.upper(), size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
        )
        cur_y += Inches(0.32)

        add_text_box(
            slide, x + pad, cur_y, col_w - 2 * pad, Inches(0.45),
            text=title, size=22, bold=True, color=C_PRIMARY,
        )
        cur_y += Inches(0.5)

        add_text_box(
            slide, x + pad, cur_y, col_w - 2 * pad, Inches(0.3),
            text=framework, size=FS_BODY_SM, italic=True, color=C_MUTED,
        )
        cur_y += Inches(0.4)

        add_horizontal_line(slide, x + pad, cur_y, Inches(0.5),
                            color=C_ACCENT, height=Emu(15000))
        cur_y += Inches(0.18)

        add_bullet_list(
            slide, x + pad, cur_y, col_w - 2 * pad,
            CONTENT_H - (cur_y - CONTENT_Y) - Inches(0.3),
            points, size=FS_BODY_SM,
        )


def body_kljucna_poruka_uvod(slide):
    # Centriran citat sa "formula" stilom
    add_text_box(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), CONTENT_W, Inches(0.5),
        text="Kljucna poruka prezentacije", size=FS_SECTION_LABEL, bold=True,
        color=C_ACCENT, align=PP_ALIGN.CENTER,
    )

    # Formula u kartici
    box_w = CONTENT_W * 0.85
    box_x = CONTENT_X + (CONTENT_W - box_w) / 2
    box_y = CONTENT_Y + Inches(1.2)
    box_h = Inches(2.2)

    add_filled_rect(slide, box_x, box_y, box_w, box_h,
                    fill=C_CARD_BG, line=C_LIGHT_LINE)
    add_horizontal_line(slide, box_x, box_y, box_w, color=C_PRIMARY,
                        height=Emu(30000))

    # Formula liniju po liniju
    add_text_box(
        slide, box_x, box_y + Inches(0.4), box_w, Inches(0.6),
        text="LLM  +  ontologija  +  pedagoska taksonomija",
        size=32, bold=True, color=C_PRIMARY,
        align=PP_ALIGN.CENTER,
    )
    add_text_box(
        slide, box_x, box_y + Inches(1.1), box_w, Inches(0.45),
        text="=", size=28, bold=True, color=C_ACCENT,
        align=PP_ALIGN.CENTER,
    )
    add_text_box(
        slide, box_x, box_y + Inches(1.55), box_w, Inches(0.5),
        text="pouzdano, pedagoski kalibrisano ITS okruzenje",
        size=22, italic=True, color=C_TEXT,
        align=PP_ALIGN.CENTER,
    )

    # Sub
    add_text_box(
        slide, CONTENT_X, box_y + box_h + Inches(0.25), CONTENT_W, Inches(0.5),
        text="vise od zbira svojih delova",
        size=FS_SUBTITLE, italic=True, color=C_MUTED,
        align=PP_ALIGN.CENTER,
    )


# --- SEGMENT A helperi -------------------------------------------------------

def body_bloom_pregled(slide):
    levels = [
        ("Pamtiti",      "Cinjenice, definicije, terminologija"),
        ("Razumeti",     "Objasnjenje, parafraziranje"),
        ("Primeniti",    "Resavanje analognih problema"),
        ("Analizirati",  "Razlaganje, poredjenje"),
        ("Evaluirati",   "Procena, kriticko misljenje"),
        ("Kreirati",     "Sinteza, novi rad"),
    ]
    half_w = CONTENT_W * 0.55
    n = len(levels)
    gap = Inches(0.08)
    row_h = (CONTENT_H - (n - 1) * gap) / n

    for i, (name, desc) in enumerate(levels):
        y = CONTENT_Y + i * (row_h + gap)
        add_filled_rect(slide, CONTENT_X, y, half_w, row_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        # broj badge
        add_filled_rect(slide, CONTENT_X, y, Inches(0.55), row_h, fill=C_PRIMARY)
        add_text_box(
            slide, CONTENT_X, y, Inches(0.55), row_h,
            text=f"L{i+1}", size=FS_BODY, bold=True, color=C_WHITE,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, CONTENT_X + Inches(0.7), y, half_w - Inches(0.8), Inches(0.35),
            text=name, size=FS_BODY, bold=True, color=C_PRIMARY,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, CONTENT_X + Inches(0.7), y + Inches(0.3),
            half_w - Inches(0.8), row_h - Inches(0.3),
            text=desc, size=FS_BODY_SM, color=C_MUTED,
        )

    # Desna polovina - placeholder za piramidu
    right_x = CONTENT_X + half_w + Inches(0.3)
    right_w = CONTENT_W - half_w - Inches(0.3)
    add_image_placeholder(slide, right_x, CONTENT_Y, right_w, CONTENT_H,
                          label="Bloomova piramida")


def body_solo_pregled(slide):
    levels = [
        ("Unistructural",
         "Jedna izolovana cinjenica",
         "Setiti se jednog elementa"),
        ("Multistructural",
         "Vise nepovezanih cinjenica",
         "Prepoznati skup elemenata"),
        ("Relational",
         "Veza izmedju koncepata",
         "Razumeti strukturu i odnose"),
        ("Extended Abstract",
         "Generalizacija i transfer",
         "Primeniti na novi kontekst"),
    ]
    half_w = CONTENT_W * 0.55
    n = len(levels)
    gap = Inches(0.12)
    row_h = (CONTENT_H - (n - 1) * gap) / n

    for i, (name, sub, desc) in enumerate(levels):
        y = CONTENT_Y + i * (row_h + gap)
        add_filled_rect(slide, CONTENT_X, y, half_w, row_h,
                        fill=C_WHITE, line=C_LIGHT_LINE)
        add_filled_rect(slide, CONTENT_X, y, Inches(0.1), row_h, fill=C_ACCENT)
        add_text_box(
            slide, CONTENT_X + Inches(0.25), y + Inches(0.1),
            half_w - Inches(0.4), Inches(0.32),
            text=name, size=FS_BODY, bold=True, color=C_PRIMARY,
        )
        add_text_box(
            slide, CONTENT_X + Inches(0.25), y + Inches(0.42),
            half_w - Inches(0.4), Inches(0.3),
            text=sub, size=FS_BODY_SM, italic=True, color=C_ACCENT,
        )
        add_text_box(
            slide, CONTENT_X + Inches(0.25), y + Inches(0.72),
            half_w - Inches(0.4), row_h - Inches(0.75),
            text=desc, size=FS_BODY_SM, color=C_MUTED,
        )

    right_x = CONTENT_X + half_w + Inches(0.3)
    right_w = CONTENT_W - half_w - Inches(0.3)
    add_image_placeholder(slide, right_x, CONTENT_Y, right_w, CONTENT_H,
                          label="SOLO taksonomija - vizuelni prikaz")


def body_bloom_vs_solo(slide):
    headers = ["Aspekt", "Bloom (revidirana)", "SOLO"]
    rows = [
        ["Tip hijerarhije",      "Tipovi misaonih operacija", "Strukturalna slozenost odgovora"],
        ["Broj nivoa",           "6 nivoa", "5 nivoa (ukljucujuci prestructural)"],
        ["Osnovna jedinica",     "Glagol akcije (znati, primeniti, ...)", "Broj i organizacija elemenata u odgovoru"],
        ["Tipicna upotreba",     "Planiranje aktivnosti i ishoda ucenja", "Evaluacija dubine ucenja"],
        ["Mi koristimo za",      "Strukturu kursa (Luka)", "Generisanje pitanja (Uros)"],
        ["Glavna prednost",      "Direktna veza sa nastavnim aktivnostima", "Direktna veza sa kognitivnom dubinom"],
    ]
    add_styled_table(slide, CONTENT_X, CONTENT_Y, CONTENT_W,
                     CONTENT_H - Inches(1.4), headers, rows,
                     col_widths=[1, 2, 2])

    # Donja poruka
    y2 = CONTENT_Y + CONTENT_H - Inches(1.2)
    add_filled_rect(slide, CONTENT_X, y2, CONTENT_W, Inches(1.0),
                    fill=C_CARD_BG, line=None)
    add_filled_rect(slide, CONTENT_X, y2, Inches(0.08), Inches(1.0), fill=C_ACCENT)
    add_text_box(
        slide, CONTENT_X + Inches(0.25), y2 + Inches(0.1),
        CONTENT_W - Inches(0.4), Inches(0.4),
        text="Zasto obe?", size=FS_BODY, bold=True, color=C_PRIMARY,
    )
    add_text_box(
        slide, CONTENT_X + Inches(0.25), y2 + Inches(0.45),
        CONTENT_W - Inches(0.4), Inches(0.6),
        text="Bloom upravlja STA generisemo (tipovi aktivnosti), SOLO upravlja KOLIKO DUBOKO se proverava (kognitivni nivoi pitanja).",
        size=FS_BODY_SM, color=C_TEXT,
    )


def body_luka_struktura(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    # Levo - opis pristupa
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Hijerarhijska dekompozicija kursa",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Ulaz", "strong"),
            "Tema kursa i opsti opis",
            "Lista ciljeva (ishoda ucenja)",
            "",
            ("LLM prolaz po nivou", "strong"),
            "Generise module iz teme",
            "Generise koncepte iz modula",
            "Generise ishode iz koncepata",
            "Generise aktivnosti iz ishoda",
            "",
            ("Bloomovo mapiranje", "strong"),
            "Svaka aktivnost dobija Bloomov nivo",
            "Aktivnosti se biraju kako bi pokrile sve nivoe",
        ],
        size=FS_BODY_SM,
    )

    # Desno - hijerarhijski dijagram + placeholder za sliku
    right_x = CONTENT_X + half_w + Inches(0.4)

    tree = [
        ("Kurs",       0, C_PRIMARY),
        ("Modul",      1, C_PRIMARY),
        ("Koncept",    2, C_ACCENT),
        ("Ishod",      3, C_ACCENT),
        ("Aktivnost",  4, C_MUTED),
    ]
    box_w = half_w - Inches(0.4)
    box_h = Inches(0.5)
    gap = Inches(0.12)
    y = CONTENT_Y + Inches(0.3)

    for i, (name, indent, color) in enumerate(tree):
        ind = Inches(0.4) * indent
        x = right_x + ind
        w = box_w - ind
        add_filled_rect(slide, x, y, w, box_h, fill=C_WHITE, line=C_LIGHT_LINE)
        add_filled_rect(slide, x, y, Inches(0.1), box_h, fill=color)
        add_text_box(
            slide, x + Inches(0.2), y, w - Inches(0.3), box_h,
            text=name, size=FS_BODY, bold=True, color=color,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        y += box_h + gap

    # caption ispod
    add_text_box(
        slide, right_x, y + Inches(0.1), box_w, Inches(0.5),
        text="Rekurzivna struktura, JSON serijalizacija.",
        size=FS_CAPTION, italic=True, color=C_MUTED,
    )


def body_luka_prompt_template(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    # Levo - opis
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Prompt template za strukturu kursa",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Sta ulazi u prompt", "strong"),
            "Tema i kontekst kursa",
            "Lista ciljeva ucenja (ishoda)",
            "Bloomov nivo koji se traci",
            "Stanje vec generisanog sadrzaja",
            "",
            ("Anti-redundancija", "strong"),
            "Prethodno generisane aktivnosti se ubacuju u prompt",
            "LLM se eksplicitno trazi da izbegne ponavljanje",
            "Daje koherentnost preko celog kursa",
        ],
        size=FS_BODY_SM,
    )

    # Desno - skica prompta (mock code)
    right_x = CONTENT_X + half_w + Inches(0.4)
    code = (
        "Role: ti si iskusan instructional designer.\n"
        "Tema kursa: {tema}\n"
        "Ciljevi ucenja:\n"
        "  - {cilj_1}\n"
        "  - {cilj_2}\n\n"
        "Generisi N aktivnosti za nivo: {bloom_level}.\n"
        "Vec generisane aktivnosti (ne ponavljaj):\n"
        "  - {prev_activity_1}\n"
        "  - {prev_activity_2}\n\n"
        "Format izlaza: JSON sa poljima\n"
        "  title, description, bloom_level, estimated_time."
    )
    add_code_block(slide, right_x, CONTENT_Y, half_w, CONTENT_H - Inches(0.5),
                   code, caption="Skica prompta - parametri u {}.")


def body_uros_solo_ps4(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="SOLO Quiz Generator (Uros)",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Ulaz", "strong"),
            "PDF lekcija (Srpski / Engleski, auto-detekcija)",
            "Parsirane sekcije i learning objekti",
            "Generisana ontologija pojmova",
            "",
            ("Izlaz", "strong"),
            "MCQ pitanja na 4 SOLO nivoa",
            "Svako pitanje: source_line citat iz PDF-a",
            "Relaciona i EA: ontology anchor u tags",
            "",
            ("Tech stack", "strong"),
            "Ollama qwen2.5:14b-instruct-q4 lokalno",
            "RDFLib + SPARQL, SQLite + Flask + React",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="PS4 prompt template",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    components = [
        ("Persona", "Role priming - LLM kao iskusan pedagog"),
        ("Pravila",  "SOLO definicija + worked example za nivo"),
        ("Putanja",  "Chain-of-thought scaffold pre odgovora"),
        ("Pitanja",  "Output schema + typed distractor strategies"),
    ]
    gap = Inches(0.12)
    box_h = (CONTENT_H - Inches(0.5) - 3 * gap) / 4
    y = CONTENT_Y + Inches(0.5)

    for label, desc in components:
        add_filled_rect(slide, right_x, y, half_w, box_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        add_filled_rect(slide, right_x, y, Inches(0.1), box_h, fill=C_ACCENT)
        add_text_box(
            slide, right_x + Inches(0.25), y + Inches(0.05),
            half_w - Inches(0.3), Inches(0.4),
            text=label, size=FS_BODY, bold=True, color=C_PRIMARY,
        )
        add_text_box(
            slide, right_x + Inches(0.25), y + Inches(0.4),
            half_w - Inches(0.3), box_h - Inches(0.4),
            text=desc, size=FS_BODY_SM, color=C_TEXT,
        )
        y += box_h + gap


def body_two_pass_ea(slide):
    half_w = CONTENT_W / 2 - Inches(0.25)

    # Levo - Pass 1 + Pass 2
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Pass 1: pitanje + odgovor + source",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, Inches(2.5),
        [
            "LLM generise pitanje za EA nivo",
            "Vraca tacan odgovor",
            "Vraca source_line - doslovni navod iz PDF-a",
            "Veza pitanje <-> izvor je eksplicitna",
        ],
        size=FS_BODY_SM,
    )

    y2 = CONTENT_Y + Inches(3.2)
    add_text_box(
        slide, CONTENT_X, y2, half_w, Inches(0.4),
        text="Pass 2: typed distractors (predictive prompting)",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, y2 + Inches(0.5), half_w, CONTENT_H - (y2 - CONTENT_Y) - Inches(0.6),
        [
            "Echo: ponovi pitanje i tacan odgovor",
            "Generisi 3 tipizirana distraktora",
            "Svaki ima zadati 'tip greske' (overgeneralizacija, povrsno citanje, ...)",
            "Strategije se biraju iz tabele po SOLO nivou",
        ],
        size=FS_BODY_SM,
    )

    # Desno - dijagram dva prolaza
    right_x = CONTENT_X + half_w + Inches(0.5)
    rw = CONTENT_W - half_w - Inches(0.5)

    # Pass 1 box
    box_h = Inches(2.4)
    p1 = add_filled_rect(slide, right_x, CONTENT_Y, rw, box_h,
                         fill=C_CARD_BG, line=C_LIGHT_LINE)
    add_filled_rect(slide, right_x, CONTENT_Y, Inches(0.12), box_h, fill=C_PRIMARY)
    add_text_box(
        slide, right_x + Inches(0.25), CONTENT_Y + Inches(0.15),
        rw - Inches(0.3), Inches(0.4),
        text="PASS 1", size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
    )
    add_text_box(
        slide, right_x + Inches(0.25), CONTENT_Y + Inches(0.5),
        rw - Inches(0.3), Inches(0.5),
        text="LLM -> { pitanje, tacan_odgovor, source_line }",
        size=FS_BODY, bold=True, color=C_PRIMARY,
    )
    add_text_box(
        slide, right_x + Inches(0.25), CONTENT_Y + Inches(1.05),
        rw - Inches(0.3), Inches(1.2),
        text="Fokus: pedagoska valjanost pitanja i grounding u izvoru. Distraktori se NE traze u ovom prolazu.",
        size=FS_BODY_SM, color=C_TEXT,
    )

    # Strelica
    arrow_y = CONTENT_Y + box_h + Inches(0.05)
    arrow = slide.shapes.add_shape(
        MSO_SHAPE.DOWN_ARROW, right_x + rw / 2 - Inches(0.15), arrow_y,
        Inches(0.3), Inches(0.4),
    )
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = C_ACCENT
    arrow.line.fill.background()

    # Pass 2 box
    p2_y = arrow_y + Inches(0.5)
    p2_h = CONTENT_H - (p2_y - CONTENT_Y) - Inches(0.2)
    add_filled_rect(slide, right_x, p2_y, rw, p2_h,
                    fill=C_CARD_BG, line=C_LIGHT_LINE)
    add_filled_rect(slide, right_x, p2_y, Inches(0.12), p2_h, fill=C_PRIMARY)
    add_text_box(
        slide, right_x + Inches(0.25), p2_y + Inches(0.15),
        rw - Inches(0.3), Inches(0.4),
        text="PASS 2", size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
    )
    add_text_box(
        slide, right_x + Inches(0.25), p2_y + Inches(0.5),
        rw - Inches(0.3), Inches(0.5),
        text="LLM -> { 3 typed distractor-a }",
        size=FS_BODY, bold=True, color=C_PRIMARY,
    )
    add_text_box(
        slide, right_x + Inches(0.25), p2_y + Inches(1.05),
        rw - Inches(0.3), p2_h - Inches(1.1),
        text="Echo prompt: LLM ponovi pitanje, pa generise distraktore po tipiziranim strategijama greske. Smanjuje 'lazne' distraktore.",
        size=FS_BODY_SM, color=C_TEXT,
    )


def body_source_line(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Source-line citation kao anti-halucinacija",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Sta", "strong"),
            "Uz svako pitanje, LLM mora vratiti doslovni navod iz PDF-a",
            "Navod opravdava tacan odgovor",
            "",
            ("Zasto", "strong"),
            "Ako navod ne postoji u izvoru = signal halucinacije",
            "Lako pregledati u UI: pitanje + citirani odlomak",
            "Recenzent (ili nastavnik) vidi gde je 'istina'",
            "",
            ("Kako", "strong"),
            "JSON polje source_line u prompt schemi",
            "Cuva se u tabeli questions",
            "QuestionBank UI prikazuje source_line uz svako pitanje",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Primer JSON izlaza",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    code = (
        '{\n'
        '  "level": "relational",\n'
        '  "question": "Kako se memorijska\\n'
        '   hijerarhija odnosi prema brzini\\n'
        '   pristupa i kapacitetu?",\n'
        '  "correct": "Brzina opada, a\\n'
        '   kapacitet raste niz hijerarhiju.",\n'
        '  "distractors": [...],\n'
        '  "source_line": "Memorijska\\n'
        '   hijerarhija je organizovana tako\\n'
        '   da brze memorije imaju manji\\n'
        '   kapacitet.",\n'
        '  "tags": { "ontology_anchor": {\\n'
        '     "source": "RAM",\\n'
        '     "target": "Cache",\\n'
        '     "type": "is_faster_than"\\n'
        '  } }\n'
        '}'
    )
    add_code_block(slide, right_x, CONTENT_Y + Inches(0.5), half_w,
                   CONTENT_H - Inches(0.6), code,
                   caption="Pojednostavljeni primer odgovora LLM-a.")


def body_stefan_dekompozicija(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Computer Use Video Instructions (Stefan)",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Ulaz", "strong"),
            "Prirodno-jezicka instrukcija praktickog zadatka",
            "Npr. \"Kreiraj C# konzolni projekat u Visual Studio-u\"",
            "",
            ("Izlaz", "strong"),
            "Snimljeni MP4 video tutorijal",
            "OWL fajl izvrsnog plana (auditable)",
            "",
            ("Hijerarhija razgradnje", "strong"),
            "Task: kompletan zadatak (zatvorena celina)",
            "Step: atomicna akcija (otvori, klikni, otkucaj)",
            "Action: tip akcije iz fiksne ontologije akcija",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Task / Step / Action hijerarhija",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    # 3 nivoa kartice
    levels = [
        ("Task", "Open VS, create C# Console app, write Hello World"),
        ("Step", "1. open_application Visual Studio   ->   2. wait   ->   3. click Create new project   ->   ..."),
        ("Action", "open_application | click | type_text | key_press | wait | capture_screen"),
    ]
    gap = Inches(0.15)
    box_h = (CONTENT_H - Inches(0.5) - 2 * gap) / 3
    y = CONTENT_Y + Inches(0.5)

    for i, (lbl, content) in enumerate(levels):
        add_filled_rect(slide, right_x, y, half_w, box_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        add_filled_rect(slide, right_x, y, Inches(0.1), box_h, fill=C_ACCENT)
        add_text_box(
            slide, right_x + Inches(0.25), y + Inches(0.1),
            half_w - Inches(0.3), Inches(0.35),
            text=lbl, size=FS_BODY, bold=True, color=C_PRIMARY,
        )
        add_text_box(
            slide, right_x + Inches(0.25), y + Inches(0.45),
            half_w - Inches(0.3), box_h - Inches(0.45),
            text=content, size=FS_BODY_SM, color=C_TEXT,
        )
        y += box_h + gap


def body_prompt_vs_finetune(slide):
    headers = ["", "Prompt engineering", "Fine-tuning"]
    rows = [
        ["Trosak razvoja",     "Nizak - samo prompt",                 "Visok - dataset + trening"],
        ["Trosak izvrsenja",   "Veci prompt = veci token cost",       "Manji prompt, ali specijalizovan model"],
        ["Brzina iteracije",   "Sekunde - promenis prompt",           "Sati / dani - rerun trening"],
        ["Domen-specificnost", "Ogranicena, oslanja se na bazni model","Visoka, prilagodjen domenu"],
        ["Kontrola halucinacija","Posredna (CoT, schema, grounding)","Direktna (kvalitet treninga)"],
        ["Nas izbor",          "Da - PS4 + ontology grounding",       "Ne, za sada"],
    ]
    add_styled_table(slide, CONTENT_X, CONTENT_Y, CONTENT_W,
                     CONTENT_H - Inches(1.4), headers, rows,
                     col_widths=[1.2, 1.7, 1.7])

    # Donja poruka
    y2 = CONTENT_Y + CONTENT_H - Inches(1.2)
    add_filled_rect(slide, CONTENT_X, y2, CONTENT_W, Inches(1.0),
                    fill=C_CARD_BG, line=None)
    add_filled_rect(slide, CONTENT_X, y2, Inches(0.08), Inches(1.0), fill=C_ACCENT)
    add_text_box(
        slide, CONTENT_X + Inches(0.25), y2 + Inches(0.1),
        CONTENT_W - Inches(0.4), Inches(0.4),
        text="Zasto prompt engineering, ne fine-tune?", size=FS_BODY, bold=True,
        color=C_PRIMARY,
    )
    add_text_box(
        slide, CONTENT_X + Inches(0.25), y2 + Inches(0.45),
        CONTENT_W - Inches(0.4), Inches(0.6),
        text="Bez stabilnog domenskog dataset-a za SOLO/Bloom u programiranju, prompt + ontology grounding daje brzu, transparentnu i jeftiniju iteraciju.",
        size=FS_BODY_SM, color=C_TEXT,
    )


def body_local_vs_cloud(slide):
    headers = ["", "Lokalni: Ollama Qwen 2.5 14B", "Cloud: Groq Llama 3.3 70B", "Cloud Vision: Qwen 2.5 VL 72B"]
    rows = [
        ["Koristi se za", "Uros - SOLO pitanja", "Stefan - planiranje koraka", "Stefan - UI detekcija"],
        ["Trosak po pozivu", "0 (lokalno)", "Mali (per token)", "Srednji (vision)"],
        ["Brzina", "Sporije, zavisi od GPU-a", "Vrlo brzo (Groq accel.)", "Brza"],
        ["Privatnost",  "Visoka (sve lokalno)", "Niska - cloud", "Niska - cloud"],
        ["Kvalitet",    "Solidan za EA, slabiji za nuansiranja", "Veoma dobar", "Najbolji za vision"],
        ["Cache",       "SHA-256 keys u SQLite",  "n/a u ovom projektu", "n/a"],
    ]
    add_styled_table(slide, CONTENT_X, CONTENT_Y, CONTENT_W,
                     CONTENT_H - Inches(0.4), headers, rows,
                     col_widths=[1.0, 1.6, 1.6, 1.6])


def body_seg_a_kljucna(slide):
    add_text_box(
        slide, CONTENT_X, CONTENT_Y + Inches(0.4), CONTENT_W, Inches(0.4),
        text="Kljucna poruka segmenta A",
        size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
        align=PP_ALIGN.CENTER,
    )

    box_w = CONTENT_W * 0.85
    box_x = CONTENT_X + (CONTENT_W - box_w) / 2
    box_y = CONTENT_Y + Inches(1.1)
    box_h = Inches(2.6)

    add_filled_rect(slide, box_x, box_y, box_w, box_h,
                    fill=C_CARD_BG, line=C_LIGHT_LINE)
    add_horizontal_line(slide, box_x, box_y, box_w, color=C_PRIMARY,
                        height=Emu(30000))

    add_text_box(
        slide, box_x + Inches(0.4), box_y + Inches(0.5), box_w - Inches(0.8),
        Inches(1.6),
        text=("Hijerarhijska dekompozicija + formalna pedagoska "
              "taksonomija ugradjena u prompt transformisu LLM iz "
              "generickog generatora teksta u pedagoski kalibrisan alat."),
        size=22, italic=True, color=C_PRIMARY, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    # ispod - 3 mini-poente
    yy = box_y + box_h + Inches(0.4)
    pts = [
        ("Bloom",       "drives the WHAT - tipovi aktivnosti"),
        ("SOLO",        "drives the HOW DEEP - kognitivne razlike"),
        ("PS4 / dekompozicija", "drives the HOW - struktura prompta"),
    ]
    gap = Inches(0.2)
    col_w = (CONTENT_W - 2 * gap) / 3
    for i, (h, b) in enumerate(pts):
        x = CONTENT_X + i * (col_w + gap)
        add_text_box(slide, x, yy, col_w, Inches(0.4),
                     text=h, size=FS_BODY, bold=True, color=C_ACCENT,
                     align=PP_ALIGN.CENTER)
        add_text_box(slide, x, yy + Inches(0.4), col_w, Inches(0.6),
                     text=b, size=FS_BODY_SM, color=C_MUTED,
                     align=PP_ALIGN.CENTER)


# --- SEGMENT B helperi -------------------------------------------------------

def body_owl_uvod(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="OWL (Web Ontology Language)",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            "W3C standard za formalne ontologije",
            "Definise klase, instance, relacije, ogranicenja",
            "Citljiv mas i mas (RDF/XML, Turtle, OWL/XML)",
            "Reasoning - moze da izvodi nove cinjenice",
            "",
            ("Mi koristimo", "strong"),
            "Stefan: OWL u Turtle (computer_use.ttl)",
            "Uros: seed OWL + RDF/XML eksport za Protege",
            "Luka: konceptualni model serijalizovan u JSON",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    code = (
        "@prefix cu: <http://example.org/computer-use#> .\n"
        "@prefix owl: <.../owl#> .\n"
        "@prefix rdf: <.../rdf-syntax-ns#> .\n\n"
        "cu:Task a owl:Class .\n"
        "cu:Step a owl:Class .\n"
        "cu:hasStep a owl:ObjectProperty ;\n"
        "    rdfs:domain cu:Task ;\n"
        "    rdfs:range  cu:Step .\n"
        "cu:stepOrder a owl:DatatypeProperty ;\n"
        "    rdfs:domain cu:Step ;\n"
        "    rdfs:range  xsd:integer ."
    )
    add_code_block(slide, right_x, CONTENT_Y + Inches(0.5), half_w,
                   CONTENT_H - Inches(0.6), code,
                   caption="Mini-isecak iz computer_use.ttl (Stefan).")


def body_rdf_sparql(slide):
    # Levo - RDF
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="RDF - tripleti",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, Inches(2.5),
        [
            "Resource Description Framework",
            "Sve cinjenice = (subject, predicate, object)",
            "Primer: (RAM, is_faster_than, Disk)",
            "OWL je TBox, RDF je ABox - klase vs instance",
        ],
        size=FS_BODY_SM,
    )

    code_rdf = (
        ":RAM   rdf:type  :MemoryDevice .\n"
        ":Disk  rdf:type  :MemoryDevice .\n"
        ":RAM   :isFasterThan  :Disk ."
    )
    add_code_block(
        slide, CONTENT_X, CONTENT_Y + Inches(3.1),
        half_w, Inches(1.3), code_rdf,
        caption="3 RDF tripleta = mini graf.",
    )

    # Desno - SPARQL
    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="SPARQL - upitni jezik nad RDF/OWL",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    code_sparql = (
        "PREFIX cu: <http://example.org/computer-use#>\n\n"
        "SELECT ?step ?order ?action ?target\n"
        "WHERE {\n"
        "  ?task a cu:Task .\n"
        "  ?task cu:hasStep ?step .\n"
        "  ?step cu:stepOrder ?order ;\n"
        "        cu:hasAction ?action ;\n"
        "        cu:targetName ?target .\n"
        "}\n"
        "ORDER BY ?order"
    )
    add_code_block(
        slide, right_x, CONTENT_Y + Inches(0.5), half_w,
        CONTENT_H - Inches(0.7), code_sparql,
        caption="Stefanov upit za citanje koraka iz OWL plana.",
    )


def body_luka_json_primer(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="JSON serijalizacija strukture kursa",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            "Hijerarhijska, rekurzivna struktura",
            "Svaki cvor ima id, name, bloom_level, children",
            "Aktivnosti su listovi u stablu",
            "Lako se mapira u Protege / OWL graph",
            "",
            ("Zasto JSON, a ne odmah OWL?", "strong"),
            "Lakse za LLM da generise validan JSON",
            "Lakse za UI prikaz i editovanje",
            "Mapiranje JSON -> OWL je determinizam, lako se radi unazad",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    code_json = (
        '{\n'
        '  "course": "Operativni sistemi",\n'
        '  "modules": [\n'
        '    {\n'
        '      "name": "Procesi i niti",\n'
        '      "concepts": [\n'
        '        {\n'
        '          "name": "Proces",\n'
        '          "outcomes": [\n'
        '            {\n'
        '              "name": "Razumeti stanja procesa",\n'
        '              "bloom": "Razumeti",\n'
        '              "activities": [...]\n'
        '            }\n'
        '          ]\n'
        '        }\n'
        '      ]\n'
        '    }\n'
        '  ]\n'
        '}'
    )
    add_code_block(
        slide, right_x, CONTENT_Y + Inches(0.5), half_w,
        CONTENT_H - Inches(0.6), code_json,
        caption="Primer Lukinog JSON izlaza (pojednostavljeno).",
    )


def body_uros_ontologija(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="ConceptRelationship ontologija",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Modeluje", "strong"),
            "Relacije izmedju learning objekata (LO) jedne lekcije",
            "Edge = (source, target, type, description)",
            "Cuva se u SQL tabeli i eksportuje u Turtle / OWL",
            "",
            ("Multi-pass ekstrakcija", "strong"),
            "5 LLM prolaza, jedan po tipu relacije",
            "Smanjuje hallucinated 'prerequisite' lance",
            "Konzervativan fallback samo ako LLM vrati nulu",
            "",
            ("Tipovi relacija", "strong"),
            "prerequisite, example_of, depends_on, ...",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Tipovi relacija u ontologiji",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    rels = [
        ("prerequisite",  "A se mora razumeti pre B"),
        ("example_of",    "A je primer pojma B"),
        ("depends_on",    "A operativno zavisi od B"),
        ("similar_to",    "A i B dele kljucnu karakteristiku"),
        ("part_of",       "A je deo strukture B"),
    ]
    gap = Inches(0.1)
    n = len(rels)
    box_h = (CONTENT_H - Inches(0.5) - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.5)
    for name, desc in rels:
        add_filled_rect(slide, right_x, y, half_w, box_h,
                        fill=C_WHITE, line=C_LIGHT_LINE)
        add_filled_rect(slide, right_x, y, Inches(0.1), box_h, fill=C_ACCENT)
        add_text_box(
            slide, right_x + Inches(0.25), y, Inches(2.0), box_h,
            text=name, size=FS_BODY, bold=True, color=C_PRIMARY,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, right_x + Inches(2.3), y, half_w - Inches(2.4), box_h,
            text=desc, size=FS_BODY_SM, color=C_MUTED,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        y += box_h + gap


def body_ontology_anchor(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Ontology anchor kao kontekst pitanja",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            "Za relacionalna i EA pitanja",
            "SPARQL bira jedan ConceptRelationship red",
            "Anchor postaje deo prompta",
            "Pitanje se gradi oko jedne konkretne veze",
            "",
            ("Posledice", "strong"),
            "Pitanje je trasabilno: zna se na koju vezu se odnosi",
            "Lakse za dedup po (anchor, normalized_correct_answer)",
            "Lakse za reviziju nastavnika",
            "",
            ("Tags polje u DB", "strong"),
            "tags.ontology_anchor = {source, target, type, description}",
            "tags.distractor_strategies = [...]",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    code = (
        "PREFIX :  <http://.../ontology#>\n\n"
        "SELECT ?src ?type ?tgt ?desc\n"
        "WHERE {\n"
        "  ?rel a :ConceptRelationship .\n"
        "  ?rel :source ?src ;\n"
        "       :target ?tgt ;\n"
        "       :relType ?type ;\n"
        "       :description ?desc .\n"
        "  ?src :inLesson <lesson_42> .\n"
        "}\n"
        "ORDER BY RAND()\n"
        "LIMIT 1"
    )
    add_code_block(
        slide, right_x, CONTENT_Y + Inches(0.5), half_w,
        CONTENT_H - Inches(0.6), code,
        caption="Biranje jedne veze kao anchor-a za prompt.",
    )


def body_stefan_owl(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)

    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Stefan - OWL ontologija plana",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Klase", "strong"),
            "Task - kompletan automatizacioni zadatak",
            "Step - atomicna akcija u zadatku",
            "Action - tip akcije iz fiksne liste",
            "UIElement, Application, ExecutionState",
            "",
            ("Object properties", "strong"),
            "hasStep, hasAction, hasTarget",
            "nextStep / previousStep, requiresApplication",
            "",
            ("Data properties", "strong"),
            "stepOrder, stepDescription, inputValue",
            "targetName, waitDuration, expectedResult",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Primer instance Step-a u OWL/RDF",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    code = (
        "<cu:Step rdf:about=\".../Step_1\">\n"
        "  <cu:stepOrder>1</cu:stepOrder>\n"
        "  <cu:hasAction rdf:resource=\"...open_application\"/>\n"
        "  <cu:targetName>Visual Studio</cu:targetName>\n"
        "  <cu:stepDescription>\n"
        "    Start Visual Studio\n"
        "  </cu:stepDescription>\n"
        "  <cu:expectedResult>\n"
        "    Visual Studio is opened\n"
        "  </cu:expectedResult>\n"
        "  <cu:hasState\n"
        "     rdf:resource=\"...PendingState\"/>\n"
        "</cu:Step>"
    )
    add_code_block(
        slide, right_x, CONTENT_Y + Inches(0.5), half_w,
        CONTENT_H - Inches(0.6), code,
        caption="Skiniran iz computer_use.ttl (Stefanov repo).",
    )


def body_validacija_pravila(slide):
    headers = ["Pravilo", "Opis", "Zasto"]
    rows = [
        ["Action validation",
         "Samo akcije iz ontologije su dozvoljene",
         "LLM ne moze izmisliti akciju koja ne moze biti izvrsena"],
        ["Sequence validation",
         "open_application mora prethoditi interakciji",
         "Sprecava klikove na nepostojeci prozor"],
        ["Wait recommendation",
         "wait posle open_application",
         "UI se ucitava - bez pauze klikovi promasuju"],
        ["Parameter validation",
         "type_text zahteva value, key_press zahteva ime tipke",
         "Hvata strukturne greske LLM-a u JSON izlazu"],
        ["Target validation",
         "click mora imati specifican target",
         "Bez targeta nema sta da se klikne"],
    ]
    add_styled_table(slide, CONTENT_X, CONTENT_Y, CONTENT_W,
                     CONTENT_H - Inches(0.4), headers, rows,
                     col_widths=[1.2, 2.2, 2.2])


def body_anti_halucinacije(slide):
    cols = [
        ("Uros",
         "source_line citat",
         ["LLM mora vratiti doslovni navod iz PDF-a",
          "Citat opravdava tacan odgovor",
          "Ako navod ne postoji = signal halucinacije",
          "Recenzent vidi gde je tacan odgovor 'usidren'"]),
        ("Luka",
         "state-of-generation kontekst",
         ["Vec generisane aktivnosti ubacuju se u prompt",
          "LLM eksplicitno cita 'ne ponavljaj ovo'",
          "Daje koherentnost kroz ceo kurs",
          "Smanjuje genericne, ponavljajuce ishode"]),
        ("Stefan",
         "SPARQL odbacuje nevalidne sekvence",
         ["Pre izvrsavanja - ontoloska provera",
          "Pravila: open_app pre interakcije, wait posle open_app",
          "type_text zahteva value, ...",
          "Nevalidan plan = stop, ne pravi se video"]),
    ]
    gap = Inches(0.25)
    col_w = (CONTENT_W - 2 * gap) / 3

    for i, (who, title, points) in enumerate(cols):
        x = CONTENT_X + i * (col_w + gap)
        add_filled_rect(slide, x, CONTENT_Y, col_w, CONTENT_H - Inches(0.1),
                        fill=C_WHITE, line=C_LIGHT_LINE)
        add_filled_rect(slide, x, CONTENT_Y, col_w, Inches(0.08), fill=C_ACCENT)

        pad = Inches(0.2)
        cur_y = CONTENT_Y + Inches(0.25)
        add_text_box(slide, x + pad, cur_y, col_w - 2 * pad, Inches(0.3),
                     text=who.upper(), size=FS_SECTION_LABEL, bold=True,
                     color=C_ACCENT)
        cur_y += Inches(0.32)
        add_text_box(slide, x + pad, cur_y, col_w - 2 * pad, Inches(0.5),
                     text=title, size=20, bold=True, color=C_PRIMARY)
        cur_y += Inches(0.55)
        add_horizontal_line(slide, x + pad, cur_y, Inches(0.5),
                            color=C_ACCENT, height=Emu(15000))
        cur_y += Inches(0.15)
        add_bullet_list(slide, x + pad, cur_y, col_w - 2 * pad,
                        CONTENT_H - (cur_y - CONTENT_Y) - Inches(0.3),
                        points, size=FS_BODY_SM)


def body_seg_b_kljucna(slide):
    add_text_box(
        slide, CONTENT_X, CONTENT_Y + Inches(0.4), CONTENT_W, Inches(0.4),
        text="Kljucna poruka segmenta B",
        size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
        align=PP_ALIGN.CENTER,
    )
    box_w = CONTENT_W * 0.85
    box_x = CONTENT_X + (CONTENT_W - box_w) / 2
    box_y = CONTENT_Y + Inches(1.1)
    box_h = Inches(2.6)
    add_filled_rect(slide, box_x, box_y, box_w, box_h,
                    fill=C_CARD_BG, line=C_LIGHT_LINE)
    add_horizontal_line(slide, box_x, box_y, box_w, color=C_PRIMARY,
                        height=Emu(30000))
    add_text_box(
        slide, box_x + Inches(0.4), box_y + Inches(0.4), box_w - Inches(0.8),
        box_h - Inches(0.8),
        text=("Ontologija je ugovor izmedju LLM-a i sistema: ono sto LLM "
              "proizvede mora se mapirati na klase i relacije. "
              "To filtrira halucinacije i omogucava ponovnu upotrebu sadrzaja."),
        size=22, italic=True, color=C_PRIMARY, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )

    yy = box_y + box_h + Inches(0.4)
    pts = [
        ("OWL/RDF",  "definicija strukture i instanci"),
        ("SPARQL",   "kontrola toka i upita"),
        ("Anchor",   "vezivanje pitanja za izvor"),
    ]
    gap = Inches(0.2)
    col_w = (CONTENT_W - 2 * gap) / 3
    for i, (h, b) in enumerate(pts):
        x = CONTENT_X + i * (col_w + gap)
        add_text_box(slide, x, yy, col_w, Inches(0.4),
                     text=h, size=FS_BODY, bold=True, color=C_ACCENT,
                     align=PP_ALIGN.CENTER)
        add_text_box(slide, x, yy + Inches(0.4), col_w, Inches(0.6),
                     text=b, size=FS_BODY_SM, color=C_MUTED,
                     align=PP_ALIGN.CENTER)


# --- SEGMENT C helperi -------------------------------------------------------

def body_pilot_metodologija(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Pilot evaluacija SOLO pitanja (Uros)",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Cilj evaluacije", "strong"),
            "Da li je kvalitet konstantan kroz SOLO nivoe?",
            "Da li distraktori postaju lazni / slabi na visim nivoima?",
            "",
            ("Metodologija", "strong"),
            "Vise PDF lekcija iz razlicitih domena",
            "N pitanja po nivou, rucna anotacija",
            "Sudije: clanovi tima + nezavisni recenzent",
            "Metrike: valjanost, kvalitet distraktora, source-line tacnost",
            "",
            ("Format izlaza", "strong"),
            "Skor na skali 1-5 po dimenziji",
            "Komentari za kvalitativnu analizu",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_image_placeholder(
        slide, right_x, CONTENT_Y, half_w, CONTENT_H,
        label="Tabela metrike po nivou / domenu",
    )


def body_kvalitet_po_nivou(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Rezultati - kvalitet po SOLO nivou",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    # Mali "graf" simuliran preko bar-charta od pravougaonika
    levels = [("Unistructural", 0.92),
              ("Multistructural", 0.85),
              ("Relational", 0.70),
              ("Extended Abstract", 0.55)]
    gap = Inches(0.18)
    n = len(levels)
    chart_y = CONTENT_Y + Inches(0.5)
    chart_h = CONTENT_H - Inches(1.2)
    row_h = (chart_h - (n - 1) * gap) / n

    max_bar_w = half_w - Inches(1.6)

    for i, (lvl, val) in enumerate(levels):
        y = chart_y + i * (row_h + gap)
        add_text_box(
            slide, CONTENT_X, y, Inches(1.5), row_h,
            text=lvl, size=FS_BODY_SM, bold=True, color=C_PRIMARY,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        # bar
        bar_w = max_bar_w * val
        bx = CONTENT_X + Inches(1.55)
        add_filled_rect(slide, bx, y + row_h * 0.15, max_bar_w,
                        row_h * 0.7, fill=C_LIGHT_LINE, line=None)
        add_filled_rect(slide, bx, y + row_h * 0.15, bar_w,
                        row_h * 0.7, fill=C_ACCENT, line=None)
        # value
        add_text_box(
            slide, bx + bar_w + Inches(0.1), y, Inches(0.7), row_h,
            text=f"{int(val*100)}%", size=FS_BODY, bold=True,
            color=C_PRIMARY, anchor=MSO_ANCHOR.MIDDLE,
        )

    # Caption ispod
    add_text_box(
        slide, CONTENT_X, chart_y + chart_h + Inches(0.1), half_w, Inches(0.5),
        text="Hipoteticne brojke iz pilot evaluacije - tacne brojke izlazu u radu.",
        size=FS_CAPTION, italic=True, color=C_MUTED,
    )

    # Desno - sta to znaci
    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Sta vidimo u rezultatima",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, right_x, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            "Validnost odgovora pada sa porastom SOLO nivoa",
            "Distraktori postaju manje uverljivi",
            "EA: cesto se 'distraktor' moze odbraniti kao tacan",
            "",
            ("Posledica za sistem", "strong"),
            "Two-pass + typed distractor strategije pomazu, ali ne resavaju potpuno",
            "Ontology anchor stabilizuje relacionalna pitanja",
            "Na EA nivou - obavezna recenzija nastavnika",
        ],
        size=FS_BODY_SM,
    )


def body_primer_solo_pitanja(slide):
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, CONTENT_W, Inches(0.4),
        text="Primeri pitanja po SOLO nivoima",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    levels = [
        ("Unistructural",
         "Koja struktura procesa cuva PID?",
         "PCB - Process Control Block"),
        ("Multistructural",
         "Koja tri stanja procesa najcesce prepoznajemo?",
         "Ready, Running, Blocked"),
        ("Relational",
         "Kako kontekstno prebacivanje uticne na throughput sistema?",
         "Povecava overhead, smanjuje koristan rad CPU-a."),
        ("Extended Abstract",
         "Kako biste prilagodili scheduling algoritam za real-time sistem sa zagarantovanim deadline-om?",
         "EDF / RM, prioritet po roku, preemption, ..."),
    ]
    gap = Inches(0.15)
    n = len(levels)
    box_h = (CONTENT_H - Inches(0.5) - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.5)

    for lvl, q, a in levels:
        add_filled_rect(slide, CONTENT_X, y, CONTENT_W, box_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        add_filled_rect(slide, CONTENT_X, y, Inches(0.12), box_h, fill=C_ACCENT)

        # level label
        add_text_box(
            slide, CONTENT_X + Inches(0.25), y + Inches(0.1),
            Inches(2.2), Inches(0.35),
            text=lvl.upper(), size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
        )
        # question
        add_text_box(
            slide, CONTENT_X + Inches(2.6), y + Inches(0.1),
            CONTENT_W - Inches(2.8), Inches(0.55),
            text=f"Q: {q}", size=FS_BODY_SM, bold=True, color=C_PRIMARY,
        )
        # answer
        add_text_box(
            slide, CONTENT_X + Inches(2.6), y + Inches(0.65),
            CONTENT_W - Inches(2.8), box_h - Inches(0.7),
            text=f"A: {a}", size=FS_BODY_SM, italic=True, color=C_TEXT,
        )

        y += box_h + gap


def body_luka_bloom_pokrivenost(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Bloomova pokrivenost (Luka)",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Metrika", "strong"),
            "Da li su zastupljeni svi Bloomovi nivoi?",
            "Koliko aktivnosti po nivou?",
            "Da li su ishodi mapirani na vise od jednog nivoa?",
            "",
            ("Pristup", "strong"),
            "LLM eksplicitno trazi po svakom nivou",
            "Sistem ne prelazi na sledeci nivo dok prethodni nije mapiran",
            "Dedup po (ishod, aktivnost_tip)",
            "",
            ("Nalaz", "strong"),
            "Najbolji rezultati: Pamtiti i Razumeti",
            "Slabiji: Evaluirati i Kreirati - cesto se 'dotiknu' samo povrsno",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_image_placeholder(
        slide, right_x, CONTENT_Y, half_w, CONTENT_H,
        label="Bar chart - broj aktivnosti po Bloomovom nivou",
    )


def body_stefan_ui_robust(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="UI detekcija - sta sve moze da pukne",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Vision model: Qwen 2.5 VL", "strong"),
            "Trazi UI element po opisu (lokacija + box)",
            "",
            ("Ogranicenja", "strong"),
            "Non-standard UI frameworks (custom controls)",
            "Brzo menjajuce interfejse (animacije, lazy load)",
            "Nizak kontrast / non-English labele",
            "Multi-monitor setup - samo primarni",
            "",
            ("Posledice", "strong"),
            "Step fail - sistem zaustavlja izvrsenje",
            "Video se zavrsava 'kratko' (do tacke prekida)",
            "Logujemo gde je tacno puklo - olaksava debug",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Testirane aplikacije",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    apps = [
        ("Visual Studio 2022", "Stabilno"),
        ("VS Code",            "Stabilno"),
        ("Chrome / Firefox / Edge / Opera", "Stabilno (web cases)"),
        ("Notepad / Notepad++","Stabilno"),
        ("Custom WinForms",    "Nestabilno"),
        ("Web app sa lazy UI", "Nestabilno"),
    ]
    n = len(apps)
    gap = Inches(0.08)
    row_h = (CONTENT_H - Inches(0.5) - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.5)

    for app, status in apps:
        color = C_PRIMARY if "Stabilno" in status else C_ACCENT
        add_filled_rect(slide, right_x, y, half_w, row_h,
                        fill=C_WHITE, line=C_LIGHT_LINE)
        add_filled_rect(slide, right_x, y, Inches(0.08), row_h, fill=color)
        add_text_box(
            slide, right_x + Inches(0.2), y, half_w * 0.6, row_h,
            text=app, size=FS_BODY_SM, bold=True, color=C_PRIMARY,
            anchor=MSO_ANCHOR.MIDDLE,
        )
        add_text_box(
            slide, right_x + half_w * 0.62, y, half_w * 0.36, row_h,
            text=status, size=FS_BODY_SM, color=color,
            anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT,
        )
        y += row_h + gap


def body_pdf_coverage(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="PDF Coverage Tracking (Uros)",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Sta meri", "strong"),
            "Koje su stranice PDF-a pokrivene pitanjima?",
            "Ponderisana pokrivenost po broju znakova",
            "Sustinska pokrivenost (izuzima skoro prazne stranice)",
            "",
            ("Vizualno - heatmap", "strong"),
            "Bar po stranici, visina = broj znakova",
            "Boja = da li je stranica pokrivena pitanjima",
            "Lista sustinskih stranica bez pitanja",
            "",
            ("Posledica", "strong"),
            "Sistem ti sam kaze gde da generises jos pitanja",
            "Heatmap je 'second pair of eyes' za nastavnika",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_image_placeholder(
        slide, right_x, CONTENT_Y, half_w, CONTENT_H,
        label="Coverage heatmap - screenshot iz UI-ja",
    )


def body_granice_hitl(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Granice automatizacije",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            "Definisanje ishoda - zahteva pedagosku procenu",
            "Procena tezine pitanja u kontekstu grupe - dinamicki kontekst",
            "Procena EA pitanja - cesto dvosmislena",
            "Recenzija video tutorijala - da li je didakticki dobar",
            "Konacno korigovanje halucinacija",
            "",
            ("Sve ovo radi - covek", "strong"),
            "Sistem snima pripremu (do 80% rada)",
            "Covek validira i fino podesava",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Human-in-the-loop u nasem pipeline-u",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    flow = [
        ("Sistem generise",   "Struktura, pitanja, video"),
        ("Nastavnik pregleda","source_line + ontology anchor + coverage"),
        ("Nastavnik koriguje","Edituje pitanja, rasporedjuje aktivnosti"),
        ("Sistem belezi",     "Promene u DB - feedback za buduce generisanje"),
    ]
    gap = Inches(0.12)
    n = len(flow)
    box_h = (CONTENT_H - Inches(0.5) - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.5)

    for i, (lbl, desc) in enumerate(flow):
        add_filled_rect(slide, right_x, y, half_w, box_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        add_filled_rect(slide, right_x, y, Inches(0.08), box_h,
                        fill=(C_ACCENT if i % 2 == 0 else C_PRIMARY))

        add_text_box(
            slide, right_x + Inches(0.25), y + Inches(0.1),
            half_w - Inches(0.3), Inches(0.35),
            text=f"{i+1}. {lbl}", size=FS_BODY, bold=True, color=C_PRIMARY,
        )
        add_text_box(
            slide, right_x + Inches(0.25), y + Inches(0.45),
            half_w - Inches(0.3), box_h - Inches(0.5),
            text=desc, size=FS_BODY_SM, color=C_TEXT,
        )
        y += box_h + gap


def body_seg_c_kljucna(slide):
    add_text_box(
        slide, CONTENT_X, CONTENT_Y + Inches(0.4), CONTENT_W, Inches(0.4),
        text="Kljucna poruka segmenta C",
        size=FS_SECTION_LABEL, bold=True, color=C_ACCENT,
        align=PP_ALIGN.CENTER,
    )
    box_w = CONTENT_W * 0.85
    box_x = CONTENT_X + (CONTENT_W - box_w) / 2
    box_y = CONTENT_Y + Inches(1.1)
    box_h = Inches(2.6)
    add_filled_rect(slide, box_x, box_y, box_w, box_h,
                    fill=C_CARD_BG, line=C_LIGHT_LINE)
    add_horizontal_line(slide, box_x, box_y, box_w, color=C_PRIMARY,
                        height=Emu(30000))
    add_text_box(
        slide, box_x + Inches(0.4), box_y + Inches(0.4), box_w - Inches(0.8),
        box_h - Inches(0.8),
        text=("Kvalitet automatski generisanog sadrzaja opada sa porastom "
              "kognitivne slozenosti - to nije bug, vec strukturna posledica, "
              "i upravo zato je human-in-the-loop obavezan."),
        size=22, italic=True, color=C_PRIMARY, align=PP_ALIGN.CENTER,
        anchor=MSO_ANCHOR.MIDDLE,
    )


# --- SEGMENT D helperi -------------------------------------------------------

def body_sinteza_arhitektura(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Integrisana arhitektura - predlog",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            ("Sloj 1 - Course Planner (Luka)", "strong"),
            "Strukturu kursa generise iz ishoda + PDF-a",
            "Izlaz: JSON / OWL hijerarhija aktivnosti",
            "",
            ("Sloj 2 - Knowledge Layer (Uros)", "strong"),
            "ConceptRelationship ontologija iz lekcija",
            "Source-of-truth za relacije pojmova",
            "",
            ("Sloj 3 - Assessment (Uros)", "strong"),
            "SOLO pitanja, ankored na knowledge layer",
            "",
            ("Sloj 4 - Demonstration (Stefan)", "strong"),
            "Praktiche aktivnosti -> NL instrukcija -> video",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_image_placeholder(
        slide, right_x, CONTENT_Y, half_w, CONTENT_H,
        label="Dijagram integrisane arhitekture (4 sloja)",
    )


def body_tok_podataka(slide):
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, CONTENT_W, Inches(0.4),
        text="Tok podataka kroz integrisan sistem",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    steps = [
        ("Nastavnik ucitava materijal",
         "PDF lekcije, lista ishoda, opis kursa"),
        ("Course Planner gradi strukturu",
         "Moduli -> koncepti -> ishodi -> aktivnosti (Bloom-aware)"),
        ("Knowledge Layer ekstrahuje ontologiju",
         "Multi-pass LLM ekstrakcija ConceptRelationship-ova"),
        ("Assessment generise SOLO pitanja",
         "Anchor iz Knowledge Layer-a, PS4 prompt, source_line"),
        ("Demonstration generise video za prakticnu aktivnost",
         "NL instrukcija -> OWL plan -> validacija -> izvrsenje + snimanje"),
        ("Nastavnik pregleda i koriguje",
         "Coverage heatmap, source_line provera, video review"),
    ]
    n = len(steps)
    gap = Inches(0.1)
    box_h = (CONTENT_H - Inches(0.5) - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.5)

    for i, (h, sub) in enumerate(steps):
        add_filled_rect(slide, CONTENT_X, y, CONTENT_W, box_h,
                        fill=C_CARD_BG if i % 2 == 0 else C_WHITE,
                        line=C_LIGHT_LINE)
        # broj
        add_filled_rect(slide, CONTENT_X, y, Inches(0.7), box_h,
                        fill=C_PRIMARY)
        add_text_box(slide, CONTENT_X, y, Inches(0.7), box_h,
                     text=f"{i+1}", size=24, bold=True, color=C_WHITE,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        add_text_box(
            slide, CONTENT_X + Inches(0.85), y + Inches(0.05),
            CONTENT_W - Inches(1.0), Inches(0.35),
            text=h, size=FS_BODY, bold=True, color=C_PRIMARY,
        )
        add_text_box(
            slide, CONTENT_X + Inches(0.85), y + Inches(0.4),
            CONTENT_W - Inches(1.0), box_h - Inches(0.4),
            text=sub, size=FS_BODY_SM, color=C_MUTED,
        )
        y += box_h + gap


def body_veze_drugim_temama(slide):
    cols = [
        ("Deep Reinforcement Learning",
         "Adaptivno biranje pitanja i puta ucenja",
         ["State: profil studenta + istorija odgovora",
          "Action: izbor narednog pitanja / aktivnosti",
          "Reward: napredak na SOLO / Bloomovoj skali",
          "Veza sa temom DRL for ITS sa liste predmeta"]),
        ("Recommender Systems",
         "Personalizacija puteva ucenja",
         ["Sadrzaj-based: ontology anchor sluzi kao feature",
          "Collaborative: studenti sa slicnim odgovorima",
          "Hibridno: koristi obe signala",
          "Veza sa temom Recommenders for ITS"]),
    ]
    gap = Inches(0.3)
    col_w = (CONTENT_W - gap) / 2

    for i, (title, sub, points) in enumerate(cols):
        x = CONTENT_X + i * (col_w + gap)
        add_filled_rect(slide, x, CONTENT_Y, col_w, CONTENT_H - Inches(0.1),
                        fill=C_WHITE, line=C_LIGHT_LINE)
        add_filled_rect(slide, x, CONTENT_Y, col_w, Inches(0.08), fill=C_ACCENT)

        pad = Inches(0.25)
        cur_y = CONTENT_Y + Inches(0.3)
        add_text_box(slide, x + pad, cur_y, col_w - 2 * pad, Inches(0.5),
                     text=title, size=20, bold=True, color=C_PRIMARY)
        cur_y += Inches(0.55)
        add_text_box(slide, x + pad, cur_y, col_w - 2 * pad, Inches(0.4),
                     text=sub, size=FS_BODY_SM, italic=True, color=C_ACCENT)
        cur_y += Inches(0.4)
        add_horizontal_line(slide, x + pad, cur_y, Inches(0.6),
                            color=C_ACCENT, height=Emu(15000))
        cur_y += Inches(0.2)
        add_bullet_list(slide, x + pad, cur_y, col_w - 2 * pad,
                        CONTENT_H - (cur_y - CONTENT_Y) - Inches(0.3),
                        points, size=FS_BODY_SM)


def body_otvorena_pitanja(slide):
    pitanja = [
        ("Kvalitet distraktora na EA nivou",
         "Kako podici uverljivost bez gubitka SOLO 'tipa' greske?"),
        ("Real-world evaluacija",
         "Sistem nije jos testiran u stvarnoj ucionici - to je sledeci korak."),
        ("Multimodalna procena znanja",
         "Kako uracunati ne samo tekstualne odgovore vec i prakticne demonstracije?"),
        ("Stabilnost vision modela",
         "Custom UI ostaju izazov; integracija sa accessibility API-jima?"),
        ("Skaliranje na vise predmeta",
         "Da li se ontologija prenosi izmedju domena ili je svaki predmet svoja TBox?"),
    ]
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, CONTENT_W, Inches(0.4),
        text="Otvorena pitanja", size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )

    n = len(pitanja)
    gap = Inches(0.12)
    box_h = (CONTENT_H - Inches(0.5) - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.5)

    for i, (h, sub) in enumerate(pitanja):
        add_filled_rect(slide, CONTENT_X, y, CONTENT_W, box_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        add_filled_rect(slide, CONTENT_X, y, Inches(0.5), box_h, fill=C_ACCENT)
        add_text_box(slide, CONTENT_X, y, Inches(0.5), box_h,
                     text=f"Q{i+1}", size=16, bold=True, color=C_WHITE,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text_box(
            slide, CONTENT_X + Inches(0.65), y + Inches(0.05),
            CONTENT_W - Inches(0.8), Inches(0.4),
            text=h, size=FS_BODY, bold=True, color=C_PRIMARY,
        )
        add_text_box(
            slide, CONTENT_X + Inches(0.65), y + Inches(0.4),
            CONTENT_W - Inches(0.8), box_h - Inches(0.4),
            text=sub, size=FS_BODY_SM, color=C_TEXT,
        )
        y += box_h + gap


def body_buduci_rad(slide):
    half_w = CONTENT_W / 2 - Inches(0.2)
    add_text_box(
        slide, CONTENT_X, CONTENT_Y, half_w, Inches(0.4),
        text="Naredni koraci - kratki rok",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, CONTENT_X, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            "Spojiti tri repozitorijuma u jedan monorepo",
            "Definisati API kontrakte izmedju slojeva",
            "Zajednicka ontoloska osnova (deljena seed_ontology)",
            "Demo end-to-end pipe za jedan predmet",
            "",
            ("Konferencija Sinteza 2026", "strong"),
            "Jedan rad iz tima vec publikovan",
            "Cilj: drugi rad o integraciji",
        ],
        size=FS_BODY_SM,
    )

    right_x = CONTENT_X + half_w + Inches(0.4)
    add_text_box(
        slide, right_x, CONTENT_Y, half_w, Inches(0.4),
        text="Naredni koraci - duzi rok",
        size=FS_SUBTITLE, bold=True, color=C_PRIMARY,
    )
    add_bullet_list(
        slide, right_x, CONTENT_Y + Inches(0.5), half_w, CONTENT_H - Inches(0.6),
        [
            "Pilot u stvarnoj ucionici",
            "Adaptivno biranje pitanja (DRL prototip)",
            "Personalizacija puta ucenja (Recommender)",
            "Multimodalna procena znanja (audio / kod / video)",
            "Skaliranje na vise predmeta",
            "",
            ("Etika i privatnost", "strong"),
            "Lokalni LLM kao osnovna opcija",
            "Pravo nastavnika na konacnu rec",
        ],
        size=FS_BODY_SM,
    )


def body_zakljucak(slide):
    add_text_box(
        slide, CONTENT_X, CONTENT_Y + Inches(0.2), CONTENT_W, Inches(0.4),
        text="Sta smo pokazali", size=FS_SECTION_LABEL, bold=True,
        color=C_ACCENT, align=PP_ALIGN.CENTER,
    )

    points = [
        ("LLM sam nije dovoljan", "halucinacije + nedostatak strukture"),
        ("Pedagoska taksonomija kalibrise", "Bloom + SOLO kao kompas"),
        ("Ontologija filtrira", "OWL/SPARQL + source_line + ontology anchor"),
        ("Human-in-the-loop ostaje obavezan", "narocito na visim SOLO nivoima"),
    ]
    gap = Inches(0.18)
    n = len(points)
    box_h = (CONTENT_H - Inches(1.2) - (n - 1) * gap) / n
    y = CONTENT_Y + Inches(0.7)

    for i, (h, sub) in enumerate(points):
        add_filled_rect(slide, CONTENT_X, y, CONTENT_W, box_h,
                        fill=C_CARD_BG, line=C_LIGHT_LINE)
        add_filled_rect(slide, CONTENT_X, y, Inches(0.1), box_h, fill=C_ACCENT)
        add_text_box(slide, CONTENT_X + Inches(0.25), y + Inches(0.1),
                     CONTENT_W - Inches(0.4), Inches(0.4),
                     text=h, size=FS_BODY, bold=True, color=C_PRIMARY)
        add_text_box(slide, CONTENT_X + Inches(0.25), y + Inches(0.45),
                     CONTENT_W - Inches(0.4), box_h - Inches(0.5),
                     text=sub, size=FS_BODY_SM, color=C_TEXT)
        y += box_h + gap


def body_literatura_lista(slide, refs):
    """Lista referenci (po kategorijama)."""
    cur_y = CONTENT_Y
    avail_h = CONTENT_H

    for cat, items in refs:
        add_text_box(
            slide, CONTENT_X, cur_y, CONTENT_W, Inches(0.32),
            text=cat, size=FS_BODY, bold=True, color=C_ACCENT,
        )
        cur_y += Inches(0.35)
        for ref in items:
            add_text_box(
                slide, CONTENT_X + Inches(0.2), cur_y,
                CONTENT_W - Inches(0.2), Inches(0.45),
                text="- " + ref, size=FS_BODY_SM, color=C_TEXT,
            )
            cur_y += Inches(0.48)
        cur_y += Inches(0.1)


# =============================================================================
# IZGRADNJA PREZENTACIJE
# =============================================================================

def build_presentation(out_path):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # Pratimo broj slidova za footer
    slide_specs = []

    # ---- Naslovna ----
    slide_specs.append(("title", {}))

    # ---- Uvod (predmet + tim + agenda) ----
    slide_specs.append(("content", dict(
        page_label=None,
        title="Tim",
        subtitle="Tri projekta sa predmeta SOTIS - jedan integrisan ITS pipeline",
        body_func=body_team,
    )))
    slide_specs.append(("content", dict(
        page_label=None,
        title="Predmet i kontekst",
        subtitle="Napredne tehnike racunarske inteligencije",
        body_func=body_predmet_context,
    )))
    slide_specs.append(("content", dict(
        page_label=None,
        title="Agenda",
        subtitle="5 segmenata - tematska organizacija, ne po projektima",
        body_func=body_agenda,
    )))

    # ============ SEGMENT 0 - UVOD I MOTIVACIJA ============
    slide_specs.append(("divider", dict(
        segment_label="Segment 0",
        title="Uvod i motivacija",
        subtitle="Polazimo od trenutnog stanja - sta nastavnik radi i zasto je to skupo.",
        time_str="10 min  /  4 slajda",
    )))
    slide_specs.append(("content", dict(
        page_label="Segment 0 - Uvod",
        title="Problem - sta nastavnik radi rucno",
        subtitle="Vreme i kognitivni napor su jaka skala bola.",
        body_func=body_problem_stats,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment 0 - Uvod",
        title="LLM otvara vrata, ali ne sam",
        subtitle="Sta dobijemo - i sta sve uneso skripa.",
        body_func=body_llm_kao_odgovor,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment 0 - Vizija",
        title="Vizija - ITS pipeline od ishoda do video tutorijala",
        subtitle="5 koraka, end-to-end, bez rucne intervencije izmedju.",
        body_func=body_vizija_pipeline,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment 0 - Komponente",
        title="Tri komponente jednog pipeline-a",
        subtitle="Svaki clan tima radi jedan deo - zajedno cine celinu.",
        body_func=body_tri_komponente,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment 0 - Kljucna poruka",
        title="Kljucna poruka prezentacije",
        body_func=body_kljucna_poruka_uvod,
    )))

    # ============ SEGMENT A - LLM PROMPTING ============
    slide_specs.append(("divider", dict(
        segment_label="Segment A",
        title="LLM prompting i pedagoske taksonomije",
        subtitle="Hijerarhijska dekompozicija + formalna taksonomija u promptu = pedagoski kalibrisan LLM.",
        time_str="25 min  /  11 slajdova",
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Bloom",
        title="Bloomova taksonomija - 6 nivoa misaonih operacija",
        subtitle="Glagol akcije: pamtiti, razumeti, primeniti, analizirati, evaluirati, kreirati.",
        body_func=body_bloom_pregled,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - SOLO",
        title="SOLO taksonomija - strukturalna slozenost odgovora",
        subtitle="Od jedne cinjenice do generalizacije i transfera znanja.",
        body_func=body_solo_pregled,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Poredjenje",
        title="Bloom vs SOLO - razliciti pogledi, komplementarni",
        body_func=body_bloom_vs_solo,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Luka",
        title="Luka - hijerarhijska struktura kursa",
        subtitle="LLM razlaze kurs u stablo modula, koncepata, ishoda i aktivnosti.",
        body_func=body_luka_struktura,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Luka",
        title="Luka - prompt template + state-of-generation",
        body_func=body_luka_prompt_template,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Uros",
        title="Uros - SOLO Quiz Generator i PS4 prompt template",
        body_func=body_uros_solo_ps4,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Uros",
        title="Two-pass generisanje za Extended Abstract",
        subtitle="Razdvajamo 'sta pitati' od 'kako napraviti dobre distraktore'.",
        body_func=body_two_pass_ea,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Uros",
        title="Source-line citation - anti-halucinacija",
        body_func=body_source_line,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Stefan",
        title="Stefan - hijerarhijska dekompozicija zadataka",
        subtitle="Task -> Step -> Action: tri nivoa razgradnje prirodno-jezicke instrukcije.",
        body_func=body_stefan_dekompozicija,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Tradeoff",
        title="Prompt engineering vs Fine-tuning",
        body_func=body_prompt_vs_finetune,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Tradeoff",
        title="Lokalni vs cloud LLM-ovi - mi koristimo oba",
        body_func=body_local_vs_cloud,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment A - Kljucna poruka",
        title="Kljucna poruka segmenta A",
        body_func=body_seg_a_kljucna,
    )))

    # ============ SEGMENT B - ONTOLOGIJE ============
    slide_specs.append(("divider", dict(
        segment_label="Segment B",
        title="Ontologije i formalni grounding",
        subtitle="Ontologija je ugovor izmedju LLM-a i sistema - sve sto LLM proizvede mora se mapirati.",
        time_str="25 min  /  10 slajdova",
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Osnove",
        title="OWL - Web Ontology Language",
        subtitle="Klase, instance, relacije, ogranicenja, mas i mas serijalizacija.",
        body_func=body_owl_uvod,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Osnove",
        title="RDF i SPARQL - tripleti i upiti",
        subtitle="Sve cinjenice su tripleti; sve upite pisemo nad gradom triplet-a.",
        body_func=body_rdf_sparql,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Luka",
        title="Luka - JSON serijalizacija strukture kursa",
        subtitle="LLM lakse generise validan JSON nego direktan OWL - JSON -> OWL je determinizam.",
        body_func=body_luka_json_primer,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Uros",
        title="Uros - ConceptRelationship ontologija",
        subtitle="Relacije izmedju learning objekata - osnova za relacionalna i EA pitanja.",
        body_func=body_uros_ontologija,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Uros",
        title="Uros - SPARQL upit za 'ontology anchor'",
        subtitle="Pitanje se ne pravi 'iz vazduha' - bira se konkretna veza i ona postaje kontekst.",
        body_func=body_ontology_anchor,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Stefan",
        title="Stefan - OWL ontologija izvrsnog plana",
        subtitle="Plan koji LLM generise mora se mapirati na klase Task / Step / Action.",
        body_func=body_stefan_owl,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Stefan",
        title="Stefan - pravila validacije plana (SPARQL)",
        subtitle="Pre nego sto se ista izvrsi, SPARQL upiti proveravaju strukturna pravila.",
        body_func=body_validacija_pravila,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Anti-halucinacija",
        title="Anti-halucinacioni mehanizmi - 3 ugla",
        subtitle="Svaki clan tima pokriva drugaciju vrstu greske LLM-a.",
        body_func=body_anti_halucinacije,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment B - Kljucna poruka",
        title="Kljucna poruka segmenta B",
        body_func=body_seg_b_kljucna,
    )))

    # ============ SEGMENT C - EVALUACIJA ============
    slide_specs.append(("divider", dict(
        segment_label="Segment C",
        title="Evaluacija, kvalitet i ogranicenja",
        subtitle="Kvalitet opada sa kognitivnom slozenoscu - to opravdava human-in-the-loop.",
        time_str="25 min  /  9 slajdova",
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - Pilot",
        title="Pilot evaluacija SOLO pitanja",
        subtitle="Metodologija: PDF lekcije iz vise domena, anotacija po dimenzijama.",
        body_func=body_pilot_metodologija,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - Rezultati",
        title="Rezultati - kvalitet po SOLO nivou",
        subtitle="Validnost i kvalitet distraktora padaju kako se ide ka EA nivou.",
        body_func=body_kvalitet_po_nivou,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - Primeri",
        title="Primeri SOLO pitanja iz domena Operativnih sistema",
        body_func=body_primer_solo_pitanja,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - Luka",
        title="Luka - Bloomova pokrivenost generisanih kurseva",
        subtitle="Vidimo gde sistem dobro pokriva nivoe, a gde tek 'dotice' visa misljenja.",
        body_func=body_luka_bloom_pokrivenost,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - Stefan",
        title="Stefan - robustnost izvrsenja video tutorijala",
        subtitle="Vision modeli su odlicni, ali ne savrseni - non-standard UI je glavna bolna tacka.",
        body_func=body_stefan_ui_robust,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - Uros",
        title="PDF Coverage Tracking - heatmap pokrivenosti",
        subtitle="Sistem sam ti kaze gde da generises jos pitanja.",
        body_func=body_pdf_coverage,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - HITL",
        title="Granice automatizacije i human-in-the-loop",
        subtitle="Sistem snima 80% rada; covek validira i fino podesava.",
        body_func=body_granice_hitl,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment C - Kljucna poruka",
        title="Kljucna poruka segmenta C",
        body_func=body_seg_c_kljucna,
    )))

    # ============ SEGMENT D - SINTEZA ============
    slide_specs.append(("divider", dict(
        segment_label="Segment D",
        title="Sinteza, buduci rad, zakljucak",
        subtitle="Kako tri komponente postaju jedan sistem - i sta sledi.",
        time_str="10-15 min  /  7 slajdova",
    )))
    slide_specs.append(("content", dict(
        page_label="Segment D - Arhitektura",
        title="Predlog integrisane arhitekture - 4 sloja",
        subtitle="Sloj po sloj: planiranje -> znanje -> procena -> demonstracija.",
        body_func=body_sinteza_arhitektura,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment D - Arhitektura",
        title="Tok podataka kroz integrisan sistem",
        subtitle="Sest koraka, od ucitavanja materijala do recenzije nastavnika.",
        body_func=body_tok_podataka,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment D - Veze",
        title="Veze sa drugim temama predmeta",
        subtitle="DRL i Recommender Systems prirodno produzavaju ovaj pipeline.",
        body_func=body_veze_drugim_temama,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment D - Buduci rad",
        title="Otvorena pitanja",
        body_func=body_otvorena_pitanja,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment D - Buduci rad",
        title="Naredni koraci - kratki i duzi rok",
        body_func=body_buduci_rad,
    )))
    slide_specs.append(("content", dict(
        page_label="Segment D - Zakljucak",
        title="Zakljucak",
        subtitle="Vracamo se na pocetnu poruku.",
        body_func=body_zakljucak,
    )))
    slide_specs.append(("quote", dict(
        page_label="Zakljucna misao",
        quote=("LLM + ontologija + pedagoska taksonomija = "
               "pouzdano, pedagoski kalibrisano ITS okruzenje koje je "
               "vise od zbira svojih delova."),
        attribution="kljucna poruka prezentacije",
    )))

    # ---- Hvala ----
    slide_specs.append(("thanks", {}))

    # ---- Literatura ----
    literatura_1 = [
        ("Pedagoske taksonomije", [
            "Biggs, J.B. & Collis, K.F. (1982). Evaluating the Quality of Learning: The SOLO Taxonomy. Academic Press.",
            "Anderson, L.W. & Krathwohl, D.R. (2001). A Taxonomy for Learning, Teaching, and Assessing. Longman.",
            "Lister, R., et al. (2006). Not Seeing the Forest for the Trees: Novice Programmers and the SOLO Taxonomy. ACM SIGCSE Bulletin, 38(3).",
        ]),
        ("LLM u obrazovanju", [
            "Kasneci, E., et al. (2023). ChatGPT for good? Learning and Individual Differences, 103.",
            "Yan, L., et al. (2024). Practical and ethical challenges of LLMs in education. BJET, 55(1).",
            "Wang, S., et al. (2024). Large Language Models for Education: A Survey. arXiv:2403.18105.",
        ]),
    ]
    literatura_2 = [
        ("Automatsko generisanje pitanja", [
            "Kurdi, G., et al. (2020). A Systematic Review of Automatic Question Generation. IJAIED, 30(1).",
            "Liang, C., et al. (2018). Distractor Generation for Multiple Choice Questions Using Learning to Rank. NAACL Workshop.",
        ]),
        ("Ontologije i Semantic Web", [
            "Vesin, B., et al. (2013). Ontology-based architecture with recommendation strategy in Java tutoring. CSIS, 10(1).",
            "Hogan, A., et al. (2021). Knowledge Graphs. ACM Computing Surveys, 54(4).",
        ]),
        ("Intelligent Tutoring Systems", [
            "Mousavinasab, E., et al. (2021). ITS: systematic review. Interactive Learning Environments, 29(1).",
            "Lin, C.-C., et al. (2023). AI in ITS toward sustainable education. Smart Learning Environments, 10(1).",
        ]),
        ("Multimodalni agenti i Computer Use", [
            "Anthropic (2024). Developing a computer use model. Technical report.",
            "Xie, T., et al. (2024). OSWorld: Benchmarking Multimodal Agents. NeurIPS 2024.",
        ]),
    ]

    slide_specs.append(("references", dict(
        page_label="Literatura -1 / 2",
        title="Literatura",
        subtitle="Pedagoske taksonomije i LLM u obrazovanju",
        refs=literatura_1,
    )))
    slide_specs.append(("references", dict(
        page_label="Literatura -2 / 2",
        title="Literatura (nastavak)",
        subtitle="Generisanje pitanja, ontologije, ITS, multimodal",
        refs=literatura_2,
    )))

    total_pages = len(slide_specs)

    # ---- Renderovanje ----
    for i, (kind, args) in enumerate(slide_specs, start=1):
        if kind == "title":
            make_title_slide(
                prs,
                title=("Od ishoda ucenja do automatizovanog ITS-a"),
                subtitle=("LLM-driven pipeline za generisanje strukture kursa, "
                          "video demonstracija i pitanja zasnovanih na SOLO taksonomiji"),
                team="Uros Petraskovic - Luka Saric - Stefan Lazarevic",
                course="Napredne tehnike racunarske inteligencije - Master, FTN UNS",
                date_str="2026.",
            )
        elif kind == "divider":
            make_section_divider(prs, **args)
        elif kind == "content":
            make_content_slide(prs, page_no=i, total_pages=total_pages, **args)
        elif kind == "quote":
            make_quote_slide(prs, page_no=i, total_pages=total_pages, **args)
        elif kind == "thanks":
            make_thank_you_slide(
                prs,
                team="Uros Petraskovic - Luka Saric - Stefan Lazarevic",
            )
        elif kind == "references":
            slide = new_slide(prs)
            add_header(slide,
                       title=args["title"], subtitle=args.get("subtitle"),
                       page_label=args.get("page_label"))
            body_literatura_lista(slide, args["refs"])
            add_footer(slide, page_number=i, total=total_pages)

    prs.save(out_path)
    print(f"Generisano: {out_path}  ({total_pages} slidova)")


if __name__ == "__main__":
    import os
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.join(here, "Prezentacija_ITS_pipeline.pptx")
    build_presentation(out)
