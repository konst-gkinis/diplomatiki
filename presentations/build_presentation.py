# -*- coding: utf-8 -*-
"""
Generator for the thesis-defense presentation:
  «Πλατφόρμα λογισμικού για την επίτευξη του στόχου μηδενικές εκπομπές CO₂:
   Η περίπτωση της Mærsk»  —  Κωνσταντίνος Γκίνης, Πανεπιστήμιο Πατρών.

Builds presentations/defense.pptx with python-pptx.

Design goals (per defense feedback):
  * Greek, University-of-Patras academic styling (burgundy + gold, serif headings).
  * Results-forward; theory compressed to a single slide.
  * Dedicated functional / non-functional (technical) requirements slides.
  * Screenshot PLACEHOLDERS mapped to the real platform features — swap in the
    actual images later by replacing add_placeholder(...) with add_image(...).
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
OUT = os.path.join(HERE, "defense.pptx")

# ── Palette (University of Patras) ────────────────────────────────────────
UP_RED   = RGBColor(0x7A, 0x16, 0x26)   # burgundy from the UP logo
UP_RED_D = RGBColor(0x57, 0x0F, 0x1B)   # darker burgundy
UP_GOLD  = RGBColor(0xC2, 0x9B, 0x52)   # gold/tan from the UP logo
DARK     = RGBColor(0x23, 0x23, 0x23)
GRAY     = RGBColor(0x6E, 0x6E, 0x6E)
LIGHT    = RGBColor(0xF6, 0xF2, 0xEA)   # warm off-white
PANEL    = RGBColor(0xEC, 0xE4, 0xD6)   # warm panel
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
PLACE_BG = RGBColor(0xF0, 0xED, 0xE6)
PLACE_BD = RGBColor(0xB9, 0xA5, 0x7E)

# Fonts bundled with LibreOffice (Greek-capable, metric-compatible with MS):
HEAD = "Caladea"     # ≈ Cambria  (serif, academic headings)
BODY = "Carlito"     # ≈ Calibri  (clean body text)

SW, SH = Inches(13.333), Inches(7.5)   # 16:9

prs = Presentation()
prs.slide_width = SW
prs.slide_height = SH
BLANK = prs.slide_layouts[6]


# ── low-level helpers ─────────────────────────────────────────────────────
def slide():
    return prs.slides.add_slide(BLANK)


def rect(s, x, y, w, h, fill=None, line=None, line_w=None, shape=MSO_SHAPE.RECTANGLE):
    sp = s.shapes.add_shape(shape, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
    return sp


def _set_run(r, text, size, color, bold, italic, font):
    r.text = text
    r.font.size = Pt(size)
    r.font.color.rgb = color
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font
    # ensure Greek glyphs use the same face for complex/east-asian slots
    rPr = r._r.get_or_add_rPr()
    for tag in ("a:latin", "a:cs", "a:ea"):
        e = rPr.find(qn(tag))
        if e is None:
            e = rPr.makeelement(qn(tag), {}); rPr.append(e)
        e.set("typeface", font)


def textbox(s, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
            wrap=True):
    """lines: list of dicts or list of [runs]; each line -> paragraph.
    A 'line' may be:
      {"t":str,"size":..,"color":..,"bold":..,"italic":..,"font":..,
       "align":..,"bullet":bool,"space_after":..,"level":int}
    or a list of run-dicts (mixed formatting in one paragraph)."""
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Pt(2)
    tf.margin_top = tf.margin_bottom = Pt(2)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        runs = ln if isinstance(ln, list) else [ln]
        meta = runs[0]
        p.alignment = meta.get("align", align)
        if meta.get("space_after") is not None:
            p.space_after = Pt(meta["space_after"])
        if meta.get("space_before") is not None:
            p.space_before = Pt(meta["space_before"])
        if meta.get("line_spacing"):
            p.line_spacing = meta["line_spacing"]
        p.level = meta.get("level", 0)
        for rd in runs:
            r = p.add_run()
            _set_run(r, rd.get("t", ""), rd.get("size", 16),
                     rd.get("color", DARK), rd.get("bold", False),
                     rd.get("italic", False), rd.get("font", BODY))
        if meta.get("bullet"):
            _bullet(p, meta.get("bullet_color", UP_RED))
        else:
            _no_bullet(p)
    return tb


def _bullet(p, color):
    pPr = p._p.get_or_add_pPr()
    pPr.set("indent", str(-Inches(0.22)))
    pPr.set("marL", str(Inches(0.22)))
    buFont = pPr.makeelement(qn("a:buFont"),
                             {"typeface": "Arial", "pitchFamily": "34", "charset": "0"})
    buChar = pPr.makeelement(qn("a:buChar"), {"char": "▪"})
    pPr.append(buFont); pPr.append(buChar)


def _no_bullet(p):
    pPr = p._p.get_or_add_pPr()
    pPr.append(pPr.makeelement(qn("a:buNone"), {}))


def fit(img_path, max_w, max_h):
    iw, ih = Image.open(img_path).size
    r = min(max_w / iw, max_h / ih)
    return Emu(int(iw * r)), Emu(int(ih * r))


def add_image(s, path, x, y, max_w, max_h, center_in=True):
    w, h = fit(path, max_w, max_h)
    if center_in:
        x = x + Emu(int((max_w - w) / 2))
        y = y + Emu(int((max_h - h) / 2))
    return s.shapes.add_picture(path, x, y, w, h)


def add_placeholder(s, x, y, w, h, title, desc):
    box = rect(s, x, y, w, h, fill=PLACE_BG, line=PLACE_BD, line_w=Pt(1.5))
    box.line.dash_style = None
    # dashed border
    ln = box.line._get_or_add_ln()
    d = ln.makeelement(qn("a:prstDash"), {"val": "dash"}); ln.append(d)
    textbox(s, x + Inches(0.2), y, w - Inches(0.4), h, [
        {"t": "▣  ΣΤΙΓΜΙΟΤΥΠΟ ΟΘΟΝΗΣ", "size": 13, "color": UP_RED,
         "bold": True, "font": BODY, "align": PP_ALIGN.CENTER, "space_after": 6},
        {"t": title, "size": 15, "color": DARK, "bold": True, "font": HEAD,
         "align": PP_ALIGN.CENTER, "space_after": 6},
        {"t": desc, "size": 11.5, "color": GRAY, "italic": True, "font": BODY,
         "align": PP_ALIGN.CENTER, "line_spacing": 1.05},
    ], anchor=MSO_ANCHOR.MIDDLE)


# ── chrome: title bar + footer ────────────────────────────────────────────
PAGE = [0]


def header(s, kicker, title):
    rect(s, 0, 0, SW, Inches(1.12), fill=UP_RED)
    rect(s, 0, Inches(1.12), SW, Inches(0.07), fill=UP_GOLD)
    if kicker:
        textbox(s, Inches(0.55), Inches(0.13), Inches(11), Inches(0.3),
                [{"t": kicker.upper(), "size": 11, "color": UP_GOLD, "bold": True,
                  "font": BODY}])
    textbox(s, Inches(0.55), Inches(0.40), Inches(12.2), Inches(0.66),
            [{"t": title, "size": 25, "color": WHITE, "bold": True, "font": HEAD}],
            anchor=MSO_ANCHOR.MIDDLE)


def footer(s, dark=False):
    PAGE[0] += 1
    c = WHITE if dark else GRAY
    textbox(s, Inches(0.55), Inches(7.06), Inches(9), Inches(0.34),
            [{"t": "Πάγκος Εργασίας Εκπομπών — Διπλωματική Εργασία, Κ. Γκίνης",
              "size": 9.5, "color": c, "font": BODY}],
            anchor=MSO_ANCHOR.MIDDLE)
    textbox(s, Inches(12.2), Inches(7.06), Inches(0.9), Inches(0.34),
            [{"t": str(PAGE[0]), "size": 9.5, "color": c, "font": BODY,
              "align": PP_ALIGN.RIGHT}], anchor=MSO_ANCHOR.MIDDLE)


def body_area():
    """Return (x, y, w, h) of the usable content region below the header."""
    return Inches(0.55), Inches(1.42), Inches(12.23), Inches(5.5)


def chip(s, x, y, w, text, color=UP_RED):
    h = Inches(0.42)
    rect(s, x, y, w, h, fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    textbox(s, x, y, w, h, [{"t": text, "size": 13, "color": WHITE, "bold": True,
            "font": BODY, "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)


# ── styled table ──────────────────────────────────────────────────────────
def table(s, x, y, w, rows, col_w, header_row=True, font_size=12.5,
          row_h=Inches(0.42)):
    nr, nc = len(rows), len(rows[0])
    gt = s.shapes.add_table(nr, nc, x, y, w, row_h * nr).table
    gt.first_row = False
    gt.horz_banding = False
    # remove default style banding by setting explicit fills
    total = sum(col_w)
    for j, cw in enumerate(col_w):
        gt.columns[j].width = Emu(int(w * cw / total))
    for i, row in enumerate(rows):
        gt.rows[i].height = row_h
        for j, val in enumerate(row):
            c = gt.cell(i, j)
            c.margin_left = Inches(0.12); c.margin_right = Inches(0.08)
            c.margin_top = Inches(0.03); c.margin_bottom = Inches(0.03)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            if header_row and i == 0:
                c.fill.solid(); c.fill.fore_color.rgb = UP_RED
                col, bold, fs = WHITE, True, font_size
            else:
                c.fill.solid()
                c.fill.fore_color.rgb = WHITE if (i % 2) else PANEL
                col, bold, fs = DARK, False, font_size
            tf = c.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.LEFT
            r = p.add_run()
            _set_run(r, str(val), fs, col, bold or (j == 0 and i > 0 and val.startswith("•")),
                     False, BODY)
            if j == 0 and i > 0:
                r.font.bold = True
                r.font.color.rgb = UP_RED_D
    return gt


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 1 — TITLE
# ══════════════════════════════════════════════════════════════════════════
def s_title():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=LIGHT)
    rect(s, 0, 0, Inches(0.32), SH, fill=UP_RED)
    rect(s, Inches(0.32), 0, Inches(0.07), SH, fill=UP_GOLD)
    add_image(s, os.path.join(ASSETS, "logo-up-4color-landscape.jpg"),
              Inches(0.9), Inches(0.55), Inches(1.7), Inches(1.7), center_in=False)
    textbox(s, Inches(2.85), Inches(0.62), Inches(9.8), Inches(1.4), [
        {"t": "ΠΑΝΕΠΙΣΤΗΜΙΟ ΠΑΤΡΩΝ", "size": 17, "color": UP_RED, "bold": True,
         "font": HEAD, "space_after": 2},
        {"t": "Πολυτεχνική Σχολή", "size": 13, "color": DARK, "font": BODY,
         "space_after": 1},
        {"t": "Τμήμα Μηχανικών Η/Υ και Πληροφορικής", "size": 13, "color": DARK,
         "font": BODY},
    ])
    rect(s, Inches(0.9), Inches(2.62), Inches(11.5), Pt(2), fill=UP_GOLD)
    textbox(s, Inches(0.9), Inches(2.95), Inches(11.6), Inches(2.0), [
        {"t": "Πλατφόρμα λογισμικού για την επίτευξη του στόχου",
         "size": 30, "color": UP_RED_D, "bold": True, "font": HEAD,
         "line_spacing": 1.05, "space_after": 0},
        {"t": "μηδενικές εκπομπές CO₂: Η περίπτωση της Mærsk",
         "size": 30, "color": UP_RED_D, "bold": True, "font": HEAD,
         "line_spacing": 1.05},
    ])
    textbox(s, Inches(0.9), Inches(4.55), Inches(11.6), Inches(0.5), [
        {"t": "Διπλωματική Εργασία", "size": 15, "color": GRAY, "italic": True,
         "font": BODY}])
    # author + committee block
    textbox(s, Inches(0.9), Inches(5.25), Inches(6.5), Inches(1.6), [
        {"t": "Κωνσταντίνος Γκίνης", "size": 19, "color": DARK, "bold": True,
         "font": HEAD, "space_after": 8},
        [{"t": "Επιβλέπων:  ", "size": 13, "color": GRAY, "font": BODY},
         {"t": "Χρήστος Μπούρας, Καθηγητής", "size": 13, "color": DARK, "font": BODY}],
    ])
    textbox(s, Inches(7.4), Inches(5.25), Inches(5.0), Inches(1.6), [
        {"t": "Εξεταστική Επιτροπή", "size": 12, "color": GRAY, "bold": True,
         "font": BODY, "space_after": 4},
        {"t": "Ιωάννης Γαροφαλάκης, Καθηγητής", "size": 12, "color": DARK,
         "font": BODY, "space_after": 1},
        {"t": "Εύη Παπαϊωάννου, Επίκουρη Καθηγήτρια", "size": 12, "color": DARK,
         "font": BODY},
    ])
    textbox(s, Inches(0.9), Inches(6.85), Inches(11.6), Inches(0.4), [
        {"t": "Πάτρα, Μάιος 2026", "size": 12, "color": UP_RED, "bold": True,
         "font": BODY}])


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 2 — THE PROBLEM
# ══════════════════════════════════════════════════════════════════════════
def s_problem():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Το πλαίσιο", "Το πρόβλημα: η χειρωνακτική διαδικασία")
    bx, by, bw, bh = body_area()
    textbox(s, bx, by, Inches(6.7), Inches(0.9), [
        {"t": "Η Maersk — στόλος 700+ πλοίων — αδυνατούσε να ανταποκριθεί "
              "αξιόπιστα στις απαιτήσεις για δεδομένα εκπομπών.",
         "size": 16, "color": DARK, "font": BODY, "line_spacing": 1.12}])
    steps = [
        ("Αίτημα μέσω email", "Ο πελάτης ή το τμήμα πωλήσεων ζητά δεδομένα εκπομπών"),
        ("Χειρωνακτική εξαγωγή", "Μη αυτόματη άντληση δεδομένων από πολλαπλές πηγές"),
        ("Excel + συντελεστές", "Εφαρμογή συντελεστών εκπομπών σε υπολογιστικά φύλλα"),
        ("Απάντηση σε 3–8 εβδομάδες", "Αποστολή αποτελέσματος ξανά μέσω email"),
    ]
    yy = by + Inches(1.0)
    for i, (t, d) in enumerate(steps):
        rect(s, bx, yy, Inches(0.42), Inches(0.42), fill=UP_RED,
             shape=MSO_SHAPE.OVAL)
        textbox(s, bx, yy, Inches(0.42), Inches(0.42),
                [{"t": str(i + 1), "size": 14, "color": WHITE, "bold": True,
                  "font": BODY, "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
        textbox(s, bx + Inches(0.6), yy - Inches(0.04), Inches(6.0), Inches(0.6), [
            [{"t": t + "  —  ", "size": 14.5, "color": UP_RED_D, "bold": True,
              "font": BODY},
             {"t": d, "size": 13, "color": DARK, "font": BODY}]])
        yy = yy + Inches(0.72)
    # right panel: the cost
    px = Inches(7.7)
    rect(s, px, by, Inches(5.05), Inches(4.55), fill=LIGHT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, px, by, Inches(0.12), Inches(4.55), fill=UP_GOLD)
    textbox(s, px + Inches(0.4), by + Inches(0.3), Inches(4.3), Inches(4.0), [
        {"t": "Τι κόστιζε αυτό", "size": 17, "color": UP_RED, "bold": True,
         "font": HEAD, "space_after": 12},
        {"t": "Ακρίβεια", "bullet": True, "size": 14.5, "color": DARK, "bold": True,
         "font": BODY, "space_after": 1},
        {"t": "λάθη μεθοδολογίας, ελλιπής κανονικοποίηση δεδομένων", "size": 12.5,
         "color": GRAY, "font": BODY, "level": 1, "space_after": 8},
        {"t": "Ταχύτητα", "bullet": True, "size": 14.5, "color": DARK, "bold": True,
         "font": BODY, "space_after": 1},
        {"t": "καθυστέρηση εβδομάδων — απώλεια συμβολαίων", "size": 12.5,
         "color": GRAY, "font": BODY, "level": 1, "space_after": 8},
        {"t": "Ιχνηλασιμότητα", "bullet": True, "size": 14.5, "color": DARK,
         "bold": True, "font": BODY, "space_after": 1},
        {"t": "καμία δυνατότητα ανεξάρτητου ελέγχου (audit)", "size": 12.5,
         "color": GRAY, "font": BODY, "level": 1, "space_after": 8},
        {"t": "Κλίμακα", "bullet": True, "size": 14.5, "color": DARK, "bold": True,
         "font": BODY, "space_after": 1},
        {"t": "1.300+ χρήστες, ανθρώπινη δυναμικότητα δεν κλιμακώνεται", "size": 12.5,
         "color": GRAY, "font": BODY, "level": 1},
    ])
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 3 — GIA AUDIT (catalyst → requirements)
# ══════════════════════════════════════════════════════════════════════════
def s_audit():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Ο καταλύτης", "Ο εσωτερικός έλεγχος GIA 2023")
    bx, by, bw, bh = body_area()
    textbox(s, bx, by, Inches(12.2), Inches(0.7), [
        [{"t": "Ο έλεγχος Group Internal Audit (2023) ", "size": 15.5,
          "color": DARK, "font": BODY},
         {"t": "μετέτρεψε ένα γνωστό λειτουργικό πρόβλημα σε επίσημη, "
               "επείγουσα προτεραιότητα — και όρισε τις απαιτήσεις του συστήματος.",
          "size": 15.5, "color": DARK, "font": BODY}]],
    )
    cards = [
        ("Παρακολούθηση δεδομένων", "Καμία ενιαία ιχνηλάτηση ροής δεδομένων από την πηγή έως την αναφορά"),
        ("Ίχνος ελέγχου (audit trail)", "Αδύνατη η ανεξάρτητη επαλήθευση των υπολογισμών — κρίσιμο για CSRD/EU ETS"),
        ("Κλιμακωσιμότητα", "Αύξηση αιτημάτων χωρίς αντίστοιχη κλιμάκωση ανθρώπινης δυναμικότητας"),
        ("Ποσοστό σφαλμάτων", "Εγγενώς υψηλότερα λάθη σε χειρωνακτικές διαδικασίες"),
    ]
    cw, ch = Inches(2.92), Inches(2.55)
    gap = Inches(0.16)
    x0 = bx
    yy = by + Inches(1.0)
    for i, (t, d) in enumerate(cards):
        x = x0 + i * (cw + gap)
        rect(s, x, yy, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, x, yy, cw, Inches(0.1), fill=UP_RED)
        textbox(s, x + Inches(0.22), yy + Inches(0.28), cw - Inches(0.44),
                ch - Inches(0.4), [
            {"t": "0" + str(i + 1), "size": 22, "color": UP_GOLD, "bold": True,
             "font": HEAD, "space_after": 6},
            {"t": t, "size": 15, "color": UP_RED_D, "bold": True, "font": HEAD,
             "space_after": 8, "line_spacing": 1.0},
            {"t": d, "size": 12.5, "color": DARK, "font": BODY, "line_spacing": 1.1},
        ])
    rect(s, bx, yy + ch + Inches(0.22), Inches(12.2), Inches(0.6), fill=PANEL,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    textbox(s, bx + Inches(0.3), yy + ch + Inches(0.22), Inches(11.6), Inches(0.6), [
        [{"t": "➜  Αποτέλεσμα:  ", "size": 14, "color": UP_RED, "bold": True,
          "font": BODY},
         {"t": "έγκριση πόρων για εξ ολοκλήρου νέο σύστημα — ανάπτυξη από τον "
               "Μάρτιο 2023, πρώτοι χρήστες παραγωγής τον Σεπτέμβριο 2023.",
          "size": 14, "color": DARK, "font": BODY}]],
        anchor=MSO_ANCHOR.MIDDLE)
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 4 — FUNCTIONAL REQUIREMENTS
# ══════════════════════════════════════════════════════════════════════════
def s_func_req():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Προδιαγραφές · 1/2", "Λειτουργικές απαιτήσεις")
    bx, by, bw, bh = body_area()
    textbox(s, bx, by, Inches(12.2), Inches(0.55), [
        {"t": "Έξι ομάδες χρηστών με διαφορετικές ανάγκες όρισαν τι πρέπει να "
              "κάνει το σύστημα:", "size": 14.5, "color": DARK, "font": BODY}])
    rows = [
        ["Ομάδα χρηστών", "Τι χρειάζεται", "Συχνότητα"],
        ["ECO Delivery Products", "Εκπομπές ανά αποστολή· πιστοποιητικά ECOv1/ECOv2", "Πολλές φορές/ημέρα"],
        ["Regional Product Mgmt", "Συγκεντρωτικά δεδομένα & τάσεις ανά περιοχή", "Εβδομαδιαία/μηνιαία"],
        ["Commercial Sustainability", "Επαληθεύσιμα στοιχεία για ESG / CSRD", "Τριμηνιαία/ετήσια"],
        ["Regional Sales", "Γρήγορη πρόσβαση για ζωντανές διαπραγματεύσεις", "Ad hoc, υψηλή πίεση"],
        ["Contract Management", "Audit-ready απόδειξη συμβατικών δεσμεύσεων", "Ανά σύμβαση"],
        ["Account Managers", "Εκπομπές πελάτη σε πραγματικό χρόνο", "Ad hoc, real-time"],
    ]
    table(s, bx, by + Inches(0.65), Inches(7.55), rows, [1.05, 1.55, 0.8],
          font_size=11.8, row_h=Inches(0.5))
    # right: derived functional capabilities
    px = Inches(8.45)
    rect(s, px, by + Inches(0.65), Inches(4.35), Inches(4.35), fill=LIGHT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, px, by + Inches(0.65), Inches(0.12), Inches(4.35), fill=UP_GOLD)
    textbox(s, px + Inches(0.34), by + Inches(0.9), Inches(3.85), Inches(4.0), [
        {"t": "Λειτουργίες που προκύπτουν", "size": 15, "color": UP_RED,
         "bold": True, "font": HEAD, "space_after": 10},
        {"t": "Αναζήτηση εκπομπών ανά πελάτη & χρονικό εύρος", "bullet": True,
         "size": 13, "color": DARK, "font": BODY, "space_after": 7, "line_spacing": 1.05},
        {"t": "Ανάλυση ανά αποστολή / δρομολόγιο", "bullet": True, "size": 13,
         "color": DARK, "font": BODY, "space_after": 7, "line_spacing": 1.05},
        {"t": "Εξαγωγή δεδομένων σε Excel", "bullet": True, "size": 13,
         "color": DARK, "font": BODY, "space_after": 7, "line_spacing": 1.05},
        {"t": "Αυτόματη έκδοση πιστοποιητικών (PDF)", "bullet": True, "size": 13,
         "color": DARK, "font": BODY, "space_after": 7, "line_spacing": 1.05},
        {"t": "Self-service — χωρίς μεσολάβηση ομάδας αναφορών", "bullet": True,
         "size": 13, "color": DARK, "font": BODY, "line_spacing": 1.05},
    ])
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 5 — NON-FUNCTIONAL / TECHNICAL REQUIREMENTS
# ══════════════════════════════════════════════════════════════════════════
def s_nonfunc_req():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Προδιαγραφές · 2/2", "Μη-λειτουργικές & τεχνικές απαιτήσεις")
    bx, by, bw, bh = body_area()
    textbox(s, bx, by, Inches(12.2), Inches(0.5), [
        {"t": "Οι ανάγκες των χρηστών και τα ευρήματα GIA ορίζουν πώς πρέπει να "
              "συμπεριφέρεται το σύστημα:", "size": 14.5, "color": DARK, "font": BODY}])
    items = [
        ("Ιχνηλασιμότητα", "Πλήρης διαδρομή δεδομένων από την πηγή έως κάθε αναφορά"),
        ("Αμετάβλητο audit trail", "Καταγραφή κάθε αλλαγής — ανεξάρτητα επαληθεύσιμη"),
        ("Δεδομένα πραγματικού χρόνου", "Freshness σε ώρες αντί για εβδομάδες"),
        ("Κλιμακωσιμότητα", "1.300+ χρήστες, αυξανόμενος όγκος αιτημάτων"),
        ("Υψηλή διαθεσιμότητα", "Production-grade, χωρίς διακοπές υπηρεσίας"),
        ("Επαληθευσιμότητα & συμμόρφωση", "Έτοιμο για CSRD, EU ETS, ISCC, CDP"),
        ("Αυτοματοποίηση", "Πιστοποιητικά χωρίς ανθρώπινη παρέμβαση"),
        ("Ασφάλεια & εξουσιοδότηση", "Πρόσβαση μόνο με εταιρικά δικαιώματα"),
    ]
    cw, ch = Inches(2.95), Inches(1.32)
    gx, gy = Inches(0.16), Inches(0.16)
    x0, y0 = bx, by + Inches(0.6)
    for i, (t, d) in enumerate(items):
        col, row = i % 4, i // 4
        x = x0 + col * (cw + gx)
        y = y0 + row * (ch + gy)
        rect(s, x, y, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, x, y, Inches(0.1), ch, fill=UP_RED)
        textbox(s, x + Inches(0.26), y + Inches(0.14), cw - Inches(0.4),
                ch - Inches(0.2), [
            {"t": t, "size": 14, "color": UP_RED_D, "bold": True, "font": HEAD,
             "space_after": 4, "line_spacing": 1.0},
            {"t": d, "size": 11.8, "color": DARK, "font": BODY, "line_spacing": 1.05},
        ])
    # bottom strip: how the design answers them
    yb = y0 + 2 * (ch + gy) + Inches(0.08)
    rect(s, bx, yb, Inches(12.2), Inches(0.62), fill=PANEL,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    textbox(s, bx + Inches(0.3), yb, Inches(11.6), Inches(0.62), [
        [{"t": "Τεχνική απάντηση:  ", "size": 13.5, "color": UP_RED, "bold": True,
          "font": BODY},
         {"t": "event-driven αρχιτεκτονική · event sourcing για το audit trail · "
               "Phoenix LiveView για real-time · BEAM/Elixir για διαθεσιμότητα.",
          "size": 13.5, "color": DARK, "font": BODY}]],
        anchor=MSO_ANCHOR.MIDDLE)
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 6 — THE SOLUTION: 4 PILLARS
# ══════════════════════════════════════════════════════════════════════════
def s_pillars():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Η λύση", "Emissions Workbench — τέσσερις πυλώνες")
    bx, by, bw, bh = body_area()
    textbox(s, bx, by, Inches(12.2), Inches(0.55), [
        [{"t": "Μία πλατφόρμα ", "size": 15, "color": DARK, "font": BODY},
         {"t": "(emissions-workbench.maersk.com)", "size": 13, "color": GRAY,
          "italic": True, "font": BODY},
         {"t": " · monorepo · 4 πυλώνες με κοινή ροή δεδομένων", "size": 15,
          "color": DARK, "font": BODY}]])
    pillars = [
        ("Α · Ocean Emissions", "Αυτόματη μέτρηση εκπομπών ανά αποστολή από τηλεμετρία στόλου (STAR Connect) + EU ETS",
         "Κεφ. 5"),
        ("Β · ECO Delivery", "Εμπορικά προϊόντα χαμηλού άνθρακα · αυτόματα πιστοποιητικά · Energy Bank (mass balance)",
         "Κεφ. 6"),
        ("Γ · Net Zero 2040", "Εσωτερικοί πίνακες για οδικούς χάρτες αποκαρβονοποίησης & δεσμεύσεις SBTi",
         "Κεφ. 7"),
        ("Δ · Bunker Optimization", "Βελτιστοποίηση πλάνων ανεφοδιασμού για ελάχιστο κόστος καυσίμου σε περιβάλλον ETS",
         "Κεφ. 7"),
    ]
    cw, ch = Inches(2.92), Inches(3.45)
    gap = Inches(0.16)
    yy = by + Inches(0.75)
    for i, (t, d, k) in enumerate(pillars):
        x = bx + i * (cw + gap)
        rect(s, x, yy, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, x, yy, cw, Inches(0.85), fill=UP_RED,
             shape=MSO_SHAPE.ROUND_2_SAME_RECTANGLE)
        textbox(s, x + Inches(0.18), yy, cw - Inches(0.36), Inches(0.85),
                [{"t": t, "size": 14.5, "color": WHITE, "bold": True, "font": HEAD,
                  "align": PP_ALIGN.CENTER, "line_spacing": 1.0}],
                anchor=MSO_ANCHOR.MIDDLE)
        textbox(s, x + Inches(0.24), yy + Inches(1.0), cw - Inches(0.48),
                Inches(2.1), [
            {"t": d, "size": 13, "color": DARK, "font": BODY, "line_spacing": 1.12}])
        chip(s, x + Inches(0.24), yy + ch - Inches(0.6), Inches(1.2), k,
             color=UP_GOLD)
    # data-flow note
    yb = yy + ch + Inches(0.18)
    textbox(s, bx, yb, Inches(12.2), Inches(0.5), [
        [{"t": "Ροή δεδομένων:  ", "size": 13, "color": UP_RED, "bold": True,
          "font": BODY},
         {"t": "Ocean Emissions → κατανομή ECO · Bunker → διαθεσιμότητα καυσίμου "
               "→ Energy Bank → εκκαθάριση αξιώσεων → πιστοποιητικά  (μέσω Kafka)",
          "size": 13, "color": DARK, "font": BODY}]])
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 7 — ARCHITECTURE
# ══════════════════════════════════════════════════════════════════════════
def s_arch():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Τεχνική επισκόπηση", "Αρχιτεκτονική του συστήματος")
    bx, by, bw, bh = body_area()
    add_image(s, os.path.join(ASSETS, "system_overview_diagram.png"),
              bx, by, Inches(8.5), Inches(5.0), center_in=True)
    px = Inches(9.2)
    rect(s, px, by, Inches(3.6), Inches(5.0), fill=LIGHT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, px, by, Inches(0.12), Inches(5.0), fill=UP_GOLD)
    textbox(s, px + Inches(0.34), by + Inches(0.3), Inches(3.1), Inches(4.5), [
        {"t": "Βασικές αρχές", "size": 16, "color": UP_RED, "bold": True,
         "font": HEAD, "space_after": 12},
        {"t": "Clean Architecture", "bullet": True, "size": 13.5, "color": DARK,
         "bold": True, "font": BODY, "space_after": 1},
        {"t": "domain ανεξάρτητο από υποδομή", "size": 11.5, "color": GRAY,
         "font": BODY, "level": 1, "space_after": 9},
        {"t": "Event-driven ingestion", "bullet": True, "size": 13.5, "color": DARK,
         "bold": True, "font": BODY, "space_after": 1},
        {"t": "Kafka + Oban workers", "size": 11.5, "color": GRAY, "font": BODY,
         "level": 1, "space_after": 9},
        {"t": "Event Sourcing (Energy Bank)", "bullet": True, "size": 13.5,
         "color": DARK, "bold": True, "font": BODY, "space_after": 1},
        {"t": "πλήρες ιστορικό & audit trail", "size": 11.5, "color": GRAY,
         "font": BODY, "level": 1, "space_after": 9},
        {"t": "Real-time web (LiveView)", "bullet": True, "size": 13.5,
         "color": DARK, "bold": True, "font": BODY, "space_after": 1},
        {"t": "ενημέρωση χωρίς reload", "size": 11.5, "color": GRAY, "font": BODY,
         "level": 1, "space_after": 9},
        {"t": "Azure / Kubernetes", "bullet": True, "size": 13.5, "color": DARK,
         "bold": True, "font": BODY, "space_after": 1},
        {"t": "auto-scaling, IaC με Terraform", "size": 11.5, "color": GRAY,
         "font": BODY, "level": 1},
    ])
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDES 8–10 — PLATFORM SCREENSHOTS (placeholders)
# ══════════════════════════════════════════════════════════════════════════
def s_demo1():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Η πλατφόρμα στην πράξη · 1/3", "Αναζήτηση & ανάλυση εκπομπών")
    bx, by, bw, bh = body_area()
    add_placeholder(s, bx, by, Inches(7.4), Inches(5.0),
        "Οθόνη αναζήτησης + αποτελέσματα",
        "Εισαγωγή customer code + χρονικό εύρος και ο πίνακας/γράφημα εκπομπών "
        "που επιστρέφεται (test data). Δείχνει το Self-service & το real-time UI.")
    px = Inches(8.1)
    textbox(s, px, by + Inches(0.1), Inches(4.7), Inches(4.8), [
        {"t": "Χαρακτηριστικό 1", "size": 13, "color": UP_GOLD, "bold": True,
         "font": BODY, "space_after": 4},
        {"t": "Αναζήτηση δεδομένων εκπομπών", "size": 18, "color": UP_RED_D,
         "bold": True, "font": HEAD, "space_after": 12, "line_spacing": 1.0},
        {"t": "Customer code + χρονικό εύρος → άμεση ανάκτηση", "bullet": True,
         "size": 14, "color": DARK, "font": BODY, "space_after": 8,
         "line_spacing": 1.08},
        {"t": "Αντικαθιστά τη ροή email-και-αναμονής 3–8 εβδομάδων", "bullet": True,
         "size": 14, "color": DARK, "font": BODY, "space_after": 8,
         "line_spacing": 1.08},
        {"t": "Polling από Dremio μέσω Oban → freshness σε ώρες", "bullet": True,
         "size": 14, "color": DARK, "font": BODY, "space_after": 18,
         "line_spacing": 1.08},
        {"t": "Χαρακτηριστικό 2", "size": 13, "color": UP_GOLD, "bold": True,
         "font": BODY, "space_after": 4},
        {"t": "Ανάλυση ανά αποστολή / δρομολόγιο", "size": 18, "color": UP_RED_D,
         "bold": True, "font": HEAD, "space_after": 10, "line_spacing": 1.0},
        {"t": "Ποια δρομολόγια βαραίνουν το αποτύπωμα του πελάτη — βάση για ECO",
         "bullet": True, "size": 14, "color": DARK, "font": BODY,
         "line_spacing": 1.08},
    ])
    footer(s)


def s_demo2():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Η πλατφόρμα στην πράξη · 2/3", "Πιστοποιητικά & εξαγωγή")
    bx, by, bw, bh = body_area()
    half = Inches(5.95)
    add_placeholder(s, bx, by + Inches(0.1), half, Inches(2.3),
        "Δημιουργία πιστοποιητικού (Certificates)",
        "Το εργαλείο Certificates με την πρόοδο των Oban workers που αντλούν "
        "από το Energy Bank (test data).")
    add_placeholder(s, bx, by + Inches(2.55), half, Inches(2.3),
        "Εξαγωγή σε Excel",
        "Η οθόνη export: πλήρη δεδομένα αποστολών + emissions saving data "
        "για πελάτες ECO.")
    add_placeholder(s, bx + half + Inches(0.3), by + Inches(0.1),
        Inches(5.95), Inches(4.75),
        "Παραγόμενο πιστοποιητικό (PDF) + αναφορά",
        "Το τελικό PDF πιστοποιητικό με τη συνοδευτική αναφορά μεθοδολογίας "
        "(ISCC / mass balance). Ολόκληρη η ροή εκτελείται αυτόματα — από "
        "εβδομάδες σε λεπτά.")
    footer(s)


def s_demo3():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Η πλατφόρμα στην πράξη · 3/3", "Energy Bank, Bunker & Observability")
    bx, by, bw, bh = body_area()
    third = Inches(3.92)
    gap = Inches(0.22)
    add_placeholder(s, bx, by + Inches(0.1), third, Inches(4.7),
        "Energy Bank — ισοζύγιο",
        "Καταθέσεις/αναλήψεις βιώσιμου καυσίμου & υπόλοιπα λογαριασμών "
        "(mass balance) — event-sourced ledger.")
    add_placeholder(s, bx + third + gap, by + Inches(0.1), third, Inches(4.7),
        "Bunker Optimization — πλάνο",
        "Αποτέλεσμα του solver: βέλτιστο πλάνο ανεφοδιασμού ανά λιμάνι/δρομολόγιο "
        "(BOPS / BoW).")
    add_placeholder(s, bx + 2 * (third + gap), by + Inches(0.1), third, Inches(4.7),
        "Observability dashboard",
        "Grafana: ρυθμός αναπτύξεων & υγεία συστήματος — τεκμηριώνει τα ~32 "
        "deploys/ημέρα και τη διαθεσιμότητα.")
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 11 — THEORY (compressed to one slide)
# ══════════════════════════════════════════════════════════════════════════
def s_theory():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Θεωρητικό υπόβαθρο (σύνοψη)", "Λογιστική εκπομπών — τα βασικά")
    bx, by, bw, bh = body_area()
    add_image(s, os.path.join(ASSETS, "Carbon_Accounting_Scopes.png"),
              bx, by + Inches(0.2), Inches(6.2), Inches(4.6), center_in=True)
    px = Inches(7.1)
    concepts = [
        ("Scope 1 / 2 / 3", "Άμεσες · έμμεσες από ενέργεια · αλυσίδα αξίας"),
        ("TTW vs WTW", "Όρια συστήματος: Tank-to-Wake vs Well-to-Wake"),
        ("Trade factors", "Συντελεστές εμπορικής διαδρομής — ο πυρήνας του υπολογισμού"),
        ("Mass balance", "Λογιστική σύνδεση βιώσιμου καυσίμου με αξιώσεις"),
        ("Book-and-claim", "Αναφορά αποσυνδεδεμένη από τη φυσική ροή"),
        ("ISCC / GLEC / GHG Protocol", "Τα πρότυπα πιστοποίησης & μεθοδολογίας"),
    ]
    items = [{"t": "Από τη θεωρία στον κώδικα", "size": 16, "color": UP_RED,
              "bold": True, "font": HEAD, "space_after": 12}]
    for t, d in concepts:
        items.append([{"t": t + "  —  ", "size": 14, "color": UP_RED_D,
                       "bold": True, "font": BODY},
                      {"t": d, "size": 13, "color": DARK, "font": BODY}])
        items[-1][0]["bullet"] = True
        items[-1][0]["space_after"] = 9
        items[-1][0]["line_spacing"] = 1.05
    textbox(s, px, by + Inches(0.15), Inches(5.6), Inches(4.6), items)
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 12 — TECHNOLOGIES
# ══════════════════════════════════════════════════════════════════════════
def s_tech():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Τεχνική επισκόπηση", "Τεχνολογίες")
    bx, by, bw, bh = body_area()
    # ── hero row: the three flagship technologies, with logos ──
    heroes = [
        ("logo-elixir.png", "BEAM / Elixir", "Υψηλή διαθεσιμότητα & ταυτοχρονισμός"),
        ("logo-phoenix.png", "Phoenix LiveView", "Real-time web χωρίς reload σελίδας"),
        ("logo-nix.png", "Nix / Flakes", "Reproducible builds & dev environments"),
    ]
    hcw = Inches(3.92)
    hgx = Inches(0.23)
    hch = Inches(2.35)
    hy = by + Inches(0.05)
    for i, (logo, t, d) in enumerate(heroes):
        x = bx + i * (hcw + hgx)
        rect(s, x, hy, hcw, hch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, x, hy, hcw, Inches(0.1), fill=UP_GOLD)
        add_image(s, os.path.join(ASSETS, logo), x + Inches(0.4),
                  hy + Inches(0.32), hcw - Inches(0.8), Inches(0.95))
        textbox(s, x + Inches(0.2), hy + Inches(1.5), hcw - Inches(0.4),
                Inches(0.8), [
            {"t": t, "size": 17, "color": UP_RED_D, "bold": True, "font": HEAD,
             "align": PP_ALIGN.CENTER, "space_after": 3},
            {"t": d, "size": 12.5, "color": DARK, "font": BODY,
             "align": PP_ALIGN.CENTER, "line_spacing": 1.05},
        ])
    # ── supporting stack: compact cards ──
    techs = [
        ("PostgreSQL / Ecto", "Σχεσιακή μονιμότητα δεδομένων"),
        ("Apache Kafka", "Event-driven ροή μηνυμάτων"),
        ("Commanded / EventStore", "Event sourcing & CQRS (Energy Bank)"),
        ("Oban", "Αξιόπιστες ασύγχρονες εργασίες"),
        ("Dremio", "Data lake — πηγή δεδομένων"),
        ("Azure · Kubernetes · Terraform", "Cloud, auto-scaling & IaC"),
    ]
    cw, ch = Inches(3.92), Inches(1.05)
    gx, gy = Inches(0.23), Inches(0.18)
    y0 = hy + hch + Inches(0.22)
    for i, (t, d) in enumerate(techs):
        col, row = i % 3, i // 3
        x = bx + col * (cw + gx)
        y = y0 + row * (ch + gy)
        rect(s, x, y, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, x, y, Inches(0.1), ch, fill=UP_GOLD)
        textbox(s, x + Inches(0.26), y + Inches(0.13), cw - Inches(0.42),
                ch - Inches(0.2), [
            {"t": t, "size": 13.5, "color": UP_RED_D, "bold": True, "font": HEAD,
             "space_after": 3},
            {"t": d, "size": 11.8, "color": DARK, "font": BODY, "line_spacing": 1.0},
        ])
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 13 — WORKING METHODS (XP)
# ══════════════════════════════════════════════════════════════════════════
def s_methods():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Εργασιακές μέθοδοι", "Extreme Programming σε ομάδα 28 μελών")
    bx, by, bw, bh = body_area()
    add_image(s, os.path.join(ASSETS, "Extreme_Programming_Loops.png"),
              bx, by, Inches(5.2), Inches(5.0), center_in=True)
    px = Inches(6.0)
    practices = [
        ("Pair programming", "συνεχής έλεγχος κώδικα & μεταφορά γνώσης"),
        ("TDD", "ανάπτυξη καθοδηγούμενη από δοκιμές — 80%+ κάλυψη"),
        ("CI/CD", "trunk-based, ~32 αναπτύξεις/ημέρα"),
        ("Vertical ownership", "κάθε μηχανικός: από τη βάση δεδομένων έως το UI"),
        ("Feedback loops", "γρήγορη ανατροφοδότηση σε κάθε επίπεδο"),
    ]
    items = [{"t": "Πρακτικές που έκαναν εφικτό τον κύκλο 6 μηνών", "size": 15.5,
              "color": UP_RED, "bold": True, "font": HEAD, "space_after": 12,
              "line_spacing": 1.0}]
    for t, d in practices:
        items.append([{"t": t + "  —  ", "size": 14.5, "color": UP_RED_D,
                       "bold": True, "font": BODY},
                      {"t": d, "size": 13.5, "color": DARK, "font": BODY}])
        items[-1][0]["bullet"] = True
        items[-1][0]["space_after"] = 11
        items[-1][0]["line_spacing"] = 1.05
    textbox(s, px, by + Inches(0.2), Inches(6.6), Inches(4.6), items)
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 14 — RESULTS: BIG NUMBERS
# ══════════════════════════════════════════════════════════════════════════
def s_results_big():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=UP_RED)
    rect(s, 0, Inches(1.12), SW, Inches(0.07), fill=UP_GOLD)
    textbox(s, Inches(0.55), Inches(0.13), Inches(11), Inches(0.3),
            [{"t": "ΑΠΟΤΕΛΕΣΜΑΤΑ", "size": 11, "color": UP_GOLD, "bold": True,
              "font": BODY}])
    textbox(s, Inches(0.55), Inches(0.40), Inches(12.2), Inches(0.66),
            [{"t": "Μετρήσιμος αντίκτυπος", "size": 25, "color": WHITE, "bold": True,
              "font": HEAD}], anchor=MSO_ANCHOR.MIDDLE)
    cards = [
        ("1.300+", "ενεργοί χρήστες", "από <100 στην εποχή Excel"),
        ("$31 εκ.", "ECO Delivery 2024", "98.000 FFE εμπορευμάτων"),
        ("6 μήνες", "έως παραγωγή", "Μάρτιος → Σεπτέμβριος 2023"),
        ("~32", "αναπτύξεις / ημέρα", "μία κάθε ~15 λεπτά"),
        ("720+", "schema migrations", "χωρίς διακοπή υπηρεσίας"),
        ("100%", "ευρημάτων GIA", "πλήρης αντιμετώπιση"),
    ]
    cw, ch = Inches(3.92), Inches(2.45)
    gx, gy = Inches(0.22), Inches(0.25)
    x0, y0 = Inches(0.55), Inches(1.55)
    for i, (big, lab, sub) in enumerate(cards):
        col, row = i % 3, i // 3
        x = x0 + col * (cw + gx)
        y = y0 + row * (ch + gy)
        rect(s, x, y, cw, ch, fill=WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, x, y, cw, Inches(0.12), fill=UP_GOLD)
        textbox(s, x + Inches(0.25), y + Inches(0.28), cw - Inches(0.5),
                ch - Inches(0.4), [
            {"t": big, "size": 44, "color": UP_RED, "bold": True, "font": HEAD,
             "align": PP_ALIGN.CENTER, "space_after": 2},
            {"t": lab, "size": 16, "color": DARK, "bold": True, "font": BODY,
             "align": PP_ALIGN.CENTER, "space_after": 4},
            {"t": sub, "size": 12, "color": GRAY, "italic": True, "font": BODY,
             "align": PP_ALIGN.CENTER},
        ], anchor=MSO_ANCHOR.MIDDLE)
    footer(s, dark=True)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 15 — RESULTS: FULL METRICS TABLE
# ══════════════════════════════════════════════════════════════════════════
def s_results_table():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Αποτελέσματα", "Πλήρης πίνακας μετρικών")
    bx, by, bw, bh = body_area()
    rows = [
        ["Διάσταση", "Μετρική", "Τιμή"],
        ["Υιοθέτηση", "Ενεργοί χρήστες", "1.300+ (από <100)"],
        ["Εμπορικός αντίκτυπος", "ECO Delivery Ocean 2024", "98.000 FFE / $31 εκ."],
        ["Κανονιστική συμμόρφωση", "Έλεγχος GIA 2023", "Πλήρης αντιμετώπιση"],
        ["Ταχύτητα παράδοσης", "Αναπτύξεις σε παραγωγή", "~32 / ημέρα"],
        ["Αξιοπιστία", "Schema migrations χωρίς διακοπή", "720+"],
        ["Έλεγχος ποιότητας", "Κάλυψη αυτοματοποιημένων tests", "80%+"],
        ["Κλίμακα κώδικα", "Αρχεία πηγαίου κώδικα", "~6.000"],
        ["Κύκλος ανάπτυξης", "Από έναρξη έως παραγωγή", "6 μήνες"],
    ]
    table(s, bx, by + Inches(0.1), Inches(8.1), rows, [1.15, 1.6, 1.1],
          font_size=13, row_h=Inches(0.5))
    px = Inches(9.0)
    rect(s, px, by + Inches(0.1), Inches(3.8), Inches(4.65), fill=LIGHT,
         shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    rect(s, px, by + Inches(0.1), Inches(0.12), Inches(4.65), fill=UP_GOLD)
    textbox(s, px + Inches(0.34), by + Inches(0.4), Inches(3.3), Inches(4.2), [
        {"t": "Με μια ματιά", "size": 16, "color": UP_RED, "bold": True,
         "font": HEAD, "space_after": 12},
        {"t": "Από proof-of-concept σε production με μετρήσιμη εμπορική αξία σε "
              "μήνες.", "bullet": True, "size": 13.5, "color": DARK, "font": BODY,
         "space_after": 10, "line_spacing": 1.12},
        {"t": "Ταχύτητα και ποιότητα δεν είναι αντίθετοι στόχοι.", "bullet": True,
         "size": 13.5, "color": DARK, "font": BODY, "space_after": 10,
         "line_spacing": 1.12},
        {"t": "Ο τεχνικός σχεδιασμός υπηρετεί και την εταιρική διακυβέρνηση.",
         "bullet": True, "size": 13.5, "color": DARK, "font": BODY,
         "line_spacing": 1.12},
    ])
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 16 — LESSONS
# ══════════════════════════════════════════════════════════════════════════
def s_lessons():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Συμπεράσματα", "Μαθήματα & συμβιβασμοί (trade-offs)")
    bx, by, bw, bh = body_area()
    lessons = [
        ("Vertical ownership vs γνωσιακό φορτίο",
         "Μεγάλη συνοχή, αλλά απαιτεί ευρεία γνώση — αντίμετρο το pairing & onboarding"),
        ("Monorepo: ισχύς με κόστος",
         "Ατομικές αναπτύξεις & refactoring, αλλά αυξανόμενος χρόνος build/CI"),
        ("Event sourcing: επιλεκτικά",
         "Ιδανικό για audit-heavy προβλήματα (Energy Bank), αλλά αυξάνει την καμπύλη μάθησης"),
        ("Cloud vs self-hosting",
         "Ευελιξία κλιμάκωσης, αλλά κόστος εξαρτώμενο από τον όγκο — απαιτεί σχεδιασμό"),
        ("Ταχύτητα ως πλεονέκτημα",
         "TDD + pair programming + CD: η ποιότητα επιτρέπει την ταχύτητα, δεν την εμποδίζει"),
    ]
    yy = by + Inches(0.15)
    for i, (t, d) in enumerate(lessons):
        rect(s, bx, yy, Inches(12.2), Inches(0.82), fill=LIGHT if i % 2 == 0 else WHITE,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        rect(s, bx, yy, Inches(0.1), Inches(0.82), fill=UP_RED)
        textbox(s, bx + Inches(0.3), yy, Inches(11.7), Inches(0.82), [
            [{"t": t + "    ", "size": 15, "color": UP_RED_D, "bold": True,
              "font": HEAD},
             {"t": d, "size": 13, "color": DARK, "font": BODY}]],
            anchor=MSO_ANCHOR.MIDDLE)
        yy = yy + Inches(0.92)
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 17 — FUTURE WORK
# ══════════════════════════════════════════════════════════════════════════
def s_future():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=WHITE)
    header(s, "Μελλοντική εργασία", "Επόμενα βήματα")
    bx, by, bw, bh = body_area()
    future = [
        ("Scope 2 εκπομπές", "Γραφεία, τερματικά και cold ironing πλοίων"),
        ("Επέκταση κάλυψης στόλου", "Περισσότερες δραστηριότητες & τύποι μεταφοράς"),
        ("Επέκταση αναφοράς SBTi", "Πληρέστερη παρακολούθηση κλιματικών δεσμεύσεων"),
        ("Προβλεπτικά analytics", "Σενάρια αποκαρβονοποίησης & πρόβλεψη εκπομπών"),
        ("Bunker ↔ Energy Bank", "Σύνδεση βελτιστοποίησης με τη λογιστική πράσινου καυσίμου"),
    ]
    cw, ch = Inches(3.92), Inches(2.0)
    gx, gy = Inches(0.22), Inches(0.25)
    x0, y0 = bx, by + Inches(0.3)
    for i, (t, d) in enumerate(future):
        col, row = i % 3, i // 3
        x = x0 + col * (cw + gx)
        y = y0 + row * (ch + gy)
        rect(s, x, y, cw, ch, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        textbox(s, x + Inches(0.28), y + Inches(0.22), cw - Inches(0.5),
                ch - Inches(0.4), [
            {"t": "→", "size": 24, "color": UP_GOLD, "bold": True, "font": HEAD,
             "space_after": 4},
            {"t": t, "size": 15.5, "color": UP_RED_D, "bold": True, "font": HEAD,
             "space_after": 6, "line_spacing": 1.0},
            {"t": d, "size": 12.5, "color": DARK, "font": BODY, "line_spacing": 1.08},
        ])
    footer(s)


# ══════════════════════════════════════════════════════════════════════════
#  SLIDE 18 — CLOSING
# ══════════════════════════════════════════════════════════════════════════
def s_closing():
    s = slide()
    rect(s, 0, 0, SW, SH, fill=LIGHT)
    rect(s, 0, 0, Inches(0.32), SH, fill=UP_RED)
    rect(s, Inches(0.32), 0, Inches(0.07), SH, fill=UP_GOLD)
    add_image(s, os.path.join(ASSETS, "logo-up-4color-landscape.jpg"),
              Inches(5.55), Inches(0.85), Inches(2.2), Inches(2.2), center_in=False)
    textbox(s, Inches(1.0), Inches(3.35), Inches(11.3), Inches(1.2), [
        {"t": "Σας ευχαριστώ", "size": 40, "color": UP_RED_D, "bold": True,
         "font": HEAD, "align": PP_ALIGN.CENTER, "space_after": 6},
        {"t": "Ερωτήσεις & συζήτηση", "size": 20, "color": GRAY, "italic": True,
         "font": BODY, "align": PP_ALIGN.CENTER}])
    rect(s, Inches(4.6), Inches(4.95), Inches(4.1), Pt(2), fill=UP_GOLD)
    textbox(s, Inches(1.0), Inches(5.3), Inches(11.3), Inches(1.4), [
        {"t": "Κωνσταντίνος Γκίνης", "size": 17, "color": DARK, "bold": True,
         "font": HEAD, "align": PP_ALIGN.CENTER, "space_after": 4},
        {"t": "Επιβλέπων: Χρήστος Μπούρας, Καθηγητής", "size": 13, "color": GRAY,
         "font": BODY, "align": PP_ALIGN.CENTER, "space_after": 2},
        {"t": "Πανεπιστήμιο Πατρών · Τμήμα Μηχανικών Η/Υ & Πληροφορικής · Μάιος 2026",
         "size": 12, "color": GRAY, "font": BODY, "align": PP_ALIGN.CENTER}])


# ── build ─────────────────────────────────────────────────────────────────
s_title()
s_problem()
s_audit()
s_func_req()
s_nonfunc_req()
s_pillars()
s_arch()
s_demo1()
s_demo2()
s_demo3()
s_theory()
s_tech()
s_methods()
s_results_big()
s_results_table()
s_lessons()
s_future()
s_closing()


# ── Speaker notes (Greek) — budgeted to ~19 min + buffer for a 20΄ defense ──
NOTES = [
    # 1 — Title
    """⏱ ~30΄΄  |  Σύνολο στόχος: 20 λεπτά.
Καλησπέρα σας. Ονομάζομαι Κωνσταντίνος Γκίνης. Η διπλωματική μου παρουσιάζει
τον «Πάγκο Εργασίας Εκπομπών» (Emissions Workbench): μια πλατφόρμα λογισμικού
παραγωγικής κλίμακας που σχεδίασε και υλοποίησε η ομάδα Energy Transition
Platform της Maersk, για τη μέτρηση, πιστοποίηση και βελτιστοποίηση των εκπομπών
αερίων θερμοκηπίου στη ναυτιλία. Επιβλέπων ο καθηγητής κ. Χρήστος Μπούρας.
→ Ξεκινώ από το πρόβλημα που γέννησε την ανάγκη.""",
    # 2 — Problem
    """⏱ ~70΄΄
Η Maersk, με στόλο 700+ πλοίων, δεν μπορούσε να ανταποκριθεί αξιόπιστα στις
αυξανόμενες απαιτήσεις για δεδομένα εκπομπών. Η διαδικασία ήταν εξ ολοκλήρου
χειρωνακτική: αίτημα μέσω email, μη αυτόματη εξαγωγή δεδομένων, εφαρμογή
συντελεστών σε φύλλα Excel, και απάντηση μετά από 3 έως 8 εβδομάδες.
Αυτό κόστιζε σε τέσσερις άξονες: ΑΚΡΙΒΕΙΑ (λάθη μεθοδολογίας), ΤΑΧΥΤΗΤΑ
(καθυστερήσεις που σήμαιναν απώλεια συμβολαίων), ΙΧΝΗΛΑΣΙΜΟΤΗΤΑ (κανένα
audit) και ΚΛΙΜΑΚΑ (η ανθρώπινη δυναμικότητα δεν κλιμακώνεται για 1.300+ χρήστες).
→ Το πρόβλημα ήταν γνωστό· χρειαζόταν όμως ένας καταλύτης.""",
    # 3 — Audit
    """⏱ ~70΄΄
Αυτός ο καταλύτης ήταν ο εσωτερικός έλεγχος Group Internal Audit του 2023.
Εντόπισε τέσσερα συγκεκριμένα ευρήματα: έλλειψη παρακολούθησης δεδομένων,
απουσία ίχνους ελέγχου (κρίσιμο εν όψει CSRD και EU ETS), αδυναμία κλιμάκωσης,
και υψηλό ποσοστό σφαλμάτων. Η σημασία του ελέγχου ήταν ότι έδωσε επίσημο
βάρος σε ένα ήδη γνωστό πρόβλημα — και έτσι εγκρίθηκαν πόροι για ένα εξ
ολοκλήρου νέο σύστημα. Η ανάπτυξη ξεκίνησε τον Μάρτιο 2023 και έφτασε σε
παραγωγή τον Σεπτέμβριο 2023: μόλις έξι μήνες.
→ Ας δούμε τι έπρεπε να κάνει αυτό το σύστημα — τις προδιαγραφές.""",
    # 4 — Functional requirements
    """⏱ ~75΄΄
Οι λειτουργικές απαιτήσεις προέκυψαν από έξι διακριτές ομάδες χρηστών, η καθεμία
με διαφορετική ανάγκη και συχνότητα: από τα ECO Delivery Products που θέλουν
αυτόματα πιστοποιητικά πολλές φορές την ημέρα, μέχρι τους Account Managers που
χρειάζονται δεδομένα σε πραγματικό χρόνο μέσα σε μια διαπραγμάτευση.
Από αυτή την ποικιλομορφία προκύπτουν οι βασικές λειτουργίες του συστήματος:
αναζήτηση εκπομπών ανά πελάτη και χρονικό εύρος, ανάλυση ανά αποστολή και
δρομολόγιο, εξαγωγή σε Excel, αυτόματη έκδοση πιστοποιητικών — και όλα αυτά
σε λογική self-service, χωρίς μεσολάβηση ειδικής ομάδας αναφορών.
→ Πέρα όμως από το «τι κάνει», υπάρχει το «πώς πρέπει να συμπεριφέρεται».""",
    # 5 — Non-functional requirements
    """⏱ ~75΄΄
Εδώ είναι το σημείο που ζήτησε ιδιαίτερα η επιτροπή: οι μη-λειτουργικές, τεχνικές
απαιτήσεις. Συνδυάζοντας τις ανάγκες των χρηστών με τα ευρήματα GIA, το σύστημα
έπρεπε να εγγυάται: πλήρη ιχνηλασιμότητα δεδομένων· αμετάβλητο audit trail·
δεδομένα πραγματικού χρόνου (freshness σε ώρες, όχι εβδομάδες)· κλιμακωσιμότητα·
υψηλή διαθεσιμότητα· επαληθευσιμότητα για CSRD, EU ETS, ISCC· πλήρη
αυτοματοποίηση· και ασφαλή εξουσιοδότηση.
Καθεμία από αυτές καθόρισε μια αρχιτεκτονική επιλογή — που φαίνεται στην κάτω
γραμμή: event-driven αρχιτεκτονική, event sourcing για το audit trail, Phoenix
LiveView για real-time, και BEAM/Elixir για διαθεσιμότητα.
→ Δείτε πώς υλοποιήθηκαν αυτές οι απαιτήσεις — η λύση.""",
    # 6 — Pillars
    """⏱ ~80΄΄
Η λύση είναι μία ενιαία πλατφόρμα, σε ένα monorepo, οργανωμένη σε τέσσερις
πυλώνες. ΠΥΛΩΝΑΣ Α – Ocean Emissions: αυτόματη μέτρηση εκπομπών ανά αποστολή,
από τηλεμετρία στόλου (STAR Connect) και ενσωμάτωση EU ETS. ΠΥΛΩΝΑΣ Β – ECO
Delivery: τα εμπορικά προϊόντα χαμηλού άνθρακα, με αυτόματα πιστοποιητικά και τη
λογιστική πράσινου καυσίμου του Energy Bank. ΠΥΛΩΝΑΣ Γ – Net Zero 2040: εσωτερικοί
πίνακες για οδικούς χάρτες αποκαρβονοποίησης. ΠΥΛΩΝΑΣ Δ – Bunker Optimization:
βελτιστοποίηση πλάνων ανεφοδιασμού για ελάχιστο κόστος καυσίμου.
Σημαντικό: οι πυλώνες δεν είναι ανεξάρτητοι — επικοινωνούν μέσω Kafka, όπως
δείχνει η ροή δεδομένων στο κάτω μέρος.
→ Πώς οργανώνεται τεχνικά όλο αυτό;""",
    # 7 — Architecture
    """⏱ ~70΄΄
Η αρχιτεκτονική στηρίζεται σε πέντε αρχές. Clean Architecture: ο επιχειρησιακός
πυρήνας είναι ανεξάρτητος από την υποδομή. Event-driven ingestion μέσω Kafka και
εργασιών Oban. Event Sourcing ειδικά για το Energy Bank, που δίνει πλήρες ιστορικό
και το αμετάβλητο audit trail. Real-time web με Phoenix LiveView, χωρίς reload
σελίδας. Και ανάπτυξη σε Azure/Kubernetes με auto-scaling και infrastructure-as-code
μέσω Terraform. Το διάγραμμα δείχνει τη συνολική ροή — από τις πηγές δεδομένων,
μέσα από τα components, έως το επίπεδο ιστού.
→ Ας δούμε τώρα την πλατφόρμα στην πράξη.""",
    # 8 — Demo 1
    """⏱ ~70΄΄  [όταν μπουν τα screenshots, δείξε τη ζωντανή ροή]
Αυτά είναι τα δύο πρώτα από τα πέντε χαρακτηριστικά του κύκλου αξίας.
ΧΑΡΑΚΤΗΡΙΣΤΙΚΟ 1 – Αναζήτηση: ο χρήστης εισάγει κωδικό πελάτη και χρονικό εύρος,
και τα δεδομένα εκπομπών επιστρέφονται άμεσα — αντικαθιστώντας τη ροή
email-και-αναμονής 3 έως 8 εβδομάδων. Τεχνικά, γίνεται polling από το Dremio μέσω
εργασιών Oban, που εξασφαλίζει freshness σε ώρες.
ΧΑΡΑΚΤΗΡΙΣΤΙΚΟ 2 – Ανάλυση ανά αποστολή και δρομολόγιο: ο πελάτης βλέπει ποια
δρομολόγια βαραίνουν περισσότερο το αποτύπωμά του — που είναι η βάση για να
προτείνουμε υπηρεσίες ECO.
→ Από τα δεδομένα, περνάμε στα πιστοποιητικά.""",
    # 9 — Demo 2
    """⏱ ~60΄΄
ΧΑΡΑΚΤΗΡΙΣΤΙΚΟ 4 – Δημιουργία πιστοποιητικών: το εργαλείο Certificates ενεργοποιεί
εργασίες Oban που αντλούν από το ισοζύγιο του Energy Bank, επαληθεύουν τα διαθέσιμα
υπόλοιπα βιώσιμου καυσίμου, και παράγουν έγγραφα συμμόρφωσης ISCC. ΧΑΡΑΚΤΗΡΙΣΤΙΚΟ
5 – το τελικό PDF συνοδεύεται από αναλυτική αναφορά μεθοδολογίας, ώστε ο πελάτης
να το χρησιμοποιήσει σε δικές του αναφορές CSRD. Όλη αυτή η ροή, που πριν απαιτούσε
χειρωνακτική εργασία εβδομάδων, εκτελείται πλέον αυτόματα — σε λεπτά. Αριστερά κάτω,
το ΧΑΡΑΚΤΗΡΙΣΤΙΚΟ 3: εξαγωγή σε Excel, τη μορφή που χρησιμοποιούν στην πράξη οι
τελικοί χρήστες.
→ Πίσω από αυτά υπάρχουν ακόμη τρία σημαντικά υποσυστήματα.""",
    # 10 — Demo 3
    """⏱ ~60΄΄
Τρεις ακόμη όψεις. Το Energy Bank: ένα event-sourced ledger με καταθέσεις και
αναλήψεις βιώσιμου καυσίμου, που επιβάλλει τον περιορισμό του ισοζυγίου μάζας.
Το Bunker Optimization: το αποτέλεσμα του solver — το βέλτιστο πλάνο ανεφοδιασμού
ανά λιμάνι και δρομολόγιο. Και το Observability dashboard, που τεκμηριώνει τον
ρυθμό αναπτύξεων και την υγεία του συστήματος — θα επανέλθω σε αυτό στα
αποτελέσματα.
→ Πριν τα αποτελέσματα, μια σύντομη θεωρητική βάση.""",
    # 11 — Theory
    """⏱ ~60΄΄  [σύντομα — η επιτροπή ζήτησε λιγότερη θεωρία]
Πολύ συνοπτικά το θεωρητικό υπόβαθρο. Οι εκπομπές κατηγοριοποιούνται σε Scope 1,
2 και 3 (άμεσες, έμμεσες από ενέργεια, αλυσίδας αξίας). Τα όρια του συστήματος
ορίζονται ως Tank-to-Wake ή Well-to-Wake. Ο πυρήνας του υπολογισμού είναι οι
trade factors — οι συντελεστές εμπορικής διαδρομής. Και η λογιστική πράσινου
καυσίμου στηρίζεται στο mass balance και στο book-and-claim, υπό τα πρότυπα
ISCC, GLEC και το GHG Protocol.
Το ουσιώδες: αυτές οι έννοιες δεν έμειναν θεωρητικές — αντιστοιχήθηκαν άμεσα σε
σχεδιαστικά πρότυπα μέσα στον κώδικα.
→ Ποια τεχνολογία τα υλοποιεί;""",
    # 12 — Tech
    """⏱ ~60΄΄
Η στοίβα στηρίζεται σε τρεις πυλώνες: την πλατφόρμα BEAM/Elixir για διαθεσιμότητα
και ταυτοχρονισμό· το Phoenix LiveView για real-time διεπαφή χωρίς reload· και το
Nix για αναπαραγώγιμα περιβάλλοντα. Τα υποστηρικτικά: PostgreSQL/Ecto για
μονιμότητα, Apache Kafka για τη ροή συμβάντων, Commanded/EventStore για το event
sourcing του Energy Bank, Oban για ασύγχρονες εργασίες, Dremio ως data lake, και
Azure με Kubernetes και Terraform για deployment.
→ Εξίσου σημαντικό με την τεχνολογία ήταν ο τρόπος εργασίας της ομάδας.""",
    # 13 — Methods
    """⏱ ~70΄΄
Η ομάδα των 28 μελών εφάρμοσε Extreme Programming. Pair programming: συνεχής
έλεγχος κώδικα και μεταφορά γνώσης. TDD, με κάλυψη 80%+. Continuous Integration/
Delivery σε λογική trunk-based, που έφτανε τις ~32 αναπτύξεις την ημέρα. Vertical
ownership: κάθε μηχανικός αναλαμβάνει λειτουργικότητα από τη βάση δεδομένων έως το
UI. Και σφιχτοί βρόχοι ανατροφοδότησης σε κάθε επίπεδο. Αυτές οι πρακτικές είναι
που έκαναν εφικτό τον κύκλο των έξι μηνών έως την παραγωγή.
→ Και ερχόμαστε στο πιο σημαντικό: τα αποτελέσματα.""",
    # 14 — Results (big numbers)
    """⏱ ~80΄΄  [η κορυφαία διαφάνεια — δώσε χρόνο και έμφαση]
Ο αντίκτυπος είναι μετρήσιμος. 1.300+ ενεργοί χρήστες — από λιγότερους από 100
στην εποχή του Excel. ECO Delivery αξίας 31 εκατομμυρίων δολαρίων, ή 98.000 FFE,
μόνο για το 2024 — μια εμπορική δραστηριότητα που πριν δεν ήταν τεχνικά εφικτή σε
αυτή την κλίμακα. Από την έναρξη έως την παραγωγή: έξι μήνες. ~32 αναπτύξεις την
ημέρα — μία κάθε δεκαπέντε λεπτά. 720+ schema migrations χωρίς καμία διακοπή
υπηρεσίας. Και πλήρης, 100%, αντιμετώπιση όλων των ευρημάτων του ελέγχου GIA.
→ Ο πλήρης πίνακας μετρικών.""",
    # 15 — Results table
    """⏱ ~60΄΄
Ο συγκεντρωτικός πίνακας προσθέτει και την τεχνική ωριμότητα: 80%+ κάλυψη tests,
~6.000 αρχεία πηγαίου κώδικα. Τρία συμπεράσματα αξίζει να κρατήσετε: πρώτον, η
πλατφόρμα πέρασε από proof-of-concept σε παραγωγή με μετρήσιμη εμπορική αξία μέσα
σε μήνες. Δεύτερον, ταχύτητα και ποιότητα δεν είναι αντίθετοι στόχοι. Και τρίτον, ο
τεχνικός σχεδιασμός — event sourcing, audit trail — δεν υπηρετεί μόνο λειτουργικούς
στόχους, αλλά και την εταιρική διακυβέρνηση.
→ Τι μάθαμε από την υλοποίηση;""",
    # 16 — Lessons
    """⏱ ~70΄΄
Πέντε μαθήματα, κυρίως συμβιβασμοί. Το vertical ownership δίνει συνοχή, αλλά αυξάνει
το γνωσιακό φορτίο — αντίμετρο το pairing και το onboarding. Το monorepo προσφέρει
ατομικές αναπτύξεις, με κόστος χρόνου build. Το event sourcing είναι ιδανικό για
audit-heavy προβλήματα όπως το Energy Bank, αλλά πρέπει να εφαρμόζεται επιλεκτικά.
Το cloud δίνει ευελιξία, με κόστος που απαιτεί σχεδιασμό. Και το βασικότερο: η
ταχύτητα παράδοσης, όταν στηρίζεται σε TDD και pair programming, γίνεται
ανταγωνιστικό πλεονέκτημα — η ποιότητα επιτρέπει την ταχύτητα, δεν την εμποδίζει.
→ Πού πηγαίνει το έργο από εδώ;""",
    # 17 — Future work
    """⏱ ~50΄΄
Οι κυριότερες επεκτάσεις που ήδη σχεδιάζονται: κάλυψη εκπομπών Scope 2 (γραφεία,
τερματικά, cold ironing)· επέκταση κάλυψης στόλου και δραστηριοτήτων· πληρέστερη
αναφορά SBTi· προβλεπτικά analytics με σενάρια αποκαρβονοποίησης· και στενότερη
σύνδεση του Bunker Optimization με τη λογιστική του Energy Bank.
→ Κλείνω.""",
    # 18 — Closing
    """⏱ ~30΄΄
Συνοψίζοντας: ο Πάγκος Εργασίας Εκπομπών αποδεικνύει ότι ένα σύστημα μπορεί να
συνδυάσει επιστημονική ακρίβεια, εμπορική αξιοπιστία και υψηλό ρυθμό παράδοσης —
αρκεί οι τεχνολογικές, αρχιτεκτονικές και μεθοδολογικές επιλογές να αντιμετωπίζονται
ως ενιαίο σύστημα. Σας ευχαριστώ για την προσοχή σας· είμαι στη διάθεσή σας για
ερωτήσεις.""",
]

for _slide, _note in zip(prs.slides, NOTES):
    _slide.notes_slide.notes_text_frame.text = _note

prs.save(OUT)
print("Saved:", OUT, "| slides:", len(prs.slides._sldIdLst),
      "| notes:", len(NOTES))
