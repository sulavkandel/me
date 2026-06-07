#!/usr/bin/env python3
"""
Generate branded .docx and .pdf versions of the Kathmandu Toastmasters
meeting agenda using only the Python standard library.

- .docx is built as an Office Open XML package (a zip of XML parts).
- .pdf is built from raw PDF syntax with a tiny layout engine.

Toastmasters brand palette:
  Maroon #772432, Loyal Blue #004165, Gold #F2DF74, Gray #A9B2B1
"""

import zipfile
import os

# --------------------------------------------------------------------------
# Agenda content (single source of truth)
# --------------------------------------------------------------------------

CLUB_NAME = "Kathmandu Toastmasters Club"
SUBTITLE = "Meeting Agenda"
SLOGAN = "\u201cWhere Unity in Diversity Comes Alive\u201d"
MEETING_INFO = ("We meet every MONDAY  \u2022  Ananda Bhumi Events, "
                "New Baneshwor  \u2022  5:45 PM (NST)")

PILLARS = [
    ("SPEAK", "with Confidence"),
    ("LISTEN", "with Purpose"),
    ("LEARN", "with Curiosity"),
    ("CONNECT", "with Others"),
    ("LEAD", "with Impact"),
]

MISSION = ("We provide a supportive and a positive learning experience in "
           "which members are empowered to develop communication and "
           "leadership skills, resulting in greater self-confidence and "
           "personal growth.")

OFFICERS = [
    ("President", "Namita Thapa"),
    ("VP Education", "Ranjana Koirala"),
    ("VP Membership", "Kanchan Subedi"),
    ("VP Public Relations", "Bigul Sapkota"),
    ("Treasurer", "Astha Shakya"),
    ("Secretary", "Sabitri Lamichhane"),
    ("Sergeant at Arms", "Ayusha Pradhan"),
    ("Immediate Past President", "Ankush Adhikari"),
]

PROGRAM = [
    ("6:00", "Sergeant-At-Arms (SAA) Calls Meeting to Order", "Ayusha Pradhan, MS3"),
    ("6:02", "Presiding Officer", "Namita Thapa, IP3"),
    ("6:07", "PO introduces the Toastmaster of the Evening (TMOE)", "\u2014"),
    ("6:15", "TMOE introduces the General Evaluator (GE) and role takers", "Sudesha Rimal"),
    ("6:25", "TMOE Introduces Featured Speakers", "\u2014"),
    ("6:35", "TMOE introduces the Educational Series Presenter", "\u2014"),
    ("6:55", "TMOE Introduces Table Topic Master", "Jipi Kalu"),
    ("7:10", "TMOE calls the General Evaluator for Evaluation", "\u2014"),
    ("7:23", "Presentation of Awards by Ballot Counter and TMOE", "Phurbu Tashi Lama"),
    ("7:30", "Closing Remarks / Announcements / Meeting Adjourned", "Namita Thapa, IP3"),
    ("7:30", "Networking Break / Postmasters", "\u2014"),
]

ROLES = [
    ("General Evaluator", "Ganesh Kunwar"),
    ("Grammarian (reviews language & introduces WOD)", "Sabitri Lamichhane"),
    ("Ah-Counter (records interjections in speeches)", "Falgun Shumsher Kunwar"),
    ("Timer (records time taken by speakers)", "Santosh Tulachan"),
    ("Ballot Counter (collects votes for better speakers)", "Phurbu Tashi Lama"),
]

WOD = ["Word of the Day:", "Meaning:", "Synonym:", "Example:"]

FEATURED_HEAD = ["Speaker", "Project", "Speech Title", "Level / Project", "Time"]
FEATURED = [["Dip Sagun Gurung", "\u2014", "\u2014", "\u2014", "\u2014"]]

EDU_HEAD = ["Presenter", "Session Title", "Time"]
EDU = [["Bipassana Shrestha", "\u2014", "\u2014"]]

TT_HEAD = ["Speaker", "Topic", "Time"]
TT = [["\u2014", "\u2014", "1\u20132 min"],
      ["\u2014", "\u2014", "1\u20132 min"],
      ["\u2014", "\u2014", "1\u20132 min"]]

EVAL_HEAD = ["Speaker", "Evaluator", "Time"]
EVAL = [["Dip Sagun Gurung", "Laxmi Pathak", "2\u20133 min"]]

EVAL_NOTES = [
    "Evaluation by Grammarian, Ah-Counter, and Timer (2 minutes each)",
    "General Evaluator provides meeting and leaders evaluations (2\u20135 mins)",
    "Return control to the TMOE \u2014 Sudesha Rimal",
]

BALLOT = [
    ("Better Evaluator", ""),
    ("Better Table Topic Speaker", ""),
    ("Better Featured Speaker", ""),
]

# Colors (hex without #)
MAROON = "772432"
BLUE = "004165"
GOLD = "F2DF74"
LIGHT = "F7F4F2"
WHITE = "FFFFFF"

# ==========================================================================
# DOCX GENERATION
# ==========================================================================

def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def run(text, *, bold=False, color=None, size=None, italic=False):
    """A single run of text. size is half-points."""
    props = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if color:
        props.append('<w:color w:val="%s"/>' % color)
    if size:
        props.append('<w:sz w:val="%d"/>' % size)
    rpr = "<w:rPr>%s</w:rPr>" % "".join(props) if props else ""
    return ('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>'
            % (rpr, esc(text)))


def para(runs, *, align=None, shading=None, space_before=0, space_after=120):
    ppr = ["<w:spacing w:before=\"%d\" w:after=\"%d\"/>"
           % (space_before, space_after)]
    if align:
        ppr.append('<w:jc w:val="%s"/>' % align)
    if shading:
        ppr.append('<w:shd w:val="clear" w:fill="%s"/>' % shading)
    ppr_xml = "<w:pPr>%s</w:pPr>" % "".join(ppr)
    return "<w:p>%s%s</w:p>" % (ppr_xml, runs)


def heading(text):
    return para(
        run(text, bold=True, color=MAROON, size=30),
        space_before=240, space_after=80)


def cell(text, *, bold=False, color=None, fill=None, width=None, align=None):
    props = ['<w:tcBorders>'
             '<w:top w:val="single" w:sz="4" w:color="DDDDDD"/>'
             '<w:bottom w:val="single" w:sz="4" w:color="DDDDDD"/>'
             '<w:left w:val="single" w:sz="4" w:color="DDDDDD"/>'
             '<w:right w:val="single" w:sz="4" w:color="DDDDDD"/>'
             '</w:tcBorders>']
    if fill:
        props.append('<w:shd w:val="clear" w:fill="%s"/>' % fill)
    if width:
        props.append('<w:tcW w:w="%d" w:type="dxa"/>' % width)
    props.append('<w:vAlign w:val="center"/>')
    tcpr = "<w:tcPr>%s</w:tcPr>" % "".join(props)
    r = run(text, bold=bold, color=color, size=20)
    p = para(r, align=align, space_after=20, space_before=20)
    return "<w:tc>%s%s</w:tc>" % (tcpr, p)


def table(headers, rows, widths):
    """Build a styled table. widths in dxa (twentieths of a point)."""
    grid = "".join('<w:gridCol w:w="%d"/>' % w for w in widths)
    out = ['<w:tbl>',
           '<w:tblPr>',
           '<w:tblW w:w="0" w:type="auto"/>',
           '<w:tblLayout w:type="fixed"/>',
           '</w:tblPr>',
           '<w:tblGrid>%s</w:tblGrid>' % grid]
    # header row
    hcells = "".join(
        cell(h, bold=True, color=WHITE, fill=BLUE, width=widths[i])
        for i, h in enumerate(headers))
    out.append("<w:tr>%s</w:tr>" % hcells)
    # data rows
    for ri, row in enumerate(rows):
        fill = LIGHT if ri % 2 else None
        rcells = "".join(
            cell(c, fill=fill, width=widths[i],
                 color=(MAROON if (headers[i] == "Time" or
                        (i == 0 and len(headers) == 2)) else None),
                 bold=(headers[i] == "Time"))
            for i, c in enumerate(row))
        out.append("<w:tr>%s</w:tr>" % rcells)
    out.append('</w:tbl>')
    # empty paragraph after table (required between tables)
    out.append("<w:p/>")
    return "".join(out)


def build_docx(path):
    body = []

    # ---- Header block (maroon shaded paragraphs) ----
    body.append(para(run("T  " + CLUB_NAME, bold=True, color=WHITE, size=40),
                     align="center", shading=MAROON,
                     space_before=120, space_after=40))
    body.append(para(run(SUBTITLE.upper(), bold=True, color=GOLD, size=22),
                     align="center", shading=MAROON, space_after=40))
    body.append(para(run(SLOGAN, italic=True, color=WHITE, size=22),
                     align="center", shading=MAROON, space_after=120))
    # meeting info (blue)
    body.append(para(run(MEETING_INFO, bold=True, color=WHITE, size=20),
                     align="center", shading=BLUE,
                     space_before=40, space_after=120))

    # ---- Pillars ----
    body.append(heading("Our Pillars"))
    pillar_rows = [[p[0] for p in PILLARS], [p[1] for p in PILLARS]]
    body.append(table([p[0] for p in PILLARS],
                      [[p[1] for p in PILLARS]],
                      [1800] * 5))

    # ---- Mission ----
    body.append(heading("Club Mission"))
    body.append(para(run(MISSION, italic=True, size=22),
                     shading=LIGHT, space_after=120))

    # ---- Theme Corner ----
    body.append(heading("Theme Corner"))
    body.append(para(run("Theme to be announced.", italic=True, size=20)))

    # ---- Meeting Program ----
    body.append(heading("Meeting Program"))
    body.append(table(["Time", "Agenda Item", "Person Responsible"],
                      PROGRAM, [1100, 5400, 2500]))

    # ---- Officers ----
    body.append(heading("Club Officers"))
    body.append(table(["Role", "Name"], OFFICERS, [4500, 4500]))

    # ---- Roles ----
    body.append(heading("Meeting Roles"))
    body.append(table(["Role", "Name"], ROLES, [5800, 3200]))

    # Word of the day
    body.append(para(run("Word of the Day", bold=True, color=BLUE, size=22),
                     space_before=120, space_after=40))
    for w in WOD:
        body.append(para(
            run(w + " ", bold=True, size=20)
            + run("_" * 45, color="999999", size=20)))

    # ---- Featured Speakers ----
    body.append(heading("Featured Speakers"))
    body.append(para(run("Speakers present speeches following the "
                         "Toastmasters Pathways program.", italic=True,
                         size=18, color="777777")))
    body.append(table(FEATURED_HEAD, FEATURED, [2200, 1700, 2200, 1700, 1200]))

    # ---- Educational Series ----
    body.append(heading("Educational Series"))
    body.append(para(run("Presenter follows the Pathway Enhancement "
                         "guidelines.", italic=True, size=18,
                         color="777777")))
    body.append(table(EDU_HEAD, EDU, [3500, 3500, 2000]))

    # ---- Table Topics ----
    body.append(heading("Table Topics"))
    body.append(para(
        run("Impromptu speaking session for guests. Table Topic Master: ",
            italic=True, size=18, color="777777")
        + run("Jipi Kalu", bold=True, size=18)))
    body.append(para(
        run("Timer Flags:  ", bold=True, size=18)
        + run("1:00 Green  ", color="3A9D4F", bold=True, size=18)
        + run("1:30 Yellow  ", color="E0A800", bold=True, size=18)
        + run("2:00 Red", color="C0392B", bold=True, size=18)))
    body.append(table(TT_HEAD, TT, [3500, 3500, 2000]))

    # ---- Evaluations ----
    body.append(heading("Evaluations"))
    body.append(para(
        run("Timer Flags:  ", bold=True, size=18)
        + run("2:00 Green  ", color="3A9D4F", bold=True, size=18)
        + run("2:30 Yellow  ", color="E0A800", bold=True, size=18)
        + run("3:00 Red", color="C0392B", bold=True, size=18)))
    body.append(table(EVAL_HEAD, EVAL, [3500, 3500, 2000]))
    for note in EVAL_NOTES:
        body.append(para(run("\u2022  " + note, size=20), space_after=40))

    # ---- Ballot ----
    body.append(heading("Ballot Paper"))
    body.append(table(["Category", "Your Vote"],
                      [[b[0], b[1] or "\u2014"] for b in BALLOT],
                      [5000, 4000]))

    # ---- Footer ----
    body.append(para(
        run("Kathmandu Toastmasters Club  \u2022  Where Leaders Are Made",
            bold=True, color=WHITE, size=18),
        align="center", shading=MAROON, space_before=200))

    body_xml = "".join(body)

    sect = ('<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="720" w:right="720" w:bottom="720" '
            'w:left="720" w:header="360" w:footer="360" w:gutter="0"/>'
            '</w:sectPr>')

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body>%s%s</w:body></w:document>' % (body_xml, sect))

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>')

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/></Relationships>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
    print("Wrote", path, os.path.getsize(path), "bytes")


# ==========================================================================
# PDF GENERATION  (tiny layout engine, single helvetica font + fills)
# ==========================================================================

class PDF:
    def __init__(self, width=595, height=842, margin=42):
        self.w = width
        self.h = height
        self.margin = margin
        self.pages = []
        self.ops = []
        self.y = height - margin
        self._new_page_ops()

    def _new_page_ops(self):
        self.ops = []
        self.y = self.h - self.margin

    def _flush_page(self):
        self.pages.append("\n".join(self.ops))

    def new_page(self):
        self._flush_page()
        self._new_page_ops()

    def ensure(self, needed):
        if self.y - needed < self.margin:
            self.new_page()

    @staticmethod
    def _rgb(hexcol):
        r = int(hexcol[0:2], 16) / 255
        g = int(hexcol[2:4], 16) / 255
        b = int(hexcol[4:6], 16) / 255
        return r, g, b

    @staticmethod
    def _esc(t):
        return (t.replace("\\", "\\\\").replace("(", "\\(")
                .replace(")", "\\)"))

    @staticmethod
    def _ascii(t):
        # PDF base font is Latin-1; map common unicode chars
        repl = {"\u2014": "-", "\u2013": "-", "\u2019": "'",
                "\u201c": '"', "\u201d": '"', "\u2022": "-",
                "\u2026": "...", "\u00a0": " "}
        for k, v in repl.items():
            t = t.replace(k, v)
        return t.encode("latin-1", "replace").decode("latin-1")

    def rect(self, x, y, w, h, fill):
        r, g, b = self._rgb(fill)
        self.ops.append("%.3f %.3f %.3f rg" % (r, g, b))
        self.ops.append("%.2f %.2f %.2f %.2f re f" % (x, y, w, h))

    def text(self, x, y, s, size=10, color="000000", bold=False):
        r, g, b = self._rgb(color)
        font = "F2" if bold else "F1"
        s = self._esc(self._ascii(s))
        self.ops.append("BT /%s %d Tf %.3f %.3f %.3f rg %.2f %.2f Td (%s) Tj ET"
                        % (font, size, r, g, b, x, y, s))

    def text_center(self, cx, y, s, size=10, color="000000", bold=False):
        width = self.text_width(s, size, bold)
        self.text(cx - width / 2, y, s, size, color, bold)

    @staticmethod
    def text_width(s, size, bold=False):
        # Approximate Helvetica width (avg factor)
        return len(PDF._ascii(s)) * size * 0.52

    # --- high level helpers ---
    def banner(self):
        # maroon header band
        x = self.margin
        w = self.w - 2 * self.margin
        bh = 78
        self.y -= bh
        self.rect(x, self.y, w, bh, MAROON)
        # gold circle badge
        self.rect(x + 16, self.y + 22, 34, 34, GOLD)
        self.text(x + 27, self.y + 33, "T", size=20, color=MAROON, bold=True)
        self.text_center(self.w / 2, self.y + 52, CLUB_NAME, size=17,
                         color="FFFFFF", bold=True)
        self.text_center(self.w / 2, self.y + 34, SUBTITLE.upper(), size=9,
                         color=GOLD, bold=True)
        self.text_center(self.w / 2, self.y + 16, SLOGAN, size=10,
                         color="FFFFFF")
        # blue info bar
        ih = 20
        self.y -= ih
        self.rect(x, self.y, w, ih, BLUE)
        self.text_center(self.w / 2, self.y + 6, MEETING_INFO, size=8.5,
                         color="FFFFFF", bold=True)
        self.y -= 14

    def heading(self, t):
        self.ensure(40)
        self.y -= 22
        x = self.margin
        self.text(x, self.y, t, size=13, color=MAROON, bold=True)
        # gold underline
        self.rect(x, self.y - 5, self.w - 2 * self.margin, 2, GOLD)
        self.y -= 10

    def paragraph(self, t, size=9.5, color="333333", italic_note=False,
                  bold=False):
        x = self.margin
        max_w = self.w - 2 * self.margin
        for line in self._wrap(t, size, max_w, bold):
            self.ensure(size + 5)
            self.y -= (size + 4)
            self.text(x, self.y, line, size=size, color=color, bold=bold)
        self.y -= 3

    def _wrap(self, text, size, max_w, bold=False):
        words = text.split()
        lines, cur = [], ""
        for word in words:
            trial = (cur + " " + word).strip()
            if self.text_width(trial, size, bold) <= max_w:
                cur = trial
            else:
                if cur:
                    lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        return lines or [""]

    def table(self, headers, rows, col_w, time_cols=None, first_col_accent=False):
        time_cols = time_cols or set()
        x0 = self.margin
        rowh = 18
        # header
        self.ensure(rowh + 4)
        self.y -= rowh
        self.rect(x0, self.y, sum(col_w), rowh, BLUE)
        cx = x0
        for i, htext in enumerate(headers):
            self.text(cx + 5, self.y + 5, htext, size=8.5, color="FFFFFF",
                      bold=True)
            cx += col_w[i]
        # data rows
        for ri, row in enumerate(rows):
            # compute needed height by wrapping the widest cell
            wrapped = []
            for i, ctext in enumerate(row):
                bold = (i in time_cols)
                lines = self._wrap(str(ctext), 8.5, col_w[i] - 10, bold)
                wrapped.append(lines)
            nlines = max(len(w) for w in wrapped)
            cellh = max(rowh, 6 + nlines * 11)
            self.ensure(cellh)
            self.y -= cellh
            if ri % 2 == 0:
                self.rect(x0, self.y, sum(col_w), cellh, LIGHT)
            cx = x0
            for i, lines in enumerate(wrapped):
                accent = (i in time_cols) or (first_col_accent and i == 0)
                col = MAROON if accent else "333333"
                bold = (i in time_cols)
                ly = self.y + cellh - 12
                for ln in lines:
                    self.text(cx + 5, ly, ln, size=8.5, color=col, bold=bold)
                    ly -= 11
                cx += col_w[i]
        self.y -= 6

    def flags(self, items):
        # items: list of (text, color)
        x = self.margin
        self.ensure(18)
        self.y -= 16
        cx = x
        self.text(cx, self.y + 3, "Timer Flags:", size=8.5, color="333333",
                  bold=True)
        cx += 70
        for txt, col in items:
            w = self.text_width(txt, 8) + 14
            self.rect(cx, self.y, w, 13, col)
            self.text(cx + 6, self.y + 3, txt, size=8, color="FFFFFF",
                      bold=True)
            cx += w + 8
        self.y -= 4

    def footer_band(self):
        self.ensure(30)
        self.y -= 26
        x = self.margin
        self.rect(x, self.y, self.w - 2 * self.margin, 22, MAROON)
        self.text_center(self.w / 2, self.y + 7,
                         "Kathmandu Toastmasters Club  -  "
                         "Where Leaders Are Made",
                         size=9, color="FFFFFF", bold=True)

    def save(self, path):
        self._flush_page()
        objects = []

        def add(obj):
            objects.append(obj)
            return len(objects)  # 1-based id

        # Fonts
        f1 = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
        f2 = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>")
        resources = ("<< /Font << /F1 %d 0 R /F2 %d 0 R >> >>" % (f1, f2))

        page_ids = []
        content_ids = []
        # Reserve pages parent id later; build content + page objs
        for ops in self.pages:
            stream = ops
            cid = add("<< /Length %d >>\nstream\n%s\nendstream"
                      % (len(stream), stream))
            content_ids.append(cid)

        pages_id = len(objects) + 1 + len(self.pages)  # placeholder
        # create page objects referencing pages_id
        for cid in content_ids:
            pid = add("<< /Type /Page /Parent %d 0 R /MediaBox [0 0 %d %d] "
                      "/Resources %s /Contents %d 0 R >>"
                      % (pages_id, self.w, self.h, resources, cid))
            page_ids.append(pid)

        kids = " ".join("%d 0 R" % p for p in page_ids)
        pages_obj = ("<< /Type /Pages /Count %d /Kids [%s] >>"
                     % (len(page_ids), kids))
        real_pages_id = add(pages_obj)
        # fix: pages_id we guessed must equal real_pages_id
        assert real_pages_id == pages_id, (real_pages_id, pages_id)

        catalog_id = add("<< /Type /Catalog /Pages %d 0 R >>" % pages_id)

        # serialize
        out = ["%PDF-1.4"]
        offsets = [0]
        pos = len(out[0]) + 1
        body_parts = []
        for i, obj in enumerate(objects, start=1):
            s = "%d 0 obj\n%s\nendobj\n" % (i, obj)
            offsets.append(pos)
            body_parts.append(s)
            pos += len(s.encode("latin-1"))
        xref_pos = pos
        n = len(objects) + 1
        xref = ["xref", "0 %d" % n, "0000000000 65535 f "]
        for off in offsets[1:]:
            xref.append("%010d 00000 n " % off)
        trailer = ("trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF"
                   % (n, catalog_id, xref_pos))
        data = (out[0] + "\n" + "".join(body_parts) + "\n".join(xref) + "\n"
                + trailer)
        with open(path, "wb") as fh:
            fh.write(data.encode("latin-1", "replace"))
        print("Wrote", path, os.path.getsize(path), "bytes")


def build_pdf(path):
    p = PDF()
    p.banner()

    p.heading("Our Pillars")
    p.table([x[0] for x in PILLARS], [[x[1] for x in PILLARS]],
            [102, 102, 102, 102, 103])

    p.heading("Club Mission")
    p.paragraph(MISSION, size=9.5, color="444444")

    p.heading("Theme Corner")
    p.paragraph("Theme to be announced.", size=9, color="777777")

    p.heading("Meeting Program")
    p.table(["Time", "Agenda Item", "Person Responsible"], PROGRAM,
            [55, 290, 166], time_cols={0})

    p.heading("Club Officers")
    p.table(["Role", "Name"], [[o[0], o[1]] for o in OFFICERS],
            [255, 256], first_col_accent=True)

    p.heading("Meeting Roles")
    p.table(["Role", "Name"], [[r[0], r[1]] for r in ROLES], [320, 191])

    p.y -= 4
    p.paragraph("Word of the Day", size=10, color=BLUE, bold=True)
    for w in WOD:
        p.paragraph(w + "  " + "_" * 40, size=9, color="555555")

    p.heading("Featured Speakers")
    p.paragraph("Speakers present speeches following the Toastmasters "
                "Pathways program.", size=8.5, color="777777")
    p.table(FEATURED_HEAD, FEATURED, [120, 95, 110, 100, 86], time_cols={4})

    p.heading("Educational Series")
    p.paragraph("Presenter follows the Pathway Enhancement guidelines.",
                size=8.5, color="777777")
    p.table(EDU_HEAD, EDU, [200, 200, 111], time_cols={2})

    p.heading("Table Topics")
    p.paragraph("Impromptu speaking session for guests. Table Topic Master: "
                "Jipi Kalu", size=8.5, color="777777")
    p.flags([("1:00 Green", "3A9D4F"), ("1:30 Yellow", "E0A800"),
             ("2:00 Red", "C0392B")])
    p.table(TT_HEAD, TT, [200, 200, 111], time_cols={2})

    p.heading("Evaluations")
    p.flags([("2:00 Green", "3A9D4F"), ("2:30 Yellow", "E0A800"),
             ("3:00 Red", "C0392B")])
    p.table(EVAL_HEAD, EVAL, [200, 200, 111], time_cols={2})
    for note in EVAL_NOTES:
        p.paragraph("- " + note, size=9, color="333333")

    p.heading("Ballot Paper")
    p.table(["Category", "Your Vote"],
            [[b[0], b[1] or "\u2014"] for b in BALLOT],
            [300, 211], first_col_accent=True)

    p.footer_band()
    p.save(path)


if __name__ == "__main__":
    base = os.path.dirname(os.path.abspath(__file__))
    build_docx(os.path.join(base, "Kathmandu_Toastmasters_Meeting_Agenda.docx"))
    build_pdf(os.path.join(base, "Kathmandu_Toastmasters_Meeting_Agenda.pdf"))
    print("Done.")
