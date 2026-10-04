# -*- coding: utf-8 -*-
"""Build the Tongues chapter ("Questions that beg for answers in today's tongue
speaking", typed sheets 1-4) as a Word doc matching the style of 'Introduction -
final.docx' and the other chapters: Garamond, A5, 13pt justified. References in
plain parentheses, hyphens in verse ranges, no em dashes. The sheets were typed
with English questions and answers and the scripture in Twi; the Twi is REPLACED
with the English NIV (kept in the VERSES table below), keeping Dad's questions,
numbering, and structure."""
import os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Garamond"
BODY_SZ = Pt(13)

doc = Document()

# ---- base style ----
normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = BODY_SZ
rpr = normal.element.get_or_add_rPr()
rf = rpr.get_or_add_rFonts()
for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
    rf.set(qn(a), FONT)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.08

# ---- page = A5, margins from the first book ----
sec = doc.sections[0]
sec.page_width = Mm(148)
sec.page_height = Mm(210)
sec.top_margin = Mm(14)
sec.bottom_margin = Mm(15)
sec.left_margin = Mm(15)
sec.right_margin = Mm(13)

# ---- footer page number ----
def add_page_number(paragraph):
    run = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_page_number(footer_p)

# ---- helpers ----
def _set_run(r, bold=False, italic=False, size=None):
    r.font.name = FONT
    rpr = r._r.get_or_add_rPr(); rf = rpr.get_or_add_rFonts()
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT)
    r.font.bold = bold
    r.font.italic = italic
    if size: r.font.size = size

def title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text.upper())
    _set_run(r, bold=True, size=Pt(22))
    return p

def subtitle(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text); _set_run(r, italic=True, size=Pt(14))
    return p

def section_heading(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(15))
    return p

def qa(n, question, answer=None, keep=False):
    """A numbered question in bold, with Dad's short answer after it."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(8)
    p.paragraph_format.first_line_indent = Mm(-8)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(1 if keep else 3)
    p.keep_with_next = keep
    r = p.add_run("%s.\t" % n); _set_run(r, bold=True)
    r = p.add_run(question); _set_run(r, bold=True)
    if answer:
        r = p.add_run(" " + answer); _set_run(r)
    return p

def textline(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(12))
    return p

def qheading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(13))
    return p

def lead(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); _set_run(r)
    return p

def bullet(text, keep=False, mark="•"):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(6)
    p.paragraph_format.first_line_indent = Mm(-4)
    p.paragraph_format.space_after = Pt(1 if keep else 3)
    p.keep_with_next = keep
    r = p.add_run(mark + "  "); _set_run(r)
    r = p.add_run(text); _set_run(r)
    return p

def sub(text):
    """A short sub-list item (no reference), one level in."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Mm(12)
    p.paragraph_format.first_line_indent = Mm(-4)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run("–  "); _set_run(r)   # en dash marker, as in the Sabbath sub-lists
    r = p.add_run(text); _set_run(r)
    return p

def numbered(n, text, keep=False):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(8)
    p.paragraph_format.first_line_indent = Mm(-8)
    p.paragraph_format.space_after = Pt(1 if keep else 3)
    p.keep_with_next = keep
    r = p.add_run(("%s\t" if str(n).startswith("(") else "%s.\t") % n); _set_run(r)
    r = p.add_run(text); _set_run(r)
    return p

def scripture(ref, label=None):
    """The NIV text of `ref`, italic and indented, with the reference after it."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(10)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(7)
    USED.append(ref)
    text = VERSES.get(ref) or "[NIV TEXT PENDING]"
    r = p.add_run("“" + text + "”"); _set_run(r, italic=True, size=Pt(12))
    r = p.add_run(" (%s)" % (label or ref)); _set_run(r, size=Pt(12))
    return p

# ---- NIV text, keyed by reference (replaces Dad's Twi quotations) ----
VERSES = {
}
USED = []

# ================= CONTENT =================

title("Tongues")
subtitle("Questions that beg for answers in today's tongue speaking")

# ---- sheet 1 ----
qa(1, "Will the speaking in tongues be forever?", "No (1 Corinthians 13:8).", keep=True)
scripture("1 Corinthians 13:8")
qa(2, "When will the speaking in tongues stop?", "(1 Corinthians 13:9-10).", keep=True)
scripture("1 Corinthians 13:9-10")
qa(3, "Which of the gifts remain till the coming of Christ?",
   "Faith, hope, and love.")
qa(4, "Tongue, as used in 1 Corinthians 14, refers to what?",
   "It refers to languages (1 Corinthians 14:10).", keep=True)
scripture("1 Corinthians 14:10")
qa(5, "How can tongue speaking be effective?",
   "Unless it is interpreted (1 Corinthians 14:13, 16-19).", keep=True)
scripture("1 Corinthians 14:13")
scripture("1 Corinthians 14:16-19")
qa(6, "Is tongue speaking a sign to believers?", "No (1 Corinthians 14:22).",
   keep=True)
scripture("1 Corinthians 14:22")
qa(7, "If, in the assembly of the church, all speak in tongues, what will "
      "unbelievers say?", "The church is crazy (1 Corinthians 14:23).",
   keep=True)
scripture("1 Corinthians 14:23")
qa(8, "Those who claim to be spiritual must understand us.",
   "(1 Corinthians 14:37-38).", keep=True)
scripture("1 Corinthians 14:37-38")

# ---- sheet 2 ----
qa(9, "What is the purpose of the spiritual gift of tongues?",
   "It was a sign to unbelievers (1 Corinthians 14:22).")
qa(10, "How many people were permitted to speak in tongues in the congregation?",
   "A minimum of two, a maximum of three (1 Corinthians 14:27a).", keep=True)
scripture("1 Corinthians 14:27a")
qa(11, "Could all who spoke in tongues speak at the same time?",
   "No (1 Corinthians 14:27b).", keep=True)
scripture("1 Corinthians 14:27b")
qa(12, "I am eager to speak in tongues, but there is no one to interpret. What "
       "should I do?", "Keep quiet (1 Corinthians 14:28).", keep=True)
scripture("1 Corinthians 14:28")
qa(13, "Why were the operations of the Spirit in the church controlled?",
   "Because God is not a confusing God (1 Corinthians 14:33).", keep=True)
scripture("1 Corinthians 14:33")

qheading("1. Is it like Acts 2:1-11?")
scripture("Acts 2:1-11")
bullet("Where are the tongues of fire?")
bullet("Where is the sound of a rushing wind?")
# sheet 3
bullet("Why can the multitudes today not all understand the tongues, as in Acts 2?")
bullet("It is not the same as Acts 2.")

qheading("2. Is it like Acts 10:44-48?")
scripture("Acts 10:44-48")
bullet("Does the gift of tongues come on non-Christians?")
bullet("Where is the purpose of Acts 10 today, to convince Jews of the equality of "
       "the Gentiles?")
bullet("Does everyone get the gift of tongues, as happened in Cornelius' house?")
bullet("Am I now your enemy because I tell you the truth? (Galatians 4:16).",
       keep=True)
scripture("Galatians 4:16")

qheading("3. Is it like Acts 19:1-7?")
scripture("Acts 19:1-7")
bullet("Where is the apostle to lay his hands on us? (Acts 8:14-18; "
       "Romans 1:10-11).", keep=True)
scripture("Acts 8:14-18")
# sheet 4
scripture("Romans 1:10-11")

qheading("4. Is it like 1 Corinthians 12-14 today?")
bullet("Instead of two or three speaking, why are all speaking? "
       "(1 Corinthians 14:23).", keep=True)
scripture("1 Corinthians 14:23")
bullet("Why are all speaking at the same time?")
bullet("Why is it that people speak though there are no interpreters? "
       "(1 Corinthians 14:28).", keep=True)
scripture("1 Corinthians 14:28")

# ================= SAVE =================
pending = sorted(set(r for r in USED if not VERSES.get(r)))
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tongues - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write '%s'. It is open in Word. Close it and run again." % out)
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs))
print("NIV text still pending for %d passages:" % len(pending), ", ".join(pending))
