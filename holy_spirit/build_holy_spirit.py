# -*- coding: utf-8 -*-
"""Build the Holy Spirit chapter (typed sheets 1-6) as a Word doc matching the
style of 'Introduction - final.docx' and the other chapters: Garamond, A5, 13pt
justified. References in plain parentheses, hyphens in verse ranges, no em dashes.
The sheets were typed with English headings and points and the scripture in Twi;
the Twi is REPLACED with the English NIV (kept in the VERSES table below), keeping
Dad's headings, numbering, and structure. Two parts: five numbered questions on
the Holy Spirit, then "Holy Spirit Baptism"."""
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

title("Holy Spirit")

# ---- sheet 1 ----
qheading("1. Who is the Holy Spirit?")
bullet("The Holy Spirit is divine, a member of the holy Godhead.")
bullet("The Holy Spirit is not a mere power, influence, or feeling.")
bullet("Masculine pronouns are applied to the Spirit (John 14:26).", keep=True)
scripture("John 14:26")
bullet("He is given human attributes:", keep=True)
numbered("(i)", "He speaks (Revelation 2:29).", keep=True)
scripture("Revelation 2:29")
numbered("(ii)", "He leads (Romans 8:14).", keep=True)
scripture("Romans 8:14")
numbered("(iii)", "He forbids (Acts 16:6).", keep=True)
scripture("Acts 16:6")
numbered("(iv)", "He gives understanding (Ephesians 3:4).", keep=True)
scripture("Ephesians 3:4")

qheading("2. Who are those who receive the Holy Spirit?")
bullet("Every Christian receives the Holy Spirit at baptism (Acts 2:37-38; "
       "1 Corinthians 12:13).", keep=True)
scripture("Acts 2:37-38")
scripture("1 Corinthians 12:13")
bullet("God puts His Spirit in all Christians (1 Thessalonians 4:8).", keep=True)
scripture("1 Thessalonians 4:8")
# sheet 2
bullet("The Bible asks us: “Don't you know that the Spirit dwells in us?” "
       "(1 Corinthians 3:16).", keep=True)
scripture("1 Corinthians 3:16")
bullet("Because of the Spirit, we are not of ourselves (1 Corinthians 6:19).", keep=True)
scripture("1 Corinthians 6:19")
bullet("The Spirit can be quenched (1 Thessalonians 5:19).", keep=True)
scripture("1 Thessalonians 5:19")
bullet("The Spirit reveals that Christ is in us (1 John 3:24).", keep=True)
scripture("1 John 3:24")

qheading("3. What role did the Holy Spirit play in the first-century Church?")
bullet("The Holy Spirit confirmed the miraculous spoken word before unbelievers "
       "(Hebrews 2:3-4).", keep=True)
scripture("Hebrews 2:3-4")
bullet("In the absence of the written word, the Spirit guided the teachers.")
bullet("In order to mature the Church, the apostles transmitted the gifts of the "
       "Spirit by laying their hands on believers (Acts 8:18).", keep=True)
scripture("Acts 8:18")
bullet("The confirming work of the Spirit before unbelievers ended when the Bible "
       "was completed (1 Corinthians 13:8-13).", keep=True)
scripture("1 Corinthians 13:8-13")

# ---- sheet 3 ----
qheading("4. What role is the Holy Spirit playing in the life of Christians today?")
bullet("The Spirit sanctifies us (2 Thessalonians 2:13-14).", keep=True)
scripture("2 Thessalonians 2:13-14")
bullet("The Spirit strengthens us (Ephesians 3:16).", keep=True)
scripture("Ephesians 3:16")
bullet("The Spirit comforts us (2 Thessalonians 2:16-17).", keep=True)
scripture("2 Thessalonians 2:16-17")
bullet("The Spirit seals us (Ephesians 1:13-14; 4:30).", keep=True)
scripture("Ephesians 1:13-14")
scripture("Ephesians 4:30")
bullet("He seeks to produce His fruits (Galatians 5:22-23).", keep=True)
scripture("Galatians 5:22-23")

qheading("5. Explain the concept of the Trinity to us.")
bullet("This is illustrated in creation (Genesis 1:26).", keep=True)
scripture("Genesis 1:26")
bullet("It is again illustrated in Christ's baptism (Luke 3:21-22).", keep=True)
scripture("Luke 3:21-22")

# ---- sheet 4 ----
section_heading("Holy Spirit Baptism")
bullet("People know that speaking in tongues is a sign of Holy Spirit baptism.")
bullet("Hence, if one does not speak in tongues, that one has not been baptized "
       "with the Holy Spirit.")
bullet("This is very strange (1 Corinthians 12:28-31).", keep=True)
scripture("1 Corinthians 12:28-31")
bullet("All members of the Corinthian church had received water baptism, but not "
       "all spoke in tongues (1 Corinthians 12:13).")
bullet("If speaking in tongues manifests Holy Spirit baptism:", keep=True)
sub("Why did not all the Corinthians speak in tongues, if they were all baptized "
    "in the Holy Spirit?")
sub("And if one is not baptized in the Holy Spirit, is he or she outside the "
    "Church?")

qheading("What makes Holy Spirit baptism unique?")
bullet("It has occurred only twice in the Bible (Acts 11:15; 15:8).", keep=True)
scripture("Acts 11:15")
scripture("Acts 15:8")
# sheet 5
bullet("Jesus was the administrator of Holy Spirit baptism (John 1:3).", keep=True)
scripture("John 1:3")
bullet("The Holy Spirit baptism was a promise, not a command (Acts 1:3-8).",
       keep=True)
scripture("Acts 1:3-8")
bullet("According to H. Leo Boles, “So the baptism of the Holy Spirit was "
       "definitely a promise, and no one was ever commanded to be baptized in the "
       "Holy Spirit. Baptism in water was a command, but Holy Spirit baptism was "
       "a promise.”")
bullet("The baptism with the Holy Spirit was unexpected (Acts 2:1-4).", keep=True)
scripture("Acts 2:1-4")
bullet("At Cornelius' house, too, it was unexpected (Acts 11:15-16).", keep=True)
scripture("Acts 11:15-16")
bullet("Water baptism is valid till the end, while Holy Spirit baptism has "
       "stopped.")

# ---- sheet 6 ----
qheading("What was the purpose of the Holy Spirit baptism?")
numbered("(i)", "The Holy Spirit baptism was to empower and inspire the apostles "
                "(John 14:26; 16:13).", keep=True)
scripture("John 14:26")
scripture("John 16:13")
bullet("Paul affirmed this (Acts 20:26-27).", keep=True)
scripture("Acts 20:26-27")
numbered("(ii)", "The Holy Spirit baptism was to inspire the New Testament prophets "
                 "(Jude 3).", keep=True)
scripture("Jude 3")
bullet("This led to the recording of the truth (2 Peter 1:3).", keep=True)
scripture("2 Peter 1:3")
numbered("(iii)", "The Holy Spirit baptism signified the beginning of the Church "
                  "(Acts 11:15).", keep=True)
scripture("Acts 11:15")
numbered("(iv)", "It was meant to bear witness through the apostles "
                 "(John 15:27).", keep=True)
scripture("John 15:27")
bullet("The apostles were the eyewitnesses of Christ (1 John 1:1-3).", keep=True)
scripture("1 John 1:1-3")

# ================= SAVE =================
pending = sorted(set(r for r in USED if not VERSES.get(r)))
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Holy Spirit - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write '%s'. It is open in Word. Close it and run again." % out)
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs))
print("NIV text still pending for %d passages:" % len(pending), ", ".join(pending))
