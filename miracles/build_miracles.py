# -*- coding: utf-8 -*-
"""Build the Miracles chapter ("Don't be deceived by miracles", typed sheets 1-8)
as a Word doc matching the style of 'Introduction - final.docx' and the other
chapters: Garamond, A5, 13pt justified. References in plain parentheses, hyphens
in verse ranges, no em dashes. The sheets were typed with English headings and
points and the scripture in Twi; the Twi is REPLACED with the English NIV (kept in
the VERSES table below), keeping Dad's headings, numbering, and structure. Dad's
handwritten additions on the sheets are worked in where he placed them."""
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
    r = p.add_run("%s.\t" % n); _set_run(r)
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

title("Miracles")
subtitle("Don't be deceived by miracles")
textline("Text: Matthew 24:4-5, 11, 23-26")

lead("Do not believe all spirits (1 John 4:1).")
scripture("Matthew 24:4-5")
scripture("Matthew 24:11")
scripture("Matthew 24:23-26")

# ---- 1 ----
qheading("1. What are miracles?")
lead("Miracles are acts that defy the laws of nature. They are effects in the "
     "physical world that surpass all known human or natural powers, and are "
     "therefore attributed to supernatural agency (Exodus 8:16-19).")
scripture("Exodus 8:16-19")
scripture("Isaiah 7:14")
bullet("God heals (Psalm 41:1-3; James 5:14-15).", keep=True)
scripture("Psalm 41:1-3")
scripture("James 5:14-15")
bullet("We walk by faith, not by sight (2 Corinthians 5:7; John 20:28-29).")

# ---- 2 ----
qheading("2. Do you believe in miracles?")
bullet("Yes, because non-miraculous Christianity is no Christianity "
       "(Matthew 1:23).", keep=True)
scripture("Matthew 1:23")
bullet("Jesus must have been a liar, because He claimed to be doing miracles.")
bullet("Without the miracle of the resurrection, we have no hope of eternal life "
       "(1 Corinthians 15:17).", keep=True)
scripture("1 Corinthians 15:17")
bullet("Christianity without miracles is a vain religion.")
bullet("The Bible couldn't have been inspired (2 Timothy 3:16).", keep=True)
scripture("2 Timothy 3:16")
bullet("We must make it clear that we do believe in miracles.")
bullet("However, most people believe that miracles still occur today "
       "(Deuteronomy 13:1-3).", keep=True)
scripture("Deuteronomy 13:1-3")

# ---- 3 ----
qheading("3. Do you want us to believe that miracles do not occur today?")
bullet("The belief that miracles still occur comes about from ignorance "
       "(2 Thessalonians 2:9-12).", keep=True)
scripture("2 Thessalonians 2:9-12")
bullet("People don't know the purpose of miracles.")
bullet("They don't understand the nature of miracles.")
bullet("They don't understand that miracles were never meant to be extended beyond "
       "the time it took to reveal the word of God (Revelation 13:4-8, 11-16).",
       keep=True)
scripture("Revelation 13:4-8")
scripture("Revelation 13:11-16")

# ---- 4 ----
qheading("4. What were the purposes of miracles?")
bullet("Miracles proved that the early Christians were of God "
       "(Mark 16:17-20).", keep=True)
scripture("Mark 16:17-20")
bullet("They helped to establish that a man was receiving power from God "
       "(Exodus 1:9).", keep=True)
scripture("Exodus 1:9")
bullet("Jesus' miracles were meant to confirm that indeed He was the coming "
       "Messiah (Deuteronomy 18:15-18; John 6:14; Matthew 11:2-5).", keep=True)
scripture("Deuteronomy 18:15-18")
scripture("John 6:14")
scripture("Matthew 11:2-5")
bullet("When Christianity was young, God used miracles to prove that He was with "
       "Christians (Hebrews 2:3-4).", keep=True)
scripture("Hebrews 2:3-4")
sub("Only the well-to-do possessed the scrolls.")
sub("In pagan lands there was no knowledge of the Bible.")
sub("There were no Bible commentaries, concordances, or encyclopaedias.")
lead("So miraculous gifts of special knowledge, wisdom, speaking in tongues, and "
     "discernment of divine utterances were vital (1 Corinthians 12:4, 11, 27-31).")
scripture("1 Corinthians 12:4")
scripture("1 Corinthians 12:11")
scripture("1 Corinthians 12:27-31")
bullet("Miracles were never used for personal benefit.")
bullet("Paul did not heal Timothy (1 Timothy 5:23).", keep=True)
scripture("1 Timothy 5:23")
bullet("Trophimus was not healed by Paul (2 Timothy 4:20).", keep=True)
scripture("2 Timothy 4:20")
bullet("Jesus is the same today (Hebrews 13:8).", keep=True)
scripture("Hebrews 13:8")

# ---- 5 ----
qheading("5. Who could perform miracles in the Bible?")
numbered(1, "Jesus could perform miracles (Luke 4:36-37).", keep=True)
scripture("Luke 4:36-37")
numbered(2, "The apostles could perform miracles (2 Corinthians 12:12).", keep=True)
scripture("2 Corinthians 12:12")
numbered(3, "Those on whom the apostles laid hands could perform miracles "
            "(Acts 8:14-18; Romans 1:10-11).", keep=True)
scripture("Acts 8:14-18")
scripture("Romans 1:10-11")
numbered(4, "That was why Paul said prophecy, speaking in tongues, and knowledge "
            "would vanish (1 Corinthians 13:8, 13).", keep=True)
scripture("1 Corinthians 13:8")
scripture("1 Corinthians 13:13")

# ---- 6 ----
qheading("6. What differentiates Biblical miracles from today's miracles?")
numbered(1, "The Biblical ones are open and public in nature (Matthew 16:9; "
            "Acts 4:16). Today's miracles are done only in their specific places "
            "(camps).", keep=True)
scripture("Matthew 16:9")
numbered(2, "Biblical miracles involved not only animate things but also inanimate "
            "things, such as:", keep=True)
sub("Calming the sea and wind")
sub("Stopping and starting rains")
sub("Changing water into wine and into blood")
numbered(3, "Biblical miracles involved all forms of cures:", keep=True)
sub("Leprosy")
sub("Blindness")
numbered(4, "Biblical miracles are marked by simplicity, as against magical feats "
            "accomplished with special props, staging, lighting, and ritual "
            "(Luke 17:11-14).", keep=True)
scripture("Luke 17:11-14")
numbered(5, "No money or offering was collected (Matthew 10:8; "
            "1 Corinthians 9:13-14; 2 Kings 5:15).", keep=True)
scripture("Matthew 10:8")
scripture("2 Kings 5:15")
numbered(6, "They were not for selfish prominence, but to glorify God "
            "(John 11:1-4; John 9:1-5).", keep=True)
scripture("John 11:1-4")
scripture("John 9:1-5")
numbered(7, "True prophecy promotes peace (Jeremiah 28:9).")

# ---- 7 ----
qheading("7. What effects does the belief in the performance of miracles have on "
         "Christianity?")
bullet("It has led to the springing up of false prophets (Matthew 24:23-26; "
       "2 Corinthians 11:13-15).", keep=True)
scripture("2 Corinthians 11:13-15")
bullet("It has led to the “sale” of God's word (2 Corinthians 2:17).", keep=True)
scripture("2 Corinthians 2:17")
bullet("It has led to the perversion of God's word.")
bullet("It has brought division among Christians.")
bullet("Other serious effects of “miracle” performance are that it has plunged "
       "Christianity into occultism and occultic practices, like:", keep=True)
sub("Divination: obtaining secret knowledge of the future through Satan "
    "(Numbers 22:7).")
sub("Magical powers practised to work wonders (Exodus 7:11).")
sub("The practice of necromancy, where the dead are consulted "
    "(Deuteronomy 18:10-11).")
sub("Sorcery, where fake prophets work with occultic powers (Isaiah 47:9).")
sub("Idolatrous worship and demonic powers through divination and witchcraft.")

# ================= SAVE =================
pending = [r for r in USED if not VERSES.get(r)]
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Miracles - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write '%s'. It is open in Word. Close it and run again." % out)
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs))
print("NIV text still pending for %d passages:" % len(pending), ", ".join(pending))
