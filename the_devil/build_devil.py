# -*- coding: utf-8 -*-
"""Build 'The Devil' chapter as a Word doc matching the style of
'Introduction - final.docx' and the other chapters: Garamond, A5, 13pt justified.
This chapter was TYPED with the scripture quotations in Twi (Akan); per the
project rule the Twi scripture is REPLACED with the English NIV, keeping the
English headings, structure and the solution table. Quotations are italic.
References in plain parentheses, hyphens in verse ranges, no em dashes."""
import os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Garamond"
BODY_SZ = Pt(13)

doc = Document()

normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = BODY_SZ
rpr = normal.element.get_or_add_rPr()
rf = rpr.get_or_add_rFonts()
for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
    rf.set(qn(a), FONT)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.08

sec = doc.sections[0]
sec.page_width = Mm(148)
sec.page_height = Mm(210)
sec.top_margin = Mm(14)
sec.bottom_margin = Mm(15)
sec.left_margin = Mm(15)
sec.right_margin = Mm(13)

def add_page_number(paragraph):
    run = paragraph.add_run()
    fld1 = OxmlElement("w:fldChar"); fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    fld2 = OxmlElement("w:fldChar"); fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)
footer_p = sec.footer.paragraphs[0]
footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_page_number(footer_p)

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
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text.upper()); _set_run(r, bold=True, size=Pt(22))
    return p

def subtitle(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run(text); _set_run(r, bold=True, italic=True, size=Pt(14))
    return p

def centre_heading(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(13))
    return p

def point(num, text):
    """A bold, numbered comparison point."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("%d.  " % num); _set_run(r, bold=True)
    r = p.add_run(text); _set_run(r, bold=True)
    return p

def label(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(13))
    return p

def lead(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); _set_run(r)
    return p

def bullet(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(6)
    p.paragraph_format.first_line_indent = Mm(-4)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("•  "); _set_run(r)
    r = p.add_run(text); _set_run(r)
    return p

def leadq(text):
    """A lead-in line that introduces a quote (keeps with the quote)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(6)
    p.paragraph_format.first_line_indent = Mm(-4)
    p.paragraph_format.space_after = Pt(1)
    p.keep_with_next = True
    r = p.add_run("•  "); _set_run(r)
    r = p.add_run(text); _set_run(r)
    return p

def scripture(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(10)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(7)
    r = p.add_run(text); _set_run(r, italic=True, size=Pt(12))
    return p

def armour_table(rows):
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    for virtue, symbol in rows:
        cells = t.add_row().cells
        pL = cells[0].paragraphs[0]
        rL = pL.add_run(virtue); _set_run(rL, bold=True)
        pR = cells[1].paragraphs[0]
        rR = pR.add_run(symbol); _set_run(rR)
    return t

LQ = "“"; RQ = "”"; LS = "‘"; RS = "’"
def q(text):
    return LQ + text + RQ

# ================= CONTENT =================

title("The Devil")
subtitle("The devil is like a roaring lion")

leadq("The devil is like a roaring lion (1 Peter 5:8):")
scripture(q("Be alert and of sober mind. Your enemy the devil prowls around like "
            "a roaring lion looking for someone to devour."))
lead("In Scripture the lion is also used to represent wicked people (Psalm 10:9).")

# ---- comparisons ----
point(1, "The lion lives on the flesh of other animals.")
bullet("The devil has no interest in any creature other than human.")

point(2, "The lion roams about in search of its prey.")

point(3, "The devil also roams about with the aim of deceiving people to sin, to "
         "bring hardships and problems.")
leadq("Satan roams the earth (Job 1:6-7; see also Matthew 4:1-11):")
scripture(q("One day the angels came to present themselves before the Lord, and "
            "Satan also came with them. The Lord said to Satan, " + LS + "Where "
            "have you come from?" + RS + " Satan answered the Lord, " + LS + "From "
            "roaming throughout the earth, going back and forth on it." + RS))

point(4, "The lion is brave and aggressive. It attacks and overpowers even other "
         "carnivorous animals. The devil also tempts both Christians and "
         "unbelievers without fear.")

point(5, "The lion has no mercy for its prey. It attacks both young and old, big "
         "or small.")
bullet("The devil is able to deceive all Christians, even God's elect.")
leadq("As wickedness increases, love grows cold (Matthew 24:12):")
scripture(q("Because of the increase of wickedness, the love of most will grow cold."))

point(6, "The lion is one of the fastest animals on earth.")
bullet("The devil is craftier and more cunning.")
leadq("Evildoers go on deceiving and being deceived (2 Timothy 3:12-13):")
scripture(q("In fact, everyone who wants to live a godly life in Christ Jesus will "
            "be persecuted, while evildoers and impostors will go from bad to "
            "worse, deceiving and being deceived."))

point(7, "The devil always makes some fearful noise that puts fear into other "
         "animals. Out of panic they try to run for safety, thereby exposing "
         "themselves to attack.")
bullet("The devil also brings troubles like loss of jobs, financial crisis, "
       "sickness, and childlessness. In the attempt to find a solution, man is "
       "led into sin.")

point(8, "The lion has the ability to detect the hiding places of other animals "
         "by their scent. The devil also has the ability to detect man's "
         "weakness, no matter how small it may be.")

# ---- solution ----
centre_heading("SOLUTION")
leadq("Put on the full armour of God (Ephesians 6:13-18):")
scripture(q("Therefore put on the full armor of God, so that when the day of evil "
            "comes, you may be able to stand your ground, and after you have done "
            "everything, to stand. Stand firm then, with the belt of truth buckled "
            "around your waist, with the breastplate of righteousness in place, and "
            "with your feet fitted with the readiness that comes from the gospel of "
            "peace. In addition to all this, take up the shield of faith, with "
            "which you can extinguish all the flaming arrows of the evil one. Take "
            "the helmet of salvation and the sword of the Spirit, which is the word "
            "of God. And pray in the Spirit on all occasions with all kinds of "
            "prayers and requests. With this in mind, be alert and always keep on "
            "praying for all the Lord's people."))

armour_table([
    ("Truth", "belt"),
    ("Righteousness", "breastplate"),
    ("Peace", "sandals"),
    ("Faith", "water to quench"),
    ("Salvation", "helmet"),
])

# ---- questions ----
label("Questions")
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.left_indent = Mm(6); p.paragraph_format.first_line_indent = Mm(-4)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("1.  Mention some things we do against God."); _set_run(r)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.left_indent = Mm(6); p.paragraph_format.first_line_indent = Mm(-4)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("2.  State the four things that constitute obedience."); _set_run(r)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.left_indent = Mm(6); p.paragraph_format.first_line_indent = Mm(-4)
p.paragraph_format.space_after = Pt(2)
r = p.add_run("3.  Why do we have to obey God?"); _set_run(r)

# ================= SAVE =================
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "The Devil - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write '%s'. It is open in Word. Close it and run again." % out)
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs))
