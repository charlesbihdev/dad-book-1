# -*- coding: utf-8 -*-
"""Build the 'Light' chapter as a Word doc matching the style of
'Introduction - final.docx' and the other chapters: Garamond, A5, 13pt justified.
This chapter was TYPED with the scripture quotations in Twi (Akan); per the
project rule the Twi scripture is REPLACED with the English NIV, keeping the
English headings and structure. Quotations are italic. References in plain
parentheses, hyphens in verse ranges, no em dashes."""
import os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

def point(num, text):
    """A bold, numbered main point."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.keep_with_next = True
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

LQ = "“"; RQ = "”"; LS = "‘"; RS = "’"
def q(text):
    return LQ + text + RQ

# ================= CONTENT =================

title("Light")
subtitle("You are the light of the world")

leadq("You are the light of the world (Matthew 5:14-16):")
scripture(q("You are the light of the world. A town built on a hill cannot be "
            "hidden. Neither do people light a lamp and put it under a bowl. "
            "Instead they put it on its stand, and it gives light to everyone in "
            "the house. In the same way, let your light shine before others, that "
            "they may see your good deeds and glorify your Father in heaven."))

bullet("Jesus said that His redeemed people become " + q("salt") + " to preserve "
       "the world.")
bullet("The redeemed people are " + q("light") + " to help the world find their "
       "way out of the darkness of the bondage of sin.")
leadq("Christians are seen to be transformed people (Romans 12:2):")
scripture(q("Do not conform to the pattern of this world, but be transformed by "
            "the renewing of your mind. Then you will be able to test and approve "
            "what God's will is, his good, pleasing and perfect will."))

label("What takes place in the Christian's life")

# ---- 1 ----
point(1, "He becomes the temple of God (1 Corinthians 6:19-20).")
scripture(q("Do you not know that your bodies are temples of the Holy Spirit, who "
            "is in you, whom you have received from God? You are not your own; you "
            "were bought at a price. Therefore honor God with your bodies."))
bullet("The indwelling of the Holy Spirit serves as a seal for the new Christian.")

# ---- 2 ----
point(2, "The redeemed are seen to be transformed people (2 Corinthians 5:17).")
scripture(q("Therefore, if anyone is in Christ, the new creation has come: The "
            "old has gone, the new is here!"))
leadq("The new Christian is dead to the old way of living (Romans 6:1-6):")
scripture(q("What shall we say, then? Shall we go on sinning so that grace may "
            "increase? By no means! We are those who have died to sin; how can we "
            "live in it any longer? Or don't you know that all of us who were "
            "baptized into Christ Jesus were baptized into his death? We were "
            "therefore buried with him through baptism into death in order that, "
            "just as Christ was raised from the dead through the glory of the "
            "Father, we too may live a new life. For if we have been united with "
            "him in a death like his, we will certainly also be united with him in "
            "a resurrection like his. For we know that our old self was crucified "
            "with him so that the body ruled by sin might be done away with, that "
            "we should no longer be slaves to sin."))
leadq("Christians become servants of righteousness (Romans 6:17-18):")
scripture(q("But thanks be to God that, though you used to be slaves to sin, you "
            "have come to obey from your heart the pattern of teaching that has "
            "now claimed your allegiance. You have been set free from sin and have "
            "become slaves to righteousness."))
leadq("As Christians we are expected to abstain from sinful activities, and we "
      "are expected to pursue good deeds (Ephesians 4:21-25):")
scripture(q("When you heard about Christ and were taught in him in accordance with "
            "the truth that is in Jesus. You were taught, with regard to your "
            "former way of life, to put off your old self, which is being "
            "corrupted by its deceitful desires; to be made new in the attitude of "
            "your minds; and to put on the new self, created to be like God in "
            "true righteousness and holiness. Therefore each of you must put off "
            "falsehood and speak truthfully to your neighbor, for we are all "
            "members of one body."))
leadq("Christians are to pursue virtues (Colossians 3:12-15):")
scripture(q("Therefore, as God's chosen people, holy and dearly loved, clothe "
            "yourselves with compassion, kindness, humility, gentleness and "
            "patience. Bear with each other and forgive one another if any of you "
            "has a grievance against someone. Forgive as the Lord forgave you. And "
            "over all these virtues put on love, which binds them all together in "
            "perfect unity. Let the peace of Christ rule in your hearts, since as "
            "members of one body you were called to peace. And be thankful."))

# ---- 3 ----
point(3, "The Christian is a harmonious part of the body of Christ.")
bullet("After baptism, Christ makes the saved part of His body, the Church.")
leadq("Each Christian is urged to maintain the unity of the body of Christ, the "
      "Church (1 Corinthians 12:25):")
scripture(q("So that there should be no division in the body, but that its parts "
            "should have equal concern for each other."))
leadq("Believers are appealed to, to be united with no divisions "
      "(1 Corinthians 1:10):")
scripture(q("I appeal to you, brothers and sisters, in the name of our Lord Jesus "
            "Christ, that all of you agree with one another in what you say and "
            "that there be no divisions among you, but that you be perfectly "
            "united in mind and thought."))

# ---- 4 ----
point(4, "The Christian is supposed to be a fruitful worker (John 15:1-6).")
scripture(q("I am the true vine, and my Father is the gardener. He cuts off every "
            "branch in me that bears no fruit, while every branch that does bear "
            "fruit he prunes so that it will be even more fruitful. You are "
            "already clean because of the word I have spoken to you. Remain in me, "
            "as I also remain in you. No branch can bear fruit by itself; it must "
            "remain in the vine. Neither can you bear fruit unless you remain in "
            "me. I am the vine; you are the branches. If you remain in me and I in "
            "you, you will bear much fruit; apart from me you can do nothing. If "
            "you do not remain in me, you are like a branch that is thrown away and "
            "withers; such branches are picked up, thrown into the fire and "
            "burned."))
leadq("Christians are to serve (Matthew 20:25-28):")
scripture(q("Jesus called them together and said, " + LS + "You know that the "
            "rulers of the Gentiles lord it over them, and their high officials "
            "exercise authority over them. Not so with you. Instead, whoever wants "
            "to become great among you must be your servant, and whoever wants to "
            "be first must be your slave, just as the Son of Man did not come to "
            "be served, but to serve, and to give his life as a ransom for "
            "many." + RS))

# ================= SAVE =================
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Light - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write '%s'. It is open in Word. Close it and run again." % out)
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs))
