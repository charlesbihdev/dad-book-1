# -*- coding: utf-8 -*-
"""Build the Tithe chapter (handwritten pages 1-8, Questions 1-11) as a Word doc
matching the style of 'Introduction - final.docx' and the Sabbath chapter:
Garamond, A5, 13pt justified. References in plain parentheses, hyphens in verse
ranges, no em dashes. This chapter is all prose teaching (no Twi appendix, no
block scripture quotations), so references sit inline at the end of sentences."""
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
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run(text.upper())
    _set_run(r, bold=True, size=Pt(22))
    return p

def qheading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.keep_with_next = True
    r = p.add_run(text)
    _set_run(r, bold=True, size=Pt(13))
    return p

def subheading(text):
    """A plain centred/left intro heading (used for FACTS ABOUT TITHING)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4)
    p.keep_with_next = True
    r = p.add_run(text)
    _set_run(r, bold=True, size=Pt(13))
    return p

def lead(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
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

def numlist(text):
    """Roman/numeric sub-point kept as Dad labelled it, e.g. '(i) ...'."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(8)
    p.paragraph_format.first_line_indent = Mm(-8)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); _set_run(r)
    return p

# ================= CONTENT =================

title("Tithe")

# ---- intro: Facts about tithing ----
subheading("Facts about tithing")
bullet("No one is against the use of money in worshipping God. No one is saying "
       "that it is wrong to give money or items to support God's work. Neither is "
       "anyone saying that when you give, God does not bless you (Acts 20:35).")
bullet("But let us speak like the oracle of God (1 Peter 4:11).")
bullet("Let us also note that money not properly acquired is the root cause of all evil.")

# ---- Q1 ----
qheading("1. What is tithing?")
lead("In the Bible, tithing means giving a tenth part (10%) of one's crops or "
     "livestock to support the Levites, the poor, the stranger, or the widows and "
     "orphans. In the Old Testament, it was a legal command for the nation of "
     "Israel to support the priests and the poor.")

# ---- Q2 ----
qheading("2. Give a brief history about tithing.")
lead("It was Abraham who gave a tenth of the spoils of his victory over "
     "Chedorlaomer and his allies (Genesis 14:18-20). The Apostle Paul cites this "
     "incident as proof that Christ's priesthood, according to the manner of "
     "Melchizedek, is superior to that of Levi, since Levi, being in the loins of "
     "Abraham, paid tithe in effect to Melchizedek (Hebrews 7:4-10).")
lead("The second case concerned Jacob, who vowed at Bethel to give one-tenth of "
     "his substance to God (Genesis 28:20-22). These two accounts, however, are "
     "merely instances of voluntarily giving one-tenth. There is no record to the "
     "effect that Abraham, in his 175 years in this life, and Jacob in his 147 "
     "years, commanded their descendants to follow such examples, thereby "
     "establishing a religious practice, custom, or law.")
lead("It is important to note that the Patriarchal period (from Adam to Moses) "
     "lasted for about 2,511 years, and in that time there was no law from God for "
     "the payment of tithe. Also, although Abraham was rich in possessions, he did "
     "not use his livestock regularly for the payment of tithe to God in his worship.")
lead("Payment of tithe became a law for the Israelites after they had come to "
     "Canaan in 1406 BC and the land had been divided among the eleven tribes of "
     "Israel, excluding the tribe of Levi (Leviticus 18:20-26; Leviticus 27:30-32; "
     "Joshua 13:32-33).")

# ---- Q3 ----
qheading("3. What were the things used to pay tithe?")
lead("In the Bible, tithes were primarily paid using agricultural products and "
     "livestock from the land of Israel. This included grain, wine, oil, and the "
     "firstborn animals of the herds and flocks (Leviticus 27:30-32; "
     "Deuteronomy 14:22-23; Matthew 23:23).")
lead("In fact, no money was used to pay tithe, as commanded by God. So those who "
     "collect money as tithe are contravening the will of God.")

# ---- Q4 ----
qheading("4. Maybe, during Bible time, there was no money.")
lead("The above argument is not supported by facts.")
bullet("In the Bible, mankind began using money during the Patriarchal era, "
       "specifically in the time of Abraham (around 2,000 BC) (Genesis 17:12).")
bullet("Abimelech gives Sarah a thousand pieces of silver (Genesis 20:16).")
bullet("Joseph is later sold by his brothers to Ishmaelite traders for twenty "
       "pieces of silver (Genesis 37:28).")
bullet("Later in Israel there was a law on lending money (Exodus 22:25).")
bullet("In the Jewish religion, apart from tithing and other offerings, they were "
       "required to give a collection, as attested to by Jesus (Mark 12:41).")
lead("From the foregoing, the argument that God did not include money in tithing "
     "because there was no money is completely false.")

# ---- Q5 ----
qheading("5. Why members of the Churches of Christ do not pay tithe.")
lead("Churches of Christ believe the 10% tithe was specifically tied to ancient "
     "Israel's agricultural economy and the Levite priesthood, which does not "
     "directly apply to the New Testament Church (2 Corinthians 3:4-6; "
     "Hebrews 7:5, 11-19).")
lead("Again, the Church places emphasis on grace giving and not a fixed "
     "percentage, because the New Testament teaches believers to give cheerfully, "
     "freely, and as they are able, rather than under compulsion "
     "(2 Corinthians 9:6-7).")
lead("Also, there is no New Testament command or approved example that Christians "
     "should tithe.")

# ---- Q6 ----
qheading("6. How many types of tithes did God order the Jews to observe?")
lead("There are four different types of tithing for the Jews. These were:")
numlist("(i) Tithe for the Levites, paid yearly. God tells the Levites that all "
        "the tithes in Israel are their inheritance. The Levites got no land "
        "inheritance among the people (Numbers 18:21-24).")
numlist("(ii) In the second type of tithe, God tells Moses that the Levites must "
        "give a tenth of the tithe they receive from the Israelites as an offering "
        "to the Lord, through Aaron the chief priest (Numbers 18:25-28).")
numlist("(iii) The third form of tithe explains that if the designated central "
        "place of worship (where a Levite resides) is too far to transport your "
        "agricultural tithe, you may convert your crops and livestock into silver, "
        "travel to the location, and use the money to buy food, meat, or wine for "
        "a joyful family feast while remembering the local Levites "
        "(Deuteronomy 14:24-27).")
numlist("(iv) The final type of tithe commands the Israelites to gather a special "
        "third-year tithe of their harvest, storing it in their town to feed "
        "vulnerable members of society, including Levites, foreigners, orphans, "
        "and widows, so that God will bless their work (Deuteronomy 14:28-29).")

# ---- Q7 ----
qheading("7. Why Malachi 3:10 cannot be used by Christians to pay tithe.")
lead("Christians reject using the above quotation to enforce a 10% tithe today "
     "for the following reasons:")
numlist("(i) Malachi's instruction was for the Jews and not Christians, as he "
        "himself was not a Christian.")
numlist("(ii) Malachi accused the Jews of offering defiled food on God's altar, "
        "thereby rendering the Lord's table contemptible. Again, God rebukes the "
        "Jewish priests for offering defective, blind, lame, or sick animals as "
        "sacrifices (Malachi 1:7-8).")
lead("The question is: if Malachi 1:7-8 is obviously not applicable in "
     "Christianity, how can Malachi 3:10 be applicable?")

# ---- Q8 ----
qheading("8. Why Matthew 23:23 cannot be justified to pay tithe.")
lead("In the above quotation, Jesus spoke to the Jews under the old law, before "
     "His death. He targeted justice and mercy over money. Also, the New Testament "
     "teaches a different way of giving, based on a cheerful heart instead of a "
     "strict rule. At that time the old law of Moses was still in force.")
bullet("Jesus told them to obey the law while they lived under it.")
bullet("His death changed the rules for believers (Galatians 4:4).")

# ---- Q9 ----
qheading("9. How does the Church do projects without collecting tithes?")
lead("Churches do projects, pay church workers, and carry out benevolence without "
     "tithe by relying on voluntary freewill offerings, member labour, and "
     "material donations, since tithes were solely for the Levites and the needy "
     "and not for projects (Acts 4:32-35).")
lead("Again, when the tithing law was binding on Israel, projects were still done "
     "through voluntary contribution (Exodus 36:4-7).")

# ---- Q10 ----
qheading("10. Did Jesus or the apostles collect, receive, or instruct "
         "Christians to pay tithe?")
lead("The answer to the above question is a big NO. This is because, under the "
     "Mosaic Law, tithes were strictly designated for the Levitical priests. "
     "Because Jesus descended from the tribe of Judah, rather than Levi, He had no "
     "legal or traditional right to collect the temple tithe (Hebrews 7:14). With "
     "this, even if Jesus had demanded tithing, the Jews would not have paid, as "
     "Jesus' tribe, Judah, received a portion of the land.")
bullet("The apostles did not collect tithe, because tithe belonged only to the "
       "Levites and priests who worked in the temple. Jesus and the apostles were "
       "not from the tribe of Levi.")
bullet("Again, the New Testament teaches giving from the heart with joy, not by an "
       "old-law command (2 Corinthians 3:4-6; Hebrews 7:5, 11-19).")
bullet("Christianity operates differently from Judaism when it comes to tithing. A "
       "Levite who was supposed to have received tithe converted into Christianity, "
       "sold his piece of land, and voluntarily brought the proceeds to the "
       "apostles to support the needy (Acts 4:36-37).")

# ---- Q11 ----
qheading("11. What problems do payment of tithe pose to Christianity?")
bullet("Critics argue that biblical tithing was agricultural (crops and "
       "livestock), meant for ancient Israel's Levite system, and not a mandatory "
       "10% cash levy for modern Christians.")
bullet("Opponents of tithing point out that the New Testament promotes voluntary, "
       "generous, and cheerful giving rather than a fixed legal percentage.")
bullet("The argument is further advanced: why should registers and cards meant for "
       "tithe payment be kept, and references made to them to determine defaulters, "
       "those whose payments are in arrears, and why should it even be used to "
       "decide whether a dead member should be buried or not?")
bullet("A question which still remains unanswered is why, these days, tithes are "
       "collected weekly or monthly, while the original tithing was supposed to be "
       "yearly (Deuteronomy 14:22-23).")
bullet("Payment of tithe poses a big challenge to Christianity, as it has "
       "succeeded in working against Jesus' prayer that all His followers are to be "
       "one, so that unbelievers may know that Christ was from God (John 17:20-21).")

# ================= SAVE =================
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tithe - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write '%s'. It is open in Word. Close it and run again." % out)
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs))
