# -*- coding: utf-8 -*-
"""Build the Sabbath chapter (Part A: Q1-Q10, pages 1-12) as a Word doc
matching the style of 'Introduction - final.docx': Garamond, A5, 13pt justified.
Punctuation: references in plain parentheses, hyphens in verse ranges, no em dashes."""
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

def _segments(p, segs):
    if isinstance(segs, str):
        segs = [(segs, False)]
    for seg in segs:
        txt = seg[0]; bold = seg[1] if len(seg) > 1 else False
        ital = seg[2] if len(seg) > 2 else False
        r = p.add_run(txt); _set_run(r, bold=bold, italic=ital)

def lead(segs):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _segments(p, segs)
    return p

def bullet(segs, level=1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    indent = Mm(6) if level == 1 else Mm(12)
    p.paragraph_format.left_indent = indent
    p.paragraph_format.first_line_indent = Mm(-4)
    p.paragraph_format.space_after = Pt(3)
    mark = "•  " if level == 1 else "–  "   # bullet / short dash marker only
    r = p.add_run(mark); _set_run(r)
    _segments(p, segs)
    return p

# ============================= CONTENT =============================
title("Sabbath")

qheading("1. Are those churches who worship on Sundays lost because they do not observe the Sabbath?")
lead("In the first place, those churches are not lost. “Sabbath” comes from the Hebrew word "
     "“Shuvath”, meaning to rest from labour. It is a day set apart by God for rest from regular labours.")
bullet("Sabbath appears 86 times in the Old Testament (Moses alone gave it 22, while the other prophets mentioned it 64 times).")
bullet("It was first mentioned at creation (Gen 2:2-3). It was again mentioned during the supply of manna to the Israelites (Ex 16:22-24).")
bullet("The weekly Sabbath was made an integral part of a system of Sabbaths when the law covenant was formally inaugurated at Mount Sinai in 1513 BC (Ex 19:1; 20:8-11).")

qheading("2. Was the Sabbath law meant for the whole world?")
lead("The answer to the above question is completely no. When we consider the 430 years the Israelites spent in Egypt, "
     "there was no mention of the keeping of the Sabbath. When God covenanted with Israel, He gave this law to them alone "
     "among all the people in the world. The Sabbath was not observed in Egypt where they stayed, neither was it observed "
     "in Canaan where they were heading during the Exodus.")
bullet("Moses gave this information, that the Sabbath was for Israel and Israel only (Deut 4:7-8).")
bullet("The prophet Amos also testified that it was only Israel that God knew (Amos 3:1-2).")
bullet("Moses again reminded the Jews that it is only Israel who should observe the Sabbath, as a sign between Him and them (Ex 31:12-16).")
bullet("In the book of Samuel, God indicated that He has done great things for Israel, including His Sabbaths (2 Sam 7:23).")
bullet("In the book of Psalms, God affirmed that He has not dealt with any other nation as He has with Israel (Ps 147:19-20).")
bullet("Finally, part of the reasons why God caused the Israelites to observe the Sabbath was that they should use it to "
       "remember their past as slaves in Egypt, and how God freed them with power (Deut 5:12-15).")

qheading("3. Didn’t Sabbath keeping start from the Garden of Eden?")
lead("The argument that Sabbath observation started from the Garden of Eden is not supported by biblical facts. "
     "In fact, the Bible warns us not to think above what is written for us (Rom 12:3).")
bullet("None of the Ten Commandments given to Israel was given to Adam in Eden.")
bullet("The Ten Commandments, including the Sabbath, were given to Israel, about 1,450 years after creation.")
bullet("Three strong reasons stand in favour of the idea that no servant of God observed the Sabbath until Israel did, from the book of Exodus:")
bullet("The first reason is that there is no explicit record in the book of Genesis which mentioned Adam, Noah, Abraham, or Jacob keeping a weekly Sabbath.", level=2)
bullet("Also, Sabbath keeping was a national covenant; there are passages which describe the Sabbath as a specific sign between God and the nation of Israel (Ex 31:16-17).", level=2)
bullet("Finally, from a historical point of view, some theologians argue that the Eden account foreshadowed the law, but the actual practice of Sabbath keeping started during the Exodus, with the collection of manna and the Ten Commandments.", level=2)

qheading("4. How many types of Sabbaths are in the Bible?")
lead("The Bible, in the Old Testament, outlines several types of Sabbaths. Beyond the weekly seventh-day Sabbath, "
     "Scripture commands annual festival rest days, land-rest years, and a special jubilee cycle.")
bullet([("(i) The Weekly Sabbath: ", True),
        ("the foundational weekly day of rest, observed from Friday sunset to Saturday sunset. It commemorates God’s finishing of creation and is listed as one of the Ten Commandments (Ex 20:8-11).", False)])
bullet([("(ii) Annual Sabbaths: ", True),
        ("these are also called festival Sabbaths, specific holy days occurring on fixed dates in the Hebrew calendar, regardless of the day of the week. Regular secular work was prohibited on them.", False)])
bullet("First and Last Days of Unleavened Bread, occurring during Passover week (Lev 23:7-8).", level=2)
bullet("Feast of Weeks, or Pentecost: the harvest celebration (Lev 23:21).", level=2)
bullet("Feast of Trumpets: this refers to the start of the civil new year (Lev 23:24-25).", level=2)
bullet("Day of Atonement: referred to as the “Sabbath of Sabbaths,” a solemn day of strict fasting, complete rest, and national atonement.", level=2)
bullet([("(iii) Long-Term Sabbatical Cycles: ", True),
        ("this occurs every seventh year; the land itself was left unplowed and allowed to rest, without sowing or reaping (Lev 25:1-7).", False)])
bullet("The Year of Jubilee: this occurred after seven cycles of seven years (the 50th year). This super-Sabbath involved land rest, the cancelling of debts, and the return of property to the original owners (Lev 16:31; 25:8-10).", level=2)
bullet("God made a food-security plan to save the nation of Israel from famine during the period in which the land was to rest (Lev 25:20-22).")
bullet("The nation of Israel was to observe all the above Sabbaths without any exceptions (Lev 19:30).")

qheading("5. Was it not Constantine who changed the Sabbath to Sunday?")
lead("The assertion that it was Emperor Constantine the Great who changed the biblical Sabbath from Saturday to Sunday is false. "
     "Although, on March 7, AD 321, Constantine issued a civil law making Sunday an official legal day of rest in the Roman Empire, "
     "Christians had been meeting already (1 Cor 16:1-2).")
bullet("Before Constantine embraced and converted to Christianity in AD 312, Christians had been meeting and worshipping on Sundays; "
       "the book of First Corinthians was written in AD 65.")
bullet("Constantine’s role in affirming Sunday as a public holiday, as part of political strategy, is likened to the effort made by "
       "the former president J. J. Rawlings, who also passed a law to declare both Eid-al-Fitr and Eid-al-Adha (two Muslim festivals) "
       "as public holidays in Ghana in 1995, in solidarity with Ghanaian Muslims who form part of the support base of Rawlings’ party (the N.D.C.).")
bullet("An ensuing question which begs for an answer is this: Can we say that, with the passage of the law to declare these two holidays for Muslims, "
       "the Eid-al-Fitr and Eid-al-Adha festivals were brought to Islam by J. J. Rawlings? If the answer is no, why should Emperor Constantine, "
       "who also passed a law in favour of Sunday rest for the same political expediency for his newly accepted faith, be said to have changed "
       "a Sabbath when Christians had been meeting and worshipping on Sunday already?")
bullet("Many years before Constantine accepted Christianity, the prophet Hosea had prophesied that the observation of festivals, including Sabbaths, would end (Hosea 2:11).")
bullet("Those who cite Daniel 7:25 (that someone would come and change times and laws) lose sight of the same Daniel 2:20-21, where it is said that it is God who has the power to change times and seasons.")
bullet("With these, we can conclude that if there was any change, then that change was from God and not Constantine.")

qheading("6. Did Jesus and His apostles keep the Sabbath?")
bullet("Jesus kept the Sabbath as a Jewish man living under the law of Moses. He regularly attended the synagogue on the seventh day and taught.")
bullet("For Jesus to observe the Sabbath was not news, because, born as a Jew, He was circumcised, named, and purified according to the Jewish traditions and customs (Lk 2:21-22).")
bullet("Even with these, Jesus began to “punch holes” into the Sabbath observation. This caused the Jews (their leaders, the Pharisees, and the Sadducees) to attack and criticise Him on many occasions.")
bullet("For Jesus and His apostles to pass through a cornfield on the Sabbath (Matt 12:1-7), they contravened the Sabbath law in Exodus 34:21.")
bullet("The Jews again accused Jesus of not coming from God because He did not observe the Sabbath (Jn 9:14-17).")
bullet("Jesus was gradually bringing the law to an end, to redeem those under the law (Gal 4:4-5).")
bullet("We should note that Christ is the end of the Old Testament law (Rom 10:4).")
bullet("After Christ’s ascension to heaven, and Christianity had fully come, the apostles neither kept the Sabbath nor taught disciples to observe the Sabbath.")
bullet("After Paul had converted to Christianity, he did not observe the Sabbath any more (Gal 1:13-15).")
bullet("In fact, Paul, on several occasions, was accused by the Jews as one not observing the law, the Sabbath (Acts 18:13-18; 21:28).")
bullet("Paul’s effort to convert the Jews to Christianity made him appear in the Jewish synagogue. People misconstrued this to mean it was his culture to be in the synagogue to keep the Sabbath (1 Cor 9:19-21).")
bullet("If Paul’s intention of going to the synagogue was to keep the Sabbath, how could the Jews beat him after observing it with them? (Acts 13:13-15, 38, 44-45, 48-51).")
bullet("It is not true that the apostles kept the Sabbath.")

qheading("7. What were some rules for Sabbath keeping in Israel?")
bullet("According to biblical Mosaic law, intentionally breaking the Sabbath-day rules (doing routine work) attracted the death penalty, carried out by stoning (Ex 35:1-3). As an enforcement of this Sabbath, an Israelite gathered wood on the Sabbath and was indeed stoned to death under God’s command (Num 15:32-36).")
bullet("Sabbath keeping in the Old Testament also involved specific ritual sacrifices, like:")
bullet([("Burnt Offering: ", True), ("two flawless lambs were presented for sacrifice, to symbolize total devotion to God.", False)], level=2)
bullet([("Grain Offerings: ", True), ("fine flour mixed with olive oil, representing human labour and provision.", False)], level=2)
bullet([("Drink Offerings: ", True), ("wine poured out on the altar as an act of service (Num 28:4-10).", False)], level=2)
bullet("No fire was to be prepared for cooking on the day of the Sabbath (Ex 35:2-3).")
bullet("Showbread was to be eaten on the Sabbath (Lev 24:5-8).")

qheading("8. What was the attitude of the Jews towards Sabbath observation?")
bullet("In historical and theological context, particularly within the New Testament Gospels, the religious leaders of ancient Judaism, such as the Pharisees and scribes, were viewed as having “perverted” the Sabbath.")
bullet("They transformed a day intended as a joyful, spiritual gift of rest into an intolerable, legalistic burden, by adding a vast system of rigid, minute human restrictions.")
bullet("The attitude of the Jewish leaders led to the missing of the spirit of the law. Critics, including Jesus in the Gospels, pointed out that their interpretations inverted the purpose of the commandment. The core principle, that the Sabbath was made to benefit humanity through physical and spiritual renewal, was subordinated to legal technicalities.")
bullet("The Jews carried loads in violation of the Sabbath law, and God had to warn them through the prophet Nehemiah (Neh 13:19-22).")
bullet("The Jews again engaged in their normal trading business on Sabbath days (Neh 13:15-16).")
bullet("Again, Nehemiah and his allies had to close the city gates and appoint gatekeepers to prevent those Jews who would go out to do their own business from defiling the Sabbath (Neh 13:19).")
bullet("God protested, through the prophet Ezekiel, that Israel had perverted the Sabbath (Ezek 20:12-13, 20-21).")
bullet("God swore an oath to cancel the Sabbath (Ps 95:8-11).")
bullet("Stephen, in the New Testament, through the Holy Spirit, also said that the Jews did not obey any of the laws given them (Acts 7:51-53).")
bullet("Consequently, the Jews wanted to find out when the Sabbath would end, so that they could do their trading on Sabbath days (Amos 8:4-5, 7).")
bullet("Indeed, several years after Amos’ prophecy, when Christ was dying on the cross, this prophecy got fulfilled (Matt 27:45-50).")

qheading("9. Why does the Lord’s Church not observe the Sabbath?")
bullet("The Lord’s Church does not keep the traditional Saturday Sabbath because they have the firm belief that Jesus Christ fulfilled the Old Testament law and established a New Covenant (Rom 10:4).")
bullet("Instead of a strict Saturday rest, the Lord’s Church worships on Sunday, the “Lord’s Day,” to honour the resurrection of Jesus, viewing physical Sabbath-keeping as a matter of spiritual freedom rather than a binding commandment (Col 2:16-17).")
bullet("The Jews were given the Sabbath till Christ came (Gal 3:23-25).")
bullet("Sunday better replaces the Saturday Sabbath law (Heb 7:18-19).")

qheading("10. Was the Sabbath meant to be everlasting?")
bullet("To some, the Sabbath is called a perpetual or everlasting covenant in Scripture, because Exodus 31:16-17 says so.")
bullet("They argue that it mirrors God’s eternal rhythm of creation, serves as an everlasting sign of identity between God and His people, and points toward an ultimate spiritual rest.")
bullet("The perpetuity of the Sabbath was limited to the life of the Old Testament laws and the nation of Israel, just like other ordinances in the law.")
bullet("It is not only the Sabbath that was supposed to be everlasting, but also the Passover feast of the Jews (Ex 12:23-24).")
bullet("Again, the burning of incense in the temple, which was supposed to be forever, has ceased (Ex 30:7-8).")
bullet("The Pentecost feast carried an idea of being an everlasting feast, but no Christian Church celebrates it.")
bullet("In Genesis 17:13, God calls physical circumcision an “everlasting covenant” for Abraham and his descendants.")
bullet("In all the above situations, the word translated as “everlasting” can mean enduring, for a very long time, an age, or having perpetual significance, rather than an endless literal requirement.")

# Always written into this script's own folder (the sabath topic folder).
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Sabbath - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write 'Sabbath - draft.docx' - it is open in Word. "
                     "Close it and run again.")
print("saved", out, "| paragraphs:", len(doc.paragraphs))
