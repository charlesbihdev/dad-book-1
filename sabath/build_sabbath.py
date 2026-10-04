# -*- coding: utf-8 -*-
"""Build the Sabbath chapter (Part A: Q1-Q10, pages 1-12) as a Word doc
matching the style of 'Introduction - final.docx': Garamond, A5, 13pt justified.
Punctuation: references in plain parentheses, hyphens in verse ranges, no em dashes."""
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

def section_heading(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(6)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(15))
    return p

def point_label(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(13))
    return p

def scripture(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(6)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text); _set_run(r, size=Pt(12))
    return p

def cmd_heading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(3)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(13))
    return p

def sub_point(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Mm(4)
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(1)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, italic=True, size=Pt(12))
    return p

def numlist(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Mm(8)
    p.paragraph_format.first_line_indent = Mm(-8)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text); _set_run(r)
    return p

TBL_SZ = Pt(10)

def _fill_scr_cell(cell, ref, text):
    p0 = cell.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p0.paragraph_format.space_after = Pt(1)
    r = p0.add_run(ref); _set_run(r, bold=True, size=TBL_SZ)
    p1 = cell.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r2 = p1.add_run(text); _set_run(r2, size=TBL_SZ)

def prophecy_table():
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit = False
    for c, t in zip(tbl.rows[0].cells, ["Prophecy", "Subject", "Fulfilment"]):
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(t); _set_run(r, bold=True, size=TBL_SZ)
    return tbl

def table_group(tbl, label):
    cells = tbl.add_row().cells
    m = cells[0].merge(cells[1]).merge(cells[2])
    p = m.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(label); _set_run(r, bold=True, size=TBL_SZ)

def table_row(tbl, pref, ptext, subject, fref, ftext):
    cells = tbl.add_row().cells
    _fill_scr_cell(cells[0], pref, ptext)
    ps = cells[1].paragraphs[0]; ps.alignment = WD_ALIGN_PARAGRAPH.LEFT
    rs = ps.add_run(subject); _set_run(rs, size=TBL_SZ)
    _fill_scr_cell(cells[2], fref, ftext)

def set_table_widths(tbl, widths):
    for row in tbl.rows:
        if len(row.cells) != len(widths):
            continue
        for i, c in enumerate(row.cells):
            c.width = widths[i]

def worship_table():
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit = False
    for c, t in zip(tbl.rows[0].cells, ["Old Testament worship", "New Testament worship"]):
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(t); _set_run(r, bold=True, size=TBL_SZ)
    return tbl

def _wcell(cell, label, text):
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(label); _set_run(r, bold=True, size=TBL_SZ)
    if text:
        p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r2 = p2.add_run(text); _set_run(r2, size=TBL_SZ)

def w_split(tbl, old_label, old_text, new_label, new_text):
    cells = tbl.add_row().cells
    _wcell(cells[0], old_label, old_text)
    _wcell(cells[1], new_label, new_text)

def w_shared(tbl, old_label, new_label, text):
    cells = tbl.add_row().cells
    r0 = cells[0].paragraphs[0].add_run(old_label); _set_run(r0, bold=True, size=TBL_SZ)
    r1 = cells[1].paragraphs[0].add_run(new_label); _set_run(r1, bold=True, size=TBL_SZ)
    cells2 = tbl.add_row().cells
    m = cells2[0].merge(cells2[1])
    rm = m.paragraphs[0].add_run(text); _set_run(rm, size=TBL_SZ)

def simple2col(hleft, hright):
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit = False
    for c, t in zip(tbl.rows[0].cells, [hleft, hright]):
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(t); _set_run(r, bold=True, size=TBL_SZ)
    return tbl

def row2(tbl, left, right):
    cells = tbl.add_row().cells
    r0 = cells[0].paragraphs[0].add_run(left); _set_run(r0, size=TBL_SZ)
    r1 = cells[1].paragraphs[0].add_run(right); _set_run(r1, size=TBL_SZ)

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

section_heading("Note these from the Old Testament laws")
point_label("Text: Jeremiah 1:17-19")
scripture("17 “Get yourself ready! Stand up and say to them whatever I command you. Do not be terrified by them, or I will terrify you before them. 18 Today I have made you a fortified city, an iron pillar and a bronze wall to stand against the whole land—against the kings of Judah, its officials, its priests and the people of the land. 19 They will fight against you but will not overcome you, for I am with you and will rescue you,” declares the Lord.")

point_label("Those under the law have fallen from grace (Galatians 5:1-4):")
scripture("1 It is for freedom that Christ has set us free. Stand firm, then, and do not let yourselves be burdened again by a yoke of slavery. 2 Mark my words! I, Paul, tell you that if you let yourselves be circumcised, Christ will be of no value to you at all. 3 Again I declare to every man who lets himself be circumcised that he is obligated to obey the whole law. 4 You who are trying to be justified by the law have been alienated from Christ; you have fallen away from grace.")

point_label("The law kept the Jews till Christ came (Galatians 3:23-25):")
scripture("23 Before the coming of this faith, we were held in custody under the law, locked up until the faith that was to come would be revealed. 24 So the law was our guardian until Christ came that we might be justified by faith. 25 Now that this faith has come, we are no longer under a guardian.")

point_label("Paul was accused of setting aside the law (Acts 18:12-13):")
scripture("12 While Gallio was proconsul of Achaia, the Jews of Corinth made a united attack on Paul and brought him to the place of judgment. 13 “This man,” they charged, “is persuading the people to worship God in ways contrary to the law.”")

point_label("Christians are under a new law of the Spirit (Romans 8:1-4):")
scripture("1 Therefore, there is now no condemnation for those who are in Christ Jesus, 2 because through Christ Jesus the law of the Spirit who gives life has set you free from the law of sin and death. 3 For what the law was powerless to do because it was weakened by the flesh, God did by sending his own Son in the likeness of sinful flesh to be a sin offering. And so he condemned sin in the flesh, 4 in order that the righteous requirement of the law might be fully met in us, who do not live according to the flesh but according to the Spirit.")

point_label("3,000 people killed (Exodus 32:27-29):")
scripture("27 Then he said to them, “This is what the Lord, the God of Israel, says: ‘Each man strap a sword to his side. Go back and forth through the camp from one end to the other, each killing his brother and friend and neighbor.’” 28 The Levites did as Moses commanded, and that day about three thousand of the people died. 29 Then Moses said, “You have been set apart to the Lord today, for you were against your own sons and brothers, and he has blessed you this day.”")

point_label("3,000 were saved (Acts 2:40-41):")
scripture("40 With many other words he warned them; and he pleaded with them, “Save yourselves from this corrupt generation.” 41 Those who accepted his message were baptized, and about three thousand were added to their number that day.")

point_label("This is better explained in 2 Corinthians 3:4-17:")
scripture("4 Such confidence we have through Christ before God. 5 Not that we are competent in ourselves to claim anything for ourselves, but our competence comes from God. 6 He has made us competent as ministers of a new covenant—not of the letter but of the Spirit; for the letter kills, but the Spirit gives life. 7 Now if the ministry that brought death, which was engraved in letters on stone, came with glory, so that the Israelites could not look steadily at the face of Moses because of its glory, transitory though it was, 8 will not the ministry of the Spirit be even more glorious? 9 If the ministry that brought condemnation was glorious, how much more glorious is the ministry that brings righteousness! 10 For what was glorious has no glory now in comparison with the surpassing glory. 11 And if what was transitory came with glory, how much greater is the glory of that which lasts! 12 Therefore, since we have such a hope, we are very bold. 13 We are not like Moses, who would put a veil over his face to prevent the Israelites from seeing the end of what was passing away. 14 But their minds were made dull, for to this day the same veil remains when the old covenant is read. It has not been removed, because only in Christ is it taken away. 15 Even to this day when Moses is read, a veil covers their hearts. 16 But whenever anyone turns to the Lord, the veil is taken away. 17 Now the Lord is the Spirit, and where the Spirit of the Lord is, there is freedom.")

point_label("We are under Christ’s law (1 Corinthians 9:19-21):")
scripture("19 Though I am free and belong to no one, I have made myself a slave to everyone, to win as many as possible. 20 To the Jews I became like a Jew, to win the Jews. To those under the law I became like one under the law (though I myself am not under the law), so as to win those under the law. 21 To those not having the law I became like one not having the law (though I am not free from God’s law but am under Christ’s law), so as to win those not having the law.")

point_label("(Romans 3:27):")
scripture("27 Where, then, is boasting? It is excluded. Because of what law? The law that requires works? No, because of the law that requires faith.")

point_label("The law persisted till Christ (Galatians 4:4):")
scripture("4 But when the set time had fully come, God sent his Son, born of a woman, born under the law,")

point_label("The old law has been changed (Hebrews 7:1-12):")
scripture("1 This Melchizedek was king of Salem and priest of God Most High. He met Abraham returning from the defeat of the kings and blessed him, 2 and Abraham gave him a tenth of everything. First, the name Melchizedek means “king of righteousness”; then also, “king of Salem” means “king of peace.” 3 Without father or mother, without genealogy, without beginning of days or end of life, resembling the Son of God, he remains a priest forever. 4 Just think how great he was: Even the patriarch Abraham gave him a tenth of the plunder! 5 Now the law requires the descendants of Levi who become priests to collect a tenth from the people—that is, from their fellow Israelites—even though they also are descended from Abraham. 6 This man, however, did not trace his descent from Levi, yet he collected a tenth from Abraham and blessed him who had the promises. 7 And without doubt the lesser is blessed by the greater. 8 In the one case, the tenth is collected by people who die; but in the other case, by him who is declared to be living. 9 One might even say that Levi, who collects the tenth, paid the tenth through Abraham, 10 because when Melchizedek met Abraham, Levi was still in the body of his ancestor. 11 If perfection could have been attained through the Levitical priesthood—and indeed the law given to the people established that priesthood—why was there still need for another priest to come, one in the order of Melchizedek, not in the order of Aaron? 12 For when the priesthood is changed, the law must be changed also.")

point_label("(Hebrews 7:18-19):")
scripture("18 The former regulation is set aside because it was weak and useless 19 (for the law made nothing perfect), and a better hope is introduced, by which we draw near to God.")

point_label("Paul had to insult the Galatians (Galatians 3:1-3):")
scripture("1 You foolish Galatians! Who has bewitched you? Before your very eyes Jesus Christ was clearly portrayed as crucified. 2 I would like to learn just one thing from you: Did you receive the Spirit by the works of the law, or by believing what you heard? 3 Are you so foolish? After beginning by means of the Spirit, are you now trying to finish by means of the flesh?")

point_label("We should not be judged (Colossians 2:16-17):")
scripture("16 Therefore do not let anyone judge you by what you eat or drink, or with regard to a religious festival, a New Moon celebration or a Sabbath day. 17 These are a shadow of the things that were to come; the reality, however, is found in Christ.")

point_label("Be mindful of Jewish teachings (Titus 1:10-16):")
scripture("10 For there are many rebellious people, full of meaningless talk and deception, especially those of the circumcision group. 11 They must be silenced, because they are disrupting whole households by teaching things they ought not to teach—and that for the sake of dishonest gain. 12 One of Crete’s own prophets has said it: “Cretans are always liars, evil brutes, lazy gluttons.” 13 This saying is true. Therefore rebuke them sharply, so that they will be sound in the faith 14 and will pay no attention to Jewish myths or to the merely human commands of those who reject the truth. 15 To the pure, all things are pure, but to those who are corrupted and do not believe, nothing is pure. In fact, both their minds and consciences are corrupted. 16 They claim to know God, but by their actions they deny him. They are detestable, disobedient and unfit for doing anything good.")

point_label("The law mainly refers to the Ten Commandments (Deuteronomy 4:13):")
scripture("13 He declared to you his covenant, the Ten Commandments, which he commanded you to follow and then wrote them on two stone tablets.")

point_label("Three things in the ark (Hebrews 9:1-5):")
scripture("1 Now the first covenant had regulations for worship and also an earthly sanctuary. 2 A tabernacle was set up. In its first room were the lampstand and the table with its consecrated bread; this was called the Holy Place. 3 Behind the second curtain was a room called the Most Holy Place, 4 which had the golden altar of incense and the gold-covered ark of the covenant. This ark contained the gold jar of manna, Aaron’s staff that had budded, and the stone tablets of the covenant. 5 Above the ark were the cherubim of the Glory, overshadowing the atonement cover. But we cannot discuss these things in detail now.")

lead("The ark which contained the law is lost.")

point_label("The prophecy about the ark of God (Jeremiah 3:16):")
scripture("16 In those days, when your numbers have increased greatly in the land,” declares the Lord, “people will no longer say, ‘The ark of the covenant of the Lord.’ It will never enter their minds or be remembered; it will not be missed, nor will another one be made.")

point_label("Indeed, the ark is in heaven (Revelation 11:19):")
scripture("19 Then God’s temple in heaven was opened, and within his temple was seen the ark of his covenant. And there came flashes of lightning, rumblings, peals of thunder, an earthquake and a severe hailstorm.")

point_label("Those under the law are under a curse (Galatians 3:10-11):")
scripture("10 For all who rely on the works of the law are under a curse, as it is written: “Cursed is everyone who does not continue to do everything written in the Book of the Law.” 11 Clearly no one who relies on the law is justified before God, because “the righteous will live by faith.”")

point_label("The Old Testament needed to go for the New Testament (Hebrews 10:9):")
scripture("9 Then he said, “Here I am, I have come to do your will.” He sets aside the first to establish the second.")

section_heading("What is the true meaning of Matthew 5:17-18?")
scripture("17 “Do not think that I have come to abolish the Law or the Prophets; I have not come to abolish them but to fulfill them. 18 For truly I tell you, until heaven and earth disappear, not the smallest letter, not the least stroke of a pen, will by any means disappear from the Law until everything is accomplished.”")
bullet("In the above quotation, Jesus clarifies that He did not come to destroy the Hebrew scriptures (“the Law or the Prophets”), but to bring them to their intended completion.")
bullet("Meaning, everything written about Him would be fully accomplished.")
bullet("Jesus expounded the broader meaning of the above quotation in Luke 24:44-45:")
scripture("44 He said to them, “This is what I told you while I was still with you: Everything must be fulfilled that is written about me in the Law of Moses, the Prophets and the Psalms.” 45 Then he opened their minds so they could understand the Scriptures.")
bullet("In fact, the coming of Jesus Christ to the earth fulfilled more than two hundred prophecies from the first five books of Moses (the law), the Psalms, and the prophets.")
lead("The following are some of these prophecies and their fulfilment in Christ:")

_t = prophecy_table()
table_group(_t, "The Law")
table_row(_t, "Deuteronomy 18:18-19",
          "18 “I will raise up for them a prophet like you from among their fellow Israelites, and I will put my words in his mouth. He will tell them everything I command him. 19 I myself will call to account anyone who does not listen to my words that the prophet speaks in my name.”",
          "A prophet like Moses",
          "Acts 3:17-25",
          "17 “Now, fellow Israelites, I know that you acted in ignorance, as did your leaders. 18 But this is how God fulfilled what he had foretold through all the prophets, saying that his Messiah would suffer. 19 Repent, then, and turn to God, so that your sins may be wiped out, that times of refreshing may come from the Lord, 20 and that he may send the Messiah, who has been appointed for you—even Jesus. 21 Heaven must receive him until the time comes for God to restore everything, as he promised long ago through his holy prophets. 22 For Moses said, ‘The Lord your God will raise up for you a prophet like me from among your own people; you must listen to everything he tells you. 23 Anyone who does not listen to him will be completely cut off from their people.’ 24 Indeed, beginning with Samuel, all the prophets who have spoken have foretold these days. 25 And you are heirs of the prophets and of the covenant God made with your fathers. He said to Abraham, ‘Through your offspring all peoples on earth will be blessed.’”")
table_row(_t, "Genesis 3:15",
          "“And I will put enmity between you and the woman, and between your offspring and hers; he will crush your head, and you will strike his heel.”",
          "The seed of a woman",
          "Galatians 4:4",
          "But when the set time had fully come, God sent his Son, born of a woman, born under the law,")
table_row(_t, "Genesis 22:18",
          "“and through your offspring all nations on earth will be blessed, because you have obeyed me.”",
          "A descendant of Abraham",
          "Matthew 1:1",
          "This is the genealogy of Jesus the Messiah the son of David, the son of Abraham:")
table_row(_t, "Genesis 21:2-3",
          "2 Sarah became pregnant and bore a son to Abraham in his old age, at the very time God had promised him. 3 Abraham gave the name Isaac to the son Sarah bore him.",
          "A descendant of Isaac",
          "Matthew 1:2",
          "Abraham was the father of Isaac, Isaac the father of Jacob, Jacob the father of Judah and his brothers,")
table_row(_t, "Genesis 49:10",
          "“The scepter will not depart from Judah, nor the ruler’s staff from between his feet, until he to whom it belongs shall come and the obedience of the nations shall be his.”",
          "From the tribe of Judah",
          "Matthew 2:2",
          "and asked, ‘Where is the one who has been born king of the Jews? We saw his star when it rose and have come to worship him.’")
table_row(_t, "Numbers 24:17",
          "“I see him, but not now; I behold him, but not near. A star will come out of Jacob; a scepter will rise out of Israel. He will crush the foreheads of Moab, the skulls of all the people of Sheth.”",
          "A star out of Jacob",
          "Matthew 2:1-2",
          "1 After Jesus was born in Bethlehem in Judea, during the time of King Herod, Magi from the east came to Jerusalem 2 and asked, ‘Where is the one who has been born king of the Jews? We saw his star when it rose and have come to worship him.’")
table_group(_t, "The Prophets")
table_row(_t, "Isaiah 7:14",
          "“Therefore the Lord himself will give you a sign: The virgin will conceive and give birth to a son, and will call him Immanuel.”",
          "Born of a virgin",
          "Luke 1:30-34",
          "30 But the angel said to her, “Do not be afraid, Mary; you have found favor with God. 31 You will conceive and give birth to a son, and you are to call him Jesus. 32 He will be great and will be called the Son of the Most High. The Lord God will give him the throne of his father David, 33 and he will reign over Jacob’s descendants forever; his kingdom will never end.” 34 “How will this be,” Mary asked the angel, “since I am a virgin?”")
table_row(_t, "Hosea 11:1",
          "“When Israel was a child, I loved him, and out of Egypt I called my son.”",
          "The flight to Egypt",
          "Matthew 2:14-15",
          "14 So he got up, took the child and his mother during the night and left for Egypt, 15 where he stayed until the death of Herod. And so was fulfilled what the Lord had said through the prophet: “Out of Egypt I called my son.”")
table_row(_t, "Jeremiah 31:15",
          "This is what the Lord says: “A voice is heard in Ramah, mourning and great weeping, Rachel weeping for her children and refusing to be comforted, because they are no more.”",
          "Slaughter of the children",
          "Matthew 2:16-17",
          "16 When Herod realized that he had been outwitted by the Magi, he was furious, and he gave orders to kill all the boys in Bethlehem and its vicinity who were two years old and under, in accordance with the time he had learned from the Magi. 17 Then what was said through the prophet Jeremiah was fulfilled.")
table_row(_t, "Zechariah 11:12",
          "“I told them, ‘If you think it best, give me my pay; but if not, keep it.’ So they paid me thirty pieces of silver.”",
          "Betrayed for thirty pieces of silver",
          "Matthew 26:14-15",
          "14 Then one of the Twelve—the one called Judas Iscariot—went to the chief priests 15 and asked, “What are you willing to give me if I deliver him over to you?” So they counted out for him thirty pieces of silver.")
table_row(_t, "Micah 5:2",
          "“But you, Bethlehem Ephrathah, though you are small among the clans of Judah, out of you will come for me one who will be ruler over Israel, whose origins are from of old, from ancient times.”",
          "Born in Bethlehem",
          "Luke 2:4-6",
          "4 So Joseph also went up from the town of Nazareth in Galilee to Judea, to Bethlehem the town of David, because he belonged to the house and line of David. 5 He went there to register with Mary, who was pledged to be married to him and was expecting a child. 6 While they were there, the time came for the baby to be born,")

table_group(_t, "The Psalms")
table_row(_t, "Psalm 78:2-4",
          "“I will open my mouth with a parable; I will utter hidden things, things from of old— things we have heard and known, things our ancestors have told us. We will not hide them from their descendants; we will tell the next generation the praiseworthy deeds of the Lord, his power, and the wonders he has done.”",
          "Speaking in parables",
          "Matthew 13:34-35",
          "34 Jesus spoke all these things to the crowd in parables; he did not say anything to them without using a parable. 35 So was fulfilled what was spoken through the prophet: “I will open my mouth in parables, I will utter things hidden since the creation of the world.”")
table_row(_t, "Psalm 41:9",
          "“Even my close friend, someone I trusted, one who shared my bread, has turned against me.”",
          "Betrayed by a close friend",
          "Luke 22:47-48",
          "47 While he was still speaking a crowd came up, and the man who was called Judas, one of the Twelve, was leading them. He approached Jesus to kiss him, 48 but Jesus asked him, “Judas, are you betraying the Son of Man with a kiss?”")
table_row(_t, "Psalm 22:17-18",
          "“All my bones are on display; people stare and gloat over me. They divide my clothes among them and cast lots for my garment.”",
          "They gambled for His clothes",
          "Matthew 27:35-36",
          "35 When they had crucified him, they divided up his clothes by casting lots. 36 And sitting down, they kept watch over him there.")
table_row(_t, "Psalm 109:4",
          "“In return for my friendship they accuse me, but I am a man of prayer.”",
          "He prayed for His enemies",
          "Luke 23:34",
          "“Jesus said, ‘Father, forgive them, for they do not know what they are doing.’ And they divided up his clothes by casting lots.”")
table_row(_t, "Psalm 16:9-10",
          "“Therefore my heart is glad and my tongue rejoices; my body also will rest secure, because you will not abandon me to the realm of the dead, nor will you let your faithful one see decay.”",
          "His resurrection",
          "Mark 16:6-7",
          "6 “Don’t be alarmed,” he said. “You are looking for Jesus the Nazarene, who was crucified. He has risen! He is not here. See the place where they laid him. 7 But go, tell his disciples and Peter, ‘He is going ahead of you into Galilee. There you will see him, just as he told you.’”")

table_group(_t, "Further prophecies")
table_row(_t, "Zechariah 9:9",
          "“Rejoice greatly, Daughter Zion! Shout, Daughter Jerusalem! See, your king comes to you, righteous and victorious, lowly and riding on a donkey, on a colt, the foal of a donkey.”",
          "Entered Jerusalem on a donkey",
          "Matthew 21:1-11",
          "As they approached Jerusalem and came to Bethphage on the Mount of Olives, Jesus sent two disciples, saying to them, “Go to the village ahead of you, and at once you will find a donkey tied there, with her colt by her. Untie them and bring them to me. If anyone says anything to you, say that the Lord needs them, and he will send them right away.” This took place to fulfill what was spoken through the prophet: “Say to Daughter Zion, ‘See, your king comes to you, gentle and riding on a donkey, and on a colt, the foal of a donkey.’” The disciples went and did as Jesus had instructed them. They brought the donkey and the colt and placed their cloaks on them for Jesus to sit on. A very large crowd spread their cloaks on the road, while others cut branches from the trees and spread them on the road. The crowds that went ahead of him and those that followed shouted, “Hosanna to the Son of David!” “Blessed is he who comes in the name of the Lord!” “Hosanna in the highest heaven!” When Jesus entered Jerusalem, the whole city was stirred and asked, “Who is this?” The crowds answered, “This is Jesus, the prophet from Nazareth in Galilee.”")
table_row(_t, "Isaiah 53:12",
          "“Therefore I will give him a portion among the great, and he will divide the spoils with the strong, because he poured out his life unto death, and was numbered with the transgressors. For he bore the sin of many, and made intercession for the transgressors.”",
          "Crucified with criminals",
          "Mark 15:27-28",
          "27 They crucified two rebels with him, one on his right and one on his left.")
table_row(_t, "Psalm 22:16; 34:20",
          "“Dogs surround me, a pack of villains encircles me; they pierce my hands and my feet” (22:16). “He protects all his bones, not one of them will be broken” (34:20).",
          "Pierced hands and feet; no bone broken",
          "John 19:33-36",
          "33 But when they came to Jesus and found that he was already dead, they did not break his legs. 34 Instead, one of the soldiers pierced Jesus’ side with a spear, bringing a sudden flow of blood and water. 35 The man who saw it has given testimony, and his testimony is true. He knows that he tells the truth, and he testifies so that you also may believe. 36 These things happened so that the scripture would be fulfilled: “Not one of his bones will be broken,”")
table_row(_t, "Isaiah 53:9",
          "“He was assigned a grave with the wicked, and with the rich in his death, though he had done no violence, nor was any deceit in his mouth.”",
          "Buried in a rich man’s tomb",
          "Matthew 27:57-60",
          "57 As evening approached, there came a rich man from Arimathea, named Joseph, who had himself become a disciple of Jesus. 58 Going to Pilate, he asked for Jesus’ body, and Pilate ordered that it be given to him. 59 Joseph took the body, wrapped it in a clean linen cloth, 60 and placed it in his own new tomb that he had cut out of the rock. He rolled a big stone in front of the entrance to the tomb and went away.")
table_row(_t, "Exodus 17:6",
          "“I will stand there before you by the rock at Horeb. Strike the rock, and water will come out of it for the people to drink.”",
          "The spiritual rock",
          "1 Corinthians 10:4",
          "4 and drank the same spiritual drink; for they drank from the spiritual rock that accompanied them, and that rock was Christ.")
table_row(_t, "Deuteronomy 21:23",
          "“you must not leave the body hanging on the pole overnight. Be sure to bury it that same day, because anyone who is hung on a pole is under God’s curse.”",
          "Cursed; hung on a tree",
          "Galatians 3:10-13",
          "10 For all who rely on the works of the law are under a curse, as it is written: “Cursed is everyone who does not continue to do everything written in the Book of the Law.” 11 Clearly no one who relies on the law is justified before God, because “the righteous will live by faith.” 12 The law is not based on faith; on the contrary, it says, “The person who does these things will live by them.” 13 Christ redeemed us from the curse of the law by becoming a curse for us, for it is written: “Cursed is everyone who is hung on a pole.”")
table_row(_t, "Psalm 2:2",
          "“The kings of the earth rise up and the rulers band together against the Lord and against his anointed, saying,”",
          "The Anointed (Messiah)",
          "Acts 2:36",
          "36 “Therefore let all Israel be assured of this: God has made this Jesus, whom you crucified, both Lord and Messiah.”")
table_row(_t, "Psalm 22:2",
          "“My God, I cry out by day, but you do not answer, by night, but I find no rest.”",
          "Darkness upon Calvary",
          "Matthew 27:45",
          "45 From noon until three in the afternoon darkness came over all the land.")
table_row(_t, "Psalm 22:7",
          "“All who see me mock me; they hurl insults, shaking their heads.”",
          "They shook their heads",
          "Matthew 27:39",
          "39 Those who passed by hurled insults at him, shaking their heads")
table_row(_t, "Psalm 22:17-18",
          "“All my bones are on display; people stare and gloat over me. They divide my clothes among them and cast lots for my garment.”",
          "Stripped before men",
          "Luke 23:34-35",
          "34 Jesus said, “Father, forgive them, for they do not know what they are doing.” And they divided up his clothes by casting lots. 35 The people stood watching, and the rulers even sneered at him.")
table_row(_t, "Psalm 22:31",
          "“They will proclaim his righteousness, declaring to a people yet unborn: He has done it!”",
          "“It is finished”",
          "John 19:30",
          "30 When he had received the drink, Jesus said, “It is finished.” With that, he bowed his head and gave up his spirit.")
table_row(_t, "Psalm 31:11",
          "“Because of all my enemies, I am the utter contempt of my neighbors and an object of dread to my closest friends—those who see me on the street flee from me.”",
          "His people fled from Him",
          "Mark 14:50",
          "50 Then everyone deserted him and fled.")
set_table_widths(_t, [Mm(46), Mm(26), Mm(47)])

section_heading("The Old Testament and New Testament analysis of the Ten Commandments (Exodus 20:1-17)")
scripture("1 And God spoke all these words: 2 “I am the Lord your God, who brought you out of Egypt, out of the land of slavery. 3 “You shall have no other gods before me. 4 “You shall not make for yourself an image in the form of anything in heaven above or on the earth beneath or in the waters below. 5 You shall not bow down to them or worship them; for I, the Lord your God, am a jealous God, punishing the children for the sin of the parents to the third and fourth generation of those who hate me, 6 but showing love to a thousand generations of those who love me and keep my commandments. 7 “You shall not misuse the name of the Lord your God, for the Lord will not hold anyone guiltless who misuses his name. 8 “Remember the Sabbath day by keeping it holy. 9 Six days you shall labor and do all your work, 10 but the seventh day is a sabbath to the Lord your God. On it you shall not do any work, neither you, nor your son or daughter, nor your male or female servant, nor your animals, nor any foreigner residing in your towns. 11 For in six days the Lord made the heavens and the earth, the sea, and all that is in them, but he rested on the seventh day. Therefore the Lord blessed the Sabbath day and made it holy. 12 “Honor your father and your mother, so that you may live long in the land the Lord your God is giving you. 13 “You shall not murder. 14 “You shall not commit adultery. 15 “You shall not steal. 16 “You shall not give false testimony against your neighbor. 17 “You shall not covet your neighbor’s house. You shall not covet your neighbor’s wife, or his male or female servant, his ox or donkey, or anything that belongs to your neighbor.”")

cmd_heading("1. You shall have no other gods before me (v3)")
sub_point("This attracted killing when breached (Deuteronomy 13:6-11):")
scripture("6 If your very own brother, or your son or daughter, or the wife you love, or your closest friend secretly entices you, saying, “Let us go and worship other gods” (gods that neither you nor your ancestors have known, 7 gods of the peoples around you, whether near or far, from one end of the land to the other), 8 do not yield to them or listen to them. Show them no pity. Do not spare them or shield them. 9 You must certainly put them to death. Your hand must be the first in putting them to death, and then the hands of all the people. 10 Stone them to death, because they tried to turn you away from the Lord your God, who brought you out of Egypt, out of the land of slavery. 11 Then all Israel will hear and be afraid, and no one among you will do such an evil thing again.")
sub_point("The New Testament law teaches one God (1 Corinthians 8:6):")
scripture("6 yet for us there is but one God, the Father, from whom all things came and for whom we live; and there is but one Lord, Jesus Christ, through whom all things came and through whom we live.")

cmd_heading("2. You shall not make or worship idols (v4-6)")
sub_point("This imposed the death penalty (Deuteronomy 17:2-5):")
scripture("2 If a man or woman living among you in one of the towns the Lord gives you is found doing evil in the eyes of the Lord your God in violation of his covenant, 3 and contrary to my command has worshiped other gods, bowing down to them or to the sun or the moon or the stars in the sky, 4 and this has been brought to your attention, then you must investigate it thoroughly. If it is true and it has been proved that this detestable thing has been done in Israel, 5 take the man or woman who has done this evil deed to your city gate and stone that person to death.")
sub_point("The New Testament teaches against idol worship (1 John 5:21):")
scripture("21 Dear children, keep yourselves from idols.")

cmd_heading("3. You shall not misuse the name of the Lord (v7)")
sub_point("Death penalty (Leviticus 24:16):")
scripture("16 anyone who blasphemes the name of the Lord is to be put to death. The entire assembly must stone them. Whether foreigner or native-born, when they blaspheme the Name they are to be put to death.")
sub_point("New Testament (James 5:12):")
scripture("12 Above all, my brothers and sisters, do not swear—not by heaven or by earth or by anything else. All you need to say is a simple “Yes” or “No.” Otherwise you will be condemned.")
sub_point("(Mark 3:28-30):")
scripture("28 Truly I tell you, people can be forgiven all their sins and every slander they utter, 29 but whoever blasphemes against the Holy Spirit will never be forgiven; they are guilty of an eternal sin.” 30 He said this because they were saying, “He has an impure spirit.”")

cmd_heading("4. Remember the Sabbath day (v8)")
sub_point("Punishment for breaking it (Exodus 31:14-15):")
scripture("14 “‘Observe the Sabbath, because it is holy to you. Anyone who desecrates it is to be put to death; those who do any work on that day must be cut off from their people. 15 For six days work is to be done, but the seventh day is a day of sabbath rest, holy to the Lord. Whoever does any work on the Sabbath day is to be put to death.")
sub_point("(Numbers 15:32-36):")
scripture("32 While the Israelites were in the wilderness, a man was found gathering wood on the Sabbath day. 33 Those who found him gathering wood brought him to Moses and Aaron and the whole assembly, 34 and they kept him in custody, because it was not clear what should be done to him. 35 Then the Lord said to Moses, “The man must die. The whole assembly must stone him outside the camp.” 36 So the assembly took him outside the camp and stoned him to death, as the Lord commanded Moses.")
sub_point("New Testament teaching on the Sabbath (Colossians 2:16-17):")
scripture("16 Therefore do not let anyone judge you by what you eat or drink, or with regard to a religious festival, a New Moon celebration or a Sabbath day. 17 These are a shadow of the things that were to come; the reality, however, is found in Christ.")

cmd_heading("5. Honour your father and mother (v12)")
sub_point("Punishment for breaking it (Deuteronomy 21:18-21):")
scripture("18 If someone has a stubborn and rebellious son who does not obey his father and mother and will not listen to them when they discipline him, 19 his father and mother shall take hold of him and bring him to the elders at the gate of his town. 20 They shall say to the elders, “This son of ours is stubborn and rebellious. He will not obey us. He is a glutton and a drunkard.” 21 Then all the men of his town are to stone him to death. You must purge the evil from among you. All Israel will hear of it and be afraid.")
sub_point("New Testament teaching (Colossians 3:20):")
scripture("20 Children, obey your parents in everything, for this pleases the Lord.")

cmd_heading("6. You shall not murder (v13)")
sub_point("Old Testament punishment (Exodus 21:12):")
scripture("12 “Anyone who strikes a person with a fatal blow is to be put to death.")
sub_point("New Testament (1 John 3:15):")
scripture("15 Anyone who hates a brother or sister is a murderer, and you know that no murderer has eternal life residing in him.")

cmd_heading("7. You shall not commit adultery (v14)")
sub_point("Old Testament punishment (Leviticus 20:10):")
scripture("10 If a man commits adultery with another man’s wife—with the wife of his neighbor—both the adulterer and the adulteress are to be put to death.")
sub_point("New Testament teaching (Hebrews 13:4):")
scripture("4 Marriage should be honored by all, and the marriage bed kept pure, for God will judge the adulterer and all the sexually immoral.")
sub_point("They devoted themselves to the apostles’ teaching (Acts 2:42):")
scripture("42 They devoted themselves to the apostles’ teaching and to fellowship, to the breaking of bread and to prayer.")

cmd_heading("8. You shall not steal (v15)")
sub_point("Old Testament punishment for kidnapping (Exodus 21:16):")
scripture("16 “Anyone who kidnaps someone is to be put to death, whether the victim has been sold or is still in the kidnapper’s possession.")
sub_point("Old Testament restitution (Exodus 22:1):")
scripture("1 “Whoever steals an ox or a sheep and slaughters it or sells it must pay back five head of cattle for the ox and four sheep for the sheep.")
sub_point("Old Testament punishment (Exodus 22:1-4):")
scripture("1 “Whoever steals an ox or a sheep and slaughters it or sells it must pay back five head of cattle for the ox and four sheep for the sheep. 2 “If a thief is caught breaking in at night and is struck a fatal blow, the defender is not guilty of bloodshed; 3 but if it happens after sunrise, the defender is guilty of bloodshed. Anyone who steals must certainly make restitution, but if they have nothing, they must be sold to pay for their theft. 4 If the stolen animal is found alive in their possession—whether ox or donkey or sheep—they must pay back double.")
sub_point("New Testament teaching (Ephesians 4:28):")
scripture("28 Anyone who has been stealing must steal no longer, but must work, doing something useful with their own hands, that they may have something to share with those in need.")

cmd_heading("9. You shall not give false testimony (v16)")
sub_point("Old Testament punishment (Deuteronomy 19:16-21):")
scripture("16 If a malicious witness takes the stand to accuse someone of a crime, 17 the two people involved in the dispute must stand in the presence of the Lord before the priests and the judges who are in office at the time. 18 The judges must make a thorough investigation, and if the witness proves to be a liar, giving false testimony against a fellow Israelite, 19 then do to the false witness as that witness intended to do to the other party. You must purge the evil from among you. 20 The rest of the people will hear of this and be afraid, and never again will such an evil thing be done among you. 21 Show no pity: life for life, eye for eye, tooth for tooth, hand for hand, foot for foot.")
sub_point("New Testament teaching (Ephesians 4:25):")
scripture("25 Therefore each of you must put off falsehood and speak truthfully to your neighbor, for we are all members of one body.")

cmd_heading("10. You shall not covet your neighbour’s property (v17)")
sub_point("Old Testament punishment:")
bullet("Death for Achan.", level=2)
bullet("Leprosy for Gehazi.", level=2)
bullet("Death for Judas.", level=2)
sub_point("New Testament teaching (Ephesians 5:3):")
scripture("3 But among you there must not be even a hint of sexual immorality, or of any kind of impurity, or of greed, because these are improper for God’s holy people.")

lead("A breach of a Ten Commandment law in the Old Testament attracted death; but in the New Testament it is now for repentance.")

section_heading("Zion versus Sinai")
sub_point("The Lord has chosen Zion (Psalm 132:13-15):")
scripture("13 For the Lord has chosen Zion, he has desired it for his dwelling, saying, 14 “This is my resting place for ever and ever; here I will sit enthroned, for I have desired it. 15 I will bless her with abundant provisions; her poor I will satisfy with food.")
sub_point("God has placed His salvation in Zion (Isaiah 46:13):")
scripture("13 I am bringing my righteousness near, it is not far away; and my salvation will not be delayed. I will grant salvation to Zion, my splendor to Israel.")
sub_point("The law of the Lord shall come forth from Zion (Isaiah 2:2-3):")
scripture("2 In the last days the mountain of the Lord’s temple will be established as the highest of the mountains; it will be exalted above the hills, and all nations will stream to it. 3 Many peoples will come and say, “Come, let us go up to the mountain of the Lord, to the temple of the God of Jacob. He will teach us his ways, so that we may walk in his paths.” The law will go out from Zion, the word of the Lord from Jerusalem.")
sub_point("The church is God’s household (1 Timothy 3:14-15):")
scripture("14 Although I hope to come to you soon, I am writing you these instructions so that, 15 if I am delayed, you will know how people ought to conduct themselves in God’s household, which is the church of the living God, the pillar and foundation of the truth.")
sub_point("The Saviour will come from Zion to save both Jews and Gentiles (Romans 11:25-26):")
scripture("25 I do not want you to be ignorant of this mystery, brothers and sisters, so that you may not be conceited: Israel has experienced a hardening in part until the full number of the Gentiles has come in, 26 and in this way all Israel will be saved. As it is written: “The deliverer will come from Zion; he will turn godlessness away from Jacob.”")
sub_point("The Jews are of Sinai, but Christians are of Zion (Galatians 4:21-31):")
scripture("21 Tell me, you who want to be under the law, are you not aware of what the law says? 22 For it is written that Abraham had two sons, one by the slave woman and the other by the free woman. 23 His son by the slave woman was born according to the flesh, but his son by the free woman was born as the result of a divine promise. 24 These things are being taken figuratively: The women represent two covenants. One covenant is from Mount Sinai and bears children who are to be slaves: This is Hagar. 25 Now Hagar stands for Mount Sinai in Arabia and corresponds to the present city of Jerusalem, because she is in slavery with her children. 26 But the Jerusalem that is above is free, and she is our mother. 27 For it is written: “Be glad, barren woman, you who never bore a child; shout for joy and cry aloud, you who were never in labor; because more are the children of the desolate woman than of her who has a husband.” 28 Now you, brothers and sisters, like Isaac, are children of promise. 29 At that time the son born according to the flesh persecuted the son born by the power of the Spirit. It is the same now. 30 But what does Scripture say? “Get rid of the slave woman and her son, for the slave woman’s son will never share in the inheritance with the free woman’s son.” 31 Therefore, brothers and sisters, we are not children of the slave woman, but of the free woman.")
sub_point("Christians are of Zion (Hebrews 12:18-24):")
scripture("18 You have not come to a mountain that can be touched and that is burning with fire; to darkness, gloom and storm; 19 to a trumpet blast or to such a voice speaking words that those who heard it begged that no further word be spoken to them, 20 because they could not bear what was commanded: “If even an animal touches the mountain, it must be stoned to death.” 21 The sight was so terrifying that Moses said, “I am trembling with fear.” 22 But you have come to Mount Zion, to the city of the living God, the heavenly Jerusalem. You have come to thousands upon thousands of angels in joyful assembly, 23 to the church of the firstborn, whose names are written in heaven. You have come to God, the Judge of all, to the spirits of the righteous made perfect, 24 to Jesus the mediator of a new covenant, and to the sprinkled blood that speaks a better word than the blood of Abel.")
sub_point("Christ descends from Zion to judge mankind (Revelation 14:1):")
scripture("1 Then I looked, and there before me was the Lamb, standing on Mount Zion, and with him 144,000 who had his name and his Father’s name written on their foreheads.")

point_label("Mount Sinai")
bullet("Mount Sinai is also called Jabal Musa.")
bullet("It is about 2,285 metres high, in Egypt.")
bullet("It is also called Mount Horeb (7,496.8 ft).")
point_label("Mount Zion")
bullet("Mount Zion is located in Jerusalem.")
bullet("The tomb of David is here.")
bullet("The Last Supper room is also here.")
bullet("Zion symbolizes the City of God, the seat of God’s power, and the place of spiritual significance.")
bullet("It is about 765 metres (2,510 ft).")

section_heading("Names of mountains in the Bible")
numlist("1. Mount Ararat (Noah’s Ark)")
numlist("2. Mount Moriah (Abraham’s temple)")
numlist("3. Mount Carmel (Elijah)")
numlist("4. Mount Nebo (Moses’ death)")
numlist("5. Mount Gilboa (Saul and his sons died)")
numlist("6. Mount Hor (Aaron’s death)")
numlist("7. Mount Tabor (Jesus’ transfiguration)")
numlist("8. Mount Sinai (Moses received the Law)")
numlist("9. Mount Zion (God’s dwelling and kingdom)")
numlist("10. Mount Olivet (Jesus prayed, taught, and ascended into heaven)")

section_heading("The difference between the Old and New Testament worship")
_w = worship_table()
w_split(_w,
  "It was given to Israel alone (Deuteronomy 5:1-3; Psalm 147:19-20; Matthew 10:5-6).",
  "1 Moses summoned all Israel and said: Hear, Israel, the decrees and laws I declare in your hearing today. Learn them and be sure to follow them. 2 The Lord our God made a covenant with us at Horeb. 3 It was not with our ancestors that the Lord made this covenant, but with us, with all of us who are alive here today (Deuteronomy 5:1-3). He has revealed his word to Jacob, his laws and decrees to Israel. He has done this for no other nation; they do not know his laws. Praise the Lord (Psalm 147:19-20). These twelve Jesus sent out with the following instructions: Do not go among the Gentiles or enter any town of the Samaritans. Go rather to the lost sheep of Israel (Matthew 10:5-6).",
  "It was given to all nations (Matthew 28:18-20).",
  "18 Then Jesus came to them and said, “All authority in heaven and on earth has been given to me. 19 Therefore go and make disciples of all nations, baptizing them in the name of the Father and of the Son and of the Holy Spirit, 20 and teaching them to obey everything I have commanded you. And surely I am with you always, to the very end of the age.”")
w_split(_w,
  "Moses was the mediator (Deuteronomy 5:31).",
  "But you stay here with me so that I may give you all the commands, decrees and laws you are to teach them to follow in the land I am giving them to possess.",
  "Jesus is the mediator (Hebrews 9:15; 12:24).",
  "For this reason Christ is the mediator of a new covenant, that those who are called may receive the promised eternal inheritance—now that he has died as a ransom to set them free from the sins committed under the first covenant (Hebrews 9:15). You have come … to Jesus the mediator of a new covenant, and to the sprinkled blood that speaks a better word than the blood of Abel (Hebrews 12:24).")
w_split(_w,
  "Its high priest came from Levi (Hebrews 7:11-12).",
  "11 If perfection could have been attained through the Levitical priesthood—and indeed the law given to the people established that priesthood—why was there still need for another priest to come, one in the order of Melchizedek, not in the order of Aaron? 12 For when the priesthood is changed, the law must be changed also.",
  "Its High Priest came from Judah (Hebrews 7:14-16).",
  "14 For it is clear that our Lord descended from Judah, and in regard to that tribe Moses said nothing about priests. 15 And what we have said is even more clear if another priest like Melchizedek appears, 16 one who has become a priest not on the basis of a regulation as to his ancestry but on the basis of the power of an indestructible life.")
w_shared(_w,
  "Its high priests were weak men (Hebrews 7:28).",
  "Its High Priest is made perfect for ever (Hebrews 7:28).",
  "28 For the law appoints as high priests men in all their weakness; but the oath, which came after the law, appointed the Son, who has been made perfect forever.")
w_shared(_w,
  "It was a good covenant (Hebrews 8:6).",
  "It is a better, superior covenant (Hebrews 8:6).",
  "6 But in fact the ministry Jesus has received is as superior to theirs as the covenant of which he is mediator is superior to the old one, since the new covenant is established on better promises.")
w_shared(_w,
  "It was founded on a good promise, Canaan (Hebrews 8:6).",
  "It is established on a better promise (Hebrews 8:6).",
  "6 But in fact the ministry Jesus has received is as superior to theirs as the covenant of which he is mediator is superior to the old one, since the new covenant is established on better promises.")
w_split(_w,
  "It has a circumcision of the flesh (Colossians 2:11).",
  "11 In him you were also circumcised with a circumcision not performed by human hands. Your whole self ruled by the flesh was put off when you were circumcised by Christ.",
  "It has a circumcision of the spirit (Colossians 2:11-12).",
  "11 In him you were also circumcised with a circumcision not performed by human hands. Your whole self ruled by the flesh was put off when you were circumcised by Christ, 12 having been buried with him in baptism, in which you were also raised with him through your faith in the working of God, who raised him from the dead.")
w_split(_w,
  "It has a physical promise (Canaan).",
  "",
  "It has a spiritual promise, Heaven (2 Corinthians 5:1; 1 Peter 1:3-4).",
  "For we know that if the earthly tent we live in is destroyed, we have a building from God, an eternal house in heaven, not built by human hands (2 Corinthians 5:1). Praise be to the God and Father of our Lord Jesus Christ! In his great mercy he has given us new birth into a living hope through the resurrection of Jesus Christ from the dead, and into an inheritance that can never perish, spoil or fade (1 Peter 1:3-4).")
w_split(_w,
  "It is called the law of sin (Romans 7:25).",
  "Thanks be to God, who delivers me through Jesus Christ our Lord! So then, I myself in my mind am a slave to God’s law, but in my sinful nature a slave to the law of sin.",
  "It is called the law of the Spirit of life (Romans 2:25).",
  "Circumcision has value if you observe the law, but if you break the law, you have become as though you had not been circumcised.")
w_shared(_w,
  "It is the ministry of death (2 Corinthians 3:9).",
  "It is the ministry of righteousness (2 Corinthians 3:9).",
  "9 If the ministry that brought condemnation was glorious, how much more glorious is the ministry that brings righteousness!")
set_table_widths(_w, [Mm(59), Mm(59)])

sub_point("Christians are going to heaven (Ephesians 2:6-7; John 17:24):")
scripture("6 And God raised us up with Christ and seated us with him in the heavenly realms in Christ Jesus, 7 in order that in the coming ages he might show the incomparable riches of his grace, expressed in his kindness to us in Christ Jesus (Ephesians 2:6-7). “Father, I want those you have given me to be with me where I am, and to see my glory, the glory you have given me because you loved me before the creation of the world” (John 17:24).")

section_heading("Christianity is not Judaism because:")
sub_point("(i) The Old Testament has been taken away (Hebrews 10:9):")
scripture("9 Then he said, “Here I am, I have come to do your will.” He sets aside the first to establish the second.")
sub_point("(ii) Christ came to fulfil the Old Testament (Hebrews 7:18-19):")
scripture("18 The former regulation is set aside because it was weak and useless 19 (for the law made nothing perfect), and a better hope is introduced, by which we draw near to God.")
sub_point("Matthew 5:17-18:")
scripture("17 “Do not think that I have come to abolish the Law or the Prophets; I have not come to abolish them but to fulfill them. 18 For truly I tell you, until heaven and earth disappear, not the smallest letter, not the least stroke of a pen, will by any means disappear from the Law until everything is accomplished.”")
sub_point("Luke 24:44-48:")
scripture("44 He said to them, “This is what I told you while I was still with you: Everything must be fulfilled that is written about me in the Law of Moses, the Prophets and the Psalms.” 45 Then he opened their minds so they could understand the Scriptures. 46 He told them, “This is what is written: The Messiah will suffer and rise from the dead on the third day, 47 and repentance for the forgiveness of sins will be preached in his name to all nations, beginning at Jerusalem. 48 You are witnesses of these things.")
sub_point("(iii) Christ said many will not understand our message (Matthew 13:10-17):")
scripture("10 The disciples came to him and asked, “Why do you speak to the people in parables?” 11 He replied, “Because the knowledge of the secrets of the kingdom of heaven has been given to you, but not to them. 12 Whoever has will be given more, and they will have an abundance. Whoever does not have, even what they have will be taken from them. 13 This is why I speak to them in parables: “Though seeing, they do not see; though hearing, they do not hear or understand. 14 In them is fulfilled the prophecy of Isaiah: “‘You will be ever hearing but never understanding; you will be ever seeing but never perceiving. 15 For this people’s heart has become calloused; they hardly hear with their ears, and they have closed their eyes. Otherwise they might see with their eyes, hear with their ears, understand with their hearts and turn, and I would heal them.’ 16 But blessed are your eyes because they see, and your ears because they hear. 17 For truly I tell you, many prophets and righteous people longed to see what you see but did not see it, and to hear what you hear but did not hear it.")
sub_point("(iv) We need to divide the word rightly (2 Timothy 2:15):")
scripture("15 Do your best to present yourself to God as one approved, a worker who does not need to be ashamed and who correctly handles the word of truth.")
sub_point("(v) Christ prophesied a new covenant and a new law (Matthew 26:26-29):")
scripture("26 While they were eating, Jesus took bread, and when he had given thanks, he broke it and gave it to his disciples, saying, “Take and eat; this is my body.” 27 Then he took a cup, and when he had given thanks, he gave it to them, saying, “Drink from it, all of you. 28 This is my blood of the covenant, which is poured out for many for the forgiveness of sins. 29 I tell you, I will not drink from this fruit of the vine from now on until that day when I drink it new with you in my Father’s kingdom.”")
sub_point("(vi) Christians are under grace and truth (John 1:17):")
scripture("17 For the law was given through Moses; grace and truth came through Jesus Christ.")
sub_point("The Jews had wanted to kill Christ (Luke 4:28-30):")
scripture("28 All the people in the synagogue were furious when they heard this. 29 They got up, drove him out of the town, and took him to the brow of the hill on which the town was built, in order to throw him off the cliff. 30 But he walked right through the crowd and went on his way.")
sub_point("Christ’s death has taken away the old law (Colossians 2:14):")
scripture("14 having canceled the charge of our legal indebtedness, which stood against us and condemned us; he has taken it away, nailing it to the cross.")
sub_point("The Jewish priests called Jesus a liar (Matthew 27:62-66):")
scripture("62 The next day, the one after Preparation Day, the chief priests and the Pharisees went to Pilate. 63 “Sir,” they said, “we remember that while he was still alive that deceiver said, ‘After three days I will rise again.’ 64 So give the order for the tomb to be made secure until the third day. Otherwise, his disciples may come and steal the body and tell the people that he has been raised from the dead. This last deception will be worse than the first.” 65 “Take a guard,” Pilate answered. “Go, make the tomb as secure as you know how.” 66 So they went and made the tomb secure by putting a seal on the stone and posting the guard.")
sub_point("The Jewish priests bribed the soldiers (Matthew 28:11-15):")
scripture("11 While the women were on their way, some of the guards went into the city and reported to the chief priests everything that had happened. 12 When the chief priests had met with the elders and devised a plan, they gave the soldiers a large sum of money, 13 telling them, “You are to say, ‘His disciples came during the night and stole him away while we were asleep.’ 14 If this report gets to the governor, we will satisfy him and keep you out of trouble.” 15 So the soldiers took the money and did as they were instructed. And this story has been widely circulated among the Jews to this very day.")
sub_point("Stephen rebuked the Jews (Acts 7:51-53):")
scripture("51 “You stiff-necked people! Your hearts and ears are still uncircumcised. You are just like your ancestors: You always resist the Holy Spirit! 52 Was there ever a prophet your ancestors did not persecute? They even killed those who predicted the coming of the Righteous One. And now you have betrayed and murdered him— 53 you who have received the law that was given through angels but have not obeyed it.”")

section_heading("Conversion from Judaism to Christianity")
sub_point("Jewish priests became obedient to the faith, and the disciples multiplied (Acts 6:7):")
scripture("7 So the word of God spread. The number of disciples in Jerusalem increased rapidly, and a large number of priests became obedient to the faith.")
sub_point("More people were added (Acts 5:14):")
scripture("14 Nevertheless, more and more men and women believed in the Lord and were added to their number.")
sub_point("Christianity entails a new teaching (Acts 17:16-20):")
scripture("16 While Paul was waiting for them in Athens, he was greatly distressed to see that the city was full of idols. 17 So he reasoned in the synagogue with both Jews and God-fearing Greeks, as well as in the marketplace day by day with those who happened to be there. 18 A group of Epicurean and Stoic philosophers began to debate with him. Some of them asked, “What is this babbler trying to say?” Others remarked, “He seems to be advocating foreign gods.” They said this because Paul was preaching the good news about Jesus and the resurrection. 19 Then they took him and brought him to a meeting of the Areopagus, where they said to him, “May we know what this new teaching is that you are presenting? 20 You are bringing some strange ideas to our ears, and we would like to know what they mean.”")
sub_point("Paul used the law of Moses and the prophets to explain Christianity (Acts 28:23-27):")
scripture("23 They arranged to meet Paul on a certain day, and came in even larger numbers to the place where he was staying. He witnessed to them from morning till evening, explaining about the kingdom of God, and from the Law of Moses and from the Prophets he tried to persuade them about Jesus. 24 Some were convinced by what he said, but others would not believe. 25 They disagreed among themselves and began to leave after Paul had made this final statement: “The Holy Spirit spoke the truth to your ancestors when he said through Isaiah the prophet: 26 “‘Go to this people and say, “You will be ever hearing but never understanding; you will be ever seeing but never perceiving.” 27 For this people’s heart has become calloused; they hardly hear with their ears, and they have closed their eyes. Otherwise they might see with their eyes, hear with their ears, understand with their hearts and turn, and I would heal them.’”")

section_heading("Why Christianity is superior to Judaism")
lead("Theologically, Christianity is placed on a higher pedestal than Judaism. Although the two religions originated from God, their differences make Christianity more outstanding.")
_c = simple2col("Christianity", "Judaism")
row2(_c, "Christ is the prophesied Jewish Messiah.", "Judaism maintains that the Messiah has not yet come.")
row2(_c, "The Mosaic law is seen as a temporary preparation, a shadow pointing to Christ.", "The Jews see the Mosaic law as the ultimate law.")
row2(_c, "Salvation is seen as a free gift, received by grace through Christ.", "The Jews focus on keeping the commandments of the covenant.")
row2(_c, "Christianity is an inclusive, universal faith meant for all people, without dietary restrictions, circumcision or Jewish cultural laws.", "The Jewish religion was full of ritual circumcision, dietary restrictions and cultural laws.")
set_table_widths(_c, [Mm(59), Mm(59)])

lead("Apart from the above, Christianity has the following to its advantage over Judaism, as enumerated by the Bible:")
sub_point("(i) Christianity has a better revelation (Hebrews 1:1-4):")
scripture("1 In the past God spoke to our ancestors through the prophets at many times and in various ways, 2 but in these last days he has spoken to us by his Son, whom he appointed heir of all things, and through whom also he made the universe. 3 The Son is the radiance of God’s glory and the exact representation of his being, sustaining all things by his powerful word. After he had provided purification for sins, he sat down at the right hand of the Majesty in heaven. 4 So he became as much superior to the angels as the name he has inherited is superior to theirs.")
sub_point("(ii) It also has a better hope (Hebrews 7:19):")
scripture("19 (for the law made nothing perfect), and a better hope is introduced, by which we draw near to God.")
sub_point("(iii) It has a better covenant and promise (Hebrews 8:6):")
scripture("6 But in fact the ministry Jesus has received is as superior to theirs as the covenant of which he is mediator is superior to the old one, since the new covenant is established on better promises.")
sub_point("(iv) Christianity has a better sacrifice (Hebrews 9:23):")
scripture("23 It was necessary, then, for the copies of the heavenly things to be purified with these sacrifices, but the heavenly things themselves with better sacrifices than these.")
sub_point("(v) It has better possessions (Hebrews 10:34):")
scripture("34 You suffered along with those in prison and joyfully accepted the confiscation of your property, because you knew that you yourselves had better and lasting possessions.")
sub_point("(vi) Christianity has a better country, reserved for its adherents (Hebrews 11:16):")
scripture("16 Instead, they were longing for a better country—a heavenly one. Therefore God is not ashamed to be called their God, for he has prepared a city for them.")
sub_point("(vii) Finally, it has a better resurrection for its members (Hebrews 11:35):")
scripture("35 Women received back their dead, raised to life again. There were others who were tortured, refusing to be released so that they might gain an even better resurrection.")

# Always written into this script's own folder (the sabath topic folder).
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Sabbath - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write 'Sabbath - draft.docx' - it is open in Word. "
                     "Close it and run again.")
print("saved", out, "| paragraphs:", len(doc.paragraphs))
