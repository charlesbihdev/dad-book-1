# -*- coding: utf-8 -*-
"""Build the Vengeance chapter as a Word doc matching the style of
'Introduction - final.docx' and the other chapters: Garamond, A5, 13pt justified.
Sheets 1-4 were typed with the scripture in Twi (Akan); per the project rule the
Twi scripture is REPLACED with the English NIV, keeping the English headers.
Sheets 5-7 are handwritten English with inline references. Quotations are italic.
References in plain parentheses, hyphens in verse ranges, no em dashes."""
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

def qheading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.keep_with_next = True
    r = p.add_run(text); _set_run(r, bold=True, size=Pt(13))
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

def bullet2(lab, rest):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Mm(6)
    p.paragraph_format.first_line_indent = Mm(-4)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("•  "); _set_run(r)
    r = p.add_run(lab); _set_run(r, bold=True)
    r = p.add_run(rest); _set_run(r)
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

title("Vengeance")
subtitle("Vengeance is the Lord's")

# ---- 1 ----
qheading("1. What is meant by the term vengeance?")
bullet("It is an infliction of punishment in return for an injury or offence.")
bullet("It also means retribution paid by God on behalf of justice.")
bullet("Until God appoints a person to carry out vengeance on His behalf, anyone "
       "who avenges commits sin.")
leadq("Vengeance is the Lord's (Romans 12:17-21):")
scripture(q("Do not repay anyone evil for evil. Be careful to do what is right in "
            "the eyes of everyone. If it is possible, as far as it depends on you, "
            "live at peace with everyone. Do not take revenge, my dear friends, "
            "but leave room for God's wrath, for it is written: " + LS + "It is "
            "mine to avenge; I will repay," + RS + " says the Lord. On the "
            "contrary: " + LS + "If your enemy is hungry, feed him; if he is "
            "thirsty, give him something to drink. In doing this, you will heap "
            "burning coals on his head." + RS + " Do not be overcome by evil, but "
            "overcome evil with good."))

# ---- 2 ----
qheading("2. Apart from God, is there anyone who can exercise vengeance?")
leadq("God has appointed Christ to judge and repay (2 Corinthians 5:10; "
      "Acts 17:31; Romans 2:16; 2 Thessalonians 1:6-9):")
scripture(q("God is just: He will pay back trouble to those who trouble you and "
            "give relief to you who are troubled, and to us as well. This will "
            "happen when the Lord Jesus is revealed from heaven in blazing fire "
            "with his powerful angels. He will punish those who do not know God "
            "and do not obey the gospel of our Lord Jesus. They will be punished "
            "with everlasting destruction and shut out from the presence of the "
            "Lord and from the glory of his might."))
leadq("Christ appointed the apostles (2 Corinthians 10:6; 13:10):")
scripture(q("And we will be ready to punish every act of disobedience, once your "
            "obedience is complete.") + " " + q("This is why I write these things "
            "when I am absent, that when I come I may not have to be harsh in my "
            "use of authority, the authority the Lord gave me for building you up, "
            "not for tearing you down."))
leadq("Christ appointed elders (1 Corinthians 5:9-13):")
scripture(q("I wrote to you in my letter not to associate with sexually immoral "
            "people, not at all meaning the people of this world who are immoral, "
            "or the greedy and swindlers, or idolaters. In that case you would "
            "have to leave this world. But now I am writing to you that you must "
            "not associate with anyone who claims to be a brother or sister but is "
            "sexually immoral or greedy, an idolater or slanderer, a drunkard or "
            "swindler. Do not even eat with such people. What business is it of "
            "mine to judge those outside the church? Are you not to judge those "
            "inside? God will judge those outside. " + LS + "Expel the wicked "
            "person from among you." + RS))
leadq("Parents (Hebrews 12:7-11):")
scripture(q("Endure hardship as discipline; God is treating you as his children. "
            "For what children are not disciplined by their father? If you are not "
            "disciplined, and everyone undergoes discipline, then you are not "
            "legitimate, not true sons and daughters at all. Moreover, we have all "
            "had human fathers who disciplined us and we respected them for it. How "
            "much more should we submit to the Father of spirits and live! They "
            "disciplined us for a little while as they thought best; but God "
            "disciplines us for our good, in order that we may share in his "
            "holiness. No discipline seems pleasant at the time, but painful. "
            "Later on, however, it produces a harvest of righteousness and peace "
            "for those who have been trained by it."))
leadq("Rulers, or those in authority (Romans 13:1-5; 1 Peter 2:13-17):")
scripture(q("Let everyone be subject to the governing authorities, for there is "
            "no authority except that which God has established. The authorities "
            "that exist have been established by God. Consequently, whoever rebels "
            "against the authority is rebelling against what God has instituted, "
            "and those who do so will bring judgment on themselves. For rulers "
            "hold no terror for those who do right, but for those who do wrong. Do "
            "you want to be free from fear of the one in authority? Then do what "
            "is right and you will be commended. For the one in authority is God's "
            "servant for your good. But if you do wrong, be afraid, for rulers do "
            "not bear the sword for no reason. They are God's servants, agents of "
            "wrath to bring punishment on the wrongdoer. Therefore, it is "
            "necessary to submit to the authorities, not only because of possible "
            "punishment but also as a matter of conscience."))
scripture(q("Submit yourselves for the Lord's sake to every human authority: "
            "whether to the emperor, as the supreme authority, or to governors, "
            "who are sent by him to punish those who do wrong and to commend those "
            "who do right. For it is God's will that by doing good you should "
            "silence the ignorant talk of foolish people. Live as free people, but "
            "do not use your freedom as a cover-up for evil; live as God's slaves. "
            "Show proper respect to everyone, love the family of believers, fear "
            "God, honor the emperor."))

# ---- 3 ----
qheading("3. Why is man always prone to avenge?")
leadq("This is because man is imperfect (Proverbs 6:32-35):")
scripture(q("But a man who commits adultery has no sense; whoever does so "
            "destroys himself. Blows and disgrace are his lot, and his shame will "
            "never be wiped away. For jealousy arouses a husband's fury, and he "
            "will show no mercy when he takes revenge. He will not accept any "
            "compensation; he will refuse a bribe, however great it is."))
leadq("Most people are controlled by anger (James 1:19-20):")
scripture(q("My dear brothers and sisters, take note of this: Everyone should be "
            "quick to listen, slow to speak and slow to become angry, because "
            "human anger does not produce the righteousness that God desires."))
leadq("Anger resides in the bosom of the foolish (Ecclesiastes 7:9):")
scripture(q("Do not be quickly provoked in your spirit, for anger resides in the "
            "lap of fools."))

# ---- 4 ----
qheading("4. Can you give us examples from the Bible where people did not take "
         "vengeance?")
leadq("Joseph (Genesis 50:15-21):")
scripture(q("When Joseph's brothers saw that their father was dead, they said, " +
            LS + "What if Joseph holds a grudge against us and pays us back for "
            "all the wrongs we did to him?" + RS + " So they sent word to Joseph, "
            "saying, " + LS + "Your father left these instructions before he died: "
            "This is what you are to say to Joseph: I ask you to forgive your "
            "brothers the sins and the wrongs they committed in treating you so "
            "badly. Now please forgive the sins of the servants of the God of your "
            "father." + RS + " When their message came to him, Joseph wept. His "
            "brothers then came and threw themselves down before him. " + LS + "We "
            "are your slaves," + RS + " they said. But Joseph said to them, " + LS +
            "Don't be afraid. Am I in the place of God? You intended to harm me, "
            "but God intended it for good to accomplish what is now being done, "
            "the saving of many lives. So then, don't be afraid. I will provide "
            "for you and your children." + RS + " And he reassured them and spoke "
            "kindly to them."))
leadq("David (1 Samuel 24:10-15):")
scripture(q("This day you have seen with your own eyes how the Lord delivered you "
            "into my hands in the cave. Some urged me to kill you, but I spared "
            "you; I said, " + LS + "I will not lay my hand on my lord, because he "
            "is the Lord's anointed." + RS + " See, my father, look at this piece "
            "of your robe in my hand! I cut off the corner of your robe but did "
            "not kill you. See that there is nothing in my hand to indicate that I "
            "am guilty of wrongdoing or rebellion. I have not wronged you, but you "
            "are hunting me down to take my life. May the Lord judge between you "
            "and me. And may the Lord avenge the wrongs you have done to me, but "
            "my hand will not touch you. As the old saying goes, " + LS + "From "
            "evildoers come evil deeds," + RS + " so my hand will not touch you. "
            "Against whom has the king of Israel come out? Who are you pursuing? A "
            "dead dog? A flea? May the Lord be our judge and decide between us. "
            "May he consider my cause and uphold it; may he vindicate me by "
            "delivering me from your hand."))
leadq("Jesus healed the servant's cut ear (Matthew 26:51-53; Luke 22:51):")
scripture(q("With that, one of Jesus' companions reached for his sword, drew it "
            "out and struck the servant of the high priest, cutting off his ear. " +
            LS + "Put your sword back in its place," + RS + " Jesus said to him, " +
            LS + "for all who draw the sword will die by the sword. Do you think I "
            "cannot call on my Father, and he will at once put at my disposal more "
            "than twelve legions of angels?" + RS) + " " + q("But Jesus answered, " +
            LS + "No more of this!" + RS + " And he touched the man's ear and "
            "healed him."))
leadq("Christ on the cross prayed for His enemies (Luke 23:34):")
scripture(q("Jesus said, " + LS + "Father, forgive them, for they do not know "
            "what they are doing." + RS + " And they divided up his clothes by "
            "casting lots."))
leadq("Christians in general (1 Corinthians 4:12-13):")
scripture(q("We work hard with our own hands. When we are cursed, we bless; when "
            "we are persecuted, we endure it; when we are slandered, we answer "
            "kindly. We have become the scum of the earth, the garbage of the "
            "world, right up to this moment."))
leadq("Bless those who persecute you (Romans 12:14):")
scripture(q("Bless those who persecute you; bless and do not curse."))

# ---- 5 ----
qheading("5. Why do people take vengeance?")
bullet("People take vengeance to restore a sense of personal power, balance the "
       "scales of perceived injustice, and ease emotional pain.")
bullet("Vengeance acts as a deep-seated, often unconscious, drive to reclaim "
       "control after a betrayal or injury.")
label("Psychological reasons for vengeance")
bullet2("Restoring control. ", "Retaliation helps a person feel less helpless "
        "and vulnerable after being hurt.")
bullet2("Sending a message. ", "Many seek to make the wrongdoer understand the "
        "depth of their pain and realize that crossing boundaries has a cost.")
bullet2("Deterrence. ", "Vengeance serves as a personal defence mechanism, to "
        "warn an offender against repeating harmful actions.")
bullet2("Revenge as reward. ", "Brain scans show that planning and executing "
        "payback lights up the brain's reward and pleasure centres, making it feel "
        "temporarily satisfying.")
bullet2("A bruised ego. ", "Revenge is heavily tied to a bruised ego, deep "
        "humiliation, and a feeling that formal systems have failed to deliver "
        "true justice.")

# ---- 6 ----
qheading("6. Why did people take vengeance in the Bible?")
lead("People in the Bible often sought personal vengeance because of deep pain, "
     "anger, and a natural human desire for justice when wronged. However, the "
     "Bible consistently discourages personal revenge, teaching that retaliation "
     "leads to bitterness and that ultimate judgement belongs to God alone.")
label("Biblical teaching on vengeance")
bullet2("Vengeance is forbidden for humans. ", "The Bible explicitly tells "
        "individuals not to pay back evil with evil or hold grudges "
        "(Leviticus 19:18; Romans 12:19).")
bullet2("Vengeance is reserved for God. ", "Scripture states that vengeance "
        "belongs to God, who judges fairly and repays in His own time "
        "(Deuteronomy 32:35).")
bullet2("Christians are called to forgive. ", "Jesus instructed His followers to "
        "love their enemies and pray for those who hurt them, rather than seeking "
        "payback (Matthew 5:43-45).")

# ---- 7 ----
qheading("7. Examples of those who took revenge in the Bible")
lead("Several biblical figures sought personal revenge after facing a grave "
     "wrong, betrayal, or personal injury. These include:")
bullet2("Simeon and Levi, ", "who slaughtered the men of Shechem to avenge the "
        "violation of their sister Dinah (Genesis 34:25).")
bullet2("Samson, ", "who repeatedly attacked the Philistines to retaliate for "
        "personal insults, the loss of his wife, and the blinding of his eyes "
        "(Judges 15:7; 16:28).")
bullet2("Joab, ", "who assassinated Abner under the guise of conversation to "
        "avenge the blood of his brother Asahel (2 Samuel 3:27).")
bullet2("Absalom, ", "who waited years before killing his half-brother Amnon to "
        "avenge the assault of his sister Tamar (2 Samuel 13:28).")

# ================= SAVE =================
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Vengeance - draft.docx")
try:
    doc.save(out)
except PermissionError:
    raise SystemExit("Cannot write '%s'. It is open in Word. Close it and run again." % out)
print("Saved:", out)
print("Paragraphs:", len(doc.paragraphs))
