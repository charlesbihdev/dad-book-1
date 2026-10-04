# Dad's Book, Project Brief

Read this first. It explains what this project is, who is involved, how we work, and the rules that must never be broken.

## What this is

Charles Owusu Bih is helping his father transcribe and publish a religious book. His father writes the book by hand (pen on paper) and sends photos of the pages. Our job is to transcribe the handwriting and format it into a clean Microsoft Word document that matches the house style of his first book.

- Author: Paul Adasah Nkrumah, Preacher, S/Bekwai Church of Christ.
- Typed, edited and arranged by: Charles Owusu Bih.
- This is the author's SECOND book. The first, `Introduction - final.docx` (in the project root), is kept only as a style reference. Do not edit it.
- Bible translation: NIV (New International Version). Verify all quotations against the NIV.

## The 23 topics (done ONE at a time, in order)

1. Sabbath  2. Conscience  3. Lie  4. Dream  5. Anger  6. The Vengeance
7. The Devil  8. Light  9. Imitators  10. Name  11. Worship  12. Gifts of the Holy Spirit
13. Miracles  14. Holy Spirit  15. False Prophet  16. Tongues  17. Love  18. Tithe
19. Discipline  20. Holiness  21. Alcohol  22. Chief  23. Food

Finish and confirm one topic before moving to the next.

## The rules (non-negotiable; a religious book, mistakes are costly)

1. Never guess an unreadable word, letter, or character. Stop, point Charles to the exact image (and where on the page it is), and ask him to identify it. Do not make things up.
2. Verify every scripture reference against the actual verse. OCR and handwriting routinely flip references (for example Matt 1:20 misread as Mark 1:20). When a cited reference does not match the point being made, flag it and ask. Example: "Mark 1:20 is about John the Baptist, but this point is about the virgin birth; should it be Matt 1:20? Please check the image."
3. Never silently change or silently decide anything. Ask, Charles confirms; Charles asks, the agent confirms. When unsure, stop and ask.
4. Flag common-sense and context mismatches, not just scripture. Examples: an image in the `sabath/` folder whose content is actually about Love; pages out of order; a duplicate or missing page; a heading that does not match the topic folder; a stray image that does not belong. Surface it, do not fold it in quietly.

## Working directions (from Charles; apply to every topic)

- Ask less, proceed more. Do not hand Charles pages of transcription to proof-read. Transcribe and format directly. Only come with SHORT questions, and only when genuinely unsure: cannot read the handwriting, cannot see clearly, or a reference or fact looks wrong. If it reads clearly and is correct, just proceed. Prefer short multiple-choice questions over long prose.
- All content is part of the book. Typed pages and handwritten pages both go in, reformatted to flow as one continuous English book. Typed pages that are in Twi (Akan) get their scripture REPLACED with the English NIV, keeping the same structure and headers. Keep tables as tables (or propose a cleaner presentation).
- Reference fixes: when a reference is clearly a typo (a dropped or flipped digit, e.g. John 1:3 for John 1:33), correct it and list it under "Change made (please confirm)". When the right verse would be a different chapter or verse altogether (e.g. Exodus 1:9 vs Exodus 4:8-9), keep Dad's reference and flag it for him to review.
- Dad tends to keep his own wording, references, dates, and spellings. Flag only clear factual or scripture errors; do not fuss over style, transliteration, or dates he may have chosen on purpose.
- Do NOT use em dashes in the book text or in any file. Charles reads heavy em-dash use as AI slop. Put scripture references in plain parentheses, for example "(Jn 9:14-17)". Use commas, parentheses, or semicolons instead of em dashes. Use plain hyphens in verse ranges, for example 2:2-3. Exception: NIV scripture quotations are reproduced exactly as printed, em dashes included; scripture is a quotation and is never edited.
- Notes meant for Dad (the corrections files) must be PLAIN readable text: no markdown tables, no horizontal rule lines, no dotted lines. Just clear headings and sentences or short lists.
- Keep a clean environment. Everything that belongs to a topic lives inside that topic's folder, not in the project root. Keep old drafts separate from the current draft (see layout below).
- Do not edit PROJECT.md (or any file) just to log progress. Record standing directions here when Charles gives them, but not status or progress.
- Build one .docx per topic (its own review and correction cycle). Merge all topics into the full book at the very end, adding the title page, table of contents, and page numbering then. All topic builds share the same style settings so formatting stays identical across the book.

## Formatting spec (baseline from the first book)

Use this look as the default. The numbered-point structure and body layout are flexible for this book; propose a layout after seeing a topic's real pages.

- Font: Garamond, entire book.
- Body text: 13 pt, justified.
- Page size: A5 (148 by 210 mm).
- Margins: top about 0.55 in, bottom about 0.6 in, left about 0.6 in, right about 0.5 in.
- Chapter and question headings: bold (chapter title centered; numbered questions left-aligned).
- Scripture references: in plain parentheses at the end of the sentence (this book is prose teaching, so inline reads better than the first book's dotted-leader columns). Use dotted-leader columns only for genuine short "point then verse" lists.
- Quotations: italic.
- Footer: page numbers (roman i, ii for front matter; arabic 1, 2 for the body).
- Credit line: "Typed, edited, and arranged by Charles Owusu Bih." (italic, centered.)

## Folder layout (per topic)

Project root holds only: `PROJECT.md`, the style reference `Introduction - final.docx`, and one folder per topic.

Each topic folder (for example `sabath/`) holds:
- `images/` : all source photos for the topic, plus the original zip files.
- `build_<topic>.py` : the script that generates this topic's Word draft.
- `<Topic> - draft.docx` : the single CURRENT draft.
- `archive/` : superseded or old drafts only (never mixed with the current draft).
- `PAGE-MAP.md` : which image file maps to which numbered page.
- `CORRECTIONS - <Topic> (for Dad).md` : plain-text list of flagged facts and references.

To regenerate a topic draft: run its `build_<topic>.py` from inside the topic folder.

Agent memory (durable context across sessions) lives in the Claude projects memory directory, not here.

## Filling in NIV text (handoff for an agent with web access)

Four chapters were typed with Twi scripture and are fully transcribed, but their NIV text could not be fetched (Bible sites were blocked in that session). Every quoted passage prints as "[NIV TEXT PENDING]" until filled in. Nothing else in these chapters needs redoing.

Where to fill it in: each script has a `VERSES = { }` dictionary near the top, just above the CONTENT section. Add one entry per passage, keyed by the exact reference string the script uses:
- `miracles/build_miracles.py` (40 passages)
- `holy_spirit/build_holy_spirit.py` (36 passages)
- `tongues/build_tongues.py` (18 passages)
- `love/build_love.py` (35 passages)

Example entry: `"Hebrews 13:8": "Jesus Christ is the same yesterday and today and forever.",`

To see exactly which keys are still missing, run the script from inside its topic folder (for example `cd tongues && python3 build_tongues.py`). The last line lists every pending reference. Repeat until it reports 0, then commit the rebuilt drafts.

How to enter the text:
- Source: the NIV (2011) from an authoritative site such as biblegateway.com (version=NIV). Never type it from memory.
- Copy the words exactly as printed, em dashes and inner quotation marks included. Scripture is a quotation and is never edited (see the em dash rule above).
- Leave out verse numbers, section headings, and footnote or cross-reference markers. Join a multi-verse passage into one string with single spaces.
- Do not add outer quotation marks or the reference; `scripture()` adds both.
- Split references: `1 Corinthians 14:27a` is the first clause of the verse (how many may speak), and `1 Corinthians 14:27b` is the rest (one at a time, someone must interpret). Split the NIV verse at that point, without changing any words.
- The keys are the references Dad cited, already checked. Some were deliberately set or kept (for example John 1:33 in Holy Spirit; Acts 8:14-18 in Tongues, where the Twi had wrongly pasted Acts 10:14-18). Quote the key's own passage; do not swap references. Questions about references belong in that topic's CORRECTIONS file, not in the script.
