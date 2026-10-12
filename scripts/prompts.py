"""Schemas and prompts sent to Claude for each kind of ledger page.

The wording here is the result of real transcription mistakes (see
HOW_IT_WORKS.md); change it deliberately, not as cleanup.
"""

# Right-hand pages: Offence / Parish / Residence (process_reports.py)
EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "rows": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "row_number": {"type": "integer"},
                    "crime": {"type": "string"},
                    "parish": {"type": "string"},
                    "home_address": {"type": "string"},
                    "status": {"type": "string", "enum": ["ok", "review"]},
                    "notes": {"type": "string"},
                },
                "required": [
                    "row_number", "crime", "parish",
                    "home_address", "status", "notes",
                ],
            },
        }
    },
    "required": ["rows"],
}

PROMPT_TEMPLATE = """Read the image at this exact path: {image_path}

This is a scanned page from a historical convict/prison register. It contains a table with multiple prisoner records, one per row. Most pages in this register are ruled for 21 entries, but a genuine page can legitimately have fewer (or, rarely, more) -- 21 is a common count, not a required one. Never force your count to hit 21.

Each physical entry in the table is delimited by the horizontal ruled lines. Count rows by those ruled lines, top to bottom. Do not invent an extra row out of a remarks/continuation annotation, and do not merge two distinct ruled rows into one. If your count comes out different from 21, treat that as a normal possibility, not a problem -- but as a one-time sanity check, briefly re-scan the page for a row you may have accidentally split into two or merged into one. If the page genuinely only shows fewer (or more) real entries, report that true count and do not force it to 21.

IMPORTANT -- wrapped text is not a new row: when a field's text (crime/offence, residence, remarks, etc.) is too long to fit on one written line, the writer often continued it on a second handwritten line squeezed into the SAME ruled row, or slightly overflowing into the row below's blank space, before the next ruled line starts. That overflow/continuation text belongs to the row it visually overflowed from -- concatenate it into that same row's field instead of treating it as a separate row for a different person. Only start a new row when a genuinely new ruled line with its own new set of column entries (a new parish/residence/term/etc.) begins.

To get the row count right, use the "When Received" column (or whichever date column is always filled in) as your ground truth: it has exactly ONE date per real physical entry and is never wrapped across two lines. First count the dates in that column top to bottom -- that is the true number of rows on the page. Then match each crime/parish/home_address to one of those dated rows. If a long crime or address description looks like it might span two lines, but there is only one date for it, it is ONE row, not two -- merge the text back together rather than splitting it.

For EVERY row on the page, in top-to-bottom order, transcribe these fields EXACTLY as written. Do not correct spelling, do not modernize wording, do not expand abbreviations, do not paraphrase:
- crime: just the offense/crime name itself (may be labeled "Offence"). Leave out repeat-offense counters/numbers such as "(2nd)", "(3)", "2", etc. that are appended to the offense -- e.g. write "Burglary and Larceny", not "Burglary(2) and Larceny(1)".
- parish: the parish column
- home_address: the residence/home address column

Also include for each row:
- row_number: position on the page, starting at 1, top to bottom, matching the physical ruled-line order exactly
- status: "ok" unless you are genuinely unsure of the transcription; otherwise "review"
- notes: only if status is "review" -- briefly state which field is uncertain and why. Leave empty if status is "ok".

Critical rules:
- Ditto marks (", -"-, "do.", "same") and standard period ledger abbreviations are NORMAL and LEGIBLE -- resolve them from the value directly above/context and mark status "ok". Do not flag a row as "review" just because a value was written as a ditto mark or abbreviation you could confidently resolve.
- Only use status "review" when a field is genuinely illegible, ambiguous, smudged, torn, or you cannot confidently resolve it -- not for routine, confidently-resolved shorthand. Most rows should come back "ok".
- Never guess or fabricate a value you cannot actually read at all. If a field is truly illegible, leave that field as an empty string and set status to "review" with a short reason.
- Keep each field's text attached to the correct physical row -- do not shift values between rows.
- Include every row on the page, even ones partially cut off at the top or bottom edge.
"""


# Left-hand pages: Name / Age / Complexion (process_complexion.py, export_complexion.py)
COMPLEXION_SCHEMA = {
    "type": "object",
    "properties": {
        "rows": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "row_number": {"type": "integer"},
                    "complexion": {"type": "string"},
                    "status": {"type": "string", "enum": ["ok", "review"]},
                    "notes": {"type": "string"},
                },
                "required": ["row_number", "complexion", "status", "notes"],
            },
        }
    },
    "required": ["rows"],
}

COMPLEXION_PROMPT_TEMPLATE = """Read the image at this exact path: {image_path}

This is the OTHER half of a scanned historical convict/prison register page -- the half with columns like Name, Age, Height, and Complexion (sometimes labeled "Color"). It contains one row per prisoner, same physical rows as the matching offence/parish/residence half of this same page. Most pages in this register are ruled for 21 entries, but a genuine page can legitimately have fewer (or, rarely, more) -- 21 is a common count, not a required one. Never force your count to hit 21.

Each physical entry in the table is delimited by the horizontal ruled lines. Count rows by those ruled lines, top to bottom -- use whichever column is always filled in and never wrapped (e.g. Age, or a register/entry number) as your check for the true row count. If your count comes out different from 21, treat that as a normal possibility, not a problem -- but as a one-time sanity check, briefly re-scan the page for a row you may have accidentally split into two or merged into one. If the page genuinely only shows fewer (or more) real entries, report that true count and do not force it to 21.

IMPORTANT -- wrapped text is not a new row: if any field's text is too long to fit on one written line and continues on a second handwritten line squeezed into the SAME ruled row (or slightly overflowing into the row below's blank space) before the next ruled line starts, that overflow belongs to the row it visually overflowed from -- concatenate it into that row's field instead of treating it as a new row for a different person. Only start a new row when a genuinely new ruled line with its own new set of column entries begins.

For EVERY row on the page, in top-to-bottom order, transcribe:
- complexion: just the complexion/color entry, exactly as written (e.g. "White", "Black", "Mulatto", "Dark", etc.) -- do not paraphrase or standardize the wording.

Also include for each row:
- row_number: position on the page, starting at 1, top to bottom, matching the physical ruled-line order exactly
- status: "ok" unless you are genuinely unsure of the transcription; otherwise "review"
- notes: only if status is "review" -- briefly state why. Leave empty if status is "ok".

Critical rules:
- Ditto marks and standard ledger abbreviations are NORMAL and LEGIBLE -- resolve them from context and mark status "ok".
- Only use status "review" when genuinely illegible or ambiguous -- most rows should come back "ok".
- Never guess a value you cannot actually read. If illegible, leave complexion as an empty string and set status "review".
- Include every row on the page, even ones partially cut off at the top or bottom edge.
"""
