"""Generate May 2026 Committee Meeting Minutes as a .docx file matching the April 2026 styling."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLUE = RGBColor(0x2E, 0x75, 0xB6)
LIGHT_GREY = "F2F2F2"

doc = Document()

# ---- Page setup (A4, 1 inch margins) ----
section = doc.sections[0]
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1)
section.right_margin = Inches(1)

# ---- Default font ----
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

# ---- Title ----
title = doc.add_paragraph()
title.paragraph_format.space_after = Pt(12)
run = title.add_run("Meeting Minutes")
run.font.name = "Calibri"
run.font.size = Pt(28)
run.font.bold = True
run.font.color.rgb = BLUE

# ---- Header info table ----
header_table = doc.add_table(rows=4, cols=2)
header_table.autofit = False
header_rows = [
    ("Organisation", "Southern Cross 4x4 Club"),
    ("Date", "19/05/2026"),
    ("Time", "7:32 PM – [end time TBC]"),
    ("Chair", "Trevor Ryan"),
]
for i, (label, value) in enumerate(header_rows):
    row = header_table.rows[i]
    row.cells[0].width = Inches(1.6)
    row.cells[1].width = Inches(4.4)
    p_label = row.cells[0].paragraphs[0]
    r = p_label.add_run(label)
    r.bold = True
    r.font.size = Pt(11)
    p_value = row.cells[1].paragraphs[0]
    rv = p_value.add_run(value)
    rv.font.size = Pt(11)

# Light shading on label column
for i in range(4):
    tc_pr = header_table.rows[i].cells[0]._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), LIGHT_GREY)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)

# Spacer
doc.add_paragraph()


def add_heading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Calibri"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = BLUE


def add_para(text):
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(6)
    return p


def add_bullets(items):
    for item in items:
        p = doc.add_paragraph(item, style="List Bullet")
        p.paragraph_format.space_after = Pt(2)


# ---- Sections ----
add_heading("1. Opening")
add_para("The meeting was opened at 7:32 PM by Trevor, who welcomed the committee members in attendance.")

add_heading("2. Mid-Year Event")
add_para(
    "The committee discussed options for the mid-year function, including Christmas in July, a Winter Wonderland "
    "theme, and a mystery event proposed by David. The committee resolved to run the event in late June under the "
    "title “Kebabs in the Kountry”."
)
add_bullets([
    "When: 27–28 June 2026",
    "Where: Warburtons Bridge Campground, Drummond-Vaughan Rd, Glenluce VIC 3451",
    "Catering: Jamie to bring a spit, with Jerome and Ed assisting. Jamie may also donate the meat.",
    "Funding: Kebabs to be sold as a fundraiser; the club to cover incidentals. Donations from attendees to be collected in advance via cash or bank transfer.",
    "Promotion: Chris to set up and promote the event on Facebook.",
])

add_heading("3. Christmas Event")
add_para(
    "The committee broadly agreed to run a Christmas event in late November. It is expected there will be enough "
    "children in attendance to justify a Santa appearance. Brett’s property was discussed as a possible venue, "
    "and Trevor will speak with Brett about both hosting the event and joining the club."
)

add_heading("4. Membership Cap")
add_para(
    "The committee discussed whether to introduce a cap on membership numbers. It was agreed that no cap is required, "
    "as the nomination process will provide a natural limit on growth."
)

add_heading("5. Membership Process")
add_para("The committee confirmed the following process for new members:")
add_bullets([
    "A prospective member completes the application form.",
    "The prospect is then invited to the club Facebook group.",
    "The member liaison checks in with the prospect.",
    "The prospect must complete at least one trip and one meeting, typically over a two-month period, before being put to a vote (extenuating circumstances aside).",
])
add_para(
    "Pete and Pam will be put to a vote at the next meeting. Current prospects on record: Riddi, Jason, Joe, Marc, "
    "and Anonymous. Chris will investigate setting up anonymous voting via a QR code. Tynke was confirmed as the "
    "club’s member liaison."
)

add_heading("6. Email Signatures")
add_para("David has set up Gmail signatures for all committee members.")

add_heading("7. Multi-Club Trips")
add_para(
    "A member has proposed running trips jointly with other clubs. The committee does not oppose multi-club trips in "
    "principle but will consider each on a case-by-case basis."
)

add_heading("8. Printed Event Calendar")
add_para(
    "Trevor proposed printing a calendar of events, likely in A5 format. Sponsors to be listed include:"
)
add_bullets([
    "ARB",
    "Extreme",
    "On Track",
    "Werribee Isuzu",
    "Max Liner Australia",
    "Concept 2",
])
add_para(
    "Shirt printing is currently delayed as the printer is at capacity. Chris will mock up a draft of the printed "
    "calendar."
)

add_heading("9. Trip Co-ordination")
add_para(
    "Trevor raised the need for a trip co-ordinator to ensure trips are well organised and that the correct "
    "documentation is completed — including sign-in sheets, visitor forms, a trip-running checklist, and "
    "visitor waivers. Forms will need to be linked from the Facebook page. Chris volunteered to co-ordinate trip "
    "documentation but not to take on the role of trip co-ordinator."
)

add_heading("10. Car Boot Sale")
add_para("Chris proposed running a car boot sale as a fundraiser for the club. Possible venues include:")
add_bullets([
    "Tarneit Football Club — possibly on a Saturday; to enquire about availability on away-game weekends.",
    "The Racecourse — suggested by Trevor.",
])
add_para(
    "A permit may be required, and a certificate of currency for public liability will be obtained from 4WD Victoria."
)

add_heading("11. By-Laws / Model Rules")
add_para(
    "Chris will share the proposed changes to the model rules on the Facebook page for member review ahead of the AGM."
)

add_heading("12. AGM and Committee Continuity")
add_para(
    "The committee proposes putting to the members that the current committee continue into the next year. Daryl "
    "will run the AGM."
)

add_heading("13. Welfare Officer")
add_para(
    "Trevor proposed the appointment of a welfare officer to reach out to members who are going through difficult "
    "times. The club will be mindful of opportunities to send cards, check in with members, and pass on get-well "
    "wishes."
)

add_heading("14. Next Meeting")
add_para("Tynke will speak at the next meeting about the recent club trip.")

add_heading("15. Closure")
add_para("There being no further business, the meeting was declared closed.")

p_conf = doc.add_paragraph()
p_conf.paragraph_format.space_before = Pt(12)
r1 = p_conf.add_run("Confirmed by: ")
r1.bold = True
p_conf.add_run("Chris Butterworth")
r2 = p_conf.add_run("\t\tDate: ")
r2.bold = True
p_conf.add_run("19/05/2026")

# Spacer
doc.add_paragraph()

# ---- Attendee table ----
attendees = ["Trevor Ryan", "David Powell", "Tynke Vrowe", "Jerome Muller", "Chris Butterworth"]
att_table = doc.add_table(rows=1, cols=len(attendees))
att_table.style = "Table Grid"
for i, name in enumerate(attendees):
    cell = att_table.rows[0].cells[i]
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(name)
    r.font.size = Pt(10)
    # cell padding via tcMar
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for side, val in [("top", "80"), ("bottom", "80"), ("left", "120"), ("right", "120")]:
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), val)
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)

out_path = r"C:\Users\cbutt\Downloads\May 2026 - Committee Meeting Minutes.docx"
doc.save(out_path)
print(f"Wrote {out_path}")
