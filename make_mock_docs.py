"""Create FICTIONAL source documents for a past exercise ("Exercise RIVER WATCH").

These stand in for your real exercise documents so the pipeline can be tried
end to end. Replace the contents of source_docs/ with your real files later.

Run:  python make_mock_docs.py
"""
from pathlib import Path

from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

OUT = Path(__file__).parent / "source_docs"

NAME = "Exercise RIVER WATCH 2025"
ORG = "Marlow Valley Regional Emergency Management Group"
BANNER = "EXERCISE EXERCISE EXERCISE"

AIM = (
    "To practise the activation and operation of the Marlow Valley Emergency "
    "Operations Centre (EOC) and the coordination of partner agencies during a "
    "rapid-onset riverine flood affecting the township of Wexley."
)

OBJECTIVES = [
    "O1 - Activate the EOC and establish a common operating picture within 30 minutes of the first warning.",
    "O2 - Make and communicate a timely evacuation decision for at-risk zones (target: within 15 minutes of trigger).",
    "O3 - Coordinate multi-agency resources, including requests for mutual aid.",
    "O4 - Deliver consistent public information and manage rumour through a single spokesperson.",
    "O5 - Identify and protect vulnerable people and facilities.",
    "O6 - Maintain operations through a loss of primary communications.",
]

AGENCIES = [
    ("Marlow Valley EOC", "Lead agency. Controller, Operations, Planning, Logistics, Public Information (PIO)."),
    ("Marlow Valley Fire and Rescue", "Swift-water rescue, pumping, door-knocking support."),
    ("Marlow Valley Police", "Road closures, cordons, evacuation support."),
    ("Regional Ambulance Service", "Casualty care, movement of non-ambulant patients."),
    ("Council Public Works", "River gauges, roads, sandbags, plant and equipment."),
    ("Community Welfare Team", "Evacuation centres, registration, welfare needs."),
    ("Valley Power", "Electricity network status and isolation decisions."),
    ("Exercise Control (EXCON)", "Delivers injects, simulates all non-playing agencies, media and public."),
]

SCENARIO = [
    "A slow-moving low pressure system has stalled over the upper Tarrin catchment after a week of steady rain. Soils are saturated and the Tarrin Dam is at 96 percent capacity.",
    "The Tarrin River runs through the centre of Wexley (population 8,400). Flood Zone A (about 340 homes, Millbrook Caravan Park and Wexley Aged Care) floods when the Wexley Bridge gauge passes 4.0 m.",
    "Gauge levels: minor flood 3.5 m, moderate 4.0 m, major 4.5 m. The record flood of 1998 peaked at 4.7 m.",
    "It is a weekday in late winter. Schools are open. The exercise starts at 09:00 with the gauge at 2.9 m and rising.",
]

PHASES = [
    ("Phase 1 - Warning and activation", "09:00", "10:00", "Initial warnings, EOC activation, first reports of surface flooding."),
    ("Phase 2 - Escalation and evacuation", "10:00", "11:20", "Forecast upgraded, evacuation decision for Zone A, resources stretched."),
    ("Phase 3 - Peak impact", "11:20", "12:40", "Power isolation, vulnerable facilities, evacuation centre capacity, communications failure."),
    ("Phase 4 - Transition to recovery", "12:40", "13:00", "River peaks and recedes, handover to recovery, ENDEX."),
]

# number, time, from, to, method, message, expected actions, objective
INJECTS = [
    (1, "09:00", "Met Service duty forecaster", "EOC Duty Officer", "Email",
     "Severe weather warning issued for the upper Tarrin catchment. A further 120 to 160 mm of rain is expected in the next 24 hours. The Tarrin River is likely to reach moderate flood level at Wexley Bridge by 14:00.",
     ["Activate the EOC to Level 2", "Notify partner agencies and request liaison officers", "Open the incident log and set a reporting rhythm"], "O1"),
    (2, "09:20", "Council Public Works (gauge telemetry)", "EOC Operations", "Phone",
     "Wexley Bridge gauge now reads 3.1 m and is rising at 0.2 m per hour. Minor flood level is 3.5 m.",
     ["Agree a gauge reporting interval", "Map properties at risk in Zone A", "Brief the Controller"], "O1"),
    (3, "09:40", "Police Communications", "EOC Operations", "Radio",
     "Riverside Road is closed by surface water between Mill Lane and the bridge. Two vehicles are stranded. No injuries reported.",
     ["Record the road closure and update the map", "Agree a traffic management plan with Police", "Pass road information to the PIO"], "O3"),
    (4, "10:00", "Member of the public (via call centre)", "Community Welfare Team", "Phone",
     "The manager of Millbrook Caravan Park is asking whether residents should leave. There are 60 residents on site, many of them elderly, and water is pooling at the entrance.",
     ["Assess the risk to the caravan park", "Refer an evacuation recommendation to the Controller", "Identify an evacuation centre and transport"], "O2, O5"),
    (5, "10:20", "Met Service duty forecaster", "EOC Controller", "Phone, confirmed by email",
     "Forecast upgraded. The river is now expected to peak at 4.6 m (major flood level) at about 16:00. Confidence is high.",
     ["Escalate the EOC to Level 3", "Decide on evacuation of Zone A within 15 minutes", "Issue an evacuation notice and task door-knocking teams"], "O2"),
    (6, "10:40", "Fire and Rescue Incident Commander", "EOC Operations", "Radio",
     "Both swift-water rescue teams are committed on Riverside Road and at Millbrook. We have no reserve. Requesting mutual aid for two further teams.",
     ["Submit a resource request to the regional coordination centre", "Prioritise current rescue tasks", "Record the resource gap in the situation report"], "O3"),
    (7, "11:00", "Valley FM radio reporter", "Public Information Officer", "Phone",
     "We would like a live interview at 11:15. We are also seeing posts on social media saying the Tarrin Dam is about to fail. Can you confirm?",
     ["Issue a holding statement", "Rebut the dam rumour with confirmed facts from the dam operator", "Nominate a single spokesperson"], "O4"),
    (8, "11:20", "Valley Power control room", "EOC Operations", "Email",
     "Wexley substation will be isolated at 12:00 for safety. About 2,100 customers will lose supply, including Wexley Aged Care (48 residents).",
     ["Identify vulnerable facilities and customers in the outage area", "Arrange a generator or relocation for Wexley Aged Care", "Include the outage in public messaging"], "O5"),
    (9, "11:40", "Community Welfare Team leader", "EOC Logistics", "Phone",
     "The evacuation centre at Wexley Community Hall has 160 people registered against a capacity of 200. We have no facilities for pets and people are arriving with animals.",
     ["Open a second evacuation centre", "Arrange catering, bedding and an animal holding area", "Update public information on where to go"], "O3, O5"),
    (10, "12:00", "Ambulance duty manager", "EOC Operations", "Radio",
     "Access to Wexley Aged Care via Bridge Street is now cut. We have six non-ambulant residents to move and need an alternative route or high-clearance vehicles.",
     ["Run a joint planning huddle with Ambulance, Fire and Public Works", "Source high-clearance vehicles or boats", "Confirm a receiving facility"], "O3, O5"),
    (11, "12:20", "Council Public Works", "EOC Controller", "Runner (written message)",
     "The radio repeater at Kestrel Hill has failed. Primary radio is unavailable across the district. Mobile coverage in Wexley is congested.",
     ["Activate the backup communications plan", "Confirm contact with every field team by alternative means", "Adjust the reporting rhythm"], "O6"),
    (12, "12:40", "Exercise Control", "All players", "Verbal announcement",
     "The river peaked at 4.4 m at 12:30 and is now falling. No further rain is forecast. Begin the transition to recovery. ENDEX will be called at 13:00.",
     ["Produce a final situation report", "Brief the incoming recovery manager", "List the first five recovery priorities"], "O1, O3"),
]

RULES = [
    f"All exercise messages begin and end with the words {BANNER}.",
    "The code words NO DUFF mean a real emergency or injury. Play stops immediately.",
    "Players respond as they would in a real event, using current plans and procedures.",
    "Do not contact anyone outside the exercise. EXCON simulates all outside agencies.",
    "Exercise time is real time. There are no time jumps.",
]

ARTIFICIALITIES = [
    "Weather and river behaviour are scripted and may move faster than a real event.",
    "No field resources are deployed. Field teams are simulated by EXCON.",
    "The dam operator, schools and media are simulated.",
]

DEBRIEF = [
    "What went well?",
    "What slowed us down?",
    "Did everyone share the same picture of the situation?",
    "What one thing would you change before the next real event?",
]

AAR_STRENGTHS = [
    "The EOC was activated and staffed within 18 minutes of the first warning (target 30 minutes).",
    "Joint planning between Ambulance, Fire and Public Works for the aged care move was fast and well structured.",
    "The incident log was complete and timed throughout.",
]

AAR_IMPROVE = [
    "The Zone A evacuation decision took 35 minutes against a target of 15. The trigger level was debated rather than applied.",
    "Situation reports from Operations and Planning used different figures for properties at risk.",
    "The PIO was not told about the dam rumour for 20 minutes and learned of it from the media call.",
    "Few players knew the backup communications plan. Contact with two field teams was not confirmed for 25 minutes.",
    "The vulnerable facilities register was out of date and did not list Millbrook Caravan Park.",
]

AAR_RECS = [
    "Adopt pre-agreed evacuation triggers tied to gauge levels and forecast confidence.",
    "Use one situation report template owned by Planning.",
    "Add a social media monitoring role to the PIO cell.",
    "Run a short communications-failure drill each quarter.",
    "Review the vulnerable facilities register every six months.",
    "Next exercise: use a different hazard and include an out-of-hours start.",
]


def bullets(doc, items):
    for item in items:
        doc.add_paragraph(item, style="List Bullet")


def table(doc, header, rows):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Light Grid Accent 1"
    for cell, text in zip(t.rows[0].cells, header):
        cell.text = text
    for row in rows:
        for cell, text in zip(t.add_row().cells, row):
            cell.text = str(text)
    doc.add_paragraph()


def exercise_plan():
    d = Document()
    d.add_heading(f"{NAME} - Exercise Plan", 0)
    d.add_paragraph(f"{ORG}. Functional exercise. Duration 4 hours (09:00 to 13:00). FICTIONAL SAMPLE DOCUMENT.")
    d.add_heading("1. Aim", 1)
    d.add_paragraph(AIM)
    d.add_heading("2. Objectives", 1)
    bullets(d, OBJECTIVES)
    d.add_heading("3. Participating agencies", 1)
    table(d, ["Agency", "Role in the exercise"], AGENCIES)
    d.add_heading("4. Scenario background", 1)
    for p in SCENARIO:
        d.add_paragraph(p)
    d.add_heading("5. Exercise phases", 1)
    table(d, ["Phase", "Start", "End", "Summary"], PHASES)
    d.add_heading("6. Rules of play", 1)
    bullets(d, RULES)
    d.add_heading("7. Assumptions and artificialities", 1)
    bullets(d, ARTIFICIALITIES)
    d.add_heading("8. Evaluation", 1)
    d.add_paragraph(
        "Evaluators are assigned to each EOC function. They record the time each expected "
        "action is completed and rate each objective as Met, Partly met or Not met. "
        "A hot debrief follows ENDEX and a written after-action report is issued within 14 days."
    )
    d.save(OUT / "01_Exercise_Plan.docx")


def events_list():
    d = Document()
    d.add_heading(f"{NAME} - Master Scenario Events List", 0)
    d.add_paragraph("EXCON EYES ONLY. Not for release to players. FICTIONAL SAMPLE DOCUMENT.")
    table(
        d,
        ["No.", "Time", "From", "To", "Method", "Inject message", "Expected actions", "Obj."],
        [(n, t, f, to, m, msg, "; ".join(exp), obj) for n, t, f, to, m, msg, exp, obj in INJECTS],
    )
    d.save(OUT / "02_Master_Scenario_Events_List.docx")


def after_action():
    d = Document()
    d.add_heading(f"{NAME} - After-Action Report", 0)
    d.add_paragraph("Summary of evaluator observations and the hot debrief. FICTIONAL SAMPLE DOCUMENT.")
    d.add_heading("1. Strengths", 1)
    bullets(d, AAR_STRENGTHS)
    d.add_heading("2. Areas for improvement", 1)
    bullets(d, AAR_IMPROVE)
    d.add_heading("3. Recommendations", 1)
    bullets(d, AAR_RECS)
    d.add_heading("4. Objective ratings", 1)
    table(d, ["Objective", "Rating"], [
        ("O1 EOC activation", "Met"), ("O2 Evacuation decision", "Not met"),
        ("O3 Resource coordination", "Met"), ("O4 Public information", "Partly met"),
        ("O5 Vulnerable people", "Partly met"), ("O6 Communications failure", "Not met"),
    ])
    d.save(OUT / "03_After_Action_Report.docx")


def content_slide(prs, title, lines, size=20, notes=None):
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = title
    tf = s.placeholders[1].text_frame
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
    if notes:
        s.notes_slide.notes_text_frame.text = notes
    return s


def table_slide(prs, title, header, rows, widths):
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = title
    shape = s.shapes.add_table(len(rows) + 1, len(header), Inches(0.5), Inches(1.6),
                               prs.slide_width - Inches(1), Inches(0.4) * (len(rows) + 1))
    for col, w in zip(shape.table.columns, widths):
        col.width = Inches(w)
    for r, row in enumerate([header] + list(rows)):
        for c, text in enumerate(row):
            cell = shape.table.cell(r, c)
            cell.text = str(text)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(12)


def briefing_deck():
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = NAME
    s.placeholders[1].text = f"Exercise briefing and inject deck\n{ORG}\n{BANNER}"

    content_slide(prs, "Exercise administration and safety", RULES)
    content_slide(prs, "Aim and objectives", [AIM] + OBJECTIVES, size=16)
    table_slide(prs, "Participating agencies", ["Agency", "Role in the exercise"], AGENCIES, [3.0, 6.0])
    content_slide(prs, "Scenario background", SCENARIO, size=16)
    table_slide(prs, "Exercise timeline", ["Phase", "Start", "End", "Summary"], PHASES, [3.0, 0.8, 0.8, 4.4])

    for n, t, f, to, m, msg, exp, obj in INJECTS:
        content_slide(
            prs, f"Inject {n} - {t}",
            [BANNER, f"From: {f}", f"To: {to}", f"Method: {m}", msg, BANNER],
            size=18,
            notes=f"FACILITATOR ONLY. Objective(s): {obj}. Expected actions: " + "; ".join(exp),
        )

    content_slide(prs, "ENDEX and hot debrief", DEBRIEF)
    content_slide(prs, "Next steps", [
        "Evaluator notes to EXCON by end of day.",
        "After-action report issued within 14 days.",
        "Thank you for taking part.",
    ])
    prs.save(OUT / "RIVER_WATCH_Briefing_Deck.pptx")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    exercise_plan()
    events_list()
    after_action()
    briefing_deck()
    for f in sorted(OUT.iterdir()):
        print("created", f.relative_to(OUT.parent))
