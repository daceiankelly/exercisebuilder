"""Exercise builder: learn the format of a past exercise, then generate new ones.

Steps (each is a subcommand):

  extract   Read every file in source_docs/ (.pptx .docx .pdf .txt .md) into text.
  profile   Ask the LLM to describe the exercise FORMAT (done once, then hand-edit).
  generate  Ask the LLM for a new exercise as structured JSON, then render it.
  render    Re-render the deck and facilitator guide from an (edited) exercise.json.
  models    List the models your API key can use.

Add --mock to profile/generate to use canned answers instead of calling the API.
"""
import argparse
import json
import os
import re
import sys
from pathlib import Path

ROOT = Path(__file__).parent
SOURCE = ROOT / "source_docs"
OUTPUT = ROOT / "output"
MOCK = ROOT / "mock_llm"
EXTRACTED = OUTPUT / "extracted_text.md"
PROFILE = OUTPUT / "exercise_profile.md"

# Change this, or set the OPENAI_MODEL environment variable, or pass --model.
# Run "python exercise_builder.py models" to see what your key can use.
DEFAULT_MODEL = "gpt-4.1"

# Above this size the raw documents are left out of the generate call and only
# the profile is sent (keeps the request inside the model's context window).
MAX_RAW_CHARS = 150_000


# ---------------------------------------------------------------- extraction

def read_pptx(path):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    def shape_text(shape):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            return [t for s in shape.shapes for t in shape_text(s)]
        if getattr(shape, "has_table", False) and shape.has_table:
            return [" | ".join(c.text.strip().replace("\n", " / ") for c in row.cells)
                    for row in shape.table.rows]
        if shape.has_text_frame and shape.text_frame.text.strip():
            return [shape.text_frame.text.strip()]
        return []

    out = []
    for i, slide in enumerate(Presentation(path).slides, 1):
        out.append(f"--- Slide {i} (layout: {slide.slide_layout.name}) ---")
        for shape in slide.shapes:
            out.extend(shape_text(shape))
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                out.append(f"[Speaker notes] {notes}")
        out.append("")
    return "\n".join(out)


def read_docx(path):
    from docx import Document
    from docx.table import Table

    out = []
    for block in Document(path).iter_inner_content():
        if isinstance(block, Table):
            for row in block.rows:
                out.append(" | ".join(c.text.strip().replace("\n", " / ") for c in row.cells))
            out.append("")
            continue
        text = block.text.strip()
        if not text:
            continue
        style = block.style.name if block.style is not None else ""
        if style == "Title":
            text = "# " + text
        elif style.startswith("Heading"):
            text = "## " + text
        elif style.startswith("List"):
            text = "- " + text
        out.append(text)
    return "\n".join(out)


def read_pdf(path):
    from pypdf import PdfReader

    return "\n\n".join((page.extract_text() or "") for page in PdfReader(path).pages)


READERS = {
    ".pptx": read_pptx,
    ".docx": read_docx,
    ".pdf": read_pdf,
    ".txt": lambda p: p.read_text(encoding="utf-8", errors="replace"),
    ".md": lambda p: p.read_text(encoding="utf-8", errors="replace"),
}


def cmd_extract(args):
    files = [f for f in sorted(SOURCE.iterdir())
             if f.suffix.lower() in READERS and not f.name.startswith("~$")]
    if not files:
        sys.exit(f"No readable documents found in {SOURCE}")
    parts = []
    for f in files:
        text = READERS[f.suffix.lower()](f).strip()
        print(f"  {f.name}: {len(text):,} characters")
        if not text:
            print("    WARNING: no text found (scanned PDF or image-only slides?)")
        parts.append(f"===== FILE: {f.name} =====\n\n{text}")
    OUTPUT.mkdir(exist_ok=True)
    EXTRACTED.write_text("\n\n\n".join(parts), encoding="utf-8")
    print(f"Wrote {EXTRACTED.relative_to(ROOT)}")


# ----------------------------------------------------------------------- LLM

def call_llm(system, user, model, schema=None):
    if not os.environ.get("OPENAI_API_KEY"):
        sys.exit("OPENAI_API_KEY is not set. Set it, or add --mock to try the pipeline without the API.")
    from openai import OpenAI

    kwargs = {}
    if schema:
        kwargs["response_format"] = {
            "type": "json_schema",
            "json_schema": {"name": "exercise", "strict": True, "schema": schema},
        }
    print(f"Calling {model} ...")
    resp = OpenAI().chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": system}, {"role": "user", "content": user}],
        **kwargs,
    )
    return resp.choices[0].message.content


def cmd_models(args):
    from openai import OpenAI

    for m in sorted(m.id for m in OpenAI().models.list()):
        print(m)


# ------------------------------------------------------------------- profile

PROFILE_SYSTEM = """You are an experienced emergency management exercise designer.
You will be given the documents from one past exercise (plan, events list, after-action
report, briefing deck). Describe the FORMAT of this exercise so that another designer
could write a new exercise, on a different scenario, that the same organisation would
recognise as theirs. Describe the pattern, do not just summarise the content.

Write Markdown with these sections:
1. Exercise type and overall shape (type, duration, real time or compressed, who plays, who is simulated)
2. Organisation and roles (agencies, EOC functions, reporting lines, exercise control)
3. Objectives (how they are worded and numbered, what capabilities they cover)
4. Phases and pacing (number of phases, inject spacing, how pressure escalates, where the unexpected complication falls)
5. Inject format (every field, typical length and tone of the message, mix of senders and methods, how expected actions are written)
6. Conventions (safety code words, message prefixes, terminology, time format, units)
7. Deck structure (slide by slide pattern, what goes in speaker notes)
8. Evaluation and debrief (how objectives are rated, debrief questions)
9. Lessons from the after-action report that the next exercise should test again
"""


def cmd_profile(args):
    if args.mock:
        text = (MOCK / "profile.md").read_text(encoding="utf-8")
    else:
        if not EXTRACTED.exists():
            cmd_extract(args)
        text = call_llm(PROFILE_SYSTEM, EXTRACTED.read_text(encoding="utf-8"), args.model)
    OUTPUT.mkdir(exist_ok=True)
    PROFILE.write_text(text, encoding="utf-8")
    print(f"Wrote {PROFILE.relative_to(ROOT)} - read it and correct anything that is wrong.")


# ------------------------------------------------------------------ generate

def _obj(props):
    return {"type": "object", "properties": props, "required": list(props), "additionalProperties": False}


_STR = {"type": "string"}
_STRS = {"type": "array", "items": _STR}

EXERCISE_SCHEMA = _obj({
    "title": _STR,
    "organisation": _STR,
    "exercise_type": _STR,
    "hazard": _STR,
    "location": _STR,
    "start_time": _STR,
    "end_time": _STR,
    "aim": _STR,
    "objectives": _STRS,
    "agencies": {"type": "array", "items": _obj({"agency": _STR, "role": _STR})},
    "scenario_background": _STRS,
    "phases": {"type": "array", "items": _obj({"name": _STR, "start": _STR, "end": _STR, "summary": _STR})},
    "injects": {"type": "array", "items": _obj({
        "number": {"type": "integer"},
        "time": _STR,
        "sender": _STR,
        "recipient": _STR,
        "method": _STR,
        "message": _STR,
        "expected_actions": _STRS,
        "objectives": _STR,
    })},
    "rules_of_play": _STRS,
    "assumptions_and_artificialities": _STRS,
    "evaluation_criteria": _STRS,
    "debrief_questions": _STRS,
    "facilitator_notes": _STRS,
})

GENERATE_SYSTEM = """You design emergency management exercises for the organisation whose
previous exercise is described below. Write a NEW exercise that this organisation would
recognise as its own: same structure, pacing, terminology, conventions, inject format and
level of detail as the previous one, but a different scenario.

Rules:
- Follow the exercise profile. Where the request conflicts with the profile, the request wins.
- Do not reuse injects or scenario details from the previous exercise.
- Inject times must fall inside the exercise window and inside their phase, in time order.
- Where the previous after-action report lists areas for improvement, include injects that test them again, and say so in facilitator_notes.
- Use the organisation's own agency names and roles unless the request changes them.
- Invent place names and details for the scenario, but do not invent real procedures, statutory powers or resource numbers for the organisation. Where you had to assume something, list it in assumptions_and_artificialities so a subject-matter expert can check it.
- Include at least one unexpected complication that is not signalled in the scenario background.
"""


def slugify(text):
    return re.sub(r"[^a-z0-9]+", "_", text.lower()).strip("_")[:60] or "exercise"


def cmd_generate(args):
    if args.mock:
        raw = (MOCK / "exercise.json").read_text(encoding="utf-8")
    else:
        if not PROFILE.exists():
            sys.exit("No exercise profile yet. Run:  python exercise_builder.py profile")
        system = GENERATE_SYSTEM + "\n\n# EXERCISE PROFILE\n\n" + PROFILE.read_text(encoding="utf-8")
        raw_docs = EXTRACTED.read_text(encoding="utf-8") if EXTRACTED.exists() else ""
        if raw_docs and len(raw_docs) <= MAX_RAW_CHARS:
            system += "\n\n# PREVIOUS EXERCISE DOCUMENTS (style reference)\n\n" + raw_docs
        request = [
            "Create a new exercise.",
            f"Hazard / scenario: {args.hazard}",
            f"Location: {args.location or 'designer to choose, within the organisation area'}",
            f"Start time: {args.start}",
            f"Duration: {args.duration} hours",
            f"Number of injects: {args.injects or 'similar to the previous exercise'}",
        ]
        if args.agencies:
            request.append(f"Participating agencies: {args.agencies}")
        if args.objectives:
            request.append(f"Objectives to emphasise: {args.objectives}")
        if args.notes:
            request.append(f"Other requirements: {args.notes}")
        raw = call_llm(system, "\n".join(request), args.model, schema=EXERCISE_SCHEMA)

    exercise = json.loads(raw)
    folder = OUTPUT / slugify(exercise["title"])
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "exercise.json"
    path.write_text(json.dumps(exercise, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {path.relative_to(ROOT)}")
    render(exercise, folder, args.template)


# -------------------------------------------------------------------- render

BANNER = "EXERCISE EXERCISE EXERCISE"


def find_template(explicit):
    if explicit:
        return Path(explicit)
    decks = [f for f in sorted(SOURCE.glob("*.pptx")) if not f.name.startswith("~$")]
    return decks[0] if decks else None


def pick_layouts(prs):
    """Find title, title+content and title-only layouts in whatever template is used."""
    from pptx.enum.shapes import PP_PLACEHOLDER as PH

    def kinds(layout):
        return [p.placeholder_format.type for p in layout.placeholders]

    layouts = list(prs.slide_layouts)
    title = next((l for l in layouts if PH.CENTER_TITLE in kinds(l)), layouts[0])

    def is_content(l):
        k = kinds(l)
        return PH.TITLE in k and sum(t in (PH.BODY, PH.OBJECT) for t in k) == 1

    content = (next((l for l in layouts if l.name == "Title and Content" and is_content(l)), None)
               or next((l for l in layouts if is_content(l)), layouts[min(1, len(layouts) - 1)]))
    title_only = next((l for l in layouts if PH.TITLE in kinds(l)
                       and not any(t in (PH.BODY, PH.OBJECT, PH.PICTURE) for t in kinds(l))), content)
    return title, content, title_only


def body_placeholder(slide):
    from pptx.enum.shapes import PP_PLACEHOLDER as PH

    for p in slide.placeholders:
        if p.placeholder_format.type in (PH.BODY, PH.OBJECT, PH.SUBTITLE):
            return p
    return None


def render_deck(ex, path, template):
    from pptx import Presentation
    from pptx.util import Inches, Pt

    prs = Presentation(str(template)) if template else Presentation()
    # Keep the template's theme and layouts, remove its slides.
    id_list = prs.slides._sldIdLst
    for sld_id in list(id_list):
        prs.part.drop_rel(sld_id.rId)
        id_list.remove(sld_id)
    title_layout, content_layout, table_layout = pick_layouts(prs)

    def fill(slide, lines, size):
        body = body_placeholder(slide)
        if body is None:
            return
        tf = body.text_frame
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = line
            if size:
                run.font.size = Pt(size)

    def content(title, lines, size=20, notes=None):
        s = prs.slides.add_slide(content_layout)
        s.shapes.title.text = title
        fill(s, lines, size)
        if notes:
            s.notes_slide.notes_text_frame.text = notes

    def table(title, header, rows, weights):
        s = prs.slides.add_slide(table_layout)
        s.shapes.title.text = title
        for p in list(s.placeholders):  # fallback layout may carry an empty body box
            if p.placeholder_format.idx != 0:
                p._element.getparent().remove(p._element)
        width = prs.slide_width - Inches(1)
        shape = s.shapes.add_table(len(rows) + 1, len(header), Inches(0.5), Inches(1.6),
                                   width, Inches(0.4) * (len(rows) + 1))
        for col, w in zip(shape.table.columns, weights):
            col.width = int(width * w / sum(weights))
        for r, row in enumerate([header] + rows):
            for c, text in enumerate(row):
                cell = shape.table.cell(r, c)
                cell.text = str(text)
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(12)

    s = prs.slides.add_slide(title_layout)
    s.shapes.title.text = ex["title"]
    fill(s, ["Exercise briefing and inject deck", ex["organisation"], BANNER], None)

    content("Exercise administration and safety", ex["rules_of_play"])
    content("Aim and objectives", [ex["aim"]] + ex["objectives"], size=16)
    table("Participating agencies", ["Agency", "Role in the exercise"],
          [[a["agency"], a["role"]] for a in ex["agencies"]], [1, 2])
    content("Scenario background", ex["scenario_background"], size=16)
    table("Exercise timeline", ["Phase", "Start", "End", "Summary"],
          [[p["name"], p["start"], p["end"], p["summary"]] for p in ex["phases"]], [3, 0.8, 0.8, 4.4])

    for inj in ex["injects"]:
        size = 18 if len(inj["message"]) < 320 else 15
        content(
            f"Inject {inj['number']} - {inj['time']}",
            [BANNER, f"From: {inj['sender']}", f"To: {inj['recipient']}",
             f"Method: {inj['method']}", inj["message"], BANNER],
            size=size,
            notes=f"FACILITATOR ONLY. Objective(s): {inj['objectives']}. Expected actions: "
                  + "; ".join(inj["expected_actions"]),
        )

    content("ENDEX and hot debrief", ex["debrief_questions"])
    prs.save(str(path))


def render_guide(ex, path):
    from docx import Document

    d = Document()

    def bullets(items):
        for item in items:
            d.add_paragraph(item, style="List Bullet")

    def table(header, rows):
        t = d.add_table(rows=1, cols=len(header))
        t.style = "Light Grid Accent 1"
        for cell, text in zip(t.rows[0].cells, header):
            cell.text = text
        for row in rows:
            for cell, text in zip(t.add_row().cells, row):
                cell.text = str(text)
        d.add_paragraph()

    d.add_heading(f"{ex['title']} - Facilitator Guide", 0)
    d.add_paragraph(f"{ex['organisation']}. {ex['exercise_type']}. {ex['start_time']} to {ex['end_time']}. "
                    f"Hazard: {ex['hazard']}. Location: {ex['location']}.")
    d.add_paragraph("DRAFT generated by an AI model. Must be reviewed by a subject-matter expert before use.")
    d.add_heading("1. Aim", 1)
    d.add_paragraph(ex["aim"])
    d.add_heading("2. Objectives", 1)
    bullets(ex["objectives"])
    d.add_heading("3. Participating agencies", 1)
    table(["Agency", "Role in the exercise"], [[a["agency"], a["role"]] for a in ex["agencies"]])
    d.add_heading("4. Scenario background", 1)
    for p in ex["scenario_background"]:
        d.add_paragraph(p)
    d.add_heading("5. Exercise phases", 1)
    table(["Phase", "Start", "End", "Summary"],
          [[p["name"], p["start"], p["end"], p["summary"]] for p in ex["phases"]])
    d.add_heading("6. Rules of play", 1)
    bullets(ex["rules_of_play"])
    d.add_heading("7. Assumptions and artificialities (check these)", 1)
    bullets(ex["assumptions_and_artificialities"])
    d.add_heading("8. Master scenario events list", 1)
    table(["No.", "Time", "From", "To", "Method", "Inject message", "Expected actions", "Obj."],
          [[i["number"], i["time"], i["sender"], i["recipient"], i["method"], i["message"],
            "; ".join(i["expected_actions"]), i["objectives"]] for i in ex["injects"]])
    d.add_heading("9. Evaluation criteria", 1)
    bullets(ex["evaluation_criteria"])
    d.add_heading("10. Debrief questions", 1)
    bullets(ex["debrief_questions"])
    d.add_heading("11. Facilitator notes", 1)
    bullets(ex["facilitator_notes"])
    d.save(str(path))


def check(ex):
    """Cheap sanity checks on what the model returned."""
    problems = []
    times = [i["time"] for i in ex["injects"]]
    if times != sorted(times):
        problems.append("inject times are not in order")
    for i in ex["injects"]:
        if not (ex["start_time"] <= i["time"] <= ex["end_time"]):
            problems.append(f"inject {i['number']} at {i['time']} is outside {ex['start_time']}-{ex['end_time']}")
    return problems


def render(exercise, folder, template_arg=None):
    template = find_template(template_arg)
    deck = folder / "briefing_deck.pptx"
    guide = folder / "facilitator_guide.docx"
    render_deck(exercise, deck, template)
    render_guide(exercise, guide)
    print(f"Wrote {deck} (template: {template.name if template else 'blank'})")
    print(f"Wrote {guide}")
    for p in check(exercise):
        print("  CHECK:", p)


def cmd_render(args):
    path = Path(args.json)
    render(json.loads(path.read_text(encoding="utf-8")), path.parent, args.template)


# ----------------------------------------------------------------------- CLI

def main():
    model = os.environ.get("OPENAI_MODEL", DEFAULT_MODEL)
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="command", required=True)

    sub.add_parser("extract").set_defaults(func=cmd_extract)
    sub.add_parser("models").set_defaults(func=cmd_models)

    p = sub.add_parser("profile")
    p.add_argument("--model", default=model)
    p.add_argument("--mock", action="store_true")
    p.set_defaults(func=cmd_profile)

    g = sub.add_parser("generate")
    g.add_argument("--hazard", default="designer to choose, different from the previous exercise")
    g.add_argument("--location")
    g.add_argument("--start", default="09:00")
    g.add_argument("--duration", type=float, default=4)
    g.add_argument("--injects", type=int)
    g.add_argument("--agencies")
    g.add_argument("--objectives")
    g.add_argument("--notes")
    g.add_argument("--template", help="deck to use as the slide template (default: first .pptx in source_docs)")
    g.add_argument("--model", default=model)
    g.add_argument("--mock", action="store_true")
    g.set_defaults(func=cmd_generate)

    r = sub.add_parser("render")
    r.add_argument("json", help="path to an exercise.json")
    r.add_argument("--template")
    r.set_defaults(func=cmd_render)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
