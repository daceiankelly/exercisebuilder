# Exercise builder

Learns the format of a past disaster management exercise from its documents,
then uses the OpenAI API to write new exercises in the same format and renders
each one as a briefing deck and a facilitator guide.

## Folders

| Path | What it is |
|---|---|
| `source_docs/` | The past exercise. Currently FICTIONAL mock files (Exercise RIVER WATCH). Replace with your real `.pptx`, `.docx`, `.pdf`, `.txt` files. |
| `output/extracted_text.md` | Everything read from `source_docs/`, as plain text. |
| `output/exercise_profile.md` | The model's description of your exercise format. Hand-edit this. |
| `output/<exercise name>/` | One folder per generated exercise: `exercise.json`, `briefing_deck.pptx`, `facilitator_guide.docx`. |
| `mock_llm/` | Canned answers used by `--mock` so the pipeline runs without an API key. |

## Use

All commands are run from this folder in PowerShell.

Set your key and model for the current window (do not put the key in a file):

    $env:OPENAI_API_KEY = "sk-..."
    $env:OPENAI_MODEL = "gpt-4.1"

See which models your key can use:

    .\.venv\Scripts\python.exe exercise_builder.py models

1. Read the documents:

       .\.venv\Scripts\python.exe exercise_builder.py extract

2. Build the format profile (once per set of documents), then open
   `output/exercise_profile.md` and correct it:

       .\.venv\Scripts\python.exe exercise_builder.py profile

3. Generate a new exercise:

       .\.venv\Scripts\python.exe exercise_builder.py generate --hazard "magnitude 6.8 earthquake" --location "Wexley" --start 02:00 --duration 4 --notes "Out-of-hours start. Test the backup communications plan."

   Other options: `--injects 12`, `--agencies "..."`, `--objectives "..."`, `--template other_deck.pptx`.

4. Edit `exercise.json` by hand if needed, then rebuild the deck and guide:

       .\.venv\Scripts\python.exe exercise_builder.py render output\<exercise name>\exercise.json

Add `--mock` to `profile` or `generate` to try the pipeline without calling the API.

## Notes

- The new deck uses the first `.pptx` in `source_docs/` as its template, so it
  keeps your theme and slide layouts. The slide order is fixed in `render_deck`
  in `exercise_builder.py`; change it there to match your real deck.
- Scanned PDFs and text inside images are not read. `extract` warns when a file
  yields no text.
- Generated exercises are drafts. Every one needs review by a subject-matter
  expert, especially the "Assumptions and artificialities" section, where the
  model is told to list anything it had to assume.
- Check that you are allowed to send your real documents to an external API
  before replacing the mock files. Redact names and contact details first if needed.
- To recreate the mock documents: `.\.venv\Scripts\python.exe make_mock_docs.py`
