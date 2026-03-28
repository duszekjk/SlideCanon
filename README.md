# Markdown -> PPTX builder with real template layouts

This script is meant for a lecture-style markdown format like:

- `#slide N`
- `## layout`
- `## slide title`
- `## subtitle`
- `## bullets`
- `## term` / `## definition`
- `## quote` / `## attribution`
- `## teleprompter`
- optional: `## image`, `## photo`, `## visual`, `## visualization`, `## chart`, `## youtube`

## What it does

- loads an existing PPTX template,
- creates new slides from the template's **real layouts**,
- fills **template placeholders** instead of drawing text boxes by hand,
- copies teleprompter text into presenter notes,
- optionally inserts images into picture placeholders,
- removes **local text/font overrides** from slides and notes,
- removes slide numbers from layouts/masters.

## Expected template layouts

The script was written against templates like:

- `Title`
- `Title & Bullets`
- `Quote`
- `Title, Bullets & Photo`
- `Title & Photo`
- `Title Only`
- `Section`

It tries these by name, in order, and uses the first matching layout.

## Repository contents

- `md_template_pptx_builder.py`: main builder script
- `lecture_template.pptx`: minimal starter template with the original slide master and layout set preserved
- `example_lecture.md`: small example input file
- `example-assets/sample-photo.png`: tiny image asset used by the example template and markdown

## Usage

```bash
python md_template_pptx_builder.py \
  --markdown /path/to/example_lecture.md \
  --template /path/to/lecture_template.pptx \
  --output /path/to/example_lecture_out.pptx \
  --verbose
```

From this repository root, that becomes:

```bash
python md_template_pptx_builder.py \
  --markdown example_lecture.md \
  --template lecture_template.pptx \
  --output example_lecture_out.pptx \
  --verbose
```

## Notes

- Local run-level font overrides are removed from slides and notes after the file is saved.
- It supports image fields written either as a plain path or markdown image syntax, for example:
  - `hotp_formula.png`
  - `![Wzór HOTP](hotp_formula.png)`
  - `![Wzór HOTP](sandbox:/mnt/data/hotp_formula.png)`
- For `sandbox:/...` references copied from ChatGPT, it also tries the basename, so `sandbox:/mnt/data/hotp_formula.png` resolves to `hotp_formula.png` next to the markdown file.
- If the exact path is not found, it also searches by filename recursively under the markdown directory, which helps when the markdown contains abstract paths but the assets sit next to it in a local folder.
- The script does not invent diagrams from prose. It only places existing image assets into picture placeholders.
- For YouTube links, it adds the URL into the body placeholder as clickable text.
- Run with `--verbose` to print warnings for slides whose images could not be resolved.
- The repository intentionally does not include a `presentations/` archive of generated lecture decks.

## Important limitation

If the markdown content is much longer than the template placeholder was designed for, PowerPoint will still overflow or shrink text visually depending on the template behavior. This script intentionally does **not** set per-slide font sizes, because the goal is to inherit from the template.

## Environment setup

Create and populate a virtual environment:

```bash
cd /path/to/pptx_markdown_builder
./setup_venv.sh .venv
source .venv/bin/activate
```

If your system uses a specific Python binary:

```bash
PYTHON_BIN=python3.11 ./setup_venv.sh .venv
```

Manual alternative:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip setuptools wheel
python -m pip install -r requirements.txt
```
