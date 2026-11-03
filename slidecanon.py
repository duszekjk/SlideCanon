#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

from PIL import Image as PILImage
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt
try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (
        Flowable,
        HRFlowable,
        Image as RLImage,
        KeepTogether,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )
except ImportError:  # pragma: no cover - handled at runtime
    colors = None
    TA_CENTER = None
    A4 = None
    ParagraphStyle = None
    getSampleStyleSheet = None
    mm = None
    pdfmetrics = None
    TTFont = None
    Flowable = None
    HRFlowable = None
    KeepTogether = None
    PageBreak = None
    RLImage = None
    Paragraph = None
    SimpleDocTemplate = None
    Spacer = None
    Table = None
    TableStyle = None

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
for k, v in NS.items():
    ET.register_namespace(k, v)

IMAGE_KEYS = ["image", "photo", "visual", "visualization", "visualisation", "chart", "picture"]
IMAGE_MARKDOWN_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
SLIDE_SPLIT_RE = re.compile(r"(?m)^#slide\s+(\d+)\s*$")
SECTION_RE = re.compile(r"(?m)^##\s+([^\n]+)\n")
STORY_LAYOUTS = {"story", "narrative", "case", "case study"}
COMPARISON_LAYOUTS = {"comparison", "compare", "split", "two-column", "two columns", "two column"}
IMAGE_LAYOUTS = {
    "photo",
    "image",
    "visual",
    "visualisation",
    "visualization",
    "chart",
    "bullet + image",
    "bullet+image",
    "image + bullet",
    "image+bullet",
}


@dataclass
class SlideSpec:
    number: int
    fields: Dict[str, str] = field(default_factory=dict)

    def get(self, key: str, default: str = "") -> str:
        return self.fields.get(key.lower().strip(), default)


@dataclass
class ImageSpec:
    kind: str
    raw_ref: str
    alt_text: str = ""
    path: Optional[Path] = None


@dataclass
class PicturePlaceholderInfo:
    shape: object
    name: str
    idx: Optional[int]


@dataclass
class PdfStyles:
    cover_kicker: ParagraphStyle
    cover_title: ParagraphStyle
    cover_subtitle: ParagraphStyle
    toc_heading: ParagraphStyle
    toc_item: ParagraphStyle
    title: ParagraphStyle
    slide_heading: ParagraphStyle
    subtitle: ParagraphStyle
    body: ParagraphStyle
    bullet: ParagraphStyle
    callout_label: ParagraphStyle
    callout_body: ParagraphStyle
    notes_label: ParagraphStyle
    notes: ParagraphStyle
    caption: Par# The same DejaVu Sans TTF files ship with Matplotlib on macOS and Linux.
# Unlike PDF's built-in Helvetica, this family covers Polish diacritics,
# Greek letters, and common mathematical superscript/subscript characters.
_PDF_GLYPH_COVERAGE = None
_PDF_MATH_FONT = None


def register_pdf_fonts() -> Dict[str, str]:
    """Register portable, embedded PDF fonts; never silently use Helvetica."""
    global _PDF_GLYPH_COVERAGE, _PDF_MATH_FONT
    if pdfmetrics is None or TTFont is None:
        raise RuntimeError("PDF generation requires ReportLab.")

    try:
        import matplotlib
    except ImportError as exc:
        raise RuntimeError(
            "PDF generation requires Matplotlib's bundled DejaVu Sans fonts. "
            "Run: python -m pip install -r requirements.txt"
        ) from exc

    font_dir = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
    paths = {
        "regular": font_dir / "DejaVuSans.ttf",
        "bold": font_dir / "DejaVuSans-Bold.ttf",
        "italic": font_dir / "DejaVuSans-Oblique.ttf",
        "bold_italic": font_dir / "DejaVuSans-BoldOblique.ttf",
    }
    missing = [str(path) for path in paths.values() if not path.is_file()]
    if missing:
        raise RuntimeError(f"Missing bundled Unicode PDF fonts: {', '.join(missing)}")

    names = {
        "regular": "SlideCanonDejaVu-Regular",
        "bold": "SlideCanonDejaVu-Bold",
        "italic": "SlideCanonDejaVu-Italic",
        "bold_italic": "SlideCanonDejaVu-BoldItalic",
    }
    for style, path in paths.items():
        if names[style] not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(names[style], str(path)))
    pdfmetrics.registerFontFamily(
        names["regular"],
        normal=names["regular"], bold=names["bold"],
        italic=names["italic"], boldItalic=names["bold_italic"],
    )
    # Only characters available across all styles are used without font tags.
    # This avoids glyph loss in bold/italic headings and notes.
    coverage = [set(pdfmetrics.getFont(name).face.charToGlyph) for name in names.values()]
    _PDF_GLYPH_COVERAGE = set.intersection(*coverage)

    # Supplementary mathematical Unicode symbols (e.g. U+1D465 MATHEMATICAL
    # ITALIC SMALL X) are in Noto Sans Math, not the main DejaVu family.
    # This optional fallback is used only for glyphs missing from DejaVu.
    candidates = [
        Path("/usr/share/fonts/truetype/noto/NotoSansMath-Regular.ttf"),
        Path("/usr/local/share/fonts/NotoSansMath-Regular.ttf"),
        Path.home() / "Library/Fonts/NotoSansMath-Regular.ttf",
        Path("/Library/Fonts/NotoSansMath-Regular.ttf"),
    ]
    _PDF_MATH_FONT = None
    for candidate in candidates:
        if candidate.is_file():
            fontname = "SlideCanonNotoMath"
            if fontname not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont(fontname, str(candidate)))
            _PDF_MATH_FONT = fontname
            break
    return names


def escape(value: str) -> str:
    """XML-escape text, switching font only for non-DejaVu math glyphs.

    Missing glyphs produce actionable errors instead of black squares in PDFs.
    This helper is used for ReportLab Paragraph text, not PPTX XML.
    """
    if _PDF_GLYPH_COVERAGE is None:
        register_pdf_fonts()
    fragments = []
    plain = []
    math = []
    math_coverage = (
        set(pdfmetrics.getFont(_PDF_MATH_FONT).face.charToGlyph)
        if _PDF_MATH_FONT else set()
    )

    def flush_plain():
        if plain:
            fragments.append(xml_escape("".join(plain)))
            plain.clear()

    def flush_math():
        if math:
            fragments.append(f'<font name="{_PDF_MATH_FONT}">' + xml_escape("".join(math)) + '</font>')
            math.clear()

    for char in str(value):
        # Newlines, tabs and carriage returns are structural, not glyphs.
        if char in "\n\r\t" or ord(char) in _PDF_GLYPH_COVERAGE:
            flush_math()
            plain.append(char)
        elif ord(char) in math_coverage:
            flush_plain()
            math.append(char)
        else:
            raise ValueError(
                f"No embedded PDF font contains U+{ord(char):04X} "
                f"({char!r}). Install Noto Sans Math for additional math "
                "symbols, or use a formula image. SlideCanon will not "
                "silently render missing glyphs."
            )
    flush_math()
    flush_plain()
    return "".join(fragments)
turn fallback


class OverlapImage(Flowable):
    def __init__(
        self,
        path: Path,
        width: float,
        height: float,
        layout_reduction_ratio: float = 0.25,
        bleed_ratio: float = 0.05,
    ):
        super().__init__()
        self.path = str(path)
        self.drawWidth = width
        self.drawHeight = height
        self.layout_reduction_ratio = max(0.0, min(layout_reduction_ratio, 0.3))
        self.bleed_ratio = max(0.0, min(bleed_ratio, 0.2))
        self._image = RLImage(self.path, width=width, height=height)
        self.width = width
        self.height = height * max(0.1, 1.0 - self.layout_reduction_ratio)

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    def drawOn(self, canvas, x, y, _sW=0):
        upward_bleed = self.drawHeight * self.bleed_ratio
        self._image.drawOn(canvas, x, y - upward_bleed, _sW=_sW)


def parse_markdown(md_path: Path) -> List[SlideSpec]:
    text = md_path.read_text(encoding="utf-8", errors="ignore").replace("\r\n", "\n")
    matches = list(SLIDE_SPLIT_RE.finditer(text))
    slides: List[SlideSpec] = []
    for i, m in enumerate(matches):
        number = int(m.group(1))
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        block = text[start:end].strip()
        fields: Dict[str, str] = {}
        sec_matches = list(SECTION_RE.finditer(block))
        for j, sm in enumerate(sec_matches):
            key = sm.group(1).strip().lower().rstrip(":")
            sstart = sm.end()
            send = sec_matches[j + 1].start() if j + 1 < len(sec_matches) else len(block)
            value = block[sstart:send].strip()
            fields[key] = value
        # Auto-split long titles into subtitle when subtitle is missing.
        # If the author writes e.g. "Ciphering mode: decyzja, negocjacja i konsekwencje" as slide title
        # and omits subtitle, move the part after ':' into subtitle.
        if "subtitle" not in fields or not fields.get("subtitle", "").strip():
            raw_title = (fields.get("slide title") or fields.get("title") or "").strip()
            if ":" in raw_title:
                left, right = raw_title.split(":", 1)
                left = left.strip()
                right = right.strip()
                if left and right and len(right) >= 6:
                    if fields.get("slide title"):
                        fields["slide title"] = left
                    elif fields.get("title"):
                        fields["title"] = left
                    fields["subtitle"] = right

        slides.append(SlideSpec(number=number, fields=fields))
    return slides


def normalize_bullets(raw: str) -> List[Tuple[int, str]]:
    if not raw.strip():
        return []
    out: List[Tuple[int, str]] = []
    for line in raw.splitlines():
        if not line.strip():
            continue
        m = re.match(r"^(\s*)[-*]\s+(.*)$", line)
        if m:
            indent = len(m.group(1).replace("\t", "    "))
            level = min(indent // 2, 4)
            out.append((level, m.group(2).strip()))
        else:
            out.append((0, line.strip()))
    return out


def norm_layout_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (name or "").lower())


def find_layout(prs: Presentation, preferred_names: List[str]):
    names = {norm_layout_name(layout.name): layout for layout in prs.slide_layouts}
    for pref in preferred_names:
        key = norm_layout_name(pref)
        if key in names:
            return names[key]
    raise KeyError(f"No matching layout among: {preferred_names}")


def body_placeholders(slide):
    items = []
    for shp in slide.placeholders:
        try:
            ph = shp.placeholder_format
            if ph.type == PP_PLACEHOLDER.BODY:
                items.append(shp)
        except Exception:
            pass
    return sorted(items, key=lambda s: (s.top, s.left))


def picture_placeholders(slide) -> List[PicturePlaceholderInfo]:
    items: List[PicturePlaceholderInfo] = []
    for shp in slide.placeholders:
        try:
            if shp.placeholder_format.type == PP_PLACEHOLDER.PICTURE:
                items.append(
                    PicturePlaceholderInfo(
                        shape=shp,
                        name=(getattr(shp, "name", "") or "").strip().lower(),
                        idx=getattr(shp.placeholder_format, "idx", None),
                    )
                )
        except Exception:
            pass
    return sorted(items, key=lambda s: (s.shape.top, s.shape.left))


def choose_layout_name(spec: SlideSpec, base_dir: Optional[Path] = None) -> List[str]:
    layout = spec.get("layout").lower().strip()
    resolve_base = base_dir or Path(".")
    images = extract_images_from_spec(spec, resolve_base, resolve=base_dir is not None)
    has_img = bool(images)
    has_bullets = bool(normalize_bullets(spec.get("bullets")))
    has_subtitle = bool((spec.get("subtitle") or "").strip())
    has_title = bool((spec.get("slide title") or spec.get("title") or "").strip())
    has_comparison = any(
        (spec.get(key) or "").strip()
        for key in ("left title", "left bullets", "right title", "right bullets")
    )

    if layout in {"definition", "quote", "statement", "fact"}:
        return ["Quote", "Statement", "Big Fact", "Definition"]
    if layout in STORY_LAYOUTS:
        return ["Statement", "Quote", "Big Fact", "Title Only", "Title & Bullets"]
    if layout in COMPARISON_LAYOUTS or has_comparison:
        return [
            "Comparison",
            "Two Content",
            "Two Columns",
            "Title + Two Content",
            "Title + 2 Content",
            "Title & Bullets",
            "Bullets",
        ]
    if layout == "section" and not has_img:
        return ["Section", "Title", "Title Only"]
    if layout == "title" and not has_img:
        return ["Title", "Section", "Title Only"]

    if has_img or layout in IMAGE_LAYOUTS:
        wide_image = False
        if base_dir is not None:
            for image in images:
                if not image.path:
                    continue
                try:
                    with PILImage.open(image.path) as img:
                        width, height = img.size
                    if height and width / height > 1.15:
                        wide_image = True
                        break
                except Exception:
                    continue
        # Prefer the layout that best matches the updated template structure.
        prefs: List[str] = []
        if wide_image:
            if has_title and has_bullets:
                prefs += ["Title & Photo Alt", "Title & Photo", "Title, Bullets & Photo"]
            elif has_title and has_subtitle:
                prefs += ["Title & Photo Alt", "Title & Photo", "Title, Bullets & Photo"]
            elif has_title:
                prefs += ["Title & Photo Alt", "Title & Photo", "Photo"]
        else:
            if has_title and has_bullets:
                prefs += ["Title, Bullets & Photo", "Title & Photo", "Title & Photo Alt"]
            elif has_title and has_subtitle:
                prefs += ["Title & Photo", "Title & Photo Alt", "Title, Bullets & Photo"]
            if has_title and not has_bullets and not has_subtitle:
                prefs += ["Title & Photo", "Title & Photo Alt", "Photo"]
        prefs += [
            "Title, Bullets & Photo",
            "Title & Photo",
            "Title & Photo Alt",
            "Photo",
            "Photo - 3 Up",
            # Fallback-only candidates if the template has custom names.
            "Title + Subtitle + Bullets + Photo",
            "Bullets & Photo",
            "Comparison",
        ]
        return prefs

    return ["Title & Bullets", "Title + Subtitle + Bullets", "Bullets", "Agenda"]


def first_existing_path(base_dir: Path, raw: str) -> Optional[Path]:
    results = resolve_image_candidates(base_dir, raw)
    return results[0] if results else None


def clean_image_ref(raw: str) -> str:
    raw = raw.strip().strip('"').strip("'")
    if raw.startswith("sandbox:/mnt/data/"):
        return os.path.basename(raw)
    if raw.startswith("file://"):
        return raw[7:]
    return raw


def resolve_image_candidates(base_dir: Path, raw: str) -> List[Path]:
    found: List[Path] = []
    if not raw.strip():
        return found
    seen = set()
    for piece in re.split(r"[|,]", raw):
        ref = clean_image_ref(piece)
        if not ref:
            continue
        candidates = []
        p = Path(ref)
        if p.is_absolute():
            candidates.append(p)
        else:
            candidates.append(base_dir / ref)
            candidates.append(Path(ref))
            candidates.append(base_dir / p.name)
        for cand in candidates:
            try:
                resolved = cand.resolve()
            except Exception:
                resolved = cand
            if cand.exists() and resolved not in seen:
                found.append(cand)
                seen.add(resolved)
        if found:
            continue
        # Fallback: recursively search by basename below markdown directory.
        basename = p.name
        if basename:
            for cand in base_dir.rglob(basename):
                try:
                    resolved = cand.resolve()
                except Exception:
                    resolved = cand
                if cand.is_file() and resolved not in seen:
                    found.append(cand)
                    seen.add(resolved)
                    break
    return found


def extract_images_from_text(text: str, kind: str) -> List[ImageSpec]:
    items: List[ImageSpec] = []
    if not text.strip():
        return items
    matches = list(IMAGE_MARKDOWN_RE.finditer(text))
    if matches:
        for m in matches:
            items.append(ImageSpec(kind=kind, raw_ref=m.group(2).strip(), alt_text=m.group(1).strip()))
        return items
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        items.append(ImageSpec(kind=kind, raw_ref=line))
    return items


def extract_images_from_spec(spec: SlideSpec, base_dir: Path, resolve: bool = True) -> List[ImageSpec]:
    images: List[ImageSpec] = []
    for key in IMAGE_KEYS:
        value = spec.get(key)
        if value:
            images.extend(extract_images_from_text(value, key))
    if resolve:
        for item in images:
            candidates = resolve_image_candidates(base_dir, item.raw_ref)
            item.path = candidates[0] if candidates else None
    return images


def set_textframe_plain(tf, text: str):
    tf.clear()
    tf.word_wrap = True
    lines = [ln.rstrip() for ln in text.strip().splitlines()]
    if not lines:
        return
    p = tf.paragraphs[0]
    p.text = lines[0]
    for line in lines[1:]:
        np = tf.add_paragraph()
        np.text = line


def fill_bullets(tf, bullets: List[Tuple[int, str]], trailing_link: Optional[str] = None):
    tf.clear()
    tf.word_wrap = True
    first = True
    for level, text in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = level
        p.text = text
    if trailing_link:
        p = tf.add_paragraph() if not first else tf.paragraphs[0]
        p.text = trailing_link


def comparison_text(title: str, bullets: List[Tuple[int, str]]) -> str:
    lines: List[str] = []
    if title.strip():
        lines.append(title.strip())
    for level, text in bullets:
        lines.append(f'{"  " * level}• {text}')
    return "\n".join(lines).strip()


def add_youtube_link_to_body(tf, url: str):
    url = url.strip()
    if not url:
        return
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = url
    try:
        run.hyperlink.address = url
    except Exception:
        p.text = url


def picture_match_score(ph: PicturePlaceholderInfo, image: ImageSpec) -> int:
    score = 0
    name = ph.name
    kind = image.kind.lower()
    alt = image.alt_text.lower()

    keyword_groups = {
        "photo": ["photo", "image", "picture", "zdję", "foto"],
        "image": ["photo", "image", "picture", "zdję", "foto"],
        "picture": ["photo", "image", "picture", "zdję", "foto"],
        "chart": ["chart", "graph", "wykres"],
        "visual": ["visual", "diagram", "schem", "flow", "timeline"],
        "visualization": ["visual", "diagram", "schem", "flow", "timeline"],
        "visualisation": ["visual", "diagram", "schem", "flow", "timeline"],
    }
    for kw in keyword_groups.get(kind, []):
        if kw in name:
            score += 10
        if kw in alt:
            score += 2

    position_words = ["left", "right", "top", "bottom", "center", "środk", "lewy", "prawy"]
    for word in position_words:
        if word in alt and word in name:
            score += 5

    if ph.idx is not None:
        score += max(0, 3 - int(ph.idx))
    return score


def map_images_to_placeholders(images: List[ImageSpec], placeholders: List[PicturePlaceholderInfo]) -> List[Tuple[ImageSpec, PicturePlaceholderInfo]]:
    mappings: List[Tuple[ImageSpec, PicturePlaceholderInfo]] = []
    if not images or not placeholders:
        return mappings

    remaining_ph = placeholders[:]
    for image in images:
        best = None
        best_score = None
        for ph in remaining_ph:
            score = picture_match_score(ph, image)
            if best is None or score > best_score:
                best = ph
                best_score = score
        if best is None:
            continue
        mappings.append((image, best))
        remaining_ph.remove(best)
    return mappings


def populate_slide(slide, spec: SlideSpec, base_dir: Path, verbose: bool = False):
    title = spec.get("slide title") or spec.get("title")
    subtitle = spec.get("subtitle")
    bullets = normalize_bullets(spec.get("bullets"))
    teleprompter = spec.get("teleprompter") or spec.get("notes")
    layout = spec.get("layout").lower().strip()
    quote = spec.get("quote")
    attribution = spec.get("attribution")
    term = spec.get("term")
    definition = spec.get("definition")
    story = spec.get("story")
    left_title = spec.get("left title")
    left_bullets = normalize_bullets(spec.get("left bullets"))
    right_title = spec.get("right title")
    right_bullets = normalize_bullets(spec.get("right bullets"))
    youtube = spec.get("youtube") or spec.get("video") or spec.get("link")
    images = [img for img in extract_images_from_spec(spec, base_dir, resolve=True) if img.path]
    missing_images = [img for img in extract_images_from_spec(spec, base_dir, resolve=True) if not img.path]
    is_story = layout in STORY_LAYOUTS or bool(story.strip())
    is_comparison = layout in COMPARISON_LAYOUTS or any(
        (value or "").strip() for value in (left_title, spec.get("left bullets"), right_title, spec.get("right bullets"))
    )

    if slide.shapes.title is not None and title:
        slide.shapes.title.text = title

    bodies = body_placeholders(slide)
    pics = picture_placeholders(slide)

    if is_comparison:
        upper = bodies[0] if bodies else None
        columns = bodies[1:3] if len(bodies) >= 3 else bodies[:2]
        left_text = comparison_text(left_title, left_bullets)
        right_text = comparison_text(right_title, right_bullets)
        if upper is not None:
            if subtitle:
                set_textframe_plain(upper.text_frame, subtitle)
            elif len(bodies) >= 3:
                upper.text_frame.clear()
        if len(columns) >= 2:
            if left_text:
                set_textframe_plain(columns[0].text_frame, left_text)
            if right_text:
                set_textframe_plain(columns[1].text_frame, right_text)
        elif len(bodies) == 1:
            blocks = [block for block in (left_text, right_text) if block]
            if blocks:
                set_textframe_plain(bodies[0].text_frame, "\n\n".join(blocks))
    elif is_story:
        main_text = story or spec.get("content") or subtitle or ""
        if bodies:
            set_textframe_plain(bodies[0].text_frame, main_text)
        elif slide.shapes.title is not None and not title and main_text:
            slide.shapes.title.text = main_text
    elif layout in {"definition", "quote", "statement", "fact"} or (term and definition):
        main_text = definition or quote or spec.get("content")
        if term and definition and term.strip().casefold() != title.strip().casefold():
            main_text = f"{term}\n\n{definition}"
        if bodies:
            set_textframe_plain(bodies[0].text_frame, main_text)
        if len(bodies) > 1 and attribution:
            set_textframe_plain(bodies[1].text_frame, attribution)
        # The supplied template maps definitions to its Quote layout, which has
        # no title placeholder. Add the concept name in a separate title box.
        if layout == "definition" and title and slide.shapes.title is None:
            reference = bodies[0] if bodies else None
            left = reference.left if reference is not None else 736600
            width = reference.width if reference is not None else 11531600
            heading = slide.shapes.add_textbox(left, 1550000, width, 850000)
            heading.name = "Definition Title"
            frame = heading.text_frame
            frame.clear()
            frame.margin_left = frame.margin_right = 0
            frame.margin_top = frame.margin_bottom = 0
            frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            paragraph = frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.CENTER
            run = paragraph.add_run()
            run.text = title
            run.font.name = "Arial"
            run.font.size = Pt(32)
            run.font.bold = True
            run.font.color.rgb = RGBColor(19, 79, 120)
    else:
        if len(bodies) >= 2:
            upper, lower = bodies[0], bodies[-1]
            if subtitle:
                set_textframe_plain(upper.text_frame, subtitle)
            else:
                upper.text_frame.clear()
            if bullets:
                fill_bullets(lower.text_frame, bullets)
                if youtube:
                    add_youtube_link_to_body(lower.text_frame, youtube)
            elif subtitle and upper is not lower:
                lower.text_frame.clear()
            elif bodies and subtitle:
                set_textframe_plain(upper.text_frame, subtitle)
            elif youtube:
                add_youtube_link_to_body(upper.text_frame, youtube)
        elif len(bodies) == 1:
            content = ""
            if subtitle:
                content += subtitle
            if bullets:
                if content:
                    content += "\n"
                content += "\n".join([("  " * lvl + "• " + txt) for lvl, txt in bullets])
            if content:
                set_textframe_plain(bodies[0].text_frame, content)
            if youtube:
                add_youtube_link_to_body(bodies[0].text_frame, youtube)

    mappings = map_images_to_placeholders(images, pics)
    used_image_paths = set()
    for image, ph in mappings:
        try:
            ph.shape.insert_picture(str(image.path))
            used_image_paths.add(image.path)
        except Exception as exc:
            if verbose:
                print(f"[slide {spec.number}] Could not insert image {image.path} into placeholder '{ph.name}': {exc}")

    if verbose:
        for image in images:
            if image.path not in used_image_paths and pics:
                print(f"[slide {spec.number}] Image resolved but not mapped: {image.path}")
        if images and not pics:
            print(f"[slide {spec.number}] Layout has no picture placeholder for {len(images)} image(s).")
        for image in missing_images:
            print(f"[slide {spec.number}] Image not found: {image.raw_ref}")

    if teleprompter:
        try:
            ntf = slide.notes_slide.notes_text_frame
            ntf.clear()
            ntf.text = teleprompter
        except Exception as exc:
            if verbose:
                print(f"[slide {spec.number}] Could not write notes: {exc}")


def create_pdf_styles() -> PdfStyles:
    if getSampleStyleSheet is None or ParagraphStyle is None or colors is None:
        raise RuntimeError("PDF generation requires reportlab. Install it with: python -m pip install -r requirements.txt")
    styles = getSampleStyleSheet()
    fonts = register_pdf_fonts()
    scale = 1.5
    return PdfStyles(
        cover_kicker=ParagraphStyle(
            "CoverKicker",
            parent=styles["BodyText"],
            fontName=fonts["bold"],
            fontSize=10 * scale,
            leading=13 * scale,
            textColor=colors.HexColor("#A44914"),
            spaceAfter=4 * mm,
            tracking=0.5,
        ),
        cover_title=ParagraphStyle(
            "CoverTitle",
            parent=styles["Title"],
            fontName=fonts["bold"],
            fontSize=24 * scale,
            leading=30 * scale,
            textColor=colors.HexColor("#1C1C1C"),
            spaceAfter=4 * mm,
        ),
        cover_subtitle=ParagraphStyle(
            "CoverSubtitle",
            parent=styles["BodyText"],
            fontName=fonts["regular"],
            fontSize=12 * scale,
            leading=17 * scale,
            textColor=colors.HexColor("#4B5563"),
            spaceAfter=3 * mm,
        ),
        toc_heading=ParagraphStyle(
            "TocHeading",
            parent=styles["Heading2"],
            fontName=fonts["bold"],
            fontSize=14 * scale,
            leading=18 * scale,
            textColor=colors.HexColor("#1C1C1C"),
            spaceAfter=4 * mm,
        ),
        toc_item=ParagraphStyle(
            "TocItem",
            parent=styles["BodyText"],
            fontName=fonts["regular"],
            fontSize=10.2 * scale,
            leading=13.5 * scale,
            textColor=colors.HexColor("#374151"),
            spaceAfter=1.6 * mm,
        ),
        title=ParagraphStyle(
            "BookTitle",
            parent=styles["Title"],
            fontName=fonts["bold"],
            fontSize=15 * scale,
            leading=19 * scale,
            textColor=colors.HexColor("#1C1C1C"),
            spaceAfter=4 * mm,
        ),
        slide_heading=ParagraphStyle(
            "SlideHeading",
            parent=styles["Heading2"],
            fontName=fonts["bold"],
            fontSize=9.2 * scale,
            leading=12 * scale,
            textColor=colors.HexColor("#A44914"),
            backColor=colors.HexColor("#F3E6D8"),
            borderPadding=(3, 5, 3, 5),
            spaceBefore=0.5 * mm,
            spaceAfter=2 * mm,
        ),
        subtitle=ParagraphStyle(
            "Subtitle",
            parent=styles["BodyText"],
            fontName=fonts["bold"],
            fontSize=9.6 * scale,
            leading=12 * scale,
            textColor=colors.HexColor("#4F4F4F"),
            spaceAfter=2.6 * mm,
        ),
        body=ParagraphStyle(
            "Body",
            parent=styles["BodyText"],
            fontName=fonts["regular"],
            fontSize=9.2 * scale,
            leading=12.2 * scale,
            textColor=colors.HexColor("#202020"),
            spaceAfter=1.9 * mm,
        ),
        bullet=ParagraphStyle(
            "Bullet",
            parent=styles["BodyText"],
            fontName=fonts["regular"],
            fontSize=9.2 * scale,
            leading=11.8 * scale,
            leftIndent=6 * mm,
            firstLineIndent=-3.2 * mm,
            bulletIndent=0,
            spaceAfter=1 * mm,
        ),
        callout_label=ParagraphStyle(
            "CalloutLabel",
            parent=styles["Heading4"],
            fontName=fonts["bold"],
            fontSize=9.6 * scale,
            leading=12 * scale,
            textColor=colors.HexColor("#8A3F10"),
            backColor=colors.HexColor("#F8EADB"),
            borderPadding=(3, 5, 3, 5),
            spaceBefore=1.5 * mm,
            spaceAfter=1.8 * mm,
        ),
        callout_body=ParagraphStyle(
            "CalloutBody",
            parent=styles["BodyText"],
            fontName=fonts["regular"],
            fontSize=9.1 * scale,
            leading=12.2 * scale,
            textColor=colors.HexColor("#202020"),
            backColor=colors.HexColor("#FBF7F0"),
            borderPadding=(6, 8, 6, 8),
            borderRadius=4,
            spaceAfter=2 * mm,
        ),
        notes_label=ParagraphStyle(
            "NotesLabel",
            parent=styles["Heading4"],
            fontName=fonts["bold"],
            fontSize=9.5 * scale,
            leading=12 * scale,
            textColor=colors.HexColor("#6B7280"),
            spaceBefore=3 * mm,
            spaceAfter=1.5 * mm,
        ),
        notes=ParagraphStyle(
            "Notes",
            parent=styles["BodyText"],
            fontName=fonts["italic"],
            fontSize=9.4 * scale,
            leading=13 * scale,
            textColor=colors.HexColor("#4B5563"),
            backColor=colors.HexColor("#F6F1EA"),
            borderPadding=(5, 7, 5, 7),
            borderRadius=4,
            spaceAfter=2 * mm,
        ),
        caption=ParagraphStyle(
            "Caption",
            parent=styles["BodyText"],
            alignment=TA_CENTER,
            fontName=fonts["regular"],
            fontSize=8.5 * scale,
            leading=10.5 * scale,
            textColor=colors.HexColor("#6B7280"),
            spaceBefore=1.2 * mm,
            spaceAfter=1 * mm,
        ),
    )


def split_text_blocks(text: str) -> List[str]:
    return [block.strip() for block in re.split(r"\n\s*\n", text.strip()) if block.strip()]


def append_paragraph_blocks(story_list, text: str, style: ParagraphStyle):
    for block in split_text_blocks(text):
        story_list.append(Paragraph(escape(block).replace("\n", "<br/>"), style))


def append_bullet_flowables(story_list, bullets: List[Tuple[int, str]], style: ParagraphStyle):
    for level, text in bullets:
        bullet_style = ParagraphStyle(
            f"{style.name}-{level}",
            parent=style,
            leftIndent=style.leftIndent + (level * 5 * mm),
            firstLineIndent=style.firstLineIndent,
        )
        story_list.append(Paragraph(f"- {escape(text)}", bullet_style))


def build_pdf_cover(slides: List[SlideSpec], styles: PdfStyles):
    if not slides:
        return []
    first = slides[0]
    cover_title = first.get("slide title") or first.get("title") or Path("lecture").stem
    cover_subtitle = first.get("subtitle") or "Lecture notes generated from SlideCanon markdown"
    return [
        Spacer(1, 35 * mm),
        Paragraph("SlideCanon Course Notes", styles.cover_kicker),
        Paragraph(escape(cover_title), styles.cover_title),
        Paragraph(escape(cover_subtitle), styles.cover_subtitle),
        Spacer(1, 6 * mm),
        HRFlowable(width="100%", color=colors.HexColor("#C37A48"), thickness=1.1),
        Spacer(1, 7 * mm),
        Paragraph(
            f"Generated from structured markdown on {escape(dt.date.today().isoformat())}. "
            "This PDF keeps slide content, images, and speaker notes in a more readable A4 note format.",
            styles.body,
        ),
        Spacer(1, 10 * mm),
        PageBreak() if PageBreak is not None else Spacer(1, 0),
    ]


def build_pdf_toc(slides: List[SlideSpec], styles: PdfStyles):
    items = [Paragraph("Contents", styles.toc_heading), Spacer(1, 2 * mm)]
    for spec in slides:
        title = spec.get("slide title") or spec.get("title") or spec.get("term")
        if title.strip():
            items.append(Paragraph(f"{spec.number}. {escape(title)}", styles.toc_item))
    items.append(PageBreak() if PageBreak is not None else Spacer(1, 0))
    return items


def build_pdf_image(
    path: Path,
    max_width: float,
    max_height: float,
    layout_reduction_ratio: float = 0.0,
    bleed_ratio: float = 0.0,
):
    try:
        with PILImage.open(path) as img:
            width, height = img.size
    except Exception:
        return None
    if width <= 0 or height <= 0:
        return None
    scale = min(max_width / width, max_height / height, 1.0)
    render_width = width * scale
    render_height = height * scale
    if layout_reduction_ratio > 0 or bleed_ratio > 0:
        return OverlapImage(
            path,
            render_width,
            render_height,
            layout_reduction_ratio=layout_reduction_ratio,
            bleed_ratio=bleed_ratio,
        )
    return RLImage(str(path), width=render_width, height=render_height)


def estimate_flowables_height(flowables, width: float, height_limit: float) -> float:
    used = 0.0
    remaining = max(height_limit, 1)
    prev_space_after = 0.0
    first = True
    for flowable in flowables:
        if not first:
            space_before = float(flowable.getSpaceBefore()) if hasattr(flowable, "getSpaceBefore") else 0.0
            used += max(space_before - prev_space_after, 0.0)
            remaining = max(1, height_limit - used)
        try:
            _, flow_height = flowable.wrap(width, remaining)
        except Exception:
            _, flow_height = flowable.wrap(width, height_limit)
        used += flow_height
        prev_space_after = float(flowable.getSpaceAfter()) if hasattr(flowable, "getSpaceAfter") else 0.0
        first = False
        remaining = max(1, height_limit - used)
    if not first:
        used += max(prev_space_after, 0.0)
    return used


def render_slide_to_pdf_story(
    spec: SlideSpec,
    base_dir: Path,
    styles: PdfStyles,
    doc_width: float,
    doc_height: float,
):
    pre_image_elements = []
    post_image_elements = []
    title = spec.get("slide title") or spec.get("title") or spec.get("term")
    subtitle = spec.get("subtitle")
    bullets = normalize_bullets(spec.get("bullets"))
    teleprompter = spec.get("teleprompter") or spec.get("notes")
    quote = spec.get("quote")
    attribution = spec.get("attribution")
    definition = spec.get("definition")
    story = spec.get("story")
    content = spec.get("content")
    left_title = spec.get("left title")
    left_bullets = normalize_bullets(spec.get("left bullets"))
    right_title = spec.get("right title")
    right_bullets = normalize_bullets(spec.get("right bullets"))
    youtube = spec.get("youtube") or spec.get("video") or spec.get("link")
    layout = spec.get("layout").lower().strip()
    images = [img for img in extract_images_from_spec(spec, base_dir, resolve=True) if img.path]

    pre_image_elements.append(Paragraph(f"Slide {spec.number}", styles.slide_heading))
    if title.strip():
        pre_image_elements.append(Paragraph(escape(title), styles.title))

    if subtitle:
        pre_image_elements.append(Paragraph(escape(subtitle), styles.subtitle))

    is_story = layout in STORY_LAYOUTS or bool(story.strip())
    is_comparison = layout in COMPARISON_LAYOUTS or any(
        (value or "").strip() for value in (left_title, spec.get("left bullets"), right_title, spec.get("right bullets"))
    )

    if is_comparison:
        left_lines = []
        if left_title.strip():
            left_lines.append(f"<b>{escape(left_title.strip())}</b>")
        for level, text in left_bullets:
            prefix = "&nbsp;" * (level * 4)
            left_lines.append(f"{prefix}• {escape(text)}")
        right_lines = []
        if right_title.strip():
            right_lines.append(f"<b>{escape(right_title.strip())}</b>")
        for level, text in right_bullets:
            prefix = "&nbsp;" * (level * 4)
            right_lines.append(f"{prefix}• {escape(text)}")
        table = Table(
            [[Paragraph("<br/>".join(left_lines) or "&nbsp;", styles.body), Paragraph("<br/>".join(right_lines) or "&nbsp;", styles.body)]],
            colWidths=[doc_width / 2 - 4 * mm, doc_width / 2 - 4 * mm],
            hAlign="LEFT",
        )
        table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FBF7F0")),
                    ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#D6C6B4")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E6D9CA")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        pre_image_elements.append(table)
        pre_image_elements.append(Spacer(1, 2.4 * mm))
    elif is_story:
        pre_image_elements.append(Paragraph("Story", styles.callout_label))
        append_paragraph_blocks(pre_image_elements, story or content or "", styles.callout_body)
    elif definition or quote:
        label = "Definition" if definition else "Quote"
        pre_image_elements.append(Paragraph(label, styles.callout_label))
        if spec.get("term") and definition:
            pre_image_elements.append(Paragraph(f"<b>{escape(spec.get('term'))}</b>", styles.callout_body))
        append_paragraph_blocks(pre_image_elements, definition or quote, styles.callout_body)
        if attribution:
            pre_image_elements.append(Paragraph(f"Source: {escape(attribution)}", styles.caption))
    else:
        if bullets:
            append_bullet_flowables(pre_image_elements, bullets, styles.bullet)
        elif content:
            append_paragraph_blocks(pre_image_elements, content, styles.body)

    # Keep PDF figures in the document flow. The old overlap offsets could
    # place tall portrait figures on top of bullets in lecture handouts.
    layout_reduction_ratio = 0.0
    bleed_ratio = 0.0
    image_gap = 0
    min_image_height = 16 * mm
    trailing_reserve = 3 * mm
    post_image_gap = 0

    if youtube:
        post_image_elements.append(Paragraph(f"Video link: {escape(youtube.strip())}", styles.body))

    if teleprompter:
        post_image_elements.append(Paragraph("In-Depth Explanation", styles.notes_label))
        append_paragraph_blocks(post_image_elements, teleprompter, styles.notes)

    reserved_after_images = estimate_flowables_height(post_image_elements, doc_width, doc_height)
    elements = list(pre_image_elements)

    for idx, image in enumerate(images):
        remaining_images = len(images) - idx
        consumed = estimate_flowables_height(elements, doc_width, doc_height)
        remaining_height = doc_height - consumed - reserved_after_images - trailing_reserve
        if remaining_height <= 1 * mm:
            break
        usable_height = remaining_height
        if remaining_images > 1:
            usable_height = usable_height / remaining_images
        if usable_height < min_image_height:
            break
        flowable = build_pdf_image(
            image.path,
            max_width=doc_width * 0.99,
            max_height=usable_height,
            layout_reduction_ratio=layout_reduction_ratio,
            bleed_ratio=bleed_ratio,
        )
        if flowable is not None:
            elements.append(flowable)

    elements.extend(post_image_elements)

    total_height = estimate_flowables_height(elements, doc_width, doc_height)
    if KeepTogether is not None and total_height <= doc_height:
        return [KeepTogether(elements)]
    return elements


def draw_pdf_page(canvas, doc):
    canvas.saveState()
    page_width, page_height = A4
    fonts = register_pdf_fonts()
    canvas.setStrokeColor(colors.HexColor("#E6D9CA"))
    canvas.setLineWidth(0.6)
    canvas.line(doc.leftMargin, page_height - 12 * mm, page_width - doc.rightMargin, page_height - 12 * mm)
    canvas.line(doc.leftMargin, 12 * mm, page_width - doc.rightMargin, 12 * mm)
    canvas.setFont(fonts["regular"], 8.5)
    canvas.setFillColor(colors.HexColor("#6B7280"))
    canvas.drawString(doc.leftMargin, 7.5 * mm, "SlideCanon notes")
    canvas.drawRightString(page_width - doc.rightMargin, 7.5 * mm, f"Page {canvas.getPageNumber()}")
    canvas.restoreState()


def draw_pdf_cover_page(canvas, doc):
    canvas.saveState()
    page_width, page_height = A4
    canvas.setStrokeColor(colors.HexColor("#E6D9CA"))
    canvas.setLineWidth(0.8)
    canvas.line(doc.leftMargin, page_height - 18 * mm, page_width - doc.rightMargin, page_height - 18 * mm)
    canvas.restoreState()


def build_pdf_document(markdown: Path, output_pdf: Path):
    if SimpleDocTemplate is None or A4 is None or mm is None:
        raise RuntimeError("PDF generation requires reportlab. Install it with: python -m pip install -r requirements.txt")
    slides = parse_markdown(markdown)
    styles = create_pdf_styles()
    doc = SimpleDocTemplate(
        str(output_pdf),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
        title=markdown.stem,
        author="SlideCanon",
    )
    story = []
    base_dir = markdown.parent
    story.extend(build_pdf_cover(slides, styles))
    story.extend(build_pdf_toc(slides, styles))
    for index, spec in enumerate(slides):
        story.extend(render_slide_to_pdf_story(spec, base_dir, styles, doc.width, doc.height))
        if index != len(slides) - 1:
            if PageBreak is not None:
                story.append(PageBreak())
            else:
                story.append(Spacer(1, 4 * mm))
                story.append(HRFlowable(width="100%", color=colors.HexColor("#D5B89C"), thickness=0.8))
                story.append(Spacer(1, 5 * mm))
    doc.build(story, onFirstPage=draw_pdf_cover_page, onLaterPages=draw_pdf_page)


def build_presentation(markdown: Path, template: Path, output: Path, verbose: bool = False):
    prs = Presentation(str(template))

    while len(prs.slides):
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    slides = parse_markdown(markdown)
    base_dir = markdown.parent

    for spec in slides:
        layout = find_layout(prs, choose_layout_name(spec, base_dir))
        slide = prs.slides.add_slide(layout)
        populate_slide(slide, spec, base_dir, verbose=verbose)

    prs.save(str(output))


def scale_font_sizes_in_xml(root: ET.Element, factor: float):
    for tag in [f"{{{NS['a']}}}defRPr", f"{{{NS['a']}}}rPr", f"{{{NS['a']}}}endParaRPr"]:
        for el in root.iter(tag):
            sz = el.get("sz")
            if sz:
                try:
                    new_val = max(100, int(round(int(sz) * factor)))
                    el.set("sz", str(new_val))
                except Exception:
                    pass


def strip_local_text_overrides(root: ET.Element):
    prop_tags = [f"{{{NS['a']}}}rPr", f"{{{NS['a']}}}endParaRPr"]
    font_children = {f"{{{NS['a']}}}{name}" for name in ["latin", "ea", "cs", "sym", "solidFill", "gradFill", "pattFill", "ln", "highlight"]}
    remove_attrs = {
        "sz", "b", "i", "u", "strike", "cap", "spc", "kern", "baseline", "normalizeH",
        "noProof", "dirty", "smtClean", "lang", "altLang", "kumimoji",
    }
    for tag in prop_tags:
        for el in root.iter(tag):
            for attr in list(el.attrib):
                if attr in remove_attrs:
                    del el.attrib[attr]
            for child in list(el):
                if child.tag in font_children:
                    el.remove(child)


def remove_slide_number_placeholders(root: ET.Element):
    spTree = root.find(".//p:spTree", NS)
    if spTree is None:
        return
    for child in list(spTree):
        ph = child.find(".//p:ph", NS)
        if ph is not None and ph.get("type") == "sldNum":
            spTree.remove(child)


def postprocess_pptx(pptx_path: Path, font_scale: float = 1.0):
    tmpdir = Path(tempfile.mkdtemp(prefix="pptx_post_"))
    try:
        with zipfile.ZipFile(pptx_path, "r") as zf:
            zf.extractall(tmpdir)

        if abs(font_scale - 1.0) > 1e-6:
            for pattern in ["ppt/slideLayouts/*.xml", "ppt/slideMasters/*.xml", "ppt/theme/*.xml", "ppt/notesMasters/*.xml"]:
                for xml_path in tmpdir.glob(pattern):
                    root = ET.parse(xml_path).getroot()
                    scale_font_sizes_in_xml(root, font_scale)
                    ET.ElementTree(root).write(xml_path, encoding="utf-8", xml_declaration=True)

        for pattern in ["ppt/slides/*.xml", "ppt/notesSlides/*.xml"]:
            for xml_path in tmpdir.glob(pattern):
                root = ET.parse(xml_path).getroot()
                strip_local_text_overrides(root)
                ET.ElementTree(root).write(xml_path, encoding="utf-8", xml_declaration=True)

        for pattern in ["ppt/slideLayouts/*.xml", "ppt/slideMasters/*.xml"]:
            for xml_path in tmpdir.glob(pattern):
                root = ET.parse(xml_path).getroot()
                remove_slide_number_placeholders(root)
                ET.ElementTree(root).write(xml_path, encoding="utf-8", xml_declaration=True)

        out_tmp = pptx_path.with_suffix(".postprocessed.pptx")
        with zipfile.ZipFile(out_tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in sorted(tmpdir.rglob("*")):
                if fp.is_file():
                    zf.write(fp, fp.relative_to(tmpdir))
        shutil.move(out_tmp, pptx_path)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="Build a PPTX and companion A4 PDF from lecture markdown using a real PPTX template.")
    ap.add_argument("--markdown", required=True, help="Path to markdown file.")
    ap.add_argument("--template", required=True, help="Path to template PPTX (e.g. Operating systems002.pptx).")
    ap.add_argument("--output", required=True, help="Output PPTX path.")
    ap.add_argument("--pdf-output", help="Optional output PDF path. Defaults to the PPTX path with .pdf suffix.")
    ap.add_argument("--skip-pdf", action="store_true", help="Skip companion PDF generation.")
    ap.add_argument("--font-scale", type=float, default=1.0, help="Scale font sizes in masters/layouts/notes, e.g. 3.0 for 3x larger.")
    ap.add_argument("--verbose", action="store_true", help="Print warnings about unresolved images and layout/image placeholder mismatches.")
    args = ap.parse_args()

    markdown = Path(args.markdown)
    template = Path(args.template)
    output = Path(args.output)
    pdf_output = Path(args.pdf_output) if args.pdf_output else output.with_suffix(".pdf")
    output.parent.mkdir(parents=True, exist_ok=True)
    if not args.skip_pdf:
        pdf_output.parent.mkdir(parents=True, exist_ok=True)

    build_presentation(markdown, template, output, verbose=args.verbose)
    postprocess_pptx(output, font_scale=args.font_scale)
    print(f"Saved {output}")
    if not args.skip_pdf:
        build_pdf_document(markdown, pdf_output)
        print(f"Saved {pdf_output}")


if __name__ == "__main__":
    main()
