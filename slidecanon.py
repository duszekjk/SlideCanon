#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

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


def choose_layout_name(spec: SlideSpec) -> List[str]:
    layout = spec.get("layout").lower().strip()
    images = extract_images_from_spec(spec, Path("."), resolve=False)
    has_img = bool(images)
    has_bullets = bool(normalize_bullets(spec.get("bullets")))
    has_subtitle = bool((spec.get("subtitle") or "").strip())
    has_title = bool((spec.get("slide title") or spec.get("title") or "").strip())

    if layout in {"definition", "quote", "statement", "fact"}:
        return ["Quote", "Statement", "Big Fact", "Definition"]
    if layout in {"title", "section"} and not has_img:
        return ["Title", "Section", "Title Only"]

    if has_img or layout in {"photo", "image", "visual", "visualisation", "visualization", "chart"}:
        # Prefer real picture-placeholder layouts first.
        # In Operating systems002.pptx, 'Title & Photo' has a real picture placeholder,
        # while 'Title, Bullets & Photo' does not.
        prefs: List[str] = []
        if has_title and (has_subtitle or has_bullets):
            prefs += ["Title & Photo", "Title & Photo Alt"]
        if has_title and not has_bullets and not has_subtitle:
            prefs += ["Title & Photo", "Title & Photo Alt", "Photo"]
        prefs += [
            "Title & Photo",
            "Title & Photo Alt",
            "Photo",
            "Photo - 3 Up",
            # Fallback-only candidates if the template has custom names.
            "Title + Subtitle + Bullets + Photo",
            "Bullets & Photo",
            "Title, Bullets & Photo",
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
    youtube = spec.get("youtube") or spec.get("video") or spec.get("link")
    images = [img for img in extract_images_from_spec(spec, base_dir, resolve=True) if img.path]
    missing_images = [img for img in extract_images_from_spec(spec, base_dir, resolve=True) if not img.path]

    if slide.shapes.title is not None and title:
        slide.shapes.title.text = title

    bodies = body_placeholders(slide)
    pics = picture_placeholders(slide)

    if layout in {"definition", "quote", "statement", "fact"} or (term and definition):
        main_text = definition or quote or spec.get("content")
        if term and definition:
            main_text = f"{term}\n\n{definition}"
        if bodies:
            set_textframe_plain(bodies[0].text_frame, main_text)
        if len(bodies) > 1 and attribution:
            set_textframe_plain(bodies[1].text_frame, attribution)
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


def build_presentation(markdown: Path, template: Path, output: Path, verbose: bool = False):
    prs = Presentation(str(template))

    while len(prs.slides):
        rId = prs.slides._sldIdLst[0].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    slides = parse_markdown(markdown)
    base_dir = markdown.parent

    for spec in slides:
        layout = find_layout(prs, choose_layout_name(spec))
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
    ap = argparse.ArgumentParser(description="Build a PPTX from lecture markdown using a real PPTX template and strip local overrides.")
    ap.add_argument("--markdown", required=True, help="Path to markdown file.")
    ap.add_argument("--template", required=True, help="Path to template PPTX (e.g. Operating systems002.pptx).")
    ap.add_argument("--output", required=True, help="Output PPTX path.")
    ap.add_argument("--font-scale", type=float, default=1.0, help="Scale font sizes in masters/layouts/notes, e.g. 3.0 for 3x larger.")
    ap.add_argument("--verbose", action="store_true", help="Print warnings about unresolved images and layout/image placeholder mismatches.")
    args = ap.parse_args()

    markdown = Path(args.markdown)
    template = Path(args.template)
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)

    build_presentation(markdown, template, output, verbose=args.verbose)
    postprocess_pptx(output, font_scale=args.font_scale)
    print(f"Saved {output}")


if __name__ == "__main__":
    main()
