#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
for k, v in NS.items():
    ET.register_namespace(k, v)


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


def process_pptx(inp: Path, out: Path):
    tmpdir = Path(tempfile.mkdtemp(prefix="strip_overrides_"))
    try:
        with zipfile.ZipFile(inp, "r") as zf:
            zf.extractall(tmpdir)
        for pattern in ["ppt/slides/*.xml", "ppt/notesSlides/*.xml"]:
            for xml_path in tmpdir.glob(pattern):
                root = ET.parse(xml_path).getroot()
                strip_local_text_overrides(root)
                ET.ElementTree(root).write(xml_path, encoding="utf-8", xml_declaration=True)
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
            for fp in sorted(tmpdir.rglob("*")):
                if fp.is_file():
                    zf.write(fp, fp.relative_to(tmpdir))
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser(description="Remove local text/font overrides from a PPTX so text inherits from template/theme.")
    ap.add_argument("input", help="Input PPTX")
    ap.add_argument("output", help="Output PPTX")
    args = ap.parse_args()
    process_pptx(Path(args.input), Path(args.output))
    print(f"Saved {args.output}")


if __name__ == "__main__":
    main()
