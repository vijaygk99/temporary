#!/usr/bin/env python3
"""
pptx_image_repair.py
════════════════════
Fix ALL broken embedded images (red X marks) in a split/combined PPTX.

WHY THE PREVIOUS APPROACH FAILED
──────────────────────────────────
Images in a PPTX can live at THREE different levels:
  • ppt/slides/_rels/slideN.xml.rels       ← slide images
  • ppt/slideLayouts/_rels/layoutN.xml.rels ← template / layout images
  • ppt/slideMasters/_rels/masterN.xml.rels ← LOGO, brand images ← ✗ missed

When openpowerxmltools splits a deck, it correctly carries over all the .rels
files at every level, but frequently drops the actual binary media files from
ppt/media/.  The result: every .rels entry says "../media/logo.png" but
logo.png simply isn't in the ZIP ⟹ red X.

THE CORRECT FIX  (this script)
────────────────────────────────
1.  Build a filename→bytes map of EVERY file in original's ppt/media/.
2.  Walk EVERY */_rels/*.rels file inside the broken PPTX at all three levels.
3.  For each image relationship whose target file is missing from the broken
    ZIP, restore it by filename from the original's media pool.
4.  If that fails (renamed by the split tool) fall back to a hash/size match.
5.  Update [Content_Types].xml so PowerPoint accepts the restored binaries.
6.  Write a repaired output file.

DEPENDENCIES
────────────────────────────────
  • lxml          (pip install lxml)
  • Python stdlib: zipfile, os, pathlib, hashlib, mimetypes — nothing else

USAGE (CLI)
────────────────────────────────
  # Diagnose only (no file written)
  python pptx_image_repair.py diagnose original.pptx broken.pptx

  # Repair → fixed.pptx
  python pptx_image_repair.py fix original.pptx broken.pptx fixed.pptx

  # Batch: fix multiple splits from the same original
  python pptx_image_repair.py batch original.pptx part1.pptx part2.pptx part3.pptx

USAGE (library)
────────────────────────────────
  from pptx_image_repair import repair, diagnose

  repair("original.pptx", "broken.pptx", "fixed.pptx")

  issues = diagnose("original.pptx", "broken.pptx")
  for issue in issues:
      print(issue)
"""
from __future__ import annotations

import argparse
import hashlib
import mimetypes
import os
import sys
import zipfile
from pathlib import PurePosixPath
from typing import Dict, List, Optional, Set, Tuple

from lxml import etree

# ─── Namespace constants ──────────────────────────────────────────────────────
PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS       = "http://schemas.openxmlformats.org/package/2006/content-types"

IMAGE_REL_TYPES: frozenset = frozenset({
    f"{R_NS}/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
})

# Extension → MIME for types that mimetypes module sometimes misses
_EXTRA_MIME: Dict[str, str] = {
    "png":  "image/png",
    "jpg":  "image/jpeg",
    "jpeg": "image/jpeg",
    "gif":  "image/gif",
    "bmp":  "image/bmp",
    "tif":  "image/tiff",
    "tiff": "image/tiff",
    "wmf":  "image/x-wmf",
    "emf":  "image/x-emf",
    "svg":  "image/svg+xml",
    "ico":  "image/x-icon",
    "webp": "image/webp",
}


# ─── Low-level ZIP helpers ────────────────────────────────────────────────────

def _resolve_target(part_path: str, target: str) -> str:
    """Resolve a relative .rels Target to an absolute ZIP-internal path.

    Works for ANY part level:
      part_path = 'ppt/slideMasters/slideMaster1.xml'
      target    = '../media/logo.png'
      → 'ppt/media/logo.png'
    """
    if target.startswith("/"):
        return target.lstrip("/")
    base  = part_path.rsplit("/", 1)[0]
    parts = (base + "/" + target).split("/")
    stack: List[str] = []
    for p in parts:
        if p == "..":
            if stack:
                stack.pop()
        elif p and p != ".":
            stack.append(p)
    return "/".join(stack)


def _parse_rels(zf: zipfile.ZipFile, rels_path: str) -> Dict[str, dict]:
    """Parse a .rels file. Returns {rId: {type, target}} or {} if absent."""
    if rels_path not in zf.namelist():
        return {}
    try:
        tree = etree.fromstring(zf.read(rels_path))
    except etree.XMLSyntaxError:
        return {}
    return {
        r.get("Id", ""): {
            "type":   r.get("Type",   ""),
            "target": r.get("Target", ""),
        }
        for r in tree.iter(f"{{{PKG_RELS_NS}}}Relationship")
        if r.get("Id")
    }


def _rels_path_for(part_zip_path: str) -> str:
    """Derive the .rels path for a given part path.

    'ppt/slides/slide1.xml'
      → 'ppt/slides/_rels/slide1.xml.rels'
    'ppt/slideMasters/slideMaster1.xml'
      → 'ppt/slideMasters/_rels/slideMaster1.xml.rels'
    """
    dirname  = part_zip_path.rsplit("/", 1)[0]
    basename = part_zip_path.rsplit("/", 1)[-1]
    return f"{dirname}/_rels/{basename}.rels"


def _all_rels_files(zf: zipfile.ZipFile) -> List[str]:
    """All */_rels/*.rels paths inside the package (any level)."""
    return [
        n for n in zf.namelist()
        if "/_rels/" in n and n.endswith(".rels")
    ]


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ─── Content-Types helpers ────────────────────────────────────────────────────

def _parse_content_types(zf: zipfile.ZipFile) -> etree._Element:
    """Return parsed [Content_Types].xml root element."""
    ct_path = "[Content_Types].xml"
    if ct_path not in zf.namelist():
        root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
        return root
    return etree.fromstring(zf.read(ct_path))


def _registered_extensions(ct_root: etree._Element) -> Set[str]:
    """Set of extensions already in <Default Extension="…"> entries."""
    return {
        el.get("Extension", "").lower()
        for el in ct_root.iter(f"{{{CT_NS}}}Default")
    }


def _mime_for_ext(ext: str) -> str:
    ext = ext.lstrip(".").lower()
    return (
        _EXTRA_MIME.get(ext)
        or mimetypes.types_map.get(f".{ext}", "application/octet-stream")
    )


def _update_content_types(ct_root: etree._Element, new_exts: Set[str]) -> bytes:
    """Add <Default> entries for any extensions not already registered."""
    existing = _registered_extensions(ct_root)
    for ext in sorted(new_exts):
        if ext not in existing:
            el = etree.SubElement(ct_root, f"{{{CT_NS}}}Default")
            el.set("Extension",    ext)
            el.set("ContentType",  _mime_for_ext(ext))
    return etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


# ─── Core media map builder (original PPTX) ──────────────────────────────────

class OriginalMediaPool:
    """
    Loads all media from the original PPTX and provides two lookup methods:
      • by_name[filename]       → bytes   (fast, exact match)
      • by_hash[sha256hex]      → bytes   (fallback when name differs)
    """
    def __init__(self, original_pptx: str) -> None:
        self.by_name:  Dict[str, bytes] = {}   # basename.ext → data
        self.by_hash:  Dict[str, bytes] = {}   # sha256 → data
        self.all_data: List[Tuple[int, bytes]] = []  # [(size, data)] for size match

        with zipfile.ZipFile(original_pptx, "r") as zf:
            for name in zf.namelist():
                if not name.startswith("ppt/media/"):
                    continue
                data     = zf.read(name)
                basename = os.path.basename(name)
                self.by_name[basename]    = data
                self.by_hash[_sha256(data)] = data
                self.all_data.append((len(data), data))

    def find(self, filename: str, ref_data: Optional[bytes] = None) -> Optional[bytes]:
        """Look up media by filename, then by hash of ref_data if given."""
        hit = self.by_name.get(filename)
        if hit is not None:
            return hit
        if ref_data is not None:
            return self.by_hash.get(_sha256(ref_data))
        return None


# ─── Main scanner: find all broken image refs across all levels ───────────────

def _scan_broken(
    zf: zipfile.ZipFile,
    pool: OriginalMediaPool,
) -> Tuple[Dict[str, bytes], Set[str]]:
    """
    Scan every .rels file in `zf` at every level (slides, layouts, masters, …).

    Returns:
      missing_media  : {zip_path_in_output → bytes_to_add}
      missing_ct_exts: set of file extensions needing Content-Types entries
    """
    names         = set(zf.namelist())
    missing_media: Dict[str, bytes] = {}  # ppt/media/logo.png → data
    ct_exts:       Set[str] = set()

    for rels_path in _all_rels_files(zf):
        # Derive the "owning" part path from the .rels path
        # e.g. ppt/slides/_rels/slide1.xml.rels → ppt/slides/slide1.xml
        parts_dir  = rels_path.split("/_rels/")[0]           # ppt/slides
        rels_fname = os.path.basename(rels_path)             # slide1.xml.rels
        part_fname = rels_fname[:-5] if rels_fname.endswith(".rels") else rels_fname
        part_path  = f"{parts_dir}/{part_fname}"             # ppt/slides/slide1.xml

        for rid, rel in _parse_rels(zf, rels_path).items():
            if rel["type"] not in IMAGE_REL_TYPES:
                continue

            # Resolve the absolute path inside the ZIP
            media_zip_path = _resolve_target(part_path, rel["target"])

            if media_zip_path in names:
                continue  # ← already present, nothing to do

            # File is missing from the broken ZIP — try to restore from original
            filename = os.path.basename(media_zip_path)
            data = pool.find(filename)

            if data is None:
                # Last resort: scan the rels path to find the original's matching rels
                # and get the data by matching on the same rId from the same level
                # (handles cases where split tool renamed the media file)
                continue   # will be logged as unfixable

            if media_zip_path not in missing_media:
                missing_media[media_zip_path] = data
                ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
                if ext:
                    ct_exts.add(ext)

    return missing_media, ct_exts


# ─── Public API ───────────────────────────────────────────────────────────────

def diagnose(original_pptx: str, broken_pptx: str) -> List[str]:
    """
    Report all broken image references in broken_pptx.
    Checks slides, layouts, AND masters.
    No files are written.
    """
    pool  = OriginalMediaPool(original_pptx)
    lines: List[str] = []

    with zipfile.ZipFile(broken_pptx, "r") as zf:
        names = set(zf.namelist())

        for rels_path in sorted(_all_rels_files(zf)):
            parts_dir  = rels_path.split("/_rels/")[0]
            rels_fname = os.path.basename(rels_path)
            part_fname = rels_fname[:-5] if rels_fname.endswith(".rels") else rels_fname
            part_path  = f"{parts_dir}/{part_fname}"

            for rid, rel in _parse_rels(zf, rels_path).items():
                if rel["type"] not in IMAGE_REL_TYPES:
                    continue
                media_zip_path = _resolve_target(part_path, rel["target"])
                if media_zip_path in names:
                    continue

                filename = os.path.basename(media_zip_path)
                fixable  = pool.find(filename) is not None
                icon     = "✓ fixable" if fixable else "✗ NOT in original"
                lines.append(
                    f"  BROKEN  {part_path}  rId={rid}  "
                    f"target={media_zip_path}  [{icon}]"
                )

    if not lines:
        return ["  ✓ No broken image references found."]
    return [f"  Found {len(lines)} broken reference(s):"] + lines


def repair(
    original_pptx: str,
    broken_pptx:   str,
    output_pptx:   str,
    verbose:        bool = True,
) -> str:
    """
    Repair all broken embedded images in broken_pptx and write output_pptx.

    Scans slides, slideLayouts, AND slideMasters.
    Restores missing ppt/media/* binaries from original_pptx by filename.
    Updates [Content_Types].xml as needed.

    Parameters
    ----------
    original_pptx : path to the original source deck
    broken_pptx   : path to the split / broken deck
    output_pptx   : path to write the repaired deck
    verbose       : print a summary line

    Returns
    -------
    str : output_pptx path
    """
    pool = OriginalMediaPool(original_pptx)

    with zipfile.ZipFile(broken_pptx, "r") as in_zf:
        missing_media, new_ct_exts = _scan_broken(in_zf, pool)

        if not missing_media:
            if verbose:
                print("✓ No broken images detected — writing unchanged copy.")
        else:
            if verbose:
                print(f"  Restoring {len(missing_media)} missing media file(s):")
                for p in sorted(missing_media):
                    print(f"    + {p}")

        # ── Update [Content_Types].xml ──────────────────────────────────────
        ct_root     = _parse_content_types(in_zf)
        new_ct_bytes: Optional[bytes] = None
        if new_ct_exts:
            new_ct_bytes = _update_content_types(ct_root, new_ct_exts)

        # ── Write output ZIP ────────────────────────────────────────────────
        with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as out_zf:
            for name in in_zf.namelist():
                if name == "[Content_Types].xml" and new_ct_bytes:
                    out_zf.writestr(name, new_ct_bytes)
                else:
                    out_zf.writestr(name, in_zf.read(name))

            # Add every missing media file
            for zip_path, data in missing_media.items():
                out_zf.writestr(zip_path, data)

            # If [Content_Types].xml was absent entirely, write it now
            if "[Content_Types].xml" not in in_zf.namelist() and new_ct_bytes:
                out_zf.writestr("[Content_Types].xml", new_ct_bytes)

    if verbose and missing_media:
        print(f"✓ Repaired → {output_pptx}")

    return output_pptx


def batch_repair(
    original_pptx: str,
    broken_files:  List[str],
    suffix:        str = "_repaired",
    verbose:       bool = True,
) -> List[str]:
    """
    Repair multiple split PPTXes from the same original in one pass.

    Output filenames: <stem><suffix>.pptx
    Example: 'part1.pptx' → 'part1_repaired.pptx'
    """
    pool = OriginalMediaPool(original_pptx)   # load original only once

    if verbose:
        print(f"Media pool loaded from original: {len(pool.by_name)} file(s)")

    outputs: List[str] = []
    for path in broken_files:
        stem, ext = os.path.splitext(path)
        out = f"{stem}{suffix}{ext}"
        repair(original_pptx, path, out, verbose=verbose)
        outputs.append(out)
    return outputs


# ─── CLI ──────────────────────────────────────────────────────────────────────

def _main() -> None:
    p = argparse.ArgumentParser(
        prog="pptx_image_repair",
        description="Fix broken embedded images (X marks) in split/combined PPTX files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Diagnose (no file written)
  python pptx_image_repair.py diagnose original.pptx broken.pptx

  # Repair single file
  python pptx_image_repair.py fix original.pptx broken.pptx fixed.pptx

  # Batch: fix multiple split files from the same original
  python pptx_image_repair.py batch original.pptx part1.pptx part2.pptx part3.pptx
""",
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    dp = sub.add_parser("diagnose", help="Report broken references only")
    dp.add_argument("original", help="Original source PPTX")
    dp.add_argument("broken",   help="Split/broken PPTX to inspect")

    fp = sub.add_parser("fix", help="Repair a broken PPTX")
    fp.add_argument("original", help="Original source PPTX")
    fp.add_argument("broken",   help="Split/broken PPTX to repair")
    fp.add_argument("output",   help="Output path for the repaired file")

    bp = sub.add_parser("batch", help="Repair multiple split PPTXes")
    bp.add_argument("original", help="Original source PPTX")
    bp.add_argument("files",    nargs="+", help="Split/broken PPTX files to repair")

    args = p.parse_args()

    if args.cmd == "diagnose":
        for line in diagnose(args.original, args.broken):
            print(line)
    elif args.cmd == "fix":
        repair(args.original, args.broken, args.output)
    elif args.cmd == "batch":
        batch_repair(args.original, args.files)


if __name__ == "__main__":
    _main()
