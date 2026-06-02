#!/usr/bin/env python3
"""
pptx_deep_repair.py  — Deep scanner & fixer for PPTX image problems
════════════════════════════════════════════════════════════════════════

Handles TWO separate failure classes:

  CLASS A  "red X / broken ref"   ← previous script fixed these
    • ppt/media/logo.png is literally missing from the ZIP
    • Happens at all levels: slides, slideLayouts, slideMasters

  CLASS B  "This picture can't be displayed"  ← this script fixes these
    • Media file IS in the ZIP but PowerPoint still won't render it
    • Root causes, all detected and fixed here:

      B1  [Content_Types].xml is missing a <Default Extension="emf"/>
          (or png, jpeg, wmf, etc.) — most common cause after splitting
      B2  Image binary is zero bytes or corrupt (wrong magic header)
      B3  Relationship uses r:link (external URL) instead of r:embed
          — the file was linked externally and the link is now dead
      B4  Relationship Target is literally "NULL" or empty string

USAGE (CLI)
───────────
    python pptx_deep_repair.py scan  broken.pptx
    python pptx_deep_repair.py scan  broken.pptx  --original original.pptx

    python pptx_deep_repair.py fix   broken.pptx  fixed.pptx
    python pptx_deep_repair.py fix   broken.pptx  fixed.pptx  --original original.pptx

    python pptx_deep_repair.py batch original.pptx  part1.pptx part2.pptx ...

USAGE (library)
───────────────
    from pptx_deep_repair import deep_scan, deep_fix

    report = deep_scan("broken.pptx", original_pptx="original.pptx")
    for line in report:
        print(line)

    deep_fix("broken.pptx", "fixed.pptx", original_pptx="original.pptx")

DEPENDENCIES: lxml  +  Python stdlib only
"""
from __future__ import annotations

import argparse
import os
import struct
import sys
import zipfile
from typing import Dict, List, Optional, Set, Tuple

from lxml import etree

# ──────────────────────────────────────────────────────────────────────────────
# Namespace constants
# ──────────────────────────────────────────────────────────────────────────────
PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS       = "http://schemas.openxmlformats.org/package/2006/content-types"

IMAGE_REL_TYPES = frozenset({
    f"{R_NS}/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
})

# ──────────────────────────────────────────────────────────────────────────────
# MIME / extension tables
# ──────────────────────────────────────────────────────────────────────────────
# Every extension that can appear in ppt/media/ → correct MIME type
MEDIA_CONTENT_TYPES: Dict[str, str] = {
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
    "mp4":  "video/mp4",
    "mov":  "video/quicktime",
    "avi":  "video/avi",
    "wmv":  "video/x-ms-wmv",
    "mp3":  "audio/mpeg",
    "wav":  "audio/wav",
    "m4a":  "audio/mp4",
}

# Magic-byte signatures → (offset, bytes)
# Used to detect files whose binary content doesn't match their extension
_MAGIC: Dict[str, List[Tuple[int, bytes]]] = {
    "png":  [(0, b"\x89PNG\r\n\x1a\n")],
    "jpg":  [(0, b"\xff\xd8\xff")],
    "jpeg": [(0, b"\xff\xd8\xff")],
    "gif":  [(0, b"GIF87a"), (0, b"GIF89a")],
    "bmp":  [(0, b"BM")],
    "tif":  [(0, b"II*\x00"), (0, b"MM\x00*")],
    "tiff": [(0, b"II*\x00"), (0, b"MM\x00*")],
    "wmf":  [(0, b"\xd7\xcd\xc6\x9a"), (0, b"\x01\x00\x09\x00")],
    "emf":  [(0, b"\x01\x00\x00\x00")],   # EMR_HEADER record type = 1
    "svg":  [(0, b"<?xml"), (0, b"<svg")],
    "webp": [(0, b"RIFF"), (8, b"WEBP")],
    "mp4":  [(4, b"ftyp")],
    "mov":  [(4, b"ftyp"), (4, b"wide"), (4, b"moov")],
}


# ──────────────────────────────────────────────────────────────────────────────
# Low-level helpers
# ──────────────────────────────────────────────────────────────────────────────

def _ext(filename: str) -> str:
    """Return lowercase extension without dot; e.g. 'logo.PNG' → 'png'."""
    return os.path.splitext(filename)[1].lstrip(".").lower()


def _resolve_target(part_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base  = part_path.rsplit("/", 1)[0]
    parts = (base + "/" + target).split("/")
    stack: List[str] = []
    for p in parts:
        if p == "..":
            if stack: stack.pop()
        elif p and p != ".":
            stack.append(p)
    return "/".join(stack)


def _parse_rels(zf: zipfile.ZipFile, rels_path: str) -> Dict[str, dict]:
    if rels_path not in zf.namelist():
        return {}
    try:
        tree = etree.fromstring(zf.read(rels_path))
    except etree.XMLSyntaxError:
        return {}
    return {
        r.get("Id", ""): {
            "type":        r.get("Type",       ""),
            "target":      r.get("Target",     ""),
            "target_mode": r.get("TargetMode", "Internal"),
        }
        for r in tree.iter(f"{{{PKG_RELS_NS}}}Relationship")
        if r.get("Id")
    }


def _all_rels_files(zf: zipfile.ZipFile) -> List[str]:
    return [n for n in zf.namelist() if "/_rels/" in n and n.endswith(".rels")]


def _build_rels_xml(rels: List[dict]) -> bytes:
    root = etree.Element(f"{{{PKG_RELS_NS}}}Relationships", nsmap={None: PKG_RELS_NS})
    for r in rels:
        el = etree.SubElement(root, f"{{{PKG_RELS_NS}}}Relationship")
        el.set("Id",     r["id"])
        el.set("Type",   r["type"])
        el.set("Target", r["target"])
        if r.get("target_mode", "Internal") != "Internal":
            el.set("TargetMode", r["target_mode"])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ──────────────────────────────────────────────────────────────────────────────
# [Content_Types].xml helpers
# ──────────────────────────────────────────────────────────────────────────────

def _parse_content_types(zf: zipfile.ZipFile) -> Optional[etree._Element]:
    if "[Content_Types].xml" not in zf.namelist():
        return None
    try:
        return etree.fromstring(zf.read("[Content_Types].xml"))
    except etree.XMLSyntaxError:
        return None


def _registered_extensions(ct_root: etree._Element) -> Set[str]:
    return {
        el.get("Extension", "").lower()
        for el in ct_root.iter(f"{{{CT_NS}}}Default")
    }


def _add_missing_defaults(ct_root: etree._Element, extensions: Set[str]) -> int:
    """Add <Default> entries for missing extensions. Returns count added."""
    existing = _registered_extensions(ct_root)
    added    = 0
    for ext in sorted(extensions):
        if ext in existing:
            continue
        mime = MEDIA_CONTENT_TYPES.get(ext, f"image/{ext}")
        el   = etree.SubElement(ct_root, f"{{{CT_NS}}}Default")
        el.set("Extension",   ext)
        el.set("ContentType", mime)
        added += 1
    return added


def _serialize_ct(ct_root: etree._Element) -> bytes:
    return etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ──────────────────────────────────────────────────────────────────────────────
# Image binary validator
# ──────────────────────────────────────────────────────────────────────────────

def _validate_magic(data: bytes, ext: str) -> Tuple[bool, str]:
    """
    Check that `data` starts with the expected magic bytes for `ext`.

    Returns (ok: bool, reason: str).
    """
    if len(data) == 0:
        return False, "zero-byte file"

    if len(data) < 16:
        return False, f"suspiciously small ({len(data)} bytes)"

    rules = _MAGIC.get(ext)
    if not rules:
        return True, "format not validated"  # unknown format — assume OK

    for offset, magic in rules:
        if data[offset : offset + len(magic)] == magic:
            return True, "ok"

    # Report actual vs expected
    head = data[:8].hex()
    expected = " or ".join(m.hex() for _, m in rules)
    return False, f"wrong header: got 0x{head}, expected {expected}"


# ──────────────────────────────────────────────────────────────────────────────
# Original media pool
# ──────────────────────────────────────────────────────────────────────────────

class MediaPool:
    """Loads every ppt/media/* file from the original PPTX."""

    def __init__(self, path: str) -> None:
        self.by_name: Dict[str, bytes] = {}
        with zipfile.ZipFile(path, "r") as zf:
            for name in zf.namelist():
                if name.startswith("ppt/media/"):
                    self.by_name[os.path.basename(name)] = zf.read(name)

    def get(self, filename: str) -> Optional[bytes]:
        return self.by_name.get(filename)


# ──────────────────────────────────────────────────────────────────────────────
# CLASS A check — missing media files (from previous script logic)
# ──────────────────────────────────────────────────────────────────────────────

def _find_class_a(zf: zipfile.ZipFile, pool: Optional[MediaPool]) -> List[dict]:
    """
    Class A: image rels that point to files missing from the ZIP.
    Returns list of issue dicts.
    """
    names  = set(zf.namelist())
    issues = []
    for rp in _all_rels_files(zf):
        parts_dir  = rp.split("/_rels/")[0]
        part_fname = os.path.basename(rp)[:-5]   # strip .rels
        part_path  = f"{parts_dir}/{part_fname}"

        for rid, rel in _parse_rels(zf, rp).items():
            if rel["type"] not in IMAGE_REL_TYPES:
                continue
            if rel.get("target_mode") == "External":
                continue   # handled in class B3

            target = rel["target"]
            if not target or target.upper() == "NULL":
                continue   # handled in class B4

            media_path = _resolve_target(part_path, target)
            if media_path in names:
                continue   # present — class A is fine

            fixable = pool is not None and pool.get(os.path.basename(media_path)) is not None
            issues.append({
                "class":      "A",
                "part":       part_path,
                "rp":         rp,
                "rid":        rid,
                "media_path": media_path,
                "filename":   os.path.basename(media_path),
                "fixable":    fixable,
            })
    return issues


# ──────────────────────────────────────────────────────────────────────────────
# CLASS B checks — present but not displayable
# ──────────────────────────────────────────────────────────────────────────────

def _find_class_b1(zf: zipfile.ZipFile) -> List[dict]:
    """
    B1: ppt/media/* files whose extension is NOT in [Content_Types].xml.
    This is the single most common cause of 'picture can't be displayed'.
    """
    ct_root = _parse_content_types(zf)
    registered = _registered_extensions(ct_root) if ct_root is not None else set()

    issues = []
    for name in zf.namelist():
        if not name.startswith("ppt/media/"):
            continue
        ext = _ext(name)
        if ext and ext not in registered:
            mime = MEDIA_CONTENT_TYPES.get(ext, f"application/{ext}")
            issues.append({
                "class":    "B1",
                "file":     name,
                "ext":      ext,
                "mime":     mime,
                "fixable":  True,
                "fix":      f"Add <Default Extension='{ext}' ContentType='{mime}'/> to [Content_Types].xml",
            })
    return issues


def _find_class_b2(zf: zipfile.ZipFile, pool: Optional[MediaPool]) -> List[dict]:
    """
    B2: ppt/media/* files that exist but have zero bytes or wrong magic header.
    """
    issues = []
    for name in zf.namelist():
        if not name.startswith("ppt/media/"):
            continue
        data    = zf.read(name)
        ext     = _ext(name)
        ok, why = _validate_magic(data, ext)
        if not ok:
            fname   = os.path.basename(name)
            fixable = pool is not None and pool.get(fname) is not None
            issues.append({
                "class":      "B2",
                "file":       name,
                "ext":        ext,
                "size":       len(data),
                "reason":     why,
                "fixable":    fixable,
                "fix":        "Replace with original's copy" if fixable else "Cannot auto-fix",
            })
    return issues


def _find_class_b3(zf: zipfile.ZipFile) -> List[dict]:
    """
    B3: Relationships using r:link (TargetMode=External) for images.
    These point to URLs/paths that no longer exist after splitting.
    """
    names  = set(zf.namelist())
    issues = []
    for rp in _all_rels_files(zf):
        for rid, rel in _parse_rels(zf, rp).items():
            if rel["type"] not in IMAGE_REL_TYPES:
                continue
            if rel.get("target_mode", "Internal") == "External":
                issues.append({
                    "class":   "B3",
                    "rp":      rp,
                    "rid":     rid,
                    "target":  rel["target"],
                    "fixable": False,
                    "fix":     "External image links cannot be auto-embedded; re-insert image manually",
                })
    return issues


def _find_class_b4(zf: zipfile.ZipFile) -> List[dict]:
    """
    B4: Relationships with Target='NULL' or Target='' (empty / placeholder).
    """
    issues = []
    for rp in _all_rels_files(zf):
        for rid, rel in _parse_rels(zf, rp).items():
            if rel["type"] not in IMAGE_REL_TYPES:
                continue
            target = rel["target"]
            if not target or target.strip().upper() == "NULL":
                issues.append({
                    "class":   "B4",
                    "rp":      rp,
                    "rid":     rid,
                    "target":  repr(target),
                    "fixable": False,
                    "fix":     "NULL/empty target in rels — image was never embedded; re-insert manually",
                })
    return issues


# ──────────────────────────────────────────────────────────────────────────────
# Public API — deep_scan
# ──────────────────────────────────────────────────────────────────────────────

def deep_scan(broken_pptx: str, original_pptx: Optional[str] = None) -> List[str]:
    """
    Comprehensive scan for all image issues.
    Returns list of human-readable report lines.
    """
    pool = MediaPool(original_pptx) if original_pptx else None

    with zipfile.ZipFile(broken_pptx, "r") as zf:
        a  = _find_class_a(zf, pool)
        b1 = _find_class_b1(zf)
        b2 = _find_class_b2(zf, pool)
        b3 = _find_class_b3(zf)
        b4 = _find_class_b4(zf)

    lines = []

    if not any([a, b1, b2, b3, b4]):
        return ["  ✓ No image issues detected."]

    total = len(a) + len(b1) + len(b2) + len(b3) + len(b4)
    lines.append(f"  Found {total} issue(s):")

    def _yesno(v): return "✓ fixable" if v else "✗ manual"

    if a:
        lines.append(f"\n  [Class A] Missing media files ({len(a)}):")
        for i in a:
            lines.append(f"    {i['part']}  rId={i['rid']}  → {i['media_path']}  [{_yesno(i['fixable'])}]")

    if b1:
        lines.append(f"\n  [Class B1] Unregistered extensions in [Content_Types].xml ({len(b1)}):")
        lines.append(f"  ← THIS IS THE MOST COMMON CAUSE OF 'picture can't be displayed'")
        for i in b1:
            lines.append(f"    {i['file']}  ext={i['ext']}  →  {i['fix']}")

    if b2:
        lines.append(f"\n  [Class B2] Corrupt / zero-byte image files ({len(b2)}):")
        for i in b2:
            lines.append(f"    {i['file']}  size={i['size']}  reason={i['reason']}  [{_yesno(i['fixable'])}]")

    if b3:
        lines.append(f"\n  [Class B3] External image links (r:link, not embedded) ({len(b3)}):")
        for i in b3:
            lines.append(f"    {i['rp']}  rId={i['rid']}  target={i['target']}")
        lines.append(f"    → Re-insert these images manually in PowerPoint")

    if b4:
        lines.append(f"\n  [Class B4] NULL / empty relationship targets ({len(b4)}):")
        for i in b4:
            lines.append(f"    {i['rp']}  rId={i['rid']}  target={i['target']}")

    auto_fixable = sum(1 for lst in [a, b1, b2] for i in lst if i["fixable"])
    manual       = total - auto_fixable
    lines.append(f"\n  Summary: {auto_fixable} auto-fixable, {manual} need manual intervention")

    return lines


# ──────────────────────────────────────────────────────────────────────────────
# Public API — deep_fix
# ──────────────────────────────────────────────────────────────────────────────

def deep_fix(
    broken_pptx:   str,
    output_pptx:   str,
    original_pptx: Optional[str] = None,
    verbose:        bool = True,
) -> str:
    """
    Apply all auto-fixable repairs and write output_pptx.

    Fixes applied (when possible):
      A  — restore missing ppt/media/* files from original
      B1 — add missing <Default> entries to [Content_Types].xml
      B2 — replace corrupt image binaries from original
      B4 — remove NULL/empty-target rels (stops the broken placeholder)

    B3 (external links) cannot be auto-fixed.
    """
    pool = MediaPool(original_pptx) if original_pptx else None

    with zipfile.ZipFile(broken_pptx, "r") as in_zf:
        names = set(in_zf.namelist())

        # ── Collect all issues ─────────────────────────────────────────────
        a_issues  = _find_class_a(in_zf, pool)
        b1_issues = _find_class_b1(in_zf)
        b2_issues = _find_class_b2(in_zf, pool)
        b4_issues = _find_class_b4(in_zf)

        # ── Build repair maps ──────────────────────────────────────────────

        # Media files to ADD (class A: missing from ZIP)
        media_to_add: Dict[str, bytes] = {}
        for i in a_issues:
            if pool and i["fixable"] and i["media_path"] not in names:
                media_to_add[i["media_path"]] = pool.get(i["filename"])

        # Media files to REPLACE (class B2: corrupt/zero-byte)
        media_to_replace: Dict[str, bytes] = {}
        for i in b2_issues:
            if pool and i["fixable"]:
                fresh = pool.get(os.path.basename(i["file"]))
                if fresh:
                    media_to_replace[i["file"]] = fresh

        # All extensions that need content-type entries (B1 + new from A)
        all_media_exts: Set[str] = set()
        for nm in names:
            if nm.startswith("ppt/media/"):
                e = _ext(nm)
                if e: all_media_exts.add(e)
        for p in media_to_add:
            e = _ext(p)
            if e: all_media_exts.add(e)

        ct_root = _parse_content_types(in_zf)
        if ct_root is None:
            ct_root = etree.Element(f"{{{CT_NS}}}Types", nsmap={None: CT_NS})
            # Add mandatory defaults
            for ext_, ct_ in [("rels","application/vnd.openxmlformats-package.relationships+xml"),
                               ("xml","application/xml")]:
                el = etree.SubElement(ct_root, f"{{{CT_NS}}}Default")
                el.set("Extension", ext_); el.set("ContentType", ct_)

        added_ct = _add_missing_defaults(ct_root, all_media_exts)
        new_ct_bytes = _serialize_ct(ct_root)

        # .rels files to REWRITE (class B4: strip NULL targets)
        b4_by_rp: Dict[str, Set[str]] = {}   # rels_path → set of bad rIds
        for i in b4_issues:
            b4_by_rp.setdefault(i["rp"], set()).add(i["rid"])

        updated_rels: Dict[str, bytes] = {}
        for rp, bad_rids in b4_by_rp.items():
            rels_list = []
            for rid, rel in _parse_rels(in_zf, rp).items():
                if rid in bad_rids:
                    continue   # drop the NULL entry
                rels_list.append({
                    "id":          rid,
                    "type":        rel["type"],
                    "target":      rel["target"],
                    "target_mode": rel.get("target_mode", "Internal"),
                })
            updated_rels[rp] = _build_rels_xml(rels_list)

        # ── Write output ZIP ───────────────────────────────────────────────
        with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as out_zf:
            for name in in_zf.namelist():
                if name == "[Content_Types].xml":
                    out_zf.writestr(name, new_ct_bytes)
                elif name in media_to_replace:
                    out_zf.writestr(name, media_to_replace[name])
                elif name in updated_rels:
                    out_zf.writestr(name, updated_rels[name])
                else:
                    out_zf.writestr(name, in_zf.read(name))

            # Add missing media (class A)
            for mp, data in media_to_add.items():
                out_zf.writestr(mp, data)

            # Write [Content_Types].xml if it was absent
            if "[Content_Types].xml" not in names:
                out_zf.writestr("[Content_Types].xml", new_ct_bytes)

    if verbose:
        parts = []
        if media_to_add:
            parts.append(f"{len(media_to_add)} missing media restored")
        if media_to_replace:
            parts.append(f"{len(media_to_replace)} corrupt files replaced")
        if added_ct:
            parts.append(f"{added_ct} content-type(s) added to [Content_Types].xml")
        if b4_issues:
            parts.append(f"{len(b4_issues)} NULL rels removed")
        if parts:
            print("✓ Fixed: " + ", ".join(parts))
            print(f"  → {output_pptx}")
        else:
            print(f"✓ No fixable issues found — wrote unchanged copy → {output_pptx}")

    return output_pptx


def batch_fix(
    broken_files:   List[str],
    original_pptx:  Optional[str] = None,
    suffix:         str = "_fixed",
    verbose:        bool = True,
) -> List[str]:
    pool_loaded = original_pptx is not None
    if verbose and pool_loaded:
        pool = MediaPool(original_pptx)
        print(f"Media pool: {len(pool.by_name)} file(s) from {original_pptx}")

    outputs = []
    for path in broken_files:
        stem, ext = os.path.splitext(path)
        out = f"{stem}{suffix}{ext}"
        deep_fix(path, out, original_pptx=original_pptx, verbose=verbose)
        outputs.append(out)
    return outputs


# ──────────────────────────────────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────────────────────────────────

def _main() -> None:
    p = argparse.ArgumentParser(
        prog="pptx_deep_repair",
        description="Deep scanner + fixer for PPTX image issues",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python pptx_deep_repair.py scan  broken.pptx
  python pptx_deep_repair.py scan  broken.pptx --original original.pptx
  python pptx_deep_repair.py fix   broken.pptx  fixed.pptx
  python pptx_deep_repair.py fix   broken.pptx  fixed.pptx --original original.pptx
  python pptx_deep_repair.py batch part1.pptx part2.pptx --original original.pptx
""",
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    sp = sub.add_parser("scan",  help="Diagnose all image issues")
    sp.add_argument("broken",                help="PPTX to scan")
    sp.add_argument("--original", default=None, help="Original source PPTX (optional)")

    fp = sub.add_parser("fix",   help="Repair all auto-fixable issues")
    fp.add_argument("broken",               help="PPTX to repair")
    fp.add_argument("output",               help="Output path")
    fp.add_argument("--original", default=None, help="Original source PPTX (optional)")

    bp = sub.add_parser("batch", help="Repair multiple files")
    bp.add_argument("files",    nargs="+",  help="PPTX files to repair")
    bp.add_argument("--original", default=None, help="Original source PPTX (optional)")

    args = p.parse_args()

    if args.cmd == "scan":
        for line in deep_scan(args.broken, original_pptx=args.original):
            print(line)
    elif args.cmd == "fix":
        deep_fix(args.broken, args.output, original_pptx=args.original)
    elif args.cmd == "batch":
        batch_fix(args.files, original_pptx=args.original)


if __name__ == "__main__":
    _main()
