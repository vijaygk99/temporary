#!/usr/bin/env python3
"""
pptx_image_fix.py
═════════════════
Fix broken embedded images (✕ / X marks) in split or combined PPTX files.
Uses ONLY lxml + Python's built-in zipfile – no python-pptx needed.

────────────────────────────────────────────────────────────────────────────
WHY IMAGES BREAK AFTER SPLITTING
────────────────────────────────────────────────────────────────────────────
A .pptx is a ZIP archive with this structure:

  ppt/slides/slide3.xml              ← slide content
  ppt/slides/_rels/slide3.xml.rels   ← maps rId → media file
  ppt/media/company_logo.png         ← actual image binary

In a slide's XML an image is referenced like:
    <a:blip r:embed="rId2"/>

The .rels file resolves rId2:
    <Relationship Id="rId2" Type=".../image"
                  Target="../media/company_logo.png"/>

openpowerxmltools (and similar) commonly causes one of three failures:
  ① .rels file copied, but ppt/media/ binaries NOT copied into new ZIP
  ② .rels file dropped entirely
  ③ rId renumbered in .rels but the slide XML still has the original rId

All three produce the red X mark. This script auto-detects and repairs all.

────────────────────────────────────────────────────────────────────────────
QUICK START
────────────────────────────────────────────────────────────────────────────
As a script:
    # Diagnose (nothing modified)
    python pptx_image_fix.py diagnose original.pptx broken_split.pptx

    # Fix (slide N in broken = slide N in original – default)
    python pptx_image_fix.py fix original.pptx broken.pptx fixed.pptx

    # Fix with explicit mapping (broken slide 1 = original slide 4, etc.)
    python pptx_image_fix.py fix original.pptx broken.pptx fixed.pptx \
        --map '{"1":4,"2":5,"3":6}'

As a library:
    from pptx_image_fix import fix_split_pptx, batch_fix_splits, diagnose

    # Single file
    fix_split_pptx("original.pptx", "broken.pptx", "fixed.pptx")

    # With explicit slide map
    fix_split_pptx("original.pptx", "broken.pptx", "fixed.pptx",
                   slide_map={1: 4, 2: 5, 3: 6})

    # Batch: fix three split files from same original
    batch_fix_splits(
        "original.pptx",
        ["part1.pptx", "part2.pptx", "part3.pptx"],
        slide_maps={
            "part1.pptx": {1:1, 2:2, 3:3},
            "part2.pptx": {1:4, 2:5, 3:6},
            "part3.pptx": {1:7, 2:8, 3:9, 4:10},
        }
    )
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import zipfile
from typing import Dict, List, Optional, Set

from lxml import etree


# ═══════════════════════════════════════════════════════════════════════════════
# Namespace constants
# ═══════════════════════════════════════════════════════════════════════════════

PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Both the standard and OOXML-alt image relationship types
IMAGE_REL_TYPES: frozenset = frozenset({
    f"{R_NS}/image",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/image",
})


# ═══════════════════════════════════════════════════════════════════════════════
# Low-level ZIP / XML helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _slide_num(path: str) -> int:
    """Extract the integer from a slide filename.
    'ppt/slides/slide3.xml'  →  3
    """
    return int("".join(c for c in os.path.basename(path) if c.isdigit()))


def _sorted_slides(namelist) -> List[str]:
    """Return slide XML paths in ascending slide-number order."""
    return sorted(
        (n for n in namelist
         if n.startswith("ppt/slides/slide") and n.endswith(".xml")),
        key=_slide_num,
    )


def _rels_path(slide_path: str) -> str:
    """Map a slide XML path to its companion .rels path.
    'ppt/slides/slide3.xml'  →  'ppt/slides/_rels/slide3.xml.rels'
    """
    return f"ppt/slides/_rels/{os.path.basename(slide_path)}.rels"


def _resolve(slide_path: str, target: str) -> str:
    """Resolve a relative .rels Target to an absolute ZIP-internal path.

    Example:
        slide_path = 'ppt/slides/slide1.xml'
        target     = '../media/logo.png'
        →            'ppt/media/logo.png'
    """
    if target.startswith("/"):
        return target.lstrip("/")
    base  = slide_path.rsplit("/", 1)[0]   # 'ppt/slides'
    parts = (base + "/" + target).split("/")
    out: List[str] = []
    for p in parts:
        if p == "..":
            if out:
                out.pop()
        elif p and p != ".":
            out.append(p)
    return "/".join(out)


def _parse_rels(zf: zipfile.ZipFile, rp: str) -> Dict[str, dict]:
    """Parse a .rels file from an open ZipFile.

    Returns:
        {rId: {type, target, target_mode}}
        Empty dict if the .rels file does not exist in the archive.
    """
    if rp not in zf.namelist():
        return {}
    tree = etree.fromstring(zf.read(rp))
    return {
        r.get("Id"): {
            "type":        r.get("Type",       ""),
            "target":      r.get("Target",     ""),
            "target_mode": r.get("TargetMode", "Internal"),
        }
        for r in tree.iter(f"{{{PKG_RELS_NS}}}Relationship")
    }


def _embedded_rids(zf: zipfile.ZipFile, slide_path: str) -> Set[str]:
    """Return all rId values referenced via r:embed or r:link in a slide's XML."""
    tree = etree.fromstring(zf.read(slide_path))
    rids: Set[str] = set()
    for el in tree.iter():
        for attr in (f"{{{R_NS}}}embed", f"{{{R_NS}}}link"):
            v = el.get(attr)
            if v:
                rids.add(v)
    return rids


def _build_rels_xml(rels: List[dict]) -> bytes:
    """Serialize a list of relationship dicts to .rels XML bytes.

    Each dict must have keys: id, type, target, and optionally target_mode.
    """
    root = etree.Element(
        f"{{{PKG_RELS_NS}}}Relationships",
        nsmap={None: PKG_RELS_NS},
    )
    for r in rels:
        el = etree.SubElement(root, f"{{{PKG_RELS_NS}}}Relationship")
        el.set("Id",     r["id"])
        el.set("Type",   r["type"])
        el.set("Target", r["target"])
        if r.get("target_mode", "Internal") != "Internal":
            el.set("TargetMode", r["target_mode"])
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


# ═══════════════════════════════════════════════════════════════════════════════
# Image-map builder
# ═══════════════════════════════════════════════════════════════════════════════

def _build_image_map(zf: zipfile.ZipFile) -> Dict[int, Dict[str, dict]]:
    """Build a complete image reference map for every slide in a PPTX.

    Returns:
        {
          slide_num: {
            rId: {
              type, target, media_path,   ← strings
              media_data,                  ← raw bytes of the image binary
              filename,                    ← basename of media_path
            }
          }
        }
    Only IMAGE relationship types are included; non-image rels are skipped.
    Media files that do not exist inside `zf` are also skipped.
    """
    names  = set(zf.namelist())
    result: Dict[int, Dict[str, dict]] = {}

    for slide_path in _sorted_slides(names):
        snum = _slide_num(slide_path)
        rp   = _rels_path(slide_path)

        img_rels: Dict[str, dict] = {}
        for rid, rel in _parse_rels(zf, rp).items():
            if rel["type"] not in IMAGE_REL_TYPES:
                continue                           # skip non-image rels
            mp = _resolve(slide_path, rel["target"])
            if mp not in names:
                continue                           # skip unreachable targets
            img_rels[rid] = {
                "type":       rel["type"],
                "target":     rel["target"],
                "media_path": mp,
                "media_data": zf.read(mp),
                "filename":   os.path.basename(mp),
            }
        result[snum] = img_rels

    return result


# ═══════════════════════════════════════════════════════════════════════════════
# Public API: diagnose
# ═══════════════════════════════════════════════════════════════════════════════

def diagnose(
    broken_pptx:   str,
    original_pptx: str,
    slide_map:     Optional[Dict[int, int]] = None,
) -> List[str]:
    """Report broken image references in a split PPTX (read-only).

    Parameters
    ----------
    broken_pptx   : Path to the split PPTX with broken images.
    original_pptx : Path to the original source PPTX.
    slide_map     : {broken_slide_num → original_slide_num}.
                    None = identity mapping (slide N maps to slide N).

    Returns
    -------
    List of human-readable diagnostic lines.
    No files are created or modified.
    """
    lines: List[str] = []

    with zipfile.ZipFile(original_pptx, "r") as orig_zf:
        orig_map = _build_image_map(orig_zf)

        with zipfile.ZipFile(broken_pptx, "r") as bk_zf:
            bk_names  = set(bk_zf.namelist())
            bk_slides = _sorted_slides(bk_names)

            if slide_map is None:
                slide_map = {_slide_num(s): _slide_num(s) for s in bk_slides}

            for sp in bk_slides:
                bnum      = _slide_num(sp)
                onum      = slide_map.get(bnum, bnum)
                cur_rels  = _parse_rels(bk_zf, _rels_path(sp))
                refs      = _embedded_rids(bk_zf, sp)
                orig_rels = orig_map.get(onum, {})

                for rid in refs:
                    if rid not in orig_rels:
                        continue  # not an image rId in the original

                    # Check if .rels entry exists AND media file is present
                    if rid in cur_rels and cur_rels[rid]["type"] in IMAGE_REL_TYPES:
                        mp = _resolve(sp, cur_rels[rid]["target"])
                        if mp in bk_names:
                            continue  # ← healthy
                        lines.append(
                            f"  Slide {bnum}  rId={rid}: "
                            f".rels entry exists but media file '{mp}' is "
                            f"MISSING from the ZIP  "
                            f"[cause ①]"
                        )
                    else:
                        lines.append(
                            f"  Slide {bnum}  rId={rid}: "
                            f"relationship entry is MISSING entirely  "
                            f"(original media: {orig_rels[rid]['filename']})  "
                            f"[cause ② or ③]"
                        )

    if not lines:
        return ["  ✓ No broken image references detected."]

    return [f"  ✗ {len(lines)} broken image reference(s):"] + lines


# ═══════════════════════════════════════════════════════════════════════════════
# Public API: fix_split_pptx
# ═══════════════════════════════════════════════════════════════════════════════

def fix_split_pptx(
    original_pptx: str,
    broken_pptx:   str,
    output_pptx:   str,
    slide_map:     Optional[Dict[int, int]] = None,
    verbose:       bool = True,
) -> str:
    """Repair broken embedded images in a split PPTX.

    The function repairs all three known failure modes:
      ① Media files missing from ZIP  → copies them from original
      ② .rels file missing entirely   → creates it from original's rels
      ③ rId in .rels / slide mismatch → re-links the correct media

    Parameters
    ----------
    original_pptx : str
        The original (unbroken) PPTX file the broken one was split from.
    broken_pptx : str
        The split PPTX showing X-mark images.
    output_pptx : str
        Output path for the repaired file (may be the same as broken_pptx
        to overwrite in-place, but a new path is safer).
    slide_map : dict, optional
        {broken_slide_num → original_slide_num}.
        Omit for identity mapping (slide 1→1, 2→2, …).
        Use when the split file's slide numbering differs from the original.
    verbose : bool
        Print a summary line to stdout.

    Returns
    -------
    str : Path of the repaired output file.

    Examples
    --------
    # Simplest case: slide numbers match
    fix_split_pptx("deck.pptx", "split_broken.pptx", "split_fixed.pptx")

    # Slides 4-6 of original became slides 1-3 of the split file
    fix_split_pptx("deck.pptx", "split_broken.pptx", "split_fixed.pptx",
                   slide_map={1: 4, 2: 5, 3: 6})
    """

    with zipfile.ZipFile(original_pptx, "r") as orig_zf:
        orig_img_map = _build_image_map(orig_zf)

        with zipfile.ZipFile(broken_pptx, "r") as bk_zf:
            bk_names  = set(bk_zf.namelist())
            bk_slides = _sorted_slides(bk_names)

            if slide_map is None:
                slide_map = {_slide_num(s): _slide_num(s) for s in bk_slides}

            # ── Phase 1: Detect which rIds need repair ────────────────────
            # repair_plan[broken_slide_num] = {rId: orig_rel_dict}
            repair_plan: Dict[int, Dict[str, dict]] = {}

            for sp in bk_slides:
                bnum      = _slide_num(sp)
                onum      = slide_map.get(bnum, bnum)
                orig_rels = orig_img_map.get(onum, {})
                cur_rels  = _parse_rels(bk_zf, _rels_path(sp))
                refs      = _embedded_rids(bk_zf, sp)

                to_fix: Dict[str, dict] = {}
                for rid in refs:
                    if rid not in orig_rels:
                        continue   # rId is not an image, or original doesn't know it

                    # Is the current entry already healthy?
                    if rid in cur_rels and cur_rels[rid]["type"] in IMAGE_REL_TYPES:
                        mp = _resolve(sp, cur_rels[rid]["target"])
                        if mp in bk_names:
                            continue  # ← already working

                    to_fix[rid] = orig_rels[rid]

                if to_fix:
                    repair_plan[bnum] = to_fix

            # Nothing broken → fast-path copy
            if not repair_plan:
                if verbose:
                    print("✓ No broken images found – copying file unchanged.")
                shutil.copy2(broken_pptx, output_pptx)
                return output_pptx

            # ── Phase 2: Assign collision-free filenames for incoming media ─
            # Multiple slides may reference the same original image (e.g. logo).
            # We only add it to ppt/media/ once and reuse the same filename.
            existing_media: Set[str] = {
                os.path.basename(n) for n in bk_names
                if n.startswith("ppt/media/")
            }
            # orig_zip_path → filename that will be used in the output archive
            orig_path_to_fname: Dict[str, str] = {}

            for to_fix in repair_plan.values():
                for rel in to_fix.values():
                    opath = rel["media_path"]
                    if opath in orig_path_to_fname:
                        continue  # already assigned a name

                    fname = rel["filename"]

                    # Avoid overwriting an existing (different) file with the same name
                    if fname in existing_media:
                        stem, ext = os.path.splitext(fname)
                        counter = 1
                        while f"{stem}_{counter}{ext}" in existing_media:
                            counter += 1
                        fname = f"{stem}_{counter}{ext}"

                    orig_path_to_fname[opath] = fname
                    existing_media.add(fname)

            # ── Phase 3: Build updated .rels XML for each affected slide ──
            # rels_path (str) → new XML bytes
            updated_rels: Dict[str, bytes] = {}

            for bnum, to_fix in repair_plan.items():
                sp  = next(s for s in bk_slides if _slide_num(s) == bnum)
                rp  = _rels_path(sp)
                cur = _parse_rels(bk_zf, rp)

                # Start from current rels, keeping every non-broken entry as-is
                rels_list: List[dict] = [
                    {
                        "id":          rid,
                        "type":        r["type"],
                        "target":      r["target"],
                        "target_mode": r.get("target_mode", "Internal"),
                    }
                    for rid, r in cur.items()
                    if rid not in to_fix           # exclude the ones we're fixing
                ]

                # Append correct entries for every broken rId
                for rid, orig_rel in to_fix.items():
                    fname = orig_path_to_fname[orig_rel["media_path"]]
                    rels_list.append({
                        "id":          rid,
                        "type":        orig_rel["type"],
                        "target":      f"../media/{fname}",  # standard relative path
                        "target_mode": "Internal",
                    })

                updated_rels[rp] = _build_rels_xml(rels_list)

            # ── Phase 4: Collect media binaries that must be added ─────────
            # new_zip_path → raw bytes; only files not already in the broken ZIP
            new_media: Dict[str, bytes] = {}
            for to_fix in repair_plan.values():
                for rel in to_fix.values():
                    fname    = orig_path_to_fname[rel["media_path"]]
                    new_path = f"ppt/media/{fname}"
                    if new_path not in bk_names and new_path not in new_media:
                        new_media[new_path] = rel["media_data"]

            # ── Phase 5: Write the output ZIP ─────────────────────────────
            with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as out_zf:
                # 5a. Copy every file from broken PPTX, substituting updated .rels
                for name in bk_zf.namelist():
                    data = (
                        updated_rels[name]
                        if name in updated_rels
                        else bk_zf.read(name)
                    )
                    out_zf.writestr(name, data)

                # 5b. Write .rels files that were absent in the broken ZIP
                for rp, content in updated_rels.items():
                    if rp not in bk_names:
                        out_zf.writestr(rp, content)

                # 5c. Write all new media binaries
                for mpath, mdata in new_media.items():
                    out_zf.writestr(mpath, mdata)

    if verbose:
        total_rids   = sum(len(v) for v in repair_plan.values())
        slides_fixed = len(repair_plan)
        print(
            f"✓ Repaired {total_rids} image reference(s) "
            f"across {slides_fixed} slide(s)  →  {output_pptx}"
        )
    return output_pptx


# ═══════════════════════════════════════════════════════════════════════════════
# Public API: batch_fix_splits
# ═══════════════════════════════════════════════════════════════════════════════

def batch_fix_splits(
    original_pptx: str,
    broken_files:  List[str],
    slide_maps:    Optional[Dict[str, Dict[int, int]]] = None,
    suffix:        str = "_fixed",
) -> List[str]:
    """Repair multiple split PPTXes that all came from the same original.

    Parameters
    ----------
    original_pptx : str
        Common source PPTX.
    broken_files : list of str
        Paths to the split PPTXes that need repair.
    slide_maps : dict, optional
        {broken_file_path: {broken_slide_num → original_slide_num}}.
        Omit for identity mapping on every file.
    suffix : str
        String appended to the stem of each input filename for the output.
        Default: "_fixed"  (e.g. "part1.pptx" → "part1_fixed.pptx")

    Returns
    -------
    List[str] : Paths of all repaired output files.

    Example
    -------
    # original.pptx was split into three sequential parts
    batch_fix_splits(
        "original.pptx",
        ["slides_1_to_3.pptx", "slides_4_to_6.pptx", "slides_7_to_10.pptx"],
        slide_maps={
            "slides_1_to_3.pptx":  {1: 1, 2: 2, 3: 3},
            "slides_4_to_6.pptx":  {1: 4, 2: 5, 3: 6},
            "slides_7_to_10.pptx": {1: 7, 2: 8, 3: 9, 4: 10},
        },
    )
    """
    outputs: List[str] = []
    for path in broken_files:
        stem, ext = os.path.splitext(path)
        out       = f"{stem}{suffix}{ext}"
        sm        = (slide_maps or {}).get(path)
        fix_split_pptx(original_pptx, path, out, slide_map=sm)
        outputs.append(out)
    return outputs


# ═══════════════════════════════════════════════════════════════════════════════
# Convenience: fix a single file in-place (makes a .bak first)
# ═══════════════════════════════════════════════════════════════════════════════

def fix_inplace(
    original_pptx: str,
    target_pptx:   str,
    slide_map:     Optional[Dict[int, int]] = None,
    make_backup:   bool = True,
) -> str:
    """Repair `target_pptx` in-place, optionally backing up the original.

    Parameters
    ----------
    original_pptx : str  Source PPTX.
    target_pptx   : str  The split PPTX to repair (overwritten on success).
    slide_map     : dict  Optional {broken_slide_num → original_slide_num}.
    make_backup   : bool  If True, renames target to target.bak before fixing.

    Returns
    -------
    str : Path of the repaired file (same as target_pptx).
    """
    if make_backup:
        bak = target_pptx + ".bak"
        shutil.copy2(target_pptx, bak)
        print(f"  Backup saved: {bak}")

    import tempfile
    tmp = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False)
    tmp.close()
    try:
        fix_split_pptx(original_pptx, target_pptx, tmp.name,
                       slide_map=slide_map, verbose=False)
        shutil.move(tmp.name, target_pptx)
    except Exception:
        os.unlink(tmp.name)
        raise

    print(f"✓ Repaired in-place: {target_pptx}")
    return target_pptx


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

def _main() -> None:
    parser = argparse.ArgumentParser(
        prog="pptx_image_fix",
        description="Fix broken embedded images (X marks) in split PPTX files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Diagnose only — no file written
  python pptx_image_fix.py diagnose original.pptx broken_split.pptx

  # Auto slide mapping (slide N in broken = slide N in original)
  python pptx_image_fix.py fix original.pptx broken.pptx fixed.pptx

  # Manual mapping (slides 4,5,6 of original became 1,2,3 of the split)
  python pptx_image_fix.py fix original.pptx broken.pptx fixed.pptx \\
      --map '{"1":4,"2":5,"3":6}'

  # Fix in-place (creates .bak automatically)
  python pptx_image_fix.py fix original.pptx broken.pptx broken.pptx
""",
    )

    sub = parser.add_subparsers(dest="cmd", required=True)

    # ── diagnose ──────────────────────────────────────────────────────────────
    dp = sub.add_parser(
        "diagnose",
        help="Report broken references without modifying files",
    )
    dp.add_argument("original", help="Original source PPTX")
    dp.add_argument("broken",   help="Split / broken PPTX to inspect")
    dp.add_argument(
        "--map", default=None,
        help="JSON slide-number map  e.g. '{\"1\":4,\"2\":5}'",
    )

    # ── fix ───────────────────────────────────────────────────────────────────
    fp = sub.add_parser(
        "fix",
        help="Repair a split PPTX's broken image references",
    )
    fp.add_argument("original", help="Original source PPTX")
    fp.add_argument("broken",   help="Split / broken PPTX to repair")
    fp.add_argument("output",   help="Output path for the repaired file")
    fp.add_argument(
        "--map", default=None,
        help="JSON slide-number map  e.g. '{\"1\":4,\"2\":5}'",
    )

    args = parser.parse_args()

    sm: Optional[Dict[int, int]] = None
    if getattr(args, "map", None):
        try:
            sm = {int(k): int(v) for k, v in json.loads(args.map).items()}
        except (json.JSONDecodeError, ValueError) as exc:
            print(f"ERROR: --map is not valid JSON: {exc}", file=sys.stderr)
            sys.exit(1)

    if args.cmd == "diagnose":
        for line in diagnose(args.broken, args.original, slide_map=sm):
            print(line)

    elif args.cmd == "fix":
        fix_split_pptx(args.original, args.broken, args.output, slide_map=sm)


if __name__ == "__main__":
    _main()
