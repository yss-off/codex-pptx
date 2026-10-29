#!/usr/bin/env python3
"""Check a local PPTX edit against explicit preservation invariants.

Examples:
  compare_decks.py before.pptx after.pptx --mode palette --slides 4,7,11
  compare_decks.py before.pptx after.pptx --mode layout --slides 9
  compare_decks.py before.pptx after.pptx --mode reorder

Slide numbers are 1-based presentation order, not slideN.xml suffixes. Reorder
mode requires stable slide-part identities and checks content/notes by identity.
This is an edit-scope check, not an OOXML, visual, or scientific validator.
"""
from __future__ import annotations

import argparse
import json
import posixpath
from pathlib import Path
from urllib.parse import unquote, urlsplit
import zipfile

from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"p": P, "a": A, "r": R}
COLOR_TAGS = {f"{{{A}}}{name}" for name in (
    "srgbClr", "scrgbClr", "hslClr", "sysClr", "schemeClr", "prstClr",
)}
METADATA = {"docProps/core.xml", "docProps/app.xml"}


def xml(data: bytes):
    return etree.fromstring(data, etree.XMLParser(
        remove_blank_text=True, resolve_entities=False, no_network=True,
    ))


def normalized(data: bytes, name: str, *, ignore_colors=False) -> bytes:
    if not name.endswith((".xml", ".rels")):
        return data
    root = xml(data)
    if ignore_colors:
        for node in list(root.iter()):
            if node.tag in COLOR_TAGS:
                node.tag = "color-value"
                node.attrib.clear()
                node.text = None
                for child in list(node):
                    node.remove(child)
    return etree.tostring(root, method="c14n")


def text_content(root):
    """Preserve text objects, paragraphs and explicit breaks, not run splitting."""
    bodies = root.xpath('//p:txBody | //a:txBody', namespaces=NS)
    return [
        [''.join(''.join(node.itertext()) if node.tag == f'{{{A}}}t'
               else '\n' if node.tag == f'{{{A}}}br' else '\t'
               for node in paragraph.iter()
               if node.tag in {f'{{{A}}}t', f'{{{A}}}br', f'{{{A}}}tab'})
         for paragraph in body.findall('a:p', NS)]
        for body in bodies
    ]


def run_links(root):
    """Preserve linked character spans while allowing equivalent run splitting."""
    result = []
    for body in root.xpath('//p:txBody | //a:txBody', namespaces=NS):
        paragraphs = []
        for paragraph in body.findall('a:p', NS):
            offset, spans = 0, []
            for run in paragraph:
                text = ''.join(run.xpath('.//a:t/text()', namespaces=NS))
                if run.tag == f'{{{A}}}br':
                    text = '\n'
                links = tuple(etree.tostring(link, method='c14n') for link in
                              run.xpath('./a:rPr/a:hlinkClick | ./a:rPr/a:hlinkMouseOver', namespaces=NS))
                end = offset + len(text)
                if links and end > offset:
                    if spans and spans[-1][1] == offset and spans[-1][2] == links:
                        spans[-1] = (spans[-1][0], end, links)
                    else:
                        spans.append((offset, end, links))
                offset = end
            paragraphs.append(spans)
        result.append(paragraphs)
    return result


def layout_content(data):
    root = xml(data)
    # Keep relationship uses tied to object identity, not just the set of rIds.
    refs = []
    for node in root.iter():
        is_link = node.tag in {f'{{{A}}}hlinkClick', f'{{{A}}}hlinkMouseOver'}
        if is_link and any(ancestor.tag in {f'{{{A}}}r', f'{{{A}}}fld'}
                           for ancestor in node.iterancestors()):
            continue  # Covered by character spans, independent of run boundaries.
        attrs = tuple(sorted((key, val) for key, val in node.attrib.items()
                             if is_link or key.startswith(f'{{{R}}}')))
        if not attrs and not is_link:
            continue
        owner = next((ancestor for ancestor in node.iterancestors()
                      if ancestor.tag in {f'{{{P}}}{name}' for name in
                                          ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp')}), root)
        identity = owner.find('.//p:cNvPr', NS)
        refs.append((identity.get('id') if identity is not None else None,
                     node.tag, attrs))
    controls = [etree.tostring(node, method='c14n') for node in root
                if node.tag in {f'{{{P}}}transition', f'{{{P}}}timing', f'{{{P}}}extLst'}]
    return text_content(root), run_links(root), refs, controls, tuple(sorted(root.attrib.items()))


class Deck:
    def __init__(self, path: str | Path):
        with zipfile.ZipFile(path) as archive:
            names = [n for n in archive.namelist() if not n.endswith("/")]
            if len(set(names)) != len(names):
                raise ValueError("Duplicate package part names")
            self.parts = {n: archive.read(n) for n in names}
        roots = self.relationships("")
        documents = [target for _, kind, target in roots if kind.endswith("/officeDocument")]
        if len(documents) != 1:
            raise ValueError("Expected one officeDocument relationship")
        self.presentation = documents[0]
        root = xml(self.parts[self.presentation])
        rels = {rid: (kind, target) for rid, kind, target in self.relationships(self.presentation)}
        self.order = []
        for node in root.findall("p:sldIdLst/p:sldId", NS):
            kind, target = rels[node.get(f"{{{R}}}id")]
            if not kind.endswith("/slide"):
                raise ValueError("Slide-list entry does not refer to a slide")
            self.order.append(target)
        if len(set(self.order)) != len(self.order):
            raise ValueError("Duplicate slide-part identity in presentation order")
        size = root.find("p:sldSz", NS)
        if size is None:
            raise ValueError("Missing presentation slide size")
        self.size = tuple(sorted(size.attrib.items()))
        self.slides = {part: self.snapshot(part) for part in self.order}

    @staticmethod
    def relpath(part: str) -> str:
        return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels") if part else "_rels/.rels"

    def relationships(self, part: str):
        name = self.relpath(part)
        if name not in self.parts:
            return []
        result = []
        for rel in xml(self.parts[name]):
            if rel.get("TargetMode") == "External":
                continue
            uri = urlsplit(rel.get("Target", ""))
            if uri.scheme or uri.netloc:
                raise ValueError(f"Invalid internal relationship in {name}")
            target = unquote(uri.path)
            target = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(part), target))
            if target.startswith("../") or target == ".." or target not in self.parts:
                raise ValueError(f"Missing or escaping relationship target: {target}")
            result.append((rel.get("Id"), rel.get("Type", ""), target))
        return result

    def dependencies(self, part: str) -> set[str]:
        seen = set()
        def walk(owner):
            if owner in seen:
                return
            seen.add(owner)
            relpath = self.relpath(owner)
            if relpath in self.parts:
                seen.add(relpath)
            for _, kind, target in self.relationships(owner):
                # Back-pointers / master layout catalogs are not sibling slide content.
                if kind.endswith("/slide") or ("/slideMasters/" in owner and kind.endswith("/slideLayout")):
                    continue
                walk(target)
        walk(part)
        return seen

    def global_dependencies(self):
        """Presentation defaults and global styles cannot be scoped to one slide.

        Master catalogs are intentionally excluded: slide-local traversal already
        covers used masters, and traversing the catalog would pull in siblings.
        """
        result = {self.presentation, self.relpath(self.presentation)}
        for _, kind, target in self.relationships(self.presentation):
            if kind.rsplit('/', 1)[-1] not in {
                'slide', 'slideMaster', 'notesMaster', 'handoutMaster',
            }:
                result.update(self.dependencies(target))
        return result

    def snapshot(self, part):
        root = xml(self.parts[part])
        note_parts = [target for _, kind, target in self.relationships(part) if kind.endswith("/notesSlide")]
        if len(note_parts) > 1:
            raise ValueError(f"Multiple notes slides for {part}")
        notes = None
        if note_parts:
            note_root = xml(self.parts[note_parts[0]])
            bodies = note_root.xpath('//p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]', namespaces=NS)
            notes = [etree.tostring(body.find("p:txBody", NS), method="c14n").decode() for body in bodies]
        return {
            "text": text_content(root),
            "notes_part": note_parts[0] if note_parts else None,
            "notes": notes,
            "geometry": [etree.tostring(node, method="c14n").decode() for node in root.xpath('//a:xfrm | //p:xfrm', namespaces=NS)],
            "hidden": root.get("show", "1"),
        }


def compare(before: str | Path, after: str | Path, mode="palette", slides=None) -> dict:
    if mode not in {"palette", "layout", "reorder"}:
        raise ValueError(f"Unsupported mode: {mode}")
    a, b = Deck(before), Deck(after)
    selected = set(range(1, len(a.order) + 1)) if slides is None else set(slides)
    if not selected <= set(range(1, len(a.order) + 1)) or not selected:
        raise ValueError("--slides must contain valid 1-based presentation positions")
    if mode == "reorder" and slides is not None:
        raise ValueError("--slides is not supported in reorder mode")
    failures = []
    if a.size != b.size:
        failures.append("Slide size changed")
    if set(a.order) != set(b.order):
        failures.append("Slide-part identities changed; preservation cannot be established")
    if mode != "reorder" and a.order != b.order:
        failures.append("Presentation order changed")
    changes = []
    for part in sorted(a.parts.keys() | b.parts.keys()):
        old, new = a.parts.get(part), b.parts.get(part)
        if old is None or new is None or normalized(old, part) != normalized(new, part):
            changes.append(part)
    metadata_changes = sorted(set(changes) & METADATA)
    meaningful = set(changes) - METADATA
    if mode != "reorder" and len(selected) < len(a.order):
        global_changes = meaningful & (a.global_dependencies() | b.global_dependencies())
        if global_changes:
            failures.append("Global rendering changes cannot be limited to selected slides: "
                            + ', '.join(sorted(global_changes)))
    shared = set(a.order) & set(b.order)
    for position, part in enumerate(a.order, 1):
        if part not in shared:
            continue
        old, new = a.slides[part], b.slides[part]
        for facet in ("text", "notes_part", "notes", "hidden"):
            if old[facet] != new[facet]:
                failures.append(f"Slide {position} ({part}): {facet} changed")
        if mode in {"palette", "reorder"} and old["geometry"] != new["geometry"]:
            failures.append(f"Slide {position} ({part}): geometry changed")
        touched = meaningful & (a.dependencies(part) | b.dependencies(part))
        if mode != "reorder" and position not in selected and touched:
            failures.append(f"Unselected slide {position} affected by: {', '.join(sorted(touched))}")
    if mode == "palette":
        for part in sorted(meaningful):
            if part not in a.parts or part not in b.parts or normalized(a.parts[part], part, ignore_colors=True) != normalized(b.parts[part], part, ignore_colors=True):
                failures.append(f"Non-color package change: {part}")
    elif mode == "reorder":
        for part in sorted(meaningful):
            if part == a.presentation == b.presentation:
                ar, br = xml(a.parts[part]), xml(b.parts[part])
                for root in (ar, br):
                    node = root.find("p:sldIdLst", NS)
                    if node is not None:
                        # Sort complete entries: only their order may differ.
                        entries = list(node)
                        for entry in entries:
                            node.remove(entry)
                        for entry in sorted(entries, key=lambda e: etree.tostring(e, method="c14n")):
                            node.append(entry)
                if etree.tostring(ar, method="c14n") == etree.tostring(br, method="c14n"):
                    continue
            failures.append(f"Change beyond presentation reordering: {part}")
    else:
        for part in sorted(meaningful):
            if any(segment in part for segment in ("/media/", "/embeddings/", "/charts/")):
                failures.append(f"Embedded content changed outside layout-only scope: {part}")
            if part not in a.parts or part not in b.parts:
                failures.append(f"Added or removed part outside layout scope: {part}")
                continue
            if part.endswith('.rels'):
                failures.append(f"Relationship change outside layout scope: {part}")
                continue
            if part.endswith('.xml'):
                ar, br = xml(a.parts[part]), xml(b.parts[part])
                allowed = {f'{{{P}}}{name}' for name in ('sld', 'sldLayout', 'sldMaster')}
                if ar.tag not in allowed or br.tag != ar.tag:
                    failures.append(f"Non-layout part changed: {part}")
                elif layout_content(a.parts[part]) != layout_content(b.parts[part]):
                    failures.append(f"Content, references or playback settings changed: {part}")
            else:
                failures.append(f"Non-layout part changed: {part}")
        reachable = set().union(*(a.dependencies(p) for p in a.order), *(b.dependencies(p) for p in b.order))
        for part in sorted(meaningful - reachable):
            failures.append(f"Non-slide package change outside layout scope: {part}")
    return {
        "passed": not failures, "mode": mode, "failures": failures,
        "before_order": a.order, "after_order": b.order,
        "selected_slides": sorted(selected), "changed_parts": changes,
        "metadata_changes_not_checked": metadata_changes,
        "limits": [
            "No schema, rendering, visual quality, scientific accuracy, or application compatibility validation.",
            "Stable slide-part identities are required; equivalent regenerated decks may be conservatively rejected.",
            "Layout mode allows edits within existing slide/layout/master XML only; text, content references, root settings, playback controls, notes and relationships are preserved. Layout switches and dependency rewrites require separate review; graphical meaning and typography are not validated.",
            "Shared rendering dependencies can mark multiple slides affected; global style changes are conservatively rejected for partial slide scopes. Palette mode compares existing color slots only; adding an inherited-color override may be conservatively rejected. Core/app document metadata changes are reported but not rejected.",
        ],
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before")
    parser.add_argument("after")
    parser.add_argument("--mode", choices=("palette", "layout", "reorder"), required=True)
    parser.add_argument("--slides", help="Comma-separated 1-based presentation positions allowed to change")
    args = parser.parse_args()
    try:
        pages = [int(x) for x in args.slides.split(",")] if args.slides else None
        report = compare(args.before, args.after, args.mode, pages)
    except (ValueError, KeyError, OSError, zipfile.BadZipFile, etree.XMLSyntaxError) as exc:
        report = {"passed": False, "error": str(exc)}
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
