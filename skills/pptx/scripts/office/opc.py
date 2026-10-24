"""Resolve OPC relationship targets inside an unpacked Office package."""

from pathlib import Path
from urllib.parse import unquote, urlsplit

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


def resolve_target(root: Path, rels_file: Path, target: str, mode: str = "") -> Path | None:
    """Return an internal target; external URLs and escaping paths are not parts."""
    uri = urlsplit(target)
    if not target or mode == "External" or uri.scheme or uri.netloc:
        return None
    path = unquote(uri.path)
    root = root.resolve()
    if path.startswith("/"):
        resolved = (root / path.lstrip("/")).resolve()
    else:
        base = root if rels_file.name == ".rels" else rels_file.parent.parent
        resolved = (base / path).resolve()
    try:
        resolved.relative_to(root)
    except ValueError:
        return None
    return resolved
