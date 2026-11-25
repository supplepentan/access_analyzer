"""Recursive descent parser for Microsoft Access SaveAsText files."""

from __future__ import annotations

from dataclasses import dataclass, field
import re


@dataclass
class SaveAsTextNode:
    """A Begin <node_type> ... End block inside a SaveAsText dump."""

    node_type: str
    properties: dict[str, str] = field(default_factory=dict)
    children: list[SaveAsTextNode] = field(default_factory=list)


@dataclass
class SaveAsTextDocument:
    """Parsed representation of a SaveAsText file."""

    headers: dict[str, str] = field(default_factory=dict)
    root_nodes: list[SaveAsTextNode] = field(default_factory=list)
    code_behind: str = ""

    @property
    def header_properties(self) -> dict[str, str]:
        return self.headers


_HEX_LINE_RE = re.compile(r"^0x[0-9a-fA-F]+$")
_ASSIGN_RE = re.compile(r"^([A-Za-z0-9_.\-]+)\s*=\s*(.*)$")


def _unescape_access_string(raw: str) -> str:
    """Unquote and unescape a single Access SaveAsText string literal."""
    s = raw.strip()
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        s = s[1:-1]
    s = s.replace('""', '"')
    s = s.replace("\\015\\012", "\n").replace("\\012", "\n").replace("\\015", "")
    s = s.replace("\\r\\n", "\n").replace("\\n", "\n")
    return s


def _consume_multiline_quoted(first_val: str, lines: list[str], next_idx: int) -> tuple[str, int]:
    """Consume consecutive indented quoted string lines following `Key ="..."`."""
    parts = [_unescape_access_string(first_val)]
    idx = next_idx
    while idx < len(lines):
        stripped = lines[idx].strip()
        if stripped.startswith('"') and stripped.endswith('"') and len(stripped) >= 2:
            parts.append(_unescape_access_string(stripped))
            idx += 1
        else:
            break
    return "".join(parts), idx


def _parse_value_block(lines: list[str], start_idx: int) -> tuple[str, int]:
    """Parse a `Key = Begin ... End` value block (multi-line string or hex blob)."""
    idx = start_idx
    chunks: list[str] = []
    hex_lines = 0
    total_lines = 0

    while idx < len(lines):
        stripped = lines[idx].strip()
        idx += 1
        if stripped == "End":
            break
        if not stripped:
            continue
        total_lines += 1
        if _HEX_LINE_RE.match(stripped):
            hex_lines += 1
        elif stripped.startswith('"') and stripped.endswith('"'):
            chunks.append(_unescape_access_string(stripped))
        else:
            chunks.append(stripped)

    if total_lines > 0 and hex_lines == total_lines:
        return f"<binary:{hex_lines}_lines>", idx
    return "".join(chunks), idx


def _parse_block(lines: list[str], start_idx: int, node_type: str) -> tuple[SaveAsTextNode, int]:
    """Recursively parse `Begin <node_type> ... End`."""
    node = SaveAsTextNode(node_type=node_type)
    idx = start_idx
    n = len(lines)

    while idx < n:
        stripped = lines[idx].strip()
        if not stripped:
            idx += 1
            continue
        if stripped == "End":
            return node, idx + 1
        if stripped in ("CodeBehindForm", "CodeBehindReport"):
            return node, idx

        if stripped == "Begin" or stripped.startswith("Begin "):
            child_type = stripped[6:].strip() if stripped.startswith("Begin ") else "Anonymous"
            child_node, idx = _parse_block(lines, idx + 1, child_type)
            node.children.append(child_node)
            continue

        match = _ASSIGN_RE.match(stripped)
        if match:
            key, val = match.group(1), match.group(2).strip()
            if val == "Begin":
                parsed_val, idx = _parse_value_block(lines, idx + 1)
                node.properties[key] = parsed_val
            elif val.startswith('"') and val.endswith('"') and len(val) >= 2:
                parsed_val, idx = _consume_multiline_quoted(val, lines, idx + 1)
                node.properties[key] = parsed_val
            else:
                node.properties[key] = val
                idx += 1
            continue

        idx += 1

    return node, idx


def parse_saveastext(text: str) -> SaveAsTextDocument:
    """Parse a full SaveAsText document into headers, AST nodes, and optional CodeBehindForm."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    doc = SaveAsTextDocument()

    idx = 0
    n = len(lines)
    while idx < n:
        stripped = lines[idx].strip()
        if not stripped:
            idx += 1
            continue
        if stripped in ("CodeBehindForm", "CodeBehindReport"):
            doc.code_behind = "\n".join(lines[idx + 1 :]).strip("\n")
            break
        if stripped == "Begin" or stripped.startswith("Begin "):
            node_type = stripped[6:].strip() if stripped.startswith("Begin ") else "Anonymous"
            node, idx = _parse_block(lines, idx + 1, node_type)
            doc.root_nodes.append(node)
            continue

        match = _ASSIGN_RE.match(stripped)
        if match:
            key, val = match.group(1), match.group(2).strip()
            if val == "Begin":
                parsed_val, idx = _parse_value_block(lines, idx + 1)
                doc.headers[key] = parsed_val
            elif val.startswith('"') and val.endswith('"') and len(val) >= 2:
                parsed_val, idx = _consume_multiline_quoted(val, lines, idx + 1)
                doc.headers[key] = parsed_val
            else:
                doc.headers[key] = val
                idx += 1
            continue

        idx += 1

    return doc
