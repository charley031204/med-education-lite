"""Shared representation for a native page draft and its review packet."""

from dataclasses import dataclass, field
import re


def literal(text):
    """Escape Markdown controls while preserving the displayed characters."""
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return re.sub(r"([\\`*_$~#\[\]])", r"\\\1", text)


def prose_line(text):
    """Keep a source line from becoming an unintended Markdown block."""
    text = re.sub(r"^(\s{0,3})([+-])(?=\s|$)", r"\1\\\2", text)
    text = re.sub(r"^(\s{0,3}\d{1,9})([.)])(?=\s|$)", r"\1\\\2", text)
    text = re.sub(r"^(\s{0,3})([=-])(?=[=\s-]*$)", r"\1\\\2", text)
    if re.fullmatch(r"[ \t|:-]+", text):
        return text.replace("|", r"\|")
    return text


def style_text(text, bold=False, italic=False, script=""):
    text = literal(text)
    core = text.strip()
    if not core:
        return text
    start = len(text) - len(text.lstrip())
    end = len(text.rstrip())
    emphasis = "*" * (2 * bool(bold) + bool(italic))
    core = emphasis + core + emphasis
    if script:
        core = f"<{script}>{core}</{script}>"
    return text[:start] + core + text[end:]


def table_markdown(rows):
    def row_line(row):
        cells = [cell.replace("|", r"\|").replace("\n", "<br>") for cell in row]
        return "| " + " | ".join(cells) + " |"

    return "\n".join([row_line(rows[0]), row_line(["---"] * len(rows[0])),
                      *(row_line(row) for row in rows[1:])])


def box_text(coords):
    return ",".join(f"{value:.1f}" for value in coords)


@dataclass
class NativePage:
    number: int
    geometry: str
    title: str = ""
    body: list[str] = field(default_factory=list)
    tables: list[str] = field(default_factory=list)
    details: list[str] = field(default_factory=list)
    warnings: set[str] = field(default_factory=set)
    notes: str = ""
    hidden: bool = False

    def draft(self):
        blocks = [f"<!-- page {self.number} -->"]
        if self.title:
            blocks.append(f"## {self.title}")
        blocks.extend(self.body)
        blocks.extend(self.tables)
        if len(blocks) == 1:
            if self.warnings:
                blocks.append("[불명확: 원본 화면에서 이 페이지의 표시 내용을 확인해야 함]")
            else:
                blocks.append("*(빈 페이지)*")
        return "\n\n".join(blocks)

    def packet(self):
        hidden = " [hidden]" if self.hidden else ""
        warnings = ", ".join(sorted(self.warnings)) or "none"
        lines = [f"== Page {self.number}{hidden} | {self.geometry} ==",
                 f"Review: {warnings}", *self.details]
        if self.notes:
            lines.extend(["[speaker notes — separate]", self.notes])
        return "\n".join(lines)
