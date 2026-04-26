from __future__ import annotations

import io
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def extract_docx_text(path_or_bytes: Path | bytes) -> str:
    data = path_or_bytes.read_bytes() if isinstance(path_or_bytes, Path) else path_or_bytes
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        names = [
            name
            for name in archive.namelist()
            if name.startswith("word/")
            and name.endswith(".xml")
            and (
                name == "word/document.xml"
                or name.startswith("word/header")
                or name.startswith("word/footer")
            )
        ]
        parts: list[str] = []
        for name in names:
            try:
                root = ET.fromstring(archive.read(name))
            except ET.ParseError:
                continue
            for para in root.iter(WORD_NS + "p"):
                text = _paragraph_text(para).strip()
                if text:
                    parts.append(text)
    return "\n".join(parts)


def replace_docx_text(template_bytes: bytes, replacements: dict[str, str]) -> bytes:
    source = io.BytesIO(template_bytes)
    output = io.BytesIO()
    with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _is_word_xml(item.filename):
                data = _replace_in_xml(data, replacements)
            zout.writestr(item, data)
    return output.getvalue()


def normalize_mail_template_text(text: str) -> str:
    text = text.strip()
    if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
        text = text[1:-1]
    return text.replace("\\n", "\n").strip()


def _is_word_xml(filename: str) -> bool:
    return filename.startswith("word/") and filename.endswith(".xml")


def _replace_in_xml(data: bytes, replacements: dict[str, str]) -> bytes:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        return data
    changed = False
    for para in root.iter(WORD_NS + "p"):
        text_nodes = [node for node in para.iter(WORD_NS + "t")]
        if not text_nodes:
            continue
        original = "".join(node.text or "" for node in text_nodes)
        updated = original
        for old, new in replacements.items():
            updated = updated.replace(old, str(new))
        if updated != original:
            text_nodes[0].text = updated
            for node in text_nodes[1:]:
                node.text = ""
            changed = True
    return ET.tostring(root, encoding="utf-8", xml_declaration=True) if changed else data


def _paragraph_text(para: ET.Element) -> str:
    values: list[str] = []
    for node in para.iter():
        if node.tag == WORD_NS + "t" and node.text:
            values.append(node.text)
        elif node.tag == WORD_NS + "tab":
            values.append("\t")
        elif node.tag == WORD_NS + "br":
            values.append("\n")
    return "".join(values)
