from __future__ import annotations

import io
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

WORD_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_NS_URI = "http://www.w3.org/XML/1998/namespace"
XMLNS_RE = re.compile(rb'\sxmlns:([A-Za-z_][\w.\-]*)="([^"]+)"')
IGNORABLE_RE = re.compile(rb'\s(?:[A-Za-z_][\w.\-]*:)?Ignorable="([^"]*)"')
EDITABLE_WORD_XML = (
    "word/document.xml",
    "word/header",
    "word/footer",
)


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


def build_docx_from_paragraphs(paragraphs: list[str]) -> bytes:
    document = "\n".join(_paragraph_xml(paragraph) for paragraph in paragraphs)
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {document}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
""",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
""",
        )
        archive.writestr("word/document.xml", document_xml)
    return output.getvalue()


def normalize_mail_template_text(text: str) -> str:
    text = text.strip()
    if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
        text = text[1:-1]
    return text.replace("\\n", "\n").strip()


def _is_word_xml(filename: str) -> bool:
    return filename.endswith(".xml") and (
        filename == EDITABLE_WORD_XML[0]
        or filename.startswith(EDITABLE_WORD_XML[1])
        or filename.startswith(EDITABLE_WORD_XML[2])
    )


def _replace_in_xml(data: bytes, replacements: dict[str, str]) -> bytes:
    try:
        namespace_context = _namespace_context(data)
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
            updated = updated.replace(old, _xml_safe_text(str(new)))
        if updated != original:
            text_nodes[0].text = updated
            for node in text_nodes[1:]:
                node.text = ""
            changed = True
    if not changed:
        return data
    _register_namespaces(namespace_context)
    rendered = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return _restore_ignorable_namespace_declarations(rendered, namespace_context)


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


def _paragraph_xml(paragraph: str) -> str:
    runs = []
    parts = paragraph.split("\n")
    for index, part in enumerate(parts):
        if index:
            runs.append("<w:r><w:br/></w:r>")
        runs.append(f"<w:r><w:t>{escape(part)}</w:t></w:r>")
    return "<w:p>" + "".join(runs) + "</w:p>"


def _namespace_context(data: bytes) -> tuple[dict[str, str], set[str]]:
    declarations = {prefix.decode("utf-8"): uri.decode("utf-8") for prefix, uri in XMLNS_RE.findall(data)}
    ignorable: set[str] = set()
    match = IGNORABLE_RE.search(data)
    if match:
        ignorable = {prefix.decode("utf-8") for prefix in match.group(1).split() if prefix}
    return declarations, ignorable


def _register_namespaces(namespace_context: tuple[dict[str, str], set[str]]) -> None:
    declarations, _ = namespace_context
    for prefix, uri in declarations.items():
        if prefix == "xml" or uri == XML_NS_URI:
            continue
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            continue


def _restore_ignorable_namespace_declarations(
    rendered: bytes, namespace_context: tuple[dict[str, str], set[str]]
) -> bytes:
    declarations, ignorable = namespace_context
    if not ignorable:
        return rendered

    missing = [
        prefix
        for prefix in sorted(ignorable)
        if prefix in declarations and f"xmlns:{prefix}=".encode("utf-8") not in rendered
    ]
    if not missing:
        return rendered

    root_start = 0
    if rendered.startswith(b"<?xml"):
        declaration_end = rendered.find(b"?>")
        if declaration_end != -1:
            root_start = declaration_end + 2
    root_open = rendered.find(b"<", root_start)
    root_close = rendered.find(b">", root_open)
    if root_open == -1 or root_close == -1:
        return rendered

    additions = b"".join(
        f' xmlns:{prefix}="{declarations[prefix]}"'.encode("utf-8") for prefix in missing
    )
    return rendered[:root_close] + additions + rendered[root_close:]


def _xml_safe_text(value: str) -> str:
    return "".join(
        char
        for char in value
        if char in "\t\n\r" or ord(char) >= 0x20
    )
