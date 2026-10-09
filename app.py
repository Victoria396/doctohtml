from __future__ import annotations

import html
import re
import zipfile
from dataclasses import dataclass
from io import BytesIO
from typing import Optional
from xml.etree import ElementTree as ET

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.oxml.ns import qn

st.set_page_config(page_title="DOCX → HTML", page_icon="🧩", layout="wide")

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


@dataclass
class RunPart:
    text: str
    url: Optional[str] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False


@dataclass
class ParagraphPart:
    text: str
    style: str
    runs: list[RunPart]
    is_table: bool = False


def get_relationships(docx_bytes: bytes) -> dict[str, str]:
    """Read external and internal hyperlink relationships directly from DOCX package."""
    with zipfile.ZipFile(BytesIO(docx_bytes)) as archive:
        try:
            root = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
        except KeyError:
            return {}
    result = {}
    for rel in root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        rel_type = rel.attrib.get("Type", "")
        if rid and target and rel_type.endswith("/hyperlink"):
            result[rid] = target
    return result


def paragraph_runs(paragraph_element, rels: dict[str, str]) -> list[RunPart]:
    """Read text in document order, including runs inside w:hyperlink."""
    parts: list[RunPart] = []

    def visit(node, inherited_url: Optional[str] = None):
        tag = node.tag
        if tag == qn("w:hyperlink"):
            rid = node.attrib.get(qn("r:id"))
            url = rels.get(rid, inherited_url)
            for child in node:
                visit(child, url)
            return
        if tag == qn("w:t"):
            text = node.text or ""
            if text:
                parent_rpr = None
                # Formatting is primarily read from the nearest run properties.
                parts.append(RunPart(text=text, url=inherited_url))
            return
        if tag == qn("w:tab"):
            parts.append(RunPart(text="\t", url=inherited_url))
            return
        if tag in (qn("w:br"), qn("w:cr")):
            parts.append(RunPart(text="\n", url=inherited_url))
            return
        for child in node:
            visit(child, inherited_url)

    for child in paragraph_element:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag == qn("w:r"):
            text_nodes = child.findall(".//w:t", NS)
            text = "".join(n.text or "" for n in text_nodes)
            if not text:
                # Keep explicit breaks/tabs within otherwise empty runs.
                for sub in child:
                    if sub.tag == qn("w:br") or sub.tag == qn("w:cr"):
                        parts.append(RunPart(text="\n"))
                    elif sub.tag == qn("w:tab"):
                        parts.append(RunPart(text="\t"))
                continue
            rpr = child.find("w:rPr", NS)
            bold = rpr is not None and rpr.find("w:b", NS) is not None
            italic = rpr is not None and rpr.find("w:i", NS) is not None
            underline = rpr is not None and rpr.find("w:u", NS) is not None and rpr.find("w:u", NS).attrib.get(qn("w:val"), "single") != "none"
            # Preserve run chunks and their formatting.
            for node in child:
                if node.tag == qn("w:t") and node.text:
                    parts.append(RunPart(node.text, bold=bool(bold), italic=bool(italic), underline=bool(underline)))
                elif node.tag == qn("w:tab"):
                    parts.append(RunPart("\t", bold=bool(bold), italic=bool(italic), underline=bool(underline)))
                elif node.tag in (qn("w:br"), qn("w:cr")):
                    parts.append(RunPart("\n", bold=bool(bold), italic=bool(italic), underline=bool(underline)))
        elif child.tag == qn("w:hyperlink"):
            rid = child.attrib.get(qn("r:id"))
            url = rels.get(rid)
            for run in child.findall(".//w:r", NS):
                rpr = run.find("w:rPr", NS)
                bold = rpr is not None and rpr.find("w:b", NS) is not None
                italic = rpr is not None and rpr.find("w:i", NS) is not None
                underline_el = rpr.find("w:u", NS) if rpr is not None else None
                underline = underline_el is not None and underline_el.attrib.get(qn("w:val"), "single") != "none"
                for node in run:
                    if node.tag == qn("w:t") and node.text:
                        parts.append(RunPart(node.text, url=url, bold=bool(bold), italic=bool(italic), underline=bool(underline)))
                    elif node.tag == qn("w:tab"):
                        parts.append(RunPart("\t", url=url))
                    elif node.tag in (qn("w:br"), qn("w:cr")):
                        parts.append(RunPart("\n", url=url))
        else:
            visit(child)
    return parts


def element_text(el) -> str:
    return "".join(t.text or "" for t in el.findall(".//w:t", NS))


def read_document(docx_bytes: bytes) -> tuple[list[ParagraphPart], list[str]]:
    rels = get_relationships(docx_bytes)
    doc = Document(BytesIO(docx_bytes))
    paragraphs: list[ParagraphPart] = []
    warnings: list[str] = []

    # Include document body paragraphs and table paragraphs in their document order.
    with zipfile.ZipFile(BytesIO(docx_bytes)) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find("w:body", NS)
    if body is None:
        return [], ["Не удалось найти содержимое документа."]
    for child in body:
        if child.tag == qn("w:p"):
            parts = paragraph_runs(child, rels)
            text = "".join(part.text for part in parts)
            style = ""
            # Match paragraph element to python-docx paragraph by order for style name.
            paragraphs.append(ParagraphPart(text=text, style=style, runs=parts, is_table=False))
        elif child.tag == qn("w:tbl"):
            for row in child.findall("w:tr", NS):
                for cell in row.findall("w:tc", NS):
                    for p in cell.findall("w:p", NS):
                        parts = paragraph_runs(p, rels)
                        text = "".join(part.text for part in parts)
                        paragraphs.append(ParagraphPart(text=text, style="", runs=parts, is_table=True))

    # Paragraph styles aren't easy to align when tables and XML are traversed separately.
    # Map normal document paragraphs in sequence to their python-docx styles.
    normal_styles = iter(p.style.name if p.style else "" for p in doc.paragraphs)
    normal_index = 0
    for item in paragraphs:
        if not item.is_table:
            try:
                item.style = next(normal_styles)
            except StopIteration:
                pass
            normal_index += 1

    if not rels:
        warnings.append("В DOCX не найдено отношений гиперссылок. Если в Word видны ссылки, проверьте исходный файл.")
    missing = [part.text for p in paragraphs for part in p.runs if part.underline and not part.url and part.text.strip()]
    if missing:
        preview = ", ".join(repr(x) for x in missing[:8])
        warnings.append(f"Обнаружены подчёркнутые фрагменты без извлечённого URL: {preview}. Они сохранены как обычный текст.")
    return paragraphs, warnings


def inline_html(parts: list[RunPart]) -> str:
    rendered = []
    for part in parts:
        text = html.escape(part.text, quote=False).replace("\n", "<br>").replace("\t", "    ")
        if part.url:
            url = html.escape(part.url, quote=True)
            text = f'<a href="{url}" target="_blank">{text}</a>'
        if part.bold:
            text = f"<strong>{text}</strong>"
        if part.italic:
            text = f"<em>{text}</em>"
        rendered.append(text)
    return "".join(rendered)


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def detect_heading(item: ParagraphPart) -> bool:
    style = (item.style or "").lower()
    if "heading" in style or "заголовок" in style:
        return True
    # Only use Word formatting as a signal; don't invent or rewrite a heading.
    return bool(item.text.strip()) and all(run.bold for run in item.runs if run.text.strip()) and len(item.text.strip()) < 100


def label_for(text: str, index: int) -> str:
    known = [
        ("рекламный слоган", "Рекламный слоган и краткая информация об объекте"),
        ("краткая информация", "Рекламный слоган и краткая информация об объекте"),
        ("адрес", "Основная информация об объекте"),
        ("архитектурный стиль", "Основная информация об объекте"),
        ("годы постройки", "Основная информация об объекте"),
        ("станции метро", "Основная информация об объекте"),
        ("округ", "Основная информация об объекте"),
        ("предыстория", "Предыстория"),
        ("современное использование", "Современное использование"),
        ("исторический статус", "Исторический статус"),
        ("информация о здании", "Информация о здании"),
        ("информация об объекте", "Информация об объекте"),
    ]
    lower = text.lower()
    for phrase, label in known:
        if phrase in lower:
            return label
    return clean_text(text)[:72] or f"Блок {index + 1}"


def to_blocks(paragraphs: list[ParagraphPart]) -> list[dict]:
    """Create stable copy-ready blocks based only on document paragraphs and headings."""
    content = [p for p in paragraphs if p.text.strip()]
    if not content:
        return []
    blocks: list[dict] = []
    current: list[ParagraphPart] = []
    current_label = "Основная информация об объекте"

    def flush():
        nonlocal current
        if not current:
            return
        body = []
        for p in current:
            tag = "h3" if detect_heading(p) else "p"
            value = inline_html(p.runs)
            if tag == "p" and value.strip():
                body.append(f"<p>{value}</p>")
            elif tag == "h3":
                body.append(f"<h3>{value}</h3>")
        if body:
            block_html = f"<!-- Блок: {current_label} -->\n" + "\n".join(body)
            blocks.append({"title": current_label, "html": block_html})
        current = []

    for i, p in enumerate(content):
        text = p.text.strip()
        if detect_heading(p) and current:
            flush()
            current_label = label_for(text, i)
        elif not current:
            current_label = label_for(text, i)
        current.append(p)
    flush()
    return blocks


st.markdown(
    """
    <style>
      .block-container {max-width: 1180px; padding-top: 2.2rem; padding-bottom: 4rem;}
      .hero {padding: 1.5rem 0 1rem 0;}
      .hero h1 {letter-spacing: -0.035em; font-size: 2.25rem; margin-bottom: .35rem;}
      .hero p {color: #777; font-size: 1rem;}
      div[data-testid="stCode"] {border-radius: 12px;}
      .hint {color: #737373; font-size: .9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="hero"><h1>DOCX <span style="color:#888">→</span> HTML</h1><p>Преобразование документов в готовые HTML-блоки. Без ИИ и без генерации текста.</p></div>', unsafe_allow_html=True)
st.caption("Файл обрабатывается в памяти текущей сессии. Содержимое документа не отправляется в сторонние ИИ-сервисы.")

uploaded = st.file_uploader("Загрузите документ Word", type=["docx"], help="Поддерживается современный формат .docx. Старый .doc сначала сохраните из Word как .docx.")
if uploaded:
    data = uploaded.getvalue()
    try:
        paragraphs, warnings = read_document(data)
        blocks = to_blocks(paragraphs)
        st.success(f"Документ прочитан: {len(paragraphs)} абзацев, {len(blocks)} HTML-блоков.")
        for warning in warnings:
            st.warning(warning)
        if not blocks:
            st.error("Не удалось найти непустые текстовые блоки в документе.")
        else:
            all_html = "\n\n-----\n\n".join(block["html"] for block in blocks)
            col1, col2 = st.columns([3, 1])
            with col1:
                st.subheader("Результат")
                st.caption("Каждый блок можно скопировать отдельно. Используйте «Копировать всё» для полного результата.")
            with col2:
                st.download_button("⬇ Скачать HTML", data=all_html.encode("utf-8"), file_name=f"{uploaded.name.rsplit('.', 1)[0]}.html", mime="text/html; charset=utf-8", use_container_width=True)
            components.html(
                """
                <script>
                // Clipboard controls are rendered for a consistent, one-click copy experience.
                </script>
                """,
                height=0,
            )
            for i, block in enumerate(blocks):
                with st.container(border=True):
                    top = st.columns([5, 1])
                    with top[0]:
                        st.markdown(f"**{i + 1}. {block['title']}**")
                    with top[1]:
                        st.button("Копировать", key=f"copy_{i}", on_click=lambda value=block["html"]: st.session_state.update({"copy_payload": value}), use_container_width=True)
                    st.code(block["html"], language="html", line_numbers=False)
                    if st.session_state.get("copy_payload") == block["html"]:
                        # Streamlit's code block provides a native copy icon; also offer an explicit clipboard button above.
                        components.html(
                            f"""
                            <button id="copy-{i}" style="display:none">copy</button>
                            <script>
                            const payload = {block["html"]!r};
                            if (navigator.clipboard && window.parent) {{
                              navigator.clipboard.writeText(payload).catch(() => {{}});
                            }}
                            </script>
                            """,
                            height=0,
                        )
                        st.success("HTML блока подготовлен. Если браузер не разрешил копирование автоматически, используйте иконку копирования у кода.")
            with st.expander("Проверка и ограничения"):
                st.write(f"- Найдено абзацев: {len(paragraphs)}")
                st.write(f"- Найдено HTML-блоков: {len(blocks)}")
                st.write(f"- Найдено URL-гиперссылок: {sum(1 for p in paragraphs for r in p.runs if r.url)}")
                st.write("- Табличные ячейки извлекаются как последовательность абзацев, без HTML-таблиц.")
                st.write("- В этой первой версии смысловые блоки определяются по заголовкам и форматированию Word. Проверьте разбиение перед публикацией.")
    except Exception as exc:
        st.error(f"Не удалось обработать файл: {exc}")
        st.info("Убедитесь, что это корректный .docx, а не старый .doc, переименованный в .docx.")
else:
    st.info("Загрузите .docx, чтобы сформировать блоки HTML с сохранением обнаруженных гиперссылок.")
