from __future__ import annotations

import base64
import html
import re
from io import BytesIO
from urllib.parse import urlparse

import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.oxml.ns import qn

st.set_page_config(page_title="DOCX → HTML", page_icon="🧩", layout="wide")

URL_RE = re.compile(r"https?://[^\s<>]+")


def render_run(run) -> str:
    """Render a Word run while preserving its text and basic inline formatting."""
    text = html.escape(run.text or "", quote=False).replace("\n", "<br>").replace("\t", "    ")
    if not text:
        return ""
    if run.bold:
        text = f"<strong>{text}</strong>"
    if run.italic:
        text = f"<em>{text}</em>"
    if run.underline:
        text = f"<u>{text}</u>"
    return text


def render_paragraph(paragraph) -> str:
    """Render a paragraph in document order, including actual Word hyperlinks."""
    parts = []
    for child in paragraph._p:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag == qn("w:r"):
            # Resolve this XML run back to its formatting and text.
            from docx.text.run import Run
            parts.append(render_run(Run(child, paragraph)))
        elif child.tag == qn("w:hyperlink"):
            rid = child.get(qn("r:id"))
            relationship = paragraph.part.rels.get(rid) if rid else None
            url = relationship.target_ref if relationship and relationship.is_external else ""
            chunks = []
            for run_el in child.findall(".//w:r", {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}):
                from docx.text.run import Run
                chunks.append(render_run(Run(run_el, paragraph)))
            label = "".join(chunks)
            if url:
                parts.append(f'<a href="{html.escape(url, quote=True)}" target="_blank" rel="noopener noreferrer">{label}</a>')
            else:
                parts.append(label)
        elif child.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:proofErr")):
            continue
    return "".join(parts).strip()


def is_heading(paragraph) -> bool:
    name = (paragraph.style.name if paragraph.style else "").lower()
    return name.startswith("heading") or name.startswith("заголовок")


def block(title: str, body: list[str]) -> dict | None:
    body = [part for part in body if part and part.strip()]
    if not body:
        return None
    return {"title": title, "html": f"<!-- Блок: {title} -->\n" + "\n".join(body)}


def p_tag(value: str, *, strong: bool = False) -> str:
    if not value.strip():
        return ""
    return f"<p>{'<strong>' + value + '</strong>' if strong else value}</p>"


def render_cell_paragraphs(cell):
    result = []
    for paragraph in cell.paragraphs:
        rendered = render_paragraph(paragraph)
        # Word frequently stores separate editorial paragraphs as hard line breaks
        # inside a single paragraph. Convert those breaks to separate <p> elements,
        # matching the expected publishing HTML rather than emitting <br> tags.
        html_parts = rendered.split("<br>")
        text_parts = paragraph.text.split("\n")
        for index, html_part in enumerate(html_parts):
            if not html_part.strip():
                continue
            text = text_parts[index].strip() if index < len(text_parts) else html_part
            result.append({
                "text": text,
                "html": html_part,
                "heading": is_heading(paragraph),
            })
    return result
def find_url_in_title(value: str) -> str:
    match = URL_RE.search(value)
    return match.group(0).rstrip(".,;") if match else ""


def make_blocks(docx_bytes: bytes) -> tuple[list[dict], list[str]]:
    """Read only the editor column (last column) and ignore all guidance columns."""
    document = Document(BytesIO(docx_bytes))
    if not document.tables:
        raise ValueError("В документе не найдена таблица с полями редактора.")

    table = document.tables[0]
    rows = {}
    warnings = []
    for row in table.rows[1:]:
        cells = row.cells
        if len(cells) < 5:
            continue
        field = cells[1].text.strip().replace("\n", " ")
        editor_cell = cells[4]
        paragraphs = render_cell_paragraphs(editor_cell)
        if paragraphs:
            rows[field] = paragraphs

    blocks: list[dict] = []

    # 1. Card title + URL. The URL in this field is plain text in some templates,
    # so convert just that URL into an anchor without changing its visible text.
    title_key = next((key for key in rows if key.lower().startswith("название")), None)
    if title_key:
        value = rows[title_key][0]["html"]
        plain = rows[title_key][0]["text"]
        url = find_url_in_title(plain)
        if url and "<a " not in value:
            title_text = plain[:plain.find(url)].strip()
            value = f'{html.escape(title_text, quote=False)} <a href="{html.escape(url, quote=True)}" target="_blank" rel="noopener noreferrer">{html.escape(url, quote=False)}</a>'
        item = block("Название карточки", [p_tag(value)])
        if item:
            blocks.append(item)

    # 2. The short teaser and short description are one publishing block.
    teaser_key = next((key for key in rows if "рекламный слоган" in key.lower()), None)
    short_key = next((key for key in rows if "краткая информация об объекте" in key.lower()), None)
    if teaser_key or short_key:
        body = []
        if teaser_key:
            body.append(p_tag(rows[teaser_key][0]["html"], strong=True))
        if short_key:
            body.append("<p>&nbsp;&nbsp;&nbsp;&nbsp;</p>")
            body.extend(p_tag(p["html"]) for p in rows[short_key])
        item = block("Рекламный слоган и краткая информация об объекте", body)
        if item:
            blocks.append(item)

    # 3. Compact object metadata: take values only from populated editor cells.
    meta_fields = [
        ("Адрес", "Адрес"),
        ("Годы постройки/век", "Годы постройки/век"),
        ("Архитекторы", "Архитекторы"),
        ("Архитектурный стиль", "Архитектурный стиль"),
        ("Округ", "Округ, в котором находится здание"),
        ("Станции метро", "Станции метро"),
    ]
    meta_body = []
    for output_label, prefix in meta_fields:
        key = next((key for key in rows if key.lower().startswith(prefix.lower())), None)
        if key:
            value = " ".join(p["html"] for p in rows[key]).strip()
            # Avoid duplicating the field label in the value; labels are added here.
            meta_body.append(p_tag(f"{output_label}: {value}"))
    item = block("Основная информация об объекте", meta_body)
    if item:
        blocks.append(item)

    # 4. Long building history: the opening paragraph gets its own block, then
    # each actual Word heading starts a new block with the following paragraphs.
    history_key = next((key for key in rows if "информация о здании" in key.lower()), None)
    if history_key:
        history = rows[history_key]
        intro = []
        sections = []
        current_title = None
        current_body = []
        for paragraph in history:
            if paragraph["heading"]:
                if current_title and current_body:
                    sections.append((current_title, current_body))
                current_title = paragraph["text"].strip()
                current_body = [f'<h3>{paragraph["html"]}</h3>']
            elif current_title is None:
                intro.append(p_tag(paragraph["html"]))
            else:
                current_body.append(p_tag(paragraph["html"]))
        if current_title and current_body:
            sections.append((current_title, current_body))
        item = block("Общая информация о здании", intro)
        if item:
            blocks.append(item)
        for title, body in sections:
            item = block(title, body)
            if item:
                blocks.append(item)

    # 5. Remaining populated editor fields that have meaningful content.
    extras = []
    author_key = next((key for key in rows if key.lower() == "автор"), None)
    if author_key:
        extras.extend(p_tag(p["html"]) for p in rows[author_key])
    status_key = next((key for key in rows if "исторический статус" in key.lower()), None)
    if status_key:
        extras.extend(p_tag(p["html"]) for p in rows[status_key])
    item = block("Дополнительная информация об объекте", extras)
    if item:
        blocks.append(item)

    # Warn if populated editor fields were not mapped by the current template.
    known_prefixes = ("название", "адрес", "годы постройки", "архитекторы", "архитектурный стиль",
                      "округ", "станции метро", "рекламный слоган", "краткая информация об объекте",
                      "информация о здании", "автор", "исторический статус")
    for key in rows:
        if not key.lower().startswith(known_prefixes):
            warnings.append(f'Поле «{key}» заполнено, но пока не включено в итоговые блоки.')
    return blocks, warnings


def clipboard_button(label: str, payload: str, key: str) -> None:
    encoded = base64.b64encode(payload.encode("utf-8")).decode("ascii")
    components.html(
        f"""
        <button id="{key}"
          style="width:100%;padding:0.42rem 0.65rem;border:1px solid #d0d7de;border-radius:8px;background:transparent;cursor:pointer;font-size:14px"
          onclick="const bytes=Uint8Array.from(atob('{encoded}'),c=>c.charCodeAt(0));const text=new TextDecoder().decode(bytes);navigator.clipboard.writeText(text).then(()=>this.innerText='Скопировано').catch(()=>this.innerText='Не удалось скопировать')">
          {label}
        </button>
        """,
        height=43,
    )


st.markdown(
    """
    <style>
      .block-container {max-width: 1180px; padding-top: 2.2rem; padding-bottom: 4rem;}
      .hero {padding: 1.5rem 0 1rem 0;}
      .hero h1 {letter-spacing: -0.035em; font-size: 2.25rem; margin-bottom: .35rem;}
      .hero p {color: #777; font-size: 1rem;}
      div[data-testid="stCode"] {border-radius: 12px;}
    </style>
    """,
    unsafe_allow_html=True,
)
st.markdown('<div class="hero"><h1>DOCX <span style="color:#888">→</span> HTML</h1><p>Извлекает только заполненные поля редактора, игнорируя SEO-рекомендации и остальные служебные колонки.</p></div>', unsafe_allow_html=True)
st.caption("Документ обрабатывается локально в процессе приложения. Содержимое не отправляется в сторонние ИИ-сервисы.")

uploaded = st.file_uploader("Загрузите документ Word", type=["docx"])
if uploaded:
    try:
        blocks, warnings = make_blocks(uploaded.getvalue())
        st.success(f"Извлечено HTML-блоков: {len(blocks)}")
        for warning in warnings:
            st.warning(warning)
        if not blocks:
            st.error("В колонке «Поля для заполнения редактором» не найдено заполненных поддерживаемых полей.")
        else:
            all_html = "\n\n".join(item["html"] for item in blocks)
            left, right = st.columns([3, 1])
            with left:
                st.subheader("Результат")
                st.caption("Каждый блок соответствует полям для публикации. Служебные рекомендации не включаются.")
            with right:
                st.download_button("Скачать HTML", all_html.encode("utf-8"), file_name=f"{uploaded.name.rsplit('.', 1)[0]}.html", mime="text/html; charset=utf-8", use_container_width=True)
            for i, item in enumerate(blocks):
                with st.container(border=True):
                    header, copy_col = st.columns([5, 1])
                    with header:
                        st.markdown(f"**{i + 1}. {item['title']}**")
                    with copy_col:
                        clipboard_button("Копировать", item["html"], f"copy-{i}")
                    st.code(item["html"], language="html", line_numbers=False)
            clipboard_button("Копировать всё", all_html, "copy-all")
            with st.expander("Детали извлечения"):
                st.write(f"HTML-блоков: {len(blocks)}")
                st.write(f"URL-ссылок в результате: {all_html.count('<a href=')}")
                st.write("Источник данных: только колонка «Поля для заполнения редактором».")
    except Exception as exc:
        st.error(f"Не удалось обработать файл: {exc}")
        st.info("Проверьте, что файл является корректным .docx и содержит таблицу с колонкой «Поля для заполнения редактором».")
else:
    st.info("Загрузите .docx. Приложение возьмёт только заполненные поля редактора и сохранит встроенные ссылки.")
