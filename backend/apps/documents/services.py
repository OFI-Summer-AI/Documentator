from __future__ import annotations

import base64
import json
import io
import os
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from pypdf import PdfReader

from docx import Document
from docx.shared import Cm
from django.utils.text import slugify
from openai import OpenAI
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

ASSETS_DIR = Path(__file__).resolve().parents[2] / "assets"

DocSection = dict[str, Any]


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass(slots=True)
class RenderedDocument:
    title: str
    client_name: str
    source_text: str
    document_sections: list[DocSection]
    filename: str
    pdf_bytes: bytes
    docx_bytes: bytes
    tex_zip_bytes: bytes
    latex_source: str
    generation_mode: str


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def build_document_payload(validated_data: dict[str, Any]) -> RenderedDocument:
    source_text = validated_data["source_text"].strip()
    agent_instructions = validated_data.get("agent_instructions", "").strip()
    document_language = detect_document_language(source_text, agent_instructions)

    pdf_texts = extract_pdf_texts(validated_data.get("source_pdfs") or [])
    if pdf_texts:
        joined = "\n\n---\n\n".join(pdf_texts)
        source_text = f"{source_text}\n\n--- Extracted from uploaded PDFs ---\n\n{joined}".strip()

    # Read uploaded images into memory before any async/generator usage
    source_images = validated_data.get("source_images") or []
    image_descriptions = validated_data.get("image_descriptions") or []
    image_data: list[dict[str, Any]] = []
    for i, img_file in enumerate(source_images):
        desc = image_descriptions[i] if i < len(image_descriptions) else ""
        try:
            img_file.seek(0)
            image_data.append({"bytes": img_file.read(), "description": str(desc).strip()})
        except Exception:
            pass

    doc_data = generate_document_data(source_text, agent_instructions, document_language, image_data)
    generation_mode = str(doc_data.get("generation_mode", "fallback"))
    title = str(doc_data.get("title", "")).strip() or _fallback_title(source_text)
    client_name = str(doc_data.get("client_name", "")).strip()
    document_sections: list[DocSection] = doc_data.get("document_sections", [])
    latex_source = str(doc_data.get("latex_source", "")).strip()

    if not document_sections:
        document_sections = _fallback_sections(source_text, document_language, image_data)
    elif image_data:
        document_sections = _ensure_figures(document_sections, image_data)

    if not latex_source:
        latex_source = _build_latex(title, client_name, document_sections, document_language, image_data)

    filename = slugify(title) or "documentation"
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    pdf_bytes = render_pdf(
        title=title,
        client_name=client_name,
        source_text=source_text,
        document_sections=document_sections,
        document_language=document_language,
        logo=validated_data.get("logo"),
        timestamp=timestamp,
        image_data=image_data,
    )
    docx_bytes = render_docx(
        title=title,
        client_name=client_name,
        document_sections=document_sections,
        document_language=document_language,
        timestamp=timestamp,
        image_data=image_data,
    )
    tex_zip_bytes = build_tex_zip(
        latex_source=latex_source,
        filename=filename,
        image_data=image_data,
    )

    return RenderedDocument(
        title=title,
        client_name=client_name,
        source_text=source_text,
        document_sections=document_sections,
        filename=filename,
        pdf_bytes=pdf_bytes,
        docx_bytes=docx_bytes,
        tex_zip_bytes=tex_zip_bytes,
        latex_source=latex_source,
        generation_mode=generation_mode,
    )


# ---------------------------------------------------------------------------
# LLM generation
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = (
    "You are an expert technical writer. Transform the user-provided transcript or notes into a "
    "polished, client-ready document. You have COMPLETE freedom to decide:\n"
    "- How many sections the document needs.\n"
    "- The title and content of every section.\n"
    "- Whether each section is best represented as a paragraph of text, a bullet list, or a table.\n\n"
    "Rules:\n"
    "1. Never paste raw transcript text into the document. Always synthesise and rewrite.\n"
    "2. Use tables when comparing options, listing attributes, or showing structured data.\n"
    "3. Use bullet lists for action items, requirements, or enumerable facts.\n"
    "4. Use paragraphs for narrative explanations, summaries, and analysis.\n"
    "5. Always include a short executive summary as the first section.\n"
    "6. Always include a version/control table as the last section.\n"
    "7. If the transcript mentions a client name, extract it.\n\n"
    "Return ONLY a valid JSON object with this exact schema (no markdown fences):\n"
    '{\n'
    '  "title": "<concise document title>",\n'
    '  "client_name": "<client name or empty string>",\n'
    '  "document_sections": [\n'
    '    {"title": "<section title>", "type": "paragraph", "content": "<synthesised prose>"},\n'
    '    {"title": "<section title>", "type": "bullets", "content": ["<item>", "<item>"]},\n'
    '    {"title": "<section title>", "type": "table", "content": {"headers": ["<col1>", "<col2>"], "rows": [["<val>", "<val>"]]}},\n'
    '    {"title": "", "type": "figure", "content": {"figure_index": 0, "figure_number": 1, "caption": "<Figure N: descriptive caption>"}}\n'
    '  ],\n'
    '  "latex_source": "<complete compilable LaTeX document>"\n'
    '}'
)


def generate_document_data(
    source_text: str,
    agent_instructions: str = "",
    document_language: str = "en",
    image_data: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    api_key = _env_value("OPENAI_API_KEY")
    model = _env_value("OPENAI_MODEL", default="gpt-4.1-mini")

    if not api_key:
        return _fallback_document_data(source_text, document_language, image_data or [])

    try:
        client = OpenAI(api_key=api_key)
        messages: list[dict[str, str]] = [{"role": "system", "content": _SYSTEM_PROMPT}]
        messages.append({
            "role": "system",
            "content": (
                "Write the section titles and content in Spanish if the transcript/instructions are mainly Spanish. "
                "Write them in English if the transcript/instructions are mainly English. "
                f"Detected language preference: {document_language}."
            ),
        })
        if image_data:
            img_list = "\n".join(
                f"  - figure_index={i}: {d['description'] or '(no description)'}"
                for i, d in enumerate(image_data)
            )
            messages.append({
                "role": "system",
                "content": (
                    f"MANDATORY: The user has uploaded {len(image_data)} image(s) that MUST appear in the "
                    f"document_sections array. Every figure_index from 0 to {len(image_data) - 1} must have "
                    "exactly one corresponding section of type 'figure' in the output.\n\n"
                    f"Available images:\n{img_list}\n\n"
                    "For each image, decide the best position in the document (where the image most supports "
                    "the surrounding content) and insert a figure section there. Write a professional caption "
                    "that describes what the image shows — it does not need to copy the user's description verbatim.\n\n"
                    "Figure section format (title must be empty string):\n"
                    '{"title": "", "type": "figure", "content": {"figure_index": 0, "figure_number": 1, "caption": "Figure 1: Architecture overview"}}\n\n'
                    "Number figures sequentially starting at 1. Omitting any image is NOT allowed."
                ),
            })
        if agent_instructions:
            messages.append({"role": "system", "content": f"Additional author instructions: {agent_instructions}"})
        messages.append({"role": "user", "content": source_text})

        response = client.chat.completions.create(
            model=model,
            temperature=0.3,
            response_format={"type": "json_object"},
            messages=messages,
        )
        raw = response.choices[0].message.content or "{}"
        parsed = json.loads(raw)
        parsed["generation_mode"] = "openai"
        parsed["document_sections"] = _normalise_sections(
            parsed.get("document_sections", []), source_text, document_language, image_data or []
        )
        return parsed
    except Exception:
        return _fallback_document_data(source_text, document_language, image_data or [])


# ---------------------------------------------------------------------------
# Section normalisation & fallback
# ---------------------------------------------------------------------------

def _ensure_figures(sections: list[DocSection], image_data: list[dict[str, Any]]) -> list[DocSection]:
    """Guarantee every uploaded image has a figure section in the document.
    Any image the LLM forgot to place is injected before the last section."""
    placed = {
        int(s["content"]["figure_index"])
        for s in sections
        if s.get("type") == "figure" and isinstance(s.get("content"), dict)
    }
    missing = [i for i in range(len(image_data)) if i not in placed]
    if not missing:
        return sections

    # Count already-placed figures to continue numbering correctly
    next_fig_num = sum(1 for s in sections if s.get("type") == "figure") + 1
    inject = []
    for i in missing:
        img = image_data[i]
        caption = f"Figure {next_fig_num}: {img['description']}" if img["description"] else f"Figure {next_fig_num}"
        inject.append({
            "title": "",
            "type": "figure",
            "content": {"figure_index": i, "figure_number": next_fig_num, "caption": caption},
        })
        next_fig_num += 1

    # Insert before the last section (typically the version control table)
    insert_pos = max(len(sections) - 1, 0)
    return sections[:insert_pos] + inject + sections[insert_pos:]


def _normalise_sections(
    raw: Any,
    source_text: str,
    document_language: str,
    image_data: list[dict[str, Any]],
) -> list[DocSection]:
    if not isinstance(raw, list):
        return _fallback_sections(source_text, document_language, image_data)

    sections: list[DocSection] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        sec_type = str(item.get("type", "paragraph")).strip().lower()
        title = str(item.get("title", "")).strip()
        content = item.get("content", "")

        if sec_type == "bullets":
            content = [str(b).strip() for b in (content if isinstance(content, list) else [str(content)]) if str(b).strip()]
            if not content:
                continue
        elif sec_type == "table":
            if not isinstance(content, dict):
                continue
            headers = [str(h) for h in content.get("headers", [])]
            rows = [[str(c) for c in row] for row in content.get("rows", [])]
            if not headers:
                continue
            content = {"headers": headers, "rows": rows}
        elif sec_type == "figure":
            if not isinstance(content, dict):
                continue
            try:
                figure_index = int(content.get("figure_index", 0))
            except (TypeError, ValueError):
                figure_index = 0
            try:
                figure_number = int(content.get("figure_number", 1))
            except (TypeError, ValueError):
                figure_number = 1
            caption = str(content.get("caption", f"Figure {figure_number}")).strip()
            if figure_index >= len(image_data):
                continue
            content = {"figure_index": figure_index, "figure_number": figure_number, "caption": caption}
        else:
            sec_type = "paragraph"
            content = str(content).strip()
            if not content:
                continue

        sections.append({"title": title, "type": sec_type, "content": content})

    return sections if sections else _fallback_sections(source_text, document_language, image_data)


def _fallback_sections(
    source_text: str,
    document_language: str,
    image_data: list[dict[str, Any]] | None = None,
) -> list[DocSection]:
    summary = _fallback_summary(source_text)
    lines = _extract_lines(source_text)[:6] or ["See source context for details."]
    labels = localized_labels(document_language)
    sections: list[DocSection] = [
        {"title": labels["executive_summary"], "type": "paragraph", "content": summary},
        {"title": labels["key_points"], "type": "bullets", "content": lines},
    ]
    for i, img in enumerate(image_data or []):
        caption = f"Figure {i + 1}: {img['description']}" if img["description"] else f"Figure {i + 1}"
        sections.append({
            "title": "",
            "type": "figure",
            "content": {"figure_index": i, "figure_number": i + 1, "caption": caption},
        })
    sections.append({
        "title": labels["version_control"],
        "type": "table",
        "content": {
            "headers": labels["version_headers"],
            "rows": [["1.0", datetime.now(timezone.utc).strftime("%Y-%m-%d"), "", labels["initial_draft"]]],
        },
    })
    return sections


def _fallback_document_data(
    source_text: str,
    document_language: str,
    image_data: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    title = _fallback_title(source_text)
    client_name = _fallback_client_name(source_text)
    document_sections = _fallback_sections(source_text, document_language, image_data)
    latex_source = _build_latex(title, client_name, document_sections, document_language, image_data or [])
    return {
        "title": title,
        "client_name": client_name,
        "document_sections": document_sections,
        "latex_source": latex_source,
        "generation_mode": "fallback",
    }


# ---------------------------------------------------------------------------
# LaTeX generation
# ---------------------------------------------------------------------------

def _build_latex(
    title: str,
    client_name: str,
    document_sections: list[DocSection],
    document_language: str,
    image_data: list[dict[str, Any]] | None = None,
) -> str:
    logo_path = resolve_ofi_logo_path()
    labels = localized_labels(document_language)
    ofi_logo_block = "\\includegraphics[width=1.8cm]{ofi-logo}" if logo_path else "\\fbox{\\textbf{OFI}}"
    client_line = f"\\textbf{{{_le(labels['client'])}}}: {_le(client_name)}\\\\" if client_name else ""

    toc_items = "\n".join(
        f"\\item {_le(str(sec.get('title', '')))}"
        for sec in document_sections
        if sec.get("title")
    )
    body_blocks: list[str] = []
    has_figures = any(s.get("type") == "figure" for s in document_sections)

    for sec in document_sections:
        sec_title = str(sec.get("title", ""))
        sec_type = str(sec.get("type", "paragraph"))
        content = sec.get("content", "")

        if sec_type == "figure" and isinstance(content, dict):
            fig_num = content.get("figure_number", 1)
            caption = content.get("caption", f"Figure {fig_num}")
            fig_idx = content.get("figure_index", 0)
            body_blocks.append(
                f"\\begin{{figure}}[H]\n"
                f"\\centering\n"
                f"\\includegraphics[width=0.8\\textwidth]{{figure_{fig_idx}}}\n"
                f"\\caption{{{_le(caption)}}}\n"
                f"\\end{{figure}}"
            )
        elif sec_type == "bullets":
            items_str = "\n".join(f"\\item {_le(str(b))}" for b in (content if isinstance(content, list) else [str(content)]))
            body_blocks.append(
                f"\\section*{{{_le(sec_title)}}}\n"
                f"\\begin{{itemize}}[leftmargin=1.2em]\n{items_str}\n\\end{{itemize}}"
            )
        elif sec_type == "table" and isinstance(content, dict):
            headers = content.get("headers", [])
            rows = content.get("rows", [])
            col_count = max(len(headers), max((len(r) for r in rows), default=0), 1)
            col_fmt = "|".join(["l"] * col_count)
            header_row = " & ".join(_le(str(h)) for h in headers) + " \\\\ \\hline"
            data_rows = "\n".join(" & ".join(_le(str(c)) for c in row) + " \\\\" for row in rows)
            body_blocks.append(
                f"\\section*{{{_le(sec_title)}}}\n"
                f"\\begin{{tabular}}{{|{col_fmt}|}}\n"
                f"\\hline\n{header_row}\n{data_rows}\n"
                f"\\hline\n\\end{{tabular}}"
            )
        else:
            if sec_title:
                body_blocks.append(f"\\section*{{{_le(sec_title)}}}\n{_le(str(content))}")
            else:
                body_blocks.append(_le(str(content)))

    body_block = "\n\n".join(body_blocks)
    float_pkg = "\\usepackage{float}\n" if has_figures else ""

    return (
        "\\documentclass[11pt,a4paper]{article}\n"
        "\\usepackage[margin=1in]{geometry}\n"
        "\\usepackage[T1]{fontenc}\n"
        "\\usepackage[utf8]{inputenc}\n"
        "\\usepackage{graphicx}\n"
        f"{float_pkg}"
        "\\usepackage{longtable}\n"
        "\\usepackage{booktabs}\n"
        "\\usepackage{enumitem}\n"
        "\\usepackage{xcolor}\n"
        "\\usepackage{hyperref}\n"
        "\\hypersetup{colorlinks=true, linkcolor=teal!60!black, urlcolor=teal!60!black}\n"
        "\\begin{document}\n"
        f"\\noindent\\hfill{ofi_logo_block}\\par\n"
        "\\vspace{1.2cm}\n"
        "\\begin{center}\n"
        f"{{\\LARGE\\bfseries {_le(title)}}}\\\\[0.45cm]\n"
        f"{client_line}\n"
        f"\\textbf{{{_le(labels['version'])}}}: 1.0\n"
        "\\end{center}\n"
        "\\newpage\n"
        f"\\noindent\\hfill{ofi_logo_block}\\par\n"
        f"\\section*{{{_le(labels['contents'])}}}\n"
        "\\begin{itemize}[leftmargin=1.2em]\n"
        f"{toc_items}\n"
        "\\end{itemize}\n"
        "\\newpage\n\n"
        f"{body_block}\n"
        "\\end{document}\n"
    )


def _to_png_bytes(raw: bytes) -> bytes:
    """Normalise any PIL-readable image to PNG so reportlab and python-docx never choke."""
    from PIL import Image as PILImage
    buf = io.BytesIO(raw)
    img = PILImage.open(buf)
    img.load()
    if img.mode not in ("RGB", "RGBA", "L"):
        img = img.convert("RGBA")
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def _le(value: str) -> str:
    """Escape a string for LaTeX."""
    replacements = [
        ("\\", r"\textbackslash{}"),
        ("{", r"\{"), ("}", r"\}"), ("$", r"\$"), ("&", r"\&"),
        ("#", r"\#"), ("_", r"\_"), ("%", r"\%"),
        ("~", r"\textasciitilde{}"), ("^", r"\textasciicircum{}"),
    ]
    out = value
    for src, dst in replacements:
        out = out.replace(src, dst)
    return out


# ---------------------------------------------------------------------------
# PDF rendering
# ---------------------------------------------------------------------------

def render_pdf(
    *,
    title: str,
    client_name: str,
    source_text: str,
    document_sections: list[DocSection],
    document_language: str,
    logo: Any,
    timestamp: str,
    image_data: list[dict[str, Any]] | None = None,
) -> bytes:
    image_data = image_data or []
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        topMargin=1.8 * cm, bottomMargin=1.5 * cm,
    )
    styles = getSampleStyleSheet()

    S = {
        "title": ParagraphStyle("DocTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=26,
                                leading=30, textColor=colors.HexColor("#0f3d4a"), alignment=TA_CENTER, spaceAfter=14),
        "subtitle": ParagraphStyle("DocSub", parent=styles["BodyText"], fontName="Helvetica", fontSize=10,
                                   textColor=colors.HexColor("#415d66"), alignment=TA_CENTER, spaceAfter=4),
        "h1": ParagraphStyle("SecH", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=14,
                              leading=18, textColor=colors.HexColor("#0e7490"), spaceBefore=14, spaceAfter=8),
        "body": ParagraphStyle("Body", parent=styles["BodyText"], fontName="Helvetica", fontSize=10.5,
                               leading=15, textColor=colors.HexColor("#1f2933"), spaceAfter=4),
        "bullet": ParagraphStyle("Bullet", parent=styles["BodyText"], fontName="Helvetica", fontSize=10.5,
                                 leading=15, textColor=colors.HexColor("#1f2933"), leftIndent=14, spaceAfter=2),
        "toc": ParagraphStyle("Toc", parent=styles["BodyText"], fontName="Helvetica", fontSize=11,
                              leading=16, textColor=colors.HexColor("#1f2933"), leftIndent=10, spaceAfter=3),
        "th": ParagraphStyle("TH", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=9.5,
                             textColor=colors.white),
        "td": ParagraphStyle("TD", parent=styles["BodyText"], fontName="Helvetica", fontSize=9.5,
                             textColor=colors.HexColor("#1f2933")),
        "caption": ParagraphStyle("FigCaption", parent=styles["BodyText"], fontName="Helvetica-Oblique",
                                  fontSize=9, leading=12, textColor=colors.HexColor("#415d66"),
                                  alignment=TA_CENTER, spaceAfter=8),
    }

    story: list[Any] = []

    logo_reader: ImageReader | None = None
    ofi_logo_reader: ImageReader | None = None
    client_logo_bytes: io.BytesIO | None = None

    if logo:
        logo.seek(0)
        client_logo_bytes = io.BytesIO(logo.read())
        logo_reader = ImageReader(client_logo_bytes)

    logo_path = resolve_ofi_logo_path()
    labels = localized_labels(document_language)
    if logo_path:
        ofi_logo_reader = ImageReader(str(logo_path))

    # Cover page
    if client_logo_bytes:
        client_logo_bytes.seek(0)
        story.append(Image(client_logo_bytes, width=3.2 * cm, height=3.2 * cm, kind="proportional"))
        story.append(Spacer(1, 0.35 * cm))
    story.append(Spacer(1, 4.5 * cm))
    story.append(Paragraph(_p(title or "Document"), S["title"]))
    if client_name:
        story.append(Paragraph(_p(client_name), S["subtitle"]))
    story.append(Paragraph(_p(f"{labels['version']} 1.0  |  {timestamp}"), S["subtitle"]))
    story.append(PageBreak())

    # Table of contents — own page
    story.append(Paragraph(labels["contents"], S["h1"]))
    for sec in document_sections:
        if sec.get("title"):
            story.append(Paragraph(_p(str(sec.get("title", ""))), S["toc"], bulletText="-"))
    story.append(PageBreak())

    # Body sections
    available_width = A4[0] - 3.0 * cm

    for sec in document_sections:
        sec_title = str(sec.get("title", ""))
        sec_type = str(sec.get("type", "paragraph"))
        content = sec.get("content")

        if sec_type == "figure" and isinstance(content, dict):
            fig_idx = int(content.get("figure_index", 0))
            caption = str(content.get("caption", ""))
            story.append(Spacer(1, 0.3 * cm))
            if fig_idx < len(image_data):
                try:
                    img_bytes = _to_png_bytes(image_data[fig_idx]["bytes"])
                    fig_img = Image(io.BytesIO(img_bytes), width=min(13 * cm, available_width), kind="proportional")
                    story.append(fig_img)
                except Exception:
                    pass  # image unreadable; caption still renders below
            if caption:
                story.append(Paragraph(_p(caption), S["caption"]))
            story.append(Spacer(1, 0.4 * cm))
            continue

        if sec_title:
            story.append(Paragraph(_p(sec_title), S["h1"]))

        if sec_type == "bullets" and isinstance(content, list):
            for item in content:
                story.append(Paragraph(_p(str(item)), S["bullet"], bulletText="-"))

        elif sec_type == "table" and isinstance(content, dict):
            headers = [str(h) for h in content.get("headers", [])]
            rows = [[str(c) for c in row] for row in content.get("rows", [])]
            if headers:
                col_w = available_width / len(headers)
                tbl_data = [[Paragraph(_p(h), S["th"]) for h in headers]] + \
                           [[Paragraph(_p(c), S["td"]) for c in row] for row in rows]
                pdf_tbl = Table(tbl_data, colWidths=[col_w] * len(headers), repeatRows=1)
                pdf_tbl.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0e7490")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#f5fafb"), colors.white]),
                    ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor("#c9e4e8")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#dbeaec")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]))
                story.append(pdf_tbl)

        else:
            story.append(Paragraph(_p(str(content or "")), S["body"]))

        story.append(Spacer(1, 0.2 * cm))

    def draw_page(canvas, doc_obj) -> None:
        canvas.saveState()
        if ofi_logo_reader is not None:
            canvas.drawImage(
                ofi_logo_reader,
                A4[0] - doc_obj.rightMargin - 1.8 * cm, A4[1] - 2.15 * cm,
                width=1.4 * cm, height=1.4 * cm, preserveAspectRatio=True, mask="auto",
            )
        else:
            canvas.setStrokeColor(colors.HexColor("#c7b37a"))
            canvas.circle(A4[0] - doc_obj.rightMargin - 0.9 * cm, A4[1] - 1.0 * cm, 0.45 * cm, stroke=1, fill=0)
            canvas.setFont("Helvetica-Bold", 8)
            canvas.setFillColor(colors.HexColor("#3a3a3a"))
            canvas.drawCentredString(A4[0] - doc_obj.rightMargin - 0.9 * cm, A4[1] - 1.06 * cm, "OFI")
        canvas.setStrokeColor(colors.HexColor("#d1e6e9"))
        canvas.setLineWidth(0.8)
        canvas.line(doc_obj.leftMargin, A4[1] - 1.15 * cm, A4[0] - doc_obj.rightMargin, A4[1] - 1.15 * cm)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#5b7280"))
        canvas.drawString(doc_obj.leftMargin, 0.9 * cm, "Documentator")
        canvas.drawRightString(A4[0] - doc_obj.rightMargin, 0.9 * cm, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    return buffer.getvalue()


def _p(value: str) -> str:
    return (
        value.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace("\n", "<br/>")
    )


# ---------------------------------------------------------------------------
# DOCX rendering
# ---------------------------------------------------------------------------

def render_docx(
    *,
    title: str,
    client_name: str,
    document_sections: list[DocSection],
    document_language: str,
    timestamp: str,
    image_data: list[dict[str, Any]] | None = None,
) -> bytes:
    image_data = image_data or []
    document = Document()
    labels = localized_labels(document_language)

    header = document.sections[0].header
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.alignment = 2
    logo_path = resolve_ofi_logo_path()
    if logo_path:
        hp.add_run().add_picture(str(logo_path), width=Cm(1.4))
    else:
        run = hp.add_run("OFI")
        run.bold = True

    document.add_heading(title or "Document", level=1)
    if client_name:
        document.add_paragraph(f"{labels['client']}: {client_name}")
    document.add_paragraph(f"{labels['version']}: 1.0")
    document.add_paragraph(f"Generated: {timestamp}")
    document.add_page_break()

    document.add_heading(labels["contents"], level=1)
    for sec in document_sections:
        if sec.get("title"):
            document.add_paragraph(str(sec.get("title", "")), style="List Bullet")
    document.add_page_break()

    for sec in document_sections:
        sec_title = str(sec.get("title", ""))
        sec_type = str(sec.get("type", "paragraph"))
        content = sec.get("content")

        if sec_type == "figure" and isinstance(content, dict):
            fig_idx = int(content.get("figure_index", 0))
            caption = str(content.get("caption", ""))
            if fig_idx < len(image_data):
                try:
                    img_bytes = _to_png_bytes(image_data[fig_idx]["bytes"])
                    document.add_picture(io.BytesIO(img_bytes), width=Cm(12))
                except Exception:
                    pass  # image unreadable; caption still renders below
            if caption:
                cap_para = document.add_paragraph(caption)
                cap_para.alignment = 1  # center
            continue

        if sec_title:
            document.add_heading(sec_title, level=2)

        if sec_type == "bullets" and isinstance(content, list):
            for item in content:
                document.add_paragraph(str(item), style="List Bullet")

        elif sec_type == "table" and isinstance(content, dict):
            headers = [str(h) for h in content.get("headers", [])]
            rows = [[str(c) for c in row] for row in content.get("rows", [])]
            if headers:
                tbl = document.add_table(rows=1 + len(rows), cols=len(headers))
                tbl.style = "Table Grid"
                for i, h in enumerate(headers):
                    tbl.rows[0].cells[i].text = h
                    tbl.rows[0].cells[i].paragraphs[0].runs[0].bold = True
                for ri, row_data in enumerate(rows):
                    for ci, cell_val in enumerate(row_data):
                        tbl.rows[ri + 1].cells[ci].text = cell_val

        else:
            document.add_paragraph(str(content or ""))

    buf = io.BytesIO()
    document.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def _env_value(name: str, default: str = "") -> str:
    return os.getenv(name, default) or default


def _fallback_title(source_text: str) -> str:
    first = next((ln.strip(" :-") for ln in source_text.splitlines() if ln.strip()), "Document")
    return first[:80] or "Document"


def _fallback_client_name(source_text: str) -> str:
    match = re.search(r"(?:client|customer|company|for)[:\-]\s*([^\n]+)", source_text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def _fallback_summary(source_text: str) -> str:
    parts = re.split(r"(?<=[.!?])\s+", re.sub(r"\s+", " ", source_text.strip()))
    return " ".join(parts[:2]) if parts else source_text[:240].strip() or "Generated documentation."


def _extract_lines(source_text: str) -> list[str]:
    lines = []
    for raw in source_text.splitlines():
        clean = re.sub(r"^[-*•\d+.\)]\s*", "", raw.strip())
        if clean:
            lines.append(clean)
    return lines


def build_tex_zip(
    latex_source: str,
    filename: str,
    image_data: list[dict[str, Any]],
) -> bytes:
    """Bundle LaTeX source + all referenced images into a self-contained zip."""
    tex_name = f"{filename}.tex"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(tex_name, latex_source.encode("utf-8"))

        for i, img in enumerate(image_data):
            try:
                zf.writestr(f"figure_{i}.png", _to_png_bytes(img["bytes"]))
            except Exception:
                pass

        logo_path = resolve_ofi_logo_path()
        if logo_path and logo_path.exists():
            try:
                with open(logo_path, "rb") as f:
                    zf.writestr("ofi-logo.png", _to_png_bytes(f.read()))
            except Exception:
                pass

        zf.writestr("compile.sh", f"#!/bin/bash\npdflatex {tex_name}\npdflatex {tex_name}\n")
        zf.writestr("compile.bat", f"@echo off\npdflatex {tex_name}\npdflatex {tex_name}\n")

    return buf.getvalue()


def pdf_base64(pdf_bytes: bytes) -> str:
    return base64.b64encode(pdf_bytes).decode("ascii")


def docx_base64(docx_bytes: bytes) -> str:
    return base64.b64encode(docx_bytes).decode("ascii")


def tex_zip_base64(tex_zip_bytes: bytes) -> str:
    return base64.b64encode(tex_zip_bytes).decode("ascii")


def detect_document_language(source_text: str, agent_instructions: str = "") -> str:
    combined = f"{source_text}\n{agent_instructions}".lower()
    spanish_hits = sum(
        word in combined
        for word in [" el ", " la ", " los ", " las ", " para ", " documento ", " cliente ", " seccion ", " secciones ", " que ", " con "]
    )
    english_hits = sum(
        word in combined
        for word in [" the ", " document ", " client ", " section ", " sections ", " with ", " for ", " and ", " requirements "]
    )
    return "es" if spanish_hits >= english_hits else "en"


def localized_labels(language: str) -> dict[str, Any]:
    if language == "es":
        return {
            "contents": "Contenido",
            "client": "Cliente",
            "version": "Versión",
            "executive_summary": "Resumen Ejecutivo",
            "key_points": "Puntos Clave",
            "version_control": "Control de Versiones",
            "version_headers": ["Versión", "Fecha", "Autor", "Notas"],
            "initial_draft": "Borrador inicial",
        }
    return {
        "contents": "Contents",
        "client": "Client",
        "version": "Version",
        "executive_summary": "Executive Summary",
        "key_points": "Key Points",
        "version_control": "Version Control",
        "version_headers": ["Version", "Date", "Author", "Notes"],
        "initial_draft": "Initial draft",
    }


def resolve_ofi_logo_path() -> Path | None:
    candidates = [
        ASSETS_DIR / "ofi-logo.png",
        ASSETS_DIR / "ofi-logo.png.png",
        ASSETS_DIR / "ofi-logo.jpg",
        ASSETS_DIR / "ofi-logo.jpeg",
        ASSETS_DIR / "ofi-logo.webp",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate

    prefix_matches = sorted(ASSETS_DIR.glob("ofi-logo*")) if ASSETS_DIR.exists() else []
    return prefix_matches[0] if prefix_matches else None


def extract_pdf_texts(pdf_files: list[Any]) -> list[str]:
    """Extract plain text from a list of uploaded PDF file objects."""
    texts: list[str] = []
    for pdf_file in pdf_files:
        try:
            pdf_file.seek(0)
            reader = PdfReader(io.BytesIO(pdf_file.read()))
            pages = []
            for page in reader.pages:
                page_text = page.extract_text() or ""
                if page_text.strip():
                    pages.append(page_text.strip())
            if pages:
                texts.append("\n\n".join(pages))
        except Exception:
            pass
    return texts
