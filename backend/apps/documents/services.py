from __future__ import annotations

import base64
import json
import io
import os
import re
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

OFI_LOGO_PATH = Path(__file__).resolve().parents[2] / "assets" / "ofi-logo.png"

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
    latex_source: str
    generation_mode: str


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def build_document_payload(validated_data: dict[str, Any]) -> RenderedDocument:
    source_text = validated_data["source_text"].strip()
    agent_instructions = validated_data.get("agent_instructions", "").strip()

    pdf_texts = extract_pdf_texts(validated_data.get("source_pdfs") or [])
    if pdf_texts:
        joined = "\n\n---\n\n".join(pdf_texts)
        source_text = f"{source_text}\n\n--- Extracted from uploaded PDFs ---\n\n{joined}".strip()

    doc_data = generate_document_data(source_text, agent_instructions)
    generation_mode = str(doc_data.get("generation_mode", "fallback"))
    title = str(doc_data.get("title", "")).strip() or _fallback_title(source_text)
    client_name = str(doc_data.get("client_name", "")).strip()
    document_sections: list[DocSection] = doc_data.get("document_sections", [])
    latex_source = str(doc_data.get("latex_source", "")).strip()

    if not document_sections:
        document_sections = _fallback_sections(source_text)

    if not latex_source:
        latex_source = _build_latex(title, client_name, document_sections)

    filename = slugify(title) or "documentation"
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")

    pdf_bytes = render_pdf(
        title=title,
        client_name=client_name,
        source_text=source_text,
        document_sections=document_sections,
        logo=validated_data.get("logo"),
        timestamp=timestamp,
    )
    docx_bytes = render_docx(
        title=title,
        client_name=client_name,
        document_sections=document_sections,
        timestamp=timestamp,
    )

    return RenderedDocument(
        title=title,
        client_name=client_name,
        source_text=source_text,
        document_sections=document_sections,
        filename=filename,
        pdf_bytes=pdf_bytes,
        docx_bytes=docx_bytes,
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
    '    {"title": "<section title>", "type": "table", "content": {"headers": ["<col1>", "<col2>"], "rows": [["<val>", "<val>"]]}}\n'
    '  ],\n'
    '  "latex_source": "<complete compilable LaTeX document>"\n'
    '}'
)


def generate_document_data(source_text: str, agent_instructions: str = "") -> dict[str, Any]:
    api_key = _env_value("OPENAI_API_KEY")
    model = _env_value("OPENAI_MODEL", default="gpt-4.1-mini")

    if not api_key:
        return _fallback_document_data(source_text)

    try:
        client = OpenAI(api_key=api_key)
        messages: list[dict[str, str]] = [{"role": "system", "content": _SYSTEM_PROMPT}]
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
        parsed["document_sections"] = _normalise_sections(parsed.get("document_sections", []), source_text)
        return parsed
    except Exception:
        return _fallback_document_data(source_text)


# ---------------------------------------------------------------------------
# Section normalisation & fallback
# ---------------------------------------------------------------------------

def _normalise_sections(raw: Any, source_text: str) -> list[DocSection]:
    if not isinstance(raw, list):
        return _fallback_sections(source_text)

    sections: list[DocSection] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        sec_type = str(item.get("type", "paragraph")).strip().lower()
        title = str(item.get("title", "Section")).strip()
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
        else:
            sec_type = "paragraph"
            content = str(content).strip()
            if not content:
                continue

        sections.append({"title": title, "type": sec_type, "content": content})

    return sections if sections else _fallback_sections(source_text)


def _fallback_sections(source_text: str) -> list[DocSection]:
    summary = _fallback_summary(source_text)
    lines = _extract_lines(source_text)[:6] or ["See source context for details."]
    return [
        {"title": "Executive Summary", "type": "paragraph", "content": summary},
        {"title": "Key Points", "type": "bullets", "content": lines},
        {
            "title": "Version Control",
            "type": "table",
            "content": {
                "headers": ["Version", "Date", "Author", "Notes"],
                "rows": [["1.0", datetime.now(timezone.utc).strftime("%Y-%m-%d"), "", "Initial draft"]],
            },
        },
    ]


def _fallback_document_data(source_text: str) -> dict[str, Any]:
    title = _fallback_title(source_text)
    client_name = _fallback_client_name(source_text)
    document_sections = _fallback_sections(source_text)
    latex_source = _build_latex(title, client_name, document_sections)
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

def _build_latex(title: str, client_name: str, document_sections: list[DocSection]) -> str:
    ofi_logo_block = "\\includegraphics[width=1.8cm]{ofi-logo}" if OFI_LOGO_PATH.exists() else "\\fbox{\\textbf{OFI}}"
    client_line = f"\\textbf{{Client}}: {_le(client_name)}\\\\" if client_name else ""

    toc_items = "\n".join(f"\\item {_le(str(sec.get('title', '')))}" for sec in document_sections)
    body_blocks: list[str] = []

    for sec in document_sections:
        sec_title = str(sec.get("title", "Section"))
        sec_type = str(sec.get("type", "paragraph"))
        content = sec.get("content", "")

        if sec_type == "bullets":
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
            body_blocks.append(f"\\section*{{{_le(sec_title)}}}\n{_le(str(content))}")

    body_block = "\n\n".join(body_blocks)

    return (
        "\\documentclass[11pt,a4paper]{article}\n"
        "\\usepackage[margin=1in]{geometry}\n"
        "\\usepackage[T1]{fontenc}\n"
        "\\usepackage[utf8]{inputenc}\n"
        "\\usepackage{graphicx}\n"
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
        "\\textbf{Version}: 1.0\n"
        "\\end{center}\n"
        "\\newpage\n"
        f"\\noindent\\hfill{ofi_logo_block}\\par\n"
        "\\section*{Contents}\n"
        "\\begin{itemize}[leftmargin=1.2em]\n"
        f"{toc_items}\n"
        "\\end{itemize}\n\n"
        f"{body_block}\n"
        "\\end{document}\n"
    )


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
    logo: Any,
    timestamp: str,
) -> bytes:
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
    }

    story: list[Any] = []

    logo_reader: ImageReader | None = None
    ofi_logo_reader: ImageReader | None = None
    client_logo_bytes: io.BytesIO | None = None

    if logo:
        logo.seek(0)
        client_logo_bytes = io.BytesIO(logo.read())
        logo_reader = ImageReader(client_logo_bytes)

    if OFI_LOGO_PATH.exists():
        ofi_logo_reader = ImageReader(str(OFI_LOGO_PATH))

    # Cover page
    if client_logo_bytes:
        client_logo_bytes.seek(0)
        story.append(Image(client_logo_bytes, width=3.2 * cm, height=3.2 * cm, kind="proportional"))
        story.append(Spacer(1, 0.35 * cm))
    story.append(Spacer(1, 4.5 * cm))
    story.append(Paragraph(_p(title or "Document"), S["title"]))
    if client_name:
        story.append(Paragraph(_p(client_name), S["subtitle"]))
    story.append(Paragraph(_p(f"Version 1.0  |  {timestamp}"), S["subtitle"]))
    story.append(PageBreak())

    # Table of contents
    story.append(Paragraph("Contents", S["h1"]))
    for sec in document_sections:
        story.append(Paragraph(_p(str(sec.get("title", ""))), S["toc"], bulletText="-"))
    story.append(Spacer(1, 0.5 * cm))

    # Body sections
    for sec in document_sections:
        sec_title = str(sec.get("title", ""))
        sec_type = str(sec.get("type", "paragraph"))
        content = sec.get("content")

        story.append(Paragraph(_p(sec_title), S["h1"]))

        if sec_type == "bullets" and isinstance(content, list):
            for item in content:
                story.append(Paragraph(_p(str(item)), S["bullet"], bulletText="-"))

        elif sec_type == "table" and isinstance(content, dict):
            headers = [str(h) for h in content.get("headers", [])]
            rows = [[str(c) for c in row] for row in content.get("rows", [])]
            if headers:
                available_width = A4[0] - 3.0 * cm
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
    timestamp: str,
) -> bytes:
    document = Document()

    header = document.sections[0].header
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.alignment = 2
    if OFI_LOGO_PATH.exists():
        hp.add_run().add_picture(str(OFI_LOGO_PATH), width=Cm(1.4))
    else:
        run = hp.add_run("OFI")
        run.bold = True

    document.add_heading(title or "Document", level=1)
    if client_name:
        document.add_paragraph(f"Client: {client_name}")
    document.add_paragraph("Version: 1.0")
    document.add_paragraph(f"Generated: {timestamp}")
    document.add_page_break()

    document.add_heading("Contents", level=1)
    for sec in document_sections:
        document.add_paragraph(str(sec.get("title", "")), style="List Bullet")

    for sec in document_sections:
        sec_title = str(sec.get("title", ""))
        sec_type = str(sec.get("type", "paragraph"))
        content = sec.get("content")

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
        clean = re.sub(r"^[-*\u2022\d+.\)]\s*", "", raw.strip())
        if clean:
            lines.append(clean)
    return lines


def pdf_base64(pdf_bytes: bytes) -> str:
    return base64.b64encode(pdf_bytes).decode("ascii")


def docx_base64(docx_bytes: bytes) -> str:
    return base64.b64encode(docx_bytes).decode("ascii")


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


