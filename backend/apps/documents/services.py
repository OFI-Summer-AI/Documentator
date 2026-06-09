from __future__ import annotations

import base64
import json
import io
import os
import re
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

from docx import Document
from django.utils.text import slugify
from openai import OpenAI
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


@dataclass(slots=True)
class RenderedDocument:
    title: str
    client_name: str
    source_text: str
    summary: str
    scope: list[str]
    deliverables: list[str]
    timeline: list[str]
    notes: list[str]
    filename: str
    pdf_bytes: bytes
    docx_bytes: bytes
    latex_source: str
    generation_mode: str


def normalize_lines(value: str) -> list[str]:
    lines: list[str] = []
    for raw_line in value.splitlines():
        clean_line = raw_line.strip()
        if not clean_line:
            continue
        clean_line = re.sub(r"^[-*•\d+.\)]\s*", "", clean_line)
        if clean_line:
            lines.append(clean_line)
    return lines


def split_sentences(value: str) -> list[str]:
    cleaned = re.sub(r"\s+", " ", value.strip())
    if not cleaned:
        return []
    parts = re.split(r"(?<=[.!?])\s+", cleaned)
    return [part.strip() for part in parts if part.strip()]


def derive_section(lines: list[str], context: str, fallback: list[str]) -> list[str]:
    if lines:
        return lines[:6]
    context_lines = normalize_lines(context)
    if context_lines:
        return context_lines[:6]
    sentences = split_sentences(context)
    if sentences:
        return sentences[:6]
    return fallback


def build_document_payload(validated_data: dict[str, Any]) -> RenderedDocument:
    source_text = validated_data["source_text"].strip()
    agent_instructions = validated_data.get("agent_instructions", "").strip()
    generation_mode = "fallback"

    document_data = generate_document_data(source_text, agent_instructions)
    if document_data.get("generation_mode"):
        generation_mode = str(document_data["generation_mode"])

    title = str(document_data["title"]).strip()
    client_name = str(document_data.get("client_name", "")).strip()
    summary = str(document_data["summary"]).strip()
    scope = [str(item).strip() for item in document_data.get("scope", []) if str(item).strip()]
    deliverables = [str(item).strip() for item in document_data.get("deliverables", []) if str(item).strip()]
    timeline = [str(item).strip() for item in document_data.get("timeline", []) if str(item).strip()]
    notes = [str(item).strip() for item in document_data.get("notes", []) if str(item).strip()]
    latex_source = str(document_data["latex_source"]).strip()

    filename = slugify(title) or "documentation"
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    pdf_bytes = render_pdf(
        title=title,
        client_name=client_name,
        source_text=source_text,
        summary=summary,
        scope=scope,
        deliverables=deliverables,
        timeline=timeline,
        notes=notes,
        logo=validated_data.get("logo"),
        timestamp=timestamp,
    )
    docx_bytes = render_docx(
        title=title,
        client_name=client_name,
        source_text=source_text,
        summary=summary,
        scope=scope,
        deliverables=deliverables,
        timeline=timeline,
        notes=notes,
        timestamp=timestamp,
    )

    return RenderedDocument(
        title=title,
        client_name=client_name,
        source_text=source_text,
        summary=summary,
        scope=scope,
        deliverables=deliverables,
        timeline=timeline,
        notes=notes,
        filename=filename,
        pdf_bytes=pdf_bytes,
        docx_bytes=docx_bytes,
        latex_source=latex_source,
        generation_mode=generation_mode,
    )


def generate_document_data(source_text: str, agent_instructions: str = "") -> dict[str, Any]:
    api_key = _env_value("OPENAI_API_KEY")
    model = _env_value("OPENAI_MODEL", default="gpt-4.1-mini")

    if not api_key:
        return _fallback_document_data(source_text)

    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=model,
            temperature=0.2,
            response_format={"type": "json_object"},
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You turn raw meeting transcripts or notes into client-ready documentation. "
                        "Return only valid JSON with these keys: title, client_name, summary, scope, deliverables, timeline, notes, latex_source. "
                        "Title must be concise. client_name may be empty. summary must be 2-3 sentences. "
                        "scope, deliverables, timeline, and notes must each be arrays of short bullets. "
                        "The resulting document should follow this order: Cover page, Contents, Data Integration, Studio, Business Logic, Main KPIs, Filter dimensions, Business dimensions, Control Version. "
                        "Never include the raw transcript as-is in any section. Extract and synthesize the information instead. "
                        "latex_source must be a complete LaTeX document that includes only synthesized documentation content and is ready to compile."
                    ),
                },
                {
                    "role": "system",
                    "content": (
                        f"Additional author instructions: {agent_instructions}"
                        if agent_instructions
                        else "Additional author instructions: None provided."
                    ),
                },
                {"role": "user", "content": source_text},
            ],
        )
        payload_text = response.choices[0].message.content or "{}"
        parsed = json.loads(payload_text)
        normalized = _normalize_document_data(parsed, source_text)
        normalized["generation_mode"] = "openai"
        return normalized
    except Exception:
        return _fallback_document_data(source_text)


def _env_value(name: str, default: str = "") -> str:
    return os.getenv(name, default) or default


def _normalize_document_data(payload: dict[str, Any], source_text: str) -> dict[str, Any]:
    title = str(payload.get("title") or _fallback_title(source_text)).strip()
    client_name = str(payload.get("client_name") or "").strip()
    summary = str(payload.get("summary") or _fallback_summary(source_text)).strip()
    scope = _coerce_bullet_list(payload.get("scope"), source_text, [
        "Capture the main discussion points from the transcript.",
        "Keep the document client-facing and easy to review.",
    ])
    deliverables = _coerce_bullet_list(payload.get("deliverables"), source_text, [
        "A polished PDF document",
        "Editable LaTeX source",
        "A reusable documentation structure",
    ])
    timeline = _coerce_bullet_list(payload.get("timeline"), source_text, [
        "Draft the document from the transcript.",
        "Review and refine the generated sections.",
        "Export the final PDF for sharing.",
    ])
    notes = _coerce_bullet_list(payload.get("notes"), source_text, [
        "Generated automatically from the provided transcript.",
        "Client branding can be added when available.",
    ])
    latex_source = str(payload.get("latex_source") or _fallback_latex(title, client_name, source_text, summary, scope, deliverables, timeline, notes)).strip()

    return {
        "title": title,
        "client_name": client_name,
        "summary": summary,
        "scope": scope,
        "deliverables": deliverables,
        "timeline": timeline,
        "notes": notes,
        "latex_source": latex_source,
    }


def _coerce_bullet_list(value: Any, source_text: str, fallback: list[str]) -> list[str]:
    if isinstance(value, list):
        items = [str(item).strip() for item in value if str(item).strip()]
        if items:
            return items[:6]
    fallback_lines = normalize_lines(source_text)
    if fallback_lines:
        return fallback_lines[:6]
    return fallback


def _fallback_document_data(source_text: str) -> dict[str, Any]:
    title = _fallback_title(source_text)
    client_name = _fallback_client_name(source_text)
    summary = _fallback_summary(source_text)
    scope = derive_section([], source_text, [
        "Summarize the meeting and capture the main action items.",
        "Keep the template structure consistent and easy to edit.",
    ])
    deliverables = derive_section([], source_text, [
        "Editable source text",
        "Client-ready PDF",
        "Reusable template structure",
    ])
    timeline = derive_section([], source_text, [
        "Draft the document from the transcript.",
        "Review the generated sections.",
        "Export the final PDF.",
    ])
    notes = derive_section([], source_text, [
        "Generated automatically from the transcript.",
        "Branding is optional.",
    ])
    latex_source = _fallback_latex(title, client_name, source_text, summary, scope, deliverables, timeline, notes)
    return {
        "title": title,
        "client_name": client_name,
        "summary": summary,
        "scope": scope,
        "deliverables": deliverables,
        "timeline": timeline,
        "notes": notes,
        "latex_source": latex_source,
    }


def _fallback_title(source_text: str) -> str:
    first_line = next((line.strip(" :-") for line in source_text.splitlines() if line.strip()), "Meeting Documentation")
    return first_line[:80] or "Meeting Documentation"


def _fallback_client_name(source_text: str) -> str:
    match = re.search(r"(?:client|customer|company|for)[:\-]\s*([^\n]+)", source_text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def _fallback_summary(source_text: str) -> str:
    sentences = split_sentences(source_text)
    if sentences:
        return " ".join(sentences[:2])
    return source_text[:240].strip() or "Transcript-based documentation generated from the provided input."


def _fallback_latex(
    title: str,
    client_name: str,
    source_text: str,
    summary: str,
    scope: list[str],
    deliverables: list[str],
    timeline: list[str],
    notes: list[str],
) -> str:
    def bullet_block(items: list[str]) -> str:
        return "\n".join(f"\\item {latex_escape(item)}" for item in items)

    client_line = f"\\textbf{{Client}}: {latex_escape(client_name)}\\\\" if client_name else ""
    studio_text = _fallback_studio_text(summary, scope, deliverables)

    return f"""\\documentclass[11pt,a4paper]{{article}}
\\usepackage[margin=1in]{{geometry}}
\\usepackage[T1]{{fontenc}}
\\usepackage[utf8]{{inputenc}}
\\usepackage{{graphicx}}
\\usepackage{{enumitem}}
\\usepackage{{xcolor}}
\\usepackage{{hyperref}}
\\hypersetup{{colorlinks=true, linkcolor=teal!60!black, urlcolor=teal!60!black}}
\\begin{{document}}
\\noindent\\hfill\\fbox{{\\textbf{{OFI}}}}\\par
\\vspace{{1.2cm}}
\\begin{{center}}
{{\\LARGE\\bfseries {latex_escape(title)}}}\\\\[0.45cm]
{client_line}
\\textbf{{Version}}: 1.0
\\end{{center}}
\\newpage
\\noindent\\hfill\\fbox{{\\textbf{{OFI}}}}\\par
\\section*{{Contents}}
\\begin{{itemize}}[leftmargin=1.2em]
\\item Data Integration
\\item Studio
\\item Business Logic
\\item Main KPIs
\\item Filter dimensions
\\item Business dimensions
\\item Control Version
\\end{{itemize}}
\\section*{{Data Integration}}
{latex_escape(summary)}
\\section*{{Studio}}
{latex_escape(studio_text)}
\\section*{{Business Logic}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(scope)}
\\end{{itemize}}
\\section*{{Main KPIs}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(deliverables)}
\\end{{itemize}}
\\section*{{Filter dimensions}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(timeline)}
\\end{{itemize}}
\\section*{{Business dimensions}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(notes)}
\\end{{itemize}}
\\section*{{Control Version}}
Version 1.0
\\end{{document}}
"""


def latex_escape(value: str) -> str:
    replacements = {
        "\\": r"\textbackslash{}",
        "{": r"\{",
        "}": r"\}",
        "$": r"\$",
        "&": r"\&",
        "#": r"\#",
        "_": r"\_",
        "%": r"\%",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    escaped = value
    for source, target in replacements.items():
        escaped = escaped.replace(source, target)
    return escaped


def render_latex(
    *,
    title: str,
    client_name: str,
    context: str,
    summary: str,
    scope: list[str],
    deliverables: list[str],
    timeline: list[str],
    notes: list[str],
    timestamp: str,
    has_logo: bool,
) -> str:
    def bullet_block(items: list[str]) -> str:
        return "\n".join(f"\\item {latex_escape(item)}" for item in items)

    client_line = f"\\textbf{{Client}}: {latex_escape(client_name)}\\\\" if client_name else ""
    logo_block = "\\includegraphics[width=4cm]{client-logo}" if has_logo else ""

    return f"""\\documentclass[11pt,a4paper]{{article}}
\\usepackage[margin=1in]{{geometry}}
\\usepackage[T1]{{fontenc}}
\\usepackage[utf8]{{inputenc}}
\\usepackage{{graphicx}}
\\usepackage{{enumitem}}
\\usepackage{{xcolor}}
\\usepackage{{hyperref}}
\\hypersetup{{colorlinks=true, linkcolor=teal!60!black, urlcolor=teal!60!black}}
\\definecolor{{Accent}}{{HTML}}{{0E7490}}
\\definecolor{{Soft}}{{HTML}}{{F3F7F7}}
\\begin{{document}}
\\begin{{center}}
{logo_block}
\\vspace{{0.75cm}}
{{\\LARGE\\bfseries {latex_escape(title)}}}\\\\[0.25cm]
{client_line}
\\textbf{{Generated}}: {latex_escape(timestamp)}
\\end{{center}}
\\vspace{{0.5cm}}
\\section*{{Executive Summary}}
{latex_escape(summary)}
\\section*{{Source Context}}
{latex_escape(context)}
\\section*{{Scope}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(scope)}
\\end{{itemize}}
\\section*{{Deliverables}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(deliverables)}
\\end{{itemize}}
\\section*{{Timeline}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(timeline)}
\\end{{itemize}}
\\section*{{Notes}}
\\begin{{itemize}}[leftmargin=1.2em]
{bullet_block(notes)}
\\end{{itemize}}
\\end{{document}}
"""


def render_pdf(
    *,
    title: str,
    client_name: str,
    source_text: str,
    summary: str,
    scope: list[str],
    deliverables: list[str],
    timeline: list[str],
    notes: list[str],
    logo: Any,
    timestamp: str,
) -> bytes:
    buffer = io.BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleAccent",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=24,
        leading=28,
        textColor=colors.HexColor("#0f3d4a"),
        alignment=TA_CENTER,
        spaceAfter=12,
    )
    subtitle_style = ParagraphStyle(
        "SubtitleAccent",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#415d66"),
        alignment=TA_CENTER,
        spaceAfter=4,
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#0e7490"),
        spaceBefore=10,
        spaceAfter=8,
    )
    body_style = ParagraphStyle(
        "BodyAccent",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=15,
        textColor=colors.HexColor("#1f2933"),
        spaceAfter=4,
    )
    bullet_style = ParagraphStyle(
        "BulletAccent",
        parent=body_style,
        leftIndent=14,
        bulletIndent=0,
        spaceAfter=2,
    )

    story: list[Any] = []
    logo_reader: ImageReader | None = None
    if logo:
        logo.seek(0)
        logo_bytes = io.BytesIO(logo.read())
        logo_reader = ImageReader(logo_bytes)
        story.append(Image(logo_bytes, width=3.2 * cm, height=3.2 * cm, kind="proportional"))
        story.append(Spacer(1, 0.35 * cm))

    # Cover page
    story.append(Spacer(1, 5.0 * cm))
    story.append(Paragraph(_escape_paragraph(title or "[Use Case/App Name]"), title_style))
    story.append(Paragraph(_escape_paragraph(client_name or "Client-ready documentation"), subtitle_style))
    story.append(Paragraph(_escape_paragraph(f"Version 1.0 - {timestamp}"), subtitle_style))
    story.append(PageBreak())

    # Contents page
    story.append(Paragraph("Contents", heading_style))
    for item in [
        "Data Integration",
        "Studio",
        "Business Logic",
        "Main KPIs",
        "Filter dimensions",
        "Business dimensions",
        "Control Version",
    ]:
        story.append(Paragraph(_escape_paragraph(item), bullet_style, bulletText="•"))
    story.append(Spacer(1, 0.4 * cm))

    metadata_rows = [
        [Paragraph("<b>Source type</b>", body_style), Paragraph("Transcript / notes dump", body_style)],
        [Paragraph("<b>Client</b>", body_style), Paragraph(_escape_paragraph(client_name or "Not provided"), body_style)],
        [Paragraph("<b>Context length</b>", body_style), Paragraph(f"{len(source_text.split())} words", body_style)],
    ]
    metadata_table = Table(metadata_rows, colWidths=[4.2 * cm, 11.0 * cm])
    metadata_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f5fafb")),
                ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor("#c9e4e8")),
                ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#dbeaec")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(metadata_table)
    story.append(Spacer(1, 0.5 * cm))

    sections = [
        ("Data Integration", [summary]),
        ("Studio", [_fallback_studio_text(summary, scope, deliverables)]),
        ("Business Logic", scope),
        ("Main KPIs", deliverables),
        ("Filter dimensions", timeline),
        ("Business dimensions", notes),
        ("Control Version", ["1.0"]),
    ]

    for section_title, items in sections:
        story.append(Paragraph(_escape_paragraph(section_title), heading_style))
        if len(items) == 1 and section_title in {"Data Integration", "Studio", "Control Version"}:
            story.append(Paragraph(_escape_paragraph(items[0]), body_style))
        else:
            for item in items:
                story.append(Paragraph(_escape_paragraph(item), bullet_style, bulletText="•"))
        story.append(Spacer(1, 0.18 * cm))

    def draw_page(canvas, doc) -> None:
        canvas.saveState()
        # OFI badge at the top-right on each page.
        canvas.setStrokeColor(colors.HexColor("#c7b37a"))
        canvas.circle(A4[0] - doc.rightMargin - 0.9 * cm, A4[1] - 1.0 * cm, 0.45 * cm, stroke=1, fill=0)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(colors.HexColor("#3a3a3a"))
        canvas.drawCentredString(A4[0] - doc.rightMargin - 0.9 * cm, A4[1] - 1.06 * cm, "OFI")
        canvas.setStrokeColor(colors.HexColor("#d1e6e9"))
        canvas.setLineWidth(0.8)
        canvas.line(doc.leftMargin, A4[1] - 1.15 * cm, A4[0] - doc.rightMargin, A4[1] - 1.15 * cm)
        if logo_reader is not None:
            canvas.drawImage(logo_reader, A4[0] - doc.rightMargin - 2.6 * cm, A4[1] - 2.6 * cm, width=2.2 * cm, height=2.2 * cm, preserveAspectRatio=True, mask="auto")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#5b7280"))
        canvas.drawString(doc.leftMargin, 0.9 * cm, "Documentator")
        canvas.drawRightString(A4[0] - doc.rightMargin, 0.9 * cm, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()

    document.build(story, onFirstPage=draw_page, onLaterPages=draw_page)
    return buffer.getvalue()


def _escape_paragraph(value: str) -> str:
    escaped = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return escaped.replace("\n", "<br/>")


def render_docx(
    *,
    title: str,
    client_name: str,
    source_text: str,
    summary: str,
    scope: list[str],
    deliverables: list[str],
    timeline: list[str],
    notes: list[str],
    timestamp: str,
) -> bytes:
    document = Document()
    header = document.sections[0].header
    header_paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    header_paragraph.alignment = 2
    header_run = header_paragraph.add_run("OFI")
    header_run.bold = True
    document.add_heading(title, level=1)
    if client_name:
        document.add_paragraph(f"Client: {client_name}")
    document.add_paragraph(f"Version: 1.0")
    document.add_paragraph(f"Generated: {timestamp}")

    document.add_page_break()
    document.add_heading("Contents", level=1)
    for item in [
        "Data Integration",
        "Studio",
        "Business Logic",
        "Main KPIs",
        "Filter dimensions",
        "Business dimensions",
        "Control Version",
    ]:
        document.add_paragraph(item, style="List Bullet")

    def add_section(section_title: str, items: list[str], bullet: bool = False) -> None:
        document.add_heading(section_title, level=2)
        if bullet:
            for item in items:
                document.add_paragraph(item, style="List Bullet")
        else:
            document.add_paragraph(items[0] if items else "")

    add_section("Data Integration", [summary])
    add_section("Studio", [_fallback_studio_text(summary, scope, deliverables)])
    add_section("Business Logic", scope, bullet=True)
    add_section("Main KPIs", deliverables, bullet=True)
    add_section("Filter dimensions", timeline, bullet=True)
    add_section("Business dimensions", notes, bullet=True)
    add_section("Control Version", ["1.0"]) 

    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def pdf_base64(pdf_bytes: bytes) -> str:
    return base64.b64encode(pdf_bytes).decode("ascii")


def docx_base64(docx_bytes: bytes) -> str:
    return base64.b64encode(docx_bytes).decode("ascii")


def _fallback_studio_text(summary: str, scope: list[str], deliverables: list[str]) -> str:
    studio_parts: list[str] = []
    if summary.strip():
        studio_parts.append(summary.strip())
    if scope:
        studio_parts.append(f"Focus areas: {', '.join(scope[:3])}.")
    if deliverables:
        studio_parts.append(f"Primary assets: {', '.join(deliverables[:3])}.")
    if not studio_parts:
        return "This section summarizes the extracted studio-level implementation context."
    return " ".join(studio_parts)