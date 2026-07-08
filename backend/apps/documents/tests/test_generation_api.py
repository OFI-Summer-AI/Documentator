from __future__ import annotations

import base64
import json

import pytest
from rest_framework import status
from rest_framework.test import APIClient

from apps.documents.services import _coerce_section_type

SOURCE_TEXT = (
    "Client: Acme Studio\n"
    "We discussed the product goals, timeline, approval process, and the need for a clean PDF the team can "
    "share with stakeholders. Attendees included Sam (PM) and Lee (Eng). Action items: Sam to draft the "
    "timeline, Lee to review the architecture."
)


@pytest.fixture
def no_openai_key(monkeypatch: pytest.MonkeyPatch) -> None:
    """Force the deterministic fallback generator so tests are hermetic and free of API cost/flakiness."""
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)


@pytest.mark.django_db
def test_preview_generates_pdf_and_latex() -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/preview/",
        {
            "source_text": "Client: Acme Studio\nWe need a branded PDF document that explains the launch plan.",
            "agent_instructions": "Create a client-facing brief that focuses on database strategy.",
        },
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    assert response.data["success"] is True
    assert response.data["generation_mode"] in {"openai", "fallback"}
    assert isinstance(response.data["document"]["section_names"], list)
    assert response.data["document"]["section_count"] >= 1
    assert response.data["latex_source"].startswith("\\documentclass")
    pdf_bytes = base64.b64decode(response.data["pdf_base64"])
    assert pdf_bytes[:4] == b"%PDF"
    docx_bytes = base64.b64decode(response.data["docx_base64"])
    assert docx_bytes[:2] == b"PK"
    assert response.data["filename"].endswith(".pdf")
    assert response.data["docx_filename"].endswith(".docx")


# ---------------------------------------------------------------------------
# Feature 1: structured generation controls (length / tone / max_sections)
# ---------------------------------------------------------------------------

@pytest.mark.django_db
def test_preview_respects_brief_length_and_max_sections(no_openai_key: None) -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/preview/",
        {
            "source_text": SOURCE_TEXT,
            "target_length": "brief",
            "max_sections": "3",
        },
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    assert response.data["generation_mode"] == "fallback"
    assert response.data["document"]["section_count"] <= 3


@pytest.mark.django_db
def test_preview_detailed_length_yields_more_key_points_than_brief(no_openai_key: None) -> None:
    client = APIClient()

    brief = client.post(
        "/api/documents/preview/",
        {"source_text": SOURCE_TEXT, "target_length": "brief"},
        format="json",
    )
    detailed = client.post(
        "/api/documents/preview/",
        {"source_text": SOURCE_TEXT, "target_length": "detailed"},
        format="json",
    )

    assert brief.status_code == detailed.status_code == status.HTTP_200_OK
    assert brief.data["document"]["section_count"] <= detailed.data["document"]["section_count"]


@pytest.mark.django_db
def test_preview_rejects_invalid_options() -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/preview/",
        {"source_text": SOURCE_TEXT, "target_length": "extremely-long", "max_sections": "1"},
        format="json",
    )
    assert response.status_code == status.HTTP_400_BAD_REQUEST


# ---------------------------------------------------------------------------
# Feature 2: named templates
# ---------------------------------------------------------------------------

@pytest.mark.django_db
def test_meeting_minutes_template_adds_expected_sections(no_openai_key: None) -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/preview/",
        {"source_text": SOURCE_TEXT, "template": "meeting_minutes"},
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    section_names = response.data["document"]["section_names"]
    assert "Attendees" in section_names
    assert "Action Items" in section_names


@pytest.mark.django_db
def test_general_template_does_not_add_meeting_minutes_sections(no_openai_key: None) -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/preview/",
        {"source_text": SOURCE_TEXT, "template": "general"},
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    assert "Attendees" not in response.data["document"]["section_names"]


# ---------------------------------------------------------------------------
# Feature 3: review-before-export (generate sections -> edit -> render) + regenerate section
# ---------------------------------------------------------------------------

@pytest.mark.django_db
def test_generate_sections_returns_editable_sections_without_rendering(no_openai_key: None) -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/generate-sections/",
        {"source_text": SOURCE_TEXT, "template": "status_report"},
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    assert response.data["success"] is True
    assert "pdf_base64" not in response.data
    sections = response.data["document_sections"]
    assert isinstance(sections, list)
    assert len(sections) >= 2
    assert all("title" in s and "type" in s and "content" in s for s in sections)


@pytest.mark.django_db
def test_render_reflects_user_edited_sections(no_openai_key: None) -> None:
    client = APIClient()
    gen_response = client.post(
        "/api/documents/generate-sections/",
        {"source_text": SOURCE_TEXT},
        format="json",
    )
    assert gen_response.status_code == status.HTTP_200_OK

    sections = gen_response.data["document_sections"]
    sections[0]["title"] = "Custom Edited Title"
    sections[0]["content"] = "This paragraph was rewritten by the user before export."

    render_response = client.post(
        "/api/documents/render/",
        {
            "title": gen_response.data["title"],
            "client_name": gen_response.data["client_name"],
            "document_language": gen_response.data["document_language"],
            "document_sections": json.dumps(sections),
        },
        format="json",
    )

    assert render_response.status_code == status.HTTP_200_OK
    assert render_response.data["generation_mode"] == "edited"
    assert "Custom Edited Title" in render_response.data["document"]["section_names"]
    assert "Custom Edited Title" in render_response.data["latex_source"]
    pdf_bytes = base64.b64decode(render_response.data["pdf_base64"])
    assert pdf_bytes[:4] == b"%PDF"


@pytest.mark.django_db
def test_render_rejects_malformed_sections_payload() -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/render/",
        {"title": "Doc", "document_sections": "not-json"},
        format="json",
    )
    assert response.status_code == status.HTTP_400_BAD_REQUEST


@pytest.mark.django_db
def test_regenerate_section_returns_new_content(no_openai_key: None) -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/regenerate-section/",
        {
            "source_text": SOURCE_TEXT,
            "section_title": "Key Points",
            "section_type": "bullets",
        },
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    section = response.data["section"]
    assert section["title"] == "Key Points"
    assert section["type"] == "bullets"
    assert isinstance(section["content"], list)
    assert len(section["content"]) > 0


def test_coerce_section_type_keeps_requested_type_when_llm_drifts() -> None:
    # Regression: the LLM sometimes ignores the requested type (e.g. returns a paragraph when
    # bullets were requested). The frontend's per-type editor assumes type never changes on
    # regenerate, so the backend must force it back and reshape the content to match.
    drifted = {"title": "Blockers & Risks", "type": "paragraph", "content": "Sentence one. Sentence two."}
    coerced = _coerce_section_type(drifted, "bullets")
    assert coerced["type"] == "bullets"
    assert isinstance(coerced["content"], list)
    assert len(coerced["content"]) == 2

    drifted_table = {"title": "Status", "type": "table", "content": {"headers": ["A"], "rows": [["x"], ["y"]]}}
    coerced_para = _coerce_section_type(drifted_table, "paragraph")
    assert coerced_para["type"] == "paragraph"
    assert isinstance(coerced_para["content"], str)

    matching = {"title": "Same", "type": "bullets", "content": ["a", "b"]}
    assert _coerce_section_type(matching, "bullets") is matching