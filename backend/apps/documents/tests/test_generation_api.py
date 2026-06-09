from __future__ import annotations

import base64

import pytest
from rest_framework import status
from rest_framework.test import APIClient


@pytest.mark.django_db
def test_preview_generates_pdf_and_latex() -> None:
    client = APIClient()
    response = client.post(
        "/api/documents/preview/",
        {
            "source_text": "Client: Acme Studio\nWe need a branded PDF document that explains the launch plan.",
            "agent_instructions": "Create a client-facing brief that focuses on database strategy.",
            "section_structure": "Cover page\nContents\nData Integration\nBusiness Logic\nControl Version",
        },
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    assert response.data["success"] is True
    assert response.data["generation_mode"] in {"openai", "fallback"}
    assert response.data["document"]["sections"] == [
        "Cover page",
        "Contents",
        "Data Integration",
        "Business Logic",
        "Control Version",
    ]
    assert response.data["latex_source"].startswith("\\documentclass")
    pdf_bytes = base64.b64decode(response.data["pdf_base64"])
    assert pdf_bytes[:4] == b"%PDF"
    docx_bytes = base64.b64decode(response.data["docx_base64"])
    assert docx_bytes[:2] == b"PK"
    assert response.data["filename"].endswith(".pdf")
    assert response.data["docx_filename"].endswith(".docx")