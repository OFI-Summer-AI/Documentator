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
        },
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    assert response.data["success"] is True
    assert response.data["generation_mode"] in {"openai", "fallback"}
    assert response.data["latex_source"].startswith("\\documentclass")
    pdf_bytes = base64.b64decode(response.data["pdf_base64"])
    assert pdf_bytes[:4] == b"%PDF"
    assert response.data["filename"].endswith(".pdf")