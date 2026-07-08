from __future__ import annotations

from django.urls import path

from apps.documents.api.views import (
    DocumentPreviewView,
    DocumentRenderView,
    DocumentSectionsView,
    SectionRegenerateView,
)

app_name = "documents"

urlpatterns = [
    path("preview/", DocumentPreviewView.as_view(), name="preview"),
    path("generate-sections/", DocumentSectionsView.as_view(), name="generate-sections"),
    path("render/", DocumentRenderView.as_view(), name="render"),
    path("regenerate-section/", SectionRegenerateView.as_view(), name="regenerate-section"),
]