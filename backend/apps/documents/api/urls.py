from __future__ import annotations

from django.urls import path

from apps.documents.api.views import DocumentPreviewView

app_name = "documents"

urlpatterns = [
    path("preview/", DocumentPreviewView.as_view(), name="preview"),
]