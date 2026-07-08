from __future__ import annotations

import json

from rest_framework import status
from rest_framework.parsers import FormParser, JSONParser, MultiPartParser
from rest_framework.permissions import AllowAny
from rest_framework.response import Response
from rest_framework.views import APIView

from apps.documents.api.serializers import (
    DocumentGenerationSerializer,
    DocumentRenderSerializer,
    DocumentSectionsGenerationSerializer,
    SectionRegenerateSerializer,
)
from apps.documents.services import (
    build_document_payload,
    docx_base64,
    generate_sections,
    pdf_base64,
    read_image_data,
    regenerate_section,
    render_sections,
    tex_zip_base64,
)


def _document_response_payload(document, *, success_key: str = "document") -> dict:
    section_names = [str(s.get("title", "")) for s in document.document_sections]
    figure_count = sum(1 for s in document.document_sections if s.get("type") == "figure")
    return {
        "success": True,
        success_key: {
            "title": document.title,
            "client_name": document.client_name,
            "source_text": document.source_text,
            "section_names": section_names,
            "section_count": len(section_names),
            "figure_count": figure_count,
            "filename": document.filename,
        },
        "generation_mode": document.generation_mode,
        "latex_source": document.latex_source,
        "pdf_base64": pdf_base64(document.pdf_bytes),
        "docx_base64": docx_base64(document.docx_bytes),
        "tex_zip_base64": tex_zip_base64(document.tex_zip_bytes),
        "filename": f"{document.filename}.pdf",
        "docx_filename": f"{document.filename}.docx",
        "tex_zip_filename": f"{document.filename}-latex.zip",
    }


class DocumentPreviewView(APIView):
    """One-shot generation: transcript in, rendered PDF/DOCX/LaTeX out."""

    permission_classes = [AllowAny]
    parser_classes = [JSONParser, FormParser, MultiPartParser]

    def post(self, request, *args, **kwargs):
        serializer = DocumentGenerationSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        validated = dict(serializer.validated_data)
        validated["source_pdfs"] = request.FILES.getlist("source_pdfs")
        validated["source_images"] = request.FILES.getlist("source_images")
        try:
            validated["image_descriptions"] = json.loads(request.data.get("image_descriptions") or "[]")
        except (json.JSONDecodeError, TypeError):
            validated["image_descriptions"] = []
        image_count = len(validated.get("source_images") or [])
        document = build_document_payload(validated)
        payload = _document_response_payload(document)
        payload["images_received"] = image_count
        return Response(payload, status=status.HTTP_200_OK)


class DocumentSectionsView(APIView):
    """Phase 1 of the review-before-export flow: generate structured sections without rendering."""

    permission_classes = [AllowAny]
    parser_classes = [JSONParser, FormParser, MultiPartParser]

    def post(self, request, *args, **kwargs):
        serializer = DocumentSectionsGenerationSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        validated = dict(serializer.validated_data)
        validated["source_pdfs"] = request.FILES.getlist("source_pdfs")
        validated["source_images"] = request.FILES.getlist("source_images")
        try:
            validated["image_descriptions"] = json.loads(request.data.get("image_descriptions") or "[]")
        except (json.JSONDecodeError, TypeError):
            validated["image_descriptions"] = []

        gen = generate_sections(validated)
        return Response(
            {
                "success": True,
                "title": gen["title"],
                "client_name": gen["client_name"],
                "document_language": gen["document_language"],
                "document_sections": gen["document_sections"],
                "generation_mode": gen["generation_mode"],
                "filename": gen["filename"],
            },
            status=status.HTTP_200_OK,
        )


class DocumentRenderView(APIView):
    """Phase 2 of the review-before-export flow: render PDF/DOCX/LaTeX from (possibly edited) sections."""

    permission_classes = [AllowAny]
    parser_classes = [JSONParser, FormParser, MultiPartParser]

    def post(self, request, *args, **kwargs):
        serializer = DocumentRenderSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        validated = dict(serializer.validated_data)

        try:
            document_sections = json.loads(validated.pop("document_sections"))
        except (json.JSONDecodeError, TypeError):
            return Response(
                {"success": False, "errors": {"document_sections": "Must be a valid JSON array."}},
                status=status.HTTP_400_BAD_REQUEST,
            )
        if not isinstance(document_sections, list):
            return Response(
                {"success": False, "errors": {"document_sections": "Must be a JSON array of sections."}},
                status=status.HTTP_400_BAD_REQUEST,
            )

        source_images = request.FILES.getlist("source_images")
        image_data = read_image_data(source_images, [])

        document = render_sections(
            title=validated["title"],
            client_name=validated.get("client_name", ""),
            source_text="",
            document_sections=document_sections,
            document_language=validated.get("document_language", "en"),
            logo=validated.get("logo"),
            image_data=image_data,
            generation_mode="edited",
        )
        return Response(_document_response_payload(document), status=status.HTTP_200_OK)


class SectionRegenerateView(APIView):
    """Regenerate a single section's content within the review-before-export flow."""

    permission_classes = [AllowAny]
    parser_classes = [JSONParser, FormParser, MultiPartParser]

    def post(self, request, *args, **kwargs):
        serializer = SectionRegenerateSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        validated = serializer.validated_data

        section = regenerate_section(
            source_text=validated["source_text"],
            agent_instructions=validated.get("agent_instructions", ""),
            document_language=validated.get("document_language", "en"),
            template=validated.get("template", "general"),
            target_length=validated.get("target_length", "standard"),
            tone=validated.get("tone", "formal"),
            section_title=validated.get("section_title", ""),
            section_type=validated.get("section_type", "paragraph"),
            section_instructions=validated.get("section_instructions", ""),
        )
        return Response({"success": True, "section": section}, status=status.HTTP_200_OK)