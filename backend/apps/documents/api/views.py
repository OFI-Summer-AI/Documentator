from __future__ import annotations

import json

from rest_framework import status
from rest_framework.parsers import FormParser, JSONParser, MultiPartParser
from rest_framework.permissions import AllowAny
from rest_framework.response import Response
from rest_framework.views import APIView

from apps.documents.api.serializers import DocumentGenerationSerializer
from apps.documents.services import build_document_payload, docx_base64, extract_pdf_texts, pdf_base64, tex_zip_base64


class DocumentPreviewView(APIView):
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
        section_names = [str(s.get("title", "")) for s in document.document_sections]
        figure_count = sum(1 for s in document.document_sections if s.get("type") == "figure")
        return Response(
            {
                "success": True,
                "document": {
                    "title": document.title,
                    "client_name": document.client_name,
                    "source_text": document.source_text,
                    "section_names": section_names,
                    "section_count": len(section_names),
                    "figure_count": figure_count,
                    "filename": document.filename,
                },
                "images_received": image_count,
                "generation_mode": document.generation_mode,
                "latex_source": document.latex_source,
                "pdf_base64": pdf_base64(document.pdf_bytes),
                "docx_base64": docx_base64(document.docx_bytes),
                "tex_zip_base64": tex_zip_base64(document.tex_zip_bytes),
                "filename": f"{document.filename}.pdf",
                "docx_filename": f"{document.filename}.docx",
                "tex_zip_filename": f"{document.filename}-latex.zip",
            },
            status=status.HTTP_200_OK,
        )