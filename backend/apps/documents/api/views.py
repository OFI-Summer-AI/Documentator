from __future__ import annotations

from rest_framework import status
from rest_framework.parsers import FormParser, JSONParser, MultiPartParser
from rest_framework.permissions import AllowAny
from rest_framework.response import Response
from rest_framework.views import APIView

from apps.documents.api.serializers import DocumentGenerationSerializer
from apps.documents.services import build_document_payload, docx_base64, extract_pdf_texts, pdf_base64


class DocumentPreviewView(APIView):
    permission_classes = [AllowAny]
    parser_classes = [JSONParser, FormParser, MultiPartParser]

    def post(self, request, *args, **kwargs):
        serializer = DocumentGenerationSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)
        validated = dict(serializer.validated_data)
        validated["source_pdfs"] = request.FILES.getlist("source_pdfs")
        document = build_document_payload(validated)
        section_names = [str(s.get("title", "")) for s in document.document_sections]
        return Response(
            {
                "success": True,
                "document": {
                    "title": document.title,
                    "client_name": document.client_name,
                    "source_text": document.source_text,
                    "section_names": section_names,
                    "section_count": len(section_names),
                    "filename": document.filename,
                },
                "generation_mode": document.generation_mode,
                "latex_source": document.latex_source,
                "pdf_base64": pdf_base64(document.pdf_bytes),
                "docx_base64": docx_base64(document.docx_bytes),
                "filename": f"{document.filename}.pdf",
                "docx_filename": f"{document.filename}.docx",
            },
            status=status.HTTP_200_OK,
        )