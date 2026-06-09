from __future__ import annotations

from rest_framework import status
from rest_framework.parsers import FormParser, JSONParser, MultiPartParser
from rest_framework.permissions import AllowAny
from rest_framework.response import Response
from rest_framework.views import APIView

from apps.documents.api.serializers import DocumentGenerationSerializer
from apps.documents.services import build_document_payload, pdf_base64


class DocumentPreviewView(APIView):
    permission_classes = [AllowAny]
    parser_classes = [JSONParser, FormParser, MultiPartParser]

    def post(self, request, *args, **kwargs):
        serializer = DocumentGenerationSerializer(data=request.data)
        serializer.is_valid(raise_exception=True)

        document = build_document_payload(serializer.validated_data)
        return Response(
            {
                "success": True,
                "document": {
                    "title": document.title,
                    "client_name": document.client_name,
                    "source_text": document.source_text,
                    "summary": document.summary,
                    "scope": document.scope,
                    "deliverables": document.deliverables,
                    "timeline": document.timeline,
                    "notes": document.notes,
                    "filename": document.filename,
                },
                "generation_mode": document.generation_mode,
                "latex_source": document.latex_source,
                "pdf_base64": pdf_base64(document.pdf_bytes),
                "filename": f"{document.filename}.pdf",
            },
            status=status.HTTP_200_OK,
        )