from __future__ import annotations

from rest_framework import serializers


class DocumentGenerationSerializer(serializers.Serializer):
    source_text = serializers.CharField()
    logo = serializers.ImageField(required=False, allow_null=True)