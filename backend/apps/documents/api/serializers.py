from __future__ import annotations

from rest_framework import serializers


class DocumentGenerationSerializer(serializers.Serializer):
    source_text = serializers.CharField()
    agent_instructions = serializers.CharField(required=False, allow_blank=True)
    logo = serializers.ImageField(required=False, allow_null=True)