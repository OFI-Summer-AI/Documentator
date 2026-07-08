from __future__ import annotations

from rest_framework import serializers

TARGET_LENGTH_CHOICES = ("brief", "standard", "detailed")
TONE_CHOICES = ("formal", "consulting", "technical", "casual")
TEMPLATE_CHOICES = ("general", "meeting_minutes", "technical_spec", "proposal", "status_report")


class GenerationOptionsMixin(serializers.Serializer):
    template = serializers.ChoiceField(choices=TEMPLATE_CHOICES, required=False, default="general")
    target_length = serializers.ChoiceField(choices=TARGET_LENGTH_CHOICES, required=False, default="standard")
    tone = serializers.ChoiceField(choices=TONE_CHOICES, required=False, default="formal")
    max_sections = serializers.IntegerField(required=False, allow_null=True, min_value=2, max_value=20)


class DocumentGenerationSerializer(GenerationOptionsMixin, serializers.Serializer):
    source_text = serializers.CharField()
    agent_instructions = serializers.CharField(required=False, allow_blank=True)
    logo = serializers.ImageField(required=False, allow_null=True)


class DocumentSectionsGenerationSerializer(GenerationOptionsMixin, serializers.Serializer):
    source_text = serializers.CharField()
    agent_instructions = serializers.CharField(required=False, allow_blank=True)


class DocumentRenderSerializer(serializers.Serializer):
    title = serializers.CharField()
    client_name = serializers.CharField(required=False, allow_blank=True)
    document_language = serializers.ChoiceField(choices=("en", "es"), required=False, default="en")
    document_sections = serializers.CharField()
    logo = serializers.ImageField(required=False, allow_null=True)


class SectionRegenerateSerializer(GenerationOptionsMixin, serializers.Serializer):
    source_text = serializers.CharField()
    agent_instructions = serializers.CharField(required=False, allow_blank=True)
    document_language = serializers.ChoiceField(choices=("en", "es"), required=False, default="en")
    section_title = serializers.CharField(allow_blank=True)
    section_type = serializers.ChoiceField(choices=("paragraph", "bullets", "table"), required=False, default="paragraph")
    section_instructions = serializers.CharField(required=False, allow_blank=True)