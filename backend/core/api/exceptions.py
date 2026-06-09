from __future__ import annotations

from typing import Any

from rest_framework import status
from rest_framework.response import Response
from rest_framework.views import exception_handler


def custom_exception_handler(exc: Exception, context: dict[str, Any]) -> Response:
    response = exception_handler(exc, context)

    if response is None:
        return Response(
            {
                "success": False,
                "errors": {
                    "non_field_errors": ["An unexpected error occurred."],
                },
            },
            status=status.HTTP_500_INTERNAL_SERVER_ERROR,
        )

    detail: Any = response.data
    if isinstance(detail, dict):
        errors = detail
    else:
        errors = {"detail": detail}

    response.data = {
        "success": False,
        "errors": errors,
    }
    return response
