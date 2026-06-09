from __future__ import annotations

from rest_framework.permissions import IsAuthenticated


class PermissionPolicy(IsAuthenticated):
    """Global permission policy requiring authenticated requests by default."""

    pass
