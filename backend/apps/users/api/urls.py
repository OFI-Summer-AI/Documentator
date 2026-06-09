from __future__ import annotations

from django.urls import path

from apps.users.api.views import UserProfileView

app_name = "users"

urlpatterns = [
    path("me/", UserProfileView.as_view(), name="me"),
]
