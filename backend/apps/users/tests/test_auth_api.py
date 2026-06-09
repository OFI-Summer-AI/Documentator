from __future__ import annotations

import pytest
from django.contrib.auth import get_user_model
from rest_framework import status
from rest_framework.test import APIClient


@pytest.mark.django_db
def test_login_returns_access_and_refresh_tokens() -> None:
    user_model = get_user_model()
    user_model.objects.create_user(email="john@example.com", password="strong-pass-123")

    client = APIClient()
    response = client.post(
        "/api/auth/login/",
        {"email": "john@example.com", "password": "strong-pass-123"},
        format="json",
    )

    assert response.status_code == status.HTTP_200_OK
    assert "access" in response.data
    assert "refresh" in response.data


@pytest.mark.django_db
def test_refresh_returns_new_access_token() -> None:
    user_model = get_user_model()
    user_model.objects.create_user(email="sara@example.com", password="strong-pass-123")

    client = APIClient()
    login_response = client.post(
        "/api/auth/login/",
        {"email": "sara@example.com", "password": "strong-pass-123"},
        format="json",
    )

    refresh_token = login_response.data["refresh"]
    refresh_response = client.post(
        "/api/auth/refresh/",
        {"refresh": refresh_token},
        format="json",
    )

    assert refresh_response.status_code == status.HTTP_200_OK
    assert "access" in refresh_response.data


@pytest.mark.django_db
def test_profile_requires_authentication() -> None:
    client = APIClient()
    response = client.get("/api/users/me/")

    assert response.status_code == status.HTTP_401_UNAUTHORIZED


@pytest.mark.django_db
def test_profile_returns_current_user() -> None:
    user_model = get_user_model()
    user = user_model.objects.create_user(
        email="profile@example.com",
        password="strong-pass-123",
        first_name="Jane",
        last_name="Doe",
    )

    client = APIClient()
    login_response = client.post(
        "/api/auth/login/",
        {"email": "profile@example.com", "password": "strong-pass-123"},
        format="json",
    )

    client.credentials(HTTP_AUTHORIZATION=f"Bearer {login_response.data['access']}")
    response = client.get("/api/users/me/")

    assert response.status_code == status.HTTP_200_OK
    assert response.data["email"] == user.email
    assert response.data["first_name"] == "Jane"
    assert response.data["last_name"] == "Doe"
