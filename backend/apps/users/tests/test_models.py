from __future__ import annotations

import pytest
from django.contrib.auth import get_user_model


@pytest.mark.django_db
def test_create_user_with_email_as_identifier() -> None:
    user_model = get_user_model()
    user = user_model.objects.create_user(email="user@example.com", password="safe-pass-123")

    assert user.email == "user@example.com"
    assert user.username is None
    assert user.check_password("safe-pass-123")
