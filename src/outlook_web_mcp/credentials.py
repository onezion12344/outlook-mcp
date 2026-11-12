from __future__ import annotations

import json
import uuid
from contextlib import suppress
from typing import Protocol

import keyring

from .errors import AuthenticationError
from .models import CredentialBundle

SERVICE_NAME = "ma-nakaya/outlook-web-mcp-server"
PENDING_SERVICE_NAME = f"{SERVICE_NAME}/pending-login"
SECRET_CHUNK_CHARS = 900


class CredentialStore(Protocol):
    def load(self) -> CredentialBundle | None: ...

    def save(self, bundle: CredentialBundle) -> None: ...

    def delete(self) -> None: ...


class KeyringCredentialStore:
    """Store the refresh token only in the operating-system credential vault."""

    def __init__(self, profile: str) -> None:
        self.profile = profile

    @property
    def _metadata_username(self) -> str:
        return f"{self.profile}:metadata"

    @staticmethod
    def _ensure_secure_backend() -> None:
        backend = keyring.get_keyring()
        try:
            priority = float(backend.priority)
        except (TypeError, ValueError, RuntimeError) as exc:
            raise AuthenticationError(
                "A secure keyring backend is unavailable; refusing plaintext credential storage.",
                "keyring_unavailable",
            ) from exc
        if priority <= 0:
            raise AuthenticationError(
                "A secure keyring backend is unavailable; refusing plaintext credential storage.",
                "keyring_unavailable",
            )

    @staticmethod
    def _get(service: str, username: str) -> str | None:
        KeyringCredentialStore._ensure_secure_backend()
        try:
            return keyring.get_password(service, username)
        except Exception as exc:
            raise AuthenticationError(
                "Could not read the operating-system credential vault.", "keyring_read"
            ) from exc

    @staticmethod
    def _set(service: str, username: str, value: str) -> None:
        KeyringCredentialStore._ensure_secure_backend()
        try:
            keyring.set_password(service, username, value)
        except Exception as exc:
            raise AuthenticationError(
                "Could not write the operating-system credential vault.", "keyring_write"
            ) from exc

    @staticmethod
    def _delete(service: str, username: str) -> None:
        KeyringCredentialStore._ensure_secure_backend()
        try:
            keyring.delete_password(service, username)
        except keyring.errors.PasswordDeleteError:
            return
        except Exception as exc:
            raise AuthenticationError(
                "Could not delete Outlook credentials.", "keyring_delete"
            ) from exc

    @staticmethod
    def _chunk_username(profile: str, generation: str, index: int) -> str:
        return f"{profile}:{generation}:refresh:{index}"

    @staticmethod
    def _split(value: str) -> list[str]:
        return [
            value[index : index + SECRET_CHUNK_CHARS]
            for index in range(0, len(value), SECRET_CHUNK_CHARS)
        ]

    def _delete_generation(self, metadata: dict[str, object]) -> None:
        generation = str(metadata.get("generation") or "")
        try:
            count = int(str(metadata.get("refresh_parts") or 0))
        except ValueError:
            return
        if not generation:
            return
        for index in range(min(max(count, 0), 64)):
            self._delete(SERVICE_NAME, self._chunk_username(self.profile, generation, index))

    def load(self) -> CredentialBundle | None:
        payload = self._get(SERVICE_NAME, self._metadata_username)
        if payload is None:
            return None
        try:
            metadata = json.loads(payload)
            if not isinstance(metadata, dict) or metadata.get("schema_version") != 1:
                raise ValueError("unsupported credential schema")
            generation = str(metadata["generation"])
            count = int(str(metadata["refresh_parts"]))
            if not generation or not 1 <= count <= 64:
                raise ValueError("invalid refresh-token chunks")
            parts: list[str] = []
            for index in range(count):
                part = self._get(
                    SERVICE_NAME,
                    self._chunk_username(self.profile, generation, index),
                )
                if part is None:
                    raise AuthenticationError(
                        "Stored Outlook credentials are incomplete.", "credentials_incomplete"
                    )
                parts.append(part)
            metadata["refresh_token"] = "".join(parts)
            return CredentialBundle.from_secret_dict(metadata)
        except AuthenticationError:
            raise
        except (KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
            raise AuthenticationError(
                "Stored Outlook credentials are invalid.", "credentials_invalid"
            ) from exc

    def save(self, bundle: CredentialBundle) -> None:
        previous_payload = self._get(SERVICE_NAME, self._metadata_username)
        previous: dict[str, object] | None = None
        if previous_payload:
            try:
                parsed = json.loads(previous_payload)
                previous = parsed if isinstance(parsed, dict) else None
            except json.JSONDecodeError:
                previous = None

        generation = uuid.uuid4().hex
        parts = self._split(bundle.refresh_token)
        if not 1 <= len(parts) <= 64:
            raise AuthenticationError(
                "The Outlook refresh token has an invalid size.",
                "refresh_token_size",
            )
        written: list[str] = []
        try:
            for index, part in enumerate(parts):
                username = self._chunk_username(self.profile, generation, index)
                self._set(SERVICE_NAME, username, part)
                written.append(username)
            metadata = bundle.to_secret_dict()
            metadata.pop("refresh_token")
            metadata.update(
                {
                    "schema_version": 1,
                    "generation": generation,
                    "refresh_parts": len(parts),
                }
            )
            self._set(
                SERVICE_NAME,
                self._metadata_username,
                json.dumps(metadata, separators=(",", ":"), ensure_ascii=True),
            )
        except AuthenticationError:
            for username in written:
                with suppress(AuthenticationError):
                    self._delete(SERVICE_NAME, username)
            raise
        if previous is not None:
            self._delete_generation(previous)

    def delete(self) -> None:
        payload = self._get(SERVICE_NAME, self._metadata_username)
        if payload:
            try:
                parsed = json.loads(payload)
                if isinstance(parsed, dict):
                    self._delete_generation(parsed)
            except json.JSONDecodeError:
                pass
        self._delete(SERVICE_NAME, self._metadata_username)


class KeyringPendingLoginStore:
    def __init__(self, profile: str) -> None:
        self.profile = profile

    def load(self) -> dict[str, object] | None:
        payload = KeyringCredentialStore._get(PENDING_SERVICE_NAME, self.profile)
        if payload is None:
            return None
        try:
            value = json.loads(payload)
        except json.JSONDecodeError as exc:
            raise AuthenticationError(
                "Stored pending login is invalid.", "pending_invalid"
            ) from exc
        if not isinstance(value, dict):
            raise AuthenticationError("Stored pending login is invalid.", "pending_invalid")
        return value

    def save(self, value: dict[str, object]) -> None:
        KeyringCredentialStore._set(
            PENDING_SERVICE_NAME,
            self.profile,
            json.dumps(value, separators=(",", ":"), ensure_ascii=True),
        )

    def delete(self) -> None:
        KeyringCredentialStore._delete(PENDING_SERVICE_NAME, self.profile)
