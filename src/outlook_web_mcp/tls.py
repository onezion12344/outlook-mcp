from __future__ import annotations

import ssl

import certifi


def default_context() -> ssl.SSLContext:
    """TLS trust root pinned to certifi's CA bundle.

    Built explicitly (rather than passing a path string, which httpx 0.28 deprecates)
    so a stray SSL_CERT_FILE in the environment cannot swap the trust root.
    """
    return ssl.create_default_context(cafile=certifi.where())
