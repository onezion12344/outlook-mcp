from __future__ import annotations


class OutlookWebMcpError(RuntimeError):
    def __init__(self, message: str, code: str) -> None:
        super().__init__(message)
        self.code = code


class AuthenticationError(OutlookWebMcpError):
    pass


class ApiError(OutlookWebMcpError):
    def __init__(
        self,
        message: str,
        code: str,
        *,
        http_status: int | None = None,
        response_code: str | None = None,
    ) -> None:
        super().__init__(message, code)
        self.http_status = http_status
        self.response_code = response_code
