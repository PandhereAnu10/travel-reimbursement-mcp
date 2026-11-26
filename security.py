import hmac


class BearerTokenMiddleware:
    # Require a matching Authorization: Bearer <token>

    def __init__(self, app, token: str, protected_path: str = "/mcp"):
        if not isinstance(token, str) or len(token) < 32:
            raise ValueError("Bearer token must contain at least 32 characters.")
        self.app = app
        self.token = token
        self.protected_path = protected_path

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http" or scope.get("path") != self.protected_path:
            await self.app(scope, receive, send)
            return

        authorization = next(
            (value for name, value in scope.get("headers", []) if name.lower() == b"authorization"),
            b"",
        ).decode("latin-1")
        scheme, separator, supplied_token = authorization.partition(" ")

        if (
            not separator
            or scheme.lower() != "bearer"
            or not hmac.compare_digest(
                supplied_token.strip().encode("utf-8"), self.token.encode("utf-8")
            )
        ):
            body = b'{"detail":"Unauthorized"}'
            await send({
                "type": "http.response.start",
                "status": 401,
                "headers": [
                    (b"content-type", b"application/json"),
                    (b"content-length", str(len(body)).encode("ascii")),
                    (b"www-authenticate", b"Bearer"),
                ],
            })
            await send({"type": "http.response.body", "body": body})
            return

        await self.app(scope, receive, send)
