import pytest

from security import BearerTokenMiddleware


def make_scope(headers=(), path="/mcp"):
    return {"type": "http", "method": "POST", "path": path, "headers": list(headers)}


async def invoke(middleware, scope):
    messages = []

    async def receive():
        return {"type": "http.request", "body": b"", "more_body": False}

    async def send(message):
        messages.append(message)

    await middleware(scope, receive, send)
    return messages


@pytest.mark.parametrize("headers", [
    [],
    [(b"authorization", b"Bearer wrong-token-that-is-also-long-enough")],
    [(b"authorization", b"Basic correct-token-that-is-at-least-32-chars")],
])
@pytest.mark.anyio
async def test_missing_or_invalid_token_returns_401(headers):
    async def app(scope, receive, send):
        raise AssertionError("Protected app must not be called")

    middleware = BearerTokenMiddleware(app, "correct-token-that-is-at-least-32-chars")
    messages = await invoke(middleware, make_scope(headers))

    assert messages[0]["status"] == 401
    assert (b"www-authenticate", b"Bearer") in messages[0]["headers"]


@pytest.mark.anyio
async def test_correct_token_reaches_mcp_app():
    called = False

    async def app(scope, receive, send):
        nonlocal called
        called = True
        await send({"type": "http.response.start", "status": 200, "headers": []})
        await send({"type": "http.response.body", "body": b"ok"})

    middleware = BearerTokenMiddleware(app, "correct-token-that-is-at-least-32-chars")
    messages = await invoke(
        middleware,
        make_scope([(b"authorization", b"Bearer correct-token-that-is-at-least-32-chars")]),
    )

    assert called is True
    assert messages[0]["status"] == 200


@pytest.mark.anyio
async def test_other_paths_are_not_intercepted():
    called = False

    async def app(scope, receive, send):
        nonlocal called
        called = True

    middleware = BearerTokenMiddleware(app, "correct-token-that-is-at-least-32-chars")
    await invoke(middleware, make_scope(path="/health"))

    assert called is True


def test_reject_short_bearer_secret():
    with pytest.raises(ValueError, match="at least 32 characters"):
        BearerTokenMiddleware(lambda *args: None, "too-short")
