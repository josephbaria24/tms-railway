async def app(scope, receive, send):
    if scope["type"] == "http" and scope["path"] == "/ping":
        await send({
            "type": "http.response.start",
            "status": 200,
            "headers": [
                (b"content-type", b"text/plain"),
            ],
        })
        await send({
            "type": "http.response.body",
            "body": b"ok",
        })
    else:
        # Forward to FastAPI shortly after startup
        from main import app as fastapi_app
        await fastapi_app(scope, receive, send)
