import os
from fastapi import FastAPI
from mcp.server.fastmcp import FastMCP

# import your existing server module so tools get registered
import mcp_server  # this file defines `mcp = FastMCP("proposal-assist")`

app = FastAPI()

# Optional: quick health check so opening the base URL doesn't look "blank"
@app.get("/")
def health():
    return {"status": "ok"}

# Mount MCP endpoint at /mcp
app.mount("/mcp", mcp_server.mcp.asgi_app())