"""MCP Server for Google Tasks API."""

import logging

from fastmcp import FastMCP
from fastmcp_credentials import CredentialMiddleware, HeaderCredentialBackend

from outlook_mcp_server.cli import parse_args
from outlook_mcp_server.config import configure_logging
from outlook_mcp_server.tools import register_tools

configure_logging()
logger = logging.getLogger("tasks-mcp-server")

# Use MewCP's standardized cloud credential backend
backend = HeaderCredentialBackend()
mcp = FastMCP(
    "MewCP Google Tasks MCP Server",
    middleware=[CredentialMiddleware(backend, "oauth")],
)
register_tools(mcp)

# Expose ASGI app for hosting platform's (e.g. Vercel / Cloud Run) runtime.
app = mcp.http_app(path="/mcp", transport="streamable-http", stateless_http=True)

if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("MewCP Google Tasks MCP Server Starting")
    logger.info("=" * 60)

    args = parse_args()

    run_kwargs = {}
    if args.transport:
        run_kwargs["transport"] = args.transport
        logger.info(f"Transport: {args.transport}")
    if args.host:
        run_kwargs["host"] = args.host
        logger.info(f"Host: {args.host}")
    if args.port:
        run_kwargs["port"] = args.port
        logger.info(f"Port: {args.port}")

    try:
        mcp.run(**run_kwargs)
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server crashed: {e}", exc_info=True)
        raise