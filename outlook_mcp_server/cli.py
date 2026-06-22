import argparse

def parse_args():
    """Parses command line arguments matching the enterprise template."""
    parser = argparse.ArgumentParser(description="Google Tasks MCP Server")
    parser.add_argument("--transport", type=str, help="Transport type (stdio, sse, etc.)")
    parser.add_argument("--host", type=str, help="Host to bind to")
    parser.add_argument("--port", type=int, help="Port to listen on")
    return parser.parse_args()