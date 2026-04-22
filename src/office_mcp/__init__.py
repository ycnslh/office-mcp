"""office-mcp — MCP server for Word, Excel and PowerPoint documents."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("office-mcp")
except PackageNotFoundError:  # local checkout, package not installed
    __version__ = "0.0.0+dev"

__all__ = ["__version__"]
