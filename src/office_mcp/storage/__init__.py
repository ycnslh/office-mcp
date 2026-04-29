"""Storage backends — Local, Graph (SharePoint/OneDrive), S3, Azure Blob."""

from office_mcp.storage.base import FileMetadata, StorageBackend
from office_mcp.storage.graph import GraphBackend
from office_mcp.storage.local import LocalBackend
from office_mcp.storage.registry import StorageRegistry

__all__ = [
    "FileMetadata",
    "GraphBackend",
    "LocalBackend",
    "StorageBackend",
    "StorageRegistry",
]
