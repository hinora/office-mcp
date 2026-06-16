"""
MCP Meta - MCP server for web search and web fetch tools.

Supports:
- web-search: Search the web via Brave Search API (more providers planned)
- web-fetch:  Fetch and extract text content from URLs

Configure via environment variables:
    BRAVE_API_KEY    - Brave Search API key (required for web-search)
    SEARCH_PROVIDER  - Search provider: "brave" (default, only option currently)
"""

__version__ = "0.1.0"
