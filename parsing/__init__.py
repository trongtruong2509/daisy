"""
Email parsing utilities for Office Automation Foundation.

This package provides tools for extracting data from email content:
- Plain text extraction
- HTML parsing
- Extensible parser interface for custom extraction

Design is intentionally generic - no business-specific logic.
"""

from parsing.base import BaseParser, ParseResult
from parsing.text import TextParser
from parsing.html import HtmlParser

__all__ = [
    "BaseParser",
    "ParseResult",
    "TextParser",
    "HtmlParser",
]
