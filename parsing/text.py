"""
Plain text parser for email content.

Provides utilities for extracting data from plain text email bodies.
Designed to be generic and extensible.

Usage:
    from parsing.text import TextParser
    
    parser = TextParser()
    result = parser.parse(email.body_text)
    
    # Access cleaned text
    print(result.data["clean_text"])
    
    # Access extracted lines
    for line in result.data["lines"]:
        print(line)
"""

import re
import logging
from typing import Any, Dict, List, Optional, Pattern

from parsing.base import BaseParser, ParseResult

logger = logging.getLogger(__name__)


class TextParser(BaseParser):
    """
    Basic parser for plain text content.
    
    Provides:
    - Text cleaning (whitespace normalization)
    - Line extraction
    - Pattern-based extraction (regex)
    - Key-value extraction ("Key: Value" format)
    
    This is a foundation parser. Extend it or chain it with
    other parsers for more complex extraction.
    """
    
    def __init__(
        self,
        strip_signatures: bool = True,
        normalize_whitespace: bool = True
    ):
        """
        Initialize text parser.
        
        Args:
            strip_signatures: Remove common email signatures.
            normalize_whitespace: Normalize whitespace in output.
        """
        self.strip_signatures = strip_signatures
        self.normalize_whitespace = normalize_whitespace
    
    def parse(self, content: str, **kwargs) -> ParseResult:
        """
        Parse plain text content.
        
        Args:
            content: Plain text content.
            **kwargs: Additional options.
            
        Returns:
            ParseResult with cleaned text and extracted data.
        """
        if not content:
            return ParseResult(
                success=False,
                errors=["Empty content"],
                raw_content="",
            )
        
        data = {}
        warnings = []
        
        # Store raw content
        data["raw"] = content
        
        # Clean the text
        clean = content
        
        if self.strip_signatures:
            clean = self._strip_signatures(clean)
        
        if self.normalize_whitespace:
            clean = self._normalize_whitespace(clean)
        
        data["clean_text"] = clean
        
        # Extract lines (non-empty)
        lines = [line.strip() for line in clean.split("\n") if line.strip()]
        data["lines"] = lines
        data["line_count"] = len(lines)
        
        # Extract key-value pairs (common pattern in business emails)
        key_values = self._extract_key_values(clean)
        data["key_values"] = key_values
        
        return ParseResult(
            success=True,
            data=data,
            warnings=warnings,
            raw_content=content,
        )
    
    def _strip_signatures(self, text: str) -> str:
        """
        Remove common email signature patterns.
        
        Args:
            text: Input text.
            
        Returns:
            Text with signatures removed.
        """
        # Common signature delimiters
        signature_patterns = [
            r"\n--\s*\n.*$",                    # Standard -- signature
            r"\nBest regards,?\n.*$",           # Best regards
            r"\nKind regards,?\n.*$",           # Kind regards
            r"\nRegards,?\n.*$",                # Regards
            r"\nThanks,?\n.*$",                 # Thanks
            r"\nThank you,?\n.*$",              # Thank you
            r"\nSent from my .*$",              # Mobile signatures
            r"\n_+\nFrom:.*$",                  # Forwarded email headers
        ]
        
        for pattern in signature_patterns:
            text = re.sub(pattern, "", text, flags=re.IGNORECASE | re.DOTALL)
        
        return text
    
    def _normalize_whitespace(self, text: str) -> str:
        """
        Normalize whitespace in text.
        
        Args:
            text: Input text.
            
        Returns:
            Text with normalized whitespace.
        """
        # Replace multiple blank lines with single
        text = re.sub(r"\n{3,}", "\n\n", text)
        
        # Replace multiple spaces with single
        text = re.sub(r"[ \t]+", " ", text)
        
        # Strip leading/trailing whitespace from each line
        lines = [line.strip() for line in text.split("\n")]
        text = "\n".join(lines)
        
        return text.strip()
    
    def _extract_key_values(self, text: str) -> Dict[str, str]:
        """
        Extract key-value pairs from text.
        
        Looks for patterns like:
        - "Key: Value"
        - "Key - Value"
        - "Key = Value"
        
        Args:
            text: Input text.
            
        Returns:
            Dictionary of extracted key-value pairs.
        """
        key_values = {}
        
        # Pattern: "Key: Value" or "Key - Value" (with colon or dash)
        pattern = r"^([A-Za-z][A-Za-z0-9\s]{0,30})[:=-]\s*(.+)$"
        
        for line in text.split("\n"):
            line = line.strip()
            match = re.match(pattern, line)
            if match:
                key = match.group(1).strip().lower().replace(" ", "_")
                value = match.group(2).strip()
                key_values[key] = value
        
        return key_values
    
    def extract_by_pattern(
        self,
        text: str,
        pattern: str,
        group: int = 0,
        all_matches: bool = False
    ) -> Optional[str | List[str]]:
        """
        Extract data using a regex pattern.
        
        Args:
            text: Text to search.
            pattern: Regex pattern.
            group: Capture group to return (0 for full match).
            all_matches: If True, return all matches.
            
        Returns:
            Matched string(s), or None if no match.
        """
        if all_matches:
            matches = re.findall(pattern, text, re.IGNORECASE)
            return matches if matches else None
        else:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(group)
            return None
    
    def find_lines_containing(self, text: str, keyword: str) -> List[str]:
        """
        Find all lines containing a keyword.
        
        Args:
            text: Text to search.
            keyword: Keyword to find (case-insensitive).
            
        Returns:
            List of matching lines.
        """
        lines = text.split("\n")
        keyword_lower = keyword.lower()
        return [line.strip() for line in lines if keyword_lower in line.lower()]


class RegexParser(BaseParser):
    """
    Parser that extracts data using configured regex patterns.
    
    Allows defining named patterns and extracts matching content.
    
    Example:
        parser = RegexParser({
            "email": r"[\\w\\.-]+@[\\w\\.-]+\\.\\w+",
            "phone": r"\\+?[\\d\\-\\(\\)\\s]{10,}",
            "date": r"\\d{1,2}/\\d{1,2}/\\d{2,4}",
        })
        result = parser.parse(email_body)
        print(result.data["email"])  # List of found emails
    """
    
    def __init__(self, patterns: Dict[str, str]):
        """
        Initialize with named patterns.
        
        Args:
            patterns: Dict mapping names to regex patterns.
        """
        self.patterns = {
            name: re.compile(pattern, re.IGNORECASE)
            for name, pattern in patterns.items()
        }
    
    def parse(self, content: str, **kwargs) -> ParseResult:
        """
        Extract all pattern matches from content.
        
        Args:
            content: Text to parse.
            
        Returns:
            ParseResult with matches for each pattern.
        """
        data = {}
        found_any = False
        
        for name, pattern in self.patterns.items():
            matches = pattern.findall(content)
            if matches:
                data[name] = matches
                found_any = True
            else:
                data[name] = []
        
        return ParseResult(
            success=found_any,
            data=data,
            raw_content=content,
        )


class SectionParser(BaseParser):
    """
    Parser that extracts content sections based on headers.
    
    Useful for parsing structured emails with multiple sections.
    
    Example:
        parser = SectionParser(section_headers=["Details:", "Summary:", "Notes:"])
        result = parser.parse(email_body)
        print(result.data["sections"]["Details"])
    """
    
    def __init__(
        self,
        section_headers: List[str],
        case_sensitive: bool = False
    ):
        """
        Initialize with section headers to look for.
        
        Args:
            section_headers: List of section header strings.
            case_sensitive: Whether matching is case-sensitive.
        """
        self.section_headers = section_headers
        self.case_sensitive = case_sensitive
    
    def parse(self, content: str, **kwargs) -> ParseResult:
        """
        Extract sections from content.
        
        Args:
            content: Text to parse.
            
        Returns:
            ParseResult with sections dict.
        """
        sections = {}
        current_section = "preamble"
        current_content = []
        
        for line in content.split("\n"):
            line_check = line if self.case_sensitive else line.lower()
            
            matched_header = None
            for header in self.section_headers:
                header_check = header if self.case_sensitive else header.lower()
                if line_check.strip().startswith(header_check):
                    matched_header = header
                    break
            
            if matched_header:
                # Save previous section
                if current_content:
                    sections[current_section] = "\n".join(current_content).strip()
                
                # Start new section
                current_section = matched_header.rstrip(":")
                current_content = []
            else:
                current_content.append(line)
        
        # Save final section
        if current_content:
            sections[current_section] = "\n".join(current_content).strip()
        
        return ParseResult(
            success=len(sections) > 0,
            data={"sections": sections},
            raw_content=content,
        )
