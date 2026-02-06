"""
HTML parser for email content.

Provides utilities for extracting data from HTML email bodies.
Uses BeautifulSoup for reliable HTML parsing.

Usage:
    from parsing.html import HtmlParser
    
    parser = HtmlParser()
    result = parser.parse(email.body_html)
    
    # Get plain text version
    print(result.data["text"])
    
    # Get extracted links
    for link in result.data["links"]:
        print(link["href"], link["text"])
"""

import logging
import re
from typing import Any, Dict, List, Optional

try:
    from bs4 import BeautifulSoup
    HAS_BEAUTIFULSOUP = True
except ImportError:
    HAS_BEAUTIFULSOUP = False

from parsing.base import BaseParser, ParseResult

logger = logging.getLogger(__name__)


class HtmlParser(BaseParser):
    """
    Parser for HTML email content.
    
    Provides:
    - HTML to plain text conversion
    - Link extraction
    - Image extraction
    - Table extraction
    - Tag-based content extraction
    
    Requires beautifulsoup4: pip install beautifulsoup4
    """
    
    def __init__(self, parser: str = "html.parser"):
        """
        Initialize HTML parser.
        
        Args:
            parser: BeautifulSoup parser to use.
                    Options: "html.parser", "lxml", "html5lib"
        """
        if not HAS_BEAUTIFULSOUP:
            raise ImportError(
                "BeautifulSoup is required for HTML parsing. "
                "Install: pip install beautifulsoup4"
            )
        self.parser = parser
    
    def parse(self, content: str, **kwargs) -> ParseResult:
        """
        Parse HTML content.
        
        Args:
            content: HTML content string.
            **kwargs: Additional options.
            
        Returns:
            ParseResult with extracted data.
        """
        if not content:
            return ParseResult(
                success=False,
                errors=["Empty content"],
                raw_content="",
            )
        
        try:
            soup = BeautifulSoup(content, self.parser)
        except Exception as e:
            return ParseResult(
                success=False,
                errors=[f"Failed to parse HTML: {e}"],
                raw_content=content,
            )
        
        data = {}
        warnings = []
        
        # Store raw HTML
        data["raw_html"] = content
        
        # Extract plain text
        text = self._extract_text(soup)
        data["text"] = text
        
        # Extract links
        links = self._extract_links(soup)
        data["links"] = links
        
        # Extract images
        images = self._extract_images(soup)
        data["images"] = images
        
        # Extract tables
        tables = self._extract_tables(soup)
        data["tables"] = tables
        
        # Store the soup object for advanced access
        data["_soup"] = soup
        
        return ParseResult(
            success=True,
            data=data,
            warnings=warnings,
            raw_content=content,
        )
    
    def _extract_text(self, soup: BeautifulSoup) -> str:
        """
        Extract plain text from HTML.
        
        Args:
            soup: BeautifulSoup object.
            
        Returns:
            Plain text content.
        """
        # Remove script and style elements
        for element in soup(["script", "style", "head", "meta", "link"]):
            element.decompose()
        
        # Get text
        text = soup.get_text(separator="\n")
        
        # Clean up whitespace
        lines = [line.strip() for line in text.split("\n")]
        lines = [line for line in lines if line]
        text = "\n".join(lines)
        
        # Collapse multiple newlines
        text = re.sub(r"\n{3,}", "\n\n", text)
        
        return text
    
    def _extract_links(self, soup: BeautifulSoup) -> List[Dict[str, str]]:
        """
        Extract all links from HTML.
        
        Args:
            soup: BeautifulSoup object.
            
        Returns:
            List of link dictionaries with href and text.
        """
        links = []
        
        for anchor in soup.find_all("a", href=True):
            href = anchor["href"]
            text = anchor.get_text(strip=True) or href
            
            # Skip empty or javascript links
            if not href or href.startswith("javascript:"):
                continue
            
            links.append({
                "href": href,
                "text": text,
            })
        
        return links
    
    def _extract_images(self, soup: BeautifulSoup) -> List[Dict[str, str]]:
        """
        Extract all images from HTML.
        
        Args:
            soup: BeautifulSoup object.
            
        Returns:
            List of image dictionaries with src and alt.
        """
        images = []
        
        for img in soup.find_all("img", src=True):
            src = img["src"]
            alt = img.get("alt", "")
            
            images.append({
                "src": src,
                "alt": alt,
            })
        
        return images
    
    def _extract_tables(self, soup: BeautifulSoup) -> List[List[List[str]]]:
        """
        Extract all tables from HTML.
        
        Args:
            soup: BeautifulSoup object.
            
        Returns:
            List of tables, each table is a list of rows,
            each row is a list of cell texts.
        """
        tables = []
        
        for table in soup.find_all("table"):
            table_data = []
            
            for row in table.find_all("tr"):
                row_data = []
                for cell in row.find_all(["td", "th"]):
                    cell_text = cell.get_text(strip=True)
                    row_data.append(cell_text)
                
                if row_data:
                    table_data.append(row_data)
            
            if table_data:
                tables.append(table_data)
        
        return tables
    
    def find_element(
        self,
        content: str,
        tag: str,
        attrs: Optional[Dict[str, str]] = None
    ) -> Optional[str]:
        """
        Find a specific element and return its text.
        
        Args:
            content: HTML content.
            tag: HTML tag name.
            attrs: Optional attributes to match.
            
        Returns:
            Element text, or None if not found.
        """
        soup = BeautifulSoup(content, self.parser)
        element = soup.find(tag, attrs=attrs or {})
        
        if element:
            return element.get_text(strip=True)
        return None
    
    def find_all_elements(
        self,
        content: str,
        tag: str,
        attrs: Optional[Dict[str, str]] = None
    ) -> List[str]:
        """
        Find all matching elements and return their texts.
        
        Args:
            content: HTML content.
            tag: HTML tag name.
            attrs: Optional attributes to match.
            
        Returns:
            List of element texts.
        """
        soup = BeautifulSoup(content, self.parser)
        elements = soup.find_all(tag, attrs=attrs or {})
        
        return [elem.get_text(strip=True) for elem in elements]
    
    def extract_by_selector(
        self,
        content: str,
        selector: str
    ) -> List[str]:
        """
        Extract text using CSS selector.
        
        Args:
            content: HTML content.
            selector: CSS selector string.
            
        Returns:
            List of matching element texts.
        """
        soup = BeautifulSoup(content, self.parser)
        elements = soup.select(selector)
        
        return [elem.get_text(strip=True) for elem in elements]


class TableExtractor:
    """
    Specialized extractor for HTML tables.
    
    Provides more detailed table parsing, including:
    - Header detection
    - Dictionary output (header-value pairs)
    - Specific table selection
    
    Example:
        extractor = TableExtractor()
        tables = extractor.extract_tables_as_dicts(html_content)
        for table in tables:
            for row in table:
                print(row)  # {"Column1": "Value1", "Column2": "Value2"}
    """
    
    def __init__(self, parser: str = "html.parser"):
        if not HAS_BEAUTIFULSOUP:
            raise ImportError("BeautifulSoup required")
        self.parser = parser
    
    def extract_tables_as_dicts(
        self,
        content: str,
        table_index: Optional[int] = None
    ) -> List[List[Dict[str, str]]]:
        """
        Extract tables with first row as headers.
        
        Args:
            content: HTML content.
            table_index: If specified, only extract this table (0-indexed).
            
        Returns:
            List of tables, each table is a list of row dicts.
        """
        soup = BeautifulSoup(content, self.parser)
        tables = soup.find_all("table")
        
        if table_index is not None:
            if 0 <= table_index < len(tables):
                tables = [tables[table_index]]
            else:
                return []
        
        result = []
        
        for table in tables:
            rows = table.find_all("tr")
            if not rows:
                continue
            
            # First row as headers
            headers = []
            header_row = rows[0]
            for cell in header_row.find_all(["th", "td"]):
                headers.append(cell.get_text(strip=True))
            
            if not headers:
                continue
            
            # Data rows
            table_data = []
            for row in rows[1:]:
                cells = row.find_all(["td", "th"])
                row_dict = {}
                
                for i, cell in enumerate(cells):
                    if i < len(headers):
                        row_dict[headers[i]] = cell.get_text(strip=True)
                
                if row_dict:
                    table_data.append(row_dict)
            
            if table_data:
                result.append(table_data)
        
        return result
    
    def find_table_by_header(
        self,
        content: str,
        header_text: str
    ) -> Optional[List[Dict[str, str]]]:
        """
        Find a table that contains a specific header.
        
        Args:
            content: HTML content.
            header_text: Text that should appear in table header.
            
        Returns:
            Table data as list of dicts, or None if not found.
        """
        soup = BeautifulSoup(content, self.parser)
        
        for table in soup.find_all("table"):
            # Check if header text appears in first row
            first_row = table.find("tr")
            if first_row:
                row_text = first_row.get_text(strip=True).lower()
                if header_text.lower() in row_text:
                    # Found matching table
                    return self._table_to_dicts(table)
        
        return None
    
    def _table_to_dicts(self, table) -> List[Dict[str, str]]:
        """Convert a table element to list of dicts."""
        rows = table.find_all("tr")
        if not rows:
            return []
        
        # Get headers
        headers = []
        for cell in rows[0].find_all(["th", "td"]):
            headers.append(cell.get_text(strip=True))
        
        # Get data
        result = []
        for row in rows[1:]:
            cells = row.find_all(["td", "th"])
            row_dict = {}
            for i, cell in enumerate(cells):
                if i < len(headers):
                    row_dict[headers[i]] = cell.get_text(strip=True)
            if row_dict:
                result.append(row_dict)
        
        return result
