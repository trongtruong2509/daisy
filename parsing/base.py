"""
Base parser interface for email content parsing.

Defines the contract that all parsers must follow.
Designed for extensibility without business-specific assumptions.

Usage:
    class MyCustomParser(BaseParser):
        def parse(self, content: str, **kwargs) -> ParseResult:
            # Custom parsing logic
            return ParseResult(success=True, data={"key": "value"})
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class ParseResult:
    """
    Result of a parsing operation.
    
    Attributes:
        success: Whether parsing completed successfully.
        data: Extracted data (structure depends on parser).
        errors: List of error messages.
        warnings: List of warning messages.
        raw_content: Original content that was parsed.
        metadata: Additional metadata about the parse operation.
    """
    success: bool
    data: Dict[str, Any] = field(default_factory=dict)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    raw_content: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        Get a value from the parsed data.
        
        Args:
            key: Key to look up.
            default: Default value if not found.
            
        Returns:
            Value from data, or default.
        """
        return self.data.get(key, default)
    
    def has_errors(self) -> bool:
        """Check if there were any errors."""
        return len(self.errors) > 0
    
    def has_warnings(self) -> bool:
        """Check if there were any warnings."""
        return len(self.warnings) > 0


class BaseParser(ABC):
    """
    Abstract base class for content parsers.
    
    All custom parsers should inherit from this class and implement
    the parse() method.
    
    Example:
        class EmployeeIdParser(BaseParser):
            def parse(self, content: str, **kwargs) -> ParseResult:
                # Extract employee ID from content
                match = re.search(r'Employee ID: (\\d+)', content)
                if match:
                    return ParseResult(
                        success=True,
                        data={"employee_id": match.group(1)}
                    )
                return ParseResult(success=False, errors=["Employee ID not found"])
    """
    
    @abstractmethod
    def parse(self, content: str, **kwargs) -> ParseResult:
        """
        Parse content and extract data.
        
        Args:
            content: The content to parse (text or HTML).
            **kwargs: Additional parser-specific parameters.
            
        Returns:
            ParseResult with extracted data.
        """
        pass
    
    def parse_email(self, email: "Email", use_html: bool = False, **kwargs) -> ParseResult:
        """
        Parse an email object.
        
        Convenience method that extracts the appropriate body content
        from an Email object and parses it.
        
        Args:
            email: Email object to parse.
            use_html: If True, parse HTML body; otherwise plain text.
            **kwargs: Additional parser parameters.
            
        Returns:
            ParseResult from parsing.
        """
        content = email.body_html if use_html else email.body_text
        result = self.parse(content, **kwargs)
        
        # Add email metadata
        result.metadata["email_subject"] = email.subject
        result.metadata["email_sender"] = email.sender_address
        result.metadata["email_id"] = email.entry_id
        
        return result


class ChainedParser(BaseParser):
    """
    Parser that chains multiple parsers together.
    
    Runs each parser in sequence and merges results.
    Useful for extracting different types of data from the same content.
    
    Example:
        parser = ChainedParser([
            DateParser(),
            AmountParser(),
            ReferenceNumberParser()
        ])
        result = parser.parse(email_body)
    """
    
    def __init__(self, parsers: List[BaseParser]):
        """
        Initialize with a list of parsers.
        
        Args:
            parsers: List of parsers to run in sequence.
        """
        self.parsers = parsers
    
    def parse(self, content: str, **kwargs) -> ParseResult:
        """
        Run all parsers and merge results.
        
        Args:
            content: Content to parse.
            **kwargs: Passed to all parsers.
            
        Returns:
            Merged ParseResult from all parsers.
        """
        merged_data = {}
        all_errors = []
        all_warnings = []
        overall_success = True
        
        for parser in self.parsers:
            result = parser.parse(content, **kwargs)
            
            # Merge data (later parsers override earlier)
            merged_data.update(result.data)
            all_errors.extend(result.errors)
            all_warnings.extend(result.warnings)
            
            if not result.success:
                overall_success = False
        
        return ParseResult(
            success=overall_success,
            data=merged_data,
            errors=all_errors,
            warnings=all_warnings,
            raw_content=content,
        )


class ConditionalParser(BaseParser):
    """
    Parser that applies different parsers based on content conditions.
    
    Example:
        parser = ConditionalParser({
            lambda c: "Invoice" in c: InvoiceParser(),
            lambda c: "Receipt" in c: ReceiptParser(),
        }, default_parser=GenericParser())
    """
    
    def __init__(
        self,
        condition_map: Dict[callable, BaseParser],
        default_parser: Optional[BaseParser] = None
    ):
        """
        Initialize with condition-parser mapping.
        
        Args:
            condition_map: Dict mapping condition functions to parsers.
            default_parser: Parser to use if no condition matches.
        """
        self.condition_map = condition_map
        self.default_parser = default_parser
    
    def parse(self, content: str, **kwargs) -> ParseResult:
        """
        Parse using appropriate parser based on conditions.
        
        Args:
            content: Content to parse.
            **kwargs: Passed to selected parser.
            
        Returns:
            ParseResult from selected parser.
        """
        for condition, parser in self.condition_map.items():
            try:
                if condition(content):
                    result = parser.parse(content, **kwargs)
                    result.metadata["parser_type"] = type(parser).__name__
                    return result
            except Exception:
                continue
        
        if self.default_parser:
            result = self.default_parser.parse(content, **kwargs)
            result.metadata["parser_type"] = "default"
            return result
        
        return ParseResult(
            success=False,
            errors=["No matching parser found for content"],
            raw_content=content,
        )
