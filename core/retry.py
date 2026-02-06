"""
Retry utilities for Office Automation Foundation.

Provides configurable retry logic with exponential backoff.
Designed for handling transient failures in Outlook COM operations.

Usage:
    from core.retry import retry_operation, RetryConfig
    
    config = RetryConfig(max_attempts=3, base_delay=2.0)
    
    @retry_operation(config)
    def send_email(recipient, subject, body):
        # Outlook COM operation that might fail transiently
        ...
"""

import time
import functools
import logging
from dataclasses import dataclass
from typing import Callable, Optional, TypeVar, Any, Type, Tuple

# Type variable for generic function return types
T = TypeVar("T")

logger = logging.getLogger(__name__)


# Common Outlook COM errors that are worth retrying
TRANSIENT_ERROR_CODES = {
    -2147352567,  # Call was rejected by callee (busy)
    -2147023174,  # RPC server unavailable
    -2147024891,  # Access denied (sometimes transient)
}

# Exception types that indicate transient failures
TRANSIENT_EXCEPTIONS = (
    "pywintypes.com_error",
    "win32com.client.pywintypes.com_error",
)


@dataclass
class RetryConfig:
    """
    Configuration for retry behavior.
    
    Attributes:
        max_attempts: Maximum number of attempts (including first try).
        base_delay: Base delay in seconds between retries.
        max_delay: Maximum delay in seconds (caps exponential growth).
        exponential_base: Base for exponential backoff (default: 2).
        retry_on_exceptions: Tuple of exception types to retry on.
    """
    max_attempts: int = 3
    base_delay: float = 2.0
    max_delay: float = 30.0
    exponential_base: float = 2.0
    retry_on_exceptions: Tuple[Type[Exception], ...] = (Exception,)
    
    def __post_init__(self):
        """Validate configuration."""
        if self.max_attempts < 1:
            self.max_attempts = 1
        if self.base_delay < 0:
            self.base_delay = 0
        if self.max_delay < self.base_delay:
            self.max_delay = self.base_delay


class RetryExhaustedError(Exception):
    """Raised when all retry attempts have been exhausted."""
    
    def __init__(self, message: str, last_exception: Exception, attempts: int):
        super().__init__(message)
        self.last_exception = last_exception
        self.attempts = attempts


def calculate_delay(attempt: int, config: RetryConfig) -> float:
    """
    Calculate delay before next retry using exponential backoff.
    
    Args:
        attempt: Current attempt number (1-indexed).
        config: Retry configuration.
        
    Returns:
        Delay in seconds.
    """
    delay = config.base_delay * (config.exponential_base ** (attempt - 1))
    return min(delay, config.max_delay)


def is_transient_error(exc: Exception) -> bool:
    """
    Check if an exception is likely a transient error worth retrying.
    
    Args:
        exc: The exception to check.
        
    Returns:
        True if the error appears to be transient.
    """
    # Check exception type name (to avoid import issues with pywintypes)
    exc_type_name = f"{type(exc).__module__}.{type(exc).__name__}"
    
    for transient_type in TRANSIENT_EXCEPTIONS:
        if transient_type in exc_type_name:
            # Check if it's a known transient COM error code
            if hasattr(exc, "args") and len(exc.args) > 0:
                error_code = exc.args[0] if isinstance(exc.args[0], int) else None
                if error_code in TRANSIENT_ERROR_CODES:
                    return True
            return True
    
    # Generic network/timeout errors
    error_message = str(exc).lower()
    transient_indicators = [
        "timeout",
        "timed out",
        "temporarily unavailable",
        "connection refused",
        "network",
        "rpc",
        "busy",
    ]
    
    return any(indicator in error_message for indicator in transient_indicators)


def retry_operation(
    config: Optional[RetryConfig] = None,
    on_retry: Optional[Callable[[int, Exception], None]] = None,
) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """
    Decorator for adding retry logic to a function.
    
    Args:
        config: Retry configuration. Uses defaults if None.
        on_retry: Optional callback called before each retry with (attempt, exception).
        
    Returns:
        Decorated function with retry logic.
        
    Example:
        @retry_operation(RetryConfig(max_attempts=3))
        def send_email(to, subject, body):
            outlook.send(to, subject, body)
    """
    if config is None:
        config = RetryConfig()
    
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @functools.wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> T:
            last_exception: Optional[Exception] = None
            
            for attempt in range(1, config.max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                
                except config.retry_on_exceptions as exc:
                    last_exception = exc
                    
                    # Log the failure
                    logger.warning(
                        f"Attempt {attempt}/{config.max_attempts} failed for "
                        f"{func.__name__}: {type(exc).__name__}: {exc}"
                    )
                    
                    # Check if we should retry
                    if attempt >= config.max_attempts:
                        logger.error(
                            f"All {config.max_attempts} attempts exhausted for {func.__name__}"
                        )
                        break
                    
                    # Calculate and apply delay
                    delay = calculate_delay(attempt, config)
                    logger.debug(f"Waiting {delay:.1f}s before retry...")
                    
                    # Call retry callback if provided
                    if on_retry:
                        on_retry(attempt, exc)
                    
                    time.sleep(delay)
            
            # All retries exhausted
            raise RetryExhaustedError(
                f"Operation {func.__name__} failed after {config.max_attempts} attempts",
                last_exception,
                config.max_attempts
            )
        
        return wrapper
    return decorator


def retry_with_backoff(
    operation: Callable[[], T],
    config: Optional[RetryConfig] = None,
    operation_name: str = "operation",
) -> T:
    """
    Execute an operation with retry logic (non-decorator version).
    
    Useful when you can't use decorators or need dynamic configuration.
    
    Args:
        operation: Zero-argument callable to execute.
        config: Retry configuration.
        operation_name: Name for logging purposes.
        
    Returns:
        Result of the operation.
        
    Raises:
        RetryExhaustedError: If all attempts fail.
    """
    if config is None:
        config = RetryConfig()
    
    last_exception: Optional[Exception] = None
    
    for attempt in range(1, config.max_attempts + 1):
        try:
            return operation()
        
        except Exception as exc:
            last_exception = exc
            
            logger.warning(
                f"Attempt {attempt}/{config.max_attempts} failed for "
                f"{operation_name}: {type(exc).__name__}: {exc}"
            )
            
            if attempt >= config.max_attempts:
                break
            
            delay = calculate_delay(attempt, config)
            logger.debug(f"Waiting {delay:.1f}s before retry...")
            time.sleep(delay)
    
    raise RetryExhaustedError(
        f"Operation {operation_name} failed after {config.max_attempts} attempts",
        last_exception,
        config.max_attempts
    )


class RetryContext:
    """
    Context manager for manual retry control.
    
    Useful when you need more control over retry logic.
    
    Example:
        with RetryContext(config) as ctx:
            while ctx.should_retry():
                try:
                    result = do_something()
                    ctx.success()
                    break
                except Exception as e:
                    ctx.record_failure(e)
    """
    
    def __init__(self, config: Optional[RetryConfig] = None):
        self.config = config or RetryConfig()
        self.attempt = 0
        self.last_exception: Optional[Exception] = None
        self._success = False
    
    def __enter__(self) -> "RetryContext":
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        return False
    
    def should_retry(self) -> bool:
        """Check if another attempt should be made."""
        if self._success:
            return False
        return self.attempt < self.config.max_attempts
    
    def record_failure(self, exc: Exception) -> None:
        """Record a failed attempt and wait before retry if needed."""
        self.attempt += 1
        self.last_exception = exc
        
        logger.warning(
            f"Attempt {self.attempt}/{self.config.max_attempts} failed: "
            f"{type(exc).__name__}: {exc}"
        )
        
        if self.attempt < self.config.max_attempts:
            delay = calculate_delay(self.attempt, self.config)
            logger.debug(f"Waiting {delay:.1f}s before retry...")
            time.sleep(delay)
    
    def success(self) -> None:
        """Mark the operation as successful."""
        self._success = True
    
    def raise_if_exhausted(self) -> None:
        """Raise RetryExhaustedError if all attempts failed."""
        if not self._success and self.last_exception:
            raise RetryExhaustedError(
                f"Operation failed after {self.attempt} attempts",
                self.last_exception,
                self.attempt
            )
