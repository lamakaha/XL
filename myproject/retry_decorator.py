import functools
import time
import random
import logging
from typing import Type, Union, List, Callable, Optional, Any

# Setup logging
logger = logging.getLogger(__name__)

def retry(
    exceptions: Union[Type[Exception], List[Type[Exception]]] = Exception,
    tries: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    jitter: float = 0.1,
    logger_func: Optional[Callable[[str], Any]] = None
):
    """
    Retry decorator with exponential backoff for functions that might raise exceptions.

    Specifically designed for xlwings operations that might fail with COM errors.

    Args:
        exceptions: Exception or tuple of exceptions to catch and retry on
        tries: Maximum number of attempts
        delay: Initial delay between retries in seconds
        backoff: Backoff multiplier (e.g. value of 2 will double the delay each retry)
        jitter: Jitter factor to add randomness to the delay (0.0 to 1.0)
        logger_func: Function to use for logging (defaults to logger.warning)

    Returns:
        The decorated function
    """
    if logger_func is None:
        logger_func = logger.warning

    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            mtries, mdelay = tries, delay
            # Collect retry attempt logs
            retry_logs = []

            while mtries > 1:
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    # Calculate jitter amount
                    jitter_amount = random.uniform(0, jitter * mdelay)

                    # Calculate next delay with jitter
                    next_delay = mdelay + jitter_amount

                    # Create a more detailed log message
                    msg = (
                        f"RETRY ATTEMPT: Function '{func.__name__}' failed with error: {e.__class__.__name__}: {e}\n"
                        f"  Args: {args}\n"
                        f"  Kwargs: {kwargs}\n"
                        f"  Retrying in {next_delay:.2f} seconds... ({mtries-1} tries remaining)"
                    )

                    # Add to retry logs
                    retry_logs.append(msg)

                    # Always print to console for visibility
                    print(f"RETRY LOG: {msg}")

                    # Use the logger function
                    if logger_func:
                        logger_func(msg)

                    time.sleep(next_delay)
                    mtries -= 1
                    mdelay *= backoff

            # Last attempt
            try:
                return func(*args, **kwargs)
            except exceptions as e:
                # If we get here, all retries have failed
                # Attach the retry logs to the exception for later use
                e.retry_logs = retry_logs
                raise

        return wrapper

    return decorator


# Specific retry decorator for xlwings COM errors
def retry_xlwings(
    tries: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    jitter: float = 0.1,
    logger_func: Optional[Callable[[str], Any]] = None
):
    """
    Specialized retry decorator for xlwings operations that might fail with COM errors.

    This decorator catches common COM exceptions that occur with xlwings.

    Args:
        tries: Maximum number of attempts
        delay: Initial delay between retries in seconds
        backoff: Backoff multiplier (e.g. value of 2 will double the delay each retry)
        jitter: Jitter factor to add randomness to the delay (0.0 to 1.0)
        logger_func: Function to use for logging (defaults to logger.warning)

    Returns:
        The decorated function
    """
    # Common COM exceptions to catch
    com_exceptions = (
        Exception,  # Using Exception as a fallback since we don't know the exact error types
    )

    return retry(
        exceptions=com_exceptions,
        tries=tries,
        delay=delay,
        backoff=backoff,
        jitter=jitter,
        logger_func=logger_func
    )


# Example usage:
# @retry_xlwings(tries=5, delay=1.0)
# def get_excel_range(sheet, range_address):
#     return sheet.range(range_address).value
