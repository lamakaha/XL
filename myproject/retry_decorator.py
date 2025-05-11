import functools
import time
import random
import logging
import os
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
                    logger_func(msg)

                    time.sleep(next_delay)
                    mtries -= 1
                    mdelay *= backoff

            # Last attempt
            return func(*args, **kwargs)

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
        # COMError is the most common
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


# Function to display recent retry logs
def show_recent_retry_logs(log_file_path='logs/xlwings_retries.log', num_lines=20):
    """
    Display the most recent retry log entries.

    Args:
        log_file_path: Path to the log file
        num_lines: Number of recent lines to display

    Returns:
        The most recent log entries as a string
    """
    try:
        if not os.path.exists(log_file_path):
            return f"Log file not found: {log_file_path}"

        # Read the last num_lines from the log file
        with open(log_file_path, 'r') as f:
            # Read all lines and get the last num_lines
            lines = f.readlines()
            recent_lines = lines[-num_lines:] if len(lines) > num_lines else lines

        return ''.join(recent_lines)
    except Exception as e:
        return f"Error reading log file: {e}"


# Example usage:
# @retry_xlwings(tries=5, delay=1.0)
# def get_excel_range(sheet, range_address):
#     return sheet.range(range_address).value
