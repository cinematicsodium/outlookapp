import re
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import Any

from .exceptions import OutlookError

_EMAIL = re.compile(
    r"[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@"
    r"(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,63}",
    re.IGNORECASE,
)
_RECIPIENT_SEPARATOR = re.compile(r"\s*[;,]\s*")


def validate_datetime(value: Any) -> datetime | None:
    """Validate and normalize a datetime value.

    Parameters
    ----------
    value : Any
        A datetime, an ISO 8601 string, or ``None``.

    Returns
    -------
    datetime or None
        The supplied datetime, a parsed datetime, or ``None`` when the value
        cannot be parsed.
    """
    if value is None or isinstance(value, datetime):
        return value
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def validate_email(emails: str | Iterable[str] | None) -> str:
    """Validate and normalize email recipients.

    Parameters
    ----------
    emails : str, iterable of str, or None
        Email addresses separated by commas or semicolons, or an iterable of
        addresses.

    Returns
    -------
    str
        Lowercase addresses separated by semicolons for Outlook.

    Raises
    ------
    OutlookError
        If any address is invalid.
    """
    if not emails:
        return ""
    values = [emails] if isinstance(emails, str) else list(emails)
    parsed = [
        address.strip().lower()
        for value in values
        for address in _RECIPIENT_SEPARATOR.split(str(value))
        if address.strip()
    ]
    invalid = [address for address in parsed if _EMAIL.fullmatch(address) is None]
    if invalid:
        raise OutlookError(f"Invalid emails: {invalid}")
    return "; ".join(parsed)


def validate_paths(paths: str | Path | Iterable[str | Path]) -> list[Path]:
    """Resolve existing file paths for attachments.

    Parameters
    ----------
    paths : str, Path, or iterable of str or Path
        One or more file paths to validate.

    Returns
    -------
    list of Path
        Expanded, absolute paths in input order.

    Raises
    ------
    OutlookError
        If any value is not path-like or does not identify an existing file.
    """
    values = [paths] if isinstance(paths, (str, Path)) else list(paths)
    valid: list[Path] = []
    errors: list[str] = []
    for value in values:
        if not isinstance(value, (str, Path)):
            errors.append(f"Invalid path: {value} (type: {type(value)})")
            continue
        path = Path(value).expanduser().resolve()
        if path.is_file():
            valid.append(path)
        else:
            errors.append(f"Path is not an existing file: {path}")
    if errors:
        raise OutlookError("Path validation errors:\n" + "\n".join(errors))
    return valid
