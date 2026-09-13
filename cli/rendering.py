from __future__ import annotations

from datetime import datetime

from tabulate import tabulate

from .console import console


def echo_table(rows: list[tuple[object, ...]], headers: list[str]) -> None:
    """Print rows as a GitHub-style table.

    Parameters
    ----------
    rows : list of tuple
        Table rows to render.
    headers : list of str
        Column headings.
    """
    if not rows:
        console.print("No results.")
        return
    console.print(tabulate(rows, headers=headers, tablefmt="github"), markup=False)


def format_datetime(value: datetime | None) -> str:
    """Format a datetime for compact CLI display.

    Parameters
    ----------
    value : datetime or None
        Datetime to format.

    Returns
    -------
    str
        Formatted local value, an empty string, or a string fallback.
    """
    if value is None:
        return ""
    return value.strftime("%Y-%m-%d %H:%M")


def truncate(value: str, width: int = 80) -> str:
    """Collapse whitespace and truncate text to a display width.

    Parameters
    ----------
    value : str
        Text to normalize.
    width : int, default=80
        Maximum result length.

    Returns
    -------
    str
        Normalized text, ending in an ellipsis when truncated.
    """
    normalized_value = " ".join(value.split())
    if len(normalized_value) <= width:
        return normalized_value
    return f"{normalized_value[: width - 1].rstrip()}…"
