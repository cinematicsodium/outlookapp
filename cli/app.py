from __future__ import annotations

import logging

import typer

from .commands import accounts, drafts, folders, messages
from .context import CLIState

app = typer.Typer(
    help="Command-line tools for inspecting Outlook accounts and managing messages.",
    no_args_is_help=True,
    add_completion=False,
)
app.add_typer(accounts.app, name="accounts")
app.add_typer(folders.app, name="folders")
app.add_typer(messages.app, name="messages")
app.add_typer(drafts.app, name="drafts")


def _configure_logging(verbose: bool) -> None:
    """Configure application and Rich logging levels.

    Parameters
    ----------
    verbose : bool
        Enable debug-level logging.

    Returns
    -------
    None
    """
    level = logging.DEBUG if verbose else logging.INFO
    logging.getLogger().setLevel(level)
    logging.getLogger("rich").setLevel(level)


@app.callback()
def main(
    ctx: typer.Context,
    account: str | None = typer.Option(
        None,
        "--account",
        "-m",
        help="Configured account name/address, or shared mailbox SMTP address for drafts.",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose",
        "-v",
        help="Enable verbose logging for Outlook operations.",
    ),
) -> None:
    """Configure global CLI options."""
    _configure_logging(verbose)
    ctx.obj = CLIState(account=account, verbose=verbose)


def run() -> None:
    """Run the Outlook command-line application."""
    app()
