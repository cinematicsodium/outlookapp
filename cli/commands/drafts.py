from __future__ import annotations

from pathlib import Path

import typer

from ..console import console
from ..context import abort, create_client

app = typer.Typer(
    help="Create Outlook draft messages from the command line.",
    no_args_is_help=True,
)


@app.command("create")
def create_draft(
    ctx: typer.Context,
    to: str = typer.Option(
        ...,
        "--to",
        help="Primary recipient email address or a semicolon-separated recipient list.",
    ),
    subject: str = typer.Option(
        ...,
        "--subject",
        "-s",
        help="Subject line for the message.",
    ),
    body: str = typer.Option(
        "",
        "--body",
        help="Plain-text message body.",
    ),
    cc: str | None = typer.Option(
        None,
        "--cc",
        help="Optional CC recipients separated by commas or semicolons.",
    ),
    bcc: str | None = typer.Option(
        None,
        "--bcc",
        help="Optional BCC recipients separated by commas or semicolons.",
    ),
    attachment: list[Path] | None = typer.Option(  # ruff: ignore[B008]
        None,
        "--attachment",
        "-a",
        help="Path to an attachment. Repeat the option to attach multiple files.",
    ),
    display: bool = typer.Option(
        False,
        "--display",
        help="Open the Outlook compose window after creating the message.",
    ),
    send: bool = typer.Option(
        False,
        "--send",
        help="Send the message immediately instead of leaving it as a draft.",
    ),
) -> None:
    """Create a draft email, optionally display it, and optionally send it."""
    client = create_client(ctx)
    message = client.new_email()
    if message is None:
        abort("Unable to create a new Outlook message.")
    message.to = to
    message.subject = subject
    message.body = body
    if cc:
        message.cc = cc
    if bcc:
        message.bcc = bcc
    if attachment:
        message.add_attachments(attachment)
    if send:
        message.send()
        console.log("[green]Message sent.[/green]")
        return
    message.save()
    if display:
        message.show()
    console.log("[green]Draft created.[/green]")
