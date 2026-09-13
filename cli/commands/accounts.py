from __future__ import annotations

import typer

from ..context import create_client, get_state
from ..rendering import echo_table

app = typer.Typer(
    help="Inspect Outlook accounts that are visible to the current Outlook profile.",
    no_args_is_help=True,
)


@app.command("list")
def list_accounts(ctx: typer.Context) -> None:
    """List available Outlook accounts."""
    client = create_client(ctx)
    selected_account = get_state(ctx).account
    default_account = client.account
    rows: list[tuple[object, ...]] = []
    for account in client.accounts:
        matches_global_selection = account.matches(selected_account)
        is_default_account = default_account is not None and account.matches(
            default_account.email_address
        )
        root = account.root_folder
        rows.append(
            (
                account.name,
                account.email_address,
                root.subfolder_count if root is not None else 0,
                "yes" if matches_global_selection or is_default_account else "",
            )
        )
    echo_table(rows, headers=["Name", "Address", "Folders", "Selected"])
