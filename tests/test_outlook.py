from types import SimpleNamespace

import pytest
from typer.testing import CliRunner

import outlook
from outlook import Outlook, OutlookError
from outlook.cli.app import app
from outlook.enums import ItemType
from outlook.services.client import _select_account
from outlook.tests.test_model_navigation import FakeAccount, FakeCollection, FakeFolder


def _accounts() -> list[outlook.Account]:
    """Create two fake Outlook accounts.

    Parameters
    ----------
    None

    Returns
    -------
    list of outlook.Account
        Fake accounts sharing a root folder.
    """
    root = FakeFolder("Root")
    return [
        outlook.Account(FakeAccount("First", "first@example.com", root)),
        outlook.Account(FakeAccount("Second", "second@example.com", root)),
    ]


def test_public_package_and_cli_help_import_normally() -> None:
    """Verify public imports and CLI help work.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    assert outlook.Outlook is Outlook
    result = CliRunner().invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "accounts" in result.stdout


def test_account_selection_handles_multiple_accounts_without_guessing() -> None:
    """Verify ambiguous accounts require an explicit selection.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    accounts = _accounts()
    assert _select_account(accounts) is None
    assert _select_account(accounts, "SECOND") is accounts[1]
    with pytest.raises(OutlookError):
        _select_account(accounts, "missing@example.com")


def test_close_releases_cached_com_wrappers() -> None:
    """Verify closing clears cached Outlook COM wrappers.

    Parameters
    ----------
    None

    Returns
    -------
    None
    """
    account_items = [
        FakeAccount("Only", "only@example.com", FakeFolder("Root")),
    ]
    namespace = SimpleNamespace(
        Accounts=FakeCollection(account_items),
        Class=ItemType.NAMESPACE,
    )
    client = Outlook(app=SimpleNamespace(), mapi=namespace)
    assert client.accounts

    client.close()

    assert client.account is None
    assert "accounts" not in client.__dict__
    with pytest.raises(OutlookError):
        _ = client.accounts


def test_cli_account_totals_do_not_enumerate_child_folders(monkeypatch) -> None:
    """Verify account totals include unopened entries without loading them.

    Parameters
    ----------
    monkeypatch : pytest.MonkeyPatch
        Temporary replacement of the CLI connection factory.
    """
    from outlook.cli.commands import accounts

    root = FakeFolder("Root", subfolders=[object() for _ in range(1000)])
    account = outlook.Account(FakeAccount("Only", "only@example.com", root))
    client = SimpleNamespace(accounts=[account], account=account)
    monkeypatch.setattr(accounts, "create_client", lambda ctx: client)
    result = CliRunner().invoke(app, ["accounts", "list"])
    assert result.exit_code == 0, result.output
    assert "1000" in result.output
    assert root.Folders.item_calls == 0
