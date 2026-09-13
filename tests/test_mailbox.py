from types import SimpleNamespace

import pytest
from typer.testing import CliRunner

from outlook import Outlook, OutlookError
from outlook.cli import context
from outlook.cli.app import app as cli_app
from outlook.enums import FolderEnum
from outlook.services import client as connection
from outlook.tests.test_model_navigation import (
    FakeAccount,
    FakeCollection,
    FakeFolder,
    FakeMailItem,
)


class COMFailure(Exception):
    """Simulate a pywintypes COM failure without requiring Windows."""


def test_shared_mailbox_is_verified_and_used_as_from(monkeypatch) -> None:
    """Verify a mailbox outside Accounts is accessible and becomes From."""
    raw = FakeMailItem("draft")
    inbox = FakeFolder("Group Inbox")
    recipient = SimpleNamespace(Resolve=lambda: True)
    lookups = []

    def shared_folder(resolved, kind):
        """Record the mailbox and folder requested for access verification."""
        lookups.append((resolved, kind))
        return inbox

    mapi = SimpleNamespace(
        Accounts=FakeCollection(
            [
                FakeAccount("Personal", "personal@example.com", FakeFolder("Root")),
                FakeAccount("Second", "second@example.com", FakeFolder("Root")),
            ]
        ),
        CreateRecipient=lambda address: (
            recipient if address == "group@example.com" else None
        ),
        GetSharedDefaultFolder=shared_folder,
    )
    client = Outlook(
        address=" Group@Example.com ",
        app=SimpleNamespace(CreateItem=lambda kind: raw),
        mapi=mapi,
    )

    assert client.account is None  # Shared mailboxes need not be profile accounts.
    assert client.address == "group@example.com"
    assert lookups == [(recipient, FolderEnum.INBOX)]
    assert client.new_email().sent_for == "group@example.com"
    monkeypatch.setattr(
        context,
        "Outlook",
        lambda address: Outlook(
            address=address, app=SimpleNamespace(CreateItem=lambda kind: raw), mapi=mapi
        ),
    )
    result = CliRunner().invoke(
        cli_app,
        [
            "--account",
            "group@example.com",
            "drafts",
            "create",
            "--to",
            "a@example.com",
            "--subject",
            "Team update",
        ],
    )
    assert result.exit_code == 0, result.output
    assert raw.SentOnBehalfOfName == "group@example.com"
    assert raw.saved == 1


@pytest.mark.parametrize("resolved", [True, False])
def test_inaccessible_shared_mailbox_is_rejected(monkeypatch, resolved: bool) -> None:
    """Reject unresolved recipients and denied Inbox access before composition."""
    monkeypatch.setattr(connection, "COM_ERRORS", (AttributeError, COMFailure))

    def denied(*args):
        """Simulate Exchange denying access to the shared Inbox."""
        raise COMFailure("access denied")

    mapi = SimpleNamespace(
        Accounts=FakeCollection(
            [FakeAccount("Personal", "personal@example.com", FakeFolder("Root"))]
        ),
        CreateRecipient=lambda address: SimpleNamespace(Resolve=lambda: resolved),
        GetSharedDefaultFolder=denied,
    )
    with pytest.raises(OutlookError, match="not accessible|could not be resolved"):
        Outlook(address="group@example.com", app=SimpleNamespace(), mapi=mapi)


def test_configured_account_sets_transport_and_verifies_inbox() -> None:
    """Verify explicit profile selection uses the selected transport account."""
    accounts = [
        FakeAccount("Personal", "personal@example.com", FakeFolder("Root")),
        FakeAccount("Group", "group@example.com", FakeFolder("Group Inbox")),
    ]
    raw = FakeMailItem("draft")
    client = Outlook(
        address="Group",
        app=SimpleNamespace(CreateItem=lambda kind: raw),
        mapi=SimpleNamespace(Accounts=FakeCollection(accounts)),
    )
    assert accounts[1].DeliveryStore.calls == 1
    client.new_email()
    assert raw.SendUsingAccount is accounts[1]
    assert raw.SentOnBehalfOfName == "group@example.com"


def test_none_keeps_single_account_selection_and_rejects_ambiguous_sending() -> None:
    """Verify omitted addresses preserve account selection without guessing."""
    accounts = [FakeAccount("Personal", "personal@example.com", FakeFolder("Root"))]
    raw = FakeMailItem("draft")
    app = SimpleNamespace(CreateItem=lambda kind: raw)
    client = Outlook(app=app, mapi=SimpleNamespace(Accounts=FakeCollection(accounts)))
    client.new_email()
    assert raw.SendUsingAccount is accounts[0]
    accounts.append(FakeAccount("Second", "second@example.com", FakeFolder("Root")))
    client = Outlook(app=app, mapi=SimpleNamespace(Accounts=FakeCollection(accounts)))
    with pytest.raises(OutlookError, match="No Outlook account"):
        client.new_email()


@pytest.mark.parametrize("address", ["", " ", "a@example.com;b@example.com", "invalid"])
def test_invalid_sending_address_is_rejected(address: str) -> None:
    """Reject empty, malformed, and multiple sender addresses."""
    mapi = SimpleNamespace(
        Accounts=FakeCollection(
            [FakeAccount("Personal", "personal@example.com", FakeFolder("Root"))]
        )
    )
    with pytest.raises(OutlookError):
        Outlook(address=address, app=SimpleNamespace(), mapi=mapi)


def test_connection_errors_are_wrapped(monkeypatch) -> None:
    """Verify COM failures become public OutlookError exceptions."""
    monkeypatch.setattr(connection, "COM_ERRORS", (AttributeError, COMFailure))

    def fail(*args):
        """Raise a simulated Outlook COM failure."""
        raise COMFailure("server unavailable")

    monkeypatch.setattr(
        connection, "_load_dispatch", lambda: SimpleNamespace(Dispatch=fail)
    )
    with pytest.raises(OutlookError, match="Unable to connect"):
        connection._connect()
    with pytest.raises(OutlookError, match="MAPI namespace"):
        connection._open_mapi(SimpleNamespace(GetNamespace=fail))
