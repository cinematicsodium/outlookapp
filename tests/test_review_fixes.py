from io import StringIO
from types import SimpleNamespace

import pytest
from rich.console import Console

from outlook import MailItem, OutlookError, utils
from outlook.cli import rendering
from outlook.models import mail_item
from outlook.tests.test_mailbox import COMFailure
from outlook.tests.test_model_navigation import FakeCollection, FakeMailItem


def test_message_without_delegated_sender_can_send() -> None:
    """Verify normal drafts can use Outlook's default sending identity."""
    raw = FakeMailItem("draft")
    raw.SentOnBehalfOfName = ""
    sent = []
    raw.Send = lambda: sent.append(True)
    MailItem(raw).send()
    assert sent == [True]
    raw.To = ""
    with pytest.raises(ValueError, match="recipient is missing"):
        MailItem(raw).send()
    assert sent == [True]


def test_directory_rejection_precedes_all_attachment_additions(tmp_path) -> None:
    """Verify a directory prevents even valid preceding files being attached."""
    file = tmp_path / "report.txt"
    file.write_text("report")
    raw = FakeMailItem("draft")
    with pytest.raises(OutlookError, match="not an existing file"):
        MailItem(raw).add_attachments([file, tmp_path])
    assert raw.Attachments.Count == 0
    MailItem(raw).add_attachments(file)
    assert raw.Attachments.items == [str(file)]


def test_table_preserves_literal_email_markup(monkeypatch) -> None:
    """Verify bracketed email text is displayed verbatim without crashing."""
    output = StringIO()
    monkeypatch.setattr(rendering, "console", Console(file=output, width=120))
    subject = "[red]Invoice[/red] [/missing]"
    rendering.echo_table([(subject,)], ["Subject"])
    assert subject in output.getvalue()


def test_export_returns_false_on_com_failure(monkeypatch, tmp_path) -> None:
    """Verify Outlook SaveAs errors follow the export failure contract."""
    monkeypatch.setattr(mail_item, "COM_ERRORS", (AttributeError, COMFailure))
    raw = FakeMailItem("draft")

    def fail(path):
        """Simulate Outlook rejecting a save operation."""
        raise COMFailure("save denied")

    raw.SaveAs = fail
    assert MailItem(raw).export(tmp_path / "message.msg") is False


@pytest.mark.parametrize("exchange_available", [True, False])
def test_smtp_address_falls_back_after_com_failures(
    monkeypatch, exchange_available
) -> None:
    """Verify Exchange and MAPI-property failures retain the plain address fallback."""
    monkeypatch.setattr(utils, "COM_ERRORS", (AttributeError, COMFailure))

    def fail(*args):
        """Simulate an inaccessible COM method or property."""
        raise COMFailure("property unavailable")

    class ExchangeUser:
        @property
        def PrimarySmtpAddress(self):
            """Simulate an inaccessible Exchange primary address."""
            return fail()

    user = SimpleNamespace(
        GetExchangeUser=(lambda: ExchangeUser()) if exchange_available else fail,
        PropertyAccessor=SimpleNamespace(GetProperty=fail),
        Address="Sender@Example.com",
    )
    assert utils.get_smtp_address(user) == "sender@example.com"


def test_export_returns_false_when_parent_cannot_be_created(tmp_path) -> None:
    """Verify filesystem errors follow the same export failure contract."""
    parent = tmp_path / "file"
    parent.write_text("not a directory")
    assert MailItem(FakeMailItem("draft")).export(parent / "message.msg") is False


def test_message_table_counts_attachments_without_loading_filenames() -> None:
    """Verify attachment totals require no attachment item lookups."""
    raw = FakeMailItem("draft")
    raw.Attachments = FakeCollection([object() for _ in range(100)])
    assert "100 file(s)" in MailItem(raw).as_table()
    assert raw.Attachments.item_calls == 0
