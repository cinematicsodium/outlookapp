from __future__ import annotations

import logging
from collections.abc import Iterable
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, TypedDict

from ..enums import ItemType
from ..exceptions import COM_ERRORS
from ..protocols import OlMailItem
from ..types import StrPath
from ..utils import get_smtp_address, unpack_collection
from ..validation import validate_datetime, validate_email, validate_paths
from .address_entry import AddressEntry
from .base import ItemModel

log = logging.getLogger(__name__)
if TYPE_CHECKING:
    from .folder import Folder


class Recipient(TypedDict):
    """Name and SMTP address for a message recipient."""

    name: str
    address: str


class MailItem(ItemModel):
    """Represent an Outlook email message.

    Parameters
    ----------
    item : OlMailItem
        Outlook mail item COM object to wrap.
    """

    _UPDATABLE_FIELDS = frozenset(
        {
            "sent_for",
            "to",
            "cc",
            "bcc",
            "subject",
            "body",
            "html",
            "deliver_at",
            "unread",
        }
    )
    item_type = ItemType.MAIL_ITEM
    required_properties = (
        "SenderEmailAddress",
        "SentOnBehalfOfName",
        "To",
        "CC",
        "BCC",
        "Subject",
        "Body",
        "HTMLBody",
        "Attachments",
        "DeferredDeliveryTime",
        "SentOn",
        "ReceivedTime",
        "ConversationID",
    )
    inaccessible_error_message = "Provided Outlook item is not an accessible mail item."

    def __init__(self, item: OlMailItem):
        """Initialize a mail-item wrapper.

        Parameters
        ----------
        item : OlMailItem
            Outlook mail-item COM object.

        Returns
        -------
        None
        """
        super().__init__(item)
        self._item = item

    # Identity
    @property
    def id(self) -> str:
        """The unique hexadecimal identifier for the Outlook item."""
        return self._item.EntryID or ""

    @property
    def thread_id(self) -> str:
        """A unique string identifying all messages within the same thread."""
        return self._item.ConversationID or ""

    @property
    def thread_index(self) -> str:
        """A hexadecimal string representing the message's hierarchical position in a thread."""
        return self._item.ConversationIndex or ""

    @property
    def folder(self) -> Folder | None:
        """Returns the Folder containing the mail item."""
        from .folder import Folder

        return Folder.from_outlook_item(self._item.Parent)

    # Addresses
    @property
    def sender_entry(self) -> AddressEntry | None:
        """Returns an AddressEntry object representing the email sender."""
        return AddressEntry.from_outlook_item(self._item.Sender)

    @property
    def sender_name(self) -> str:
        """Returns the display name of the email sender."""
        return str(self._item.SenderName or "")

    @property
    def sender_address(self) -> str:
        """Retrieves the sender's email address in lowercase format."""
        address = get_smtp_address(self._item.Sender)
        return str(address or self._item.SenderEmailAddress or "").lower()

    @property
    def sent_for(self) -> str:
        """Gets or sets the name or address of the person the email is sent on behalf of."""
        return self._item.SentOnBehalfOfName or ""

    @sent_for.setter
    def sent_for(self, value: str) -> None:
        """Set the sending identity.

        Parameters
        ----------
        value : str
            Sender name or email address.

        Returns
        -------
        None
        """
        self._item.SentOnBehalfOfName = validate_email(value) if value else ""

    @property
    def to(self) -> str:
        """Gets or sets the primary recipients of the email message."""
        return self._item.To or ""

    @to.setter
    def to(self, value: str | Iterable[str]) -> None:
        """Set primary recipients.

        Parameters
        ----------
        value : str or iterable of str
            Recipient addresses.

        Returns
        -------
        None
        """
        self._item.To = validate_email(value) if value else ""

    @property
    def cc(self) -> str:
        """Gets or sets the carbon copy recipients of the email message."""
        return self._item.CC or ""

    @cc.setter
    def cc(self, value: str | Iterable[str]) -> None:
        """Set carbon-copy recipients.

        Parameters
        ----------
        value : str or iterable of str
            Recipient addresses.

        Returns
        -------
        None
        """
        self._item.CC = validate_email(value) if value else ""

    @property
    def bcc(self) -> str:
        """Gets or sets the blind carbon copy recipients of the email message."""
        return self._item.BCC or ""

    @bcc.setter
    def bcc(self, value: str | Iterable[str]) -> None:
        """Set blind-carbon-copy recipients.

        Parameters
        ----------
        value : str or iterable of str
            Recipient addresses.

        Returns
        -------
        None
        """
        self._item.BCC = validate_email(value) if value else ""

    @property
    def delivery_recipients(self):
        """Return the To, CC, and BCC recipient strings."""
        return self.to, self.cc, self.bcc

    @property
    def recipients(self) -> list[Recipient]:
        """Return names and SMTP addresses for all resolved recipients."""
        entries = (
            AddressEntry.from_outlook_item(recipient.AddressEntry)
            for recipient in unpack_collection(self._item.Recipients)
            if recipient is not None
        )
        return [
            Recipient(name=entry.name, address=entry.email_address)
            for entry in entries
            if entry is not None
        ]

    # Content
    @property
    def subject(self) -> str:
        """Gets or sets the subject line for the email message."""
        return self._item.Subject or ""

    @subject.setter
    def subject(self, value: str) -> None:
        """Set the message subject.

        Parameters
        ----------
        value : str
            Subject text.

        Returns
        -------
        None
        """
        self._item.Subject = value or ""

    @property
    def body(self) -> str:
        """Gets or sets the plain text content of the email body."""
        return self._item.Body

    @body.setter
    def body(self, value: str) -> None:
        """Set the plain-text message body.

        Parameters
        ----------
        value : str
            Body text.

        Returns
        -------
        None
        """
        self._item.Body = value or ""

    @property
    def html(self) -> str:
        """Gets or sets HTML formatted content of the email body."""
        return self._item.HTMLBody

    @html.setter
    def html(self, value: str) -> None:
        """Set the HTML message body.

        Parameters
        ----------
        value : str
            HTML text containing an ``<html>`` element.

        Returns
        -------
        None
        """
        value = value or ""
        if "html" not in value.lower():
            raise ValueError("HTML body must contain an <html> root element.")
        self._item.HTMLBody = value

    @property
    def attachments(self) -> list[str]:
        """Gets a list of filenames for all current attachments."""
        return [
            str(attachment.FileName)
            for attachment in unpack_collection(self._item.Attachments)
            if attachment is not None
        ]

    def add_attachments(self, paths: str | Path | Iterable[str | Path]) -> None:
        """Attach one or more files to the message.

        Parameters
        ----------
        paths : str, Path, or iterable of str or Path
            Existing files to attach.

        Raises
        ------
        OutlookError
            If any supplied path is invalid or does not exist.
        """
        for path in validate_paths(paths):
            self._item.Attachments.Add(str(path))

    # Timing and status
    @property
    def deliver_at(self) -> datetime | None:
        """Gets or sets the date and time for deferred delivery."""
        return self._item.DeferredDeliveryTime

    @deliver_at.setter
    def deliver_at(self, value: datetime) -> None:
        """Set the deferred-delivery timestamp.

        Parameters
        ----------
        value : datetime
            Deferred-delivery time.

        Returns
        -------
        None
        """
        dt = validate_datetime(value)
        if dt:
            self._item.DeferredDeliveryTime = dt

    @property
    def sent_at(self) -> datetime | None:
        """Returns when the email message was sent."""
        return self._item.SentOn

    @property
    def received_at(self) -> datetime | None:
        """Returns when the email message was received."""
        return self._item.ReceivedTime

    @property
    def size(self) -> int:
        """Retrieves the total size of the email item in bytes."""
        return self._item.Size

    @property
    def size_mb(self) -> float:
        """Return the message size in mebibytes, rounded to two decimals."""
        return round(self.size / (1024**2), 2)

    @property
    def unread(self) -> bool:
        """Gets or sets whether the email message is unread."""
        return self._item.UnRead

    @unread.setter
    def unread(self, value: bool) -> None:
        """Set the message unread state.

        Parameters
        ----------
        value : bool
            Whether the message is unread.

        Returns
        -------
        None
        """
        self._item.UnRead = value

    # Actions
    def show(self) -> None:
        """Open the message in an Outlook inspector window."""
        self._item.Display()

    def send(self) -> None:
        """Resolve recipients and send the message.

        Raises
        ------
        ValueError
            If all delivery recipients are missing.
        """
        if not any(self.delivery_recipients):
            raise ValueError(f"Cannot send {self.subject!r}: recipient is missing.")
        if self._item.Recipients.ResolveAll() is False:
            log.warning("Sending %r with unresolved recipients", self.subject)
        self._item.Send()

    def save(self) -> None:
        """Save the message in Outlook."""
        self._item.Save()

    def export(self, path: StrPath) -> bool:
        """Export the message to a file.

        Parameters
        ----------
        path : str or Path
            Destination file path. Missing parent directories are created.

        Returns
        -------
        bool
            ``True`` when Outlook saves the message successfully.
        """
        try:
            target = Path(path).expanduser().resolve()
            target.parent.mkdir(parents=True, exist_ok=True)
            self._item.SaveAs(str(target))
            return True
        except (OSError, *COM_ERRORS):
            log.exception("Failed to export %r to %r", self.subject, path)
            return False

    def delete(self) -> None:
        """Delete the message from Outlook."""
        self._item.Delete()

    def move(self, folder: Folder) -> MailItem | None:
        """Move the message to another folder.

        Parameters
        ----------
        folder : Folder
            Destination folder.

        Returns
        -------
        MailItem or None
            The moved message, if Outlook returns an accessible object.
        """
        item = self._item.Move(folder._ol_folder_item)
        return MailItem.from_outlook_item(item)

    def update(self, **kwargs: Any) -> None:
        """Set supported fields and save the message.

        Parameters
        ----------
        **kwargs : Any
            Message fields and their new values. Supported fields are
            ``sent_for``, ``to``, ``cc``, ``bcc``, ``subject``, ``body``,
            ``html``, ``deliver_at``, and ``unread``.

        Raises
        ------
        ValueError
            If any field is unsupported or a property rejects its value.
        """
        unsupported = [key for key in kwargs if key not in self._UPDATABLE_FIELDS]
        if unsupported:
            raise ValueError(f"Unsupported fields: {unsupported}")
        for key, value in kwargs.items():
            setattr(self, key, value)
        self.save()

    # Serialization
    def as_dict(self) -> dict[str, Any]:
        """Return the message's serializable fields as a dictionary."""
        data = {
            "thread_index": self.thread_index,
            "thread_id": self.thread_id,
            "sender_address": self.sender_address,
            "sent_for": self.sent_for,
            "to": self.to,
            "cc": self.cc,
            "bcc": self.bcc,
            "recipients": self.recipients,
            "subject": self.subject,
            "body": self.body,
            "attachments": self.attachments,
            "deliver_at": self.deliver_at,
            "sent_at": self.sent_at,
            "received_at": self.received_at,
            "size": self.size,
        }
        return {key: self._format(value) for key, value in data.items()}

    def as_table(
        self,
        format: str = "github",
        width: int | None = 100,
        body_limit: int | None = 80,
    ) -> str:
        """Render message details as a two-column table.

        Parameters
        ----------
        format : str, default="github"
            Table format accepted by :func:`tabulate.tabulate`.
        width : int or None, default=100
            Maximum rendered column width.
        body_limit : int or None, default=80
            Maximum number of body characters, or ``None`` for no limit.

        Returns
        -------
        str
            Rendered table text.
        """
        folder = self.folder
        from tabulate import tabulate

        attachments = self.attachments
        size = self.size
        rows = {
            "Subject": self.subject,
            "Sender": self.sender_address,
            "To": self.to,
            "CC": self.cc,
            "BCC": self.bcc,
            "Sent For": self.sent_for,
            "Deliver At": self._format_dt(self.deliver_at),
            "Sent At": self._format_dt(self.sent_at),
            "Received At": self._format_dt(self.received_at),
            "Body": self._format_body(self.body, body_limit),
            "Attachments": f"{len(attachments)} file(s)",
            "Size": f"{size} bytes ({round(size / (1024**2), 2)} MB)",
            "Folder": folder.name if folder else "",
            "Thread ID": self.thread_id,
            "Thread Index": self.thread_index,
            "ID": self.id,
        }
        return tabulate(
            rows.items(),
            tablefmt=format,
            maxcolwidths=width,
        ).strip()

    # Internals
    @staticmethod
    def _format(value: Any):
        """Normalize a display value.

        Parameters
        ----------
        value : Any
            Value to format.

        Returns
        -------
        Any
            Normalized string or original value.
        """
        if isinstance(value, str):
            return "\n".join(
                " ".join(line.split()) for line in value.splitlines() if line.strip()
            )
        return str(value) if isinstance(value, datetime) else value

    @classmethod
    def _format_body(cls, body: str, char_limit: int | None = None):
        """Normalize and optionally truncate message body text.

        Parameters
        ----------
        body : str
            Body text to format.
        char_limit : int, optional
            Maximum number of characters to return.

        Returns
        -------
        str
            Formatted body text.
        """
        if not isinstance(body, str) or not body:
            return ""
        body = cls._format(body)
        return body[:char_limit] if isinstance(char_limit, int) else body

    @staticmethod
    def _format_dt(value: datetime | None) -> str:
        """Format a plausible Outlook timestamp.

        Parameters
        ----------
        value : datetime, optional
            Timestamp to format.

        Returns
        -------
        str
            Timestamp text or an empty string.
        """
        if value is None:
            return ""
        try:
            return "" if abs(value.year - datetime.now().year) >= 100 else str(value)  # ruff: ignore[DTZ005]
        except (AttributeError, OverflowError, ValueError):
            return ""

    @property
    def _status(self) -> str:
        """Return the message status timestamp.

        Parameters
        ----------
        None

        Returns
        -------
        str
            Sent time, received time, or ``"Draft"``.
        """
        return str(self.sent_at or self.received_at or "Draft")

    def __str__(self) -> str:
        """Return a concise message description.

        Parameters
        ----------
        None

        Returns
        -------
        str
            Message description.
        """
        return (
            f"[{self._status}] sender={self.sender_address}, "
            f"subject={self.subject}, recipients={self.recipients}"
        )

    def __repr__(self) -> str:
        """Return a developer representation of the message.

        Parameters
        ----------
        None

        Returns
        -------
        str
            Message representation.
        """
        return (
            f"Mail(subject={self.subject!r}, sender={self.sender_address!r}, "
            f"recipients={self.recipients!r}, sent_at={self.sent_at!r}, "
            f"deliver_at={self.deliver_at!r})"
        )
