from __future__ import annotations

from datetime import datetime
from typing import Literal, Protocol

from .types import T


class OlCollection(Protocol[T]):
    """Structural type for a one-indexed Outlook COM collection."""

    @property
    def Count(self) -> int:
        """Return the number of collection entries."""
        ...

    def Item(self, index: int, /) -> T:
        """Return the item at a one-based collection index.

        Parameters
        ----------
        index : int
            One-based item index.

        Returns
        -------
        T
            Item at ``index``.
        """
        ...


class OlObject(Protocol):
    """Expose the Outlook class identifier used for model validation."""

    @property
    def Class(self) -> int:
        """Return the Outlook object-class identifier."""
        ...


class OlAccount(OlObject, Protocol):
    """Expose configured account metadata and its delivery store."""

    @property
    def DisplayName(self) -> str:
        """Return the account display name."""
        ...

    @property
    def SmtpAddress(self) -> str:
        """Return the account SMTP address."""
        ...

    @property
    def DeliveryStore(self) -> OlStore:
        """Return the account's delivery store."""
        ...


class OlAddressEntry(OlObject, Protocol):
    """Expose the address-book fields used for SMTP resolution."""

    @property
    def Address(self) -> str | None:
        """Return the underlying address, when available."""
        ...

    @property
    def Name(self) -> str:
        """Return the entry display name."""
        ...

    @property
    def Type(self) -> str:
        """Return the address type checked by accessibility validation."""
        ...

    @property
    def AddressEntryUserType(self) -> int:
        """Return the address-entry user type."""
        ...

    @property
    def PropertyAccessor(self) -> _OlPropertyAccessor:
        """Return the accessor used to read the MAPI SMTP property."""
        ...

    def GetExchangeUser(self) -> _OlExchangeUser | None:
        """Return the corresponding Exchange user, if available."""
        ...


class OlApplication(Protocol):
    """Expose the application operations used to connect and compose mail."""

    def GetNamespace(self, name: Literal["MAPI"], /) -> OlNamespace:
        """Open the MAPI namespace.

        Parameters
        ----------
        name : {"MAPI"}
            Namespace used by this package.

        Returns
        -------
        OlNamespace
            Active Outlook namespace.
        """
        ...

    def CreateItem(self, item_type: Literal[0], /) -> OlMailItem:
        """Create a mail item.

        Parameters
        ----------
        item_type : {0}
            Outlook mail-item creation constant.

        Returns
        -------
        OlMailItem
            New unsaved message.
        """
        ...


class OlFolder(OlObject, Protocol):
    """Expose folder navigation, message access, and deletion."""

    @property
    def Name(self) -> str:
        """Return the folder name."""
        ...

    @property
    def FolderPath(self) -> str:
        """Return the full Outlook folder path."""
        ...

    @property
    def Items(self) -> _OlItems:
        """Return the collection, which may contain non-mail Outlook objects."""
        ...

    @property
    def Folders(self) -> _OlFolders:
        """Return the direct child-folder collection."""
        ...

    def Delete(self) -> None:
        """Delete the folder."""
        ...


class OlMailItem(OlObject, Protocol):
    """Expose message fields and operations used by the mail wrapper."""

    SentOnBehalfOfName: str | None
    To: str | None
    CC: str | None
    BCC: str | None
    Subject: str | None
    Body: str
    HTMLBody: str
    DeferredDeliveryTime: datetime | None
    UnRead: bool
    SendUsingAccount: OlAccount

    @property
    def EntryID(self) -> str | None:
        """Return the message identifier, if assigned."""
        ...

    @property
    def ConversationID(self) -> str | None:
        """Return the conversation identifier."""
        ...

    @property
    def ConversationIndex(self) -> str | None:
        """Return the conversation index."""
        ...

    @property
    def Parent(self) -> OlFolder | None:
        """Return the parent folder, when available."""
        ...

    @property
    def Sender(self) -> OlAddressEntry | None:
        """Return the sender's address-book entry, when available."""
        ...

    @property
    def SenderName(self) -> str | None:
        """Return the sender display name."""
        ...

    @property
    def SenderEmailAddress(self) -> str | None:
        """Return the sender's underlying email address."""
        ...

    @property
    def Attachments(self) -> _OlAttachments:
        """Return the attachment collection."""
        ...

    @property
    def Recipients(self) -> _OlRecipients:
        """Return the recipient collection."""
        ...

    @property
    def SentOn(self) -> datetime | None:
        """Return the sent timestamp, when available."""
        ...

    @property
    def ReceivedTime(self) -> datetime | None:
        """Return the received timestamp, when available."""
        ...

    @property
    def Size(self) -> int:
        """Return the message size in bytes."""
        ...

    def Display(self) -> None:
        """Display the message inspector."""
        ...

    def Send(self) -> None:
        """Submit the message for sending."""
        ...

    def Save(self) -> None:
        """Save the message in Outlook."""
        ...

    def SaveAs(self, path: str, /) -> None:
        """Export the message to a file.

        Parameters
        ----------
        path : str
            Destination file path.
        """
        ...

    def Delete(self) -> None:
        """Delete the message."""
        ...

    def Move(self, destination: OlFolder, /) -> OlMailItem:
        """Move the message to another folder.

        Parameters
        ----------
        destination : OlFolder
            Target Outlook folder.

        Returns
        -------
        OlMailItem
            Message returned by Outlook after the move.
        """
        ...


class OlNamespace(Protocol):
    """Expose configured accounts and shared-mailbox resolution."""

    @property
    def Accounts(self) -> OlCollection[OlAccount]:
        """Return the accounts configured in the Outlook profile."""
        ...

    def CreateRecipient(self, address: str, /) -> _OlRecipient:
        """Create a recipient for address-book resolution.

        Parameters
        ----------
        address : str
            Mailbox SMTP address.

        Returns
        -------
        _OlRecipient
            Recipient to resolve before accessing a shared folder.
        """
        ...

    def GetSharedDefaultFolder(
        self, recipient: _OlRecipient, folder_type: int, /
    ) -> OlFolder:
        """Open a resolved recipient's shared default folder.

        Parameters
        ----------
        recipient : _OlRecipient
            Resolved mailbox recipient.
        folder_type : int
            Outlook default-folder identifier.

        Returns
        -------
        OlFolder
            Requested shared folder.
        """
        ...


class OlStore(Protocol):
    """Expose the root and default folders of an account's store."""

    def GetRootFolder(self) -> OlFolder:
        """Return the store's root folder."""
        ...

    def GetDefaultFolder(self, folder_type: int, /) -> OlFolder:
        """Open a default folder in the store.

        Parameters
        ----------
        folder_type : int
            Outlook default-folder identifier.

        Returns
        -------
        OlFolder
            Requested default folder.
        """
        ...


class _OlItems(OlCollection[OlObject], Protocol):
    """Expose filtering and sorting for a mixed Outlook item collection."""

    def Sort(self, field: str, descending: bool, /) -> None:
        """Sort the collection in place.

        Parameters
        ----------
        field : str
            Outlook property to sort by.
        descending : bool
            Whether to sort in descending order.
        """
        ...

    def Restrict(self, query: str, /) -> _OlItems:
        """Return items matching an Outlook filter.

        Parameters
        ----------
        query : str
            Outlook restriction expression.

        Returns
        -------
        _OlItems
            Filtered collection.
        """
        ...


class _OlFolders(OlCollection[OlFolder], Protocol):
    """Expose child-folder enumeration and creation."""

    def Add(self, name: str, /) -> OlFolder:
        """Create a child folder.

        Parameters
        ----------
        name : str
            New folder name.

        Returns
        -------
        OlFolder
            Created folder.
        """
        ...


class _OlAttachment(Protocol):
    """Expose the attachment filename used in listings."""

    @property
    def FileName(self) -> str:
        """Return the attachment filename."""
        ...


class _OlAttachments(OlCollection[_OlAttachment], Protocol):
    """Expose attachment enumeration, counts, and file attachment."""

    def Add(self, path: str, /) -> _OlAttachment:
        """Attach a local file.

        Parameters
        ----------
        path : str
            Existing attachment file path.

        Returns
        -------
        _OlAttachment
            Added attachment.
        """
        ...


class _OlRecipient(Protocol):
    """Expose recipient resolution and address-book access."""

    @property
    def AddressEntry(self) -> OlAddressEntry | None:
        """Return the recipient's address-book entry, when available."""
        ...

    def Resolve(self) -> bool:
        """Return whether Outlook resolves the recipient."""
        ...


class _OlRecipients(OlCollection[_OlRecipient], Protocol):
    """Expose recipient enumeration and bulk resolution."""

    def ResolveAll(self) -> bool:
        """Return whether Outlook resolves every recipient."""
        ...


class _OlPropertyAccessor(Protocol):
    """Expose the MAPI property read used for SMTP lookup."""

    def GetProperty(self, schema: str, /) -> object:
        """Read a property identified by its MAPI schema.

        Parameters
        ----------
        schema : str
            MAPI property schema URI.

        Returns
        -------
        object
            Raw property value converted to text by the caller.
        """
        ...


class _OlExchangeUser(Protocol):
    """Expose the primary SMTP address of a resolved Exchange user."""

    @property
    def PrimarySmtpAddress(self) -> str | None:
        """Return the primary SMTP address, when available."""
        ...
