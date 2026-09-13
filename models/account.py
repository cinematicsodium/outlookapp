from __future__ import annotations

from collections.abc import Iterable
from functools import cached_property

from ..enums import FolderEnum, ItemType
from ..protocols import OlAccount, OlStore
from .base import ItemModel
from .folder import Folder


class Account(ItemModel):
    """Represent an Outlook account.

    Parameters
    ----------
    ol_acct_item : OlAccount
        Outlook account COM object to wrap.
    """

    item_type = ItemType.ACCOUNT
    required_properties = ("DisplayName", "SmtpAddress", "DeliveryStore")
    inaccessible_error_message = "Provided Outlook item is not an accessible account."

    def __init__(self, ol_acct_item: OlAccount) -> None:
        """Initialize an account wrapper.

        Parameters
        ----------
        ol_acct_item : OlAccount
            Outlook account COM object.

        Returns
        -------
        None
        """
        super().__init__(ol_acct_item)
        self.ol_item = ol_acct_item
        self._default_folders: dict[FolderEnum, Folder] = {}

    @cached_property
    def name(self) -> str:
        """Return the display name of the account."""
        return self.ol_item.DisplayName

    @cached_property
    def email_address(self) -> str:
        """Return the email address of the account."""
        return self.ol_item.SmtpAddress

    @cached_property
    def store(self) -> OlStore:
        """Return the default store associated with this account."""
        return self.ol_item.DeliveryStore

    @property
    def root_folder(self) -> Folder | None:
        """Return the root folder of the account's default store, if available."""
        return Folder.from_outlook_item(self.store.GetRootFolder())

    @property
    def folders(self) -> list[Folder]:
        """Return a list of folders in the account's root folder."""
        if (root := self.root_folder) is None:
            return []
        return root.subfolders

    def default_folder(self, folder: FolderEnum) -> Folder | None:
        """Return one of the account's default folders.

        Parameters
        ----------
        folder : FolderEnum
            Built-in folder to resolve.

        Returns
        -------
        Folder or None
            The wrapped default folder, if it is accessible.
        """
        if folder not in self._default_folders:
            resolved = Folder.from_outlook_item(self.store.GetDefaultFolder(folder))
            if resolved is not None:
                self._default_folders[folder] = resolved
        return self._default_folders.get(folder)

    def find_folder(self, path: str | Iterable[str]) -> Folder | None:
        """Find a folder below the account root.

        Parameters
        ----------
        path : str or iterable of str
            Slash-delimited path or ordered folder names.

        Returns
        -------
        Folder or None
            The matching folder, or ``None`` when the path cannot be resolved.
        """
        parts: list[str]
        if isinstance(path, str):
            parts = [segment.strip() for segment in path.split("/") if segment.strip()]
        else:
            parts = [str(segment).strip() for segment in path if str(segment).strip()]
        if not parts:
            return None
        root = self.root_folder
        if root is None:
            return None
        current = root.get_subfolder(parts[0])
        for segment in parts[1:]:
            if current is None:
                return None
            current = current.get_subfolder(segment)
        return current

    def matches(self, value: str | None) -> bool:
        """Check an account name or address for a case-insensitive match.

        Parameters
        ----------
        value : str or None
            Display name or SMTP address to compare.

        Returns
        -------
        bool
            ``True`` when the value identifies this account.
        """
        if not value:
            return False
        target = value.lower()
        return target in {self.name.lower(), self.email_address.lower()}

    def __str__(self) -> str:
        """Return the display name of the account."""
        return self.name

    def __repr__(self) -> str:
        """Return the string representation of the Account."""
        return f"Account({self.name})"
