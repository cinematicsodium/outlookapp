from __future__ import annotations

import logging
from collections.abc import Iterable, Iterator
from dataclasses import dataclass
from functools import cached_property
from typing import cast

from ..enums import ItemType
from ..exceptions import COM_ERRORS
from ..protocols import OlFolder, OlMailItem
from .base import BaseModel
from .mail_item import MailItem

logger = logging.getLogger(__name__)


@dataclass(frozen=True, slots=True)
class FolderListing:
    """Represent a folder in a tree listing.

    Parameters
    ----------
    path : str
        Slash-delimited folder path.
    depth : int
        Depth relative to the walk root.
    subfolder_count : int
        Outlook's reported number of direct children, including entries that
        cannot be opened.
    """

    path: str
    depth: int
    subfolder_count: int

    def as_row(self) -> tuple[object, ...]:
        """Return the listing fields as a table row."""
        return (self.path, self.depth, self.subfolder_count)


class Folder(BaseModel):
    """Represent an Outlook folder.

    Parameters
    ----------
    item : OlFolder
        Outlook folder COM object to wrap.
    """

    item_type = ItemType.FOLDER
    required_properties = ("Name", "Items", "Folders")

    def __init__(self, folder: OlFolder) -> None:
        """Initialize a folder wrapper.

        Parameters
        ----------
        folder : OlFolder
            Outlook folder COM object.

        Returns
        -------
        None
        """
        super().__init__(folder)
        self._protocol = folder

    @cached_property
    def name(self) -> str:
        """Return the folder name."""
        return self._protocol.Name

    @cached_property
    def folder_path(self) -> str:
        """Return the full Outlook folder path."""
        return self._protocol.FolderPath

    @property
    def mail_items(self) -> list[MailItem]:
        """Return a list of MailItem objects in the folder."""
        return self.list_messages()

    @property
    def subfolders(self) -> list[Folder]:
        """Return a list of subfolders."""
        return list(self.iter_subfolders())

    @property
    def subfolder_count(self) -> int:
        """Return Outlook's reported number of direct child folders.

        The count can include folders that cannot be opened through this
        wrapper. Return zero if the count cannot be read.
        """
        try:
            return self._protocol.Folders.Count
        except COM_ERRORS:
            logger.warning("Unable to count subfolders for '%s'", self.name)
            return 0

    def list_messages(
        self,
        limit: int | None = None,
        unread_only: bool = False,
    ) -> list[MailItem]:
        """Return messages from the folder.

        Parameters
        ----------
        limit : int, optional
            Maximum number of messages to return.
        unread_only : bool, default=False
            Return only unread messages.

        Returns
        -------
        list of MailItem
            Accessible messages ordered by received time when Outlook supports
            sorting.

        Raises
        ------
        ValueError
            If ``limit`` is less than one.
        """
        if limit is not None and limit < 1:
            raise ValueError("limit must be >= 1")
        return list(self._iter_messages(limit=limit, unread_only=unread_only))

    def _iter_messages(
        self, limit: int | None = None, unread_only: bool = False
    ) -> Iterator[MailItem]:
        """Yield accessible messages from the folder.

        Parameters
        ----------
        limit : int, optional
            Maximum number of messages to yield.
        unread_only : bool, default=False
            Yield only unread messages.

        Yields
        ------
        MailItem
            Accessible messages ordered by received time when Outlook supports
            sorting.
        """
        items = self._protocol.Items
        if unread_only:
            try:
                items = items.Restrict("[UnRead] = True")
                unread_only = False
            except COM_ERRORS:
                logger.debug("Unable to filter unread messages in '%s'", self.name)
        try:
            items.Sort("[ReceivedTime]", True)
        except COM_ERRORS:
            logger.debug("Unable to sort messages in '%s'", self.name)
        count = 0
        # ponytail: live indices; use list(folder) before moving/deleting messages.
        for index in range(1, items.Count + 1):
            try:
                item = items.Item(index)
                if unread_only and not cast(OlMailItem, item).UnRead:
                    continue
                message = MailItem.from_outlook_item(item)
            except COM_ERRORS:
                logger.debug("Unable to read message %s in '%s'", index, self.name)
                continue
            if message is not None:
                yield message
                count += 1
                if limit is not None and count >= limit:
                    return

    def iter_subfolders(self) -> Iterable[Folder]:
        """Iterate over accessible child folders."""
        try:
            folders = self._protocol.Folders
            count = folders.Count
        except COM_ERRORS:
            logger.warning("Unable to enumerate subfolders for '%s'", self.name)
            return
        for index in range(1, count + 1):
            try:
                folder = Folder.from_outlook_item(folders.Item(index))
            except COM_ERRORS:
                logger.debug("Unable to read subfolder %s in '%s'", index, self.name)
                continue
            if folder is not None:
                yield folder

    def walk(
        self,
        recursive: bool = False,
        max_depth: int = 0,
        depth: int = 0,
        parent_path: str = "",
    ) -> list[FolderListing]:
        """Return folder tree entries rooted at this folder.

        Parameters
        ----------
        recursive : bool, default=False
            Include descendant folders.
        max_depth : int, default=0
            Maximum descendant depth when walking recursively.
        depth : int, default=0
            Depth assigned to this folder.
        parent_path : str, default=""
            Path prepended to this folder's name.

        Returns
        -------
        list of FolderListing
            Folder entries in depth-first order.

        Raises
        ------
        ValueError
            If ``max_depth`` is negative for a recursive walk.
        """
        if recursive and max_depth < 0:
            raise ValueError("max_depth must be >= 0")
        path = f"{parent_path}/{self.name}" if parent_path else self.name
        folder_listing_entries = [
            FolderListing(path=path, depth=depth, subfolder_count=self.subfolder_count)
        ]
        if not recursive or depth >= max_depth:
            return folder_listing_entries
        for subfolder in self.iter_subfolders():
            folder_listing_entries.extend(
                subfolder.walk(
                    recursive=recursive,
                    max_depth=max_depth,
                    depth=depth + 1,
                    parent_path=path,
                )
            )
        return folder_listing_entries

    def get_item(self, index: int) -> MailItem | None:
        """Return a message by zero-based index.

        Parameters
        ----------
        index : int
            Zero-based message index.

        Returns
        -------
        MailItem or None
            The accessible message, or ``None`` when the index is out of range.

        Raises
        ------
        ValueError
            If ``index`` is negative.
        """
        if index < 0:
            raise ValueError("index must be >= 0")
        items = self._protocol.Items
        if index >= items.Count:
            return None
        return MailItem.from_outlook_item(items.Item(index + 1))

    def get_subfolder(self, folder_name: str) -> Folder | None:
        """Return a direct child folder by name.

        Parameters
        ----------
        folder_name : str
            Case-insensitive folder name.

        Returns
        -------
        Folder or None
            The matching child folder, if present.

        Raises
        ------
        ValueError
            If ``folder_name`` is not a string.
        """
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")  # ruff: ignore[TRY004]
        normalized_folder_name: str = folder_name.lower()
        for folder in self.iter_subfolders():
            if (folder.name or "").lower() == normalized_folder_name:
                return folder
        return None

    def create_subfolder(self, folder_name: str) -> Folder | None:
        """Create a direct child folder.

        Parameters
        ----------
        folder_name : str
            Name of the new folder.

        Returns
        -------
        Folder or None
            The new folder, if Outlook returns an accessible object.

        Raises
        ------
        ValueError
            If ``folder_name`` is not a nonempty string.
        AttributeError
            If the folder has no child-folder collection.
        """
        if not isinstance(folder_name, str):
            raise ValueError("folder_name must be a string")  # ruff: ignore[TRY004]
        if not folder_name.strip():
            raise ValueError("folder_name cannot be empty")
        folders = self._protocol.Folders
        if folders is None:
            raise AttributeError("Folder does not expose a Folders collection")
        return Folder(folders.Add(folder_name))

    def delete_subfolder(self, folder_name: str) -> bool:
        """Delete a direct child folder by name.

        Parameters
        ----------
        folder_name : str
            Case-insensitive folder name.

        Returns
        -------
        bool
            ``True`` when a matching folder was deleted.
        """
        folder: Folder | None = self.get_subfolder(folder_name)
        if folder is None:
            return False
        folder.delete()
        return True

    def move_item(self, mail_item: MailItem, destination: Folder) -> MailItem | None:
        """Move a message to another folder.

        Parameters
        ----------
        mail_item : MailItem
            Message to move.
        destination : Folder
            Destination folder.

        Returns
        -------
        MailItem or None
            The moved message, if Outlook returns an accessible object.

        Raises
        ------
        ValueError
            If either argument has the wrong model type.
        """
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")  # ruff: ignore[TRY004]
        if not isinstance(destination, Folder):
            raise ValueError("destination must be a Folder")  # ruff: ignore[TRY004]
        return mail_item.move(destination)

    def delete_item(self, mail_item: MailItem) -> None:
        """Delete a message from the folder.

        Parameters
        ----------
        mail_item : MailItem
            Message to delete.

        Raises
        ------
        ValueError
            If ``mail_item`` is not a :class:`MailItem`.
        """
        if not isinstance(mail_item, MailItem):
            raise ValueError("mail_item must be a MailItem")  # ruff: ignore[TRY004]
        mail_item.delete()

    def delete(self) -> None:
        """Delete this folder."""
        self._protocol.Delete()

    def __iter__(self) -> Iterator[MailItem]:
        """Iterate over the folder's messages.

        Parameters
        ----------
        None

        Returns
        -------
        Iterator of MailItem
            Messages in the folder.
        """
        yield from self._iter_messages()

    def __repr__(self) -> str:
        """Return repr string for the folder."""
        name = self.name
        folder_path = self.folder_path
        return f"Folder({name=!r}, {folder_path=!r})"

    def __str__(self) -> str:
        """Return string representation of the folder."""
        return self.__repr__()
