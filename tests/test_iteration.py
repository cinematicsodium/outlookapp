from __future__ import annotations

from types import SimpleNamespace

from outlook import Account, Folder, utils
from outlook.enums import ItemType
from outlook.models.folder import FolderListing
from outlook.tests.test_model_navigation import (
    FakeAccount,
    FakeCollection,
    FakeFolder,
    FakeMailItem,
)


class OrderedCollection(FakeCollection):
    """Provide a fake collection that supports Outlook restrictions."""

    def Restrict(self, query: str) -> OrderedCollection:
        """Return unread messages for Outlook's unread restriction."""
        assert query == "[UnRead] = True"
        self.restricted = OrderedCollection(
            [item for item in self.items if getattr(item, "UnRead", False)]
        )
        return self.restricted

    def Sort(self, field: str, descending: bool) -> None:
        """Sort fake mail by received time."""
        super().Sort(field, descending)
        self.items.sort(key=lambda item: item.ReceivedTime, reverse=descending)


class FlakyCollection(FakeCollection):
    """Fail while resolving one collection entry."""

    def Item(self, index: int) -> object:
        """Raise an accessibility error for the third entry."""
        if index == 3:
            raise AttributeError
        return super().Item(index)


def test_folder_iteration_reads_only_requested_message() -> None:
    """Verify direct iteration does not materialize the entire folder."""
    items = OrderedCollection([FakeMailItem(str(index)) for index in range(10_000)])
    for index, item in enumerate(items.items):
        item.ReceivedTime = index

    first = next(iter(Folder(FakeFolder("Inbox", items=items))))

    assert first.subject == "9999"
    assert items.item_calls == 1


def test_subfolder_iteration_reads_only_requested_folder() -> None:
    """Verify child-folder iteration does not materialize the collection."""
    folder = Folder(
        FakeFolder(
            "Root", subfolders=[FakeFolder(str(index)) for index in range(1_000)]
        )
    )
    first = next(iter(folder.iter_subfolders()))

    assert first.name == "0"
    assert folder._protocol.Folders.item_calls == 1


def test_folder_walk_uses_native_count_without_opening_children() -> None:
    """Verify a bounded walk does not open child folders."""
    raw = FakeFolder(
        "Root", subfolders=[FakeFolder(str(index)) for index in range(1_000)]
    )
    folder = Folder(raw)

    assert folder.walk() == [FolderListing("Root", 0, 1_000)]
    assert raw.Folders.item_calls == 0

    assert folder.walk(recursive=True, max_depth=0)[0].subfolder_count == 1_000
    assert raw.Folders.item_calls == 0


def test_subfolder_iteration_skips_inaccessible_entries() -> None:
    """Verify inaccessible child folders do not stop later enumeration."""
    first = FakeFolder("first")
    last = FakeFolder("last")
    folders = FlakyCollection(
        [first, SimpleNamespace(Class=ItemType.MAIL_ITEM), object(), last]
    )
    folder = Folder(FakeFolder("Root"))
    folder._protocol.Folders = folders

    assert [child.name for child in folder.iter_subfolders()] == ["first", "last"]


def test_account_folder_lookup_stops_after_root_match() -> None:
    """Verify root-folder lookup stops once it finds the requested child."""
    raw_root = FakeFolder(
        "Root",
        subfolders=[
            FakeFolder("Inbox"),
            *[FakeFolder(str(index)) for index in range(999)],
        ],
    )
    account = Account(FakeAccount("Primary", "user@example.com", raw_root))

    assert account.find_folder("Inbox") is not None
    assert raw_root.Folders.item_calls == 1


def test_unread_restriction_is_sorted_after_filtering() -> None:
    """Verify Outlook restriction retains received-time ordering."""
    old = FakeMailItem("old")
    old.ReceivedTime = 1
    read = FakeMailItem("read", unread=False)
    read.ReceivedTime = 3
    new = FakeMailItem("new")
    new.ReceivedTime = 2
    items = OrderedCollection([old, read, new])

    messages = Folder(FakeFolder("Inbox", items=items)).list_messages(unread_only=True)

    assert [message.subject for message in messages] == ["new", "old"]
    assert items.sort_calls == []
    assert items.restricted is not None
    assert items.restricted.sort_calls == [("[ReceivedTime]", True)]


def test_iteration_skips_inaccessible_collection_entries() -> None:
    """Verify non-mail and inaccessible COM entries do not stop enumeration."""
    first = FakeMailItem("first")
    second = FakeMailItem("second")
    items = FlakyCollection(
        [first, SimpleNamespace(Class=ItemType.FOLDER), object(), second]
    )

    messages = Folder(FakeFolder("Inbox", items=items)).list_messages()

    assert [message.subject for message in messages] == ["first", "second"]


def test_inaccessible_class_property_is_rejected(monkeypatch) -> None:
    """Verify a COM failure while reading Class is treated as inaccessible."""

    class ComFailure(Exception):
        """Represent a COM property-access error."""

    class BrokenItem:
        @property
        def Class(self) -> ItemType:
            """Raise while resolving the COM Class property."""
            raise ComFailure

    monkeypatch.setattr(utils, "COM_ERRORS", (ComFailure,))

    assert not utils.is_accessible_ol_item(BrokenItem(), ItemType.MAIL_ITEM)
