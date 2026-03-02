from typing import Iterator
from win32com.client import CDispatch


def get_all_folders(folders: CDispatch) -> Iterator[CDispatch]:
    """
    Рекурсивно обходит все папки Outlook и возвращает конечные папки (листья).

    Args:
        folders: Коллекция папок Outlook (обычно Namespace.Folders или Folder.Folders)

    Yields:
        Объекты папок Outlook (как CDispatch), не содержащих вложенных папок.

    Notes:
        Возвращаемые объекты имеют тип CDispatch, но содержат все свойства и методы
        папки Outlook: Name, EntryID, Folders и т.д.
    """
    for folder in folders:
        if folder.Folders.Count > 0:
            yield from get_all_folders(folder.Folders)
        else:
            yield folder
