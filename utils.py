from typing import Iterator
from win32com.client import CDispatch
from user_docs import LEXICON
from enums import MyApp


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


def generate_docs(cmd: MyApp):
    """Декоратор для установки документации из LEXICON"""
    def decorator(func):
        func.__doc__ = LEXICON[cmd]
        return func

    return decorator