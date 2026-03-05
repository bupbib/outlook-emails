from typing import Iterator
from win32com.client import CDispatch
import locale
from datetime import datetime
from user_docs import LEXICON
from enums import MyApp, EmailStatus, FlagStatus


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


def build_message_filter(
        status: EmailStatus,
        sender: str | None,
        flag: FlagStatus,
        date_from: datetime | None,
        date_to: datetime | None
) -> str:
    """
    Функция для построения фильтра поиска сообщений.
    """
    conditions = []
    locale.setlocale(locale.LC_ALL, '')

    if status in (EmailStatus.READ, EmailStatus.UNREAD):
        conditions.append('[Unread] = True' if status == EmailStatus.UNREAD else '[Unread] = False')

    if sender:
        conditions.append(f'[SenderEmailAddress] = "{sender}"')

    if flag != FlagStatus.ALL:
        conditions.append({
            FlagStatus.ANY: '[FlagStatus] <> 0',
            FlagStatus.NONE: '[FlagStatus] = 0',
            FlagStatus.EXEC: '[FlagStatus] = 1',
            FlagStatus.COMP: '[FlagStatus] = 2'
        }.get(flag))

    if date_from:
        start = date_from.replace(hour=0, minute=0, second=0).strftime('%x %H:%M')
        conditions.append(f'[ReceivedTime] >= "{start}"')

    if date_to:
        end = date_to.replace(hour=23, minute=59, second=59).strftime('%x %H:%M')
        conditions.append(f'[ReceivedTime] <= "{end}"')

    return ' AND '.join(conditions)