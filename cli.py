import os
import sys
import typer
from typer import colors
import pythoncom
import win32com.client
from utils import get_all_folders, generate_docs
from enums import MyApp, EmailStatus, FlagStatus
from validators import parse_date, parse_email, validate_dates


app = typer.Typer(
    name=MyApp.NAME,
    help="""Утилита для работы с Outlook с помощью COM-интерфейса\n
        Требования:
            • Windows OS (используется COM API)
            • Microsoft Outlook (установлен и настроен)
            • Запущенный экземпляр Outlook
    """,
    no_args_is_help=True
)


@app.callback()
def main(ctx: typer.Context):
    """
    Callback-функция, выполняемая перед каждой командой.

    Вызывает ошибку, если:
        • Операционная система не Windows
        • Microsoft Outlook не запущен
    """
    if not sys.platform.startswith('win') or os.name != 'nt':
        typer.secho(
            'Ошибка: утилита работает только на Windows',
            fg=colors.RED
        )
        raise typer.Exit(1)  # Завершаем выполнение утилиты

    try:
        outlook = win32com.client.GetActiveObject('Outlook.Application')
        ctx.obj = outlook.GetNamespace('MAPI')
    except pythoncom.com_error as err:
        typer.secho(
            'Ошибка: Outlook не запущен. Запустите Outlook и повторите попытку',
            fg=colors.RED
        )
        raise typer.Exit(1) from err


@app.command(MyApp.FOLDERS)
@generate_docs(MyApp.FOLDERS)
def all_folders(ctx: typer.Context):
    """
    Выводит все папки и их EntryID
    """
    namespace: win32com.client.CDispatch = ctx.obj
    total = 0

    typer.secho('Сканирую папки...', fg=colors.CYAN)

    for total, folder in enumerate(get_all_folders(namespace.Folders), 1):
        typer.secho(f'{total}. Имя папки: «{folder.Name}», EntryID: {folder.EntryID}')

    typer.secho(f'Готово! Найдено папок: {total}', fg=colors.CYAN)


@app.command(MyApp.FIND_FOLDERS)
@generate_docs(MyApp.FIND_FOLDERS)
def find_folders(
        ctx: typer.Context,
        name: str = typer.Argument(..., help='Название папки'),
        partial: bool = typer.Option(False, '--partial', '-p', help='Искать частичное совпадение'),
        ignore_case: bool = typer.Option(False, '--ignore-case', '-i', help='Игнорировать регистр при поиске')
):
    """
    Найти папки по имени и вывести их EntryID
    """
    namespace: win32com.client.CDispatch = ctx.obj
    total_finds = 0
    target_folder = name.lower() if ignore_case else name

    typer.secho('Сканирую папки...', fg=colors.CYAN)

    for folder in get_all_folders(namespace.Folders):
        current_folder = folder.Name.lower() if ignore_case else folder.Name

        if current_folder == target_folder or ((target_folder in current_folder) if partial else False):
            typer.secho(f'Имя папки: «{folder.Name}», EntryID: {folder.EntryID}')
            total_finds += 1

    if total_finds:
        typer.secho(f'Найдено папок: {total_finds}', fg=colors.CYAN)
    else:
        typer.secho(f'Ни одной папки с совпадением "{name}" не найдено', fg=colors.CYAN)


@app.command(MyApp.EMAILS)
@generate_docs(MyApp.EMAILS)
def emails(
        ctx: typer.Context,
        entry_id: str = typer.Argument(..., help='EntryID папки'),
        status: EmailStatus = typer.Option(EmailStatus.BOTH, help='Фильтр по прочитанности'),
        sender: str | None = typer.Option(
            None, '--sender', callback=parse_email, help='Фильтр по email отправителя'
        ),
        flag: FlagStatus = typer.Option(
            FlagStatus.ALL, '--flag', help='Фильтр по флагам'
        ),
        date_from: str | None = typer.Option(
            None, '--from', callback=parse_date, help='Отбор писем начиная с указанной даты (ДД.ММ.ГГГГ)'
        ),
        date_to: str | None = typer.Option(
            None, '--to', callback=parse_date, help='Отбор писем заканчивая указанной датой, включительно (ДД.ММ.ГГГГ)'
        ),
        count: bool = typer.Option(False, '--count', help='Показать только количество писем (без вывода EntryID)')
):
    """
    Получить письма из папки с фильтрацией
    """
    validate_dates(date_from, date_to)
    namespace: win32com.client.CDispatch = ctx.obj

    try:
        current_folder = namespace.GetFolderFromID(entry_id)
    except pythoncom.com_error as err:
        typer.secho(
            f'Ошибка: Папка с EntryID "{entry_id}" не найдена\n'
            f'Используйте {MyApp.NAME} {MyApp.FOLDERS} для просмотра доступных папок',
            fg=colors.RED
        )
        raise typer.Exit(1) from err


@app.command(MyApp.UPDATE)
@generate_docs(MyApp.UPDATE)
def update(
        ctx: typer.Context,
        entry_id: str = typer.Argument(..., help='EntryID письма'),
        set_exec: bool = typer.Option(False, '--exec', help="Установить флаг 'На исполнение'"),
        set_complete: bool = typer.Option(False, '--complete', help="Установить флаг 'Завершено'"),
        clear_flag: bool = typer.Option(False, '--clear', help='Снять все флаги'),
        read: bool = typer.Option(False, '--read', help='Пометить как прочитанное'),
        unread: bool = typer.Option(False, '--unread', help='Пометить как непрочитанное')
):
    """
    Обновляет статус письма в Outlook по EntryID
    """
    pass


if __name__ == '__main__':
    app()