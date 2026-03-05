import typer
import re
from datetime import datetime


def parse_date(date_str: str | None) -> datetime | None:
    """Проверяет формат даты"""
    if date_str is not None:
        # валидация только если есть значение
        try:
            return datetime.strptime(date_str, '%d.%m.%Y')
        except ValueError as err:
            raise typer.BadParameter(
                f'Неверный формат даты: {date_str}\n'
                f'Корректный формат: ДД.ММ.ГГГГ (например: 02.03.2026)'
            ) from err
    return None


def parse_email(email_str: str | None) -> None:
    if (email_str is not None) and (not re.fullmatch(r'[\w.-]+@[\w.-]+\.\w+', email_str)):
        raise typer.BadParameter(
            f'Неверный формат почты: {email_str}\n'
            f'Корректный формат: username@example.com'
        )
    return email_str