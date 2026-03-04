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


def validate_dates(date_from: str | None, date_to: str | None) -> None:
    """Проверяет, что даты указаны корректно и не противоречат друг другу"""
    if date_from is None or date_to is None:
        return

    from_dt = datetime.strptime(date_from, '%d.%m.%Y')
    to_dt = datetime.strptime(date_to, '%d.%m.%Y')

    if from_dt > to_dt:
        raise typer.BadParameter(
            f'Дата начала {date_from} позже даты окончания {date_to}'
        )