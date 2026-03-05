import typer
import re


def parse_email(email_str: str | None) -> None:
    if (email_str is not None) and (not re.fullmatch(r'[\w.-]+@[\w.-]+\.\w+', email_str)):
        raise typer.BadParameter(
            f'Неверный формат почты: {email_str}\n'
            f'Корректный формат: username@example.com'
        )
    return email_str