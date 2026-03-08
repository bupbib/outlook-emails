from enum import Enum


class MyApp(str, Enum):
    """Название приложения и команд утилиты"""
    NAME = 'outlook-emails'
    FOLDERS = 'folders'
    FIND_FOLDERS = 'find-folders'
    EMAILS = 'emails'
    UPDATE = 'update'

    def __str__(self):
        return self.value


class EmailStatus(str, Enum):
    READ = 'read'
    UNREAD = 'unread'
    BOTH = 'both'


class FlagStatus(str, Enum):
    ANY = 'any'    # любой флаг
    EXEC = 'exec'  # на исполнение
    COMP = 'comp'  # завершено
    NONE = 'none'  # без флага
    ALL = 'all'    # все письма (без фильтра по флагам)


class Period(str, Enum):
    TODAY = 'today'  
    WEEK = 'week'
    MONTH = 'month'