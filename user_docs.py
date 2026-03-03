# ====================================================
# ПОЛЬЗОВАТЕЛЬСКАЯ ДОКУМЕНТАЦИЯ ДЛЯ CLI
# ====================================================
# Этот файл содержит тексты, которые видят пользователи
# в командах --help. Документация вынесена сюда, чтобы:
#   1. Использовать f-строки с подстановкой {MyApp.NAME}
#   2. Держать все пользовательские тексты в одном месте
#   3. Не загромождать cli.py длинными строками
#
# ВАЖНО: Эти docstrings ПЕРЕЗАПИСЫВАЮТ короткие docstrings
# из cli.py через декоратор @generate_docs. Пользователь
# видит именно эти тексты в --help.
# ====================================================

from enums import MyApp


LEXICON = {
    MyApp.FOLDERS:
        """
        Выводит все папки и их EntryID
        """,

    MyApp.FIND_FOLDERS:
        f"""
        Найти папки по имени и вывести их EntryID

        Примеры:
            {MyApp.NAME} {MyApp.FIND_FOLDERS} Входящие               # точное совпадение
            {MyApp.NAME} {MyApp.FIND_FOLDERS} -i входящие            # точное, но регистр не важен
            {MyApp.NAME} {MyApp.FIND_FOLDERS} -p Входящие            # все папки с "Входящие" в названии
            {MyApp.NAME} {MyApp.FIND_FOLDERS} -pi входящие           # все папки, содержащие "входящие" (регистр не важен)
        """,

    MyApp.EMAILS:
        f"""
        Получить письма из папки с фильтрацией

        Выводит:
            - EntryID писем (каждый с новой строки) по умолчанию
            - Количество писем (если указан --count)
    
        Примеры:
            {MyApp.NAME} {MyApp.EMAILS} <EntryID>                                 # список EntryID
            {MyApp.NAME} {MyApp.EMAILS} --count <EntryID>                         # только количество
            {MyApp.NAME} {MyApp.EMAILS} --status unread <EntryID>                 # только непрочитанные
            {MyApp.NAME} {MyApp.EMAILS} --status unread --flag exec <EntryID>     # непрочитанные с флагом
        """,

    MyApp.UPDATE:
        f"""
        Обновляет статус письма в Outlook по EntryID

        Используется для отметки писем как 'взято в работу' или 'выполнено',
        а также для управления флагами прочтения.
    
        Флаги разделены на две независимые группы (можно выбрать ТОЛЬКО ОДИН из каждой):
    
        Группа 1 (флажки): --exec, --complete, --clear
        Группа 2 (чтение): --read, --unread
    
        Всего возможных комбинаций: 3 × 2 = 6
    
        Примеры:
            {MyApp.NAME} {MyApp.UPDATE} --exec <EntryID>                # взять в работу
            {MyApp.NAME} {MyApp.UPDATE} --complete <EntryID>            # завершить
            {MyApp.NAME} {MyApp.UPDATE} --clear --read <EntryID>        # снять флажки и прочитать
            {MyApp.NAME} {MyApp.UPDATE} --exec --read <EntryID>         # взять в работу и прочитать
            {MyApp.NAME} {MyApp.UPDATE} --complete --unread <EntryID>   # завершить, но оставить непрочитанным
        """
}