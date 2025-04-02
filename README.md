# society_research

Программа для составления таблицы, согласно данному заданию. Конвертирует таблицу ответов на [форму](https://docs.google.com/forms/d/1DqNI7FpK0Tc3gMhmwqlesUNj930d-cXYYx7fxzQ7Pmc/edit) в Excel-файл с "матрицами" на 6 разных листах.

Для запуска необходимо ввести в терминал <code> python main.py "[полный путь к xls/xlsx файлу, полученному из гугл форм]" "[полный путь для сохранения итогового файла(включая xlsx)]" [количество студентов в группе] [номер группы] </code>

Пример:

<code> python main.py "C:\Prog\society_research\google_doc_raw.xlsx" "C:\Prog\society_research\result_table.xlsx" 25 6114 </code>
