import re
from openpyxl import Workbook, load_workbook
from utils import valid_xlsx


def parse_xlsx(path: str) -> Workbook:
    valid_xlsx(path)

    return load_workbook(path)


def create_result_table(table: Workbook, group: int) -> Workbook:
    wb = Workbook()

    test_string = "1 2 3 4 12 24 я тупой пишу тут зачем то текст 97"
    template = r"\d+"

    res = re.findall(template, test_string)
    print(res)


def save_xlsx(table: Workbook, path: str):
    valid_xlsx(path, exists=False)
    table.save(path)


if __name__ == "__main__":
    # create_result_table(None, None)
    parsed = parse_xlsx(r"C:\Users\user\Downloads\социометрия (Responses).xlsx")
    new = save_xlsx(parsed, r"C:\Users\user\Downloads\социометрия (Responses)2.xlsx")