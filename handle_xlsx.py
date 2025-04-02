from openpyxl import Workbook, load_workbook
from utils import valid_xlsx




def parse_xlsx(path: str) -> Workbook:
    valid_xlsx(path)
    return load_workbook(path)


def create_result_table(input_table: Workbook, group_size: int) -> Workbook:
    """
    Creates table as described in technical task
    :param input_table: Table from Google Forms, cleared from crap
    :param group_size: Number of academical group members
    :return: Table as described in technical task
    """
    wb = Workbook()
    def_sheet = wb.active
    wb.remove(def_sheet) # delete default sheet
    for sheet in range(1,7): # create 6 sheets
        new_sheet = wb.create_sheet(title=f"Вопрос № {sheet}")
        for j in range(1, group_size + 1): # create table NxN, N - group size
            new_sheet.cell(j + 2, 2).value = j
            new_sheet.cell(2, j + 2).value = j

        for row in range(3, group_size + 3): # fill table with zeros
            for col in range(3, group_size + 3):
                new_sheet.cell(row, col).value = 0

        for row in range(1, group_size + 1): # fill table with ones
            cur_student_number = input_table.active.cell(row + 1, 3).value
            response_value = str(input_table.active.cell(row + 1, sheet + 3).value)
            print(cur_student_number, " ", response_value)
            numbers = [int(x) for x in response_value.split()]
            for number in numbers:
                new_sheet.cell(cur_student_number + 2, number + 2 ).value = 1
    return wb


def save_xlsx(table: Workbook, path: str) -> None:
    """

    :param table:
    :param path:
    :return:
    """
    valid_xlsx(path, exists=False)
    table.save(path)


if __name__ == "__main__":
    # create_result_table(None, None)
    parsed = parse_xlsx(r"C:\Torrent\society_research\google_doc_table.xlsx")
    new = save_xlsx(parsed, r"C:\Users\user\Downloads\социометрия (Responses)2.xlsx")
    for i in range (6):
        print(i)