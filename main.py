from typing import List

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from Comparator import Comparator
from ExcelHandler import ExcelHandler

if __name__ == '__main__':
    PATH_TO_EXCEL: str = "assignment.xlsx"
    excelHandler: ExcelHandler = ExcelHandler()
    comparator: Comparator = Comparator()

    workbook: Workbook = excelHandler.read_workbook(PATH_TO_EXCEL)
    worksheets: List[Worksheet] = workbook.worksheets
    legacy_sheet: Worksheet = worksheets[0]
    scalable_sheet: Worksheet = worksheets[1]
    mapping_sheet: Worksheet = worksheets[2]

    comparator.compare(PATH_TO_EXCEL, legacy_sheet, scalable_sheet, mapping_sheet)