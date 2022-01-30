from typing import List

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame

from Comparator import Comparator
from ExcelHandler import ExcelHandler
import pandas as pd

if __name__ == '__main__':
    PATH_TO_EXCEL: str = "assignment.xlsx"
    excelHandler: ExcelHandler = ExcelHandler()
    comparator: Comparator = Comparator()

    workbook: Workbook = excelHandler.read_workbook(PATH_TO_EXCEL)
    worksheets: List[Worksheet] = workbook.worksheets
    legacy_sheet: Worksheet = worksheets[0]
    scalable_sheet: Worksheet = worksheets[1]
    mapping_sheet: Worksheet = worksheets[2]

    legacy: DataFrame = pd.read_excel(PATH_TO_EXCEL, legacy_sheet.title)
    scalable: DataFrame = pd.read_excel(PATH_TO_EXCEL, scalable_sheet.title)
    mapping: DataFrame = pd.read_excel(PATH_TO_EXCEL, mapping_sheet.title)
    print()