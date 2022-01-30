from zipfile import Path

import openpyxl
from openpyxl import Workbook


class ExcelHandler:

    def read_workbook(self, path: str) -> Workbook:
        return openpyxl.load_workbook(path)

