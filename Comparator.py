from typing import Dict, Tuple, List

from openpyxl.worksheet.worksheet import Worksheet
from pandas import DataFrame
import pandas as pd


class Comparator:

    def compare(self, path: str, left_table: Worksheet, right_table: Worksheet, mapper_table: Worksheet):
        left: DataFrame = pd.read_excel(path, left_table.title)
        right: DataFrame = pd.read_excel(path, right_table.title)
        mapping: DataFrame = pd.read_excel(path, mapper_table.title)

        mapper_list: List[Tuple[str, str]] = list()

        # Build mapper list
        for i in range(len(mapping)):
            mapper_list.append((mapping.values[i][0], mapping.values[i][1]))

        # Compare columns by row
        for l, r in mapper_list:
            result = pd.DataFrame([left[l].isin([right[r]])]).T

            # Report to xlsx
            output = pd.ExcelWriter(f"Compare_{l}_{r}.xlsx")
            result.to_excel(output)
            output.save()

        # Compare by unique values
        for l, r in mapper_list:
            lu = left[l].unique().T
            ru = right[r].unique().T
            rl = pd.Series(lu, name=str([left[l]])).array
            rr = pd.Series(ru, name=str([right[r]])).array

            # Swap to make sure rr is the longest
            if len(rr) < len(rl):
                temp = rr
                rr = rl
                rl = temp

            result = pd.DataFrame([rr.T, rr.isin(rl).T]).T

            # Report to xlsx
            output = pd.ExcelWriter(f"Unique {l}.xlsx")
            result.to_excel(output)
            output.save()
