"""
ExcelExporter: export RecipeTree instances to an .xlsx workbook.
"""

import os
import logging
from openpyxl import Workbook
from core.base import EXCEL_COLUMNS


class ExcelExporter:
    """
    Export parameters and formula values into an Excel workbook.
    """

    def export(self, trees, excel_path: str):
        """
        Write one sheet per RecipeTree.filepath.
        """
        log = logging.getLogger(__name__)
        if not isinstance(trees, list):
            trees = [trees]
        wb = Workbook()
        # remove default sheet
        wb.remove(wb.active)

        all_extras = set()
        # collect rows per sheet
        sheet_data = {}
        for t in trees:
            rows = []
            for node in t.parameters + t.formula_values:
                row = node.to_excel_row()
                rows.append(row)
                all_extras.update(k for k in row if k not in EXCEL_COLUMNS)
            sheet = os.path.basename(t.filepath)
            sheet_data[sheet] = rows
            log.info("Prepared %d rows for sheet %s", len(rows), sheet)

        extras = sorted(all_extras)
        header = EXCEL_COLUMNS + extras

        for sheet, rows in sheet_data.items():
            ws = wb.create_sheet(sheet)
            ws.append(header)
            for row in rows:
                ws.append([row.get(col, "") for col in header])

        wb.save(excel_path)
        log.info("Excel written to %s", excel_path)
