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
        Write out one Excel sheet per RecipeTree, exporting all parameters and formula values.

        This method accepts either a single RecipeTree or a list of them, then removes the default
        OpenPyXL sheet and prepares a new workbook.  It iterates each tree, calls each node's
        `to_excel_row()` to build a flat dict, and collects any extra columns beyond the fixed
        schema.  After logging how many rows each sheet will have, it writes a header row (fixed
        columns + sorted extras) followed by one row per node.  Finally, it saves the workbook to
        `excel_path` and logs the completion.  Errors during file creation or serialization raise
        the underlying exception for the caller to handle.
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
