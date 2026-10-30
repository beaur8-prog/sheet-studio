"""Regenerate xlsxwriter-dynamic.xlsx with the independent XlsxWriter library.

This is an OOXML fixture, not a Microsoft Excel execution oracle. Deliberately
stale follower caches exercise recalculation after import. Run with a Python
installation containing XlsxWriter; it is not a runtime dependency of the app.
"""
from pathlib import Path
import xlsxwriter

with xlsxwriter.Workbook(Path(__file__).with_suffix('.xlsx')) as workbook:
    sheet = workbook.add_worksheet('Arrays')
    sheet.write_dynamic_array_formula('A1:B3', '=SEQUENCE(3,2)', None, 1)
    for row in range(3):
        for col in range(2):
            if row or col:
                sheet.write_number(row, col, row * 2 + col + 1)
    sheet.write_dynamic_array_formula('D1', '=SUM(_xlfn.ANCHORARRAY(A1))', None, 21)
    sheet.write_dynamic_array_formula('F1:F3', '=SORT(UNIQUE({3;1;2;1}))', None, 1)
    sheet.write_number('F2', 0)
    sheet.write_number('F3', 0)
    sheet.write_array_formula('H1:H3', '=A1:A3*2', None, 2)
    sheet.write_number('H2', 0)
    sheet.write_number('H3', 0)
