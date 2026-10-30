"""Regenerate independent typed array input (requires XlsxWriter).

An optional first argument overrides the output path. Output caches are stale
on purpose; tests must calculate from formulas instead of accepting -999.
"""
from pathlib import Path
import sys
import xlsxwriter

target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_suffix('.xlsx')
target.parent.mkdir(parents=True, exist_ok=True)
book = xlsxwriter.Workbook(target)
for name, values in [('Mixed', [True, False, 'error', 2, '', None]),
                     ('Logical', [True, False, 2, '', None])]:
    sheet = book.add_worksheet(name)
    for row, value in enumerate(values):
        if isinstance(value, bool):
            sheet.write_boolean(row, 0, value)
        elif value == 'error':
            sheet.write_formula(row, 0, '=NA()', None, '#N/A')
        elif isinstance(value, str):
            sheet.write_string(row, 0, value)
        elif value is not None:
            sheet.write_number(row, 0, value)
        sheet.write_number(row, 1, 1)
    n = len(values)
    formulas = [f'=SORT(A1:A{n})', f'=FILTER(A1:A{n},B1:B{n})',
                f'=UNIQUE(A1:A{n})', f'=IFERROR(A1:A{n},0)',
                f'=ISLOGICAL(FILTER(A1:A{n},B1:B{n}))']
    for col, formula in zip([3, 5, 7, 9, 11], formulas):
        sheet.write_dynamic_array_formula(0, col, n - 1, col, formula, None, -999)
    sheet.write_formula('N1', '=SUM(1,2)', None, -999)
book.close()
