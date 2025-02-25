from openpyxl import load_workbook
import re

def extract_refrences(formula):
    ref_pattern = r"""
        (?:'([^']+)'!)?                 # ('Sheet Name'!)
        (?:\[([^\])+)\])?               # ([Workbook.xlsx])
        ([A-Z]+[0-9]+(:[A-Z]+[0-9]+)?)  # (A1, A1:B10)

    """

    named_range_pattern = r"\b[A-Za-z_]+\b(?!\d|\()"

    matches = re.findall(ref_pattern, formula, re.VERBOSE)
    named_ranges = re.findall(named_range_pattern, formula)

    refs = []
    for match in matches:
        sheet, workbook, cell_ref, _ = match
        ref = {"cell": cell_ref}
        if sheet:
            ref["sheet"] = sheet
        if workbook:
            ref["workbook"] = workbook
        refs.append(ref)

    return refs + [{"named_range": name} for name in named_ranges if name.upper() not in ["SUM", "AVERAGE", "IF"]]


def parse_excel_formulas(sheet):
    formulas = {}

    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                formulas[cell.coordinate] = {
                    "formula": cell.value,
                    "refrences": extract_refrences(cell.value),
                }

    return formulas

def parse_all_sheets(file_path):
    wb = load_workbook(file_path, data_only=False)

    all_refs = {}

    for sheet in wb.sheetnames:
        all_refs[sheet] = {
            "items": parse_excel_formulas(wb[sheet]),
        }

def parse_single_sheet(file_path, sheet_name):
    wb = load_workbook(file_path, data_only=False)
    sheet = wb[sheet_name]

    return parse_excel_formulas(sheet)
