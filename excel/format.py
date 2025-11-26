from openpyxl.utils import get_column_letter

def autofit_selected_columns(ws, column):
    for num in column:
        cellRef = get_column_letter(num)
        cell = ws[f"{cellRef}1"]
        length = len(str(cell.value)) + 2
        ws.column_dimensions[get_column_letter(num)].width = length