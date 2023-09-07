from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

from .config import movement_names


def _populate_table(sheet, data, name: str, level: str, plane: str, initial_row: int, initial_col: str) -> int:
    def col(_i):
        return chr(ord(initial_col) + _i)

    if level not in ("kinematics", "moment", "power"):
        raise ValueError("Wrong level")
    if plane not in ("sagittal", "frontal", "transversal"):
        raise ValueError("Wrong plane")

    names = movement_names[level][plane]

    sheet[f"{col(0)}{initial_row}"] = name
    for i, value in enumerate(names):
        sheet[f"{col(0)}{i + initial_row + 1}"] = value["fr"]
        sheet[f"{col(1)}{i + initial_row + 1}"] = round(data[names[i]["en"]]["Minimum Appui"], 2)
        sheet[f"{col(2)}{i + initial_row + 1}"] = round(data[names[i]["en"]]["Maximum Appui"], 2)
        sheet[f"{col(3)}{i + initial_row + 1}"] = round(data[names[i]["en"]]["Range Appui"], 2)
        sheet[f"{col(4)}{i + initial_row + 1}"] = round(data[names[i]["en"]]["Minimum Oscillation"], 2)
        sheet[f"{col(5)}{i + initial_row + 1}"] = round(data[names[i]["en"]]["Maximum Oscillation"], 2)
        sheet[f"{col(6)}{i + initial_row + 1}"] = round(data[names[i]["en"]]["Range Oscillation"], 2)

    # Formatting
    sheet[f"{col(0)}{initial_row}"].font = Font(bold=True)

    # Return the last row
    return initial_row + len(movement_names[level][plane])


def generate_excel(data: dict, save_file: str):
    workbook = Workbook()
    sheet = workbook.active

    _prepare_main_header(sheet)

    last_row = {}
    for initial_col, level in zip(("B", "J", "R"), ("kinematics", "moment", "power")):
        initial_row = 8
        last_row[level] = {}
        for name, plane in zip(("Sagittal", "Frontal", "Transversal"), ("sagittal", "frontal", "transversal")):
            last_row[level][plane] = _populate_table(
                sheet,
                data=data,
                name=name,
                level=level,
                plane=plane,
                initial_row=initial_row,
                initial_col=initial_col,
            )
            initial_row = last_row[level][plane] + 1

    # Center text in the cell
    alignment = Alignment(horizontal="center", vertical="center")
    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = alignment

    cells_to_border_1 = []  # Borders
    cells_to_border_2 = []
    cells_to_border_3 = []
    border_1 = Border(top=Side(border_style="thick", color="000000"))
    border_2 = Border(left=Side(border_style="thick", color="000000"))
    border_3 = Border(bottom=Side(border_style="medium", color="000000"))

    # First table
    for l in [2, 5, 8, last_row["kinematics"]["transversal"] + 1]:
        for m in range(2, 9):
            cells_to_border_1.append(chr(64 + m) + str(l))

    for l in range(2, last_row["kinematics"]["transversal"] + 1):
        for m in ["B", "I"]:
            cells_to_border_2.append(m + str(l))

    for l in range(5, last_row["kinematics"]["transversal"] + 1):
        cells_to_border_2.append("C" + str(l))

    for l in range(8, last_row["kinematics"]["transversal"] + 1):
        cells_to_border_2.append("F" + str(l))

    for m in range(2, 9):
        cells_to_border_3.append(chr(64 + m) + str(last_row["kinematics"]["sagittal"]))
        cells_to_border_3.append(chr(64 + m) + str(last_row["kinematics"]["frontal"]))

    # Second table
    for l in [5, 8, last_row["moment"]["transversal"] + 1]:
        for m in range(10, 17):
            cells_to_border_1.append(chr(64 + m) + str(l))

    for l in range(5, last_row["moment"]["transversal"] + 1):
        for m in ["J", "K", "Q"]:
            cells_to_border_2.append(m + str(l))

    for l in range(8, last_row["moment"]["transversal"] + 1):
        cells_to_border_2.append("N" + str(l))

    for m in range(10, 17):
        cells_to_border_3.append(chr(64 + m) + str(last_row["moment"]["sagittal"]))
        cells_to_border_3.append(chr(64 + m) + str(last_row["moment"]["frontal"]))

    # Third table
    for l in [5, 8, last_row["power"]["transversal"] + 1]:
        for m in range(18, 25):
            cells_to_border_1.append(chr(64 + m) + str(l))

    for l in range(5, last_row["power"]["transversal"] + 1):
        for m in ["R", "S", "Y"]:
            cells_to_border_2.append(m + str(l))

    for l in range(8, last_row["power"]["transversal"] + 1):
        cells_to_border_2.append("V" + str(l))

    for m in range(18, 25):
        cells_to_border_3.append(chr(64 + m) + str(last_row["power"]["sagittal"]))
        cells_to_border_3.append(chr(64 + m) + str(last_row["power"]["frontal"]))

    # Adding border lines
    for cell in cells_to_border_1:
        sheet[cell].border = border_1

    for cell in cells_to_border_2:
        existing_border = sheet[cell].border
        left_border = existing_border.left
        right_border = existing_border.right
        top_border = existing_border.top
        bottom_border = existing_border.bottom
        new_border = Border(top=top_border, bottom=bottom_border,
                            left=border_2.left, right=right_border)
        sheet[cell].border = new_border

    for cell in cells_to_border_3:
        existing_border = sheet[cell].border
        left_border = existing_border.left
        right_border = existing_border.right
        top_border = existing_border.top
        bottom_border = existing_border.bottom
        new_border = Border(top=top_border, bottom=border_3.bottom,
                            left=left_border, right=right_border)
        sheet[cell].border = new_border

    for column in sheet.columns:  # Adjust cell dimension
        sheet.column_dimensions[column[0].column_letter].width = 12
        sheet.column_dimensions["B"].width = 40
        sheet.column_dimensions["J"].width = 40
        sheet.column_dimensions["R"].width = 40

    workbook.save(save_file)


def _prepare_main_header(sheet):
    sheet["B2"] = "Patient ID"
    sheet["B3"] = "Session ID"
    sheet["B4"] = "Date"
    sheet["C5"] = "Cinématique"
    sheet["C6"] = "Phase d'appui"
    sheet["F6"] = "Phase d'oscillation"
    sheet["C7"] = "Min (deg)"
    sheet["D7"] = "Max (deg)"
    sheet["E7"] = "Range(deg)"
    sheet["F7"] = "Min (deg)"
    sheet["G7"] = "Max (deg)"
    sheet["H7"] = "Range(deg)"

    sheet["K5"] = "Moment"
    sheet["K6"] = "Phase d'appui"
    sheet["N6"] = "Phase d'oscillation"
    sheet["K7"] = "Min (deg)"
    sheet["L7"] = "Max (deg)"
    sheet["M7"] = "Range(deg)"
    sheet["N7"] = "Min (deg)"
    sheet["O7"] = "Max (deg)"
    sheet["P7"] = "Range(deg)"

    sheet["S5"] = "Power"
    sheet["S6"] = "Phase d'appui"
    sheet["V6"] = "Phase d'oscillation"
    sheet["S7"] = "Min (deg)"
    sheet["T7"] = "Max (deg)"
    sheet["U7"] = "Range(deg)"
    sheet["V7"] = "Min (deg)"
    sheet["W7"] = "Max (deg)"
    sheet["X7"] = "Range(deg)"

    # Formatting
    sheet.merge_cells("C2:H2")
    sheet.merge_cells("C3:H3")
    sheet.merge_cells("C4:H4")

    sheet.merge_cells("C5:H5")
    sheet["C5"].font = Font(bold=True)
    sheet.merge_cells("C6:E6")
    sheet.merge_cells("F6:H6")
    sheet["C6"].font = Font(bold=True)
    sheet["F6"].font = Font(bold=True)

    sheet.merge_cells("K5:P5")
    sheet["K5"].font = Font(bold=True)
    sheet.merge_cells("K6:M6")
    sheet.merge_cells("N6:P6")
    sheet["K6"].font = Font(bold=True)
    sheet["N6"].font = Font(bold=True)

    sheet.merge_cells("S5:X5")
    sheet["S5"].font = Font(bold=True)
    sheet.merge_cells("S6:U6")
    sheet.merge_cells("V6:X6")
    sheet["S6"].font = Font(bold=True)
    sheet["V6"].font = Font(bold=True)
