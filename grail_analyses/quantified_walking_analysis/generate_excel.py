from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

from .config import movement_names


def generate_excel(data: dict, save_file: str):
    # EXCEL
    workbook = Workbook()
    sheet = workbook.active
    # Titre de cellules
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

    # Titre des variables

    # Kinematics
    i1 = 8
    sheet["B" + str(i1)] = "Sagittal"

    count = 0
    names = movement_names["kinematics"]["sagittal"]
    for i1, value in enumerate(names, start=i1 + 1):
        sheet["B" + str(i1)] = value["fr"]
        sheet["C" + str(i1)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["D" + str(i1)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["E" + str(i1)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["F" + str(i1)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["G" + str(i1)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["H" + str(i1)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    j1 = i1 + 1
    sheet["B" + str(j1)] = "Frontal"

    count = 0
    names = movement_names["kinematics"]["frontal"]
    for j1, value in enumerate(names, start=(j1 + 1)):
        sheet["B" + str(j1)] = value["fr"]
        sheet["C" + str(j1)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["D" + str(j1)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["E" + str(j1)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["F" + str(j1)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["G" + str(j1)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["H" + str(j1)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    k1 = j1 + 1
    sheet["B" + str(k1)] = "Transversal"

    count = 0
    names = movement_names["kinematics"]["transversal"]
    for k1, value in enumerate(names, start=(k1 + 1)):
        sheet["B" + str(k1)] = value["fr"]
        sheet["C" + str(k1)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["D" + str(k1)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["E" + str(k1)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["F" + str(k1)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["G" + str(k1)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["H" + str(k1)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    # Moment
    i2 = 8
    sheet["J" + str(i2)] = "Sagittal"

    count = 0
    names = movement_names["moment"]["sagittal"]
    for i2, value in enumerate(names, start=i2 + 1):
        sheet["J" + str(i2)] = value["fr"]
        sheet["K" + str(i2)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["L" + str(i2)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["M" + str(i2)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["N" + str(i2)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["O" + str(i2)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["P" + str(i2)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    j2 = i2 + 1
    sheet["J" + str(j2)] = "Frontal"

    count = 0
    j2 = i2 + 1
    names = movement_names["moment"]["frontal"]
    for j2, value in enumerate(names, start=(j2 + 1)):
        sheet["J" + str(j2)] = value["fr"]
        sheet["K" + str(j2)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["L" + str(j2)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["M" + str(j2)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["N" + str(j2)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["O" + str(j2)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["P" + str(j2)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    k2 = j2 + 1
    sheet["J" + str(k2)] = "Transversal"

    count = 0
    names = movement_names["moment"]["transversal"]
    for k2, value in enumerate(names, start=(k2 + 1)):
        sheet["J" + str(k2)] = value["fr"]
        sheet["K" + str(k2)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["L" + str(k2)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["M" + str(k2)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["N" + str(k2)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["O" + str(k2)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["P" + str(k2)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    # Power
    i3 = 8
    sheet["R" + str(i3)] = "Sagittal"

    count = 0
    names = movement_names["power"]["sagittal"]
    for i3, value in enumerate(names, start=i3 + 1):
        sheet["R" + str(i3)] = value["fr"]
        sheet["S" + str(i3)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["T" + str(i3)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["U" + str(i3)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["V" + str(i3)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["W" + str(i3)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["X" + str(i3)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    j3 = i3 + 1
    sheet["R" + str(j3)] = "Frontal"

    count = 0
    j3 = i3 + 1
    names = movement_names["power"]["frontal"]
    for j3, value in enumerate(names, start=(j3 + 1)):
        sheet["R" + str(j3)] = value["fr"]
        sheet["S" + str(j3)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["T" + str(j3)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["U" + str(j3)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["V" + str(j3)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["W" + str(j3)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["X" + str(j3)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    k3 = j3 + 1
    sheet["R" + str(k3)] = "Transversal"

    count = 0
    names = movement_names["power"]["transversal"]
    for k3, value in enumerate(names, start=(k3 + 1)):
        sheet["R" + str(k3)] = value["fr"]
        sheet["S" + str(k3)] = round(data[names[count]["en"]]["Minimum Appui"], 2)
        sheet["T" + str(k3)] = round(data[names[count]["en"]]["Maximum Appui"], 2)
        sheet["U" + str(k3)] = round(data[names[count]["en"]]["Range Appui"], 2)
        sheet["V" + str(k3)] = round(data[names[count]["en"]]["Minimum Oscillation"], 2)
        sheet["W" + str(k3)] = round(data[names[count]["en"]]["Maximum Oscillation"], 2)
        sheet["X" + str(k3)] = round(data[names[count]["en"]]["Range Oscillation"], 2)
        count += 1

    # Mise en forme
    sheet.merge_cells("C2:H2")  # Merge de cellules
    sheet.merge_cells("C3:H3")
    sheet.merge_cells("C4:H4")
    sheet.merge_cells("C5:H5")
    sheet.merge_cells("K5:P5")
    sheet.merge_cells("S5:X5")
    sheet.merge_cells("C6:E6")
    sheet.merge_cells("F6:H6")
    sheet.merge_cells("K6:M6")
    sheet.merge_cells("N6:P6")
    sheet.merge_cells("S6:U6")
    sheet.merge_cells("V6:X6")

    sheet["C5"].font = Font(bold=True)  # Mise de titres en gras
    sheet["K5"].font = Font(bold=True)
    sheet["S5"].font = Font(bold=True)
    sheet["C6"].font = Font(bold=True)
    sheet["F6"].font = Font(bold=True)
    sheet["K6"].font = Font(bold=True)
    sheet["N6"].font = Font(bold=True)
    sheet["S6"].font = Font(bold=True)
    sheet["V6"].font = Font(bold=True)
    sheet["B8"].font = Font(bold=True)  # Sagittal
    sheet["J8"].font = Font(bold=True)
    sheet["R8"].font = Font(bold=True)
    sheet["B" + str(i1 + 1)].font = Font(bold=True)  # Frontal
    sheet["J" + str(i2 + 1)].font = Font(bold=True)
    sheet["R" + str(i3 + 1)].font = Font(bold=True)
    sheet["B" + str(j1 + 1)].font = Font(bold=True)  # Transversal
    sheet["J" + str(j2 + 1)].font = Font(bold=True)
    sheet["R" + str(j3 + 1)].font = Font(bold=True)

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
    for l in [2, 5, 8, k1 + 1]:
        for m in range(2, 9):
            cells_to_border_1.append(chr(64 + m) + str(l))

    for l in range(2, k1 + 1):
        for m in ["B", "I"]:
            cells_to_border_2.append(m + str(l))

    for l in range(5, k1 + 1):
        cells_to_border_2.append("C" + str(l))

    for l in range(8, k1 + 1):
        cells_to_border_2.append("F" + str(l))

    for m in range(2, 9):
        cells_to_border_3.append(chr(64 + m) + str(i1))
        cells_to_border_3.append(chr(64 + m) + str(j1))

    # Second table
    for l in [5, 8, k2 + 1]:
        for m in range(10, 17):
            cells_to_border_1.append(chr(64 + m) + str(l))

    for l in range(5, k2 + 1):
        for m in ["J", "K", "Q"]:
            cells_to_border_2.append(m + str(l))

    for l in range(8, k2 + 1):
        cells_to_border_2.append("N" + str(l))

    for m in range(10, 17):
        cells_to_border_3.append(chr(64 + m) + str(i2))
        cells_to_border_3.append(chr(64 + m) + str(j2))

    # Third table
    for l in [5, 8, k3 + 1]:
        for m in range(18, 25):
            cells_to_border_1.append(chr(64 + m) + str(l))

    for l in range(5, k3 + 1):
        for m in ["R", "S", "Y"]:
            cells_to_border_2.append(m + str(l))

    for l in range(8, k3 + 1):
        cells_to_border_2.append("V" + str(l))

    for m in range(18, 25):
        cells_to_border_3.append(chr(64 + m) + str(i2))
        cells_to_border_3.append(chr(64 + m) + str(j2))

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