from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def generate_site_class_report(site_data, result,
                               file_path="Site_Class_Report.xlsx"):

    wb = Workbook()
    ws = wb.active
    ws.title = "Site Class Report"

    # ---------------- STYLES ----------------

    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    left_align = Alignment(horizontal="left")

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    row_cursor = 1

    # ---------------- TITLE ----------------

    ws.cell(row=row_cursor, column=1,
            value="Site Class Determination Report").font = bold
    row_cursor += 2

    # ---------------- DEPTH INFO ----------------

    ws.cell(row=row_cursor, column=1, value="Depth of Influence (m):")
    ws.cell(row=row_cursor, column=2,
            value=site_data["depth_of_influence"])
    row_cursor += 2

    # ---------------- LAYER DETAILS ----------------

    for layer in result["breakdown"]:

        ws.cell(row=row_cursor, column=1,
                value=f"Layer {layer['layer']}").font = bold
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="Thickness (ti)")
        ws.cell(row=row_cursor, column=2,
                value=round(layer["effective_thickness"], 4))
        ws.cell(row=row_cursor, column=3, value="m")
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="Soil Type")
        ws.cell(row=row_cursor, column=2,
                value=layer["soil_type"])
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="Fines < 15%")
        ws.cell(row=row_cursor, column=2,
                value=layer["fines"])
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="(N1)60")
        ws.cell(row=row_cursor, column=2,
                value=layer["n1"])
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="Computed Vsi")
        ws.cell(row=row_cursor, column=2,
                value=round(layer["computed_vsi"], 4))
        ws.cell(row=row_cursor, column=3, value="m/s")
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="ti / Vsi")
        ws.cell(row=row_cursor, column=2,
                value=round(layer["ti_over_vsi"], 8))
        row_cursor += 2

    # ---------------- HARMONIC FORMULA DISPLAY ----------------

    ws.cell(row=row_cursor, column=1,
            value="Weighted Vs Formula:").font = bold
    row_cursor += 1

    ws.cell(row=row_cursor, column=1,
            value="Vs = Σti / Σ(ti / Vsi)")
    row_cursor += 2

    # ---------------- FINAL RESULT ----------------

    ws.cell(row=row_cursor, column=1,
            value="Weighted Average Vs").font = bold
    ws.cell(row=row_cursor, column=2,
            value=round(result["weighted_vs"], 4))
    ws.cell(row=row_cursor, column=3, value="m/s")
    row_cursor += 2

    ws.cell(row=row_cursor, column=1,
            value="Site Class").font = bold
    ws.cell(row=row_cursor, column=2,
            value=result["site_class"])
    row_cursor += 2

    # ---------------- TABLE 4 ----------------

    ws.cell(row=row_cursor, column=1,
            value="Table 4 - Site Classes").font = bold
    row_cursor += 1

    table_data = [
        ["A", "Vs ≥ 1500"],
        ["B", "760 ≤ Vs < 1500"],
        ["C", "360 ≤ Vs < 760"],
        ["D", "180 ≤ Vs < 360"],
        ["E", "Vs ≤ 180"],
    ]

    for entry in table_data:
        ws.cell(row=row_cursor, column=1, value=entry[0])
        ws.cell(row=row_cursor, column=2, value=entry[1])
        row_cursor += 1

    # ---------------- AUTO COLUMN WIDTH ----------------

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 22

    wb.save(file_path)
    return file_path
