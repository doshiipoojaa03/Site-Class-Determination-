from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from backend import compute_layer_vsi, compute_weighted_vs, determine_site_class


def generate_site_class_report(site_data, file_path="Site_Class_Report.xlsx"):

    wb = Workbook()
    ws = wb.active
    ws.title = "Site Class Report"

    bold = Font(bold=True)
    center = Alignment(horizontal="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    row_cursor = 1

    ws.cell(row=row_cursor, column=1, value="Report").font = bold
    row_cursor += 2

    # ---------------- LAYER DETAILS ----------------

    for layer in site_data["layers"]:

        ws.cell(row=row_cursor, column=1, value=f"Layer {layer['layer']}").font = bold
        row_cursor += 1

        ti = layer["thickness"]
        vsi = compute_layer_vsi(layer)

        ws.cell(row=row_cursor, column=1, value="Thickness ti:")
        ws.cell(row=row_cursor, column=2, value=ti)
        ws.cell(row=row_cursor, column=3, value="m")
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="Soil Type:")
        ws.cell(row=row_cursor, column=2, value=layer["soil_type"])
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="Fines < 15%:")
        ws.cell(row=row_cursor, column=2, value=layer["fines_less_than_15"])
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="(N1)60:")
        ws.cell(row=row_cursor, column=2, value=layer["n1"])
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="Vsi:")
        ws.cell(row=row_cursor, column=2, value=round(vsi, 3))
        ws.cell(row=row_cursor, column=3, value="m/s")
        row_cursor += 1

        ws.cell(row=row_cursor, column=1, value="ti / Vsi:")
        ws.cell(row=row_cursor, column=2, value=round(ti / vsi, 8))
        row_cursor += 2

    # ---------------- FINAL CALCULATION ----------------

    weighted_vs, _ = compute_weighted_vs(site_data)

    ws.cell(row=row_cursor, column=1, value="Weighted Vs =").font = bold
    ws.cell(row=row_cursor, column=2, value=round(weighted_vs, 3))
    ws.cell(row=row_cursor, column=3, value="m/s")
    row_cursor += 2

    site_class = determine_site_class(weighted_vs)

    ws.cell(row=row_cursor, column=1, value="Site Class =").font = bold
    ws.cell(row=row_cursor, column=2, value=site_class)
    row_cursor += 2

    # ---------------- TABLE 4 ----------------

    ws.cell(row=row_cursor, column=1, value="Table 4 Site Classes").font = bold
    row_cursor += 1

    table_data = [
        ["A", "Vs ≥ 1500"],
        ["B", "760 ≤ Vs < 1500"],
        ["C", "360 ≤ Vs < 760"],
        ["D", "180 ≤ Vs < 360"],
        ["E", "Vs ≤ 180"],
    ]

    for row_data in table_data:
        ws.cell(row=row_cursor, column=1, value=row_data[0])
        ws.cell(row=row_cursor, column=2, value=row_data[1])
        row_cursor += 1

    wb.save(file_path)
    return file_path
