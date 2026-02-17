from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def build_formula_text(layer):
    soil = layer["soil_type"]
    n1 = layer["n1"]
    fines = layer["fines"]

    if soil == "Others" or n1 < 10:
        return "NA"

    if soil == "Dry Sands":
        exponent = "0.5" if fines == "Yes" else "0.3"
    elif soil == "Saturated Sands":
        exponent = "0.4" if fines == "Yes" else "0.3"
    elif soil == "Clays":
        exponent = "0.3"
    else:
        return "NA"

    return f"80[(N1)60i]^{exponent}"


def highlight_site_class(ws, site_class):
    fill = PatternFill(start_color="FFF59D",
                       end_color="FFF59D",
                       fill_type="solid")

    class_row_map = {
        "A": 11,
        "B": 12,
        "C": 13,
        "D": 14,
        "E": 15
    }

    for row in range(11, 16):
        ws[f"AG{row}"].fill = PatternFill()
        ws[f"AH{row}"].fill = PatternFill()

    target_row = class_row_map.get(site_class)

    if target_row:
        ws[f"AG{target_row}"].fill = fill
        ws[f"AH{target_row}"].fill = fill


def generate_site_class_report(
    site_data,
    result,
    template_path="Site Class Report Spec.xlsx",
    output_path="Site_Class_Report.xlsx"
):

    wb = load_workbook(template_path)
    ws = wb.active

    # -------------------------
    # Top Summary Update
    # -------------------------

    ws["D3"] = result["weighted_vs"]
    ws["D4"] = result["site_class"]

    # -------------------------
    # Depth
    # -------------------------

    ws["H7"] = site_data["depth_of_influence"]

    # -------------------------
    # Dynamic Layer Rows
    # -------------------------

    start_row = 11
    original_template_rows = 6   # template initially has 6 example rows
    layer_count = len(result["breakdown"])

    # Remove template placeholder rows
    ws.delete_rows(start_row, original_template_rows)

    # Insert exact number of rows
    ws.insert_rows(start_row, layer_count)

    # Fill layer rows
    for i, layer in enumerate(result["breakdown"]):
        row = start_row + i

        ws[f"B{row}"] = f"Layer {layer['layer']}"
        ws[f"E{row}"] = layer["effective_thickness"]
        ws[f"I{row}"] = layer["soil_type"]
        ws[f"N{row}"] = layer["fines"] if layer["fines"] else "NA"
        ws[f"Q{row}"] = layer["n1"] if layer["n1"] else "NA"
        ws[f"T{row}"] = build_formula_text(layer)
        ws[f"Y{row}"] = layer["computed_vsi"]
        ws[f"AB{row}"] = layer["ti_over_vsi"]

    # -------------------------
    # Dynamic Summary Rows
    # -------------------------

    sum_row = start_row + layer_count
    final_vs_row = sum_row + 1
    final_class_row = final_vs_row + 1

    # Σti
    ws[f"E{sum_row}"] = f"=SUM(E{start_row}:E{start_row + layer_count - 1})"

    # Σ(ti/Vsi)
    ws[f"AB{sum_row}"] = f"=SUM(AB{start_row}:AB{start_row + layer_count - 1})"

    # Final Vs (use backend value)
    write_to_cell(ws, f"AL{final_vs_row}", result["weighted_vs"])

    # Bottom Site Class
    write_to_cell(ws, f"AM{final_class_row}", result["site_class"])

    # -------------------------
    # Highlight Table 4
    # -------------------------

    highlight_site_class(ws, result["site_class"])

    wb.save(output_path)
    return output_path


def write_to_cell(ws, cell_address, value):
    cell = ws[cell_address]

    # If cell is merged, find top-left of merged range
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            top_left = ws.cell(
                row=merged_range.min_row,
                column=merged_range.min_col
            )
            top_left.value = value
            return

    # If not merged
    cell.value = value

