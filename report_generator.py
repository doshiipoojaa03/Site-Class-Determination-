from openpyxl import load_workbook


def build_formula_text(layer):
    """
    Builds formula string exactly as shown in report format.
    (Display purpose only — no Excel computation)
    """

    soil = layer["soil_type"]
    n1 = layer["n1"]
    fines = layer["fines"]

    if soil == "Others":
        return "NA"

    if n1 < 10:
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


def generate_site_class_report(
    site_data,
    result,
    template_path="Site Class Report Spec.xlsx",
    output_path="Site_Class_Report.xlsx"
):
    """
    Generates Excel report using template.
    Backend remains single source of truth.
    """

    wb = load_workbook(template_path)
    ws = wb.active

    # -------------------------
    # 1️⃣ Depth of Influence
    # -------------------------

    ws["H7"] = site_data["depth_of_influence"]

    # -------------------------
    # 2️⃣ Fill Layer Data
    # -------------------------

    start_row = 11

    for i, layer in enumerate(result["breakdown"]):

        row = start_row + i

        ws[f"B{row}"] = f"Layer {layer['layer']}"
        ws[f"E{row}"] = layer["effective_thickness"]
        ws[f"I{row}"] = layer["soil_type"]

        # Fines column
        ws[f"N{row}"] = layer["fines"] if layer["fines"] else "NA"

        # N1 column
        ws[f"Q{row}"] = layer["n1"] if layer["n1"] else "NA"

        # Formula display column
        ws[f"T{row}"] = build_formula_text(layer)

        # Computed Vsi (NO rounding)
        ws[f"Y{row}"] = layer["computed_vsi"]

        # ti / Vsi (NO rounding)
        ws[f"AB{row}"] = layer["ti_over_vsi"]

    # -------------------------
    # 3️⃣ Final Results
    # -------------------------

    # IMPORTANT:
    # Overwrite any Excel formulas in template.
    # Backend result is authoritative.

    ws["AL17"] = result["weighted_vs"]
    ws["AM18"] = result["site_class"]

    # -------------------------
    # 4️⃣ Save
    # -------------------------

    wb.save(output_path)
    return output_path
