def compute_layer_vsi(layer):

    soil = layer["soil_type"]
    n1 = layer["n1"]
    fines = layer["fines_less_than_15"]
    user_vsi = layer["vsi"]

    # Always use user Vsi for Others
    if soil == "Others":
        return user_vsi

    # Correlation not valid when N1 < 10
    if n1 < 10:
        return user_vsi

    # Determine exponent
    if soil == "Dry Sands":
        exponent = 0.5 if fines == "Yes" else 0.3

    elif soil == "Saturated Sands":
        exponent = 0.4 if fines == "Yes" else 0.3

    elif soil == "Clays":
        exponent = 0.3

    else:
        return user_vsi

    return 80 * (n1 ** exponent)


def compute_weighted_vs(site_data):
    """
    Computes weighted harmonic average Vs
    considering depth of influence truncation.
    Returns weighted_vs, layers_used, breakdown list.
    """

    depth = site_data["depth_of_influence"]
    layers = site_data["layers"]

    cumulative_depth = 0
    numerator = 0
    denominator = 0
    breakdown = []

    for layer in layers:

        ti_original = layer["thickness"]

        # Apply truncation if needed
        if cumulative_depth + ti_original > depth:
            ti = depth - cumulative_depth
        else:
            ti = ti_original

        vsi = compute_layer_vsi(layer)

        if vsi <= 0:
            raise ValueError("Vsi must be greater than zero.")

        ti_over_vsi = ti / vsi

        numerator += ti
        denominator += ti_over_vsi
        cumulative_depth += ti

        breakdown.append({
            "layer": layer["layer"],
            "effective_thickness": ti,
            "soil_type": layer["soil_type"],
            "fines": layer["fines_less_than_15"],
            "n1": layer["n1"],
            "computed_vsi": vsi,
            "ti_over_vsi": ti_over_vsi
        })

        if cumulative_depth >= depth:
            break

    if cumulative_depth < depth:
        raise ValueError(
            "Total thickness is less than depth of influence."
        )

    if denominator == 0:
        raise ValueError("Invalid denominator in Vs calculation.")

    weighted_vs = numerator / denominator

    return weighted_vs, len(breakdown), breakdown
 



def determine_site_class(weighted_vs):
    """
    Determines site class per Table 4.
    """

    if weighted_vs >= 1500:
        return "A"
    elif weighted_vs >= 760:
        return "B"
    elif weighted_vs >= 360:
        return "C"
    elif weighted_vs >= 180:
        return "D"
    else:
        return "E"


def calculate_site_class(site_data):
    """
    Main backend engine.
    """

    weighted_vs, layers_used, breakdown = compute_weighted_vs(site_data)

    site_class = determine_site_class(weighted_vs)

    return {
        "weighted_vs": weighted_vs,
        "site_class": site_class,
        "layers_used": layers_used,
        "breakdown": breakdown
    }


