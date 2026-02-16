def compute_layer_vsi(layer):
    """
    Computes Vsi for a single layer based on IS 1893:2025 rules.
    """

    soil = layer["soil_type"]
    n1 = layer["n1"]
    fines = layer["fines_less_than_15"]
    user_vsi = layer["vsi"]

    # Case 1: Others → always user input
    if soil == "Others":
        return user_vsi

    # Case 2: N1 < 10 → correlation not applicable
    if n1 < 10:
        return user_vsi

    # Case 3: Correlation applies

    if soil == "Dry Sands":
        if fines == "Yes":
            exponent = 0.5
        else:
            exponent = 0.3

    elif soil == "Saturated Sands":
        if fines == "Yes":
            exponent = 0.4
        else:
            exponent = 0.3

    elif soil == "Clays":
        exponent = 0.3

    else:
        # fallback safety
        return user_vsi

    return 80 * (n1 ** exponent)


def compute_weighted_vs(site_data):
    """
    Computes weighted harmonic average Vs
    considering depth of influence truncation.
    """

    depth = site_data["depth_of_influence"]
    layers = site_data["layers"]

    cumulative_depth = 0
    numerator = 0
    denominator = 0
    layers_used = 0

    for layer in layers:

        ti = layer["thickness"]

        # If adding this layer exceeds depth → truncate
        if cumulative_depth + ti > depth:
            ti = depth - cumulative_depth

        vsi = compute_layer_vsi(layer)

        if vsi <= 0:
            raise ValueError("Vsi must be greater than zero.")

        numerator += ti
        denominator += ti / vsi

        cumulative_depth += ti
        layers_used += 1

        if cumulative_depth >= depth:
            break

    if cumulative_depth < depth:
        raise ValueError(
            "Total thickness is less than depth of influence."
        )

    if denominator == 0:
        raise ValueError("Invalid denominator in Vs calculation.")

    weighted_vs = numerator / denominator

    return weighted_vs, layers_used



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

    weighted_vs, layers_used = compute_weighted_vs(site_data)
    site_class = determine_site_class(weighted_vs)

    return {
        "weighted_vs": weighted_vs,
        "site_class": site_class,
        "layers_used": layers_used
    }

