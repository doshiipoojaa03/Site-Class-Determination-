import streamlit as st
from site_class_reference import render_site_class_table
from backend import calculate_site_class


st.set_page_config(layout="wide")
st.title("Site Class Determination - IS 1893 : 2025")

# -------------------------
# TOP INPUTS
# -------------------------

col1, col2 = st.columns(2)

with col1:
    depth = st.number_input(
        "Depth of Influence (m)",
        min_value=1.0,
        value=1.0,
        step=0.1
    )

with col2:
    num_layers = st.number_input(
        "Number of Soil Layers",
        min_value=1,
        max_value=20,
        value=1,
        step=1
    )

st.divider()

# -------------------------
# SESSION INITIALIZATION
# -------------------------

if "layers_input" not in st.session_state:
    st.session_state.layers_input = []

# Resize safely if layer count changes
if len(st.session_state.layers_input) != int(num_layers):
    st.session_state.layers_input = [
        {
            "thickness": 1.0,
            "soil_type": "Saturated Sands",
            "fines": "Yes",
            "n_value": 10,
            "vs_value": 180.0
        }
        for _ in range(int(num_layers))
    ]

# -------------------------
# TABLE HEADER
# -------------------------

header = st.columns([1,1,2,1,1,1])
header[0].markdown("**Layer**")
header[1].markdown("**Thickness (m)**")
header[2].markdown("**Soil Type**")
header[3].markdown("**Fines < 15%**")
header[4].markdown("**(N₁)₆₀**")
header[5].markdown("**Vₛᵢ (m/s)**")

# -------------------------
# DYNAMIC ROWS
# -------------------------

for i in range(int(num_layers)):

    row = st.columns([1,1,2,1,1,1])
    row[0].write(f"Layer {i+1}")

    # Thickness (> 0)
    thickness = row[1].number_input(
        "Thickness",
        min_value=0.00,
        value=float(st.session_state.layers_input[i]["thickness"]),
        step=0.1,
        key=f"th_{i}",
        label_visibility="collapsed"
    )

    # Soil Type
    soil_type = row[2].selectbox(
        "Soil Type",
        ["Saturated Sands", "Dry Sands", "Clays", "Others"],
        index=["Saturated Sands", "Dry Sands", "Clays", "Others"]
        .index(st.session_state.layers_input[i]["soil_type"]),
        key=f"soil_{i}",
        label_visibility="collapsed"
    )

    # Fines rule
    disable_fines = soil_type in ["Clays", "Others"]

    fines = row[3].selectbox(
        "Fines",
        ["Yes", "No"],
        key=f"fines_{i}",
        disabled=disable_fines,
        label_visibility="collapsed"
    )

    # N1 rule
    disable_n = soil_type == "Others"

    n_value = row[4].number_input(
        "N1",
        min_value=0,
        step=1,
        value=int(st.session_state.layers_input[i]["n_value"]),
        key=f"n_{i}",
        disabled=disable_n,
        label_visibility="collapsed"
    )

    # Vsi rule
    if soil_type == "Others":
        activate_vs = True
    else:
        activate_vs = n_value < 10

    vs_value = row[5].number_input(
        "Vsi",
        min_value=0.00,
        value=float(st.session_state.layers_input[i]["vs_value"]),
        step=1.0,
        key=f"vs_{i}",
        disabled=not activate_vs,
        label_visibility="collapsed"
    )

    # Update session state live
    st.session_state.layers_input[i] = {
        "thickness": thickness,
        "soil_type": soil_type,
        "fines": "" if disable_fines else fines,
        "n_value": 0 if disable_n else n_value,
        "vs_value": vs_value if activate_vs else 0.0
    }

# st.divider()

# -------------------------
# DEPTH VALIDATION
# -------------------------

total_thickness = sum(layer["thickness"] for layer in st.session_state.layers_input)
valid_depth = total_thickness >= depth

if not valid_depth:
    st.warning("Total thickness must be greater than or equal to Depth of Influence.")

# -------------------------
# BUTTONS (Centered)
# -------------------------

btn1, btn2, btn3 = st.columns([1,1,1])

with btn2:
    calculate_clicked = st.button(
        "Calculate",
        use_container_width=True,
        disabled=not valid_depth
    )

# st.divider()



# ---------------- CALCULATION ----------------

if calculate_clicked:

    # -------------------------
    # BUILD BACKEND DICTIONARY
    # -------------------------

    site_data = {
        "depth_of_influence": float(depth),
        "num_layers": int(num_layers),
        "layers": []
    }

    for i, layer in enumerate(st.session_state.layers_input):

        layer_dict = {
            "layer": i + 1,
            "thickness": float(layer["thickness"]),
            "soil_type": layer["soil_type"],
            "fines_less_than_15": layer["fines"],
            "n1": int(layer["n_value"]),
            "vsi": float(layer["vs_value"])
        }

        site_data["layers"].append(layer_dict)

    # Optional: store it if needed
    st.session_state["site_input_data"] = site_data
    

    result = calculate_site_class(site_data)
    print(result)

    # st.divider()

# ---------------- OUTPUT SECTION ----------------

    output_col1, output_col2 = st.columns(2)

    with output_col1:
        st.metric(
            label="Weighted Average Shear Wave Velocity (Vs)",
            value=round(result["weighted_vs"], 3)
        )

    with output_col2:
        st.success(f"Site Class: {result['site_class']}")

    st.divider()

    # Show reference table below result
    render_site_class_table()

