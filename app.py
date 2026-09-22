import io
import json
import os
import openpyxl
import pandas as pd
from rapidfuzz import fuzz, process
import streamlit as st

DB_FILE = "catalog_db.json"

st.set_page_config(page_title="Apparel Multi-Marketplace Hub", layout="wide")


def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            try:
                data = json.load(f)
                # Backward compatibility: convert old single-color entries into color_variants list
                for k, v in data.items():
                    if "color_variants" not in v:
                        v["color_variants"] = [
                            {
                                "color_name": v.get("color", "White"),
                                "color_map": v.get("color", "White"),
                                "images": v.get("images", []),
                            }
                        ]
                return data
            except Exception:
                return {}
    return {}


def save_db(data):
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=2)


db = load_db()

DROPDOWNS = {
    "colors": [
        "White",
        "Black",
        "Multi",
        "Off White",
        "Navy Blue",
        "Mustard",
        "Pink",
        "Red",
        "Green",
        "Yellow",
        "Maroon",
        "Beige",
        "Purple",
        "Grey",
        "Teal",
        "Coral",
        "Other / Custom...",
    ],
    "fabrics": [
        "Pure Cotton",
        "Cotton Blend",
        "Cotton Silk",
        "Silk Blend",
        "Georgette",
        "Chanderi",
        "Rayon",
        "Organza",
        "Tissue",
        "Modal",
        "Satin",
        "Linen Blend",
        "Other / Custom...",
    ],
    "top_patterns": [
        "Embroidered",
        "Printed",
        "Solid",
        "Woven Design",
        "Yoke Design",
        "Self Design",
        "Striped",
        "Checked",
        "Colourblocked",
        "Other / Custom...",
    ],
    "print_types": [
        "Floral",
        "Geometric",
        "Paisley",
        "Ethnic Motifs",
        "Abstract",
        "Solid",
        "Tribal",
        "Chevron",
        "Tie and Dye",
        "Polka Dot",
        "Other / Custom...",
    ],
    "neck_styles": [
        "V-Neck",
        "Round Neck",
        "Mandarin Collar",
        "Sweetheart Neck",
        "Boat Neck",
        "Square Neck",
        "Halter Neck",
        "Shirt Collar",
        "Scoop Neck",
        "Keyhole Neck",
        "Other / Custom...",
    ],
    "sleeve_lengths": [
        "Three-Quarter Sleeves",
        "Short Sleeves",
        "Long Sleeves",
        "Sleeveless",
        "Other / Custom...",
    ],
    "shapes": [
        "Straight",
        "A-Line",
        "Anarkali",
        "Flared",
        "Kaftan",
        "Pathani",
        "Other / Custom...",
    ],
    "weave_types": [
        "Machine Weave",
        "Handloom",
        "Regular",
        "Knitted",
        "Powerloom",
        "Other / Custom...",
    ],
    "wash_cares": [
        "Dry Clean",
        "Hand Wash",
        "Machine Wash",
        "Dry Clean Only",
        "Hand Wash Only",
        "Other / Custom...",
    ],
    "occasions": [
        "Festive",
        "Casual",
        "Daily",
        "Party",
        "Wedding",
        "Fusion",
        "Work",
        "Other / Custom...",
    ],
    "packages": [
        "1 Kurta, 1 Pant",
        "1 Kurta, 1 Pant, 1 Dupatta",
        "1 Kurta",
        "1 Kurta, 1 Salwar",
        "1 Kurta, 1 Palazzo",
        "Other / Custom...",
    ],
}

st.title("👗 Apparel Catalog & Multi-Marketplace Hub")

# ==============================================================================
# 1. CATALOG MANAGEMENT (ADD / EDIT)
# ==============================================================================
col_t1, col_t2, col_t3 = st.columns([2, 1, 1])
with col_t1:
    st.subheader("📦 Master Catalog")
with col_t2:
    if st.button("➕ Add New Product", use_container_width=True, type="secondary"):
        st.session_state["show_add_modal"] = not st.session_state.get("show_add_modal", False)
        st.session_state["edit_product_key"] = None
        st.session_state["temp_colors"] = [{"color_name": "White", "color_map": "White", "images": ""}]
with col_t3:
    if st.button("🗑️ Delete Selected", use_container_width=True):
        st.session_state["trigger_delete"] = True

# Add / Edit Form Drawer
if st.session_state.get("show_add_modal", False) or st.session_state.get("edit_product_key", None):
    edit_key = st.session_state.get("edit_product_key", None)
    is_edit = edit_key is not None and edit_key in db
    curr_data = db[edit_key] if is_edit else {}

    st.info(f"✏️ **{'Editing Style: ' + edit_key if is_edit else 'Add New Design to Catalog'}**")

    # Maintain color variants in session state
    if "temp_colors" not in st.session_state or is_edit and st.session_state.get("loaded_edit_key") != edit_key:
        if is_edit:
            st.session_state["temp_colors"] = [
                {
                    "color_name": c.get("color_name", "White"),
                    "color_map": c.get("color_map", "White"),
                    "images": "\n".join(c.get("images", [])),
                }
                for c in curr_data.get("color_variants", [])
            ]
            st.session_state["loaded_edit_key"] = edit_key
        else:
            st.session_state["temp_colors"] = [{"color_name": "White", "color_map": "White", "images": ""}]
            st.session_state["loaded_edit_key"] = None

    with st.form("product_form"):
        f1, f2, f3 = st.columns(3)
        with f1:
            st.markdown("### 🏷️ Base Identifiers")
            d_code = st.text_input("Design Code / Model Name*", value=curr_data.get("design_code", "NewStyle"))
            t_core = st.text_input("Core Title (Without Brand)*", value=curr_data.get("title_core", "Floral Cotton Embroidered Kurta Set with Pants"))

            fab_def = curr_data.get("fabric", "Pure Cotton")
            fab_idx = DROPDOWNS["fabrics"].index(fab_def) if fab_def in DROPDOWNS["fabrics"] else len(DROPDOWNS["fabrics"]) - 1
            fab_choice = st.selectbox("Fabric*", DROPDOWNS["fabrics"], index=fab_idx)
            fab = st.text_input("Type Custom Fabric", value=fab_def) if fab_choice == "Other / Custom..." else fab_choice

            pat_def = curr_data.get("top_pattern", "Embroidered")
            pat_idx = DROPDOWNS["top_patterns"].index(pat_def) if pat_def in DROPDOWNS["top_patterns"] else len(DROPDOWNS["top_patterns"]) - 1
            pat_choice = st.selectbox("Pattern / Work*", DROPDOWNS["top_patterns"], index=pat_idx)
            pat = st.text_input("Type Custom Pattern", value=pat_def) if pat_choice == "Other / Custom..." else pat_choice

            prt_def = curr_data.get("print_type", "Floral")
            prt_idx = DROPDOWNS["print_types"].index(prt_def) if prt_def in DROPDOWNS["print_types"] else len(DROPDOWNS["print_types"]) - 1
            prt_choice = st.selectbox("Print Type*", DROPDOWNS["print_types"], index=prt_idx)
            prt = st.text_input("Type Custom Print", value=prt_def) if prt_choice == "Other / Custom..." else prt_choice

        with f2:
            st.markdown("### ✂️ Cut & Styling")
            nck_def = curr_data.get("neck", "V-Neck")
            nck_idx = DROPDOWNS["neck_styles"].index(nck_def) if nck_def in DROPDOWNS["neck_styles"] else len(DROPDOWNS["neck_styles"]) - 1
            nck_choice = st.selectbox("Neck Type*", DROPDOWNS["neck_styles"], index=nck_idx)
            nck = st.text_input("Type Custom Neck", value=nck_def) if nck_choice == "Other / Custom..." else nck_choice

            slv_def = curr_data.get("sleeve_length", "Three-Quarter Sleeves")
            slv_idx = DROPDOWNS["sleeve_lengths"].index(slv_def) if slv_def in DROPDOWNS["sleeve_lengths"] else len(DROPDOWNS["sleeve_lengths"]) - 1
            slv_choice = st.selectbox("Sleeve Length*", DROPDOWNS["sleeve_lengths"], index=slv_idx)
            slv = st.text_input("Type Custom Sleeve", value=slv_def) if slv_choice == "Other / Custom..." else slv_choice

            shp_def = curr_data.get("shape", "Straight")
            shp_idx = DROPDOWNS["shapes"].index(shp_def) if shp_def in DROPDOWNS["shapes"] else len(DROPDOWNS["shapes"]) - 1
            shp_choice = st.selectbox("Kurta Shape / Fit*", DROPDOWNS["shapes"], index=shp_idx)
            shp = st.text_input("Type Custom Shape", value=shp_def) if shp_choice == "Other / Custom..." else shp_choice

            wev_def = curr_data.get("weave", "Machine Weave")
            wev_idx = DROPDOWNS["weave_types"].index(wev_def) if wev_def in DROPDOWNS["weave_types"] else len(DROPDOWNS["weave_types"]) - 1
            wev_choice = st.selectbox("Weave Type*", DROPDOWNS["weave_types"], index=wev_idx)
            wev = st.text_input("Type Custom Weave", value=wev_def) if wev_choice == "Other / Custom..." else wev_choice

            wsh_def = curr_data.get("wash_care", "Dry Clean")
            wsh_idx = DROPDOWNS["wash_cares"].index(wsh_def) if wsh_def in DROPDOWNS["wash_cares"] else len(DROPDOWNS["wash_cares"]) - 1
            wsh_choice = st.selectbox("Wash Care*", DROPDOWNS["wash_cares"], index=wsh_idx)
            wsh = st.text_input("Type Custom Wash Care", value=wsh_def) if wsh_choice == "Other / Custom..." else wsh_choice

            occ_def = curr_data.get("occasion", "Festive")
            occ_idx = DROPDOWNS["occasions"].index(occ_def) if occ_def in DROPDOWNS["occasions"] else len(DROPDOWNS["occasions"]) - 1
            occ_choice = st.selectbox("Occasion*", DROPDOWNS["occasions"], index=occ_idx)
            occ = st.text_input("Type Custom Occasion", value=occ_def) if occ_choice == "Other / Custom..." else occ_choice

        with f3:
            st.markdown("### 💰 Pricing & Standard Curve")
            mrp_val = st.number_input("MRP (₹)*", value=int(curr_data.get("mrp", 3999)), step=100)
            sp_val = st.number_input("Selling Price (₹)*", value=int(curr_data.get("selling_price", 1499)), step=50)
            hsn_val = st.text_input("HSN Code*", value=str(curr_data.get("hsn", "62114210")))

            pkg_def = curr_data.get("package_contains", "1 Kurta, 1 Pant")
            pkg_idx = DROPDOWNS["packages"].index(pkg_def) if pkg_def in DROPDOWNS["packages"] else len(DROPDOWNS["packages"]) - 1
            pkg_choice = st.selectbox("Package Contains*", DROPDOWNS["packages"], index=pkg_idx)
            pkg = st.text_input("Type Custom Package", value=pkg_def) if pkg_choice == "Other / Custom..." else pkg_choice

            sz_list = st.multiselect(
                "Available Sizes*",
                ["XS", "S", "M", "L", "XL", "2XL", "3XL"],
                default=curr_data.get("sizes", ["S", "M", "L", "XL", "2XL"]),
            )

        st.markdown("---")
        st.markdown("### 🎨 Color Variants & Image Sets")
        st.caption("Add every color this design is manufactured in. Each color gets its own images and color code.")

        updated_color_variants = []
        for idx, cv in enumerate(st.session_state["temp_colors"]):
            st.markdown(f"**Colorway #{idx + 1}**")
            c_col1, c_col2, c_col3 = st.columns([1.5, 1.5, 3])
            with c_col1:
                col_name = st.text_input(f"Color Name #{idx + 1}", value=cv["color_name"], key=f"cname_{idx}")
            with c_col2:
                col_map_choice = st.selectbox(
                    f"Standard Color Map #{idx + 1}",
                    DROPDOWNS["colors"][:-1],
                    index=DROPDOWNS["colors"][:-1].index(cv["color_map"]) if cv["color_map"] in DROPDOWNS["colors"][:-1] else 0,
                    key=f"cmap_{idx}",
                )
            with c_col3:
                img_text = st.text_area(f"Images for {col_name} (1 URL per line)", value=cv["images"], key=f"cimg_{idx}", height=70)

            updated_color_variants.append(
                {
                    "color_name": col_name,
                    "color_map": col_map_choice,
                    "images": [u.strip() for u in img_text.splitlines() if u.strip()],
                }
            )

        st.markdown("---")
        desc = st.text_area(
            "Product Description",
            value=curr_data.get(
                "description",
                "Crafted from pure cotton, this set offers incredible breathability and an exceptionally soft touch against your skin so you stay completely comfortable all day.",
            ),
        )

        save_c1, save_c2 = st.columns([1, 4])
        with save_c1:
            save_btn = st.form_submit_button("💾 Save Product", type="primary", use_container_width=True)
        with save_c2:
            cancel_btn = st.form_submit_button("Cancel", use_container_width=False)

        if cancel_btn:
            st.session_state["show_add_modal"] = False
            st.session_state["edit_product_key"] = None
            st.rerun()

        if save_btn:
            std_measurements = {
                "XS": {"Across Shoulder": 13.5, "Bust": 34.0, "Chest": 34.0, "Front Length": 29.0, "Hips": 32.0, "Waist": 28.0, "Inseam Length": 25.0},
                "S": {"Across Shoulder": 14.0, "Bust": 36.0, "Chest": 36.0, "Front Length": 29.0, "Hips": 34.0, "Waist": 30.0, "Inseam Length": 25.0},
                "M": {"Across Shoulder": 14.5, "Bust": 38.0, "Chest": 38.0, "Front Length": 29.0, "Hips": 36.0, "Waist": 32.0, "Inseam Length": 24.8},
                "L": {"Across Shoulder": 15.0, "Bust": 40.0, "Chest": 40.0, "Front Length": 29.0, "Hips": 38.0, "Waist": 34.0, "Inseam Length": 24.3},
                "XL": {"Across Shoulder": 15.5, "Bust": 42.0, "Chest": 42.0, "Front Length": 29.0, "Hips": 40.0, "Waist": 36.0, "Inseam Length": 23.8},
                "2XL": {"Across Shoulder": 16.0, "Bust": 44.0, "Chest": 44.0, "Front Length": 29.0, "Hips": 42.0, "Waist": 38.0, "Inseam Length": 22.0},
                "3XL": {"Across Shoulder": 16.5, "Bust": 46.0, "Chest": 46.0, "Front Length": 29.0, "Hips": 44.0, "Waist": 40.0, "Inseam Length": 22.0},
            }

            db[d_code] = {
                "design_code": d_code,
                "title_core": t_core,
                "fabric": fab,
                "top_pattern": pat,
                "print_type": prt,
                "neck": nck,
                "sleeve_length": slv,
                "shape": shp,
                "weave": wev,
                "wash_care": wsh,
                "occasion": occ,
                "mrp": mrp_val,
                "selling_price": sp_val,
                "hsn": hsn_val,
                "sizes": sz_list,
                "package_contains": pkg,
                "net_qty": 1,
                "description": desc,
                "measurements": {sz: std_measurements.get(sz, {}) for sz in sz_list},
                "color_variants": updated_color_variants,
            }
            save_db(db)
            st.session_state["show_add_modal"] = False
            st.session_state["edit_product_key"] = None
            st.success(f"✅ Product '{d_code}' with {len(updated_color_variants)} color variant(s) saved successfully!")
            st.rerun()

    # Button to add another color block
    if st.button("➕ Add Another Color Variant to this Style"):
        st.session_state["temp_colors"].append({"color_name": "", "color_map": "White", "images": ""})
        st.rerun()

# ==============================================================================
# 2. CATALOG TABLE
# ==============================================================================
if not db:
    st.info("No styles found in catalog. Click '➕ Add New Product' above to create one.")
else:
    table_data = []
    for key, item in db.items():
        colors_list = [c.get("color_name") for c in item.get("color_variants", [])]
        table_data.append({
            "Select": False,
            "Design Code": item.get("design_code"),
            "Title": item.get("title_core"),
            "Colors": ", ".join(colors_list) if colors_list else item.get("color", ""),
            "Fabric": item.get("fabric"),
            "Pattern": item.get("top_pattern"),
            "Neck": item.get("neck"),
            "Sleeve": item.get("sleeve_length"),
            "Shape": item.get("shape"),
            "MRP (₹)": item.get("mrp"),
            "Price (₹)": item.get("selling_price"),
            "Sizes": ", ".join(item.get("sizes", [])),
        })

    df_catalog = pd.DataFrame(table_data)

    edited_df = st.data_editor(
        df_catalog,
        column_config={
            "Select": st.column_config.CheckboxColumn("Select", default=False),
            "Design Code": st.column_config.TextColumn("Design Code", disabled=True),
            "Title": st.column_config.TextColumn("Title", disabled=True),
        },
        disabled=[c for c in df_catalog.columns if c != "Select"],
        hide_index=True,
        use_container_width=True,
    )

    selected_designs = edited_df[edited_df["Select"] == True]["Design Code"].tolist()

    if st.session_state.get("trigger_delete", False):
        if not selected_designs:
            st.warning("Please tick the checkbox next to the product(s) you want to delete.")
        else:
            for d in selected_designs:
                if d in db:
                    del db[d]
            save_db(db)
            st.session_state["trigger_delete"] = False
            st.success(f"Deleted {len(selected_designs)} product(s) from catalog.")
            st.rerun()

    if len(selected_designs) == 1:
        if st.button(f"✏️ Edit Selected ({selected_designs[0]})"):
            st.session_state["edit_product_key"] = selected_designs[0]
            st.session_state["show_add_modal"] = False
            st.rerun()

    st.markdown("---")

    # ==============================================================================
    # 3. LISTING HUB (AUTOMATED PARENT-CHILD & COLORWAYS GENERATION)
    # ==============================================================================
    st.subheader("🚀 List Selected Products")

    if not selected_designs:
        st.write("👈 *Tick one or more products in the catalog table above to configure and generate listing files.*")
    else:
        st.success(f"Selected **{len(selected_designs)}** product(s): **{', '.join(selected_designs)}**")

        c_plat, c_brand, c_tpl = st.columns([1.5, 2, 2.5])

        with c_plat:
            marketplace = st.selectbox("Select Target Marketplace", ["Myntra", "Amazon.in"])
            default_cat = "Kurta Sets" if marketplace == "Myntra" else "KURTA"
            category_val = st.text_input("Category / Product Type", value=default_cat)

        with c_brand:
            if marketplace == "Myntra":
                selected_brands = st.multiselect(
                    "Select Brands to Generate",
                    ["KALINI", "MITERA", "PERVAS"],
                    default=["KALINI", "MITERA", "PERVAS"],
                )
            else:
                selected_brands = st.multiselect(
                    "Select Brand(s)",
                    ["PERVAS", "BLUE RIBBON"],
                    default=["PERVAS"],
                )

        with c_tpl:
            uploaded_template = st.file_uploader(
                f"Upload Latest {marketplace} Template (.xlsx / .xlsm)",
                type=["xlsx", "xlsm"],
            )

        if st.button(f"⚡ Generate {marketplace} Bulk Upload File for {len(selected_designs)} Product(s)", type="primary", use_container_width=True):
            if not uploaded_template:
                st.error("Please upload the marketplace template file.")
            elif not selected_brands:
                st.error("Please select at least one brand.")
            else:
                is_xlsm = uploaded_template.name.lower().endswith(".xlsm")
                keep_vba_flag = True if is_xlsm else False
                out_ext = ".xlsm" if is_xlsm else ".xlsx"
                out_mime = "application/vnd.ms-excel.sheet.macroEnabled.12" if is_xlsm else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                wb = openpyxl.load_workbook(uploaded_template, keep_vba=keep_vba_flag)

                # Sheet Detection
                if marketplace == "Myntra":
                    sheet_name = "Kurta Sets" if "set" in category_val.lower() else "Kurtas"
                    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
                    header_row = 3
                    start_row = 4
                else:
                    ws = wb["Template"] if "Template" in wb.sheetnames else wb.active
                    header_row = 4
                    start_row = 7  # Amazon dataRow=7

                    # Clear row 6 dummy example data so it doesn't pollute the file
                    for c in range(1, ws.max_column + 1):
                        ws.cell(row=6, column=c).value = None

                # Build column mapping from headers
                col_map = {}
                header_names = []
                for c in range(1, ws.max_column + 1):
                    h_val = ws.cell(row=header_row, column=c).value
                    if h_val:
                        col_name_clean = str(h_val).strip()
                        if col_name_clean not in col_map:
                            col_map[col_name_clean] = []
                            header_names.append(col_name_clean)
                        col_map[col_name_clean].append(c)

                def write_cell(row_idx, col_name, value, occurrence=0):
                    if value is None:
                        return

                    target_col_name = None
                    if col_name in col_map:
                        target_col_name = col_name
                    elif header_names:
                        match, score, _ = process.extractOne(col_name, header_names, scorer=fuzz.token_sort_ratio)
                        if score >= 88:
                            target_col_name = match

                    if target_col_name and target_col_name in col_map:
                        cols = col_map[target_col_name]
                        if occurrence < len(cols):
                            ws.cell(row=row_idx, column=cols[occurrence], value=value)

                current_row = start_row
                group_id_counter = 1

                for d_name in selected_designs:
                    prod = db[d_name]
                    colorways = prod.get("color_variants", [{"color_name": "White", "color_map": "White", "images": []}])

                    for brand in selected_brands:
                        # ------------------------------------------------------
                        # AMAZON: 1 SINGLE PARENT ROW FOR THE WHOLE DESIGN
                        # ------------------------------------------------------
                        if marketplace == "Amazon.in":
                            parent_sku = f"{brand}-{prod.get('design_code')}-Parent"
                            parent_title = f"{brand} Women's {prod.get('fabric', 'Cotton')} {prod.get('top_pattern', 'Printed')} Kurta Pant Set ({prod.get('design_code')})"

                            p_row = {
                                "Status": "Active",
                                "Title": parent_title,
                                "SKU": parent_sku,
                                "Product Type": category_val.upper(),
                                "Listing Action": "Create or Replace (Full Update)",
                                "Parentage Level": "Parent",
                                "Variation Theme Name": "SIZE/COLOR",
                                "Item Name": parent_title,
                                "Brand Name": brand,
                                "Product Id Type": "GTIN Exempt",
                                "Model Name": prod.get("design_code"),
                                "Manufacturer": "Pervas, Surat, Gujarat - 395010",
                                "Product Description": prod.get("description"),
                                "Generic Keywords": "two piece suit set coord dress stylish co ord sets cord coords dresses a line neck ords long salwar trendy new printed latest design traditional suits ladies daily rayon festive kutis indian kutties v kurthi fashion mul weddings floral print",
                                "Lifestyle": "Casual",
                                "Department Name": "Womens",
                                "Target Gender": "Female",
                                "Age Range Description": "Adult",
                                "Fabric Type": prod.get("fabric", "Cotton"),
                                "Material": prod.get("fabric", "Cotton"),
                                "Item Type Name": category_val.upper(),
                                "Item Length Description": "Knee Length",
                                "Occasion": prod.get("occasion", "Festive"),
                                "Care Instructions": prod.get("wash_care", "Dry Clean Only"),
                                "Manufacturer Contact Information": "Pervas, 2087, Second Floor, The Palladium Mall, Near Apple Square, Yogi Chowk, Varaccha, Surat, Gujarat, India - 395010",
                                "Packer Contact Information": "Pervas, 2087, Second Floor, The Palladium Mall, Near Apple Square, Yogi Chowk, Varaccha, Surat, Gujarat, India - 395010",
                                "Design Name": prod.get("top_pattern", "Embroidered"),
                                "Neck Style": prod.get("neck", "V-Neck"),
                                "Sleeve Length Description": "3/4 Sleeve" if "Three" in prod.get("sleeve_length", "") else prod.get("sleeve_length", "3/4 Sleeve"),
                                "Weave Method": "Powerloom",
                                "Item Weight": 450.0,
                                "Item Weight Unit": "Grams",
                                "Country of Origin": "India",
                                "Skip Offer": "No",
                                "Item Condition": "New",
                                "Offer Condition Note": "New",
                            }

                            for col_name, val in p_row.items():
                                write_cell(current_row, col_name, val)

                            # Parent uses images of the first colorway
                            first_c_imgs = colorways[0].get("images", [])
                            if first_c_imgs:
                                write_cell(current_row, "Main Image URL", first_c_imgs[0])
                                for i_idx, img_url in enumerate(first_c_imgs[1:8]):
                                    write_cell(current_row, "Other Image URL", img_url, occurrence=i_idx)

                            write_cell(current_row, "Bullet Point", "1. Elegant Design: Features sophisticated embroidery and styling for everyday and festive elegance.", occurrence=0)
                            write_cell(current_row, "Bullet Point", "2. Soft & Breathable Fabric: Crafted from lightweight fabric offering all-day comfort.", occurrence=1)
                            write_cell(current_row, "Bullet Point", "3. Versatile Occasion: Ideal for casual outings, office wear, family gatherings, and celebrations.", occurrence=2)
                            write_cell(current_row, "Bullet Point", "4. Easy Care & Lasting Style: Durable stitching and long-lasting fabric quality.", occurrence=3)

                            current_row += 1

                        # ------------------------------------------------------
                        # CHILD ROWS: LOOP THROUGH EVERY COLORWAY & SIZE
                        # ------------------------------------------------------
                        for cway in colorways:
                            c_name = cway.get("color_name", "White")
                            c_map = cway.get("color_map", c_name)
                            c_imgs = cway.get("images", [])

                            for sz in prod.get("sizes", []):
                                amz_size = "2XL" if sz.upper() in ["2XL", "XXL"] else sz
                                myntra_size = "XXL" if sz.upper() in ["2XL", "XXL"] else sz
                                m = prod.get("measurements", {}).get(sz, {})

                                if marketplace == "Myntra":
                                    sku = f"{brand}-P-{prod.get('design_code')}-{c_name}-{myntra_size}"
                                    art_num = f"{brand}-P-{prod.get('design_code')}-{c_name}"
                                    display_name = f"{brand} {c_name} {prod.get('title_core', '')}"

                                    m_row = {
                                        "styleGroupId": group_id_counter,  # Unique group ID per colorway
                                        "vendorSkuCode": sku,
                                        "vendorArticleNumber": art_num,
                                        "vendorArticleName": display_name,
                                        "brand": brand,
                                        "Manufacturer Name and Address with Pincode": "Pervas, 2087, The Palladium Mall, Yogi Chowk, Surat, Gujarat - 395010",
                                        "Packer Name and Address with Pincode": "Pervas, 2087, The Palladium Mall, Yogi Chowk, Surat, Gujarat - 395010",
                                        "Country Of Origin": "India",
                                        "articleType": category_val,
                                        "Brand Size": myntra_size,
                                        "Standard Size": myntra_size,
                                        "is Standard Size present on Label": "Yes",
                                        "Brand Colour (Remarks)": c_name,
                                        "HSN": prod.get("hsn"),
                                        "SKUCode": sku,
                                        "MRP": prod.get("mrp"),
                                        "ISP": prod.get("selling_price"),
                                        "AgeGroup": "Adults-Women",
                                        "Prominent Colour": c_map,
                                        "FashionType": "Fashion",
                                        "Usage": "Casual",
                                        "Product Details": prod.get("description"),
                                        "productDisplayName": display_name,
                                        "Top Fabric": prod.get("fabric"),
                                        "Top Pattern": prod.get("top_pattern"),
                                        "Neck": prod.get("neck"),
                                        "Sleeve Length": prod.get("sleeve_length"),
                                        "Top Shape": prod.get("shape"),
                                        "Bottom Fabric": prod.get("fabric"),
                                        "Bottom Pattern": prod.get("top_pattern"),
                                        "Print or Pattern Type": prod.get("print_type"),
                                        "Occasion": prod.get("occasion"),
                                        "Weave Pattern": "Regular",
                                        "Weave Type": prod.get("weave"),
                                        "Wash Care": prod.get("wash_care"),
                                        "Stitch": "Ready to Wear",
                                        "Package Contains": prod.get("package_contains"),
                                        "Net Quantity": prod.get("net_qty"),
                                        "Across Shoulder ( Inches )": m.get("Across Shoulder"),
                                        "Bust ( Inches )": m.get("Bust"),
                                        "Chest ( Inches )": m.get("Chest"),
                                        "Front Length ( Inches )": m.get("Front Length"),
                                        "Hips ( Inches )": m.get("Hips"),
                                        "Waist ( Inches )": m.get("Waist"),
                                        "Inseam Length ( Inches )": m.get("Inseam Length"),
                                    }
                                    for col_name, val in m_row.items():
                                        write_cell(current_row, col_name, val)

                                else:
                                    # AMAZON CHILD ROW
                                    child_sku = f"{brand}-{prod.get('design_code')}-{c_name}-{amz_size}"
                                    parent_sku = f"{brand}-{prod.get('design_code')}-Parent"
                                    child_title = f"{brand} Women's {prod.get('fabric', 'Cotton')} Kurta Pant Set ({prod.get('design_code')} {c_name} {amz_size})"

                                    c_row = {
                                        "Status": "Active",
                                        "Title": child_title,
                                        "SKU": child_sku,
                                        "Product Type": category_val.upper(),
                                        "Listing Action": "Create or Replace (Full Update)",
                                        "Parentage Level": "Child",
                                        "Parent SKU": parent_sku,
                                        "Variation Theme Name": "SIZE/COLOR",
                                        "Item Name": child_title,
                                        "Brand Name": brand,
                                        "Product Id Type": "GTIN Exempt",
                                        "Model Name": prod.get("design_code"),
                                        "Part Number": prod.get("design_code"),
                                        "Apparel Size System": "IN",
                                        "Apparel Size Class": "Alpha",
                                        "Apparel Size Value": amz_size,
                                        "Shirt Size System": "IN",
                                        "Shirt Size Class": "Alpha",
                                        "Shirt Size Value": amz_size,
                                        "Shirt Body Type": "Regular",
                                        "Special Size": "Standard",
                                        "Color": c_name,
                                        "Color Map": c_map,
                                        "Standard Price": prod.get("selling_price"),
                                        "Maximum Retail Price": prod.get("mrp"),
                                        "Your Price INR (Sell on Amazon, IN)": prod.get("selling_price"),
                                        "Maximum Retail Price (Sell on Amazon, IN)": prod.get("mrp"),
                                        "Product Description": prod.get("description"),
                                        "Generic Keywords": "two piece suit set coord dress stylish co ord sets cord coords dresses a line neck ords long salwar trendy new printed latest design traditional suits ladies daily rayon festive kutis indian kutties v kurthi fashion mul weddings floral print",
                                        "Lifestyle": "Casual",
                                        "Style": prod.get("shape", "Straight"),
                                        "Fit Type": "Regular",
                                        "Department Name": "Womens",
                                        "Target Gender": "Female",
                                        "Age Range Description": "Adult",
                                        "Fabric Type": prod.get("fabric", "Cotton"),
                                        "Material": prod.get("fabric", "Cotton"),
                                        "Pattern": prod.get("print_type", "Floral"),
                                        "Item Type Name": category_val.upper(),
                                        "Item Length Description": "Knee Length",
                                        "Occasion": prod.get("occasion", "Festive"),
                                        "Care Instructions": prod.get("wash_care", "Dry Clean Only"),
                                        "Manufacturer": "Pervas, Surat, Gujarat - 395010",
                                        "Manufacturer Contact Information": "Pervas, 2087, Second Floor, The Palladium Mall, Near Apple Square, Yogi Chowk, Varaccha, Surat, Gujarat, India - 395010",
                                        "Packer Contact Information": "Pervas, 2087, Second Floor, The Palladium Mall, Near Apple Square, Yogi Chowk, Varaccha, Surat, Gujarat, India - 395010",
                                        "Design Name": prod.get("top_pattern", "Embroidered"),
                                        "External Product Information Entity": "HSN Code",
                                        "External Product Information": prod.get("hsn", "62114210"),
                                        "Neck Style": prod.get("neck", "V-Neck"),
                                        "Sleeve Length Description": "3/4 Sleeve" if "Three" in prod.get("sleeve_length", "") else prod.get("sleeve_length", "3/4 Sleeve"),
                                        "Weave Method": "Powerloom",
                                        "Item Weight": 450.0,
                                        "Item Weight Unit": "Grams",
                                        "Country of Origin": "India",
                                        "Skip Offer": "No",
                                        "Item Condition": "New",
                                        "Offer Condition Note": "New",
                                        "Fulfillment Channel Code (IN)": "AMAZON_IN",
                                        "Item Package Length": 25.0,
                                        "Package Length Unit": "Centimeters",
                                        "Item Package Width": 22.0,
                                        "Package Width Unit": "Centimeters",
                                        "Item Package Height": 3.0,
                                        "Package Height Unit": "Centimeters",
                                        "Package Weight": 450.0,
                                        "Package Weight Unit": "Grams",
                                    }

                                    for col_name, val in c_row.items():
                                        write_cell(current_row, col_name, val)

                                    if c_imgs:
                                        write_cell(current_row, "Main Image URL", c_imgs[0])
                                        for i_idx, img_url in enumerate(c_imgs[1:8]):
                                            write_cell(current_row, "Other Image URL", img_url, occurrence=i_idx)

                                    write_cell(current_row, "Bullet Point", "1. Elegant Design: Features sophisticated embroidery and styling for everyday and festive elegance.", occurrence=0)
                                    write_cell(current_row, "Bullet Point", "2. Soft & Breathable Fabric: Crafted from lightweight fabric offering all-day comfort.", occurrence=1)
                                    write_cell(current_row, "Bullet Point", "3. Versatile Occasion: Ideal for casual outings, office wear, family gatherings, and celebrations.", occurrence=2)
                                    write_cell(current_row, "Bullet Point", "4. Easy Care & Lasting Style: Durable stitching and long-lasting fabric quality.", occurrence=3)

                                current_row += 1
                            # Increment styleGroupId for each colorway on Myntra
                            group_id_counter += 1

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                total_rows = current_row - start_row
                st.session_state["generated_file"] = {
                    "data": output.getvalue(),
                    "filename": f"{marketplace}_bulk_listing_{len(selected_designs)}_styles{out_ext}",
                    "mime": out_mime,
                    "rows": total_rows,
                    "designs_count": len(selected_designs),
                    "brands_count": len(selected_brands),
                    "marketplace": marketplace,
                }

        if "generated_file" in st.session_state:
            gen = st.session_state["generated_file"]
            st.success(f"🎉 Generated {gen['rows']} total rows across {gen['designs_count']} style(s) and {gen['brands_count']} brand(s)!")
            st.download_button(
                label=f"📥 Download Bulk Upload Sheet ({gen['marketplace']})",
                data=gen["data"],
                file_name=gen["filename"],
                mime=gen["mime"],
                use_container_width=True,
            )
