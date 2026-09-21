import streamlit as st
import openpyxl
from rapidfuzz import process, fuzz
import pandas as pd
import json
import os
import io

DB_FILE = "catalog_db.json"

st.set_page_config(page_title="Apparel Master Lister", layout="wide")

def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            try:
                return json.load(f)
            except Exception:
                return {}
    return {}

def save_db(data):
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=2)

db = load_db()

DROPDOWNS = {
    "colors": ["White", "Black", "Multi", "Off White", "Navy Blue", "Mustard", "Pink", "Red", "Green", "Yellow", "Maroon", "Beige", "Purple", "Grey", "Teal", "Coral", "Other / Custom..."],
    "fabrics": ["Pure Cotton", "Cotton Blend", "Cotton Silk", "Silk Blend", "Georgette", "Chanderi", "Rayon", "Organza", "Tissue", "Modal", "Satin", "Linen Blend", "Other / Custom..."],
    "top_patterns": ["Embroidered", "Printed", "Solid", "Woven Design", "Yoke Design", "Self Design", "Striped", "Checked", "Colourblocked", "Other / Custom..."],
    "print_types": ["Floral", "Geometric", "Paisley", "Ethnic Motifs", "Abstract", "Solid", "Tribal", "Chevron", "Tie and Dye", "Polka Dot", "Other / Custom..."],
    "neck_styles": ["V-Neck", "Round Neck", "Mandarin Collar", "Sweetheart Neck", "Boat Neck", "Square Neck", "Halter Neck", "Shirt Collar", "Scoop Neck", "Keyhole Neck", "Other / Custom..."],
    "sleeve_lengths": ["Three-Quarter Sleeves", "Short Sleeves", "Long Sleeves", "Sleeveless", "Other / Custom..."],
    "shapes": ["Straight", "A-Line", "Anarkali", "Flared", "Kaftan", "Pathani", "Other / Custom..."],
    "weave_types": ["Machine Weave", "Handloom", "Regular", "Knitted", "Powerloom", "Other / Custom..."],
    "wash_cares": ["Dry Clean", "Hand Wash", "Machine Wash", "Dry Clean Only", "Hand Wash Only", "Other / Custom..."],
    "occasions": ["Festive", "Casual", "Daily", "Party", "Wedding", "Fusion", "Work", "Other / Custom..."],
    "packages": ["1 Kurta, 1 Pant", "1 Kurta, 1 Pant, 1 Dupatta", "1 Kurta", "1 Kurta, 1 Salwar", "1 Kurta, 1 Palazzo", "Other / Custom..."]
}

st.title("👗 Multi-Marketplace Catalog & Bulk Listing Engine")

# Three distinct tabs
tab_export, tab_view, tab_catalog = st.tabs([
    "🚀 Export Bulk Files (Reuse Saved Products)",
    "📋 View Catalog (Inventory Overview)",
    "➕ Master Catalog (Add / Edit Products)"
])

# ==============================================================================
# TAB 1: REUSE & EXPORT BULK FILES
# ==============================================================================
with tab_export:
    st.subheader("Generate Marketplace Upload Sheets from Master Catalog")
    
    if not db:
        st.info("No products found in the catalog. Please add a product in the 'Master Catalog' tab.")
    else:
        c1, c2, c3 = st.columns([1.5, 1.5, 2])
        
        with c1:
            selected_product_key = st.selectbox("Select Saved Product to List", list(db.keys()))
            prod = db[selected_product_key]
            marketplace = st.selectbox("Select Target Marketplace", ["Myntra", "Amazon.in"])
            
        with c2:
            if marketplace == "Myntra":
                default_category = st.text_input("Default Category / Article Type", value="Kurta Sets")
                selected_brands = st.multiselect(
                    "Select Brands (Generates Rows for All)",
                    ["KALINI", "MITERA", "PERVAS"],
                    default=["KALINI", "MITERA", "PERVAS"]
                )
            else:
                default_category = st.text_input("Default Category / Product Type", value="KURTA")
                selected_brands = st.multiselect(
                    "Select Brand(s)",
                    ["PERVAS", "BLUE RIBBON"],
                    default=["PERVAS"]
                )
        
        with c3:
            uploaded_template = st.file_uploader(
                f"Upload Latest {marketplace} Blank Template (.xlsx / .xlsm)",
                type=["xlsx", "xlsm"]
            )

        st.markdown("---")
        st.write(
            f"**Selected Style:** `{prod.get('design_code')}` — {prod.get('title_core')} | "
            f"**Fabric:** {prod.get('fabric')} | **Neck:** {prod.get('neck')} | "
            f"**Sizes:** {', '.join(prod.get('sizes', []))}"
        )

        def fuzzy_match(input_val, valid_list):
            if not valid_list or not input_val:
                return input_val
            match, score, _ = process.extractOne(str(input_val), valid_list, scorer=fuzz.token_sort_ratio)
            return match if score >= 65 else input_val

        if st.button("⚡ Generate Marketplace Bulk Upload File", type="primary"):
            if not uploaded_template:
                st.error("Please upload the latest marketplace template file first.")
            elif not selected_brands:
                st.error("Please select at least one brand.")
            else:
                wb = openpyxl.load_workbook(uploaded_template, keep_vba=True)
                
                if marketplace == "Myntra":
                    sheet_name = "Kurta Sets" if "set" in default_category.lower() else "Kurtas"
                    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
                    header_row = 3
                    start_row = 4
                else:
                    ws = wb["Template"] if "Template" in wb.sheetnames else wb.active
                    header_row = 4
                    start_row = 6

                col_map = {}
                for c in range(1, ws.max_column + 1):
                    h_val = ws.cell(row=header_row, column=c).value
                    if h_val:
                        col_map[str(h_val).strip()] = c

                current_row = start_row
                group_id_counter = 1

                for brand in selected_brands:
                    for sz in prod.get("sizes", []):
                        std_size = "XXL" if sz == "2XL" else sz
                        sku = f"{brand}-P-{prod.get('design_code')}-{sz}"
                        art_num = f"{brand}-P-{prod.get('design_code')}"
                        display_name = f"{brand} {prod.get('color', '')} {prod.get('title_core', '')}"
                        m = prod.get("measurements", {}).get(sz, {})

                        if marketplace == "Myntra":
                            row_dict = {
                                "styleGroupId": group_id_counter,
                                "vendorSkuCode": sku,
                                "vendorArticleNumber": art_num,
                                "vendorArticleName": display_name,
                                "brand": brand,
                                "Manufacturer Name and Address with Pincode": "Pervas, Surat, Gujarat - 395010",
                                "Packer Name and Address with Pincode": "Pervas, Surat, Gujarat - 395010",
                                "Country Of Origin": "India",
                                "articleType": default_category,
                                "Brand Size": std_size,
                                "Standard Size": std_size,
                                "is Standard Size present on Label": "Yes",
                                "Brand Colour (Remarks)": prod.get("color"),
                                "HSN": prod.get("hsn"),
                                "SKUCode": sku,
                                "MRP": prod.get("mrp"),
                                "ISP": prod.get("selling_price"),
                                "AgeGroup": "Adults-Women",
                                "Prominent Colour": prod.get("color"),
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
                                "Inseam Length ( Inches )": m.get("Inseam Length")
                            }
                        else:
                            row_dict = {
                                "SKU": sku,
                                "Product Type": default_category.upper(),
                                "Listing Action": "Create or Replace (Full Update)",
                                "Parentage Level": "Child",
                                "Parent SKU": f"{brand}-{prod.get('design_code')}",
                                "Variation Theme Name": "SIZE/COLOR",
                                "Item Name": display_name,
                                "Brand Name": brand,
                                "Product Id Type": "GTIN Exempt",
                                "Apparel Size Value": sz,
                                "Standard Price": prod.get("selling_price"),
                                "Maximum Retail Price": prod.get("mrp"),
                                "Color": prod.get("color"),
                                "Fabric Type": prod.get("fabric"),
                                "Material": prod.get("fabric"),
                                "Neck Style": prod.get("neck"),
                                "Pattern": prod.get("top_pattern"),
                                "Product Description": prod.get("description"),
                                "Main Image URL": prod.get("images", [""])[0] if len(prod.get("images", [])) > 0 else ""
                            }

                        for col_name, val in row_dict.items():
                            if col_name in col_map and val is not None:
                                ws.cell(row=current_row, column=col_map[col_name], value=val)

                        current_row += 1
                    group_id_counter += 1

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)

                rows_generated = current_row - start_row
                st.success(f"Successfully generated {rows_generated} rows across {len(selected_brands)} brand(s)!")
                st.download_button(
                    label=f"📥 Download Completed {marketplace} File",
                    data=output,
                    file_name=f"{marketplace}_{prod.get('design_code')}_bulk_upload.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==============================================================================
# TAB 2: VIEW CATALOG (OVERVIEW TABLE)
# ==============================================================================
with tab_view:
    st.subheader("Master Catalog Overview")
    if not db:
        st.info("No products found in the catalog.")
    else:
        table_rows = []
        for key, item in db.items():
            table_rows.append({
                "Design Code": item.get("design_code"),
                "Title": item.get("title_core"),
                "Color": item.get("color"),
                "Fabric": item.get("fabric"),
                "Pattern": item.get("top_pattern"),
                "Neck": item.get("neck"),
                "Sleeve": item.get("sleeve_length"),
                "Shape": item.get("shape"),
                "MRP (₹)": item.get("mrp"),
                "Selling Price (₹)": item.get("selling_price"),
                "HSN": item.get("hsn"),
                "Sizes": ", ".join(item.get("sizes", []))
            })
        
        df = pd.DataFrame(table_rows)
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.caption(f"Showing **{len(df)}** style(s) currently stored in `catalog_db.json`.")

# ==============================================================================
# TAB 3: MASTER CATALOG (ADD / EDIT PRODUCT)
# ==============================================================================
with tab_catalog:
    st.subheader("Add or Edit Master Garment Records")
    st.info("Pick from marketplace-approved dropdown values. If your specific option isn't listed, choose 'Other / Custom...' to type it.")

    with st.form("catalog_entry_form"):
        f1, f2, f3 = st.columns(3)
        
        with f1:
            st.markdown("### 🏷️ Basic Identifiers")
            d_code = st.text_input("Design Code / Model Name*", value="Asmita")
            st.caption("💡 *Common:* Asmita, Supriya, Eva, Maya, Riha")
            
            t_core = st.text_input("Core Title (Without Brand)*", value="Floral Cotton Embroidered Kurta Set with Pants")
            st.caption("💡 *Common:* Floral Cotton Embroidered Kurta Set with Pants")

            clr_choice = st.selectbox("Color*", DROPDOWNS["colors"], index=0)
            clr = st.text_input("Type Custom Color") if clr_choice == "Other / Custom..." else clr_choice
            st.caption("💡 *Common:* White, Black, Multi, Off White, Mustard")

            fab_choice = st.selectbox("Fabric*", DROPDOWNS["fabrics"], index=0)
            fab = st.text_input("Type Custom Fabric") if fab_choice == "Other / Custom..." else fab_choice
            st.caption("💡 *Common:* Pure Cotton, Cotton Blend, Cotton Silk, Georgette")

            pat_choice = st.selectbox("Pattern / Work*", DROPDOWNS["top_patterns"], index=0)
            pat = st.text_input("Type Custom Pattern") if pat_choice == "Other / Custom..." else pat_choice
            st.caption("💡 *Common:* Embroidered, Printed, Solid, Woven Design")

            prt_choice = st.selectbox("Print or Pattern Type*", DROPDOWNS["print_types"], index=0)
            prt = st.text_input("Type Custom Print") if prt_choice == "Other / Custom..." else prt_choice
            st.caption("💡 *Common:* Floral, Geometric, Paisley, Ethnic Motifs")

        with f2:
            st.markdown("### ✂️ Cut & Styling")
            nck_choice = st.selectbox("Neck Type*", DROPDOWNS["neck_styles"], index=0)
            nck = st.text_input("Type Custom Neck") if nck_choice == "Other / Custom..." else nck_choice
            st.caption("💡 *Common:* V-Neck, Round Neck, Mandarin Collar, Sweetheart Neck")

            slv_choice = st.selectbox("Sleeve Length*", DROPDOWNS["sleeve_lengths"], index=0)
            slv = st.text_input("Type Custom Sleeve") if slv_choice == "Other / Custom..." else slv_choice
            st.caption("💡 *Common:* Three-Quarter Sleeves, Short Sleeves, Sleeveless")

            shp_choice = st.selectbox("Kurta Shape / Fit*", DROPDOWNS["shapes"], index=0)
            shp = st.text_input("Type Custom Shape") if shp_choice == "Other / Custom..." else shp_choice
            st.caption("💡 *Common:* Straight, A-Line, Anarkali, Flared")

            wev_choice = st.selectbox("Weave Type*", DROPDOWNS["weave_types"], index=0)
            wev = st.text_input("Type Custom Weave") if wev_choice == "Other / Custom..." else wev_choice
            st.caption("💡 *Common:* Machine Weave, Regular, Handloom")

            wsh_choice = st.selectbox("Wash Care*", DROPDOWNS["wash_cares"], index=0)
            wsh = st.text_input("Type Custom Wash Care") if wsh_choice == "Other / Custom..." else wsh_choice
            st.caption("💡 *Common:* Dry Clean, Hand Wash, Machine Wash")

            occ_choice = st.selectbox("Occasion*", DROPDOWNS["occasions"], index=0)
            occ = st.text_input("Type Custom Occasion") if occ_choice == "Other / Custom..." else occ_choice
            st.caption("💡 *Common:* Festive, Casual, Daily, Wedding")

        with f3:
            st.markdown("### 💰 Pricing & Commercials")
            mrp_val = st.number_input("MRP (₹)*", value=3999, step=100)
            st.caption("💡 *Common:* 2499, 2999, 3499, 3999, 4999")
            
            sp_val = st.number_input("Selling Price (₹)*", value=1499, step=50)
            st.caption("💡 *Common:* 699, 799, 999, 1299, 1499")

            hsn_val = st.text_input("HSN Code*", value="62114210")
            st.caption("💡 *Common:* 62114210 (Cotton Suits), 6204 (Women Garments)")

            pkg_choice = st.selectbox("Package Contains*", DROPDOWNS["packages"], index=0)
            pkg = st.text_input("Type Custom Package") if pkg_choice == "Other / Custom..." else pkg_choice
            st.caption("💡 *Common:* 1 Kurta, 1 Pant | 1 Kurta, 1 Pant, 1 Dupatta")

            sz_list = st.multiselect(
                "Available Sizes*", 
                ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL"], 
                default=["S", "M", "L", "XL", "2XL"]
            )
            st.caption("💡 *Common standard curve:* S, M, L, XL, 2XL")

        st.markdown("---")
        st.markdown("### 📝 Details & Cloud Images")
        
        desc = st.text_area(
            "Product Description", 
            value="Crafted from pure cotton, this set offers incredible breathability and an exceptionally soft touch against your skin so you stay completely comfortable all day."
        )
        
        img_urls = st.text_area(
            "Image URLs (1 link per line)",
            placeholder="https://m.media-amazon.com/images/I/71kawFhW1mL.jpg\nhttps://m.media-amazon.com/images/I/81cFjq0ePkL.jpg"
        )

        save_btn = st.form_submit_button("💾 Save Product to Master Catalog", type="primary")

        if save_btn:
            std_measurements = {
                "XS": {"Across Shoulder": 13.5, "Bust": 34.0, "Chest": 34.0, "Front Length": 29.0, "Hips": 32.0, "Waist": 28.0, "Inseam Length": 25.0},
                "S": {"Across Shoulder": 14.0, "Bust": 36.0, "Chest": 36.0, "Front Length": 29.0, "Hips": 34.0, "Waist": 30.0, "Inseam Length": 25.0},
                "M": {"Across Shoulder": 14.5, "Bust": 38.0, "Chest": 38.0, "Front Length": 29.0, "Hips": 36.0, "Waist": 32.0, "Inseam Length": 24.8},
                "L": {"Across Shoulder": 15.0, "Bust": 40.0, "Chest": 40.0, "Front Length": 29.0, "Hips": 38.0, "Waist": 34.0, "Inseam Length": 24.3},
                "XL": {"Across Shoulder": 15.5, "Bust": 42.0, "Chest": 42.0, "Front Length": 29.0, "Hips": 40.0, "Waist": 36.0, "Inseam Length": 23.8},
                "2XL": {"Across Shoulder": 16.0, "Bust": 44.0, "Chest": 44.0, "Front Length": 29.0, "Hips": 42.0, "Waist": 38.0, "Inseam Length": 22.0},
                "3XL": {"Across Shoulder": 16.5, "Bust": 46.0, "Chest": 46.0, "Front Length": 29.0, "Hips": 44.0, "Waist": 40.0, "Inseam Length": 22.0}
            }
            
            db[d_code] = {
                "design_code": d_code,
                "title_core": t_core,
                "color": clr,
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
                "images": [u.strip() for u in img_urls.splitlines() if u.strip()]
            }
            save_db(db)
            st.success(f"✅ Product '{d_code}' successfully saved to your master database! You can now view it in 'View Catalog' and use it in 'Export Bulk Files'.")
