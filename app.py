import base64
import datetime
import io
import json
import os
import re
import zipfile
import rarfile
import openpyxl
import pandas as pd
from PIL import Image
from rapidfuzz import fuzz, process
import requests
import streamlit as st
from google import genai
from google.genai import types

DB_FILE = "catalog_db.json"

st.set_page_config(page_title="Apparel Multi-Marketplace Hub V2 (AI-Powered)", layout="wide")


# ==============================================================================
# GITHUB PERSISTENT SYNC (SECRETS-ENABLED)
# ==============================================================================
def get_github_headers():
    token = st.secrets.get("GITHUB_TOKEN", None)
    if token:
        return {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
        }
    return None


def normalize_db_records(data):
    for k, v in data.items():
        if "color_variants" not in v:
            v["color_variants"] = [
                {
                    "color_name": v.get("color", "White"),
                    "color_map": v.get("color", "White"),
                    "images": v.get("images", []),
                }
            ]
        if "top_length" not in v:
            v["top_length"] = "Knee Length"
        if "amazon_overrides" not in v:
            v["amazon_overrides"] = {}
        if "myntra_overrides" not in v:
            v["myntra_overrides"] = {}
    return data


def load_db():
    headers = get_github_headers()
    if headers and "REPO_NAME" in st.secrets:
        repo = st.secrets["REPO_NAME"]
        fpath = st.secrets.get("FILE_PATH", "catalog_db.json")
        branch = st.secrets.get("BRANCH", "main")
        url = f"https://api.github.com/repos/{repo}/contents/{fpath}?ref={branch}"
        try:
            r = requests.get(url, headers=headers, timeout=10)
            if r.status_code == 200:
                content_b64 = r.json().get("content", "")
                data_str = base64.b64decode(content_b64).decode("utf-8")
                return normalize_db_records(json.loads(data_str))
        except Exception:
            pass

    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            try:
                return normalize_db_records(json.load(f))
            except Exception:
                return {}
    return {}


def save_db(data):
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=2)

    headers = get_github_headers()
    if headers and "REPO_NAME" in st.secrets:
        repo = st.secrets["REPO_NAME"]
        fpath = st.secrets.get("FILE_PATH", "catalog_db.json")
        branch = st.secrets.get("BRANCH", "main")
        url = f"https://api.github.com/repos/{repo}/contents/{fpath}"
        try:
            r_get = requests.get(url + f"?ref={branch}", headers=headers, timeout=10)
            sha = r_get.json().get("sha", None) if r_get.status_code == 200 else None

            content_str = json.dumps(data, indent=2)
            content_b64 = base64.b64encode(content_str.encode("utf-8")).decode("utf-8")

            payload = {
                "message": "Auto-sync catalog update from Streamlit Hub V2",
                "content": content_b64,
                "branch": branch,
            }
            if sha:
                payload["sha"] = sha

            r_put = requests.put(url, headers=headers, json=payload, timeout=10)
            if r_put.status_code in [200, 201]:
                st.toast("✅ Saved permanently to GitHub!", icon="🎉")
            else:
                st.error(f"GitHub Sync Error: {r_put.status_code} - {r_put.text}")
        except Exception as e:
            st.error(f"GitHub Sync Connection Failed: {e}")


db = load_db()

# ==============================================================================
# BASE MAPPERS & FALLBACK DROPDOWNS
# ==============================================================================
COLOR_MAPPER = {
    "White": "White", "Off White": "Off White", "Cream": "Off White",
    "Beige": "Beige", "Black": "Black", "Navy Blue": "Blue", "Blue": "Blue",
    "Teal": "Turquoise", "Turquoise Blue": "Turquoise", "Green": "Green",
    "Olive": "Green", "Mehendi Green": "Green", "Sea Green": "Green",
    "Lime Green": "Green", "Mustard": "Yellow", "Yellow": "Yellow",
    "Maroon": "Red", "Red": "Red", "Burgundy": "Red", "Wine": "Red",
    "Rust": "Orange", "Orange": "Orange", "Coral": "Pink", "Peach": "Pink",
    "Pink": "Pink", "Rani Pink": "Pink", "Fuchsia": "Pink", "Magenta": "Pink",
    "Purple": "Purple", "Lavender": "Purple", "Mauve": "Purple",
    "Brown": "Brown", "Coffee Brown": "Brown", "Charcoal": "Grey",
    "Grey": "Grey", "Gold": "Gold", "Silver": "Silver", "Bronze": "Bronze",
    "Copper": "Bronze", "Multi": "Multicolor", "Other / Custom...": "Other / Custom...",
}

FABRIC_MAPPER = {
    "Pure Cotton": "Cotton", "Cotton Blend": "Cotton Blend",
    "Cotton Silk": "Silk Blend", "Silk Blend": "Silk Blend",
    "Georgette": "Georgette", "Chanderi": "Chanderi", "Rayon": "Rayon",
    "Organza": "Organza", "Tissue": "Tissue", "Modal": "Modal",
    "Satin": "Satin", "Linen Blend": "Linen Blend",
}

LENGTH_MAPPER = {
    "Knee Length": ("Knee Length", "Knee Length"),
    "Calf Length": ("Calf Length", "Calf Length"),
    "Above Knee": ("Above Knee", "Mid Thigh Length"),
    "Floor Length": ("Floor Length", "Floor Length"),
    "Ankle Length": ("Calf Length", "Ankle Length"),
    "Short / Hip Length": ("Short", "Hip Length"),
    "Other / Custom...": ("Knee Length", "Knee Length"),
}

AMAZON_SIZE_MAP = {
    "XS": "XS (x_s)", "S": "S (s)", "M": "M (m)", "L": "L (l)",
    "XL": "XL", "2XL": "2XL (2x_l)", "XXL": "2XL (2x_l)",
    "3XL": "3XL", "4XL": "4XL", "5XL": "5XL"
}

AMAZON_EXACT_DESCRIPTION = (
    "Elevate your ethnic with this elegant Women's Kurta , designed to blend comfort with timeless style. "
    "Made from a soft and breathable  fabric, this Kurta featuring delicate white floral embroidery on the sleeves and neckline, "
    "for a refined look. The neck design, 3/4 elasticated sleeves, provide effortless comfort while maintaining a flattering look. "
    "Whether you're dressing for work, casual outings, festive occasions, family gatherings, or travel, this kurta keeps you looking "
    "polished throughout the day. Style it with statement earrings, bangles, and sandals for a graceful ethnic, or keep it minimal for a chic everyday look."
)

DROPDOWNS = {
    "colors": list(COLOR_MAPPER.keys()),
    "amazon_color_maps": [
        "Beige", "Black", "Blue", "Bronze", "Brown", "Clear", "Gold",
        "Green", "Grey", "Metallic", "Multicolor", "Off White", "Orange",
        "Pink", "Purple", "Red", "Silver", "Turquoise", "White", "Yellow",
    ],
    "top_lengths": list(LENGTH_MAPPER.keys()),
    "fabrics": list(FABRIC_MAPPER.keys()) + ["Other / Custom..."],
    "top_patterns": [
        "Embroidered", "Printed", "Solid", "Woven Design", "Yoke Design",
        "Self Design", "Striped", "Checked", "Colourblocked", "Other / Custom...",
    ],
    "print_types": [
        "Floral", "Geometric", "Paisley", "Ethnic Motifs", "Abstract", "Solid",
        "Tribal", "Chevron", "Tie and Dye", "Polka Dot", "Other / Custom...",
    ],
    "neck_styles": [
        "V-Neck", "Round Neck", "Mandarin Collar", "Sweetheart Neck", "Boat Neck",
        "Square Neck", "Halter Neck", "Shirt Collar", "Scoop Neck", "Keyhole Neck",
        "Other / Custom...",
    ],
    "sleeve_lengths": [
        "Three-Quarter Sleeves", "Short Sleeves", "Long Sleeves", "Sleeveless",
        "Other / Custom...",
    ],
    "shapes": [
        "Straight", "A-Line", "Anarkali", "Flared", "Kaftan", "Pathani",
        "Other / Custom...",
    ],
    "weave_types": [
        "Machine Weave", "Handloom", "Regular", "Knitted", "Powerloom",
        "Other / Custom...",
    ],
    "wash_cares": [
        "Dry Clean", "Hand Wash", "Machine Wash", "Dry Clean Only",
        "Hand Wash Only", "Other / Custom...",
    ],
    "occasions": [
        "Festive", "Casual", "Daily", "Party", "Wedding", "Fusion", "Work",
        "Other / Custom...",
    ],
    "packages": [
        "1 Kurta, 1 Pant", "1 Kurta, 1 Pant, 1 Dupatta", "1 Kurta",
        "1 Kurta, 1 Salwar", "1 Kurta, 1 Palazzo", "1 Kurta, 1 Sharara, 1 Dupatta",
        "Other / Custom...",
    ],
}


# ==============================================================================
# V2 ENGINE: DYNAMIC TEMPLATE INTROSPECTION & DROPDOWN EXTRACTOR
# ==============================================================================
def resolve_excel_range(wb, formula_str):
    """Resolves Excel Data Validation range formulas (like =ValidValues!$B$2:$B$50 or Named Ranges)."""
    if not formula_str:
        return []
    formula = str(formula_str).strip()
    if formula.startswith("="):
        formula = formula[1:]

    # Direct comma-separated string list
    if "," in formula and "!" not in formula and ":" not in formula:
        return [o.strip().strip('"') for o in formula.split(",") if o.strip()]

    # Check if it's a Named Range in workbook
    if formula in wb.defined_names:
        defn = wb.defined_names[formula]
        formula = defn.attr_text

    # Parse Sheet!Range (e.g. 'Valid Values'!$A$2:$A$40)
    if "!" in formula:
        try:
            sheet_part, range_part = formula.split("!", 1)
            sheet_name = sheet_part.strip("'")
            if sheet_name in wb.sheetnames:
                ref_ws = wb[sheet_name]
                vals = []
                for row in ref_ws[range_part]:
                    for cell in (row if isinstance(row, tuple) else [row]):
                        if cell.value is not None and str(cell.value).strip() != "":
                            vals.append(str(cell.value).strip())
                return list(dict.fromkeys(vals))
        except Exception:
            pass
    return []


def parse_marketplace_template(uploaded_file, platform="Amazon"):
    """
    Extracts ALL column headers and their dropdown options from an Amazon or Myntra Excel template.
    """
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)

    if platform == "Myntra":
        ws = wb["Kurta Sets"] if "Kurta Sets" in wb.sheetnames else (wb["Kurtas"] if "Kurtas" in wb.sheetnames else wb.active)
        header_row = 3
    else:
        ws = wb["Template"] if "Template" in wb.sheetnames else wb.active
        header_row = 4

    schema = {}
    col_to_header = {}

    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val:
            h_name = str(val).strip()
            # Handle duplicate column names (like Bullet Point or Other Image URL)
            if h_name not in schema:
                schema[h_name] = {"col_idx": col, "options": []}
                col_to_header[col] = h_name

    # 1. Extract from worksheet Data Validations
    if ws.data_validations and ws.data_validations.dataValidation:
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                opts = resolve_excel_range(wb, dv.formula1)
                if opts:
                    for cell_range in dv.sqref.ranges:
                        for c_idx in range(cell_range.min_col, cell_range.max_col + 1):
                            if c_idx in col_to_header:
                                h_name = col_to_header[c_idx]
                                if not schema[h_name]["options"]:
                                    schema[h_name]["options"] = opts

    # 2. Check Valid Values / Drop Down sheets by matching header names
    for sname in wb.sheetnames:
        if any(w in sname.lower() for w in ["valid", "drop", "master", "value"]):
            v_ws = wb[sname]
            for col in range(1, min(v_ws.max_column + 1, 150)):
                v_head = v_ws.cell(row=1, column=col).value or v_ws.cell(row=2, column=col).value
                if v_head:
                    v_head_clean = str(v_head).strip()
                    # Match exact or close header
                    matched_key = None
                    if v_head_clean in schema:
                        matched_key = v_head_clean
                    elif schema:
                        m = process.extractOne(v_head_clean, list(schema.keys()), scorer=fuzz.WRatio)
                        if m and m[1] >= 90:
                            matched_key = m[0]
                    if matched_key and not schema[matched_key]["options"]:
                        opts = []
                        for r in range(2, min(v_ws.max_row + 1, 500)):
                            cval = v_ws.cell(row=r, column=col).value
                            if cval is not None and str(cval).strip() and str(cval).strip() != v_head_clean:
                                opts.append(str(cval).strip())
                        if opts:
                            schema[matched_key]["options"] = list(dict.fromkeys(opts))

    uploaded_file.seek(0)
    return schema


def get_best_dropdown_index(target_val: str, options: list) -> int:
    """
    Finds the exact match OR next-best fuzzy match index in `["", *options]`.
    """
    if not target_val or not options:
        return 0
    target_str = str(target_val).strip()
    full_opts = [""] + options

    if target_str in options:
        return full_opts.index(target_str)

    lower_map = {o.lower(): o for o in options}
    if target_str.lower() in lower_map:
        return full_opts.index(lower_map[target_str.lower()])

    best = process.extractOne(target_str, options, scorer=fuzz.WRatio)
    if best and best[1] >= 52:
        return full_opts.index(best[0])
    return 0


def match_value_to_list(val: str, options: list, default_val: str = "") -> str:
    """Returns the closest matching value from options list."""
    if not val or not options:
        return default_val
    clean_opts = [o for o in options if o != "Other / Custom..."]
    idx = get_best_dropdown_index(val, clean_opts)
    if idx > 0:
        return clean_opts[idx - 1]
    return default_val if default_val in options else (clean_opts[0] if clean_opts else "")


# ==============================================================================
# V2 ENGINE: ZIP/RAR ARCHIVE EXTRACTOR & GEMINI AI MULTIMODAL ANALYZER
# ==============================================================================
def extract_archive(uploaded_archive):
    texts = []
    images = []
    raw_bytes = io.BytesIO(uploaded_archive.read())
    is_rar = uploaded_archive.name.lower().endswith(".rar")
    archive_cls = rarfile.RarFile if is_rar else zipfile.ZipFile

    with archive_cls(raw_bytes) as z:
        for info in z.infolist():
            if info.is_dir():
                continue
            fname = info.filename
            lower_name = fname.lower()
            if "__macosx" in lower_name or os.path.basename(fname).startswith("."):
                continue

            if lower_name.endswith((".txt", ".md", ".csv", ".json")):
                content = z.read(fname).decode("utf-8", errors="ignore")
                texts.append(f"[{os.path.basename(fname)}]:\n{content}")
            elif lower_name.endswith((".jpg", ".jpeg", ".png", ".webp")):
                img_data = z.read(fname)
                img = Image.open(io.BytesIO(img_data)).convert("RGB")
                images.append((os.path.basename(fname), img))

    uploaded_archive.seek(0)
    return "\n\n".join(texts), images


def run_gemini_archive_analysis(api_key, model_name, raw_text, images, amz_schema, myn_schema, archive_name=""):
    client = genai.Client(api_key=api_key)

    # Compress up to 6 images in memory to stay well inside Free Tier token limits
    compressed_imgs = []
    for _, img in images[:6]:
        im_copy = img.copy()
        im_copy.thumbnail((800, 800))
        compressed_imgs.append(im_copy)

    amz_fields_spec = {
        k: (v["options"][:45] if v["options"] else "Free Text")
        for k, v in amz_schema.items()
    }
    myn_fields_spec = {
        k: (v["options"][:45] if v["options"] else "Free Text")
        for k, v in myn_schema.items()
    }

    prompt = f"""
    You are an expert Indian ethnic wear catalog manager for Amazon.in and Myntra.
    Analyze the attached product images and the raw text notes from the archive '{archive_name}'.

    RAW TEXT NOTES FROM ARCHIVE:
    {raw_text if raw_text else "No text file found. Infer all details from the archive name and product images."}

    BASE CATALOG DROPDOWN OPTIONS:
    - fabrics: {DROPDOWNS['fabrics'][:-1]}
    - top_patterns: {DROPDOWNS['top_patterns'][:-1]}
    - print_types: {DROPDOWNS['print_types'][:-1]}
    - neck_styles: {DROPDOWNS['neck_styles'][:-1]}
    - sleeve_lengths: {DROPDOWNS['sleeve_lengths'][:-1]}
    - shapes: {DROPDOWNS['shapes'][:-1]}
    - top_lengths: {DROPDOWNS['top_lengths'][:-1]}
    - weave_types: {DROPDOWNS['weave_types'][:-1]}
    - wash_cares: {DROPDOWNS['wash_cares'][:-1]}
    - occasions: {DROPDOWNS['occasions'][:-1]}
    - packages: {DROPDOWNS['packages'][:-1]}
    - colors: {DROPDOWNS['colors'][:-1]}
    - amazon_color_maps: {DROPDOWNS['amazon_color_maps']}

    AMAZON TEMPLATE FIELDS & ALLOWED DROPDOWNS:
    {json.dumps(amz_fields_spec) if amz_fields_spec else "No custom Amazon template loaded yet."}

    MYNTRA TEMPLATE FIELDS & ALLOWED DROPDOWNS:
    {json.dumps(myn_fields_spec) if myn_fields_spec else "No custom Myntra template loaded yet."}

    INSTRUCTIONS:
    1. Extract 'design_code' from SKU/Design name in text or archive filename.
    2. Identify ALL distinct color variants shown across the uploaded images (e.g., if 5 photos show 5 different colors of the same kurta set, list all 5 in 'detected_colors' with their closest 'color_name' and 'color_map').
    3. Parse size range (e.g., 'M to XXXL' -> ["M", "L", "XL", "2XL", "3XL"]).
    4. NEXT-BEST VALUE RULE: For every dropdown field (in base_fields, amazon_overrides, and myntra_overrides), if the exact value is not in the allowed options list, you MUST select the NEXT BEST, most semantically relevant value from that dropdown list. Never leave a field blank if a reasonable equivalent exists.
    5. Write a rich, compelling 'description' highlighting the fabric feel, visual embroidery/print charm, silhouette, and styling.

    Return ONLY valid JSON matching this exact structure:
    {{
      "base_fields": {{
        "design_code": "string",
        "title_core": "string (e.g. Floral Cotton Embroidered Kurta Set with Pants)",
        "fabric": "string",
        "top_pattern": "string",
        "print_type": "string",
        "neck": "string",
        "sleeve_length": "string",
        "shape": "string",
        "top_length": "string",
        "weave": "string",
        "wash_care": "string",
        "occasion": "string",
        "mrp": 3999,
        "selling_price": 1499,
        "hsn": "62114210",
        "package_contains": "string",
        "sizes": ["M", "L", "XL", "2XL", "3XL"],
        "description": "string"
      }},
      "detected_colors": [
        {{"color_name": "Navy Blue", "color_map": "Blue"}}
      ],
      "amazon_overrides": {{
        "Field Name": "Prefilled Value"
      }},
      "myntra_overrides": {{
        "Field Name": "Prefilled Value"
      }}
    }}
    """

    response = client.models.generate_content(
        model=model_name,
        contents=[prompt] + compressed_imgs,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            temperature=0.1,
        ),
    )
    return json.loads(response.text)


# ==============================================================================
# SIDEBAR: AI STUDIO KEY & GLOBAL TEMPLATE SCHEMA LOADER
# ==============================================================================
with st.sidebar:
    st.header("🤖 AI & Template Engine (V2)")

    default_key = st.secrets.get("GEMINI_API_KEY", "") if "GEMINI_API_KEY" in st.secrets else ""
    gemini_api_key = st.text_input(
        "Gemini API Key (Google AI Studio)",
        value=default_key,
        type="password",
        help="Get a 100% free key from aistudio.google.com or store GEMINI_API_KEY in Streamlit Secrets.",
    )

    model_choice = st.selectbox(
        "Free-Tier Gemini Model",
        [
            "gemini-3.8-flash",
            "gemini-3.5-flash-lite",
            "gemini-3.5-flash",
            "Custom Model ID...",
        ],
        index=0,
        help="Free tier in Google AI Studio — multimodal vision + structured JSON.",
    )
    if model_choice == "Custom Model ID...":
        gemini_model = st.text_input("Enter Custom Model ID", value="gemini-3.8-flash")
    else:
        gemini_model = model_choice
        
    st.markdown("---")
    st.subheader("📑 Load Master Templates")
    st.caption("Upload your Amazon & Myntra templates here so the Add/Edit form can display every field and dropdown from both platforms.")

    tpl_amz_sidebar = st.file_uploader("Amazon Master Template (.xlsm/.xlsx)", type=["xlsm", "xlsx"], key="sb_amz_tpl")
    tpl_myn_sidebar = st.file_uploader("Myntra Master Template (.xlsx)", type=["xlsx"], key="sb_myn_tpl")

    if tpl_amz_sidebar:
        if st.session_state.get("loaded_amz_tpl_name") != tpl_amz_sidebar.name:
            with st.spinner("Extracting Amazon columns & dropdowns..."):
                st.session_state["amz_schema"] = parse_marketplace_template(tpl_amz_sidebar, "Amazon")
                st.session_state["loaded_amz_tpl_name"] = tpl_amz_sidebar.name
        st.success(f"Amazon: {len(st.session_state.get('amz_schema', {}))} fields loaded")

    if tpl_myn_sidebar:
        if st.session_state.get("loaded_myn_tpl_name") != tpl_myn_sidebar.name:
            with st.spinner("Extracting Myntra columns & dropdowns..."):
                st.session_state["myn_schema"] = parse_marketplace_template(tpl_myn_sidebar, "Myntra")
                st.session_state["loaded_myn_tpl_name"] = tpl_myn_sidebar.name
        st.success(f"Myntra: {len(st.session_state.get('myn_schema', {}))} fields loaded")

amz_schema = st.session_state.get("amz_schema", {})
myn_schema = st.session_state.get("myn_schema", {})

st.title("👗 Apparel Catalog & Multi-Marketplace Hub (V2 AI)")

# ==============================================================================
# 1. CATALOG MANAGEMENT (ADD / EDIT / ZIP AI INGESTION)
# ==============================================================================
col_t1, col_t2, col_t3, col_t4 = st.columns([2, 1, 1, 1])
with col_t1:
    st.subheader("📦 Master Catalog")
with col_t2:
    if st.button("➕ Add New Product (AI / Manual)", use_container_width=True, type="primary"):
        st.session_state["show_add_modal"] = not st.session_state.get("show_add_modal", False)
        st.session_state["edit_product_key"] = None
        st.session_state["ai_draft"] = {}
        st.session_state["temp_colors"] = [{"color_name": "White", "color_map": "White", "images": ""}]
with col_t3:
    if st.button("🗑️ Delete Selected", use_container_width=True):
        st.session_state["trigger_delete"] = True
with col_t4:
    catalog_json_str = json.dumps(db, indent=2)
    st.download_button(
        label="💾 Backup DB",
        data=catalog_json_str,
        file_name="catalog_db.json",
        mime="application/json",
        use_container_width=True,
    )

# Add / Edit Form Drawer
if st.session_state.get("show_add_modal", False) or st.session_state.get("edit_product_key", None):
    edit_key = st.session_state.get("edit_product_key", None)
    is_edit = edit_key is not None and edit_key in db

    # Merge existing DB record with any AI prefill draft
    base_record = db[edit_key] if is_edit else {}
    ai_draft = st.session_state.get("ai_draft", {})
    curr_data = {**base_record, **ai_draft}

    st.info(f"✏️ **{'Editing Style: ' + edit_key if is_edit else 'Add New Design to Catalog (AI Zip/Rar Prefill Enabled)'}**")

    # --------------------------------------------------------------------------
    # AI ARCHIVE DROPZONE (.ZIP / .RAR)
    # --------------------------------------------------------------------------
    with st.expander("✨ **Auto-Prefill from Product Archive (.zip / .rar)** — Drop folder like `Love Sharara.rar` or `Bansari.zip`", expanded=not is_edit):
        up_col1, up_col2 = st.columns([3, 1])
        with up_col1:
            product_archive = st.file_uploader(
                "Upload Product Folder (.zip or .rar containing images & text notes)",
                type=["zip", "rar"],
                key="archive_uploader",
            )
        with up_col2:
            st.write("")
            st.write("")
            run_ai_btn = st.button("🪄 Analyze & Prefill with AI", type="primary", use_container_width=True)

        if run_ai_btn:
            if not product_archive:
                st.warning("Please upload a .zip or .rar file first.")
            elif not gemini_api_key:
                st.error("Please enter your free Gemini API Key in the left sidebar.")
            else:
                with st.spinner("Extracting archive & analyzing product photos + notes with Gemini..."):
                    try:
                        raw_txt, extracted_imgs = extract_archive(product_archive)
                        st.session_state["last_extracted_imgs"] = extracted_imgs
                        st.session_state["last_extracted_txt"] = raw_txt

                        ai_res = run_gemini_archive_analysis(
                            api_key=gemini_api_key,
                            model_name=gemini_model,
                            raw_text=raw_txt,
                            images=extracted_imgs,
                            amz_schema=amz_schema,
                            myn_schema=myn_schema,
                            archive_name=product_archive.name,
                        )

                        bf = ai_res.get("base_fields", {})
                        # Snap AI base fields to closest valid dropdown values
                        bf["fabric"] = match_value_to_list(bf.get("fabric"), DROPDOWNS["fabrics"], "Pure Cotton")
                        bf["top_pattern"] = match_value_to_list(bf.get("top_pattern"), DROPDOWNS["top_patterns"], "Embroidered")
                        bf["print_type"] = match_value_to_list(bf.get("print_type"), DROPDOWNS["print_types"], "Floral")
                        bf["neck"] = match_value_to_list(bf.get("neck"), DROPDOWNS["neck_styles"], "Round Neck")
                        bf["sleeve_length"] = match_value_to_list(bf.get("sleeve_lengths") or bf.get("sleeve_length"), DROPDOWNS["sleeve_lengths"], "Three-Quarter Sleeves")
                        bf["shape"] = match_value_to_list(bf.get("shape"), DROPDOWNS["shapes"], "Straight")
                        bf["top_length"] = match_value_to_list(bf.get("top_length"), DROPDOWNS["top_lengths"], "Calf Length")
                        bf["weave"] = match_value_to_list(bf.get("weave"), DROPDOWNS["weave_types"], "Machine Weave")
                        bf["wash_care"] = match_value_to_list(bf.get("wash_care"), DROPDOWNS["wash_cares"], "Dry Clean")
                        bf["occasion"] = match_value_to_list(bf.get("occasion"), DROPDOWNS["occasions"], "Festive")
                        bf["package_contains"] = match_value_to_list(bf.get("package_contains"), DROPDOWNS["packages"], "1 Kurta, 1 Pant")

                        bf["amazon_overrides"] = ai_res.get("amazon_overrides", {})
                        bf["myntra_overrides"] = ai_res.get("myntra_overrides", {})

                        st.session_state["ai_draft"] = bf

                        # Populate detected color variants
                        det_colors = ai_res.get("detected_colors", [])
                        if det_colors:
                            new_temp_colors = []
                            for dc in det_colors:
                                cname = match_value_to_list(dc.get("color_name"), DROPDOWNS["colors"], "White")
                                cmap = match_value_to_list(dc.get("color_map"), DROPDOWNS["amazon_color_maps"], COLOR_MAPPER.get(cname, "White"))
                                new_temp_colors.append({"color_name": cname, "color_map": cmap, "images": ""})
                            st.session_state["temp_colors"] = new_temp_colors

                        st.toast("✅ AI Prefill Complete! Check the tabs below.", icon="🪄")
                        st.rerun()
                    except Exception as e:
                        st.error(f"AI Analysis Error: {e}")

        if st.session_state.get("last_extracted_imgs") or st.session_state.get("last_extracted_txt"):
            st.caption("📸 **Extracted Archive Preview:**")
            if st.session_state.get("last_extracted_txt"):
                st.code(st.session_state["last_extracted_txt"], language="text")
            imgs_preview = st.session_state.get("last_extracted_imgs", [])
            if imgs_preview:
                p_cols = st.columns(min(len(imgs_preview), 6))
                for idx, (iname, iobj) in enumerate(imgs_preview[:6]):
                    p_cols[idx].image(iobj, caption=iname, use_container_width=True)

    # Initialize temp_colors if needed
    if "temp_colors" not in st.session_state or (is_edit and st.session_state.get("loaded_edit_key") != edit_key):
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

    # --------------------------------------------------------------------------
    # MULTI-TAB ADD/EDIT FORM (CORE + ALL AMAZON FIELDS + ALL MYNTRA FIELDS)
    # --------------------------------------------------------------------------
    with st.form("product_form"):
        tab_core, tab_amz_all, tab_myn_all = st.tabs([
            "🏷️ 1. Core Attributes & Colorways",
            f"📦 2. All Amazon Template Fields ({len(amz_schema)})",
            f"👗 3. All Myntra Template Fields ({len(myn_schema)})",
        ])

        with tab_core:
            f1, f2, f3 = st.columns(3)
            with f1:
                st.markdown("### 🏷️ Base Identifiers")
                d_code = st.text_input("Design Code / Model Name*", value=curr_data.get("design_code", "NewStyle"))
                t_core = st.text_input("Core Title (Without Brand)*", value=curr_data.get("title_core", "Floral Cotton Embroidered Kurta Set with Pants"))

                fab_def = curr_data.get("fabric", "Pure Cotton")
                fab_idx = get_best_dropdown_index(fab_def, DROPDOWNS["fabrics"])
                fab_choice = st.selectbox("Fabric*", DROPDOWNS["fabrics"], index=max(0, fab_idx - 1))
                fab = st.text_input("Type Custom Fabric", value=fab_def) if fab_choice == "Other / Custom..." else fab_choice

                pat_def = curr_data.get("top_pattern", "Embroidered")
                pat_idx = get_best_dropdown_index(pat_def, DROPDOWNS["top_patterns"])
                pat_choice = st.selectbox("Pattern / Work*", DROPDOWNS["top_patterns"], index=max(0, pat_idx - 1))
                pat = st.text_input("Type Custom Pattern", value=pat_def) if pat_choice == "Other / Custom..." else pat_choice

                prt_def = curr_data.get("print_type", "Floral")
                prt_idx = get_best_dropdown_index(prt_def, DROPDOWNS["print_types"])
                prt_choice = st.selectbox("Print Type*", DROPDOWNS["print_types"], index=max(0, prt_idx - 1))
                prt = st.text_input("Type Custom Print", value=prt_def) if prt_choice == "Other / Custom..." else prt_choice

            with f2:
                st.markdown("### ✂️ Cut & Styling")
                nck_def = curr_data.get("neck", "V-Neck")
                nck_idx = get_best_dropdown_index(nck_def, DROPDOWNS["neck_styles"])
                nck_choice = st.selectbox("Neck Type*", DROPDOWNS["neck_styles"], index=max(0, nck_idx - 1))
                nck = st.text_input("Type Custom Neck", value=nck_def) if nck_choice == "Other / Custom..." else nck_choice

                slv_def = curr_data.get("sleeve_length", "Three-Quarter Sleeves")
                slv_idx = get_best_dropdown_index(slv_def, DROPDOWNS["sleeve_lengths"])
                slv_choice = st.selectbox("Sleeve Length*", DROPDOWNS["sleeve_lengths"], index=max(0, slv_idx - 1))
                slv = st.text_input("Type Custom Sleeve", value=slv_def) if slv_choice == "Other / Custom..." else slv_choice

                shp_def = curr_data.get("shape", "Straight")
                shp_idx = get_best_dropdown_index(shp_def, DROPDOWNS["shapes"])
                shp_choice = st.selectbox("Kurta Shape / Fit*", DROPDOWNS["shapes"], index=max(0, shp_idx - 1))
                shp = st.text_input("Type Custom Shape", value=shp_def) if shp_choice == "Other / Custom..." else shp_choice

                len_def = curr_data.get("top_length", "Knee Length")
                len_idx = get_best_dropdown_index(len_def, DROPDOWNS["top_lengths"])
                len_choice = st.selectbox("Kurta / Top Length*", DROPDOWNS["top_lengths"], index=max(0, len_idx - 1))
                top_len_val = st.text_input("Type Custom Length", value=len_def) if len_choice == "Other / Custom..." else len_choice

                wev_def = curr_data.get("weave", "Machine Weave")
                wev_idx = get_best_dropdown_index(wev_def, DROPDOWNS["weave_types"])
                wev_choice = st.selectbox("Weave Type*", DROPDOWNS["weave_types"], index=max(0, wev_idx - 1))
                wev = st.text_input("Type Custom Weave", value=wev_def) if wev_choice == "Other / Custom..." else wev_choice

                wsh_def = curr_data.get("wash_care", "Dry Clean")
                wsh_idx = get_best_dropdown_index(wsh_def, DROPDOWNS["wash_cares"])
                wsh_choice = st.selectbox("Wash Care*", DROPDOWNS["wash_cares"], index=max(0, wsh_idx - 1))
                wsh = st.text_input("Type Custom Wash Care", value=wsh_def) if wsh_choice == "Other / Custom..." else wsh_choice

                occ_def = curr_data.get("occasion", "Festive")
                occ_idx = get_best_dropdown_index(occ_def, DROPDOWNS["occasions"])
                occ_choice = st.selectbox("Occasion*", DROPDOWNS["occasions"], index=max(0, occ_idx - 1))
                occ = st.text_input("Type Custom Occasion", value=occ_def) if occ_choice == "Other / Custom..." else occ_choice

            with f3:
                st.markdown("### 💰 Pricing & Standard Curve")
                mrp_val = st.number_input("MRP (₹)*", value=int(curr_data.get("mrp", 3999)), step=100)
                sp_val = st.number_input("Selling Price (₹)*", value=int(curr_data.get("selling_price", 1499)), step=50)
                hsn_val = st.text_input("HSN Code*", value=str(curr_data.get("hsn", "62114210")))

                pkg_def = curr_data.get("package_contains", "1 Kurta, 1 Pant")
                pkg_idx = get_best_dropdown_index(pkg_def, DROPDOWNS["packages"])
                pkg_choice = st.selectbox("Package Contains*", DROPDOWNS["packages"], index=max(0, pkg_idx - 1))
                pkg = st.text_input("Type Custom Package", value=pkg_def) if pkg_choice == "Other / Custom..." else pkg_choice

                valid_sz_opts = ["XS", "S", "M", "L", "XL", "2XL", "3XL"]
                default_sizes = [s for s in curr_data.get("sizes", ["S", "M", "L", "XL", "2XL"]) if s in valid_sz_opts]
                sz_list = st.multiselect(
                    "Available Sizes*",
                    valid_sz_opts,
                    default=default_sizes or ["M", "L", "XL", "2XL"],
                )

            st.markdown("---")
            st.markdown("### 🎨 Color Variants & Image Sets")
            st.caption("AI auto-populates all detected colors from your zip/rar photos. You can also edit or add more below.")

            updated_color_variants = []
            for idx, cv in enumerate(st.session_state["temp_colors"]):
                st.markdown(f"**Colorway #{idx + 1}**")
                c_col1, c_col2, c_col3 = st.columns([1.5, 1.5, 3])
                with c_col1:
                    cur_c = cv.get("color_name", "White")
                    c_idx = get_best_dropdown_index(cur_c, DROPDOWNS["colors"])
                    col_name_sel = st.selectbox(f"Color Name #{idx + 1}", DROPDOWNS["colors"], index=max(0, c_idx - 1), key=f"cname_{idx}")
                    if col_name_sel == "Other / Custom...":
                        col_name = st.text_input(f"Type Custom Color #{idx + 1}", value=cur_c if cur_c != "Other / Custom..." else "", key=f"ccustom_{idx}")
                    else:
                        col_name = col_name_sel

                with c_col2:
                    suggested_map = COLOR_MAPPER.get(col_name_sel, cv.get("color_map", "Multicolor"))
                    map_idx = get_best_dropdown_index(suggested_map, DROPDOWNS["amazon_color_maps"])
                    col_map_choice = st.selectbox(
                        f"Amazon Filter Bucket #{idx + 1}",
                        DROPDOWNS["amazon_color_maps"],
                        index=max(0, map_idx - 1),
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
                height=120,
            )

        # ----------------------------------------------------------------------
        # TAB 2: ALL AMAZON TEMPLATE FIELDS (GRANULAR CONTROL)
        # ----------------------------------------------------------------------
        amz_override_inputs = {}
        with tab_amz_all:
            if not amz_schema:
                st.info("👈 Upload your **Amazon Master Template (.xlsm / .xlsx)** in the left sidebar to unlock granular control over every Amazon column and dropdown.")
            else:
                st.caption("Every column from your Amazon template is shown below. Dropdowns are pulled directly from the template, and AI pre-selects the closest matching value.")
                saved_amz_ovr = curr_data.get("amazon_overrides", {})
                amz_grid = st.columns(3)
                for idx, (fname, fmeta) in enumerate(amz_schema.items()):
                    col_ui = amz_grid[idx % 3]
                    cur_val = str(saved_amz_ovr.get(fname, "")).strip()
                    opts = fmeta.get("options", [])
                    if opts:
                        full_opts = [""] + opts
                        best_idx = get_best_dropdown_index(cur_val, opts)
                        amz_override_inputs[fname] = col_ui.selectbox(
                            f"📦 {fname}",
                            options=full_opts,
                            index=best_idx,
                            key=f"amz_gran_{fname}",
                        )
                    else:
                        amz_override_inputs[fname] = col_ui.text_input(
                            f"📦 {fname}",
                            value=cur_val,
                            key=f"amz_gran_{fname}",
                        )

        # ----------------------------------------------------------------------
        # TAB 3: ALL MYNTRA TEMPLATE FIELDS (GRANULAR CONTROL)
        # ----------------------------------------------------------------------
        myn_override_inputs = {}
        with tab_myn_all:
            if not myn_schema:
                st.info("👈 Upload your **Myntra Master Template (.xlsx)** in the left sidebar to unlock granular control over every Myntra column and dropdown.")
            else:
                st.caption("Every column from your Myntra template is shown below. Dropdowns are pulled directly from the template, and AI pre-selects the closest matching value.")
                saved_myn_ovr = curr_data.get("myntra_overrides", {})
                myn_grid = st.columns(3)
                for idx, (fname, fmeta) in enumerate(myn_schema.items()):
                    col_ui = myn_grid[idx % 3]
                    cur_val = str(saved_myn_ovr.get(fname, "")).strip()
                    opts = fmeta.get("options", [])
                    if opts:
                        full_opts = [""] + opts
                        best_idx = get_best_dropdown_index(cur_val, opts)
                        myn_override_inputs[fname] = col_ui.selectbox(
                            f"👗 {fname}",
                            options=full_opts,
                            index=best_idx,
                            key=f"myn_gran_{fname}",
                        )
                    else:
                        myn_override_inputs[fname] = col_ui.text_input(
                            f"👗 {fname}",
                            value=cur_val,
                            key=f"myn_gran_{fname}",
                        )

        save_c1, save_c2 = st.columns([1, 4])
        with save_c1:
            save_btn = st.form_submit_button("💾 Save Product", type="primary", use_container_width=True)
        with save_c2:
            cancel_btn = st.form_submit_button("Cancel", use_container_width=False)

        if cancel_btn:
            st.session_state["show_add_modal"] = False
            st.session_state["edit_product_key"] = None
            st.session_state["ai_draft"] = {}
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

            clean_amz_ovr = {k: v for k, v in amz_override_inputs.items() if str(v).strip() != ""} if amz_schema else curr_data.get("amazon_overrides", {})
            clean_myn_ovr = {k: v for k, v in myn_override_inputs.items() if str(v).strip() != ""} if myn_schema else curr_data.get("myntra_overrides", {})

            db[d_code] = {
                "design_code": d_code,
                "title_core": t_core,
                "fabric": fab,
                "top_pattern": pat,
                "print_type": prt,
                "neck": nck,
                "sleeve_length": slv,
                "shape": shp,
                "top_length": top_len_val,
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
                "amazon_overrides": clean_amz_ovr,
                "myntra_overrides": clean_myn_ovr,
            }
            save_db(db)
            st.session_state["show_add_modal"] = False
            st.session_state["edit_product_key"] = None
            st.session_state["ai_draft"] = {}
            st.rerun()

    col_btn1, col_btn2 = st.columns([1, 4])
    with col_btn1:
        if st.button("➕ Add Another Color Variant"):
            st.session_state["temp_colors"].append({"color_name": "White", "color_map": "White", "images": ""})
            st.rerun()
    with col_btn2:
        if len(st.session_state["temp_colors"]) > 1 and st.button("➖ Remove Last Color Variant"):
            st.session_state["temp_colors"].pop()
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
            "Top Length": item.get("top_length", "Knee Length"),
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
            st.rerun()

    if len(selected_designs) == 1:
        if st.button(f"✏️ Edit Selected ({selected_designs[0]})"):
            st.session_state["edit_product_key"] = selected_designs[0]
            st.session_state["show_add_modal"] = False
            st.session_state["ai_draft"] = {}
            st.rerun()

    st.markdown("---")

    # ==============================================================================
    # 3. LISTING HUB (AUTOMATED PARENT-CHILD + GRANULAR OVERRIDE INJECTION)
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
                    ["KALINI", "MITERA", "PERVAS", "PERABIN"],
                    default=["KALINI", "MITERA", "PERVAS", "PERABIN"],
                )
            else:
                selected_brands = st.multiselect(
                    "Select Brand(s)",
                    ["PERVAS", "PERABIN"],
                    default=["PERVAS"],
                )

        with c_tpl:
            uploaded_template = st.file_uploader(
                f"Upload Latest {marketplace} Template (.xlsx / .xlsm)",
                type=["xlsx", "xlsm"],
            )

        if st.button(f"⚡ Generate {marketplace} Bulk Upload File for {len(selected_designs)} Product(s)", type="primary", use_container_width=True):
            # Fallback to sidebar master template if not uploaded in the bottom box
            active_tpl = uploaded_template or (tpl_myn_sidebar if marketplace == "Myntra" else tpl_amz_sidebar)

            if not active_tpl:
                st.error(f"Please upload the {marketplace} template file (here or in the sidebar).")
            elif not selected_brands:
                st.error("Please select at least one brand.")
            else:
                active_tpl.seek(0)
                is_xlsm = active_tpl.name.lower().endswith(".xlsm")
                keep_vba_flag = True if is_xlsm else False
                out_ext = ".xlsm" if is_xlsm else ".xlsx"
                out_mime = "application/vnd.ms-excel.sheet.macroEnabled.12" if is_xlsm else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                wb = openpyxl.load_workbook(active_tpl, keep_vba=keep_vba_flag)

                if marketplace == "Myntra":
                    sheet_name = "Kurta Sets" if "set" in category_val.lower() else "Kurtas"
                    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
                    header_row = 3
                    start_row = 4
                else:
                    ws = wb["Template"] if "Template" in wb.sheetnames else wb.active
                    header_row = 4
                    start_row = 7
                    for c in range(1, ws.max_column + 1):
                        ws.cell(row=6, column=c).value = None

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
                    if value is None or str(value).strip() == "":
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
                curr_year = datetime.datetime.now().year

                for d_name in selected_designs:
                    prod = db[d_name]
                    colorways = prod.get("color_variants", [{"color_name": "White", "color_map": "White", "images": []}])
                    amz_ovr = prod.get("amazon_overrides", {})
                    myn_ovr = prod.get("myntra_overrides", {})

                    user_top_len = prod.get("top_length", "Knee Length")
                    len_tuple = LENGTH_MAPPER.get(user_top_len, (user_top_len, user_top_len))
                    myntra_top_len_val = len_tuple[0]
                    amazon_item_len_val = len_tuple[1]

                    raw_fabric = prod.get("fabric", "Pure Cotton")
                    amz_fabric_val = FABRIC_MAPPER.get(raw_fabric, "Cotton")

                    for brand in selected_brands:
                        brand_email = "pervas0007@gmail.com" if brand == "PERVAS" else "perabin32544@gmail.com"

                        # ------------------------------------------------------
                        # AMAZON PARENT ROW
                        # ------------------------------------------------------
                        if marketplace == "Amazon.in":
                            parent_sku = f"{brand}-{prod.get('design_code')}-Parent"
                            parent_title = f"{brand} Women's {raw_fabric} {prod.get('top_pattern', 'Printed')} Kurta Pant Set ({prod.get('design_code')})"

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
                                "Manufacturer": brand_email,
                                "Product Description": prod.get("description") or AMAZON_EXACT_DESCRIPTION,
                                "Generic Keywords": "two piece suit set coord dress stylish co ord sets cord coords dresses a line neck ords long salwar trendy new printed latest design traditional suits ladies daily rayon festive kutis indian kutties v kurthi fashion mul weddings floral print",
                                "Lifestyle": "Casual",
                                "Department Name": "Womens",
                                "Target Gender": "Female",
                                "Age Range Description": "Adult",
                                "Fabric Type": amz_fabric_val,
                                "Material": amz_fabric_val,
                                "Item Type Name": category_val.upper(),
                                "Item Length Description": amazon_item_len_val,
                                "Occasion": prod.get("occasion", "Festive"),
                                "Care Instructions": prod.get("wash_care", "Dry Clean Only"),
                                "Manufacturer Contact Information": brand_email,
                                "Packer Contact Information": brand_email,
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
                                "Number of Items": 1,
                                "Item Length Longer Edge": 46.0,
                                "Item Length Unit": "Centimeters",
                                "Package Length": 25.0,
                                "Package Length Unit": "Centimeters",
                                "Package Width": 22.0,
                                "Package Width Unit": "Centimeters",
                                "Package Height": 3.0,
                                "Package Height Unit": "Centimeters",
                                "Package Weight": 450.0,
                                "Package Weight Unit": "Grams",
                            }
                            # Apply any granular Amazon overrides (without clobbering Parentage/SKU)
                            protected_parent_keys = {"SKU", "Parentage Level", "Parent SKU", "Brand Name"}
                            for ok, ov in amz_ovr.items():
                                if ok not in protected_parent_keys and str(ov).strip():
                                    p_row[ok] = ov

                            for col_name, val in p_row.items():
                                write_cell(current_row, col_name, val)

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
                                amz_size_val = AMAZON_SIZE_MAP.get(sz.upper(), sz)
                                myntra_size = "XXL" if sz.upper() in ["2XL", "XXL"] else sz
                                m = prod.get("measurements", {}).get(sz, {})

                                if marketplace == "Myntra":
                                    sku = f"{brand}-P-{prod.get('design_code')}-{c_name}-{myntra_size}"
                                    art_num = f"{brand}-P-{prod.get('design_code')}-{c_name}"
                                    display_name = f"{brand} {c_name} {prod.get('title_core', '')}"

                                    style_note_text = (
                                        f"Step into elegance with this mesmerizing Kurta set by {brand}. "
                                        f"Whether you are heading to a family dinner or a festive celebration, the lush {raw_fabric} fabric "
                                        f"creates a look that is both traditional and trendy. The {prod.get('neck', 'classic')} neckline adds a touch of modern grace, "
                                        f"while the flattering silhouette ensures you stay comfortable all day long. "
                                        f"Pair this ensemble with oxidized jewelry and block heels for a complete festive look."
                                    )

                                    m_row = {
                                        "styleGroupId": group_id_counter,
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
                                        "Year": curr_year,
                                        "Product Details": prod.get("description"),
                                        "styleNote": style_note_text,
                                        "materialCareDescription": prod.get("wash_care", "Dry Clean Only"),
                                        "productDisplayName": display_name,
                                        "Top Fabric": raw_fabric,
                                        "Top Pattern": prod.get("top_pattern"),
                                        "Neck": prod.get("neck"),
                                        "Sleeve Length": prod.get("sleeve_length"),
                                        "Top Shape": prod.get("shape"),
                                        "Top Length": myntra_top_len_val,
                                        "Bottom Fabric": raw_fabric,
                                        "Bottom Pattern": prod.get("top_pattern"),
                                        "Bottom Closure": "NA",
                                        "Waistband": "NA",
                                        "Print or Pattern Type": prod.get("print_type"),
                                        "Occasion": prod.get("occasion"),
                                        "Weave Pattern": "Regular",
                                        "Weave Type": prod.get("weave"),
                                        "Wash Care": prod.get("wash_care"),
                                        "Stitch": "Ready to Wear",
                                        "Add-Ons": "NA",
                                        "Number of Pockets": "NA",
                                        "Number of Items": 1,
                                        "Net Quantity Unit": "Piece",
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

                                    # Apply granular Myntra overrides from Tab 3
                                    protected_myn_keys = {"styleGroupId", "vendorSkuCode", "vendorArticleNumber", "brand", "Brand Size", "Standard Size", "Brand Colour (Remarks)", "Prominent Colour", "SKUCode"}
                                    for ok, ov in myn_ovr.items():
                                        if ok not in protected_myn_keys and str(ov).strip():
                                            m_row[ok] = ov

                                    for col_name, val in m_row.items():
                                        write_cell(current_row, col_name, val)

                                else:
                                    # AMAZON CHILD ROW
                                    child_sku = f"{brand}-{prod.get('design_code')}-{c_name}-{sz}"
                                    parent_sku = f"{brand}-{prod.get('design_code')}-Parent"
                                    child_title = f"{brand} Women's {raw_fabric} Kurta Pant Set ({prod.get('design_code')} {c_name} {sz})"

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
                                        "Apparel Size Value": amz_size_val,
                                        "Shirt Size System": "IN",
                                        "Shirt Size Class": "Alpha",
                                        "Shirt Size Value": amz_size_val,
                                        "Shirt Body Type": "Regular",
                                        "Special Size": "Standard",
                                        "Color": c_name,
                                        "Color Map": c_map,
                                        "Standard Price": prod.get("selling_price"),
                                        "Maximum Retail Price": prod.get("mrp"),
                                        "Your Price INR (Sell on Amazon, IN)": prod.get("selling_price"),
                                        "Maximum Retail Price (Sell on Amazon, IN)": prod.get("mrp"),
                                        "Product Description": prod.get("description") or AMAZON_EXACT_DESCRIPTION,
                                        "Generic Keywords": "two piece suit set coord dress stylish co ord sets cord coords dresses a line neck ords long salwar trendy new printed latest design traditional suits ladies daily rayon festive kutis indian kutties v kurthi fashion mul weddings floral print",
                                        "Lifestyle": "Casual",
                                        "Style": prod.get("shape", "Straight"),
                                        "Fit Type": "Regular",
                                        "Department Name": "Womens",
                                        "Target Gender": "Female",
                                        "Age Range Description": "Adult",
                                        "Fabric Type": amz_fabric_val,
                                        "Material": amz_fabric_val,
                                        "Pattern": prod.get("print_type", "Floral"),
                                        "Item Type Name": category_val.upper(),
                                        "Item Length Description": amazon_item_len_val,
                                        "Occasion": prod.get("occasion", "Festive"),
                                        "Care Instructions": prod.get("wash_care", "Dry Clean Only"),
                                        "Manufacturer": brand_email,
                                        "Manufacturer Contact Information": brand_email,
                                        "Packer Contact Information": brand_email,
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
                                        "Number of Items": 1,
                                        "Item Length Longer Edge": 46.0,
                                        "Item Length Unit": "Centimeters",
                                        "Package Length": 25.0,
                                        "Package Length Unit": "Centimeters",
                                        "Package Width": 22.0,
                                        "Package Width Unit": "Centimeters",
                                        "Package Height": 3.0,
                                        "Package Height Unit": "Centimeters",
                                        "Package Weight": 450.0,
                                        "Package Weight Unit": "Grams",
                                    }

                                    # Apply granular Amazon overrides from Tab 2
                                    protected_child_keys = {"SKU", "Parentage Level", "Parent SKU", "Brand Name", "Apparel Size Value", "Shirt Size Value", "Color", "Color Map"}
                                    for ok, ov in amz_ovr.items():
                                        if ok not in protected_child_keys and str(ov).strip():
                                            c_row[ok] = ov

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
