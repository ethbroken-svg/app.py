import streamlit as st
import pandas as pd
import numpy as np
import io

# =====================================================================
# 1. LIVE WEB ENGINE INITIALIZATION & CONFIGURATION
# =====================================================================
st.set_page_type = "wide"
st.title("🏗️ Smart PVC Inventory Manifest & Sizing Web App")
st.write("Upload a flat layout photo to detect counts, verify sizing via Google Sheets, and export to Excel.")

# Google Sheet Rate Link Setup
GOOGLE_SHEET_LINK = "https://google.com"
CSV_URL = GOOGLE_SHEET_LINK.replace("/pubhtml?widget=true&headers=false", "/pub?output=csv")

# =====================================================================
# 2. DATA PROCESSING & MODEL CLASSIFICATION ENGINES
# =====================================================================
@st.cache_data(ttl=60) # Caches the data for 60 seconds to prevent hitting Google limits on every click
def fetch_cloud_rates():
    try:
        df_rates = pd.read_csv(CSV_URL)
        df_rates.columns = df_rates.columns.str.strip()
        # Create lookup keys: (Item_Type, Nominal_Size) -> Unit_Rate
        rate_map = dict(zip(zip(df_rates['Item_Type'], df_rates['Nominal_Size']), df_rates['Unit_Rate']))
        return rate_map
    except Exception as e:
        st.warning(f"Using offline backup price catalog. Unable to read cloud spreadsheet: {e}")
        return {
            ("elbow", "20mm"): 14.50, ("elbow", "25mm"): 19.00, ("elbow", "32mm"): 28.50,
            ("tee", "20mm"): 18.00,   ("tee", "25mm"): 24.50,   ("tee", "32mm"): 36.00
        }

def classify_dimension_size(item_type, bounding_box_area, anchor_median):
    """Calibrates 8 relative size variants dynamically based on flat surface area scales"""
    if item_type == "elbow":
        if bounding_box_area < anchor_median * 0.7: return "20mm"
        elif anchor_median * 0.7 <= bounding_box_area < anchor_median * 1.2: return "25mm"
        else: return "32mm"
    elif item_type == "tee":
        if bounding_box_area < anchor_median * 0.7: return "20mm"
        elif anchor_median * 0.7 <= bounding_box_area < anchor_median * 1.1: return "25mm"
        elif anchor_median * 1.1 <= bounding_box_area < anchor_median * 1.4: return "32mm"
        else: return "40mm"
    return "20mm"

# Pull fresh rates
rate_database = fetch_cloud_rates()

# =====================================================================
# 3. INTERACTIVE WEBSITE USER INTERFACE
# =====================================================================
# Session state initialization to hold user adjustments across app clicks
if 'inventory' not in st.session_state:
    st.session_state.inventory = None

uploaded_image = st.file_uploader("Step 1: Choose a flat layout photograph...", type=["jpg", "png", "jpeg"])

if uploaded_image is not None:
    st.image(uploaded_image, caption="Uploaded Surface Layout Scan", use_container_width=True)
    
    # Process AI Detection Simulation immediately on image drop
    if st.session_state.inventory is None:
        st.info("🤖 Simulating YOLOv8 model inference extraction on flat layout...")
        
        # Simulated raw bounding box pixel matrices extracted from your image structure
        mock_detected_boxes = [
            ["elbow", 145, 120], ["elbow", 130, 115], 
            ["elbow", 190, 165], ["elbow", 175, 155], 
            ["tee", 170, 230], ["tee", 195, 260]      
        ]
        
        # Calculate dynamic median surface reference scale
        calculated_areas = [w * h for _, w, h in mock_detected_boxes]
        median_scale = np.median(calculated_areas)
        
        # Classify sizes and construct clean manifest list
        parsed_results = []
        for itype, width, height in mock_detected_boxes:
            area = width * height
            inferred_size = classify_dimension_size(itype, area, median_scale)
            parsed_results.append({"Item": itype, "Size": inferred_size, "Quantity": 1})
            
        # Group duplicates to create a base starting count list
        df_init = pd.DataFrame(parsed_results).groupby(["Item", "Size"]).sum().reset_index()
        st.session_state.inventory = df_init.to_dict('records')

    # =====================================================================
    # 4. WORKER EDIT VIEW AND MANUAL VERIFICATION ZONE
    # =====================================================================
    st.subheader("Step 2: Verify & Adjust Quantities")
    st.write("Mismatched counts can be corrected manually below prior to generating your invoice:")
    
    # Render interactive input boxes row-by-row
    updated_manifest = []
    for idx, row in enumerate(st.session_state.inventory):
        col1, col2, col3 = st.columns([3, 3, 2])
        with col1:
            st.write(f"**Item Type:** {row['Item'].upper()}")
        with col2:
            st.write(f"**Nominal Sizing:** {row['Size']}")
        with col3:
            new_qty = st.number_input(f"Qty for line {idx+1}", min_value=0, value=int(row['Quantity']), step=1, key=f"qty_{idx}")
            updated_manifest.append({"Item": row['Item'], "Size": row['Size'], "Quantity": new_qty})

    # Save tracking changes
    st.session_state.inventory = updated_manifest

    # =====================================================================
    # 5. EXCEL GENERATION ENGINE & LIVE PREVIEW
    # =====================================================================
    if st.button("Step 3: Compile Report with Google Sheet Rates"):
        st.subheader("Final Bill of Materials Ledger")
        
        final_rows = []
        for row in st.session_state.inventory:
            if row['Quantity'] > 0: # Skip empty items
                unit_price = rate_database.get((row['Item'], row['Size']), 0.0)
                subtotal = unit_price * row['Quantity']
                final_rows.append({
                    "Item Description": row['Item'].upper(),
                    "Dimension Scale": row['Size'],
                    "Verified Quantity": row['Quantity'],
                    "Unit Price (INR)": unit_price,
                    "Total Value (INR)": subtotal
                })
                
        if final_rows:
            df_final = pd.DataFrame(final_rows)
            st.dataframe(df_final, use_container_width=True)
            
            grand_total = df_final["Total Value (INR)"].sum()
            st.metric(label="Grand Invoice Total (INR)", value=f"₹{grand_total:,.2f}")
            
            # Write to in-memory Excel file stream using openpyxl
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Inventory Output', index=False, startrow=2)
                
                # Add customized design header to sheet template
                worksheet = writer.sheets['Inventory Output']
                worksheet['A1'] = "AUTOMATED SITE INVENTORY MANIFEST"
                
                # Mount summary metrics at the foot rows
                last_row = len(df_final) + 4
                worksheet[f'C{last_row}'] = "GRAND TOTAL:"
                worksheet[f'E{last_row}'] = grand_total
                
            excel_buffer.seek(0)
            
            # Streamlit Download Button triggers direct download inside browser environment
            st.download_button(
                label="📥 Download Excel Spreadsheet",
                data=excel_buffer,
                file_name="PVC_Field_Inventory_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("All quantities set to 0. Cannot compile blank spreadsheet manifest ledger.")
