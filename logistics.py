import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Global Logistics Log", layout="wide")

# 1. CONNECT TO GOOGLE SHEETS
# Note: You will set the URL in your Streamlit Cloud Secrets later
conn = st.connection("gsheets", type=GSheetsConnection)

def load_data():
    return conn.read(worksheet="Sheet1")

st.title("🌐 Shared Logistics Manager")

# 2. THE 7 SECTIONS FORM
with st.form("logistics_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        dept = st.text_input("1. To Whom (Department)")
        prod = st.text_input("2. Product Name")
        unit = st.selectbox("3. Unit", ["kg", "m", "l", "units", "pcs"])
        qty = st.number_input("4. Quantity", min_value=0.0)
    with col2:
        loc_from = st.text_input("5. Departed Location")
        loc_to = st.text_input("6. Delivery Location")
        status = st.selectbox("7. Status", ["not departed", "in transit", "late", "arrived"])
    
    submit = st.form_submit_button("Save to Cloud")

    if submit:
        if dept and prod:
            # Create new row
            new_row = pd.DataFrame([{
                "Date": datetime.now().strftime("%Y-%m-%d"),
                "Department": dept, "Product": prod, "Unit": unit,
                "Quantity": qty, "From": loc_from, "To": loc_to, "Status": status
            }])
            # Update Google Sheet
            existing_data = load_data()
            updated_df = pd.concat([existing_data, new_row], ignore_index=True)
            conn.update(worksheet="Sheet1", data=updated_df)
            st.success("✅ Shared Database Updated!")
        else:
            st.error("Please fill Department and Product.")

# 3. VIEW & EXCEL EXPORT
st.divider()
df = load_data()
st.subheader("Current Shared Records")
st.dataframe(df, use_container_width=True)

# 4. DOWNLOAD AS COLORED EXCEL
if not df.empty:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Logistics')
        workbook  = writer.book
        worksheet = writer.sheets['Logistics']
        
        # Colors (Teal, Green, Red)
        f1 = workbook.add_format({'bg_color': '#B7DEE8'}) # Not Departed
        f2 = workbook.add_format({'bg_color': '#C6EFCE'}) # Transit
        f3 = workbook.add_format({'bg_color': '#FFC7CE'}) # Late
        
        # Apply to Status column
        worksheet.conditional_format(1, 7, len(df), 7, {'type': 'cell', 'criteria': 'equal to', 'value': '"not departed"', 'format': f1})
        worksheet.conditional_format(1, 7, len(df), 7, {'type': 'cell', 'criteria': 'equal to', 'value': '"in transit"', 'format': f2})
        worksheet.conditional_format(1, 7, len(df), 7, {'type': 'cell', 'criteria': 'equal to', 'value': '"late"', 'format': f3})
        
    st.download_button(
        label="📥 Download Excel Report (Colored)",
        data=buffer.getvalue(),
        file_name=f"logistics_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.ms-excel"
    )