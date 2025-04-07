import streamlit as st
import os
import pandas as pd
from io import BytesIO
from processed import parse_json_file 

st.title("Prosparity Data Extractor")

uploaded_files = st.file_uploader(
    "Upload multiple JSON files", 
    type="json", 
    accept_multiple_files=True
)

if uploaded_files:
    results = []

    for uploaded_file in uploaded_files:
        try:
            # Load JSON content
            data = uploaded_file.read().decode("utf-8")
            json_data = pd.read_json(data)
        except Exception:
            st.error(f"Failed to read {uploaded_file.name}")
            continue

        # Save to temporary path for reuse of your function
        temp_path = f"temp_{uploaded_file.name}"
        with open(temp_path, "w", encoding="utf-8") as f:
            f.write(data)

        # Process using your existing parser
        parsed = parse_json_file(temp_path)
        if parsed:
            row = {
                "Application Current Status": "From Processed Application Report",
                "Consumer Name": parsed["consumer_name"],
                "Gender": parsed["gender"],
                "PAN": parsed["pan"],
                "Address": parsed["address"],
                "DOB": parsed["dob"],
                "State": parsed["state"],
                "Mobile": parsed["mobile"],
                "Age": parsed["age"],
                "Total Income": "",  # Optional to manually input later
                "Bureau Score": parsed["bureau_score"],
                "Account Institutions": str(parsed["institutions"]),
                "Account AccountTypes": str(parsed["account_types"]),
                "Account OwnershipTypes": str(parsed["ownership_types"]),
            }
            results.append(row)
    
    if results:
        df = pd.DataFrame(results)
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.success(f"Generated report with {len(results)} entries.")
        st.download_button(
            "⬇️ Download Combined Excel Report",
            output,
            file_name="combined_bureau_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
