import os
import glob
import json
import pandas as pd
from datetime import datetime

def parse_json_file(file_path):
    """Parse JSON file and extract required information."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        print(f"Error reading {file_path}: {str(e)}")
        return None

    report = data.get("equifaxReport", {})
    id_info = report.get("IDAndContactInfo", {})
    personal = id_info.get("PersonalInfo", {})
    identity = id_info.get("IdentityInfo", {})

    # Personal Information
    name_info = personal.get("Name", {})
    consumer_name = name_info.get("FullName", "").strip()
    
    # Date handling with validation
    dob = personal.get("DateOfBirth", "")
    try:
        dob_datetime = pd.to_datetime(dob, errors='coerce')
        formatted_dob = dob_datetime.strftime("%Y-%m-%d %H:%M:%S") if not pd.isnull(dob_datetime) else ""
    except Exception:
        formatted_dob = ""

    # Account Details Extraction
    accounts = report.get("RetailAccountDetails", [])
    institutions, account_types, ownership_types = [], [], []
    
    for acc in accounts:
        institutions.append(acc.get("Institution", ""))
        account_types.append(acc.get("AccountType", ""))
        ownership_types.append(acc.get("OwnershipType", ""))

    # Create formatted dictionaries with unique values
    def create_indexed_dict(items):
        return {idx+1: item for idx, item in enumerate(sorted(set(filter(None, items))))}

    return {
        "consumer_name": consumer_name,
        "gender": personal.get("Gender", "").strip(),
        "dob": formatted_dob,
        "age": personal.get("Age", {}).get("Age", "").strip(),
        "pan": identity.get("PANId", [{}])[-1].get("IdNumber", "").strip(),
        "address": id_info.get("AddressInfo", [{}])[-1].get("Address", "").strip(),
        "state": id_info.get("AddressInfo", [{}])[-1].get("State", "").strip(),
        "mobile": id_info.get("PhoneInfo", [{}])[-1].get("Number", "").strip(),
        "bureau_score": str(report.get("ScoreDetails", [{}])[0].get("Value", "")),
        "institutions": create_indexed_dict(institutions),
        "account_types": create_indexed_dict(account_types),
        "ownership_types": create_indexed_dict(ownership_types)
    }

def generate_excel_output(json_folder, output_file):
    """Process JSON files and generate Excel report."""
    results = []
    
    for json_file in glob.glob(os.path.join(json_folder, "*.json")):
        file_data = parse_json_file(json_file)
        if not file_data:
            continue

        row = {
            "Application Current Status": "From Processed Application Report",
            "Consumer Name": file_data["consumer_name"],
            "Gender": file_data["gender"],
            "PAN": file_data["pan"],
            "Address": file_data["address"],
            "DOB": file_data["dob"],
            "State": file_data["state"],
            "Mobile": file_data["mobile"],
            "Age": file_data["age"],
            "Total Income": "",  # Placeholder for manual input
            "Bureau Score": file_data["bureau_score"],
            "Account Institutions": str(file_data["institutions"]),
            "Account AccountTypes": str(file_data["account_types"]),
            "Account OwnershipTypes": str(file_data["ownership_types"])
        }
        results.append(row)

    # Create DataFrame with defined column order
    columns = [
        "Application Current Status", "Consumer Name", "Gender",
        "PAN", "Address", "DOB", "State", "Mobile", "Age", "Total Income",
        "Bureau Score", "Account Institutions", "Account AccountTypes",
        "Account OwnershipTypes"
    ]
    
    df = pd.DataFrame(results, columns=columns)
    df.to_excel(output_file, index=False)
    print(f"Successfully generated report with {len(results)} rows at {output_file}")

# Example Usage
if __name__ == "__main__":
    JSON_FOLDER = "bureau_reports/"
    OUTPUT_FILE = "hanced_bureau_report.xlsx"
    
    generate_excel_output(JSON_FOLDER, OUTPUT_FILE)