# Python script for automating payout grid lookup with Streamlit
import pandas as pd
import os
import streamlit as st

def determine_column(policy_type, fuel_type=None, ncb=None):
    if policy_type.lower() == "new":
        return "Pvt Car New 1+3"
    elif policy_type.lower() == "used":
        return "Pvt Car (Used Car**)"
    elif policy_type.lower() == "tp":
        if fuel_type.lower() == "petrol":
            return "Pvt car AOTP- Petrol"
        elif fuel_type.lower() == "diesel":
            return "Pvt car AOTP- Diesel"
    elif policy_type.lower() == "saod":
        if ncb and ncb.lower() == "yes":
            return "SAOD-NCB"
        elif ncb and ncb.lower() == "no":
            return "Pvt Car-0 NCB (NON NCB)"
    elif policy_type.lower() == "comp":
        if fuel_type.lower() == "petrol" and ncb.lower() == "yes":
            return "Pvt Car Petrol & CNG- 1+1 (NCB Cases)"
        elif fuel_type.lower() == "petrol" and ncb.lower() == "no":
            return "Pvt Car-0 NCB (NON NCB)"
        elif fuel_type.lower() in ["diesel", "ev"] and ncb.lower() == "yes":
            return "Pvt Car Diesel & EV - 1+1 (NCB Cases)"
        elif fuel_type.lower() in ["diesel", "ev"] and ncb.lower() == "no":
            return "Pvt Car-0 NCB (NON NCB)"
    return None

def fetch_payout(rto_code, policy_type, fuel_type=None, ncb=None, sheet1_data=None, grid_data=None):
    # Find the RTO location for the RTO code
    rto_mapping = sheet1_data[sheet1_data["RTO_CODE"].str.strip().str.casefold() == rto_code.strip().casefold()]
    if rto_mapping.empty:
        return f"Error: RTO code {rto_code} not found."
    
    rto_location = rto_mapping["GRID MAPPING"].values[0].strip()  # Get the location
    
    # Determine the column to search
    column = determine_column(policy_type, fuel_type, ncb)
    if not column or column not in grid_data.columns:
        return f"Error: Unable to determine the payout column for the inputs."

    # Extract payout from the grid using case-insensitive matching for RTO Location
    payout_row = grid_data[grid_data["RTO Location"].str.strip().str.casefold() == rto_location.casefold()]
    if payout_row.empty:
        return f"Error: No payout found for RTO {rto_code} and location {rto_location}."
    
    payout = payout_row[column].values[0]
    # Convert to percentage if the value is numeric
    if isinstance(payout, (int, float)):
        payout = round(payout * 100, 2)
    return f"Payout: {payout}%"

def main():
    st.title("Payout Grid Lookup Tool")

    # Load data from Excel using a relative path
    file_path = os.path.join(os.path.dirname(__file__), "ICICI_GRID.xlsx")  # Relative file path
    sheet1_data = pd.read_excel(file_path, sheet_name="Sheet1")
    grid_data = pd.read_excel(file_path, sheet_name="4W  Grid")

    # Preprocess the 4W Grid data to correct headers
    original_headers = ['RTO CATEGORY', 'RTO Zone', 'RTO State', 'RTO Location']
    policy_headers = grid_data.iloc[1, 4:].values  # Extract policy-specific headers
    correct_policy_headers = [
        "Pvt Car New 1+3",
        "Pvt Car Petrol & CNG- 1+1 (NCB Cases)",
        "Pvt Car Diesel & EV - 1+1 (NCB Cases)",
        "SAOD-NCB",
        "Pvt Car-0 NCB (NON NCB)",
        "Pvt Car (Used Car**)",
        "Pvt car AOTP- Petrol",
        "Pvt car AOTP- Diesel"
    ]
    final_headers = list(original_headers) + correct_policy_headers
    grid_data.columns = final_headers  # Apply corrected headers
    grid_data = grid_data[2:].reset_index(drop=True)  # Remove unnecessary rows

    # Input fields for user input
    rto_code = st.text_input("Enter RTO Code (e.g., MH04):").strip()
    policy_type = st.selectbox("Select Policy Type:", ["new", "used", "tp", "saod", "comp"])
    fuel_type = None
    ncb = None

    if policy_type in ["comp", "tp"]:
        fuel_type = st.selectbox("Select Fuel Type:", ["petrol", "diesel", "cng", "ev"])

    if policy_type in ["comp", "saod"]:
        ncb = st.selectbox("Select NCB:", ["yes", "no"])

    # Calculate and display payout
    if st.button("Calculate Payout"):
        result = fetch_payout(
            rto_code=rto_code,
            policy_type=policy_type,
            fuel_type=fuel_type,
            ncb=ncb,
            sheet1_data=sheet1_data,
            grid_data=grid_data
        )
        st.write(f"Result: {result}")

if __name__ == "__main__":
    main()
