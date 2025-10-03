
import streamlit as st
import pandas as pd
import numpy as np
from rapidfuzz import fuzz
from shapely.geometry import Point
import geopandas as gpd
# --- Parameters ---
BUFFER_DISTANCE_KM = 2
CAPACITY_TOLERANCE = 0.1   # 10%
TEXT_SIMILARITY_THRESHOLD = 70  # 0-100 scale for rapidfuzz
# --- Helper: safely convert columns to numeric ---
def safe_to_numeric(df, cols):
    for col in cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df
st.set_page_config(page_title="ECR-REPD Matcher", layout="wide")
st.title("⚡ ECR–REPD Matching Tool")
# --- Step 1: Ask file format ---
file_option = st.radio(
    "Do you have one file with two sheets or two separate files?",
    ("One file, two sheets", "Two separate files")
)
repd_df = None
ecr_df = None
# --- Step 2: File Upload ---
if file_option == "One file, two sheets":
    uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])
    if uploaded:
        try:
            repd_df = pd.read_excel(uploaded, sheet_name="REPD")
            ecr_df = pd.read_excel(uploaded, sheet_name="ECR")
            st.success("✅ Loaded REPD and ECR sheets.")
        except Exception as e:
            st.error(f"Error reading file: {e}")
else:
    repd_file = st.file_uploader("Upload REPD Excel", type=["xlsx"], key="repd")
    ecr_file = st.file_uploader("Upload ECR Excel", type=["xlsx"], key="ecr")
    if repd_file and ecr_file:
        try:
            repd_df = pd.read_excel(repd_file)
            ecr_df = pd.read_excel(ecr_file)
            st.success("✅ Loaded REPD and ECR files.")
        except Exception as e:
            st.error(f"Error reading files: {e}")
# --- Step 3: Input REPD ID range ---
if repd_df is not None and ecr_df is not None:
    start_id = st.text_input("Enter start REPD_ID")
    end_id = st.text_input("Enter end REPD_ID")
    if start_id and end_id and st.button("🔍 Run Matching"):
        try:
            start_id = int(start_id)
            end_id = int(end_id)
        except ValueError:
            st.error("REPD_ID range must be numeric")
            st.stop()
        # --- Prepare REPD subset ---
        repd = repd_df.copy()
        repd["REPD_ID"] = pd.to_numeric(repd["REPD_ID"], errors="coerce")  # force numeric
        repd = repd[(repd["REPD_ID"] >= start_id) & (repd["REPD_ID"] <= end_id)]
        repd = safe_to_numeric(repd, ["X-coordinate", "Y-coordinate", "Installed Capacity (MWelec)"])
        
        # Insert required columns at the start
        for col in ["Matched_ECR_ID", "Matching Reason", "Matched Details REPD", "Matched Details ECR"]:
            repd.insert(0, col, "NF")
        # --- Prepare ECR ---
        ecr = ecr_df[ecr_df["Energy_Source_1"].str.lower() == "solar"].copy()
        ecr = safe_to_numeric(ecr, ["Location__X_coordinate___Eastin", "Location__y_coordinate___Northi",
                                    "Accepted_to_Connect_Registered_"])
        # Convert coordinates to geodataframes
        try:
            repd_gdf = gpd.GeoDataFrame(
                repd,
                geometry=gpd.points_from_xy(repd["X-coordinate"], repd["Y-coordinate"], crs="EPSG:27700")
            )
            ecr_gdf = gpd.GeoDataFrame(
                ecr,
                geometry=gpd.points_from_xy(ecr["Location__X_coordinate___Eastin"],
                                            ecr["Location__y_coordinate___Northi"],
                                            crs="EPSG:27700")
            )
        except Exception as e:
            st.error(f"Coordinate error: {e}")
            st.stop()
        # --- Matching logic ---
        results = repd.copy()
        for idx, repd_row in repd_gdf.iterrows():
            repd_geom = repd_row.geometry
            if repd_geom is None or repd_row["X-coordinate"] is None or repd_row["Y-coordinate"] is None:
                continue
            # Spatial buffer
            buffer = repd_geom.buffer(BUFFER_DISTANCE_KM * 1000)  # meters
            ecr_candidates = ecr_gdf[ecr_gdf.intersects(buffer)]
            if ecr_candidates.empty:
                continue
            best_match = None
            best_score = -1
            reasons = []
            repd_details = []
            ecr_details = []
            for _, ecr_row in ecr_candidates.iterrows():
                match_reasons = ["spatial"]
                match_repd = []
                match_ecr = []
                # Capacity match
                repd_cap = repd_row.get("Installed Capacity (MWelec)")
                ecr_cap = ecr_row.get("Accepted_to_Connect_Registered_")
                if pd.notna(repd_cap) and pd.notna(ecr_cap):
                    if abs(repd_cap - ecr_cap) <= CAPACITY_TOLERANCE * repd_cap:
                        match_reasons.append("capacity")
                        match_repd.append(f"capacity: {repd_cap}")
                        match_ecr.append(f"capacity: {ecr_cap}")
                # Text Group A
                repd_text_a = str(repd_row.get("Operator (or Applicant)", "")) + " " + str(repd_row.get("Site Name", ""))
                ecr_text_a = str(ecr_row.get("Customer_Name", "")) + " " + str(ecr_row.get("Customer_Site", ""))
                if fuzz.token_sort_ratio(repd_text_a, ecr_text_a) >= TEXT_SIMILARITY_THRESHOLD:
                    match_reasons.append("text(GrpA)")
                    match_repd.append(f"textA: {repd_text_a}")
                    match_ecr.append(f"textA: {ecr_text_a}")
                # Text Group B
                repd_text_b = str(repd_row.get("Site Name", "")) + " " + str(repd_row.get("Address", ""))
                ecr_text_b = str(ecr_row.get("Customer_Site", "")) + " " + str(ecr_row.get("Address_Line_1", ""))
                if fuzz.token_sort_ratio(repd_text_b, ecr_text_b) >= TEXT_SIMILARITY_THRESHOLD:
                    match_reasons.append("text(GrpB)")
                    match_repd.append(f"textB: {repd_text_b}")
                    match_ecr.append(f"textB: {ecr_text_b}")
                # Postcode match
                repd_pc = str(repd_row.get("Post Code", "")).replace(" ", "").lower()
                ecr_pc = str(ecr_row.get("Postcode", "")).replace(" ", "").lower()
                if repd_pc and repd_pc == ecr_pc:
                    match_reasons.append("postcode")
                    match_repd.append(f"postcode: {repd_pc}")
                    match_ecr.append(f"postcode: {ecr_pc}")
                score = len(match_reasons)
                if score > best_score:
                    best_score = score
                    best_match = ecr_row
                    reasons = match_reasons
                    repd_details = match_repd
                    ecr_details = match_ecr
            # Save best match
            if best_match is not None:
                results.at[idx, "Matched_ECR_ID"] = best_match["ECR_ID"]
                results.at[idx, "Matching Reason"] = ", ".join(reasons)
                results.at[idx, "Matched Details REPD"] = "; ".join(repd_details)
                results.at[idx, "Matched Details ECR"] = "; ".join(ecr_details)
        # --- Show results ---
        st.subheader("Results")
        st.dataframe(results)
        # Save to Excel
        out_file = f"REPD_ECR_match_projectID_{start_id}_to_{end_id}.xlsx"
        results.to_excel(out_file, index=False)
        with open(out_file, "rb") as f:
            st.download_button("📥 Download Results as Excel", f, file_name=out_file)
        # Reset button
        if st.button("🔄 Clear and Start Again"):
            st.experimental_rerun()
