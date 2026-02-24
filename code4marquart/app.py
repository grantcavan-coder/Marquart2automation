import streamlit as st
import pandas as pd
import os
import tempfile
from process_cars import process_file

st.set_page_config(page_title="Car Ranker", page_icon="🚗", layout="centered")

st.title("🚗 Car Competitive Ranker")
st.markdown("Upload your raw car list CSV and get back a ranked Excel file — automatically!")

st.markdown("---")

uploaded_file = st.file_uploader("📂 Upload your CSV file", type=["csv"])

if uploaded_file:
    st.success(f"Uploaded: **{uploaded_file.name}**")

    if st.button("⚡ Generate Ranked Excel", type="primary"):
        with st.spinner("Processing..."):
            try:
                # Save upload to temp file
                with tempfile.NamedTemporaryFile(delete=False, suffix='.csv') as tmp_in:
                    tmp_in.write(uploaded_file.getvalue())
                    input_path = tmp_in.name

                output_path = input_path.replace('.csv', '_ranked.xlsx')

                df = process_file(input_path, output_path)

                st.success(f"✅ Done! Ranked **{len(df)} cars**.")

                # Preview top 10
                st.markdown("### 🏆 Top 10 Results Preview")
                preview_cols = ['Rank', 'Make/Model', 'Price', 'Odometer', 'Age',
                                'Relative Price', 'Relative Age', 'Score']
                available = [c for c in preview_cols if c in df.columns]
                st.dataframe(df[available].head(10), use_container_width=True)

                # Download button
                with open(output_path, 'rb') as f:
                    st.download_button(
                        label="📥 Download Ranked Excel File",
                        data=f,
                        file_name="ranked_cars.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

                # Cleanup
                os.unlink(input_path)
                os.unlink(output_path)

            except Exception as e:
                st.error(f"Something went wrong: {e}")
                st.info("Make sure your CSV has **Price**, **Odometer**, and **Age** columns.")

st.markdown("---")
st.caption("Scores use PERCENTRANK × -1. Higher score = better deal (cheaper + newer).")
