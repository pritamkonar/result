import streamlit as st
import pandas as pd
import io

# --- Page Setup ---
st.set_page_config(page_title="Student Ranker Pro", page_icon="🏆", layout="centered")

# --- Custom CSS ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stButton>button { width: 100%; background-color: #ff4b4b; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Automatic Student Ranker")
st.markdown("Upload your class Excel sheet. This tool will **Fill the Percentage & Rank columns** and **Sort** the list by rank.")

# --- File Uploader ---
uploaded_file = st.file_uploader("📂 Upload Excel File (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        # Load the file
        df = pd.read_excel(uploaded_file)
        st.success("✅ File loaded successfully!")

        st.markdown("---")
        st.subheader("⚙️ Settings")

        col1, col2 = st.columns(2)

        with col1:
            # Auto-detect 'Marks Obtained' or 'Total' column
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            default_idx = 0
            for i, col in enumerate(numeric_cols):
                if any(x in col.lower() for x in ['mark', 'obtain', 'total', 'score']):
                    default_idx = i
                    break
            
            score_col = st.selectbox("Select 'Marks Obtained' Column:", numeric_cols, index=default_idx)

        with col2:
            # Full marks input (Default 100, but user can change to 1000 etc)
            full_marks = st.number_input("Enter Full Marks:", min_value=1, value=1000, step=10)

        # --- Processing Button ---
        if st.button("🚀 Calculate & Organize"):
            with st.spinner('Calculating Ranks and Sorting...'):
                
                # 1. Calculate Percentage
                # This updates the 'Percentage' column if it exists, or creates it if it doesn't
                df['Percentage'] = (df[score_col] / full_marks) * 100
                df['Percentage'] = df['Percentage'].round(2)

                # 2. Calculate Rank
                # method='min' means if two students get Rank 1, the next is Rank 3
                df['Rank'] = df[score_col].rank(ascending=False, method='min')

                # 3. Sort the rows by Rank (1, 2, 3...)
                df_sorted = df.sort_values(by='Rank', ascending=True)

                # 4. Clean up formatting (Optional: remove decimals from Rank)
                df_sorted['Rank'] = df_sorted['Rank'].astype(int)

                # Show Result Preview
                st.write("### ✅ Ranked List Preview")
                st.dataframe(df_sorted.head(10))

                # --- Download Section ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    # We write the dataframe exactly as it is (preserving original columns)
                    df_sorted.to_excel(writer, index=False, sheet_name='Ranked Data')
                
                st.download_button(
                    label="📥 Download Final Spreadsheet",
                    data=buffer.getvalue(),
                    file_name="Final_Ranked_Student_List.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
