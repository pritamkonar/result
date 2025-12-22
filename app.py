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
st.markdown("Upload your class Excel sheet. This tool will **Sort by Merit**, rename 'Roll No' to 'Old Roll', and generate a **New Roll Number**.")

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
            full_marks = st.number_input("Enter Full Marks:", min_value=1, value=1000, step=10)

        # --- Processing Button ---
        if st.button("🚀 Calculate & Organize"):
            with st.spinner('Calculating Ranks and Sorting...'):
                
                # 1. Calculate Percentage
                df['Percentage'] = (df[score_col] / full_marks) * 100
                df['Percentage'] = df['Percentage'].round(2)

                # 2. Sort by Marks (Highest first)
                # We sort FIRST so we can assign sequential New Roll numbers (1, 2, 3...)
                df_sorted = df.sort_values(by=score_col, ascending=False).reset_index(drop=True)

                # 3. Create 'Rank/ New Roll' Column
                # Assign 1 to N based on the sorted order
                df_sorted['Rank/ New Roll'] = range(1, len(df_sorted) + 1)

                # 4. Rename 'Roll No.' to 'Old Roll'
                # We look for columns that look like "Roll No"
                for col in df_sorted.columns:
                    if 'roll' in col.lower() and 'new' not in col.lower():
                        df_sorted.rename(columns={col: 'Old Roll'}, inplace=True)
                        break
                
                # 5. Remove the old 'Rank' column if it exists in input (to avoid confusion)
                if 'Rank' in df_sorted.columns:
                    df_sorted.drop(columns=['Rank'], inplace=True)

                # 6. Reorder Columns: Put 'Rank/ New Roll' right after 'Old Roll'
                cols = list(df_sorted.columns)
                
                # Remove 'Rank/ New Roll' from the end list temporarily
                if 'Rank/ New Roll' in cols:
                    cols.remove('Rank/ New Roll')

                # Find 'Old Roll' and insert 'New Roll' after it
                if 'Old Roll' in cols:
                    old_roll_index = cols.index('Old Roll')
                    cols.insert(old_roll_index + 1, 'Rank/ New Roll')
                else:
                    # If 'Old Roll' wasn't found, put New Roll at the very start
                    cols.insert(0, 'Rank/ New Roll')
                
                # Apply the new column order
                df_final = df_sorted[cols]

                # Show Result Preview
                st.write("### ✅ Ranked List Preview")
                st.dataframe(df_final.head(10))

                # --- Download Section ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='New Roll List')
                
                st.download_button(
                    label="📥 Download Final Spreadsheet",
                    data=buffer.getvalue(),
                    file_name="Rank_Wise_Student_List.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
