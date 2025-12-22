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
st.markdown("Upload your class Excel sheet. This tool will **Sort by Merit**, apply **Tie-Breakers (Old Roll)**, and generate a **New Roll Number**.")

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
                
                # 1. Identify and Rename 'Old Roll' BEFORE sorting
                # We need this column to exist so we can use it as a tie-breaker
                roll_found = False
                for col in df.columns:
                    if 'roll' in col.lower() and 'new' not in col.lower():
                        df.rename(columns={col: 'Old Roll'}, inplace=True)
                        roll_found = True
                        break
                
                # 2. Calculate Percentage
                df['Percentage'] = (df[score_col] / full_marks) * 100
                df['Percentage'] = df['Percentage'].round(2)

                # 3. Sort by Marks (Highest) THEN by Old Roll (Lowest)
                # This handles the tie-breaker: If marks are same, lower Old Roll gets top position
                if roll_found:
                    df_sorted = df.sort_values(by=[score_col, 'Old Roll'], ascending=[False, True]).reset_index(drop=True)
                else:
                    st.warning("⚠️ 'Roll No' column not found. Sorting strictly by Marks only.")
                    df_sorted = df.sort_values(by=score_col, ascending=False).reset_index(drop=True)

                # 4. Create 'Rank/ New Roll' Column
                # Since the list is already sorted with the tie-breaker, we just number them 1 to N
                df_sorted['Rank/ New Roll'] = range(1, len(df_sorted) + 1)

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
                st.write("### ✅ Ranked List Preview (With Tie-Breaker)")
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
