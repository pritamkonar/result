import streamlit as st
import pandas as pd
import io

# --- Page Setup ---
st.set_page_config(page_title="Student Ranker Pro", page_icon="🏆", layout="centered")

# --- Custom CSS for a professional look ---
st.markdown("""
    <style>
    .main {
        background-color: #f0f2f6;
    }
    .stButton>button {
        width: 100%;
        background-color: #ff4b4b;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Automatic Student Ranker")
st.markdown("Upload your class Excel sheet to automatically calculate **Percentage**, **Rank**, and **Sort** the list.")

# --- File Uploader ---
uploaded_file = st.file_uploader("📂 Upload Excel File (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ File loaded successfully!")

        st.markdown("---")
        st.subheader("⚙️ Settings")

        col1, col2 = st.columns(2)

        with col1:
            # Attempt to auto-detect the Total Score column
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            default_idx = 0
            for i, col in enumerate(numeric_cols):
                if any(x in col.lower() for x in ['total', 'score', 'mark', 'obtained']):
                    default_idx = i
                    break
            
            score_col = st.selectbox("Select 'Total Score' Column:", numeric_cols, index=default_idx)

        with col2:
            full_marks = st.number_input("Enter Full Marks:", min_value=1, value=100, step=10)

        # --- Processing Button ---
        if st.button("🚀 Generate Ranked List"):
            with st.spinner('Processing data...'):
                
                # 1. Calculate Percentage
                df['Percentage'] = (df[score_col] / full_marks) * 100
                df['Percentage'] = df['Percentage'].round(2)

                # 2. Calculate Rank
                # method='min' handles ties (e.g. if two people are 1st, next is 3rd)
                df['Rank'] = df[score_col].rank(ascending=False, method='min')

                # 3. Sort by Rank
                df_sorted = df.sort_values(by='Rank', ascending=True)

                # 4. Reorder columns (Rank first, then Name, then others)
                # We try to put Name second if it exists
                cols = list(df_sorted.columns)
                cols.remove('Rank')
                
                # Check for name column to put it second
                name_col = next((c for c in cols if 'name' in c.lower()), None)
                if name_col:
                    cols.remove(name_col)
                    final_order = ['Rank', name_col] + cols
                else:
                    final_order = ['Rank'] + cols
                
                df_sorted = df_sorted[final_order]

                # Show Result
                st.write("### ✅ Preview of Sorted List")
                st.dataframe(df_sorted.head(10))

                # --- Download Section ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_sorted.to_excel(writer, index=False, sheet_name='Rank List')
                
                st.download_button(
                    label="📥 Download Sorted Excel File",
                    data=buffer.getvalue(),
                    file_name="Rank_Wise_Student_List.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")