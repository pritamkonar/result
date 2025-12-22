import streamlit as st
import pandas as pd
import io

# --- ReportLab Imports for PDF ---
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- Page Setup ---
st.set_page_config(page_title="Student Ranker Pro", page_icon="🏆", layout="centered")

# --- Custom CSS ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stButton>button { width: 100%; background-color: #ff4b4b; color: white; }
    div[data-testid="stExpander"] { background-color: white; border-radius: 10px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Automatic Student Ranker")
st.markdown("Upload your class Excel sheet. This tool will **Sort by Merit**, apply **Tie-Breakers (Old Roll)**, and generate a **New Roll Number**.")

# --- Demo File Generator ---
def get_demo_file():
    # Create sample data matching the expected format
    data = {
        'Roll No': [101, 102, 103, 104, 105],
        'Name': ['Amit Sharma', 'Priya Singh', 'Rahul Verma', 'Sneha Gupta', 'Vikram Rao'],
        'Father\'s Name': ['Rakesh Sharma', 'Karan Singh', 'Manish Verma', 'Suresh Gupta', 'Prakash Rao'],
        'Address': ['New Delhi', 'Mumbai', 'Pune', 'Kolkata', 'Chennai'],
        'Gender': ['Male', 'Female', 'Male', 'Female', 'Male'],
        'Marks Obtained': [850, 920, 780, 890, 910]
    }
    df_demo = pd.DataFrame(data)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_demo.to_excel(writer, index=False, sheet_name='Sheet1')
    return buffer.getvalue()

# --- Sidebar / Top Section for Template ---
col_demo, col_upload = st.columns([1, 2])
with col_demo:
    st.write("### 📝 Need a Format?")
    st.write("Download this sample file, fill it, and upload.")
    st.download_button(
        label="📥 Download Demo Excel",
        data=get_demo_file(),
        file_name="Student_List_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- File Uploader ---
st.markdown("---")
uploaded_file = st.file_uploader("📂 Upload Your Filled Excel File (.xlsx)", type=['xlsx'])

if uploaded_file:
    try:
        # Load the file
        df = pd.read_excel(uploaded_file)
        st.success("✅ File loaded successfully!")

        st.markdown("---")
        st.subheader("⚙️ Processing Settings")

        col1, col2 = st.columns(2)

        with col1:
            # Auto-detect 'Marks Obtained' or 'Total' column
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            default_idx = 0
            for i, col in enumerate(numeric_cols):
                if any(x in col.lower() for x in ['mark', 'obtain', 'total', 'score']):
                    default_idx = i
                    break
            
            # If no numeric columns found, handle gracefully
            if not numeric_cols:
                st.error("❌ No numeric columns found for Marks. Please check your Excel file.")
                st.stop()
                
            score_col = st.selectbox("Select 'Marks Obtained' Column:", numeric_cols, index=default_idx)

        with col2:
            full_marks = st.number_input("Enter Full Marks:", min_value=1, value=1000, step=10)
        
        # --- PDF Header Options ---
        with st.expander("📝 PDF Header Options (Optional)"):
            st.info("Enter details below to appear at the top of the PDF.")
            pdf_school_name = st.text_input("School Name")
            c1, c2 = st.columns(2)
            with c1:
                pdf_class = st.text_input("Class Name")
            with c2:
                pdf_year = st.text_input("Year/Session")

        # --- Processing Button ---
        if st.button("🚀 Calculate & Organize"):
            with st.spinner('Calculating Ranks and Sorting...'):
                
                # 1. Identify and Rename 'Old Roll' BEFORE sorting
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
                if roll_found:
                    df_sorted = df.sort_values(by=[score_col, 'Old Roll'], ascending=[False, True]).reset_index(drop=True)
                else:
                    st.warning("⚠️ 'Roll No' column not found. Sorting strictly by Marks only.")
                    df_sorted = df.sort_values(by=score_col, ascending=False).reset_index(drop=True)

                # 4. Create 'Rank/ New Roll' Column
                df_sorted['Rank/ New Roll'] = range(1, len(df_sorted) + 1)

                # 5. Remove the old 'Rank' column if it exists
                if 'Rank' in df_sorted.columns:
                    df_sorted.drop(columns=['Rank'], inplace=True)

                # 6. Reorder Columns
                cols = list(df_sorted.columns)
                if 'Rank/ New Roll' in cols: cols.remove('Rank/ New Roll')

                if 'Old Roll' in cols:
                    old_roll_index = cols.index('Old Roll')
                    cols.insert(old_roll_index + 1, 'Rank/ New Roll')
                else:
                    cols.insert(0, 'Rank/ New Roll')
                
                df_final = df_sorted[cols]

                # Show Result Preview
                st.write("### ✅ Ranked List Preview")
                st.dataframe(df_final.head(10))

                st.markdown("---")
                st.subheader("📥 Download Results")
                
                d_col1, d_col2 = st.columns(2)

                # --- 1. Excel Download ---
                with d_col1:
                    buffer_excel = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='New Roll List')
                    
                    st.download_button(
                        label="📥 Download Excel (.xlsx)",
                        data=buffer_excel.getvalue(),
                        file_name="Rank_Wise_Student_List.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                # --- 2. PDF Download ---
                with d_col2:
                    buffer_pdf = io.BytesIO()
                    
                    # Setup A4 document
                    doc = SimpleDocTemplate(
                        buffer_pdf, 
                        pagesize=A4, 
                        rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
                    )
                    
                    elements = []
                    styles = getSampleStyleSheet()
                    
                    # -- Custom Styles --
                    title_style = ParagraphStyle(
                        'CustomTitle',
                        parent=styles['Heading1'],
                        fontSize=16,
                        alignment=1, # Center
                        spaceAfter=10,
                        textColor=colors.darkblue
                    )
                    
                    sub_style = ParagraphStyle(
                        'CustomSub',
                        parent=styles['Normal'],
                        fontSize=12,
                        alignment=1, # Center
                        spaceAfter=20
                    )
                    
                    # Reduced body font size
                    cell_style = ParagraphStyle(
                        'CellStyle',
                        parent=styles['BodyText'],
                        fontSize=8, 
                        leading=9,
                        alignment=0 
                    )
                    
                    # Reduced HEADER font size
                    header_style = ParagraphStyle(
                        'HeaderStyle',
                        parent=styles['Normal'],
                        fontSize=7, 
                        leading=8,
                        textColor=colors.white,
                        fontName='Helvetica-Bold',
                        alignment=1 # Center
                    )

                    # -- Add Headers --
                    if pdf_school_name:
                        elements.append(Paragraph(pdf_school_name, title_style))
                    else:
                        elements.append(Paragraph("Student Merit List", title_style))
                        
                    details_text = []
                    if pdf_class: details_text.append(f"Class: {pdf_class}")
                    if pdf_year: details_text.append(f"Session: {pdf_year}")
                    if details_text:
                        elements.append(Paragraph(" | ".join(details_text), sub_style))

                    # -- Prepare Table Data --
                    headers = [Paragraph(str(col), header_style) for col in df_final.columns]
                    data = [headers]

                    for index, row in df_final.iterrows():
                        row_data = []
                        for item in row:
                            if pd.isna(item):
                                text_val = "-"
                            elif isinstance(item, float):
                                text_val = f"{item:.2f}"
                            else:
                                text_val = str(item)
                            
                            row_data.append(Paragraph(text_val, cell_style))
                        data.append(row_data)

                    # -- INTELLIGENT COLUMN SIZING --
                    # 1. Calculate max character count for each column
                    max_chars_per_col = []
                    for col in df_final.columns:
                        max_len_data = df_final[col].astype(str).map(len).max()
                        len_header = len(str(col))
                        max_chars_per_col.append(max(max_len_data, len_header))

                    # 2. Calculate total characters
                    total_chars = sum(max_chars_per_col)

                    # 3. Distribute A4 Width
                    usable_width = 555
                    col_widths = []
                    
                    for i, col_name in enumerate(df_final.columns):
                        chars = max_chars_per_col[i]
                        width = (chars / total_chars) * usable_width
                        
                        # --- MANUAL OVERRIDES ---
                        c_name_lower = str(col_name).lower()

                        if 'gender' in c_name_lower or 'sex' in c_name_lower:
                            width = max(width, 60)
                        
                        if 'percent' in c_name_lower:
                            width = max(width, 55)

                        if width < 30: width = 30
                        
                        col_widths.append(width)

                    # -- Create Table --
                    t = Table(data, colWidths=col_widths)
                    
                    t.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
                        ('LEFTPADDING', (0, 0), (-1, -1), 4),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                        ('TOPPADDING', (0, 0), (-1, -1), 4),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                    ]))

                    elements.append(t)
                    doc.build(elements)
                    
                    st.download_button(
                        label="📥 Download PDF (A4)",
                        data=buffer_pdf.getvalue(),
                        file_name="Merit_List.pdf",
                        mime="application/pdf"
                    )

    except Exception as e:
        st.error(f"An error occurred: {e}")
