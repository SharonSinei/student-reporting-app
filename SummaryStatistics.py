import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="StuReport",
    page_icon="📊",
    layout="wide"
)

# ==========================================
# STANDARD ERP COLUMN SCHEMA
# ==========================================
ERP_COLUMNS = [
    'Serial. No', 'Reg. No', 'Name', 'Gender', 'GSSP/PSSP', 
    'Course', 'Department', 'School', 'Semester', 'Residency', 
    'Reporting Date', 'Time', 'Year', 'Semester(#1)', 'Email', 'Phone Number', 'N/A'
]

# ==========================================
# ENHANCED RULE-BASED EDUCATION LEVEL CLASSIFIER
# ==========================================
def classify_education_level(row):
    """
    Categorizes course names and departments into standardized education levels 
    using an expanded rule set to catch all variations.
    """
    course = str(row.get('Course', ''))
    dept = str(row.get('Department', ''))
    
    # Rule 1: Bachelor's Degree
    if any(k in course for k in ["Bachelor", "B.Sc.", "B.A.", "B.Ed.", "Undergraduate", "BSc.", "BACHELOR", "BSc in"]) or \
       any(k in dept for k in ["Bachelor"]):
        return "Bachelor's Degree"
        
    # Rule 2: Diploma
    elif any(k in course for k in ["Diploma", "Dip.", "DIPLOMA"]) or \
         any(k in dept for k in ["Diploma"]):
        return "Diploma"
        
    # Rule 3: Certificate
    elif any(k in course for k in ["Certificate", "Cert."]) or \
         any(k in dept for k in ["Certificate"]):
        return "Certificate"
        
    # Rule 4: Master's Degree
    elif any(k in course for k in ["Master", "M.Sc.", "M.A.", "MBA", "M.Phil", "Postgraduate Diploma", "MSc.", "MSc in"]) or \
         any(k in dept for k in ["Master"]):
        return "Master's Degree"
        
    # Rule 5: Doctoral Degree
    elif any(k in course for k in ["Ph.D", "Doctor", "PhD", "DPhil", "PHD."]) or \
         any(k in dept for k in ["Doctoral", "PhD"]):
        return "Doctoral Degree"
        
    # Default outcome
    else:
        return "Other / Unclassified"


# ==========================================
# EXCEL CONVERSION HELPER FUNCTION
# ==========================================
def convert_df_to_excel(df_summary, total_count):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Summary Report', index=True)
        
        df_total = pd.DataFrame({'Metric': ['Total Students'], 'Count': [total_count]})
        df_total.to_excel(writer, sheet_name='Total Count', index=False)
        
    buffer.seek(0)
    return buffer


# ==========================================
# UI HEADER & FILE UPLOADER
# ==========================================
st.title("📊 StuReport: Student Summary Generator")
st.markdown("Upload your raw ERP export file (**`.xlsx`, `.xls`, or `.csv`**—with or without headers) to automatically generate summary reports.")

uploaded_file = st.file_uploader("Upload ERP Dataset File", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        file_name = uploaded_file.name.lower()
        
        # Check if file has headers or is headerless from ERP export
        if file_name.endswith('.csv'):
            df_temp = pd.read_csv(uploaded_file, nrows=1, header=None)
        else:
            df_temp = pd.read_excel(uploaded_file, sheet_name=0, nrows=1, header=None)
            
        first_cell = str(df_temp.iloc[0, 0])
        has_header = any(keyword in first_cell.lower() for keyword in ['serial', 'reg', 'name', 'course'])
        
        # Read file accordingly
        if file_name.endswith('.csv'):
            if has_header:
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_csv(uploaded_file, header=None)
                if len(df.columns) == len(ERP_COLUMNS):
                    df.columns = ERP_COLUMNS
        else:
            if has_header:
                df = pd.read_excel(uploaded_file, sheet_name=0)
            else:
                df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
                if len(df.columns) == len(ERP_COLUMNS):
                    df.columns = ERP_COLUMNS
                else:
                    df.columns = [f"Col_{i}" for i in range(len(df.columns))]
                    if len(df.columns) >= 6:
                        df.rename(columns={df.columns[3]: 'Gender', df.columns[5]: 'Course', df.columns[6]: 'Department'}, inplace=True)

        # Standardize column names
        df.columns = [str(c).strip() for c in df.columns]
        
        # Normalize column name if ERP exported it as "Sex" instead of "Gender"
        if 'Sex' in df.columns and 'Gender' not in df.columns:
            df.rename(columns={'Sex': 'Gender'}, inplace=True)

        # Ensure required columns exist
        if 'Gender' not in df.columns or ('Course' not in df.columns and 'Department' not in df.columns):
            st.error("⚠️ The uploaded file structure could not be automatically mapped. Please ensure it matches your standard ERP export format.")
            st.stop()

        # Preprocessing
        df['Gender'] = df['Gender'].astype(str).str.strip().str.title()
        df['Education_Level'] = df.apply(classify_education_level, axis=1)
        
        # Check unclassified rows
        unclassified_count = (df['Education_Level'] == 'Other / Unclassified').sum()
        if unclassified_count > 0:
            st.warning(f"⚠️ Note: {unclassified_count} rows were classified as 'Other / Unclassified'.")

        # ==========================================
        # GENERATE SUMMARY TABLE
        # ==========================================
        summary_table = pd.pivot_table(
            df,
            index='Education_Level',
            columns='Gender',
            values=df.columns[1], # Uses Reg. No column for counting
            aggfunc='count',
            fill_value=0
        )
        
        for col in ['Female', 'Male']:
            if col not in summary_table.columns:
                summary_table[col] = 0
                
        summary_table = summary_table[['Female', 'Male']]
        summary_table.loc['Total'] = summary_table.sum()
        
        total_students_count = len(df)
        
        st.success("✅ ERP dataset successfully processed and headers automatically mapped!")
        
        # ==========================================
        # DISPLAY SUMMARY TABLE
        # ==========================================
        st.subheader("📋 Summary Table (Academic Level vs Gender)")
        st.dataframe(summary_table, use_container_width=True)
        
        st.metric(label="Total Students Reported", value=f"{total_students_count:,}")
        
        # ==========================================
        # DOWNLOAD BUTTONS
        # ==========================================
        col_dl1, col_dl2 = st.columns(2)
        
        excel_data = convert_df_to_excel(summary_table, total_students_count)
        col_dl1.download_button(
            label="📥 Download Summary as Excel (.xlsx)",
            data=excel_data,
            file_name="student_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        csv_data = summary_table.to_csv().encode('utf-8')
        col_dl2.download_button(
            label="📄 Download Summary as CSV (.csv)",
            data=csv_data,
            file_name="student_summary_report.csv",
            mime="text/csv"
        )
        
        # ==========================================
        # VISUALIZATIONS
        # ==========================================
        st.markdown("---")
        st.subheader("📊 Visual Breakdown")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### Gender Breakdown per Course Level")
            fig, ax = plt.subplots(figsize=(6, 4))
            summary_table.drop('Total', errors='ignore')[['Female', 'Male']].plot(kind='bar', ax=ax, color=['#e056fd', '#0984e3'])
            ax.set_xlabel("Course Level")
            ax.set_ylabel("Number of Students")
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig)
            
        with col2:
            st.markdown("##### Total Students by Course Level")
            fig, ax = plt.subplots(figsize=(6, 4))
            summary_table.drop('Total', errors='ignore').sum(axis=1).plot(kind='barh', ax=ax, color='#2b5c8f')
            ax.set_xlabel("Total Count")
            ax.set_ylabel("Course Level")
            plt.tight_layout()
            st.pyplot(fig)

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info(" Please upload your ERP export file above to get started.")