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
# RULE-BASED EDUCATION LEVEL CLASSIFIER
# ==========================================
def classify_education_level(row):
    course = str(row.get('Course', ''))
    dept = str(row.get('Department', ''))
    
    # Rule 1: Bachelor's Degree
    if any(k in course for k in ["Bachelor", "B.Sc.", "B.A.", "B.Ed.", "Undergraduate", "BSc.", "Bachelor", "BSc", "BACHELOR", "Bachelors", "BSc"]) or \
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
    elif any(k in course for k in ["Master", "M.Sc.", "M.A.", "MBA", "M.Phil", "Postgraduate Diploma", "MSc.", "MSc", "Masters","MASTER"]) or \
         any(k in dept for k in ["Master"]):
        return "Master's Degree"
        
    # Rule 5: Doctoral Degree
    elif any(k in course for k in ["Ph.D", "Doctor", "PhD", "DPhil", "PHD.", "Ph.D."]) or \
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
# ROBUST ERP FILE READER WITH ENCODING FALLBACK
# ==========================================
def read_uploaded_dataset(uploaded_file):
    file_name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    
    df_raw = None
    
    if file_name.endswith('.csv'):
        # Try reading CSV with utf-8, fallback to latin1 for legacy encoding
        try:
            df_raw = pd.read_csv(uploaded_file, header=None, encoding='utf-8')
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            df_raw = pd.read_csv(uploaded_file, header=None, encoding='latin1')
    else:
        # Try multiple engines for Excel/ERP files
        read_success = False
        for engine in ['calamine', 'xlrd', 'openpyxl']:
            try:
                uploaded_file.seek(0)
                df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine=engine)
                read_success = True
                break
            except Exception:
                continue
                
        if not read_success:
            # Fallback: Treat as HTML or tab-separated text with latin1 encoding
            try:
                uploaded_file.seek(0)
                df_raw = pd.read_csv(uploaded_file, sep='\t', header=None, encoding='latin1', on_bad_lines='skip')
                read_success = True
            except Exception as e:
                raise ValueError(f"Could not parse file. Ensure 'python-calamine' and 'xlrd' are installed. Details: {e}")

    # Check if the first row contains header names
    first_cell = str(df_raw.iloc[0, 0]).lower()
    has_header = any(keyword in first_cell for keyword in ['serial', 'reg', 'name', 'course'])

    if has_header:
        df_raw.columns = df_raw.iloc[0]
        df = df_raw[1:].reset_index(drop=True)
    else:
        df = df_raw.copy()
        if len(df.columns) == len(ERP_COLUMNS):
            df.columns = ERP_COLUMNS
        else:
            df.columns = [f"Col_{i}" for i in range(len(df.columns))]
            if len(df.columns) >= 7:
                df.rename(columns={
                    df.columns[3]: 'Gender', 
                    df.columns[5]: 'Course', 
                    df.columns[6]: 'Department'
                }, inplace=True)

    return df


# ==========================================
# UI HEADER & FILE UPLOADER
# ==========================================
st.title("📊 StuReport: Student Summary Generator")
st.markdown("Upload your raw ERP export file (**`.xlsx`, `.xls`, or `.csv`**—with or without headers) to automatically generate summary reports.")

uploaded_file = st.file_uploader("Upload ERP Dataset File", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        df = read_uploaded_dataset(uploaded_file)

        # Standardize column names
        df.columns = [str(c).strip() for c in df.columns]
        
        if 'Sex' in df.columns and 'Gender' not in df.columns:
            df.rename(columns={'Sex': 'Gender'}, inplace=True)
        
        if 'Gender' not in df.columns or ('Course' not in df.columns and 'Department' not in df.columns):
            st.error("⚠️ The uploaded file structure could not be automatically mapped. Please ensure it matches your standard ERP export format.")
            st.stop()

        # Preprocessing
        df['Gender'] = df['Gender'].astype(str).str.strip().str.title()
        df['Education_Level'] = df.apply(classify_education_level, axis=1)
        
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
            values=df.columns[1], 
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

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("👆 Please upload your ERP export file above to get started.")