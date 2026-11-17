import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

# ==========================================
# PAGE CONFIGURATION
# ==========================================
st.set_page_config(
    page_title="Student Summary Statistics Generator",
    page_icon="📊",
    layout="wide"
)

# ==========================================
# RULE-BASED EDUCATION LEVEL CLASSIFIER
# ==========================================
def classify_education_level(row):
    """
    Categorizes course names and departments into standardized education levels 
    using the provided rule set.
    """
    course = str(row.get('Course', ''))
    dept = str(row.get('Department', ''))

    # Rule 1: Bachelor's Degree
    if any(k in course for k in ["Bachelor", "B.Sc.", "B.A.", "B.Ed.", "Undergraduate", "BSc.", "Bachelor", "Bachelors", "BSc", "BSc.", "BACHELOR"]) or \
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
    elif any(k in course for k in ["Master", "M.Sc.", "M.A.", "MBA", "M.Phil", "Postgraduate Diploma", "MSc.", "Masters"]) or \
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

        # Write total count on a secondary sheet
        df_total = pd.DataFrame({'Metric': ['Total Students'], 'Count': [total_count]})
        df_total.to_excel(writer, sheet_name='Total Count', index=False)

    buffer.seek(0)
    return buffer


# ==========================================
# UI HEADER & FILE UPLOADER
# ==========================================
st.title("Student Summary Statistics Generator")
st.markdown("Upload your raw student reported dataset Excel file to automatically generate the summary breakdown by Course Level and Gender.")

uploaded_file = st.file_uploader("Upload Reported Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read uploaded file
        df = pd.read_excel(uploaded_file, sheet_name=0)

        # Preprocessing
        df['Sex'] = df['Sex'].astype(str).str.strip().str.title()
        df['Education_Level'] = df.apply(classify_education_level, axis=1)

        # ==========================================
        # GENERATE SUMMARY TABLE (Matching Reference Image)
        # ==========================================
        summary_table = pd.pivot_table(
            df,
            index='Education_Level',
            columns='Sex',
            values='Reg. No',
            aggfunc='count',
            fill_value=0
        )

        # Ensure Female and Male columns exist
        for col in ['Female', 'Male']:
            if col not in summary_table.columns:
                summary_table[col] = 0

        summary_table = summary_table[['Female', 'Male']]

        # Compute Total row and column
        summary_table.loc['Total'] = summary_table.sum()

        total_students_count = len(df)

        st.success("✅ Dataset successfully processed!")

        # ==========================================
        # DISPLAY SUMMARY TABLE
        # ==========================================
        st.subheader("📋 Summary Table")

        # Display as formatted dataframe
        st.dataframe(summary_table, use_container_width=True)

        # Display Total Students callout
        st.metric(label="Total Students", value=f"{total_students_count:,}")

        # ==========================================
        # DOWNLOAD BUTTONS (EXCEL & CSV)
        # ==========================================
        col_dl1, col_dl2 = st.columns(2)

        # Excel Download Button
        excel_data = convert_df_to_excel(summary_table, total_students_count)
        col_dl1.download_button(
            label="📥 Download Summary as Excel (.xlsx)",
            data=excel_data,
            file_name="student_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # CSV Download Button
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
        xy = (e == 'Sex')
        st.error(e)
        
else:
    st.info("👆 Please upload the excel file above to get started.")
