import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="TB Processor", layout="wide")
st.title("Trial Balance Entity Processor")

# Step 1: Upload TB Files
uploaded_files = st.file_uploader(
    "Upload multiple Trial Balance Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# Upload TP Category File
tp_category_file = st.file_uploader("Upload TP Category Excel file", type=["xlsx"])

months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
]
years = ['25', '26', '27', '28', '29', '30']

selected_month = st.selectbox("Select Month", months)
selected_year = st.selectbox("Select Year", years)
selected_month_year = f"{selected_month}-{selected_year}"

def deduplicate_columns(columns):
    seen = {}
    new_columns = []
    for col in columns:
        if col not in seen:
            seen[col] = 0
            new_columns.append(col)
        else:
            seen[col] += 1
            new_columns.append(f"{col}.{seen[col]}")
    return new_columns

# Start processing if both files uploaded
if uploaded_files and tp_category_file and st.button("Process and Download Final Excel"):

    # Load TP Category once
    tp_df = pd.read_excel(tp_category_file)
    tp_df.columns = tp_df.columns.str.strip()
    if 'GL code' not in tp_df.columns:
        st.error("The TP Category file must contain a column named 'GL code'.")
        st.stop()

    entity_dfs = {}

    for file in uploaded_files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip()
        df.columns = deduplicate_columns(df.columns)

        entity_col = [col for col in df.columns if col.lower() == 'entity']
        if not entity_col:
            st.error(f"'Entity' column not found in file {file.name}")
            continue

        entity_col = entity_col[0]
        df = df.dropna(subset=[entity_col])

        try:
            df[entity_col] = df[entity_col].astype(int)
        except ValueError:
            df[entity_col] = df[entity_col].astype(str)

        df = df.sort_values(by=entity_col)

        for entity_code in df[entity_col].unique():
            entity_df = df[df[entity_col] == entity_code]
            if entity_code in entity_dfs:
                entity_dfs[entity_code] = pd.concat([entity_dfs[entity_code], entity_df], ignore_index=True)
            else:
                entity_dfs[entity_code] = entity_df

    # Map entity-specific GL column names
    entity_gl_column_map = {
        8223: 'Acc',
        8226: 'AccountNo',
        8297: 'Account Code',
        8224: 'G/L Account',
        8225: 'G/L Account',
        8229: 'G/L Account',
        8235: 'G/L Account',
        8236: 'Acc',
        
    }

    def classify_bs_pl(col, x):
        x = str(x)
        if x.startswith('M'):
            return 'Migration'
        elif x.startswith(('1', 'L1', '2', 'L2')):
            return 'BS'
        elif x.startswith(('3', 'L3', '4', 'L4', '5', '6', '7', '8', '9')):
            return 'PL'
        return None

    def add_monthly_movement(df, debit_col, credit_col):
        if debit_col in df.columns and credit_col in df.columns:
            df['Monthly Movement'] = df[debit_col].fillna(0) + df[credit_col].fillna(0)
        else:
            df['Monthly Movement'] = None
        return df

    processed_dfs = []

    for entity, df in entity_dfs.items():
        df['Month'] = selected_month_year

        # ------------------- Determine GL Column ---------------------
        gl_col = entity_gl_column_map.get(int(entity), None)

        if gl_col and gl_col in df.columns:
            df[gl_col] = df[gl_col].astype(str).str.strip().str.replace(r'^.*/', '', regex=True)
            df['BS/PL'] = df[gl_col].apply(lambda x: classify_bs_pl(gl_col, x))

            # ------------------- Merge TP Category ---------------------
            tp_df_copy = tp_df.copy()
            tp_df_copy['GL code'] = tp_df_copy['GL code'].astype(str).str.strip()

            df = df.merge(tp_df_copy, how='left', left_on=gl_col, right_on='GL code')
        else:
            df['BS/PL'] = None
            st.warning(f"GL column for entity {entity} not found or missing in data.")

        # ------------------- Monthly Movement ---------------------
        if ('Debit Balance in Company Code Currency' in df.columns and
                'Credit Balance in Company Code Currency' in df.columns):
            df = add_monthly_movement(df, 'Debit Balance in Company Code Currency', 'Credit Balance in Company Code Currency')
        elif 'Debit' in df.columns and 'Credit' in df.columns:
            df = add_monthly_movement(df, 'Debit', 'Credit')
        elif 'ActualMTD' in df.columns:
            df.rename(columns={'ActualMTD': 'Monthly Movement'}, inplace=True)
        else:
            df['Monthly Movement'] = None

        df['Entity'] = entity
        processed_dfs.append((entity, df))

    # Save to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for entity, df in processed_dfs:
            sheet_name = str(entity)[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    st.success("Processing complete!")

    st.download_button(
        label="Download Processed Final Excel File",
        data=output,
        file_name=f'TB_{selected_month_year}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

else:
    st.info("Please upload TB Excel files, TP Category file, and select month/year to begin.")



import streamlit as st
import pandas as pd
import io

st.title("Auto Consolidate Excel Files by Common Sheet Names")

uploaded_files = st.file_uploader(
    "Upload multiple Excel files", accept_multiple_files=True, type=["xlsx", "xls"]
)

if uploaded_files:
    # Step 1: Find common sheet names across all files
    sheet_sets = []
    for file in uploaded_files:
        try:
            xls = pd.ExcelFile(file)
            sheet_sets.append(set(xls.sheet_names))
        except Exception as e:
            st.error(f"Error reading {file.name}: {e}")
            st.stop()
    common_sheets = set.intersection(*sheet_sets)
    
    if not common_sheets:
        st.warning("No common sheet names found across all uploaded files.")
    else:
        st.write(f"### Common sheets found: {list(common_sheets)}")
        
        consolidated_sheets = {}
        errors = []
        
        # Step 2: Consolidate each common sheet
        for sheet in common_sheets:
            dfs = []
            for file in uploaded_files:
                try:
                    df = pd.read_excel(file, sheet_name=sheet)
                    dfs.append(df)
                except Exception as e:
                    errors.append(f"{file.name} - {sheet}: {e}")
            if dfs:
                consolidated_sheets[sheet] = pd.concat(dfs, ignore_index=True)
        
        # Step 4: Provide download option for all consolidated sheets in one Excel
        def to_excel(dfs_dict):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for sheet_name, df in dfs_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            processed_data = output.getvalue()
            return processed_data
        
        excel_data = to_excel(consolidated_sheets)
        
        st.download_button(
            label="Download All Consolidated Sheets as Excel",
            data=excel_data,
            file_name="Consolidated_TB.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        if errors:
            st.error("Errors occurred while reading some sheets:")
            for err in errors:
                st.write(err)
else:
    st.info("Upload multiple Excel files to consolidate common sheets.")
