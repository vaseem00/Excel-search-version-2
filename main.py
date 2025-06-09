import streamlit as st
import pandas as pd

# --- Constants ---
PAGE_SIZE = 10
SEARCH_COLUMNS = [
    'Class', 'Manufacturer Part Number', 'Manufacturer Number', 'Manufacturer', 'Supplier'
]
ALL_COLUMNS = [
    'Class', 'Opportunity', 'PO Number', 'Manufacturer Part Number', 'Manufacturer Number',
    'Manufacturer', 'Supplier', 'Product', 'Part No', 'Origin', 'Mail id', 'Description'
]
EXCEL_FILE_PATH = r"C:\Users\User\Desktop\Procurement automation project\old_erp_plus_outlook_full.xlsx"

# --- Config ---
st.set_page_config(page_title="ERP Search & Analysis Tool", layout="wide")
st.title("🔎 ERP + Outlook Data Explorer")

# --- Load Excel ---
@st.cache_data
def load_excel(path):
    return pd.read_excel(path)

try:
    df = load_excel(EXCEL_FILE_PATH)
    st.success("Excel file loaded successfully from path.")
except Exception as e:
    st.error(f"❌ Failed to load Excel file: {e}")
    st.stop()

# --- Helpers ---
def filter_rows(df_slice, column_name, substring):
    mask = df_slice[column_name].astype(str).str.contains(str(substring), case=False, na=False)
    return df_slice.loc[mask].copy()

def sort_rows(df_slice, columns, ascending):
    return df_slice.sort_values(by=columns, ascending=ascending)

def sort_by_frequency(df_slice, column_name):
    counts = df_slice[column_name].value_counts()
    df_copy = df_slice.copy()
    df_copy['_freq'] = df_copy[column_name].map(counts)
    return df_copy.sort_values(by='_freq', ascending=False).drop(columns=['_freq'])

def nested_frequency_sort(df_slice, columns):
    df_copy = df_slice.copy()
    freq_cols = []
    for level, col in enumerate(columns):
        group_by = columns[:level]
        freq_col = f'_freq_{level}'
        if group_by:
            counts = df_copy.groupby(group_by + [col]).size().rename('count').reset_index()
            df_copy = df_copy.merge(counts, on=group_by + [col], how='left')
            df_copy = df_copy.rename(columns={'count': freq_col})
        else:
            counts = df_copy[col].value_counts()
            df_copy[freq_col] = df_copy[col].map(counts)
        freq_cols.append(freq_col)
    return df_copy.sort_values(by=freq_cols, ascending=[False]*len(freq_cols)).drop(columns=freq_cols)

# --- Search Interface ---
st.subheader("🔍 Search Options")
search_col = st.selectbox("Choose column to search in:", SEARCH_COLUMNS)
search_term = st.text_input(f"Enter search term for '{search_col}'")

# --- State Reset ---
if "filter_reset" not in st.session_state:
    st.session_state["filter_reset"] = False

if st.button("🔄 Reset All Filters"):
    for col in ALL_COLUMNS:
        st.session_state[f"filter_dropdown_{col}"] = []
    st.session_state["filter_reset"] = True
    st.experimental_rerun()

# --- Filter Data ---
if search_term:
    filtered_df = filter_rows(df, search_col, search_term)
    st.markdown(f"✅ **{len(filtered_df)}** matching rows found.")

    with st.expander("🧰 Filter by Column Values"):
        filter_cols = st.multiselect("Select columns to filter:", ALL_COLUMNS)

        for col in filter_cols:
            unique_vals = df[col].dropna().astype(str).unique().tolist()
            selected_vals = st.multiselect(
                f"Choose values for '{col}'",
                options=sorted(unique_vals),
                default=st.session_state.get(f"filter_dropdown_{col}", []),
                key=f"filter_dropdown_{col}"
            )
            if selected_vals:
                filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]
                st.session_state[f"filter_dropdown_{col}"] = selected_vals

    # --- Sorting ---
    with st.expander("↕️ Sorting"):
        sort_cols = st.multiselect("Choose columns to sort by:", ALL_COLUMNS)
        sort_mode = {}
        for col in sort_cols:
            sort_mode[col] = st.selectbox(
                f"Sort mode for '{col}'",
                ['Ascending', 'Descending', 'Frequency'],
                key=f"sort_mode_{col}"
            )
        if sort_cols:
            if all(m == 'Frequency' for m in sort_mode.values()):
                filtered_df = nested_frequency_sort(filtered_df, sort_cols)
            else:
                for col in reversed(sort_cols):
                    mode = sort_mode[col]
                    if mode == 'Frequency':
                        filtered_df = sort_by_frequency(filtered_df, col)
                    else:
                        filtered_df = sort_rows(filtered_df, [col], [mode == 'Ascending'])

    # --- Pagination ---
    st.subheader("📄 Results")
    total_rows = len(filtered_df)
    total_pages = max((total_rows - 1) // PAGE_SIZE + 1, 1)
    page = st.number_input("Page", min_value=1, max_value=total_pages, value=1)
    start_idx = (page - 1) * PAGE_SIZE
    end_idx = start_idx + PAGE_SIZE

    page_data = filtered_df.iloc[start_idx:end_idx]

    for idx, row in page_data.iterrows():
        with st.expander(f"Row {idx+1}"):
            for col in filtered_df.columns:
                val = row[col] if pd.notna(row[col]) else ""
                st.markdown(f"**{col}:** {val}")

    st.caption(f"Showing page {page} of {total_pages} — rows {start_idx + 1} to {min(end_idx, total_rows)}")
