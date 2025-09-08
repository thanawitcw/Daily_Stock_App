import streamlit as st
import pandas as pd
import numpy as np
import io
from pandasql import sqldf
from datetime import datetime, timedelta

# Set page config
st.set_page_config(
    page_title="Daily Stock Report Generator",
    layout="wide"
)


st.title("🤖 Daily Stock Report Generator")
st.markdown("Follow the instructions below to generate your daily stock report.")

# Instructions section
with st.expander("📖 Instructions & Features"):
    st.markdown("""
    ### 📌 How to use this tool:

    **Required Files:**
    1. **Upload all required files in the order specified below.**

        1.1 Sahamit Report ..-..-2025.xlsx
        
        1.2 DC_End..-..-2025.xlsx
        
        1.3 sellout_past30D.xlsx
        
        1.4 data_from_access.xlsx

        1.5 Summary Forecast.xlsx
        
        1.6 1.Master_LeadTime.xlsx
                
    2. **Click the "Generate Daily Stock Report" button to process the data.**
                
    3. **Download the output file as a .xlsx file.**
    """)

# File upload section
st.header("Required Files")

st.subheader("Step 1: 📂 Upload Weekly CJ Stock File")
master_file = st.file_uploader("File Name => Sahamit Report ..-..-2025", type=['xlsx'], key="master")

st.subheader("Step 2: 📂 Upload Yesterday DC Stock File ")
dc_stock_file = st.file_uploader("File Name => DC_End..-..-2025", type=['xlsx'], key="dc_stock")

st.subheader("Step 3: 📂 Upload last 30 days sales File")
sellout_file = st.file_uploader("File Name => sellout_past30D", type=['xlsx'], key="sellout")

st.subheader("Step 4: 📂 Upload Master Product & PO Pending Foods / NF / Import-Local File")
access_db_extracted_file = st.file_uploader("File Name => data_from_access", type=['xlsx'], key="extract_access_db")

st.subheader("Step 5: 📂 Upload PO Pending Import-Foods/NF - Overseas")
po_import_overseas = st.file_uploader("File Name => Summary Forecast", type=['xlsx'], key="po_import")

st.subheader("Step 6: 📂 Upload Master Lead Time File")
master_leadtime_file = st.file_uploader("File Name => 1.Master_LeadTime", type=['xlsx'], key="leadtime")


# Processing functions
def process_cj_stock(master_file):
    """Process CJ Stock data"""
    master_df = pd.read_excel(master_file, sheet_name='Sahamit Report', header=2)
    
    if 'Product' in master_df.columns:
        master_df.rename(columns={'Product': 'CJ_Item'}, inplace=True)
    
    filter_master_file = ['A-HOME', 'UNO']
    master_df = master_df[~master_df['Division'].isin(filter_master_file)]
    
    return master_df



def process_dc_stock(dc_stock_file):
    """Process DC Stock data"""
    if dc_stock_file is None:
        return None
    
    try:
        # Load DC Stock file
        master_df = pd.read_excel(dc_stock_file, sheet_name='Sheet1')
        
        # Rename columns
        master_df.rename(columns={
            master_df.columns[2]: 'DC_Name', 
            master_df.columns[4]: 'Product Name'}, inplace=True)
        
        # Map Plant column
        master_df['Plant'] = master_df['Plant'].map({
            'D001': 'DC1',
            'D002': 'DC2',
            'D004': 'DC4'
        })
        
        # Replace <0 with 0
        master_df['Stock Qty'] = master_df['Stock Qty'].where(master_df['Stock Qty'] >= 0, 0)
        master_df['Stock Value'] = master_df['Stock Value'].where(master_df['Stock Value'] >= 0, 0)
        
        # Pivot data
        pivoted_df = master_df.pivot_table(
            index=['Material'],
            columns='Plant',
            values=['Stock Qty', 'Stock Value'],
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        # Flatten the multi-level columns
        pivoted_df.rename(columns={'Material': 'CJ_Item'}, inplace=True)
        pivoted_df.columns = [
            f"{plant}_Remain_StockQty" if value == 'Stock Qty' else f"{plant}_Remain_StockValue"
            if plant else 'CJ_Item'  # Keep 'CJ_Item' column unchanged
            for value, plant in pivoted_df.columns
        ]
        
        return pivoted_df
    
    except Exception as e:
        st.warning(f"Could not process DC Stock file: {str(e)}")
        return None


def process_sellout_data(sellout_file):
    """Process Daily Sell-Out data"""
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0) 
    last_7d = today - timedelta(days=7)
    
    df = pd.read_excel(sellout_file, header=0)
    
    df.columns = ['Calendar Day', 'Supplier', 'Supplier_Name', 'Product Group', 'CJ_Item',
                  'Product Name', 'Brand', 'Brand Name', 'Status', 'EAN/UPC', 'Sales unit',
                  'Total_SellOut_Qty', 'Total_Sellout_Amt', 'Normal Sale/Base QTY (TY)',
                  'Normal Sale/AMT (TY) Exc.VAT', 'Promotion Sale/Base QTY (TY)',
                  'Promotion Sale/AMT (TY) Exc.VAT']
    
    df['Calendar Day'] = pd.to_datetime(df['Calendar Day'], format='%d.%m.%Y', errors='coerce')
    
    filter_df_data = ['A-HOME','LIFESTYLE']
    df = df[~df['Product Group'].isin(filter_df_data)]
    
    pivot_sales_last30d = df.pivot_table(
        index='CJ_Item', 
        values='Total_SellOut_Qty',
        aggfunc='sum'
        ).reset_index().rename(columns={'Total_SellOut_Qty':'SO_Qty_last30D'})

    pivot_sales_last7d = df[df['Calendar Day'] >= last_7d].pivot_table(
        index='CJ_Item', 
        values='Total_SellOut_Qty', 
        aggfunc='sum', 
        fill_value=0
        ).reset_index().rename(columns={'Total_SellOut_Qty':'SO_Qty_last7D'})
    
    merged_pivot = pivot_sales_last30d.merge(pivot_sales_last7d, on='CJ_Item', how='left')
    merged_pivot = merged_pivot.fillna(0)
    
    return merged_pivot


# Hold this data HBA
#def process_po_hba(po_hba_file):
    """Process PO HBA data"""
    df = pd.read_excel(po_hba_file, sheet_name='PendingPO', header=0)
    df.columns = df.columns.str.strip()
    
    selected_columns = df[['CJ_Article', 
                           'SHM_Article', 
                           'SHM_PO NO.', 
                           'CJ_PO Date',
                           'CJ_PO NO.', 
                           'CJ_Description', 
                           'DC_Location', 
                           'PO Pending',
                           'PO_Status', 
                           'Next Delivery', 
                           'Supplier_Short_Name']]
    
    selected_columns.rename(columns={'CJ_PO Date': 'SHM_PO_Date'}, inplace=True)
    date_cols_to_convert = ['SHM_PO_Date', 'Next Delivery']
    for col in date_cols_to_convert:
        if col in selected_columns:
            selected_columns[col] = pd.to_numeric(selected_columns[col], errors = 'coerce')
            selected_columns[col] = pd.to_datetime(selected_columns[col], origin = '1899-12-30', unit = 'D', errors='coerce')

    selected_columns = selected_columns[selected_columns['PO_Status'].str.strip().str.lower() == 'pending']
    selected_columns = selected_columns[selected_columns['CJ_Article'].notna() & (~selected_columns['CJ_Article'].isin(['Tester', 'New']))]
    
    selected_columns['CJ_Article'] = selected_columns['CJ_Article'].astype(str)
    selected_columns['CJ_Article'] = selected_columns['CJ_Article'].apply(lambda x: x.rstrip('0').rstrip('.') if '.' in x else x)
    selected_columns['SHM_Article'] = selected_columns['SHM_Article'].astype(str)
    

    selected_columns.loc[:, 'DC'] = selected_columns['DC_Location'].map({
        'D001': 'DC1', 
        'D002': 'DC2', 
        'D004': 'DC4',
        'TD09': 'TD09'
    })
    
    df2 = pd.read_excel(po_hba_file, sheet_name='Supply Record', header=0)
    selected_df2 = df2[['SHM_Article', 'Unit_PC/CAR']]
    selected_df2.rename(columns={'Unit_PC/CAR': 'PC_Cartons'}, inplace=True)
    selected_df2['SHM_Article'] = selected_df2['SHM_Article'].astype(str)
    
    selected_columns = pd.merge(selected_columns, selected_df2, on='SHM_Article', how='left')

    selected_columns.loc[:, 'PendingPO (CTN)'] = selected_columns['PO Pending'] / selected_columns['PC_Cartons']
    selected_columns.loc[:, 'PendingPO (CTN)'] = selected_columns['PendingPO (CTN)'].round(0).astype(int)
    
    pivoted_df = selected_columns.pivot_table(
        index=['CJ_Article', 'SHM_Article'], 
        columns=['DC'], 
        values='PO Pending', 
        aggfunc='sum', 
        fill_value=0
        ).reset_index()
    
    pivoted_df.columns = ['CJ_Item', 'SHM_Item'] + [f'PO_Qty_to_{col}' for col in pivoted_df.columns[2:]]
    
    pivoted_min_del_date = selected_columns.pivot_table(
        index=['CJ_Article', 'SHM_Article'], 
        columns='DC', 
        values='Next Delivery', 
        aggfunc='min'
        ).reset_index()
    
    pivoted_min_del_date.columns = ['CJ_Item', 'SHM_Item'] + [f'Min_del_date_to_{col}' for col in pivoted_min_del_date.columns[2:]]
    
    merged_df = pd.merge(pivoted_df, pivoted_min_del_date, on=['CJ_Item', 'SHM_Item'], how='left')

    desired_columns = [
            'CJ_Item', 'SHM_Item',
            'PO_Qty_to_DC1', 'PO_Qty_to_DC2', 'PO_Qty_to_DC4',
            'Min_del_date_to_DC1', 'Min_del_date_to_DC2', 'Min_del_date_to_DC4'
        ]
    
    for col in desired_columns:
            if col not in merged_df.columns:
                merged_df[col] = 0 if 'PO_Qty' in col else pd.NaT

    merged_df['Min_del_date_to_DC4'] = pd.to_datetime(merged_df['Min_del_date_to_DC4'], errors = 'coerce')
    
    return merged_df, selected_columns


def process_all_data_from_access(excel_data):
    if excel_data is None:
        st.info("Please upload file named => data_from_access.xlsx")
        return None

    try:
        datasets = {}

        datasets['product_list'] = pd.read_excel(excel_data, sheet_name='Master_Product')
        st.success("Loaded 'Master_Product' sheet")

        datasets['po_foods_nf_pcb'] = pd.read_excel(excel_data, sheet_name='Pending_Foods&NF&PCB')
        st.success("Loaded 'Pending_Foods&NF' sheet")

        datasets['po_import_local'] = pd.read_excel(excel_data, sheet_name='Pending_Import_Local')
        st.success("Loaded 'Pending_Import' sheet")

        st.success("All specified sheets loaded successfully!")
        return datasets

    except Exception as e:
        st.error(f"An error occurred while processing the data_from_access file: {str(e)}")
        return None


def process_po_other(excel_datasets):
    if not excel_datasets or 'po_foods_nf_pcb' not in excel_datasets or 'product_list' not in excel_datasets:
        st.error("Missing required datasets for PO Foods/NF processing")
        return None
    
    try:
        df = excel_datasets['po_foods_nf_pcb'].copy()
        df['CJ_Item'] = df['CJ_Item'].astype(str)
        df['Item'] = df['Item'].astype(str)

        
        product_data = excel_datasets['product_list'].copy()
        product_data['CJ_Item'] = product_data['CJ_Item'].astype(str)
        product_data['SHM_Item'] = product_data['SHM_Item'].astype(str)

        df.rename(columns={'Item': 'SHM_Item'}, inplace=True)

        dc_name_mapping = {
            'CJ DC1 ราชบุรี': 'DC1',
            'CJ DC2 บางปะกง': 'DC2',
            'DC โพธาราม': 'DC1',
            'DC บางวัว 1': 'DC2',
            'DC ขอนแก่น': 'DC4',
            'DC บางวัว 2': 'TD09'
        }
        df['DC_Name'] = df['DC_Name'].replace(dc_name_mapping)

        df = pd.merge(df, product_data[['CJ_Item', 'SHM_Item', 'PC_Cartons']], on=['CJ_Item', 'SHM_Item'], how='left')

        # Handle with missing PC_Cartons
        df['PC_Cartons'] = df['PC_Cartons'].fillna(1)
        # Create new column
        df['PO_Qty'] = df.apply(
            lambda row: row['Order Qty'] if row['Unit'] == 'ชิ้น' 
            else row['Order Qty'] * row['PC_Cartons'], axis=1
            )

        pivot_po_qty = df.pivot_table(
            index=['SHM_Item','CJ_Item'],
            columns=['DC_Name'],
            values='PO_Qty',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        pivot_min_del_date = df.pivot_table(
            index=['SHM_Item','CJ_Item'],
            columns=['DC_Name'],
            values='Rec_Date',
            aggfunc='min'
        ).reset_index()

        pivoted_df = pivot_po_qty.merge(pivot_min_del_date, on=['SHM_Item', 'CJ_Item'], suffixes=('_fromPivotQty', '_fromPivotMindate'))

        pivot_po_qty_columns = [f'PO_Qty_to_{col}' if col not in ['SHM_Item', 'CJ_Item'] else col for col in pivot_po_qty.columns]
        pivot_min_del_date_columns = [f'Min_del_date_to_{col}' if col not in ['SHM_Item', 'CJ_Item'] else col for col in pivot_min_del_date.columns]
        pivoted_df.columns = pivot_po_qty_columns[:2] + pivot_po_qty_columns[2:] + pivot_min_del_date_columns[2:]

        return pivoted_df, df

    except Exception as e:
        st.warning(f"An error occurred while processing PO Other data: {str(e)}")
        return None
    


def process_po_import(excel_datasets, po_import_overseas):
    if not excel_datasets or 'po_import_local' not in excel_datasets or 'product_list' not in excel_datasets:
        st.error("Missing required datasets for PO Import processing")
        return None
     
    try:
        df = excel_datasets['po_import_local'].copy()
        product_data = excel_datasets['product_list'].copy()
        

        # Make sure data types matched
        if 'CJ_Item' in df.columns:
            df['CJ_Item'] = df['CJ_Item'].astype(str)
        if 'SHM_Item' in df.columns:
            df['SHM_Item'] = df['SHM_Item'].astype(str)

        if 'CJ_Item' in product_data.columns:
            product_data['CJ_Item'] = product_data['CJ_Item'].astype(str)
        if 'SHM_Item' in product_data.columns:
            product_data['SHM_Item'] = product_data['SHM_Item'].astype(str)
        
        df.rename(columns={'PO Ref': 'PO CJ No.'}, inplace=True)

        dc_name_mapping = {
            'CJ DC1 ราชบุรี': 'DC1',
            'CJ DC2 บางปะกง': 'DC2',
            'DC โพธาราม': 'DC1',
            'DC บางวัว 1': 'DC2',
            'DC ขอนแก่น': 'DC4',
            'DC บางวัว 2': 'TD09'
        }
        df['DC_Name'] = df['DC_Name'].replace(dc_name_mapping)


        df['PO_Qty'] = df.apply(
            lambda row: row['Order Qty'] * row['PC_Cartons'] if row['Unit'] == 'ลัง' 
            else row['Order Qty'], axis=1)
        
        df['PO_CTN'] = df.apply(
            lambda row: row['Order Qty'] if row['Unit'] == 'ลัง' 
            else row['Order Qty'] / row['PC_Cartons'], axis=1)


        pivot_df = df.pivot_table(
            index=['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'],
            columns = 'DC_Name',
            values= 'PO_Qty',
            aggfunc='sum'
        ).reset_index()
        pivot_df.columns = ['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'] + [f'PO_Qty_to_{col}' for col in pivot_df.columns[4:]]

        pivot_df2 = df.pivot_table(
            index=['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'],
            columns = 'DC_Name',
            values= 'PO_CTN',
            aggfunc='sum'
        ).reset_index()
        pivot_df2.columns = ['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'] + [f'cartons_to_{col}' for col in pivot_df2.columns[4:]]

        pivot_df3 = df.pivot_table(
            index=['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'],
            columns='DC_Name',
            values='Rec_Date',
            aggfunc='min'
        ).reset_index()
        pivot_df3.columns = ['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'] + [f'Date_to_{col}' for col in pivot_df3.columns[4:]]

        # MERGE 2 pivot tables
        merged_df = pd.merge(pivot_df, pivot_df2, on=['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'], how='left')
        merged_df2 = pd.merge(merged_df, pivot_df3, on=['SHM_Item', 'CJ_Item', 'Product Name','PO CJ No.'], how='left')

        # Work with PO Pending from Excel file
        df2 = pd.read_excel(po_import_overseas, sheet_name='vstack', header=0)
        df2['แผนส่งเข้าคลังโพธาราม'] = pd.to_datetime(df2['แผนส่งเข้าคลังโพธาราม'], errors='coerce')
        df2['แผนส่งเข้าคลังบางปะกง'] = pd.to_datetime(df2['แผนส่งเข้าคลังบางปะกง'], errors='coerce')
        df2['แผนส่งเข้าคลังขอนแก่น'] = pd.to_datetime(df2['แผนส่งเข้าคลังขอนแก่น'], errors ='coerce')
        df2['เรือ ETA'] = pd.to_datetime(df2['เรือ ETA'], errors ='coerce')
        df2['CJ_Item'] = df2['CJ_Item'].astype(str)
    
        today = pd.to_datetime(datetime.now().date())

        df2_filtered = df2[(df2['แผนส่งเข้าคลังโพธาราม'] >= today) | (df2['แผนส่งเข้าคลังบางปะกง'] >= today) | (df2['แผนส่งเข้าคลังขอนแก่น'] >= today)]
    
        selected_columns = df2_filtered[['SHM_Item', 
                                         'CJ_Item', 
                                         'Name', 
                                         'PO CJ No.', 
                                         'เรือ ETA',
                                         'แผนส่งเข้าคลังโพธาราม', 
                                         'ยอดสั่งโพธาราม(ลัง)',
                                         'แผนส่งเข้าคลังบางปะกง', 
                                         'ยอดสั่งบางปะกง(ลัง)',
                                         'แผนส่งเข้าคลังขอนแก่น', 
                                         'ยอดสั่งขอนแก่น(ลัง)']]
    
        selected_columns.rename(columns={
            'Name' : 'Product Name',
            'แผนส่งเข้าคลังโพธาราม': 'Date_to_DC1',
            'ยอดสั่งโพธาราม(ลัง)': 'cartons_to_DC1',
            'แผนส่งเข้าคลังบางปะกง': 'Date_to_DC2',
            'ยอดสั่งบางปะกง(ลัง)': 'cartons_to_DC2',
            'แผนส่งเข้าคลังขอนแก่น': 'Date_to_DC4',
            'ยอดสั่งขอนแก่น(ลัง)': 'cartons_to_DC4'
            }, inplace=True)


        # Finalize the po pedning import
        cleaned_data = pd.merge(selected_columns, product_data[['CJ_Item', 'SHM_Item', 'PC_Cartons']], on=['CJ_Item', 'SHM_Item'], how='left')

        cleaned_data['PO_Qty_to_DC1'] = cleaned_data['cartons_to_DC1'] * cleaned_data['PC_Cartons']
        cleaned_data['PO_Qty_to_DC2'] = cleaned_data['cartons_to_DC2'] * cleaned_data['PC_Cartons']
        cleaned_data['PO_Qty_to_DC4'] = cleaned_data['cartons_to_DC4'] * cleaned_data['PC_Cartons']

        existing_po_cj_no = cleaned_data['PO CJ No.'].unique()
        merged_df2_filtered = merged_df2[~merged_df2['PO CJ No.'].isin(existing_po_cj_no)]

        combined_data = pd.concat([cleaned_data, merged_df2_filtered], ignore_index=True)

        pivot_data = combined_data.groupby(['SHM_Item', 'CJ_Item']).agg({
            'PO_Qty_to_DC1': 'sum',
            'Date_to_DC1' : 'min',
            'PO_Qty_to_DC2': 'sum',
            'Date_to_DC2' : 'min',
            'PO_Qty_to_DC4': 'sum',
            'Date_to_DC4' : 'min'
        }).reset_index()


        pivot_data.rename(columns={
            'Date_to_DC1': 'Min_del_date_to_DC1',
            'Date_to_DC2': 'Min_del_date_to_DC2',
            'Date_to_DC4': 'Min_del_date_to_DC4'
        }, inplace=True)

        return pivot_data, combined_data
    
    except Exception as e:
        st.warning(f"An error occurred while processing PO Import data: {str(e)}")
        return None


def process_leadtime(master_owner_lt_file):
    """Process Lead Time data"""
    try:
        master_leadtime = pd.read_excel(master_owner_lt_file, sheet_name='All_Product', header=1)

        master_leadtime = master_leadtime[['SHM_Item', 'CJ_Item', 'OwnerSCM', 'Base Lead Time (Days)']]
        master_leadtime.rename(columns={'Base Lead Time (Days)': 'LeadTime'}, inplace=True)

        master_leadtime['CJ_Item'] = master_leadtime['CJ_Item'].astype(str)

        master_leadtime['CJ_Item'] = master_leadtime['CJ_Item'].fillna('')
        master_leadtime['CJ_Item'] = master_leadtime['CJ_Item'].apply(lambda x: x.split('.')[0] if '.0' in str(x) else x)

        return master_leadtime

    except Exception as e:
        st.warning(f"Error processing Lead Time file: {str(e)}")
        return None


def combine_all_PO_data(foods_nf_pcb_df, import_po_df, owner_scm_df):
    try:
        # Step 1: Use already processed DataFrames
        if foods_nf_pcb_df is None or import_po_df is None or owner_scm_df is None:
            st.error("One or more required data sources are missing. Please check your uploads.")
            return None, None

        # Extract product list
        access_product_list_df = access_datasets.get('product_list')
        if access_product_list_df is None:
            st.error("sheet name 'product_list' not found. Please Check 'Master_Product' sheet.")
            return None, None

        # Process Lead Time
        owner_scm_df = process_leadtime(master_leadtime_file)
        if owner_scm_df is None:
            st.error("Failed to process Lead Time file.")
            return None, None

    except Exception as e:
        st.error(f"❌ Error in Step 1 of combine_all_PO_data: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None, None

    
    # Step 2: Clean each dataframe to a unified format
    # Hold this HBA Data transformation 
    #def clean_po_hba(df):
        df = df.drop(columns=['DC_Location'])
        df.rename(columns={
            'CJ_Article': 'CJ_Item', 
            'SHM_Article': 'SHM_Item', 
            'SHM_PO NO.': 'SHM PO No.',
            'SHM_PO_Date': 'PO Date', 
            'CJ_PO NO.': 'PO CJ No.', 
            'CJ_Description' : 'Product Name',
            'DC': 'Ship to DC', 
            'PO Pending': 'PO_Qty', 
            'PO_Status': 'Delivery_Status',
            'Next Delivery': 'Delivery Date', 
            'Supplier_Short_Name': 'Supplier Name',
            'PendingPO (CTN)': 'PO Cartons',
        }, inplace=True)
        return df

    def clean_po_foods_nf_pcb(df):
        # Exclude unnecessary columns
        df = df.drop(columns=['Devision','Unit','Customer'])
        # Rename Column
        df.rename(columns={
            'PO Num': 'SHM PO No.', 
            'PO Ref': 'PO CJ No.', 
            'DC_Name': 'Ship to DC',
            'Rec_Date': 'Delivery Date', 
            'Order Qty': 'PO Cartons', 
            'CJ_Description':'Product Name'
        },inplace=True)
        return df


    def clean_po_import(df):
        df_unpivoted = df.melt(
            id_vars=['SHM_Item', 'CJ_Item', 'Product Name', 'PO CJ No.', 'เรือ ETA', 'PC_Cartons'],
            value_vars=['Date_to_DC1', 'cartons_to_DC1', 'Date_to_DC2', 'cartons_to_DC2', 'Date_to_DC4', 'cartons_to_DC4', 
                    'PO_Qty_to_DC1', 'PO_Qty_to_DC2', 'PO_Qty_to_DC4'],
            var_name='variable',
            value_name='value'
        )
        
        df_unpivoted['Ship to DC'] = df_unpivoted['variable'].apply(lambda x: 'DC1' if 'DC1' in x else ('DC2' if 'DC2' in x else 'DC4'))
        df_unpivoted['PO Cartons'] = df_unpivoted.apply(lambda row: row['value'] if 'cartons' in row['variable'] else None, axis=1)
        df_unpivoted['PO qty'] = df_unpivoted.apply(lambda row: row['value'] if 'PO_Qty' in row['variable'] else None, axis=1)
        df_unpivoted['Date to DC'] = df_unpivoted.apply(lambda row: row['value'] if 'Date' in row['variable'] else None, axis=1)
        df_unpivoted.drop(columns=['variable', 'value'], inplace=True)
        df_unpivoted.dropna(subset=['PO Cartons', 'PO qty', 'Date to DC'], how='all', inplace=True)

        df_unpivoted.fillna({"เรือ ETA": "Unknown"}, inplace=True) 
        df_unpivoted.fillna({'PC_Cartons': "0"},inplace=True)
        
        # Pivot the data back to combine rows with the same SHM_Item, CJ_Item, PO CJ No., PC_Cartons, and Ship to DC
        final_import = df_unpivoted.groupby(
            ["SHM_Item", "CJ_Item", "Product Name", "PO CJ No.", "เรือ ETA", "PC_Cartons", "Ship to DC"]
        ).agg({
            "PO Cartons": "sum",
            "PO qty": "sum",
            "Date to DC": "first"
        }).reset_index()

        final_import["เรือ ETA"] = final_import["เรือ ETA"].replace("Unknown", pd.NA)
        final_import['PC_Cartons'] = final_import['PC_Cartons'].replace("0", pd.NA)
    
        # Rename columns
        final_import.rename(columns={
            'Date to DC': 'Delivery Date',
            'PO qty': 'PO_Qty'
            }, inplace=True)

        return final_import

    # Clean each dataset
    foods_nf_pcb_clean = clean_po_foods_nf_pcb(foods_nf_pcb_df) 
    import_clean = clean_po_import(import_po_df)

    # Step 3: Combine all datasets
    combined_df = pd.concat([foods_nf_pcb_clean, import_clean], ignore_index=True)
    combined_df['CJ_Item'] = combined_df['CJ_Item'].astype(str)
    combined_df['SHM_Item'] = combined_df['SHM_Item'].astype(str)

        
    # Step 4: Enrich with master product info
    access_product_list_df['CJ_Item'] = access_product_list_df['CJ_Item'].astype(str)
    access_product_list_df['SHM_Item'] = access_product_list_df['SHM_Item'].astype(str)
    merged_df = pd.merge(combined_df, access_product_list_df, on=['CJ_Item', 'SHM_Item'], how='left', suffixes=('', '_access'))

    merged_df['Supplier Name'] = merged_df['Supplier Name'].fillna(merged_df['Supplier Name_access'])
    merged_df['Division'] = merged_df['Devision']


    # Step 5: Merge OwnerSCM
    owner_scm_df_selected = owner_scm_df[['SHM_Item', 'CJ_Item', 'OwnerSCM']]
    owner_scm_df_selected['CJ_Item'] = owner_scm_df_selected['CJ_Item'].fillna('')
    final_df = pd.merge(merged_df, owner_scm_df_selected, on=['CJ_Item', 'SHM_Item'], how='left')

    # Step 6: Final column ordering
    desired_order = [
        'Division', 'OwnerSCM', 'PO Date', 'SHM PO No.', 
        'Supplier Name', 'SHM_Item', 'CJ_Item', 'Product Name', 
        'PO CJ No.', 'PC_Cartons', 'Ship to DC', 'PO Cartons', 
        'PO_Qty', 'เรือ ETA', 'Delivery Date', 'Delivery_Status'
    ]
    
    # Add missing columns
    for col in desired_order:
        if col not in final_df.columns:
            final_df[col] = pd.NA
    final_df = final_df[desired_order]

    final_df['เรือ ETA'] = pd.to_datetime(final_df['เรือ ETA'], errors='coerce')

    # Step 7: Create pivot table
    final_df_pivot = final_df.pivot_table(
        index=['CJ_Item', 'Ship to DC'],
        values=['เรือ ETA', 'Delivery Date'],
        aggfunc='min'
    ).reset_index().rename(columns={'เรือ ETA': 'Min ETA', 'Delivery Date': 'First Delivery Date'})
    final_df_pivot['ConcatIndex'] = final_df_pivot['CJ_Item'] + final_df_pivot['Ship to DC']

    return final_df, final_df_pivot



def convert_cj_item_to_string(dataframes, access_df):
    """Convert CJ_Item to string in various dataframes"""
    def process_column(df, column):
        if df is not None and column in df.columns:
            df[column] = df[column].astype(str).str.split('.').str[0]

    for key, df in dataframes.items():
        if df is not None:
            process_column(df, 'CJ_Item')
            process_column(df, 'SHM_Item')
    
    if access_df is not None:
        process_column(access_df, 'CJ_Item')

    return dataframes, access_df



def merge_dataframes(dfs, access_df):
    """Merge All dataframes into a single dataframe"""
    required_dfs = ['CJ_Stock', 'Daily_Stock_DC', 'Daily_SO', 'PO_Foods_NF_PCB', 'PO_Import']
    for df_name in required_dfs:
        if df_name not in dfs or dfs[df_name] is None:
            st.error(f"Missing required dataframe: {df_name}. Please check your uploads.")
            return None
    
    merged_df = access_df.merge(dfs['CJ_Stock'], on='CJ_Item', how='left')

    merged_df = merged_df.merge(dfs['Daily_SO'], on='CJ_Item', how='left')  

    merged_df = merged_df.merge(dfs['PO_Import'], on='SHM_Item', how='outer', suffixes=('', '_from-Import'))

    merged_df = merged_df.merge(dfs['PO_Foods_NF_PCB'], on='SHM_Item', how='outer', suffixes=('', '_from-Access'))  

    merged_df = merged_df.merge(dfs['Daily_Stock_DC'], on='CJ_Item', how='left', suffixes=('', '_from-DailyDC'))  

    # Rename columns
    merged_df = merged_df.rename(columns={
        'Division': 'Division_CJ_stock',  
        'Devision': 'Division_SHM'
    })  

    # Create a new column for NPD Status by First_SO_Date, current logic = First SO date + 15 days
    today = pd.to_datetime(datetime.now().date())  
    merged_df['days_from_first_ATP'] = (today - pd.to_datetime(merged_df['First_SO_Date'])).dt.days  
    merged_df['NPD_Status'] = np.where(merged_df['days_from_first_ATP'] <= 15, 'NPD', '-')

    # Fill in missing values for descriptive columns  
    merged_df['Name'] = merged_df['Name'].fillna(merged_df['CJ_Description'])  
    merged_df['Category'] = merged_df['Category'].fillna(merged_df['Cat'])  
    merged_df['Subcate'] = merged_df['Subcate'].fillna(merged_df['Sub_cat'])  
    
    return merged_df



def fill_na_with_zero(df):
    """Fill Nan Values in numeric column with 0"""
    numeric_cols = df.select_dtypes(include=np.number).columns
    df[numeric_cols] = df[numeric_cols].fillna(0)
    return df



# Create new column to sum ALL PO Pending
def calculate_totals(merged_df):
    dc_columns = [1, 2, 4]
    for dc in dc_columns:
        merged_df[f'Total-PO_qty_to_DC{dc}'] = (
            merged_df[f'PO_Qty_to_DC{dc}'] +
            merged_df[f'PO_Qty_to_DC{dc}_from-Access'])

        # Calculate %Ratio with error handling
        merged_df[f'%Ratio_AvgSalesQty90D_DC{dc}'] = (
            merged_df[f'DC{dc}_AvgSaleQty90D'] / merged_df['Total_AvgSaleQty90D'].replace(0, 1)
        ).replace([np.inf, -np.inf], 0)

    # Calculate Total Remain Stock
    merged_df['Remain_StockQty_AllDC'] = merged_df['DC1_Remain_StockQty'] + merged_df['DC2_Remain_StockQty'] + merged_df['DC4_Remain_StockQty']
    merged_df['Remain_StockValue_AllDC'] = merged_df['DC1_Remain_StockValue'] + merged_df['DC2_Remain_StockValue'] + merged_df['DC4_Remain_StockValue']

    # Calculate SO Qty
    for dc in dc_columns:
        merged_df[f'DC{dc}_SO_Last30D'] = round(merged_df['SO_Qty_last30D'] * merged_df[f'%Ratio_AvgSalesQty90D_DC{dc}'])
        merged_df[f'DC{dc}_SO_Last7D'] = round(merged_df['SO_Qty_last7D'] * merged_df[f'%Ratio_AvgSalesQty90D_DC{dc}'])

        merged_df[f'DC{dc}_AvgSaleQty30D'] = merged_df[f'DC{dc}_SO_Last30D'] / 30
        merged_df[f'DC{dc}_AvgSaleQty7D'] = merged_df[f'DC{dc}_SO_Last7D'] / 7

    merged_df['Total_AvgSaleQty30D'] = merged_df[[f'DC{dc}_AvgSaleQty30D' for dc in dc_columns]].sum(axis=1)
    merged_df['Total_AvgSaleQty7D'] = merged_df[[f'DC{dc}_AvgSaleQty7D' for dc in dc_columns]].sum(axis=1)

    return merged_df


# Simplify DOH calculation
def calculate_DOH(stock_qty, avg_qty):
    return np.where(
        (stock_qty > 0) & (avg_qty == 0), 365,
        np.where((stock_qty == 0) & (avg_qty == 0), 0, stock_qty / avg_qty)
    )

def apply_doh_calculations(merged_df):
    dc_list = ['DC1', 'DC2', 'DC4']
    max_doh_value = 1825
    current_date = pd.to_datetime(datetime.now().date())

    # Calculate DC DOH and DOH after PO for each DC
    for dc in dc_list:
        merged_df[f'Current_{dc}_DOH'] = calculate_DOH(
            merged_df[f'{dc}_Remain_StockQty'],
            merged_df[f'{dc}_AvgSaleQty90D']
        )
        merged_df[f'{dc}_DOH(Stock+PO)'] = calculate_DOH(
            merged_df[f'{dc}_Remain_StockQty'] + merged_df[f'Total-PO_qty_to_{dc}'],
            merged_df[f'{dc}_AvgSaleQty90D']
        )

    # Special columns for all DCs
    merged_df['Current_DOH_All_DC'] = calculate_DOH(
        merged_df['Remain_StockQty_AllDC'],
        merged_df['Total_AvgSaleQty90D']
    )
    merged_df['Total-PO_qty_to_DC'] = merged_df[[f'Total-PO_qty_to_{dc}' for dc in dc_list]].sum(axis=1)

    merged_df['Current_DOH(Stock+PO)_All_DC'] = calculate_DOH(
        merged_df['Remain_StockQty_AllDC'] + merged_df['Total-PO_qty_to_DC'],
        merged_df['Total_AvgSaleQty90D']
    )

    # Ensure all delivery date columns are datetime
    dc_date_columns = {
        'DC1': ['Min_del_date_to_DC1', 'Min_del_date_to_DC1_from-Access'],
        'DC2': ['Min_del_date_to_DC2', 'Min_del_date_to_DC2_from-Access'],
        'DC4': ['Min_del_date_to_DC4', 'Min_del_date_to_DC4_from-Access']
    }

    for dc, cols in dc_date_columns.items():
        for col in cols:
            data_cols = merged_df[col].replace(['', '0', 0, 'null', 'None', '#N/A'], pd.NaT)
            data_cols = pd.to_datetime(data_cols, errors='coerce')
            data_cols = data_cols.where(~(data_cols.dt.date == pd.Timestamp('1970-01-01').date()), pd.NaT)

            merged_df[col] = data_cols

        # Calculate Min deld ate only if at least one column has a non-null value
        merged_df[f'Min_delivery_date_to_{dc}'] = merged_df[cols].min(axis=1)

    merged_df['Min_delivery_date_to_DC'] = merged_df[[f'Min_delivery_date_to_{dc}' for dc in dc_list]].min(axis=1)

    # Cap DOH values
    doh_columns = [
        'Current_DOH_All_DC', 'Current_DOH(Stock+PO)_All_DC'
    ] + [f'Current_{dc}_DOH' for dc in dc_list] + [f'{dc}_DOH(Stock+PO)' for dc in dc_list]

    for col in doh_columns:
        merged_df[col] = np.where(merged_df[col] > max_doh_value, np.inf, merged_df[col])

    # Initialize cover date columns
    cover_date_cols = [
        'Total_Store_cover_to_date', 'Stock_All_DC_Cover_to_date', 'Stock+PO_All_DC_Cover_to_date'
    ] + [f'{prefix}_{dc}_cover_to_date' for dc in dc_list for prefix in ['Store', 'Stock', 'Stock+PO']]

    for col in cover_date_cols:
        merged_df[col] = pd.NaT

    return merged_df, current_date, max_doh_value


# Cover date calculation function
def apply_cover_date_calculations(merged_df, current_date, max_doh_value):
    cover_date_map = {
        'Total_Store_cover_to_date': 'Total_DOHStore',
        'Stock_All_DC_Cover_to_date': 'Current_DOH_All_DC',
        'Stock+PO_All_DC_Cover_to_date': 'Current_DOH(Stock+PO)_All_DC',
        'Store_DC1_cover_to_date': 'DC1_DOHStore',
        'Stock_DC1_cover_to_date': 'Current_DC1_DOH',
        'Stock+PO_DC1_cover_to_date': 'DC1_DOH(Stock+PO)',
        'Store_DC2_cover_to_date': 'DC2_DOHStore',
        'Stock_DC2_cover_to_date': 'Current_DC2_DOH',
        'Stock+PO_DC2_cover_to_date': 'DC2_DOH(Stock+PO)',
        'Store_DC4_cover_to_date': 'DC4_DOHStore',
        'Stock_DC4_cover_to_date': 'Current_DC4_DOH',
        'Stock+PO_DC4_cover_to_date': 'DC4_DOH(Stock+PO)'
    }

    def valid_doh_to_date(doh_value):
        return pd.notnull(doh_value) and doh_value > 0 and not np.isinf(doh_value) and doh_value <= max_doh_value

    for index, row in merged_df.iterrows():
        for target_col, source_col in cover_date_map.items():
            doh_value = row[source_col]
            if valid_doh_to_date(doh_value):
                merged_df.at[index, target_col] = current_date + pd.to_timedelta(doh_value, unit='d')

    return merged_df



def replace_cj_duplicates(deduplicated_df):
    """Handles duplicates 'CJ_Item' by applying custom logic for special products"""
    numeric_cols = deduplicated_df.select_dtypes(include=np.number).columns
    special_products = ['20000408','20009203','20014191','20023778','20023779','20028264','20039186']

    def apply_logic(group):
        is_special_product = group['CJ_Item'].iloc[0] in special_products
        dc4_columns = [col for col in numeric_cols if 'DC4' in col]

        if is_special_product and len(group) == 1:
            return group
        
        if is_special_product: # if special product replace value with 0 for first rows contains 'DC4' in column name
            for col in numeric_cols:
                if col in dc4_columns or col == 'PC_Cartons':
                    continue
                if group[col].nunique() == 1:
                    group.iloc[1:, group.columns.get_loc(col)] = 0

            for col in dc4_columns:
                if group[col].nunique() == 1:
                    group.iloc[0, group.columns.get_loc(col)] = 0
        
        return group
    
    deduplicated_df = deduplicated_df.groupby('CJ_Item', group_keys=False).apply(apply_logic).reset_index(drop=True)
    return deduplicated_df



def generate_full_stock_report(
        cj_stock_df, daily_stock_dc_df, daily_so_df,
        po_foods_nf_pcb_df, po_import_df, master_leadtime_df, excel_datasets):
    
    """Generate the full stock report from processed dataframes"""
    st.header("🔍 Step 1: Validating Input File...")

    # Validate required inputs
    if any(df is None or df.empty for df in [
        cj_stock_df, daily_stock_dc_df, daily_so_df, 
        po_foods_nf_pcb_df, po_import_df, master_leadtime_df
    ]):
        st.error("One or more input DataFrames are missing or empty. Please check your uploads.")
        return None, None
    
    if not excel_datasets or 'product_list' not in excel_datasets:
        st.error("Sheet name 'product_list' is missing from the upload.")
        return None, None

    access_product_list_df = excel_datasets['product_list'].copy()
    st.success("✅ All input data validated successfully.")


    # Prepare dictionary of DataFrames
    processed_dataframes = {
        'CJ_Stock': cj_stock_df,
        'Daily_Stock_DC': daily_stock_dc_df,
        'Daily_SO': daily_so_df,
        'PO_Foods_NF_PCB': po_foods_nf_pcb_df,
        'PO_Import': po_import_df
    }

    st.header("Step 2: Transforming data...")
    try:
        # Ensure all SKU columns are string
        processed_dataframes, access_product_list_df = convert_cj_item_to_string(
            processed_dataframes, access_product_list_df
        )
        # Merge all dataframes
        merged_df = merge_dataframes(processed_dataframes, access_product_list_df)
        if merged_df is None or merged_df.empty:
            st.error("Merging failed. Check individual dataframes.")
            return None, None

        # Merge with master owner and lead time
        leadtime_data = master_leadtime_df[['SHM_Item', 'CJ_Item', 'OwnerSCM', 'LeadTime']].copy()
        merged_df = pd.merge(
            merged_df, leadtime_data,
            on=['SHM_Item'], how='left', suffixes=('', '_from-LT')
        )
        # Fill na for ownerSCM and Leadtime
        merged_df['OwnerSCM'] = merged_df['OwnerSCM'].fillna('No data')
        merged_df['LeadTime'] = merged_df['LeadTime'].fillna('No data')
        

        merged_df = fill_na_with_zero(merged_df)
        merged_df = calculate_totals(merged_df)

        merged_df, current_date, max_doh_value = apply_doh_calculations(merged_df)
        merged_df = apply_cover_date_calculations(merged_df, current_date, max_doh_value)

        st.success("✅ Transforming data completed.")

    except Exception as e:
        st.error(f"❌ Error during report generation: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None, None
    

    st.header("Step 3: Finalizing the daily stock report...")
    # Step 3: Execute query to get the final report
    try:
        query = """
        SELECT Division_SHM,
            OwnerSCM,
            [Supplier Name],
            SHM_Item,
            CJ_Item,
            Name,
            Category,
            Brand,
            LeadTime,
            Status,
            [Group],
            Unit,
            PC_Cartons,
            First_SO_Date,
            NPD_Status,
            Total_ScmAssort,
            Total_OOSAssort,
            Total_CountOKROOS,
            Total_PercOOS,
            Total_StoreStockQty,
            Total_DOHStore,
            Total_Store_cover_to_date,
            Total_AvgSaleQty90D,
            Total_AvgSaleQty30D,
            Total_AvgSaleQty7D,
            SO_Qty_last7D,
            Remain_StockQty_AllDC,
            Current_DOH_All_DC,
            Stock_All_DC_Cover_to_date,
            [Total-PO_qty_to_DC],
            Min_delivery_date_to_DC,
            [Current_DOH(Stock+PO)_All_DC],
            [Stock+PO_All_DC_Cover_to_date],
            DC1_ScmAssort,
            DC1_OOSAssort,
            DC1_CountOKROOS,
            DC1_PercOOS,
            DC1_StoreStockQty,
            DC1_DOHStore,
            Store_DC1_cover_to_date,
            DC1_AvgSaleQty90D,
            DC1_AvgSaleQty30D,
            DC1_AvgSaleQty7D,
            [%Ratio_AvgSalesQty90D_DC1],
            DC1_Remain_StockQty,
            Current_DC1_DOH,
            Stock_DC1_cover_to_date,
            [Total-PO_qty_to_DC1],
            Min_delivery_date_to_DC1,
            [DC1_DOH(Stock+PO)],
            [Stock+PO_DC1_cover_to_date],
            DC2_ScmAssort,
            DC2_OOSAssort,
            DC2_CountOKROOS,
            DC2_PercOOS,
            DC2_StoreStockQty,
            DC2_DOHStore,
            Store_DC2_cover_to_date,
            DC2_AvgSaleQty90D,
            DC2_AvgSaleQty30D,
            DC2_AvgSaleQty7D,
            [%Ratio_AvgSalesQty90D_DC2],
            DC2_Remain_StockQty,
            Current_DC2_DOH,
            Stock_DC2_cover_to_date,
            [Total-PO_qty_to_DC2],
            Min_delivery_date_to_DC2,
            [DC2_DOH(Stock+PO)],
            [Stock+PO_DC2_cover_to_date],
            DC4_ScmAssort,
            DC4_OOSAssort,
            DC4_CountOKROOS,
            DC4_PercOOS,
            DC4_StoreStockQty,
            DC4_DOHStore,
            Store_DC4_cover_to_date,
            DC4_AvgSaleQty90D,
            DC4_AvgSaleQty30D,
            DC4_AvgSaleQty7D,
            [%Ratio_AvgSalesQty90D_DC4],
            DC4_Remain_StockQty,
            Current_DC4_DOH,
            Stock_DC4_cover_to_date,
            [Total-PO_qty_to_DC4],
            Min_delivery_date_to_DC4,
            [DC4_DOH(Stock+PO)],
            [Stock+PO_DC4_cover_to_date]
        FROM merged_df
        WHERE ([Group] != 'Discontinuous' or [Group] IS NULL)
        ORDER BY CJ_Item ASC
        """
        result_df = sqldf(query, locals())

        columns_to_rename = {
            'Status': 'CJ_Status',
            'Group': 'SHM_Status', 
            'Unit': 'Unit_of_Purchase', 
            'LeadTime':'LeadTime(Days)',
            'Total_OOSAssort':'Total_ActiveAssort',
            'Total_CountOKROOS':'Total_StoreOOS',
            'Total_PercOOS':'Total_%StoreOOS'
        }
        result_df.rename(columns=columns_to_rename, inplace=True)

        # Remove rows where all numeric columns are zero
        numeric_cols = result_df.select_dtypes(include=np.number).columns
        result_df = result_df[(result_df[numeric_cols] != 0).any(axis=1)]

        # Format the date columns
        date_columns = [
            'First_SO_Date',
            'Min_delivery_date_to_DC',
            'Min_delivery_date_to_DC1',
            'Min_delivery_date_to_DC2',
            'Min_delivery_date_to_DC4',
            'Total_Store_cover_to_date',
            'Store_DC1_cover_to_date',
            'Store_DC2_cover_to_date',
            'Store_DC4_cover_to_date',
            'Stock_All_DC_Cover_to_date',
            'Stock_DC1_cover_to_date',
            'Stock_DC2_cover_to_date',
            'Stock_DC4_cover_to_date',
            'Stock+PO_All_DC_Cover_to_date',
            'Stock+PO_DC1_cover_to_date',
            'Stock+PO_DC2_cover_to_date',
            'Stock+PO_DC4_cover_to_date'
        ]

        for col in date_columns:
            if col in result_df.columns:
                result_df[col] = pd.to_datetime(result_df[col], errors='coerce')

        # transform data when CJ_Item is duplicated
        result_df = replace_cj_duplicates(result_df)

        # Rename column for converting QTY to BOX QTY
        column_rename_mapping = {
            'Total_AvgSaleQty90D': 'Total_AvgSaleCTN_Last90D',
            'Total_AvgSaleQty30D': 'Total_AvgSaleCTN_Last30D',
            'Total_AvgSaleQty7D': 'Total_AvgSaleCTN_Last7D',
            'DC1_AvgSaleQty90D': 'DC1_AvgSaleCTN_Last90Days',
            'DC1_AvgSaleQty30D': 'DC1_AvgSaleCTN_Last30Days',
            'DC1_AvgSaleQty7D': 'DC1_AvgSaleCTN_Last7Days',
            'DC2_AvgSaleQty90D': 'DC2_AvgSaleCTN_Last90Days',
            'DC2_AvgSaleQty30D': 'DC2_AvgSaleCTN_Last30Days',
            'DC2_AvgSaleQty7D': 'DC2_AvgSaleCTN_Last7Days',
            'DC4_AvgSaleQty90D': 'DC4_AvgSaleCTN_Last90Days',
            'DC4_AvgSaleQty30D': 'DC4_AvgSaleCTN_Last30Days',
            'DC4_AvgSaleQty7D': 'DC4_AvgSaleCTN_Last7Days',
            'Total_StoreStockQty': 'Total_StoreStockCTN',
            'DC1_StoreStockQty': 'DC1_StoreStockCTN',
            'DC2_StoreStockQty': 'DC2_StoreStockCTN',
            'DC4_StoreStockQty': 'DC4_StoreStockCTN',
            'SO_Qty_last7D': 'SO_CTN_last7D',
            'Remain_StockQty_AllDC': 'Remain_CTN_AllDC',
            'DC1_Remain_StockQty': 'DC1_Remain_CTN',
            'DC2_Remain_StockQty': 'DC2_Remain_CTN',
            'DC4_Remain_StockQty': 'DC4_Remain_CTN',
            'Total-PO_qty_to_DC': 'Total-CTN_to_DC',
            'Total-PO_qty_to_DC1': 'Total-CTN_to_DC1',
            'Total-PO_qty_to_DC2': 'Total-CTN_to_DC2',
            'Total-PO_qty_to_DC4': 'Total-CTN_to_DC4'
        }

        # Process the rename
        modified_df = result_df.rename(columns=column_rename_mapping)

        # Rename for sheet data by CTN
        renamed_columns_list = list(column_rename_mapping.values())

        for col in renamed_columns_list:
            if col in modified_df.columns:
                modified_df[col] = np.where(
                    modified_df['PC_Cartons'] == 0,0,
                    modified_df[col] / modified_df['PC_Cartons']
                )

        return result_df, modified_df

    except Exception as e:
        st.error(f"❌ Error in Step 3 (Finalizing report): {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None, None
    

def convert_dfs_to_multi_sheet_excel_bytes(dataframes_with_names):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in dataframes_with_names.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    processed_data = output.getvalue()
    return processed_data



# Main processing section
if st.button("🚀 Generate Daily Stock Report", type="primary", use_container_width=True):
    # Check required files
    required_files = [
        master_file, dc_stock_file, sellout_file,
        access_db_extracted_file, po_import_overseas, master_leadtime_file
    ]
    if all(file is not None for file in required_files):
        with st.spinner("Processing data... This may take a few minutes."):
            try:
                progress_bar = st.progress(0)
                status_text = st.empty()

                # ---------------- Step 1: CJ Stock ----------------
                status_text.text("Step 1/7: Processing CJ Stock data...")
                progress_bar.progress(14)
                cj_stock_df = process_cj_stock(master_file)
                st.success("✅ Step 1: CJ Stock Data File has been completely processed")


                # ---------------- Step 2: DC Stock ----------------
                status_text.text("Step 2/7: Processing DC Stock data...")
                progress_bar.progress(29)
                dc_stock_df = process_dc_stock(dc_stock_file)
                st.success("✅ Step 2: DC Stock Data File has been completely processed")


                # ---------------- Step 3: Sell-Out ----------------
                status_text.text("Step 3/7: Processing Daily Sell-Out data...")
                progress_bar.progress(42)
                sellout_df = process_sellout_data(sellout_file)
                st.success("✅ Step 3: Sell out past 30 days has been completely processed")


                # ---------------- Step 4: PO Foods & NF & PCB ----------------
                status_text.text("Step 4/7: Processing PO Pending Foods and NF...")
                progress_bar.progress(57)
                access_datasets = process_all_data_from_access(access_db_extracted_file)
                po_other_df, po_other_raw = process_po_other(access_datasets)
                st.success("✅ Step 4: PO pending Foods and NF has been completely processed")


                # ---------------- Step 6: PO Import ----------------
                status_text.text("Step 5/7: Processing PO Pending Import...")
                progress_bar.progress(71)
                po_import_df, po_import_raw = process_po_import(access_datasets, po_import_overseas)
                st.success("✅ Step 5: PO pending Import has been completely processed")


                # ---------------- Step 7: Lead Time ----------------
                status_text.text("Step 6/7: Processing Master Lead Time data...")
                progress_bar.progress(85)
                master_leadtime_df = process_leadtime(master_leadtime_file)
                st.success("✅ Step 6: Master Lead time has been completely processed")


                # ---------------- Step 8: Final Report ----------------
                status_text.text("Step 7/7: Generating final report...")
                progress_bar.progress(95)
                result_df, modified_df = generate_full_stock_report(
                    cj_stock_df,
                    dc_stock_df,
                    sellout_df,
                    po_other_df,
                    po_import_df,
                    master_leadtime_df,
                    access_datasets
                )

                # Combine all PO data
                final_df, final_df_pivot = combine_all_PO_data(
                    po_other_raw, 
                    po_import_raw,
                    master_leadtime_df,
                )

                # Prepare data for export
                dfs_to_export = {
                    'Data by Qty': result_df.copy(),
                    'Data by CTN': modified_df.copy(),
                    'All PO Pending': final_df.copy(),
                    'MIN ETA': final_df_pivot.copy()
                }

                # Generate Excel file
                current_date = datetime.now().strftime('%d-%m-%Y')
                excel_bytes_multi_sheet = convert_dfs_to_multi_sheet_excel_bytes(dfs_to_export)
                st.success(f"✅ Step 8: Daily Stock Report has been generated with {len(result_df)} rows of data")
                progress_bar.progress(100)
                status_text.text("✅ All steps completed successfully!")
                st.success("✅ The daily stock report is ready for download.")

                # Final Download Button
                st.download_button(
                    label="📥 Download Final Stock Report (Excel)",
                    data=excel_bytes_multi_sheet,
                    file_name=f"Sahamit_Daily_Stock_Report_{current_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )

            except Exception as e:
                progress_bar.progress(0)
                status_text.text("Error occurred during processing")
                st.error(f"❌ Error processing data: {str(e)}")
                st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p><strong>Sahamit Daily Stock Report Generator</strong> - Powered by Streamlit<br>
    Modified by Thanawit.C for generate daily stock report as an Excel file only</p>
</div>
""", unsafe_allow_html=True)
