import pandas as pd
import glob
import numpy as np
import time
import datetime as dt
import re
import streamlit as st
from io import BytesIO
from xlsxwriter import Workbook
import pytz
import requests
import os

def download_file_from_github(url, save_path):
    response = requests.get(url)
    if response.status_code == 200:
        with open(save_path, 'wb') as file:
            file.write(response.content)
        print(f"File downloaded successfully and saved to {save_path}")
    else:
        print(f"Failed to download file. Status code: {response.status_code}")

def load_excel(file_path):
    with open(file_path, 'rb') as file:
        model = pd.read_excel(file, engine='openpyxl')
    return model

st.title('Promix')
uploaded_file = st.file_uploader("Upload File", type="xlsx", accept_multiple_files=False)

def get_current_time_gmt7():
    tz = pytz.timezone('Asia/Jakarta')
    return dt.datetime.now(tz).strftime('%Y%m%d_%H%M%S')
    
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Mengakses workbook dan worksheet untuk format header
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Menambahkan format khusus untuk header
        header_format = workbook.add_format({'border': 0, 'bold': False, 'font_size': 12})
        
        # Menulis header manual dengan format khusus
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
    processed_data = output.getvalue()
    return processed_data
            
if uploaded_file is not None:
    #st.write('File berhasil diupload')
    # Baca konten zip file

    if st.button('Process'):
        with st.spinner('Data sedang diproses...'):
            df_promix = pd.read_excel(uploaded_file,header=1)
            df_cab = pd.read_excel(uploaded_file,header=2).dropna(subset=df_promix.iloc[0,0]).iloc[:,:5].drop_duplicates()
            df_promix = df_promix.T
            df_promix[0] = df_promix[0].ffill()
            df_promix = df_promix.reset_index()
            df_promix['index'] = df_promix['index'].apply(lambda x: np.nan if 'Unnamed' in str(x) else x).ffill()
            df_promix.columns = df_promix.loc[0,:].fillna('')
            df_promix = df_promix.iloc[5:,:].groupby(df_promix.columns[:3].to_list())[df_promix.columns[3:]].sum().reset_index()
            df_promix = df_promix.melt(id_vars=df_promix.columns[:3], value_vars=df_promix.columns[3:])
            df_promix.columns = ['TANGGAL','NAMA BAHAN','SUMBER','CABANG','QTY']
            df_promix = df_promix.merge(df_cab,
                            how='left', left_on='CABANG', right_on=df_cab.columns[0]).drop(columns='CABANG').iloc[:,[0,4,5,6,7,8,1,2,3]]
            st.download_button(
                    label="Download Excel",
                    data=to_excel(df_promix),
                    file_name=f'promix_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
