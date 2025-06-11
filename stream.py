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

st.error("Platform web ini sudah tidak lagi menerima pembaruan dan akan dinonaktifkan pada 20/06/2025. 
Segera gunakan aplikasi versi offline (DEX) untuk mengakses fitur dan data terbaru.")

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

st.title('GIS')
selected_option = st.selectbox("Pilih salah satu:", ['13.01','13.10','13.22','13.31','13.33','13.55','13.66','22.05','22.16','22.19','32.07','32.15','32.23','32.24','32.43','41.01','41.04','41.04.B','41.09','42.02','42.05','42.06','42.08','42.15','42.17','42.17 (Rev1)','42.18','44.06','44.08','51.01','99.01'])
uploaded_file = st.file_uploader("Upload File", type="xlsx", accept_multiple_files=True)

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
            
            if selected_option=='13.01':
                concatenated_df = []
                for file in uploaded_file:
                    df_1301 = pd.read_excel(file, skiprows=4)
                    df_1301 = df_1301.loc[:, ~df_1301.columns.str.startswith('Unnamed')]
                    df_1301['Keterangan'] = df_1301['Keterangan'].bfill()
                    df_1301 = df_1301.iloc[1:,:].iloc[:-6,][(df_1301['Keterangan']!='Keterangan')].reset_index(drop=True)
                    df_1301['Tanggal']  =  pd.to_datetime(df_1301['Tanggal'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d/%m/%Y')
                    for i in df_1301[df_1301['Tanggal'].isna()].index:
                        df_1301.loc[i-1,'Keterangan'] = df_1301.loc[i-1,'Keterangan']+df_1301.loc[i,'Keterangan']
                    df_1301 = df_1301[~df_1301['Tanggal'].isna()]
                    concatenated_df.append(df_1301)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'13.01_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='13.10':
                concatenated_df = []
                for file in uploaded_file:
                    df_1310 = pd.read_excel(file, skiprows=4).fillna('')
                    df_1310z = df_1310.iloc[:-5]
                    df_1310 = df_1310z.loc[:, ~df_1310z.columns.str.startswith('Unnamed:')]
                    concatenated_df.append(df_1310)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'13.10_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='13.22':
                concatenated_df = []
                for file in uploaded_file:
                    df_1322 = pd.read_excel(file, skiprows=4).fillna('')
                    df_1322 = df_1322[~df_1322['Tanggal'].isna()][[x for x in df_1322.columns if 'Unnamed' not in x]]
                    concatenated_df.append(df_1322)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'13.22_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )  

            if selected_option=='13.31':
                concatenated_df = []
                for file in uploaded_file:
                    df_1331 = pd.read_excel(file, skiprows=4).fillna('')
                    df_1331 = df_1331[~df_1331['Tanggal'].isna()][[x for x in df_1331.columns if 'Unnamed' not in x]]
                    concatenated_df.append(df_1331)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'13.31_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )  
                
            if selected_option=='13.33':
                concatenated_df = []
                for file in uploaded_file:
                    df_1333 = pd.read_excel(file, skiprows=4).fillna('')
                    df_1333 = df_1333[~df_1333['Tanggal'].isna()][[x for x in df_1333.columns if 'Unnamed' not in x]]
                    concatenated_df.append(df_1333)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'13.33_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='13.55':
                concatenated_df = []
                for file in uploaded_file:
                    df_1355 = pd.read_excel(file, skiprows=4).fillna('')
                    df_1355 = df_1355[~df_1355['Tanggal'].isna()][[x for x in df_1355.columns if 'Unnamed' not in x]]
                    concatenated_df.append(df_1355)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'13.55_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
            
            if selected_option=='13.66':
                concatenated_df = []
                for file in uploaded_file:
                    df_1366 = pd.read_excel(file, header=4).fillna('')
                    df_1366 = df_1366.iloc[:-5]
                    df_1366 = df_1366.loc[:, ~df_1366.columns.str.startswith('Unnamed:')]
                    df_1366['Tanggal']           =   pd.to_datetime(df_1366['Tanggal'], format='%Y/%m/%d').dt.strftime('%d/%m/%Y')
                    df_1366['Debit']    =   pd.to_numeric(df_1366['Debit'])
                    df_1366['Kredit']   =   pd.to_numeric(df_1366['Kredit'])
                    df_1366['Hari']     =   pd.to_numeric(df_1366['Hari'])
                    concatenated_df.append(df_1366)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'13.66_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
            
            
            if selected_option=='22.05':
                concatenated_df = []
                for file in uploaded_file:
                    df_2205 = pd.read_excel(file, skiprows=4).fillna('')
                    df_2205 = df_2205.iloc[:-5]
                    df_2205 = df_2205.loc[:, ~df_2205.columns.str.startswith('Unnamed:')]
                    concatenated_df.append(df_2205)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'22.05_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
                
            if selected_option=='22.16':
                concatenated_df = []
                for file in uploaded_file:
                    df_2216 = pd.read_excel(file)#.fillna('')
                    
                    # Drop the first four columns
                    df_2216 = df_2216.loc[:df_2216.shape[0]-2,df_2216.columns.tolist()[:2] + df_2216.columns.tolist()[4:]].drop(columns=df_2216.columns[-2])
                    dfs = []
                    x=0
                    for i in range(1,int((len(df_2216.columns)-2)/4)):
                        df = df_2216.loc[:,df_2216.columns[:2].tolist()[:2]+df_2216.columns[2+x:6+x].tolist()]
                        df.columns = df.iloc[3,:2].values.tolist() + df.iloc[4,2:].values.tolist()
                        df['Bulan'] = df.iloc[3,2]
                        df = df.iloc[5:,]
                        x =+4
                        dfs.append(df)
                    
                    concatenated_df.append(pd.concat(dfs,ignore_index=True))
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'22.16_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   


            if selected_option=='22.19':
                concatenated_df = []
                for file in uploaded_file:
                    df_2219 = pd.read_excel(file).fillna('')
    
                    # Drop the first four columns
                    df_2219 = df_2219.iloc[:, 2:]
                    # Drop the first three rows
                    df_2219 = df_2219.iloc[3:, :]
                    # Reset the index (optional, if you want a clean index)
                    df_2219.reset_index(drop=True, inplace=True)
                    
                    # Set the first row as the header
                    df_2219.columns = df_2219.iloc[0]  # Set the first row as the column headers
                    df_2219 = df_2219.drop(df_2219.index[0])  # Drop the first row now that it's the header
                    
                    # Reset the index again (optional, if you want a clean index)
                    df_2219.reset_index(drop=True, inplace=True)
                    # Fill the blank "Pelanggan" cells with the preceding value
                    df_2219['Nama Cabang'] = df_2219['Nama Cabang'].replace('', None).ffill()
                    
                    # Fill the blank "Pelanggan" cells with the preceding value
                    df_2219['Pelanggan'] = df_2219['Pelanggan'].replace('', None).ffill()
                    
                    # Convert "Tgl. SI #" column to datetime format
                    df_2219['Tgl. SI #'] = pd.to_datetime(df_2219['Tgl. SI #'], format='%d/%m/%Y')
                    
                    # Format "Total" column as numbers (assuming they are stored as strings)
                    df_2219['Total'] = pd.to_numeric(df_2219['Total'])
                    
                    df_2219 =       df_2219[df_2219['Nama Cabang']      !=      "Total Nama Cabang"]
                    concatenated_df.append(df_2219)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'22.19_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='32.07':
                concatenated_df = []
                for file in uploaded_file:
                    df  = pd.read_excel(file,header=1).fillna('')
                    
                    # Find the indices of start and end markers
                    start_indices = df[df.apply(lambda row: 'Cabang :' in str(row.values), axis=1)].index
                    end_indices = df[df.apply(lambda row: 'ACCURATE Accounting System Report' in str(row.values), axis=1)].index
                    
                    # Concatenate sections into a single DataFrame
                    df_3207 = pd.concat([df.iloc[start_idx+1:end_idx] for start_idx, end_idx in zip(start_indices, end_indices)], ignore_index=True)
                    
                    # Remove existing header
                    df_3207.columns = range(df_3207.shape[1])
                    
                    # Set the second row as the new header
                    new_header = df_3207.iloc[0]
                    df_3207 = df_3207[1:]
                    df_3207.columns = new_header
                    
                    # Delete the blank Column
                    df_3207 = df_3207.loc[:,['Nomor # PR','Tanggal # PR','Nomor # PO','Tanggal # PO','Pemasok','Kode #','Nama Barang','Kuantitas','@Harga','Total Harga','Rasio Satuan','Nama Satuan','Tgl/Jam Pembuatan PO#','Tgl/Jam Pembuatan PR#']]
                    
                    # Drop Unnecessary Column
                    df_3207 = df_3207[df_3207['Nomor # PO']     !=      'Nomor # PO']
                    df_3207 = df_3207[df_3207['Nomor # PO']     !=      '']
                    
                    
                    # Reset the index
                    df_3207.reset_index(drop=True, inplace=True)
                    
                    df_3207['Tanggal # PR']           =   pd.to_datetime(df_3207['Tanggal # PR'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y')
                    df_3207['Tanggal # PO']           =   pd.to_datetime(df_3207['Tanggal # PO'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y')
                    df_3207['Tgl/Jam Pembuatan PO#']  =   pd.to_datetime(df_3207['Tgl/Jam Pembuatan PO#'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')
                    df_3207['Tgl/Jam Pembuatan PR#']  =   pd.to_datetime(df_3207['Tgl/Jam Pembuatan PR#'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')
                    concatenated_df.append(df_3207)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'32.07_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        
            if selected_option=='32.15':
                concatenated_df=[]
                for file in uploaded_file:
                    # Read the Excel file, skipping blank lines and using the second row as the header
                    df = pd.read_excel(file, header=1).fillna('')
                
                    # Initialize lists to store extracted dataframes
                    extracted_dfs = []
                
                    # Find the indices of start and end markers
                    start_indices = df[df.apply(lambda row: 'Permintaan Barang' in str(row.values), axis=1)].index
                    end_indices = df[df.apply(lambda row: 'ACCURATE Accounting System Report' in str(row.values), axis=1)].index
                
                    # Loop through each pair of start and end indices
                    for start_idx, end_idx in zip(start_indices, end_indices):
                        # Extract the desired range of rows
                        selected_rows = df.loc[start_idx:end_idx-1]
                
                        # Remove existing header
                        selected_rows.columns = range(selected_rows.shape[1])
                
                        # Set the first row as the new header
                        new_header = selected_rows.iloc[0]
                        selected_rows = selected_rows[1:]
                        selected_rows.columns = new_header
                
                        # Delete the blank Column
                        selected_rows = selected_rows.loc[:,['Permintaan Barang', 'Pesanan Pembelian', 'Penerimaan Barang', 'Uang Muka Pembelian', 'Faktur Pembelian', 'Retur Pembelian', 'Pembayaran Pembelian']]
                
                        # Drop Unnecessary Column
                        selected_rows = selected_rows[selected_rows['Permintaan Barang'] != 'Permintaan Barang']
                        selected_rows = selected_rows[selected_rows['Permintaan Barang'] != '']
                
                        # Reset the index
                        selected_rows.reset_index(drop=True, inplace=True)
                
                        # Append the selected DataFrame to the list
                        extracted_dfs.append(selected_rows)
                
                    # Concatenate all extracted dataframes
                    concatenated_df.append(pd.concat(extracted_dfs, ignore_index=True))
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'32.15_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )          

            if selected_option=='32.23':
                concatenated_df = []
                for file in uploaded_file:
                    df_3223 = pd.read_excel(file, header=4).fillna('')
                    df_3223 = df_3223.iloc[:-5]
                    dfDPB       =       df_3223.loc[:,["Nama Cabang",
                                         "Nomor #",
                                         "Tanggal",
                                         "Tgl/Jam Pembuatan",
                                         "Pemasok",
                                         "Pengiriman"]].rename(columns={'Nomor #':'Nomor'}).fillna("")
                    dfDPB['Tanggal']                = pd.to_datetime(dfDPB['Tanggal'], format='%Y-%m-%d')
                    dfDPB['Tgl/Jam Pembuatan']      = pd.to_datetime(dfDPB['Tgl/Jam Pembuatan'], format='%Y-%m-%d %H:%M:%S')
                    concatenated_df.append(dfDPB)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'32.23_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='32.24':
                concatenated_df = []
                for file in uploaded_file:
                    df_3224 = pd.read_excel(file, skiprows=range(0, 4))
                    df_3224 = df_3224.loc[:,['Tanggal','Nomor # PO','Nomor # RI','Pemasok','Kode #','Nama Barang','Kts Terima','Satuan','@Harga','Total Harga','#Kts Ditagih','Nama Gudang','Nama Cabang Penerimaan Barang','Status Penerimaan Barang','Pembuat Data','Tgl/Jam Pembuatan']]
                    df_3224 = df_3224[df_3224['Nomor # PO' ] != ""]
                    concatenated_df.append(df_3224)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'32.24_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='32.43':
                concatenated_df = []
                for file in uploaded_file:
                    df_3243 = pd.read_excel(file, skiprows=4).fillna('')
                    df_3243 = df_3243[~df_3243['Tanggal'].isna()][[x for x in df_3243.columns if 'Unnamed' not in x]]
                    concatenated_df.append(df_3243)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'32.43_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )  
                
            if selected_option=='41.01':
                concatenated_df = []
                for file in uploaded_file:
                    df_4101 =   pd.read_excel(file, header=4)
                    df_4101['Nama Cabang'] = df_4101['Nama Cabang'].astype(str)
                    
                    data_remove = ["Nama Cabang", "GiS", "#41.01", "Dari", "Cabang"]
                    df_4101a = df_4101[~df_4101['Nama Cabang'].str.startswith(tuple(data_remove))]
                    
                    df_4101a['Nama Cabang'] = df_4101a['Nama Cabang'].replace('nan', "")
                    
                    df_4101a['Keterangan'] = df_4101a['Keterangan'].fillna("").astype(str)
                    
                    df_4101a['Keterangan'] = df_4101a.groupby((df_4101a['Nomor #'].notna()).cumsum())['Keterangan'].transform(lambda x: ' '.join(x))
                    
                    df_4101a = df_4101a.loc[:, ~df_4101.columns.str.startswith('Unnamed')]
                    
                    df_4101a = df_4101a[df_4101a['Nama Cabang'] != ""]
                    
                    df_4101a['Tanggal'] = pd.to_datetime(df_4101a['Tanggal'], format='%d/%m/%Y %H:%M:%S')
                    
                    df_4101a['Tanggal'] = df_4101a['Tanggal'].dt.strftime('%d/%m/%Y')
                    concatenated_df.append(df_4101a)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'41.01_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
                
            if selected_option=='41.04':
                concatenated_df = []
                for file in uploaded_file:
                    db_4104 = pd.read_excel(file, header=3)
                    db_4104 = db_4104.loc[:,['Unnamed: 2','Unnamed: 4']]
                    db_4104 = db_4104.dropna(subset=['Unnamed: 2'])

                    kode_barang = []
                    nama_barang = []
                    
                    for i in range(0, len(db_4104), 2):
                        kode_barang.append(db_4104.iloc[i, 1])
                        nama_barang.append(db_4104.iloc[i+1, 1])
                        
                    db = pd.DataFrame({"Kode Barang": kode_barang, "Nama Barang": nama_barang})
                    
                    df_4104 = pd.read_excel(file, header=6).fillna('')
                    df_4104 = df_4104.loc[:, ~df_4104.columns.str.contains('^Unnamed')]
                    
                    df_4104 = df_4104[~df_4104['Nomor #'].isin(['Nomor #', ''])]
                    df_4104 = df_4104.dropna(subset=['Nomor #'])

                    df_4104 = pd.merge(df_4104, db, how='left', on='Nama Barang').fillna('Cek')
                    df_4104 = df_4104.loc[:,['Nama Gudang', 'Kode Barang', 'Nama Barang', 'Nomor #', 'Tanggal', 'Kts Masuk', 'Nilai Masuk/Sat', 'Nilai Masuk', 'Kts Keluar', 'Nilai Keluar/Sat', 'Nilai Keluar', 'Kts Akhir', 'Nilai Akhir']]
                    
                    concatenated_df.append(df_4104)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'41.04_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='41.04.B':
                concatenated_df = []
                for file in uploaded_file:
                    df_4104b = pd.read_excel(file,header=4)
                    df_4104b = df_4104b.dropna(how='all',axis=1)
                    df_4104b = df_4104b.iloc[1:-1,:]
                    df_db = pd.DataFrame({'variable':[x for x in df_4104b.columns if 'Unnamed' in x],'Nama Barang':[x for x in df_4104b.columns if 'Unnamed' not in x][1:]})
                    df_4104b = df_4104b.melt(id_vars='Nama Gudang', value_vars=[x for x in df_4104b.columns if 'Unnamed' not in x][1:],value_name='Nominal Kts Keluar', var_name='Nama Barang').merge(
                    df_4104b.melt(id_vars='Nama Gudang', value_vars=[x for x in df_4104b.columns if 'Unnamed' in x],value_name='Kts Keluar').merge(df_db,how='left',on='variable'), how='left', on=['Nama Gudang','Nama Barang']).drop(columns='variable')
                    concatenated_df.append(df_4104b)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'41.04.B_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                ) 
            
            if selected_option=='41.09':
                concatenated_df = []
                for file in uploaded_file:
                    df_4109 = pd.read_excel(file, header=4)#.fillna('')
                    df_4109 = df_4109.dropna(axis=1,how='all')
                    df_4109.columns = (df_4109.T.reset_index()['index'].apply(lambda x: np.nan if 'Unnamed' in x else x).ffill()+df_4109.T.reset_index().iloc[:,1].fillna('')).str.replace('Saldo Awal','').str.replace('Saldo Akhir','').str.replace('Masuk','').str.replace('Keluar','')
                    
                    df_4109 = df_4109[~df_4109['Kode Barang'].isna()]
                    df_4109['Kategori Barang'] = df_4109['Kategori Barang'].ffill()
                    
                    df_4109 = df_4109.iloc[:,[0,1,2,3,4,5]].melt(id_vars=['Kategori Barang','Nama Barang','Kode Barang','Nama Satuan'],
                                                           value_vars=['Kuantitas','Nilai'],value_name='Saldo Awal',var_name='Konversi'
                                                           ).merge(df_4109.iloc[:,[0,1,2,3,6,7]].melt(id_vars=['Kategori Barang','Nama Barang','Kode Barang','Nama Satuan'],
                                                           value_vars=['Kuantitas','Nilai'],value_name='Masuk',var_name='Konversi'), on=df_4109.columns[:4].to_list() + ['Konversi']
                                                           ).merge(df_4109.iloc[:,[0,1,2,3,8,9]].melt(id_vars=['Kategori Barang','Nama Barang','Kode Barang','Nama Satuan'],
                                                           value_vars=['Kuantitas','Nilai'],value_name='Keluar',var_name='Konversi'), on=df_4109.columns[:4].to_list() + ['Konversi']
                                                           ).merge(df_4109.iloc[:,[0,1,2,3,10,11]].melt(id_vars=['Kategori Barang','Nama Barang','Kode Barang','Nama Satuan'],
                                                           value_vars=['Kuantitas','Nilai'],value_name='Saldo Akhir',var_name='Konversi'), on=df_4109.columns[:4].to_list() + ['Konversi']).sort_values(['Kategori Barang','Nama Barang','Konversi'])
                    concatenated_df.append(df_4109)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'41.09_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            if selected_option=='42.02':
                concatenated_df = []
                for file in uploaded_file:
                    file_name = os.path.basename(file.name)
                    
                    match = re.search(r'_(\d{4}\.[A-Z]+)', file_name)
                    cabang = match.group(1) if match else ''
                    
                    df_4202 = pd.read_excel(file, header=4).fillna('')
                    
                    df_4202 = df_4202.loc[:, ~df_4202.columns.str.startswith('Unnamed')]
                    
                    df_4202['Cabang'] = cabang                    
                    concatenated_df.append(df_4202)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'42.02_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            if selected_option=='42.05':
                concatenated_df = []
                for file in uploaded_file:
                    df_4205 = pd.read_excel(file, header=4).fillna('')
                    df_4205 = df_4205.iloc[:-5]
                    df_4205 = df_4205.drop(columns=['Unnamed: 0'])
                    
                    # Rename columns with names like "Unnamed: 1", "Unnamed: 2", etc. to empty strings
                    #df_4205.rename(columns=lambda x: '' if 'Unnamed' in x else x, inplace=True)
                    df_4205 = df_4205.loc[:, ~df_4205.columns.str.startswith('Unnamed')]
                    df_4205['Tanggal #Kirim']           =   pd.to_datetime(df_4205['Tanggal #Kirim'], format='%d-%b-%y').dt.strftime('%d %b %Y')
                    df_4205['Tanggal #Terima']          =   pd.to_datetime(df_4205['Tanggal #Terima'], format='%d-%b-%y').dt.strftime('%d %b %Y')
                    df_4205['#Tgl/Jam Pembuatan RI']    =   pd.to_datetime(df_4205['#Tgl/Jam Pembuatan RI'], format='%d-%b-%Y %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')
                    concatenated_df.append(df_4205)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'42.05_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
            if selected_option=='42.06':
                concatenated_df = []
                for file in uploaded_file:
                    df_4206     =   pd.read_excel(file, header=3).fillna('')
                    df_4206.columns = [f'Unnamed: {i}' for i in range(len(df_4206.columns))]
                    df_4206     =   df_4206.drop(df_4206.columns[[0, 0]], axis=1)
                    
                    df_4206 = df_4206.rename(columns={
                        'Unnamed: 1': 'Kode Barang',
                        'Unnamed: 5': 'Nama Barang',
                        'Unnamed: 7': 'Nama Gudang',
                        'Unnamed: 11': 'Nomor #',
                        'Unnamed: 13': 'Tanggal',
                        'Unnamed: 15': 'Deskripsi',
                        'Unnamed: 18': 'Keterangan Transaksi',
                        'Unnamed: 20': 'Satuan',
                        'Unnamed: 22': 'Masuk',
                        'Unnamed: 24': 'Keluar',
                        'Unnamed: 26': 'Saldo'
                    })
                    
                    df_4206     =       df_4206.loc[:, ~df_4206.columns.str.startswith('Unnamed')]
                    
                    df_4206['Nama Gudang'] = df_4206['Nama Gudang'].replace('', pd.NA)
                    df_4206['Nama Gudang'] = df_4206['Nama Gudang'].fillna(method='ffill')
                    
                    filter_strings = ['ACCURATE', 'Tercetak', 'Halaman', 'PPA', '#42.06', 'Dari', 'Kode Barang','Kode','Barang']
                    mask = ~df_4206['Kode Barang'].str.startswith(tuple(filter_strings))
                    df_4206 = df_4206[mask]
                    
                    df_4206     =       df_4206[(df_4206['Kode Barang'] != '') & ~(df_4206['Kode Barang'].str.contains('Filter'))]
                    concatenated_df.append(df_4206)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)      
                excel_data = to_excel(concatenated_df)

                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'42.06_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
                
            if selected_option=='42.08':
                concatenated_df = []
                for file in uploaded_file:
                    df_4208     =   pd.read_excel(file, header=4).fillna('')
                
                    df_4208 = df_4208.drop(df_4208.columns[[0, 1]], axis=1)
                
                    df_4208     =   df_4208.rename(columns={'Kode Barang':'Nama Barang','Unnamed: 3':'Cabang','Unnamed: 4':'Nomor #',df_4208.columns[4]:'Barang','Unnamed: 9':'Tanggal','Unnamed: 11':'Deskripsi','Unnamed: 14':'Satuan','Unnamed: 16':'Masuk','Unnamed: 18':'Keluar','Unnamed: 20':'Saldo'})
                
                    # Drop columns that start with 'Unnamed'
                    df_4208 = df_4208.loc[:, ~df_4208.columns.str.startswith('Unnamed')].drop(columns=(':'))
                
                    for i in range(len(df_4208)):
                        if df_4208['Deskripsi'][i].startswith("Saldo Barang"):
                            if ((i + 1) < len(df_4208)) and (df_4208['Tanggal'][i+1] != ""):
                                df_4208.at[i, 'Tanggal'] = 'BOTTOM'
                            else:
                                df_4208.at[i, 'Tanggal'] = 'TOP'
                
                    # Forward fill 'BOTTOM' and backward fill 'TOP'
                    df_4208['Tanggal'] = df_4208['Tanggal'].replace('BOTTOM', method='bfill').replace('TOP', method='ffill')
                
                    # Check for consecutive blank rows
                    def is_blank_row(row):
                        return all(cell == '' for cell in row)
                
                    # Track indices of rows to delete
                    rows_to_delete = []
                    consecutive_blanks = 0
                
                    for i, row in df_4208.iterrows():
                        if is_blank_row(row):
                            consecutive_blanks += 1
                            if consecutive_blanks == 9:
                                # Mark the 9 rows for deletion
                                rows_to_delete.extend(range(i - 8, i + 1))
                                consecutive_blanks = 0  # Reset the counter after marking
                        else:
                            consecutive_blanks = 0  # Reset the counter if a row is not blank
                
                    # Drop the rows
                    df_4208 = df_4208.drop(rows_to_delete)
                
                    # Reset the index of the DataFrame
                    df_4208.reset_index(drop=True, inplace=True)
                
                    # Forward fill the 'Barang' column
                    df_4208['Barang'] = df_4208['Barang'].replace('', pd.NA).ffill()
                    # Forward fill the 'Cabang' column
                    df_4208['Cabang'] = df_4208['Cabang'].replace('', pd.NA).ffill()
                
                    df_4208     =   df_4208[df_4208['Deskripsi']      !=      ""]
                    df_4208     =   df_4208[df_4208['Nomor #']      !=      "Nomor #"]
                    #df_4208     =   df_4208[df_4208['Nomor #']      !=      ""]
                
                    df_4208['Nama Barang']     =   df_4208['Barang']
                
                    # Drop the 'Barang' column
                    df_4208 = df_4208.drop(columns='Barang').reset_index(drop=True)
                    df_4208['Tanggal']      =   pd.to_datetime(df_4208['Tanggal'], format='%Y-%m-%d %H:%M:%S').dt.strftime('%d/%m/%Y')
                    for i in df_4208[df_4208['Masuk']==''].index:
                        df_4208.loc[i-1,'Deskripsi'] = df_4208.loc[i-1,'Deskripsi']+df_4208.loc[i,'Deskripsi']
                    df_4208 = df_4208[df_4208['Masuk']!='']
                    concatenated_df.append(df_4208)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'42.08_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='42.15':
                concatenated_df = []
                for file in uploaded_file:
                    df_4215 = pd.read_excel(file, header=4).fillna('')
                    df_4215 = df_4215.iloc[:-5]
                    df_4215 = df_4215.drop(columns=['Unnamed: 0','Nomor # Permintaan Barang'])
                    df_4215.rename(columns=lambda x: '' if 'Unnamed' in x else x, inplace=True)
                    df_4215['Tanggal']              =   pd.to_datetime(df_4215['Tanggal'], format='%d-%b-%y').dt.strftime('%d %b %Y')
                    df_4215['Tgl/Jam Pembuatan']    =   pd.to_datetime(df_4215['Tgl/Jam Pembuatan'], format='%d-%b-%Y %H:%M:%S').dt.strftime('%d %b %Y %H:%M:%S')
                    concatenated_df.append(df_4215)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'42.15_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='42.17':
                concatenated_df = []
                for file in uploaded_file:
                    df_4217     =   pd.read_excel(file, header=4).fillna('')
                    df_4217 = df_4217.drop(columns=[x for x in df_4217.reset_index().T[(df_4217.reset_index().T[1]=='')].index if 'Unnamed' in x])
                    df_4217.columns = df_4217.T.reset_index()['index'].apply(lambda x: np.nan if 'Unnamed' in x else x).ffill().values
                    df_4217 = df_4217.iloc[1:,:-3]
                    
                    df_melted =pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang'], 
                        value_vars=df_4217.columns[6:].values,
                        var_name='Nama Cabang', value_name='Total Stok').reset_index(drop=True)
    
                    df_melted2 = pd.melt(pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Satuan #1','Satuan #2','Satuan #3'], 
                        value_vars=df_4217.columns[6:].values,
                        var_name='Nama Cabang', value_name='Total Stok').drop_duplicates(),
                        id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Nama Cabang','Total Stok'],
                        var_name='Variabel', value_name='Satuan')
    
                    df_melted2 = df_melted2[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Satuan','Variabel']].drop_duplicates().reset_index(drop=True)
    
                    df_melted = df_melted.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)
                    df_melted2 = df_melted2.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)
                    
                    df_4217_final = pd.concat([df_melted2, df_melted[['Total Stok']]], axis=1)
                    df_4217_final = df_4217_final.rename(columns={'Variabel':'Kategori'})[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Kategori','Satuan','Total Stok']]
                    df_4217_final['Kode Barang'] = df_4217_final['Kode Barang'].astype('int')
                    df_4217_final['Total Stok'] = df_4217_final['Total Stok'].astype('float')
                    concatenated_df.append(df_4217_final)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'42.17_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            if selected_option=='42.17 (Rev1)':
                concatenated_df = []
                for file in uploaded_file:
                    df_4217     =   pd.read_excel(file, header=4).fillna('')
                    df_4217 = df_4217.drop(columns=[x for x in df_4217.reset_index().T[(df_4217.reset_index().T[1]=='')].index if 'Unnamed' in x])
                    df_4217.columns = df_4217.T.reset_index()['index'].apply(lambda x: np.nan if 'Unnamed' in x else x).ffill().values
                    df_4217 = df_4217.iloc[1:,:-3]

                    df_melted =pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang'],
                                        value_vars=df_4217.columns[6:].values,
                                        var_name='Nama Cabang', value_name='Total Stok').reset_index(drop=True)

                    df_melted2 = pd.melt(pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Satuan #1','Satuan #2','Satuan #3'],
                                        value_vars=df_4217.columns[6:].values,
                                        var_name='Nama Cabang', value_name='Total Stok').drop_duplicates(),
                                        id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Nama Cabang','Total Stok'],
                                        var_name='Variabel', value_name='Satuan')

                    df_melted2 = df_melted2[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Satuan','Variabel']].drop_duplicates().reset_index(drop=True)

                    df_melted = df_melted.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)
                    df_melted2 = df_melted2.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)

                    df_4217_final = pd.concat([df_melted2, df_melted[['Total Stok']]], axis=1)
                    df_4217_final = df_4217_final[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Variabel','Satuan','Total Stok']]
                    df_4217_final['Kode Barang'] = df_4217_final['Kode Barang'].astype('int')
                    df_4217_final['Total Stok'] = df_4217_final['Total Stok'].astype('float')

                    # Mengganti nama 'Satuan' menjadi 'Satuan Akhir' dan 'Total Stok' menjadi 'Saldo Akhir' setelah filtering
                    df_4217_final=df_4217_final[df_4217_final['Variabel'] == "Satuan #1"].rename(columns={"Total Stok":"Saldo Akhir"})

                    df_4217_final.insert(0, 'No. Urut', range(1, len(df_4217_final) + 1))

                    def format_nama_cabang(cabang):
                        match1 = re.match(r"\((\d+),\s*([A-Z]+)\)", cabang)
                        if match1:
                            return f"{match1.group(1)}.{match1.group(2)}"
                        else:
                            match2 = re.match(r"^(\d+)\..*?\((.*?)\)$", cabang)
                            if match2:
                                return f"{match2.group(1)}.{match2.group(2)}"
                            else:
                                return cabang

                    df_4217_final['Cabang'] = df_4217_final['Nama Cabang'].apply(format_nama_cabang)
                    # Memilih kolom akhir, menggunakan 'Satuan Akhir' (hasil rename dari 'Satuan')
                    df_4217_final=df_4217_final.loc[:,["No. Urut", "Kategori Barang","Kode Barang","Nama Barang","Satuan","Saldo Akhir", "Cabang"]]
                    concatenated_df.append(df_4217_final)

                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'42.17 (Rev1)_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
            if selected_option == '42.18':
                concatenated_df = []
                
                if uploaded_file:
                    for file in uploaded_file:
                        df_4218 = pd.read_excel(file, header=4).fillna("")
                        df_4218 = df_4218.loc[:, ~df_4218.columns.str.contains('^Unnamed')]
                        
                        def format_string(input_str):
                            input_str = str(input_str)
                            match = re.match(r'([^.\s]+)\.\d+.*?\((.*?)\)', input_str)
                            
                            if match:
                                return f"{match.group(1)}.{match.group(2)}"
                        
                        df_4218['Kode'] = df_4218['Nama'].apply(format_string)
                        df_4218 = df_4218.dropna(subset=['Nama'])
                        
                        df_4218 =   df_4218.loc[:,['Kode', 'Nama', 'Deskripsi', 'Jalan Alamat', 'Kota Alamat', 'Provinsi Alamat', 'K.Pos Alamat']]
                        
                        df_4218['Provinsi Alamat'] = df_4218['Provinsi Alamat'].replace({
                            'Tangerang': 'Banten', 
                            'Daerah Khusus Ibukota Jakarta': 'DKI Jakarta', 
                            'Jakarta': 'DKI Jakarta', 
                            'Jawa Timu': 'Jawa Timur'})
                        
                        concatenated_df.append(df_4218)
                    
                    if concatenated_df:  # Pastikan ada data sebelum menggabungkan
                        concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                        
                        excel_data = to_excel(concatenated_df)
                        
                        st.download_button(
                            label="Download Excel",
                            data=excel_data,
                            file_name=f'42.18_{get_current_time_gmt7()}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
            
            if selected_option=='44.06':
                concatenated_df = []
                if uploaded_file:
                    for file in uploaded_file:
                        df_4406     =    pd.read_excel(file, header=3).fillna('')

                        df_4406['No Formula #']         = df_4406['Unnamed: 8'].where(df_4406['Unnamed: 2'] == 'No Formula #')
                        df_4406['Produk Utama']         = df_4406['Unnamed: 8'].shift(-1).where(df_4406['Unnamed: 2'] == 'No Formula #')
                        df_4406['Kuantitas BOM']        = df_4406['Unnamed: 8'].shift(-2).where(df_4406['Unnamed: 2'] == 'No Formula #')
                        df_4406['Berlaku di Cabang']    = df_4406['Unnamed: 22'].where(df_4406['Unnamed: 15'] == 'Berlaku di Cabang')
                        df_4406['Non Aktif']            = df_4406['Unnamed: 22'].shift(-2).where(df_4406['Unnamed: 15'] == 'Berlaku di Cabang')
                        
                        df_4406[['No Formula #', 'Produk Utama', 'Kuantitas BOM','Berlaku di Cabang','Non Aktif']] = df_4406[['No Formula #', 'Produk Utama', 'Kuantitas BOM','Berlaku di Cabang','Non Aktif']].fillna(method='ffill')
                        df_4406[['Kuantitas BOM', 'Satuan BOM']] = df_4406['Kuantitas BOM'].str.split(' ', n=1, expand=True)
    
                        df_4406     =   df_4406.rename(columns={'Unnamed: 5'        :   'Kode #',
                                                                'Unnamed: 11'       :   'Nama Barang',
                                                                'Unnamed: 13'       :   'Satuan',
                                                                'Unnamed: 17'       :   'Kuantitas',
                                                                'Unnamed: 22'       :   'Harga Standar',
                                                                'Unnamed: 24'       :   'Total Harga Standar',
                                                                'Berlaku di Cabang' :   'Cabang',
                                                                'Satuan'            :   'Satuan Barang',
                                                                'Kuantitas'         :   'Kuantitas Barang'}).fillna("")

                        df_4406     =   df_4406.loc[:, ~df_4406.columns.str.startswith('Unnamed:')]
                        df_4406     =   df_4406[df_4406['Kode #'] != ""]
                        df_4406     =   df_4406[df_4406['Kode #'] != "Kode #"]
                        df_4406     =   df_4406.loc[:,['Cabang', 'No Formula #', 'Produk Utama', 'Satuan BOM', 'Kuantitas BOM','Non Aktif', 'Kode #', 'Nama Barang', 'Satuan', 'Kuantitas', 'Harga Standar', 'Total Harga Standar']]
                        concatenated_df.append(df_4406)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'44.06_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   
                
            if selected_option=='44.08':
                concatenated_df = []
                if uploaded_file:
                    for file in uploaded_file:
                        df_4408     =   pd.read_excel(file, header=4).fillna('')
                        df_4408 = df_4408[[x for x in df_4408.columns if 'Unnamed' not in x]]
                        df_4408 = df_4408.iloc[:-5]
    
                        concatenated_df.append(df_4408)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True)
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'44.08_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

            if selected_option=='51.01':
                concatenated_df = []
                for file in uploaded_file:
                    df_5101 = pd.read_excel(file, skiprows=4).fillna('')
                    df_5101 = df_5101[~df_5101['Tgl/Jam Pembuatan'].isna()][[x for x in df_5101.columns if 'Unnamed' not in x]]
                    concatenated_df.append(df_5101)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'51.01_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )  
                
            if selected_option=='99.01':
                # URL file model .pkl di GitHub (gunakan URL raw dari file .pkl di GitHub)
                url = 'https://raw.githubusercontent.com/ferifirmansah05/ads_mvn/main/database provinsi.xlsx'
                
                # Path untuk menyimpan file yang diunduh
                save_path = 'database provinsi.xlsx'
                
                # Unduh file dari GitHub
                download_file_from_github(url, save_path)
                
                # Muat model dari file yang diunduh
                if os.path.exists(save_path):
                    df_prov = load_excel(save_path)
                else:
                    print("file does not exist")
                    
                df_prov = df_prov[3:].dropna(subset=['Unnamed: 4']) 
                df_prov.columns = df_prov.loc[3,:].values
                df_prov = df_prov.loc[4:,]
                df_prov = df_prov.loc[:265, ['Nama','Provinsi Alamat','Kota Alamat']]
                df_prov = df_prov.rename(columns={'Nama':'Nama Cabang','Provinsi Alamat':'Provinsi Gudang', 'Kota Alamat': 'Kota/Kabupaten'})
                df_prov['Nama Cabang'] = df_prov['Nama Cabang'].str.extract(r'\((.*?)\)') 
                concatenated_df = []
                for file in uploaded_file:
                    df_9901 = pd.read_excel(file).fillna('')
                    df_9901 = df_9901.loc[:,['Nama Cabang', 'Nama Gudang', 'Nomor #', 'Tanggal', 'Pemasok',
                           'Kategori Pemasok', '#Group', 'Kode #', 'Nama Barang',
                           'Kategori Barang', '#Purch.Qty', '#Purch.UoM', '#Prime.Ratio',
                           '#Prime.Qty', '#Prime.UoM', '#Purch.@Price', '#Purch.Discount', 
                           '#Prime.NetPrice', '#Purch.Total']].fillna('')
                    df_9901['Nama Cabang'] = df_9901['Nama Cabang'].str.split('.').str[1]


                    
                    df_9901 = pd.merge(df_9901, df_prov, how='left', on='Nama Cabang').fillna('')
                    # Convert 'Tanggal' column to datetime format
                    df_9901['Tanggal'] = pd.to_datetime(df_9901['Tanggal'])
                    
                    # Extract month names from the 'Tanggal' column
                    df_9901['Month'] = df_9901['Tanggal'].dt.strftime('%B')
                    concatenated_df.append(df_9901)
                    
                concatenated_df = pd.concat(concatenated_df, ignore_index=True) 
                excel_data = to_excel(concatenated_df)
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'99.01_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

