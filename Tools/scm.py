import pandas as pd
import glob
import numpy as np
import time
import datetime as dt
import streamlit as st
#import fuzzywuzzy
from streamlit_image_zoom import image_zoom
from io import BytesIO
import pytz
import requests
import os
import zipfile
from xlsxwriter import Workbook
import tempfile
import re
import io
import traceback
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image
import sys
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, ColumnsAutoSizeMode, GridUpdateMode
import base64

key = "enterprise_disabled_grid"
license_key = None
enable_enterprise = True
if enable_enterprise:
    key = "enterprise_enabled_grid"
    license_key = license_key


if hasattr(sys, '_MEIPASS'):
    os.environ['STREAMLIT_STATIC_COMPONENT_PATH'] = os.path.join(
        sys._MEIPASS, 'st_aggrid', 'frontend', 'build'
    )

def resource_path(relative_path):
    """Dapatkan path absolut ke resource (kompatibel dengan PyInstaller)"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Path ke poppler/bin


if "process" not in st.session_state:
    st.session_state.process = False
if "selected_option" not in st.session_state:
    st.session_state.selected_option = None

if 'DEX.exe' not in os.listdir():
    dir_main = 'Main/'
else:
    dir_main = ''
        
def load_excel(file_path):
    with open(file_path, 'rb') as file:
        model = pd.read_excel(file, engine='openpyxl')
    return model
 
def to_excel(df, sheet_name='Sheet1'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Mengakses workbook dan worksheet untuk format header
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Menambahkan format khusus untuk header
        header_format = workbook.add_format({'border': 0, 'bold': False, 'font_size': 12})
        
        # Menulis header manual dengan format khusus
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
    processed_data = output.getvalue()
    return processed_data
  
def get_current_time_gmt7():
    tz = pytz.timezone('Asia/Jakarta')
    return dt.datetime.now(tz).strftime('%Y%m%d_%H%M%S')

def reset_button_state():
    st.session_state.button_clicked = False    

st.title('SCM-Processing')

col = st.columns([2,1])
df_sj = []
with col[0]:
    with st.container(border=True):
        selected_option = st.selectbox("Pilih Modul", ['REKAP MENTAH','REKAP PENYESUAIAN INPUTAN IA','REKAP DATA 42.02','REKAP DATA BOM-DEVIASI','REKAP SALES ESB & GIS','PENYESUAIAN IA','OCR-SJ','PROMIX','WEBSMART (DINE IN/TAKEAWAY)'])
        uploaded_file = st.file_uploader("Pilih File", type=["zip",'xlsx'])
    
        if selected_option == 'REKAP MENTAH':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
        if selected_option == 'REKAP PENYESUAIAN INPUTAN IA':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
            st.info("""
        **File** (*zip*):
        - `[Nama file].zip`
            - `4217_....xlsx` → *(Raw Data GIS)*
            - `SALDO_SO.xlsx` → kolom: `{CABANG, NAMA BARANG, #Hasil Stock Opname}`
        """)
        if selected_option == 'REKAP DATA 42.02':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
        if selected_option == 'REKAP DATA BOM-DEVIASI':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
        if selected_option == 'PENYESUAIAN IA':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
        if selected_option == 'PROMIX':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>xlsx</em></span>",unsafe_allow_html=True)
        if selected_option == 'WEBSMART (DINE IN/TAKEAWAY)':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
        if selected_option == 'OCR-SJ':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
            df_prov = pd.read_excel('Master/database provinsi.xlsx')
            df_prov = df_prov[3:].dropna(subset=['Unnamed: 4']) 
            df_prov.columns = df_prov.loc[3,:].values
            df_prov = df_prov.loc[4:,]
            df_prov = df_prov.rename(columns={'Nama':'Nama Cabang','Provinsi Alamat':'Provinsi', 'Kota Alamat': 'Kota/Kabupaten'})
            list_cab = df_prov['Nama Cabang'].str.extract(r'\(([^()]*)\)')[0].values.tolist()
            all_cab = ['All']
            all_cab.extend(list_cab)
            kol = st.columns(2)
            with kol[0]:
                all_cab = st.multiselect('Pilih Cabang', all_cab, default=['All'], on_change=reset_button_state)
            with kol[1]:
                all_date = st.slider('Pilih Tanggal', 1, 31, (1, 31), on_change=reset_button_state)
                all_date = [f"{i:02}" for i in range(all_date[0], all_date[1] + 1)]
        if selected_option == 'REKAP SALES ESB & GIS':
            st.markdown("<span style='font-size:12px; font-weight:bold;'>File format <em>zip</em></span>",unsafe_allow_html=True)
        if st.button("Process"):
            if uploaded_file:
                st.session_state.selected_option = selected_option
                st.session_state.process = True
            else:
                st.warning("Silakan upload file terlebih dahulu.")
            
        db_ia = pd.read_excel(dir_main+'Master/DATABASE_IA.xlsx')



with col[1]:
    with st.container(border=True):
        st.write('')
        if st.session_state.process:
            st.session_state.process = False
            try:
                with st.spinner('Data sedang diproses...'):
                    selected_option = st.session_state.selected_option
                    if selected_option == 'REKAP MENTAH':
                        with tempfile.TemporaryDirectory() as tmpdirname:
                            # Ekstrak file ZIP ke direktori sementara
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)
                            
                            dfs=[]
                            for file in os.listdir(tmpdirname):
                                if file.endswith('.xlsx'):
                                        df = pd.read_excel(tmpdirname+'/'+file, sheet_name='REKAP MENTAH')
                                        if 'NAMA RESTO' not in df.columns:
                                            df = df.loc[:,[x for x in df.columns if 'Unnamed' not in str(x)][:-1]].fillna('')
                                            df['NAMA RESTO'] = file.split('-')[0]
                                        dfs.append(df)
                                
                            dfs = pd.concat(dfs, ignore_index=True)
                            excel_data = to_excel(dfs, sheet_name="REKAP MENTAH")
                            st.success('Success',icon='✅')
                            st.download_button(
                                label="Download",
                                data=excel_data,
                                file_name=f'LAPORAN SO HARIAN RESTO_{get_current_time_gmt7()}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )   

                    if selected_option == 'REKAP PENYESUAIAN INPUTAN IA':
                        nama_file = uploaded_file.name.replace('.zip','')
                        with tempfile.TemporaryDirectory() as tmpdirname:
                            # Ekstrak file ZIP ke direktori sementara
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)
                            non_com = ['SUPPLIES [OTHERS]','00.COST','20.ASSET.ASSET','21.COST.ASSET']
                            concatenated_df = []
                        
                            for file in os.listdir(tmpdirname):
                                if file.startswith('4217'):
                                    df_4217     =   pd.read_excel(tmpdirname+'/'+file, header=4).fillna('')
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

                                    df_4217_final=df_4217_final[df_4217_final['Variabel'] == "Satuan #1"].rename(columns={"Total Stok":"Saldo Akhir"})

                                    #df_4217_final.insert(0, 'No. Urut', range(1, len(df_4217_final) + 1))

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

                                    #df_4217_final=df_4217_final.loc[:,["No. Urut", "Kategori Barang","Kode Barang","Nama Barang","Satuan","Saldo Akhir", "Cabang"]]
                                    concatenated_df.append(df_4217_final)
                                else:
                                    df_so = pd.read_excel(tmpdirname+'/'+file)
                                    df_so['CABANG'] = df_so['CABANG'].str.upper().str[:6]

                            df_4217 = pd.concat(concatenated_df)
                            df_4217['CABANG'] = df_4217['Cabang'].str[-6:]
                            df_4217 = df_4217.merge(df_so, left_on=['CABANG','Nama Barang'], right_on=['CABANG','NAMA BARANG'], how='left').drop(columns=['CABANG','NAMA BARANG'])
                            df_4217['#Hasil Stock Opname'] = df_4217['#Hasil Stock Opname'].fillna(0)
                            df_4217['DEVIASI(Rumus)'] = df_4217['Saldo Akhir'] - df_4217['#Hasil Stock Opname']
                            df_4217 = df_4217[df_4217['DEVIASI(Rumus)']!=0].reset_index()
                            df_4217['Tipe Penyesuaian'] = ''
                            df_4217.loc[df_4217[df_4217['DEVIASI(Rumus)']>0].index, 'Tipe Penyesuaian'] = 'Pengurangan'
                            df_4217.loc[df_4217[df_4217['DEVIASI(Rumus)']<0].index, 'Tipe Penyesuaian'] = 'Penambahan'
                            df_4217['DEVIASI(Rumus)'] = df_4217['DEVIASI(Rumus)'].abs()

                            for cab in df_4217['Cabang'].unique():
                                folder = f'{tmpdirname}/{nama_file}/{df_4217[df_4217['Cabang']==cab]['Nama Cabang'].iloc[0,]}'
                                if not os.path.exists(folder):
                                    os.makedirs(folder)
                                for kat in db_ia['KATEGORI'].unique():
                                    if kat in ['Raw Material', 'Packaging']:
                                        df_ia = df_4217[(df_4217['Kategori Barang'].isin(db_ia[db_ia['KATEGORI']==kat]['FILTER'])) 
                                                        & ~(df_4217['Nama Barang'].isin(db_ia[db_ia['KATEGORI']=='Consume']['FILTER']))
                                                        & (df_4217['Cabang']==cab)] 
                                        df_ia = df_ia.rename(columns={'Kode Barang':'Kode','Satuan':'UNIT','DEVIASI(Rumus)':'Kuantitas','Nama Cabang':'Gudang'}).loc[:,['Nama Barang','Kode','UNIT','Kuantitas','Gudang','Tipe Penyesuaian']]
                                        if not df_ia.empty:
                                            df_ia.to_excel(f'{folder}/{kat}_{cab}_{nama_file}.xlsx', index=False)
                                    if kat in ['Consume']:
                                        df_ia = df_4217[(df_4217['Nama Barang'].isin(db_ia[db_ia['KATEGORI']==kat]['FILTER']))
                                                        & (df_4217['Cabang']==cab)] 
                                        df_ia = df_ia.rename(columns={'Kode Barang':'Kode','Satuan':'UNIT','DEVIASI(Rumus)':'Kuantitas','Nama Cabang':'Gudang'}).loc[:,['Nama Barang','Kode','UNIT','Kuantitas','Gudang','Tipe Penyesuaian']]
                                        if not df_ia.empty:
                                            df_ia.to_excel(f'{folder}/{kat}_{cab}_{nama_file}.xlsx', index=False)
                                    else:         
                                        df_ia = df_4217[(df_4217['Nama Barang'].isin(db_ia[db_ia['KATEGORI']==kat]['FILTER']))
                                                        & (df_4217['Cabang']==cab) & (df_4217['Kategori Barang'].isin(non_com))] 
                                        df_ia = df_ia.rename(columns={'Kode Barang':'Kode','Satuan':'UNIT','DEVIASI(Rumus)':'Kuantitas','Nama Cabang':'Gudang'}).loc[:,['Nama Barang','Kode','UNIT','Kuantitas','Gudang','Tipe Penyesuaian']]
                                        if not df_ia.empty:
                                            df_ia.to_excel(f'{folder}/{kat}_{cab}_{nama_file}.xlsx', index=False)

                            folder_path = f'{tmpdirname}/{nama_file}'
                            zip_buffer = io.BytesIO()
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                for root, dirs, files in os.walk(folder_path):
                                    for file in files:
                                        file_path = os.path.join(root, file)
                                        arcname = os.path.relpath(file_path, start=folder_path)
                                        zip_file.write(file_path, arcname)

                            # Pindahkan ke awal buffer agar bisa dibaca
                            zip_buffer.seek(0)
                            st.success('Success',icon='✅')
                            st.download_button(
                                label="Download Zip",
                                data=zip_buffer,
                                file_name=f"REKAP PENYESUAIAN STOK (IA)_{nama_file}_{get_current_time_gmt7()}.zip",
                                mime="application/zip"
                            )
                        
                    if selected_option == 'PROMIX':
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
                            st.success('Success',icon='✅')
                            st.download_button(
                                    label="Download",
                                    data=to_excel(df_promix),
                                    file_name=f'promix_{get_current_time_gmt7()}.xlsx',
                                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )
                    
                    if selected_option == 'REKAP DATA 42.02':
                        with tempfile.TemporaryDirectory() as tmpdirname:
                        # Ekstrak file ZIP ke folder sementara
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)
                        
                        all_dfs = []
                        for file in os.listdir(tmpdirname):
                            if file.endswith('.xlsx'):
                                file_path = os.path.join(tmpdirname, file)
                
                                # Ambil nama file dan ekstrak kode cabang
                                match = re.search(r'_(\d{4}\.[A-Z]+)', file)
                                cabang = match.group(1) if match else ''
                
                                # Baca Excel: header ke-5 (index ke-4)
                                df = pd.read_excel(file_path, header=4).fillna('')
                                df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
                                df['Cabang'] = cabang
                
                                all_dfs.append(df)
                
                        if all_dfs:
                            df_combined = pd.concat(all_dfs, ignore_index=True)
                
                            # Tombol download hasil
                            st.success('Success',icon='✅')
                            st.download_button(
                                label="Download",
                                data=to_excel(df_combined),
                                file_name=f'42.02 Combine_{get_current_time_gmt7()}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )
                        else:
                            st.warning("Tidak ada file .xlsx ditemukan dalam ZIP.")

                    if selected_option == 'WEBSMART (DINE IN/TAKEAWAY)':
                        with tempfile.TemporaryDirectory() as tmpdirname:
                        # Ekstrak file ZIP ke folder sementara
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)
                        
                            all_dfs = [] 
                            for file in os.listdir(tmpdirname):
                                if file.endswith('.xls'):
                                        file_path = os.path.join(tmpdirname, file)
                                        df_qty = pd.read_html(f'{file_path}')[0]
                                        df_qty = df_qty[df_qty.iloc[1:,].columns[df_qty.iloc[1:,].apply(lambda col: col.astype(str).str.contains('QTY', case=False, na=False).any())]]
                                        df_qty.iloc[0,:] = df_qty.iloc[0,:] + '_QTY'
                                        df_qty.columns = df_qty.iloc[0,:]
                                        try:
                                            df_qty = df_qty.iloc[2:,:-2].iloc[:-1].drop(columns='OFFLINE_QTY')
                                        except:
                                            df_qty = df_qty.iloc[2:,:-2].iloc[:-1]
                                        df_qty.iloc[:,1:] = df_qty.iloc[:,1:].astype(float)
                                        df_rp = pd.read_html(f'{file_path}')[0]
                                        df_rp = df_rp[df_rp.iloc[1:,].columns[df_rp.iloc[1:,].apply(lambda col: col.astype(str).str.contains('Rp', case=False, na=False).any())]]
                                        df_rp.iloc[0,1:] = df_rp.iloc[0,1:] + '_RP'
                                        df_rp.columns = df_rp.iloc[0,:]
                                        try:
                                            df_rp = df_rp.iloc[2:,:-2].iloc[:-1].drop(columns='OFFLINE_QTY')
                                        except:
                                            df_rp = df_rp.iloc[2:,:-2].iloc[:-1]
                                        df_rp.iloc[:,1:] = df_rp.iloc[:,1:].astype(float)
                                        df = df_rp.reset_index().merge(df_qty.reset_index(), on='index')
                                        df = df.melt(id_vars=['RESTO'], value_vars=df.iloc[:,1:].columns,value_name='RP',var_name='CATEGORY')
                                        df[['CATEGORY','Variable']] = df['CATEGORY'].str.split('_', n=1, expand=True)
                                        df = df[df['Variable']=='QTY'].reset_index(drop=True).reset_index().merge(df[df['Variable']=='RP'].reset_index(drop=True).reset_index(), on=['index','RESTO','CATEGORY']).drop(columns=['index','Variable_x','Variable_y']).rename(columns={'RP_x':'QTY','RP_y':'VALUE'})
                                        df['TYPE'] = df['CATEGORY'].apply(lambda x: x if x=='DINE IN' else 'TAKE AWAY')
                                        df['MONTH'] = re.findall(r'_(\w+)', file)[-1]
                                        df = df.sort_values(['RESTO','TYPE']).reset_index(drop=True)
                                        all_dfs.append(df)
                    
                            if all_dfs:
                                df_combined = pd.concat(all_dfs, ignore_index=True)
                    
                                # Tombol download hasil
                                st.success('Success',icon='✅')
                                st.download_button(
                                    label="Download",
                                    data=to_excel(df_combined),
                                    file_name=f'WEBSMART (DINE IN/TAKEAWAY) Combine_{get_current_time_gmt7()}.xlsx',
                                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )

                    if selected_option == 'PENYESUAIAN IA':
                        with tempfile.TemporaryDirectory() as tmpdirname:
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)

                                all_dfs = []
                                for file in os.listdir(tmpdirname):
                                    if file.startswith('penyesuaian') and file.endswith(('.xlsx', '.xls', '.csv')):
                                                if file.endswith('.csv'):
                                                    df = pd.read_csv(os.path.join(tmpdirname, file), skiprows=9)
                                                else:
                                                    df = pd.read_excel(os.path.join(tmpdirname, file), skiprows=9)

                                                df.drop(df.columns[0], axis=1, inplace=True)
                                                if 'Kode' in df.columns:
                                                    df = df[~df['Kode'].isin(['Penyesuaian Persediaan', 'Keterangan', 'Kode'])]
                                                    df = df[df['Kode'].notna()]
                                                if 'Tipe' in df.columns:
                                                    mask_penambahan = df['Tipe'].str.lower() == 'penambahan'
                                                    for col in ['Kts.', 'Total Biaya']:
                                                        if col in df.columns:
                                                            df.loc[mask_penambahan, col] = df.loc[mask_penambahan, col] * -1
                                                df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
                                                if 'Gudang' in df.columns:
                                                    def transform_gudang(val):
                                                        try:
                                                            prefix = re.search(r'^(\d+)', str(val))
                                                            kode = re.search(r'\((.*?)\)', str(val))
                                                            if prefix and kode:
                                                                return f"{prefix.group(1)}.{kode.group(1)}"
                                                            else:
                                                                return val
                                                        except:
                                                            return val
                                                    df['Gudang'] = df['Gudang'].apply(transform_gudang)
                                                all_dfs.append(df)

                                    if all_dfs:
                                        final_df = pd.concat(all_dfs, ignore_index=True)

                                harga = pd.read_excel(f"{tmpdirname}/bahan/Harga.xlsx", skiprows=4).fillna("").drop(
                                        columns={'Kategori Barang', 'Kode Barang', 'Nama Satuan', 'Saldo Awal', 'Masuk', 'Keluar'})

                                harga = harga[~harga['Nama Barang'].isin(['Total Nama Barang', ''])].rename(
                                    columns={'Saldo Akhir': 'Kuantitas', 'Unnamed: 14': 'Nilai'})
                                harga = harga[harga['Nama Barang'].notna()]
                                harga = harga.loc[:, ~harga.columns.str.startswith('Unnamed')]

                                def safe_divide(row):
                                    try:
                                        return abs(row['Nilai'] / row['Kuantitas'])
                                    except:
                                        return 0

                                harga['Harga'] = harga.apply(safe_divide, axis=1)

                                # Cari folder REKAP PENYESUAIAN STOK (IA)
                                rekap_path = ""
                                for root, dirs, files in os.walk(tmpdirname):
                                    for dir_name in dirs:
                                        if dir_name.strip().lower() == "rekap penyesuaian stok (ia)":
                                            rekap_path = os.path.join(root, dir_name)
                                            break
                                    if rekap_path:
                                        break

                                if not rekap_path:
                                    st.error("Folder 'REKAP PENYESUAIAN STOK (IA)' tidak ditemukan di dalam ZIP.")
                                    st.stop()

                                # Gabungkan semua file dari folder tersebut
                                all_files = []
                                for root, dirs, files in os.walk(rekap_path):
                                    for file in files:
                                        if file.endswith('.xlsx') or file.endswith('.xls'):
                                            all_files.append(os.path.join(root, file))

                                combined_df = pd.DataFrame()
                                for file in all_files:
                                    try:
                                        df = pd.read_excel(file)
                                        combined_df = pd.concat([combined_df, df], ignore_index=True)
                                    except Exception as e:
                                        print(f"Gagal membaca file: {file} karena {e}")

                                def format_gudang(g):
                                    match = re.match(r"(\d+)(?:\.\d+)?-.*\((\w+)\)", str(g))
                                    if match:
                                        return f"{match.group(1)}.{match.group(2)}"
                                    return g
                                                                
                                combined_df['Gudang'] = combined_df['Gudang'].apply(format_gudang)
                                
                                if 'Tipe Penyesuaian' in combined_df.columns:
                                    mask_penambahan2 = combined_df['Tipe Penyesuaian'].str.lower() == 'penambahan'
                                    if 'Kuantitas' in combined_df.columns:
                                        combined_df.loc[mask_penambahan2, 'Kuantitas'] = combined_df.loc[mask_penambahan2, 'Kuantitas'] * -1  
                                            
                                final_df['Kode'] = final_df['Kode'].astype(str)
                                combined_df['Kode'] = combined_df['Kode'].astype(str)                

                                final_df = final_df.merge(
                                    combined_df[['Kode', 'Gudang', 'Kuantitas']],
                                    on=['Kode', 'Gudang'],
                                    how='left'
                                )
                                final_df.rename(columns={'Kuantitas': 'REKAP'}, inplace=True)
                                final_df['Kts.'] = final_df['Kts.'].fillna(0).astype(float)
                                final_df['REKAP'] = final_df['REKAP'].fillna(0).astype(float)
                                final_df['Total Biaya'] = final_df['Total Biaya'].astype(float)

                                def selisih(row):
                                    try:
                                        return row['REKAP'] - row['Kts.']
                                    except:
                                        return 0

                                final_df['SELISIH'] = final_df.apply(selisih, axis=1)

                                final_df['Nama Barang'] = final_df['Nama Barang'].astype(str)
                                harga['Nama Barang'] = harga['Nama Barang'].astype(str)


                                final_df['Harga Sistem'] = final_df['Total Biaya'] / final_df['Kts.']
                                final_df = final_df.merge(
                                    harga[['Nama Barang', 'Harga']],
                                    on='Nama Barang',
                                    how='left'
                                )

                                final_df['Harga'] = final_df['Harga'].astype(float) 
                                final_df['Selisih Harga'] = (final_df['Harga Sistem'] - final_df['Harga']).replace([np.inf, -np.inf, np.nan],0)
                                
                                st.success('Success',icon='✅')
                                st.download_button(
                                    label="Download",
                                    data=to_excel(final_df),
                                    file_name=f'Penyesuaian IA Combine_{get_current_time_gmt7()}.xlsx',
                                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                                )
                            
                    if selected_option == 'REKAP DATA BOM-DEVIASI':
                        with tempfile.TemporaryDirectory() as tmpdirname:
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)
                
                            dir_db = tmpdirname+'/Database/'
                            dir_raw = tmpdirname+'/Raw Data/'
                            
                            for file in os.listdir(dir_raw):
                                if file.startswith('4101'):
                                    df_4101 =   pd.read_excel(dir_raw+file, header=4)
                                    df_4101= df_4101[~df_4101['Nama Cabang'].isna()].loc[:, ~df_4101.columns.str.startswith('Unnamed')]
                                    df_4101 = df_4101[df_4101['Akun Penyesuaian Persediaan'].isin(['COM Deviasi - Resto','COM Consume - Resto', 'Biaya Packaging - RESTO'])].reset_index(drop=True)
                                    df_4101.loc[df_4101[df_4101['Tipe Penyesuaian']=='Pengurangan'].index,['Kuantitas','Total Biaya']] = - df_4101[df_4101['Tipe Penyesuaian']=='Pengurangan'][['Kuantitas','Total Biaya']]
                                    df_4101 = df_4101.groupby(['Nama Cabang','Nama Barang'])[['Kuantitas','Total Biaya']].sum().reset_index()
                                    df_4101['Cabang'] = df_4101['Nama Cabang'].str.extract(r'\.(.+)')
                                if file.startswith('4104') and ('ACR' in file):
                                    df_4104b = pd.read_excel(dir_raw+file,header=4)
                                    df_4104b = df_4104b.dropna(how='all',axis=1)
                                    df_4104b = df_4104b.iloc[1:-1,:]
                                    df_db = pd.DataFrame({'variable':[x for x in df_4104b.columns if 'Unnamed' in x],'Nama Barang':[x for x in df_4104b.columns if 'Unnamed' not in x][1:]})
                                    df_4104b = df_4104b.melt(id_vars='Nama Gudang', value_vars=[x for x in df_4104b.columns if 'Unnamed' not in x][1:],value_name='Nominal Kts Keluar', var_name='Nama Barang').merge(
                                    df_4104b.melt(id_vars='Nama Gudang', value_vars=[x for x in df_4104b.columns if 'Unnamed' in x],value_name='Kts Keluar').merge(df_db,how='left',on='variable'), how='left', on=['Nama Gudang','Nama Barang']).drop(columns='variable')
                                    df_4104b.loc[:,['Nominal Kts Keluar', 'Kts Keluar']] = - df_4104b[['Nominal Kts Keluar', 'Kts Keluar']]
                                if file.startswith('4104') and ('CN' in file):
                                    df_4104d = pd.read_excel(dir_raw+file,header=4)
                                    df_4104d = df_4104d.dropna(how='all',axis=1)
                                    df_4104d = df_4104d.iloc[1:-1,:]
                                    df_db = pd.DataFrame({'variable':[x for x in df_4104d.columns if 'Unnamed' in x],'Nama Barang':[x for x in df_4104d.columns if 'Unnamed' not in x][1:]})
                                    df_4104d = df_4104d.melt(id_vars='Nama Gudang', value_vars=[x for x in df_4104d.columns if 'Unnamed' not in x][1:],value_name='Nominal Kts Keluar', var_name='Nama Barang').merge(
                                    df_4104d.melt(id_vars='Nama Gudang', value_vars=[x for x in df_4104d.columns if 'Unnamed' in x],value_name='Kts Keluar').merge(df_db,how='left',on='variable'), how='left', on=['Nama Gudang','Nama Barang']).drop(columns='variable')
                                if file.startswith('3224'):
                                    df_3224 = pd.read_excel(dir_raw+file, skiprows=range(0, 4))
                                    df_3224 = df_3224.loc[:,['Tanggal','Nomor # PO','Nomor # RI','Pemasok','Kode #','Nama Barang','Kts Terima','Satuan','@Harga','Total Harga','#Kts Ditagih','Nama Gudang','Nama Cabang Penerimaan Barang','Status Penerimaan Barang','Pembuat Data','Tgl/Jam Pembuatan']]
                                    df_3224 = df_3224[~df_3224['Nomor # PO' ].isna()]
                                    df_3224 = df_3224[df_3224['Nama Barang']=='KABEL TIES - RESTO'].groupby(['Nama Cabang Penerimaan Barang','Nama Barang'])[['Kts Terima','Total Harga']].sum().reset_index().rename(columns={'Nama Cabang Penerimaan Barang':'Nama Cabang','Kts Terima':'QTY RI','Total Harga':'NOMINAL RI'})
                                if file.startswith('1333'):
                                    omset = pd.read_excel(dir_raw+file, header=2).drop_duplicates().rename(columns={'Row Labels':'Nama Cabang'})[['Nama Cabang','OMSET']]
                                if file.startswith('DATA WASUTRI'):
                                    df_wst = pd.read_excel(dir_raw+file)[['NAMA RESTO','NAMA BARANG','QTY','KETERANGAN']]
                                    df_wst = df_wst.rename(columns={'NAMA BARANG':'Nama Barang','NAMA RESTO':'Nama Cabang'}).groupby(['Nama Cabang','Nama Barang','KETERANGAN'])[['QTY']].sum().reset_index()
                                if file.startswith('NOMINAL BIANG'):
                                    nb = pd.read_excel(dir_raw+file)
                                    nb['NOMINAL BIANG PER GRAM'] = nb['Nominal'] / nb['Qty']
                            
                            for file in os.listdir(dir_db):
                                if file.startswith('TEMPLATE KLASIFIKASI'):
                                    db = pd.read_excel(dir_db+file).iloc[:,:4]
                                if file.startswith('AREA'):
                                    db_area = pd.read_excel(dir_db+file).rename(columns={'KODE DAN NAMA RESTO':'Nama Cabang'})
                                if file.startswith('DATABASE PAPERBOX'):
                                    db_pkg = pd.read_excel(dir_db+file).rename(columns={'RESTO':'Nama Cabang'})
                                    
                            data = pd.concat([df_4104b,df_4104d],ignore_index=True).rename(columns={'Nominal Kts Keluar':'NOMINAL BOM','Kts Keluar':'QTY BOM'}).groupby(['Nama Gudang','Nama Barang'])[['QTY BOM','NOMINAL BOM']].sum().reset_index()
                            data['Nama Cabang'] = data['Nama Gudang'].str[:5] + data['Nama Gudang'].str.extract(r'\(([^()]*)\)')[0].values
                            data = data.merge(df_4101.rename(columns={'Kuantitas':'QTY DEVIASI','Total Biaya':'NOMINAL DEVIASI'}), on=['Nama Cabang','Nama Barang'], how='outer')
                            data = data.merge(df_3224,on=['Nama Cabang','Nama Barang'], how='outer').merge(df_wst[df_wst['KETERANGAN']=='QTY WASTE'].rename(columns={'QTY':'QTY WASTE'}).drop(columns=['KETERANGAN']),on=['Nama Cabang','Nama Barang'], how='outer').merge(
                                df_wst[df_wst['KETERANGAN']=='QTY SUSUT'].rename(columns={'QTY':'QTY SUSUT'}).drop(columns=['KETERANGAN']),on=['Nama Cabang','Nama Barang'], how='outer').merge(
                                        df_wst[df_wst['KETERANGAN']=='QTY TRIAL'].rename(columns={'QTY':'QTY TRIAL'}).drop(columns=['KETERANGAN']),on=['Nama Cabang','Nama Barang'], how='outer'
                                ).merge(db, on=['Nama Barang'], how='left')
                            data.loc[:,['QTY BOM','QTY DEVIASI','QTY SUSUT','QTY TRIAL','QTY WASTE','NOMINAL BOM','NOMINAL DEVIASI']] = data[['QTY BOM','QTY DEVIASI','QTY SUSUT','QTY TRIAL','QTY WASTE','NOMINAL BOM','NOMINAL DEVIASI']].fillna(0)
                            data['Harga'] = abs(data['NOMINAL BOM']/data['QTY BOM']).fillna(0)
                            data['QTY USAGE'] = ''
                            data['NOMINAL USAGE'] = ''
                            data.loc[data[data['STATUS']=='USAGE'].index,'QTY USAGE'] = data[data['STATUS']=='USAGE']['QTY DEVIASI']
                            data.loc[data[data['STATUS']=='USAGE'].index,'NOMINAL USAGE'] = data[data['STATUS']=='USAGE']['NOMINAL DEVIASI']
                            data['QTY USAGE'] = data['QTY USAGE'].replace('',0)
                            data['NOMINAL USAGE'] = data['NOMINAL USAGE'].replace('',0)
                            data.loc[data[data['STATUS']=='USAGE'].index,'QTY DEVIASI'] = 0
                            data.loc[data[data['STATUS']=='USAGE'].index,'NOMINAL DEVIASI'] = 0
                            data.loc[data[(data['Nama Barang']=='KABEL TIES - RESTO') & ~(data['QTY RI'].isna())].index,'QTY USAGE'] = - data[(data['Nama Barang']=='KABEL TIES - RESTO') & ~(data['QTY RI'].isna())]['QTY RI']
                            data.loc[data[(data['Nama Barang']=='KABEL TIES - RESTO') & ~(data['NOMINAL RI'].isna())].index,'NOMINAL USAGE'] = - data[(data['Nama Barang']=='KABEL TIES - RESTO') & ~(data['NOMINAL RI'].isna())]['NOMINAL RI']

                            data['NOMINAL SUSUT'] = data['QTY SUSUT'] * data['Harga']
                            data['NOMINAL WASTE'] = data['QTY WASTE'] * data['Harga']
                            data['NOMINAL TRIAL'] = data['QTY TRIAL'] * data['Harga']
                            data['QTY LOSS SURPUS'] = data['QTY DEVIASI'] - data['QTY WASTE'] - data['QTY SUSUT'] - data['QTY TRIAL']
                            data['NOMINAL LOSS SURPUS'] = data['NOMINAL DEVIASI'] - data['NOMINAL WASTE'] - data['NOMINAL SUSUT'] - data['NOMINAL TRIAL']

                            data['QTY COM'] = 0
                            data['NOMINAL COM'] = 0
                            data['QTY COM'] = data['QTY BOM'] + data['QTY DEVIASI'] + data['QTY USAGE']
                            data['NOMINAL COM'] = data['NOMINAL BOM'] + data['NOMINAL DEVIASI'] + data['NOMINAL USAGE']
                            data = data.merge(omset.iloc[:,:2], on='Nama Cabang', how='left').merge(nb[['Nama Barang','NOMINAL BIANG PER GRAM']], on='Nama Barang',how='left').rename(columns={'OMSET':'OMSET 1'})
                            omset['Nama Barang'] = 'ADONAN PANGSIT (V.20)'
                            data = data.merge(pd.concat([omset,omset.replace('ADONAN PANGSIT (V.20)','KERTAS BAKPAO 10.5 CM (V.20)')]), on=['Nama Cabang','Nama Barang'], how='left')
                            data.loc[:,['OMSET','OMSET 1']]= data[['OMSET','OMSET 1']].fillna(0)
                            data['QTY WASTE + SUSUT'] = data['QTY WASTE'] + data['QTY SUSUT']
                            data['% WASTE + SUSUT'] = -1*data['QTY WASTE + SUSUT']/data['QTY BOM']
                            data['NOMINAL BUMBU'] = data['NOMINAL BIANG PER GRAM']*data['QTY BOM']
                            data['NOMINAL BOM2'] = data['NOMINAL BOM'] -data['NOMINAL BUMBU']

                            data = data[['Akun Penyesuaian Persediaan','STATUS','Nama Cabang','Nama Barang','SATUAN',
                                'QTY BOM','QTY COM','QTY DEVIASI', 'QTY USAGE','QTY WASTE', 'QTY SUSUT','QTY TRIAL','QTY LOSS SURPUS',
                                'NOMINAL BOM','NOMINAL COM','NOMINAL DEVIASI','NOMINAL USAGE', 'NOMINAL WASTE', 'NOMINAL SUSUT', 'NOMINAL TRIAL','NOMINAL LOSS SURPUS',
                                'OMSET 1','Harga','QTY WASTE + SUSUT','% WASTE + SUSUT','NOMINAL BIANG PER GRAM','NOMINAL BUMBU','NOMINAL BOM2'
                                ]].replace([np.inf, -np.inf, np.nan], 0).merge(db_area, on='Nama Cabang', how='left').merge(db_pkg.drop_duplicates(), on='Nama Cabang',how='left').fillna('').rename(
                                    columns={'Nama Cabang':'RESTO','Nama Barang':'NAMA BAHAN','SATUAN':'Satuan'})

                            st.success('Success',icon='✅')
                            st.download_button(
                                label="Download",
                                data=to_excel(data),
                                file_name=f'REKAP DATA BOM-DEVIASI_{get_current_time_gmt7()}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )


                    if selected_option == 'OCR-SJ':
                        poppler_path = resource_path("bin")
                        tesseract_path = resource_path("Tesseract-OCR/tesseract.exe")
                        pytesseract.pytesseract.tesseract_cmd = tesseract_path
                        with tempfile.TemporaryDirectory() as tmpdirname:
                        # Ekstrak file ZIP ke folder sementara
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)
                        
                            all_dfs = []
                            namafile = []
                            fileimg = []
                            for tanggal in os.listdir(f'{tmpdirname}/File SJ/'):
                                if tanggal in all_date:
                                    for cabang in os.listdir(f'{tmpdirname}/File SJ/{tanggal}'):
                                        if (cabang in all_cab) or ('All' in all_cab):
                                            for filename in os.listdir(f'{tmpdirname}/File SJ/{tanggal}/{cabang}'):
                                                if filename.endswith('pdf'):
                                                    images = convert_from_bytes(open(f'{tmpdirname}/File SJ/{tanggal}/{cabang}/{filename}', "rb").read(), dpi=370,
                                                        poppler_path=poppler_path)
                                                else:
                                                    images = [Image.open(f'{tmpdirname}/File SJ/{tanggal}/{cabang}/{filename}')]

                                                namafile.append(filename)
                                                fileimg.append(images)

                                                for img in images:
                                                    text = pytesseract.image_to_string(img, config="--psm 6")
                                                    nomor_match = re.search(r"Nomor\.?\s+(.*)", text)
                                                    nomor = nomor_match.group(1) if nomor_match else ''
                                                    if not (re.search(r"\d{5}", nomor)) and (("Pengiriman" in text) or (len(images)==1)):
                                                        nomor = re.search(r"[\s.,]*\d{4}[\s.,]*\d{2}[\s.,]*\d{5}", text)
                                                        nomor = nomor.group() if nomor else ""

                                                    lines = text.splitlines()

                                                    keterangan_indices = [i for i, line in enumerate(lines) if "Keterangan :" in line]
                                                    terima_index = next((i for i, line in enumerate(lines) if "Terima" in line), None)

                                                    # Ambil "Keterangan" terakhir yang muncul sebelum "Terima"
                                                    start_index = None
                                                    for idx in reversed(keterangan_indices):
                                                        if terima_index is None or idx < terima_index:
                                                            start_index = idx
                                                            break

                                                    # Ekstrak bagian teks di antaranya
                                                    ket = ""
                                                    if start_index is not None and terima_index is not None and start_index < terima_index:
                                                        extracted_lines = lines[start_index + 1:terima_index]
                                                        ket = "\n".join(extracted_lines)

                                                    items_section = re.findall(r"^\s*\d+\s+\d{6}\s+.*", text, re.MULTILINE)#re.search(r"Kode.*?(?=Keterangan\s*:)", text, re.DOTALL)

                                                    if items_section:
                                                        rows = [row.strip() for row in items_section]
                                                        for row in rows:
                                                            #r"^(\d+)\.?\s+(\d{6})\s+(.*?)\s+(\d+(?:\.\d+)?)\s+([A-Z]+)(?:\s+(.*))?$"
                                                            match = re.match(r"^(\d+)\s+(\d{6})\s+(.*)\s+(\d+(?:[.,]\d+)?)\s+([A-Z]+)\s*(.*)$", row)
                                                            if match:
                                                                no, kode, nama, kts, satuan, keterangan = match.groups()
                                                                all_dfs.append({
                                                                        "Cabang Penerima_SJ":cabang,
                                                                        "Tanggal_SJ": tanggal,
                                                                        "Nomor_SJ": nomor,
                                                                        "Kode_SJ": kode,
                                                                        "Nama Barang_SJ": nama.strip(),
                                                                        "Kts_SJ": kts,
                                                                        "Satuan_SJ": satuan.strip(),
                                                                        "Keterangan_SJ": ket,
                                                                        "Nama File_SJ": filename
                                                                    })
                                                    if ket.splitlines():
                                                        for k in ket.splitlines():
                                                            match = re.match(r"^(.*\S)\s+(\d+)\s+([A-Z]+)$", k.strip())
                                                            if match:
                                                                all_dfs.append({
                                                                                "Cabang Penerima_SJ":cabang,
                                                                                "Tanggal_SJ": tanggal,
                                                                                "Nomor_SJ": nomor,
                                                                                "Kode_SJ": 0,
                                                                                "Nama Barang_SJ": match.group(1),
                                                                                "Kts_SJ": match.group(2),
                                                                                "Satuan_SJ": match.group(3),
                                                                                "Keterangan_SJ": ket,
                                                                                "Nama File_SJ": filename
                                                                            })
                            df_sj = pd.DataFrame(all_dfs)

                            all_dfs = []
                            for file in os.listdir(tmpdirname+'/4205'):
                                if file.startswith('4205'):
                                    df_4205 = pd.read_excel(tmpdirname+'/4205/'+file, header=4)
                                    df_4205 = df_4205[~df_4205['Tanggal #Kirim'].isna()][[x for x in df_4205.columns if 'Unnamed' not in x]]
                                    df_4205['Kode Barang'] = df_4205['Kode Barang'].astype(int)
                                    df_4205['Cabang Terima'] = df_4205['Gudang #Terima'].str.extract(r'\(([^()]*)\)')[0].values
                                    df_4205 = df_4205.reset_index()
                                    df_4205['Tanggal'] = df_4205['Tanggal #Kirim'].dt.strftime('%d')
                                    all_dfs.append(df_4205)
                            df_4205 = pd.concat(all_dfs, ignore_index=True)
                            all_dfs = []
                            for file in os.listdir(tmpdirname+'/History'):
                                dfh = pd.read_excel(tmpdirname+'/History/'+file)
                                dfh['Tanggal #Kirim'] = pd.to_datetime(dfh['Tanggal #Kirim'], format="%d/%m/%Y %H:%M")
                                all_dfs.append(dfh)
                            if all_dfs:
                                st.session_state.dfh = pd.concat(all_dfs, ignore_index=True)
                            else:
                                st.session_state.dfh = pd.DataFrame()

                                
                            def clean_it(entry):
                                match_start = re.search(r'202\d.*', entry)
                                if match_start:
                                    portion = match_start.group(0)

                                    portion = portion.replace(",", ".").replace(" ", "").replace("|", "").replace(":", "")

                                    match = re.search(r'(202\d)[\.]?(\d{2})[\.]?(\d{4,5})', portion)
                                    if match:
                                        year, month, number = match.groups()
                                        formatted = f"IT.{year}.{month}.{number.zfill(5)}"
                                        return(formatted)
                                    else:
                                        return(entry)

                            df_sj['Nomor_SJ'] = df_sj['Nomor_SJ'].apply(lambda x: clean_it(x))
                            df_sj['Kode_SJ'] = df_sj['Kode_SJ'].astype(int)

                            df_sj1 = df_4205.merge(df_sj, left_on=['Nomor #Kirim', 'Cabang Terima','Kode Barang'], right_on=['Nomor_SJ','Cabang Penerima_SJ','Kode_SJ'],how='inner')
                            df_sj1 = pd.concat([df_sj1,
                            df_4205.merge(df_sj[~df_sj['Nama File_SJ'].isin(df_sj1['Nama File_SJ'].unique())], left_on=['Tanggal', 'Cabang Terima','Kode Barang'], right_on=['Tanggal_SJ','Cabang Penerima_SJ','Kode_SJ'],how='inner'),
                            df_4205])
                            df_sj1 = df_sj1.drop_duplicates(subset=df_sj1.columns[:19], keep='first')
                            
                            if not df_sj[df_sj['Kode_SJ']==0].empty:
                                df_sj1 = pd.concat([df_sj1,df_sj[df_sj['Kode_SJ']==0].drop(columns='Nomor_SJ').merge(df_sj1[['Nomor #Kirim','Nama File_SJ']].drop_duplicates(), on='Nama File_SJ', how='left').merge(
                                    df_4205.iloc[:,1:8].drop_duplicates(), how='left', on=['Nomor #Kirim'])])
                            
                            df_sj = df_sj1
                            df_sj['#Qty Kirim'] = df_sj['#Qty Kirim'].astype(float)
                            df_sj['Kts_SJ'] = df_sj['Kts_SJ'].str.replace('.','').str.replace(',','.').astype(float)
                            df_sj['Selisih'] = df_sj['#Qty Kirim'] - df_sj['Kts_SJ']
                            i = df_sj[(df_sj['Selisih']!=0) & ~(df_sj['Selisih'].isna()) & ~(df_sj['Nomor #Kirim'].isna())
                                & (df_sj['Nama Barang_SJ'].fillna('').apply(lambda x: True if re.search(r"(\d+(?:[.,]\d+)?)\s+([A-Z]{2,})\s*[-–—]?$",x) else False))].index
                            if not i.empty:
                                val = df_sj.loc[i,'Nama Barang_SJ'].apply(lambda x: re.match(r"^(.*)\s+(\d+(?:[.,]\d+)?)\s+([A-Z]{2,})\s*[-–—]?\s*$", x) if re.match(r"^(.*)\s+(\d+(?:[.,]\d+)?)\s+([A-Z]{2,})\s*[-–—]?\s*$", x) else '')
                                df_sj.loc[i,'Nama Barang_SJ'] = val.apply(lambda x:x.group(1))
                                df_sj.loc[i,'Kts_SJ'] = val.apply(lambda x:x.group(2))
                                df_sj.loc[i,'Satuan_SJ'] = val.apply(lambda x:x.group(3))
                                df_sj.loc[i,'Kts_SJ'] = df_sj.loc[i,'Kts_SJ'].apply(str).str.replace('.','').str.replace(',','.').astype(float)
                                df_sj['Selisih'] = df_sj['#Qty Kirim'] - df_sj['Kts_SJ']
                                
                            df_sj = df_sj.drop(columns=['Nama File_SJ','Keterangan_SJ']).merge(df_sj[['Nomor #Kirim','Nama File_SJ']].sort_values(['Nomor #Kirim','Nama File_SJ']).drop_duplicates().dropna(),
                                        on='Nomor #Kirim', how='left')
                            df_sj = df_sj.drop(columns=['index','Tanggal','Cabang Terima','Status Pengiriman #','#Tgl/Jam Pembuatan RI','#Tgl Kirim vs Tgl Terima','Cabang Penerima_SJ','Nomor_SJ','Tanggal_SJ']).sort_values(['Nama File_SJ']).fillna('')
                            
                            st.session_state.df_sj = df_sj 
                            st.session_state.df_file_sj = pd.DataFrame({'Nama File_SJ':namafile,'Images':fileimg})
                            st.success('Success',icon='✅')
                            st.download_button(
                                label="Download",
                                data=to_excel(df_sj),
                                file_name=f'CHECK SJ_{get_current_time_gmt7()}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

                    if selected_option == 'REKAP SALES ESB & GIS':
                        with tempfile.TemporaryDirectory() as tmpdirname:
                            with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                zip_ref.extractall(tmpdirname)

                            for file in os.listdir(tmpdirname):
                                if file.startswith('4121'):
                                    df_4121 = pd.read_excel(os.path.join(tmpdirname, file), header=4)
                                    df_4121 = df_4121.loc[~df_4121['Kode Barang'].isna(), ~df_4121.columns.str.startswith('Unnamed')]
                                    df_4121 = df_4121[df_4121['Non Aktif Barang & Jasa']=='Tidak'].reset_index(drop=True)
                                if file.startswith('2205'):
                                    df_2205 = pd.read_excel(os.path.join(tmpdirname, file), skiprows=4).fillna('')
                                    df_2205 = df_2205.iloc[:-5]
                                    df_2205 = df_2205.loc[:, ~df_2205.columns.str.startswith('Unnamed:')]
                                    db_2205 = df_2205[['Kode #','Nama Barang','Satuan']].drop_duplicates()
                                if file.startswith('Daily'):
                                    df_esb = pd.read_excel(os.path.join(tmpdirname, file), header=12)
                                    df_esb = df_esb[df_esb['Type'].isin(['Ala Carte', 'Package Head'])]
                            df_esb = df_esb[['Branch','Sales Date','Menu Name','Menu Code','Qty']].assign(**{'Menu Code': df_esb['Menu Code'].astype(str)}).merge(df_4121[['Kode Barang Grup Barang','Kode Barang','Kuantitas']].rename(columns={'Kode Barang Grup Barang':'Menu Code'}), on='Menu Code', how='left')
                            df_esb = df_esb.assign(**{'Kuantitas_ESB':df_esb['Kuantitas'] * df_esb['Qty'],
                                            'Branch':df_esb['Branch'].str.extract(r'\.(.+)')}).groupby(['Branch','Sales Date','Kode Barang'])[['Kuantitas_ESB']].sum().reset_index().merge(
                                df_2205.assign(**{'Nama Pelanggan':df_2205['Nama Pelanggan'].str.extract(r'\(([^()]*)\)')[0].values})[df_2205['Nomor #'].str.startswith('ACR')][['Nama Pelanggan','Tanggal','Kode #','Kuantitas']].rename(
                                    columns={'Nama Pelanggan':'Branch','Tanggal':'Sales Date','Kode #':'Kode Barang','Kuantitas':'Kuantitas_GIS'}),
                                on=['Branch','Sales Date','Kode Barang'], how='outer').merge(
                                db_2205.rename(columns={'Kode #':'Kode Barang'}), on='Kode Barang', how='left')
                            df_esb['Selisih'] = (df_esb['Kuantitas_ESB'].fillna(0) - df_esb['Kuantitas_GIS'].fillna(0)).round(2).replace(-0.0,0.0)

                            st.success('Success',icon='✅')
                            st.download_button(
                                label="Download",
                                data=to_excel(df_esb),
                                file_name=f'REKAP SALES ESB & GIS_{get_current_time_gmt7()}.xlsx',
                                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                            )

            except Exception as e:
                st.error('Failed', icon='🛑')
                try:
                    st.error(file)
                    st.error(e)
                    traceback.print_exc()
                except:
                    st.error(e)
                    traceback.print_exc()


if selected_option == 'OCR-SJ':
    try:
        st.write(' ')
        df_sj = st.session_state.df_sj
        df_file_sj = st.session_state.df_file_sj 
        df_sj_cek = df_sj[['Nomor #Kirim','Gudang #Terima','Tanggal #Kirim','Nama File_SJ']].drop_duplicates().merge(
            df_sj[['Nomor #Kirim', 'Selisih']].replace('',1)
            .groupby(['Nomor #Kirim'])[['Selisih']].sum()
            .apply(lambda row: True if row['Selisih'] == 0 else False, axis=1).reset_index().rename(columns={0:'Check'}),
            how = 'left', on='Nomor #Kirim')
        df_sj_cek['Note'] = ''
        if not st.session_state.dfh.empty:
            df_sj_cek = pd.concat([st.session_state.dfh,df_sj_cek]).sort_values(['Check','Nama File_SJ','Gudang #Terima','Nomor #Kirim'], ascending=[False,False,False,False]).drop_duplicates(['Nomor #Kirim','Tanggal #Kirim','Gudang #Terima'],keep='first')
            
        gb = GridOptionsBuilder.from_dataframe(df_sj_cek)
        
        gb.configure_default_column(filter=True, sortable=True,
            flex=0,
            resizable=True,
            minWidth=130,
            enableValue=True,
            enableRowGroup=True
        )
        editable_columns = ['Check', 'Note']
        for col in editable_columns:
            gb.configure_column(col, editable=True)
        gb = gb.build()
        AgGrid(
            df_sj_cek,
            gb, update_mode=GridUpdateMode.NO_UPDATE,
            enable_enterprise_modules=enable_enterprise,
            allow_unsafe_jscode=True,
            fit_columns_on_grid_load=True,
            #license_key=license_key,
            #key=key,
        )

        gb = GridOptionsBuilder.from_dataframe(df_sj)
        
        gb.configure_default_column(filter=True, sortable=True,
            flex=1,
            minWidth=130,
            enableValue=True,
            enableRowGroup=True,
            enablePivot=True
        )
        #gb.configure_column('Check', cellRenderer='agCheckboxCellRenderer')
        gb = gb.build()

        # Menambahkan pengaturan untuk auto group column (kolom grup otomatis)
        gb['autoGroupColumnDef'] = {
            'minWidth': 200,
            'pinned': 'left'
        }
        gb['sideBar'] = {
            "toolPanels": ['columns', 'filters']  # Tidak ada panel default saat pertama kali dibuka
        }

        gb['pivotMode'] = False
        gb['pivotPanelShow'] = "always"
        kol = st.columns(2)
        with kol[0]:
            st.markdown("<div style='height:150px;'></div>", unsafe_allow_html=True)
            AgGrid(
                df_sj,
                gb, update_mode=GridUpdateMode.NO_UPDATE,
                enable_enterprise_modules=enable_enterprise,
                allow_unsafe_jscode=True,
                height=700,
                #license_key=license_key,
                #key=key,
            )
        with kol[1]:
                
                filename = st.selectbox("Pilih File",df_file_sj['Nama File_SJ'].unique().tolist(),key='select_file')
                images = df_file_sj[df_file_sj['Nama File_SJ']==filename]['Images'].values[0]
                if 'halaman' not in st.session_state:
                    st.session_state.halaman = 0

                hal, col_kiri, col_kanan = st.columns([3,3,5])
                with col_kiri:
                    if st.button("⬅️ Previous") and st.session_state.halaman > 0:
                        st.session_state.halaman -= 1
                with col_kanan:
                    if st.button("Next ➡️") and st.session_state.halaman < len(images) - 1:
                        st.session_state.halaman += 1
                with hal:
                    st.write(f"Halaman {st.session_state.halaman + 1} dari {len(images)}")
                img = images[st.session_state.halaman]
                lebar_target = 590
                width_asli, height_asli = img.size
                tinggi_target = int((height_asli / width_asli) * lebar_target)

                
                image_zoom(
                    img,
                    mode="both",
                    size=(lebar_target, tinggi_target),
                    keep_aspect_ratio=True,
                    keep_resolution=True,
                    zoom_factor=3.0,
                    increment=0.2
                )


    except Exception as ex:
        print('')
