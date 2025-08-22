import pandas as pd
import streamlit as st
import numpy as np
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, ColumnsAutoSizeMode, GridUpdateMode
import xlsxwriter
import os
from io import BytesIO

css = '''
<style>
    section:nth-of-type(2) div.stButton > button {
        background-color: #af0d1a !important;  /* Merah saat hover */
        color: white !important;
    }   
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }

    .stTabs [data-baseweb="tab"] {
        height: 30px;
        border-radius: 7px 7px 0 0; /* Rounded top */
        color: black;
        white-space: pre-wrap;
        background-color: #cccccc;
        gap: 1px;
        padding-right: 10px;
        padding-left: 10px;
        padding-top: 0px;
        padding-bottom: 0px;
    }

    .stTabs [aria-selected="true"] {
        background-color: #af0d1a !important;
        color: white;
        font-weight: bold;
    }
</style>
'''

st.markdown(css, unsafe_allow_html=True)
custom_css = {
    ".ag-root-wrapper": {
        "border-radius": "10px",     # Ujung luar
        "overflow": "hidden",        # Pastikan kontennya tidak melebihi radius
    },
    ".ag-header-container": {
        "border-top-left-radius": "10px",
        "border-top-right-radius": "10px",
    },
    ".ag-center-cols-clipper": {
        "border-bottom-left-radius": "10px",
        "border-bottom-right-radius": "10px",
    },
    ".ag-header": {
        "background-color": "#001C53", 
        "color": "white",
        "font-weight": "bold"
    },
    ".ag-header-cell-label": {
        "font-size": "10px",
        "padding": "2px"
    },
    ".ag-cell": {
        "font-size": "10px"
    }
}


kol = st.columns([0.6,1.3])
with kol[0]:
    st.title('COM Monitoring')
st.markdown("<hr style='margin:0; padding:0; border:1px solid #ccc'>", unsafe_allow_html=True)


if 'button_clicked' not in st.session_state:
    st.session_state.button_clicked = False

def reset_button_state():
    st.session_state.button_clicked = False


key = "enterprise_disabled_grid"
license_key = None
enable_enterprise = True
if enable_enterprise:
    key = "enterprise_enabled_grid"
    license_key = license_key

def move_column(df, col_to_move, col_target, after=True):
    cols = list(df.columns)
    cols.remove(col_to_move)
    
    idx = cols.index(col_target)
    if after:
        idx += 1
    cols.insert(idx, col_to_move)
    
    return df[cols]

def to_excel(df, sheet_name='Sheet1',output='Data/COM Monitoring/Output/REKAP MENTAH.xlsx'):
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        header_format = workbook.add_format({'border': 0, 'bold': False, 'font_size': 12})
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
with kol[1]:
    st.write(' ')
    st.write(' ')
    if st.button('Update'):
        all_dfs = []
        for file in os.listdir('Data/COM Monitoring/Raw Data/REKAP MENTAH'):
            if file.endswith('.xlsx'):
                    df = pd.read_excel('Data/COM Monitoring/Raw Data/REKAP MENTAH'+'/'+file, sheet_name='REKAP MENTAH')
                    if 'CABANG' not in df.columns:
                        df = df.loc[:,[x for x in df.columns if 'Unnamed' not in str(x)][:-1]].fillna('')
                        df['CABANG'] = file.replace('.xlsx','').split('-')[0]
                    all_dfs.append(df)
            
        all_dfs = pd.concat(all_dfs, ignore_index=True)
        all_dfs = all_dfs[['CABANG', 'KATEGORI','SOURCE DATA','JENIS','STATUS','NAMA BARANG','KETERANGAN','PENYEBAB TERJADINYA WASTE','SATUAN','1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13',
        '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25',
        '26', '27', '28', '29', '30', '31']]
        to_excel(all_dfs, sheet_name="REKAP MENTAH")

        for file in os.listdir('Data/COM Monitoring/Raw Data'):
            if file.startswith('2205'):
                all_dfs = []
                df_2205 = pd.read_excel('Data/COM Monitoring/Raw Data'+'/'+file, skiprows=4)
                df_2205 = df_2205.loc[:, ~df_2205.columns.str.startswith('Unnamed:')].iloc[:-5,:]
                df_2205 = df_2205[df_2205['Nomor #'].str.startswith(('ACR','SRT'))]
                df_2205.assign(**{'Nama Pelanggan':df_2205['Nama Pelanggan'].str.extract(r'\(([^()]*)\)')[0].values,
                    'TANGGAL':(df_2205['Tanggal'].dt.strftime('%d').astype('int').astype('str'))}).rename(columns={'Nama Pelanggan':'CABANG','Nama Barang':'NAMA BARANG','Kuantitas':'BOM'}).groupby(['CABANG','TANGGAL','NAMA BARANG'])[['BOM']].sum().reset_index().to_csv('Data/COM Monitoring/Output/.csv/2205.csv',index=False)
                
            if file.startswith('4205'):
                df_4205 = pd.read_excel('Data/COM Monitoring/Raw Data'+'/'+file, skiprows=4)
                df_4205 = df_4205.iloc[:-5]
                df_4205 = df_4205.loc[:, ~df_4205.columns.str.startswith('Unnamed')]
                df_4205['Tanggal #Terima'] = pd.to_datetime(df_4205['Tanggal #Terima'], format='%d-%b-%y')
                df_4205 = df_4205.assign(**{'Gudang #Terima':df_4205['Gudang #Terima'].str.extract(r'\(([^()]*)\)')[0].values,
                                'TANGGAL':(df_4205['Tanggal #Terima'].dt.strftime('%d').astype('int').astype('str'))}).rename(columns={'Gudang #Terima':'CABANG','Nama Barang':'NAMA BARANG','#Qty. Terkecil':'4205'}).groupby(['CABANG','TANGGAL','NAMA BARANG'])[['4205']].sum().reset_index().to_csv('Data/COM Monitoring/Output/.csv/4205.csv',index=False)
                
            if file.startswith('4217'):
                df_4217 = pd.read_excel('Data/COM Monitoring/Raw Data'+'/'+file, header=4).fillna('')
                df_4217 = df_4217.drop(columns=[x for x in df_4217.reset_index().T[(df_4217.reset_index().T[1]=='')].index if 'Unnamed' in x])
                df_4217.columns = df_4217.T.reset_index()['index'].apply(lambda x: np.nan if 'Unnamed' in x else x).ffill().values
                df_4217 = df_4217.iloc[1:,:-3]

                df_melted =pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang'], 
                    value_vars=df_4217.columns[6:].values,
                    var_name='Nama Cabang', value_name='subtotalStok').reset_index(drop=True)

                df_melted2 = pd.melt(pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Satuan #1','Satuan #2','Satuan #3'], 
                    value_vars=df_4217.columns[6:].values,
                    var_name='Nama Cabang', value_name='subtotalStok').drop_duplicates(),
                    id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Nama Cabang','subtotalStok'],
                    var_name='Variabel', value_name='Satuan')

                df_melted2 = df_melted2[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Satuan','Variabel']].drop_duplicates().reset_index(drop=True)

                df_melted = df_melted.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)
                df_melted2 = df_melted2.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)

                df_4217_final = pd.concat([df_melted2, df_melted[['subtotalStok']]], axis=1)
                df_4217_final = df_4217_final.rename(columns={'Variabel':'Kategori'})[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Kategori','Satuan','subtotalStok']]
                df_4217_final['Kode Barang'] = df_4217_final['Kode Barang'].astype('int')
                df_4217_final['subtotalStok'] = df_4217_final['subtotalStok'].astype('float')

                df_4217_final=df_4217_final[df_4217_final['Kategori']   ==  "Satuan #1"].rename(columns={"Kategori":"Satuan","subtotalStok":"Saldo Akhir"})

                df_4217_final=df_4217_final.loc[:,["Kategori Barang","Kode Barang","Nama Barang","Satuan","Saldo Akhir","Nama Cabang"]]
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
                df_4217_final.assign(**{'Nama Cabang':df_4217_final['Nama Cabang'].str.extract(r'\(([^()]*)\)')[0].values})[['Nama Cabang','Nama Barang','Saldo Akhir']].rename(columns={'Nama Cabang':'CABANG','Nama Barang':'NAMA BARANG','Saldo Akhir':'SO_4217'}).to_csv('Data/COM Monitoring/Output/.csv/4217.csv',index=False)

try:
    @st.cache_data
    def load_data():
        df = pd.read_excel('Data/COM Monitoring/Output/REKAP MENTAH.xlsx')
        return df

    df = load_data()
    df_4217 = pd.read_csv('Data/COM Monitoring/Output/.csv/4217.csv')
    df_4217['NAMA BARANG'] = df_4217['NAMA BARANG'].str.replace('MINYAK MIE SHALLOT OIL','MINYAK MIE (V.20)')
    df_4217 = df_4217.groupby(['CABANG','NAMA BARANG'])[['SO_4217']].sum().reset_index()
    df_4205 = pd.read_csv('Data/COM Monitoring/Output/.csv/4205.csv')
    df_4205['NAMA BARANG'] = df_4205['NAMA BARANG'].str.replace('MINYAK MIE SHALLOT OIL','MINYAK MIE (V.20)')
    df_4205 = df_4205.groupby(['CABANG','TANGGAL','NAMA BARANG'])[['4205']].sum().reset_index()
    df_2205 = pd.read_csv('Data/COM Monitoring/Output/.csv/2205.csv')

except FileNotFoundError:
    st.error("File 'REKAP MENTAH.xlsx' tidak ditemukan")   
    st.stop()             

angka_cols = df.loc[[0], 'SATUAN':].columns[1:]
#df_bom = df.assign(TOTAL=df[angka_cols].sum(axis=1))[df['JENIS']=='SO'].groupby(['CABANG','NAMA BARANG','KATEGORI','SATUAN'])[['TOTAL']].sum().reset_index()
df_bom = pd.concat([pd.DataFrame({'Kolom':['CABANG','NAMA BARANG']+angka_cols.tolist()}).set_index('Kolom').T, df_2205.assign(TANGGAL = df_2205['TANGGAL'].astype(str)).pivot(index=['CABANG','NAMA BARANG'], columns='TANGGAL', values='BOM').reset_index()])
df_bom = df_bom.assign(TOTAL=df_bom[angka_cols].sum(axis=1)).groupby(['CABANG','NAMA BARANG'])[['TOTAL']].sum().reset_index()

tanggal = st.number_input("Tanggal:", value=7, min_value=1, step=1, key='tanggal')
angka_cols=angka_cols[:tanggal]

tab = st.tabs(['DEVIASI','WASTE','SUSUT','TRIAL'])


with tab[0]:
    tanggal = str(tanggal)
    df_so = df[(df['JENIS']=='SO') & (df['KATEGORI'].str.upper().isin(['COM DEVIASI - RESTO', 'COM CONSUME - RESTO', 'BIAYA PACKAGING - RESTO']))]    
    df_so = df_so.groupby(['CABANG','NAMA BARANG','KATEGORI','SATUAN'])[angka_cols].sum().reset_index().replace([0], np.nan)
    df_deviasi = df_so.assign(Cabang=df_so['CABANG'].str[5:])[['Cabang','CABANG','NAMA BARANG','KATEGORI','SATUAN']+[tanggal]].rename(columns={tanggal:'SO'}).merge(df_4217.rename(columns={'CABANG':'Cabang'}), on=['Cabang','NAMA BARANG'], how='left').merge(
    df_2205[df_2205['TANGGAL']<=int(tanggal)].groupby(['CABANG','NAMA BARANG'])[['BOM']].sum().reset_index().rename(columns={'CABANG':'Cabang'}), on=['Cabang','NAMA BARANG'], how='left').merge(df_4205[df_4205['TANGGAL']==int(tanggal)].rename(columns={'CABANG':'Cabang'}), on=['Cabang','NAMA BARANG'], how='left').drop(columns=['Cabang']).fillna(0)
    df_deviasi['DEVIASI'] = df_deviasi['SO'] - df_deviasi['SO_4217']
    df_deviasi['BOM'] = -df_deviasi['BOM']
    df_deviasi['COM'] = df_deviasi['BOM'] + df_deviasi['DEVIASI']
    df_deviasi['%'] = -(df_deviasi['DEVIASI']/df_deviasi['BOM'])
    df_deviasi = df_deviasi.replace([np.nan, np.inf,-np.inf],0)[['CABANG','NAMA BARANG','KATEGORI','SATUAN','BOM','COM','DEVIASI','%','4205']].merge(df_so[['CABANG','NAMA BARANG']+[str(i) for i in range(int(tanggal)-6,int(tanggal)+1)]], on=['CABANG','NAMA BARANG'],how='left').fillna(0)
    df_deviasi['% '] = ((df_deviasi.iloc[:,-1] - df_deviasi.iloc[:,-2] )/ df_deviasi.iloc[:,-2]).fillna(0)
    df_deviasi['subtotal'] = df_deviasi.assign(**{'DEVIASI':df_deviasi.apply(lambda x: abs(x['DEVIASI']) if x['BOM']!=0 else 0, axis=1)}).groupby('NAMA BARANG')['DEVIASI'].transform('sum')
    df_deviasi = df_deviasi.sort_values(['subtotal','NAMA BARANG'],ascending=[False,True]).reset_index(drop=True).drop(columns=['subtotal'])
    df_deviasi = df_deviasi[df_deviasi.drop(columns='4205').columns.to_list()+['4205']]

    def export_with_excel_icons_inplace(df, angka_cols, filename='output.xlsx'):
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            def get_excel_column_letter(n):
                result = ''
                while n >= 0:
                    result = chr(n % 26 + 65) + result
                    n = n // 26 - 1
                return result

            green = workbook.add_format({'font_color': 'green'})
            red = workbook.add_format({'font_color': 'red'})
            yellow = workbook.add_format({'font_color': 'black'})

            for i in range(int(tanggal)-5, int(tanggal)+1):  # mulai dari col kedua tanggal
                curr_col = f"{i}"
                prev_col = f"{i-1}"
                if curr_col in df.columns and prev_col in df.columns:
                    curr_idx = df.columns.get_loc(curr_col)
                    prev_idx = df.columns.get_loc(prev_col) 
                    curr_letter = get_excel_column_letter(curr_idx)
                    prev_letter = get_excel_column_letter(prev_idx)

                    # range baris (Excel mulai baris data di 2)
                    start_row = 2
                    end_row = len(df) + 1
                    cell_range = f"{curr_letter}{start_row}:{curr_letter}{end_row}"

                    # Formula Excel untuk bandingkan isi sel di kolom ini (misal C2) dengan sel di kolom sebelumnya (B2)
                    # Misal C2 > B2 maka hijau
                    formula_green = f'={curr_letter}2>{prev_letter}2'
                    formula_red = f'={curr_letter}2<{prev_letter}2'
                    formula_yellow = f'={curr_letter}2={prev_letter}2'

                    worksheet.conditional_format(cell_range, {
                        'type': 'formula',
                        'criteria': formula_green,
                        'format': green
                    })
                    worksheet.conditional_format(cell_range, {
                        'type': 'formula',
                        'criteria': formula_red,
                        'format': red
                    })
                    worksheet.conditional_format(cell_range, {
                        'type': 'formula',
                        'criteria': formula_yellow,
                        'format': yellow
                    })


    gb = GridOptionsBuilder.from_dataframe(df_deviasi)

    for col in ['COM','BOM','DEVIASI','%','4205']+[str(i) for i in range(int(tanggal)-6,int(tanggal)+1)]:
        value_formatter_jscode = JsCode("""
        function(params) {
            if (params.value == null || isNaN(params.value)) return '';
            return (params.value).toLocaleString();
        }
        """)
        gb.configure_column(col, valueFormatter=value_formatter_jscode)

    def make_value_formatter(curr_col, prev_col):
        return JsCode(f"""
            function(params) {{
                const val = params.value;
                const prev = params.data['{prev_col}'];
                if (val == null || prev == null) return '';
                const diff = val - prev;
                let icon = '';
                if (diff > 0) icon = ' 🟢';
                else if (diff < 0) icon = ' 🔴';
                else icon = '';
                return (params.value).toLocaleString() + icon;
            }}
        """)

    for i in range(int(tanggal)-5, int(tanggal)+1):
        col = f"{i}"
        prev = f"{i - 1}"
        formatter = make_value_formatter(col, prev)
        gb.configure_column(col, valueFormatter=formatter)

    cell_style_deviasi = JsCode("""
        function(params) {
            if (params.value < 0) {
                return { color: 'red', fontWeight: 'bold' };
            } else {
                return { color: 'black' };
            }
        }
    """)
    
    for col in ['DEVIASI','%']:
        gb.configure_column(
            col,
            cellStyle=cell_style_deviasi
        )

    percent_formatter = JsCode("""
        function(params) {
            if (params.value == null) return '';
            return (params.value * 100).toFixed(1) + '%';
        }
    """)

    for col in ['%','% ']:
        gb.configure_column(col, valueFormatter=percent_formatter)
    for col in ['CABANG','NAMA BARANG','KATEGORI','SATUAN']:
        gb.configure_column(col, pinned='left')
    gb.configure_default_column(filter=True, sortable=True,
        flex=1,
        minWidth=100,
        enableValue=True,
        enableRowGroup=True,
        enablePivot=True
    )

    gridOptions = gb.build()
    gridOptions['autoGroupColumnDef'] = {
        'minWidth': 200,
        'pinned': 'left'
    }
    gridOptions['sideBar'] = {
        "toolPanels": ['columns', 'filters'] 
    }

    gridOptions['pivotMode'] = False
    gridOptions['pivotPanelShow'] = "always"
    if "columnDefs" in gridOptions:
        for colDef in gridOptions["columnDefs"]:
            colDef["autoWidth"] = True

    AgGrid(
        df_deviasi,
        gridOptions, update_mode=GridUpdateMode.NO_UPDATE,
        enable_enterprise_modules=enable_enterprise,
        allow_unsafe_jscode=True,custom_css=custom_css)
    
    if st.button('Export', key='export-deviasi'):
        export_with_excel_icons_inplace(df_deviasi, angka_cols,'Data/COM Monitoring/Output/REKAP_DEVIASI.xlsx')

with tab[1]:
    df_2205 = df_2205[df_2205['TANGGAL']<=int(tanggal)]
    kol = st.columns([1,1,4])
    with kol[0]:
        ts_waste_down = st.number_input("Threshold Lower (Z-Score):", value=1.0, min_value=1.0, step=0.1, key='threshold_waste2')
    with kol[1]:
        ts_waste_up = st.number_input("Threshold Upper (Z-Score):", value=1.0, min_value=1.0, step=0.1, key='threshold_waste')
    df_waste = df[(df['JENIS']=='WASTE') & (df['KATEGORI'].str.upper().isin(['COM DEVIASI - RESTO', 'COM CONSUME - RESTO', 'BIAYA PACKAGING - RESTO']))]
    keterangan_waste = st.multiselect("KETERANGAN:", ['All'] + df_waste['KETERANGAN'].unique().tolist(), default=['All'], on_change=reset_button_state)
    df_waste = df_waste[(df_waste['KETERANGAN'].isin(df_waste['KETERANGAN'].unique() if keterangan_waste==['All'] else keterangan_waste))].groupby(['CABANG','NAMA BARANG','KATEGORI','SATUAN'])[angka_cols].sum().reset_index().fillna(0)

    df_waste['mean'] = df_waste[angka_cols].mean(axis=1)
    df_waste['std'] = df_waste[angka_cols].std(axis=1)
    df_waste['TOTAL '] = df_waste[angka_cols].sum(axis=1)
    df_waste['subtotal'] = df_waste.groupby('NAMA BARANG')['TOTAL '].transform('sum')
    df_waste = df_waste.sort_values(['subtotal','NAMA BARANG'],ascending=[False,True]).reset_index(drop=True)

    merged = pd.merge(df_waste[['CABANG','NAMA BARANG','SATUAN']+[col for col in angka_cols]].assign(Cabang=df_waste['CABANG'].str[5:]), df_2205.assign(TANGGAL = df_2205['TANGGAL'].astype(str)).pivot(index=['CABANG','NAMA BARANG'], columns='TANGGAL',values='BOM').reset_index().rename(columns={'CABANG':'Cabang'}), on=['Cabang', 'NAMA BARANG'], how='left').fillna(0)

    for col in angka_cols:
        merged[col] = merged[f"{col}_x"] / merged[f"{col}_y"] * 100

    result = merged[['CABANG','NAMA BARANG','SATUAN'] + [col for col in angka_cols]]
    result['mean'] = result[angka_cols].mean(axis=1)
    result['std'] = result[angka_cols].std(axis=1)
    result['TOTAL '] = result[angka_cols].sum(axis=1)

    df_waste = move_column(df_waste, col_to_move='TOTAL ', col_target='SATUAN', after=True)
    
    color1 = result.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) < -ts_waste_down))if (x['std'] != 0 and x['SATUAN'] not in ['PCS']) else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )

    color4 = result.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) > ts_waste_up))if (x['std'] != 0)  else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )

    color2 = df_waste.apply(
        lambda x: x[angka_cols] < 0,
        axis=1, result_type='expand'
    )

    color3 = df_waste.assign(CABANG=df_waste['CABANG'].str[5:]).merge(df_bom, how='left', on=['CABANG','NAMA BARANG']).fillna(0).apply(
        lambda x: np.abs(x[angka_cols]) > 0 if x['TOTAL'] == 0 else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )
    color3.columns = angka_cols
    color = color1 | color2 | color3 | color4
    color.columns = angka_cols

    for col in angka_cols:
        df_waste[f"{col}_outlier"] = color[col]

    df_waste['is_outlier'] = df_waste[[f"{col}_outlier" for col in angka_cols]].any(axis=1)

    def export_with_color(df, angka_cols, filename='output.xlsx'):
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            df[['CABANG','NAMA BARANG','KATEGORI','SATUAN','TOTAL ']+angka_cols.to_list()+['is_outlier']].to_excel(writer, index=False, sheet_name='Sheet1')
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']

            red_format = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'white'})

            for row_idx, row in df.iterrows():
                for col_idx, col in enumerate(angka_cols):
                    if row[f"{col}_outlier"]:
                        worksheet.write(row_idx+1, df.columns.get_loc(col), row[col], red_format)

    outlier_waste = st.radio(
        "Hanya Tampilkan Outlier:",
        ("Iya","Tidak"), key='outlier_waste'
    )

    if outlier_waste == "Iya": 
        df_waste = df_waste[df_waste['is_outlier']]

    gb = GridOptionsBuilder.from_dataframe(df_waste)
    for col in angka_cols.to_list()+['TOTAL ']:
        cell_style_jscode = JsCode(f"""
        function(params) {{
            if (params.data['{col}_outlier']) {{
                return {{'backgroundColor': 'red', 'color': 'white'}};
            }}
        }}
        """)
        gb.configure_column(col, cellStyle=cell_style_jscode)
        
        value_formatter_jscode = JsCode("""
        function(params) {
            if (params.value == null || isNaN(params.value)) return '';
            return (params.value).toLocaleString();
        }
        """)
        gb.configure_column(col, valueFormatter=value_formatter_jscode)

    for col in angka_cols.to_list():
        gb.configure_column(f"{col}_outlier", hide=True, suppressColumnsToolPanel=True)
    for col in ['mean','std','subtotal','is_outlier']:
        gb.configure_column(col, hide=True, suppressColumnsToolPanel=True)
    for col in ['CABANG','NAMA BARANG','KATEGORI','SATUAN']:
        gb.configure_column(col, pinned='left')
    gb.configure_default_column(filter=True, sortable=True,
        flex=1,
        minWidth=100,
        enableValue=True,
        enableRowGroup=True,
        enablePivot=True
    )

    gridOptions = gb.build()
    gridOptions['autoGroupColumnDef'] = {
        'minWidth': 200,
        'pinned': 'left'
    }
    gridOptions['sideBar'] = {
        "toolPanels": ['columns', 'filters']  
    }

    gridOptions['pivotMode'] = False
    gridOptions['pivotPanelShow'] = "always"
    if "columnDefs" in gridOptions:
        for colDef in gridOptions["columnDefs"]:
            colDef["autoWidth"] = True

    AgGrid(
        df_waste,
        gridOptions, update_mode=GridUpdateMode.NO_UPDATE,
        enable_enterprise_modules=enable_enterprise,
        allow_unsafe_jscode=True,custom_css=custom_css)
    if st.button('Export', key='export-waste'):
        export_with_color(df_waste, angka_cols,'Data/COM Monitoring/Output/REKAP_WASTE.xlsx')

with tab[2]:
    kol = st.columns([1,1,4])
    with kol[0]:
        ts_susut_down = st.number_input("Threshold Lower (Z-Score):", value=1.0, min_value=1.0, step=0.1, key='threshold_susut2')
    with kol[1]:
        ts_susut_up = st.number_input("Threshold Upper (Z-Score):", value=1.0, min_value=1.0, step=0.1, key='threshold_susut')

    br_susut = pd.read_excel('Data/COM Monitoring/Database/BARANG SUSUT.xlsx')
    df_susut = df[(df['JENIS']=='SUSUT') & (df['KATEGORI'].str.upper().isin(['COM DEVIASI - RESTO', 'COM CONSUME - RESTO', 'BIAYA PACKAGING - RESTO']))].groupby(['CABANG','NAMA BARANG','KATEGORI','SATUAN'])[angka_cols].sum().reset_index().fillna(0)
    df_susut['mean'] = df_susut[angka_cols].mean(axis=1)
    df_susut['std'] = df_susut[angka_cols].std(axis=1)
    df_susut['TOTAL '] = df_susut[angka_cols].sum(axis=1)
    df_susut['subtotal'] = df_susut.groupby('NAMA BARANG')['TOTAL '].transform('sum')
    df_susut = df_susut.sort_values(['subtotal','NAMA BARANG'],ascending=[False,True]).reset_index(drop=True)

    df_susut = move_column(df_susut, col_to_move='TOTAL ', col_target='SATUAN', after=True)

    merged = pd.merge(df_susut[['CABANG','NAMA BARANG','SATUAN']+[col for col in angka_cols]].assign(Cabang=df_susut['CABANG'].str[5:]), df_2205.assign(TANGGAL = df_2205['TANGGAL'].astype(str)).pivot(index=['CABANG','NAMA BARANG'], columns='TANGGAL',values='BOM').reset_index().rename(columns={'CABANG':'Cabang'}), on=['Cabang', 'NAMA BARANG'], how='left').fillna(0)

    for col in angka_cols:
        merged[col] = merged[f"{col}_x"] / merged[f"{col}_y"] * 100

    result = merged[['CABANG','NAMA BARANG','SATUAN'] + [col for col in angka_cols]]
    result['mean'] = result[angka_cols].mean(axis=1)
    result['std'] = result[angka_cols].std(axis=1)
    result['TOTAL '] = result[angka_cols].sum(axis=1)

    color1 = result.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) < -ts_susut_down))if ((x['std'] != 0) and (x['SATUAN'] not in ['PCS']) & (x['NAMA BARANG'] not in ['CABE RAWIT - RESTO (V.20)'])) else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )

    color4 = result.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) > ts_susut_up))if ((x['std'] != 0) & (x['NAMA BARANG'] not in ['CABE RAWIT - RESTO (V.20)']))  else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )
    color2 = df_susut.apply(
        lambda x: x[angka_cols] < 0,
        axis=1, result_type='expand'
    )

    color3 = df_susut.assign(CABANG=df_susut['CABANG'].str[5:]).merge(df_bom, how='left', on=['CABANG','NAMA BARANG']).fillna(0).apply(
        lambda x: np.abs(x[angka_cols]) > 0 if x['TOTAL'] == 0 else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )
    color3.columns = angka_cols

    color = color1 | color2 | color3 | color4
    not_in_br_susut = ~df_susut['NAMA BARANG'].isin(br_susut['NAMA BARANG'])
    
    color4 = df_susut.apply(
        lambda x: [val > 0 and not_in_br_susut.loc[x.name] for val in x[angka_cols]],
        axis=1, result_type='expand'
    )
    color4 = color4.apply(lambda x: [val > 0 for val in x], axis=1, result_type='expand')
    color4.columns = angka_cols

    df_susut_cabe = (df_susut[angka_cols] - df_susut.assign(CABANG=df_susut['CABANG'].str[5:])[['CABANG','NAMA BARANG']].merge(pd.concat([pd.DataFrame({'Kolom':['CABANG','NAMA BARANG']+angka_cols.tolist()}).set_index('Kolom').T,((df_2205.assign(TANGGAL = df_2205['TANGGAL'].astype(str))[df_2205['NAMA BARANG'].str.contains('CABE RAWIT')].pivot(index=['CABANG','NAMA BARANG'], columns='TANGGAL', values='BOM')/0.89)*0.11).reset_index()]),
                                                                                    how='left')[angka_cols])
    df_susut_cabe['mean'] = df_susut_cabe[angka_cols].mean(axis=1)
    df_susut_cabe['std'] = df_susut_cabe[angka_cols].std(axis=1)
    color1 = df_susut_cabe.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) < -ts_susut_down))if ((x['std'] != 0)) else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )

    color2 = df_susut_cabe.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) > ts_susut_up))if (x['std'] != 0)  else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )

    color = color | color4 | color1 | color2
    color.columns = angka_cols

    for col in angka_cols:
        df_susut[f"{col}_outlier"] = color[col]

    df_susut['is_outlier'] = df_susut[[f"{col}_outlier" for col in angka_cols]].any(axis=1)

    def export_with_color(df, angka_cols, filename='output.xlsx'):
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            df[['CABANG','NAMA BARANG','KATEGORI','SATUAN','TOTAL ']+angka_cols.to_list()+['is_outlier']].to_excel(writer, index=False, sheet_name='Sheet1')
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']

            red_format = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'white'})

            for row_idx, row in df.iterrows():
                for col_idx, col in enumerate(angka_cols):
                    if row[f"{col}_outlier"]:
                        worksheet.write(row_idx+1, df.columns.get_loc(col), row[col], red_format)

    outlier_susut = st.radio(
        "Hanya Tampilkan Outlier:",
        ("Iya","Tidak"), key='outlier_susut'
    )

    if outlier_susut == "Iya": 
        df_susut = df_susut[df_susut['is_outlier']]

    gb = GridOptionsBuilder.from_dataframe(df_susut)
    for col in angka_cols.to_list() + ['TOTAL ']:
        cell_style_jscode = JsCode(f"""
        function(params) {{
            if (params.data['{col}_outlier']) {{
                return {{'backgroundColor': 'red', 'color': 'white'}};
            }}
        }}
        """)
        gb.configure_column(col, cellStyle=cell_style_jscode)
        
        value_formatter_jscode = JsCode("""
        function(params) {
            if (params.value == null || isNaN(params.value)) return '';
            return (params.value).toLocaleString();
        }
        """)
        gb.configure_column(col, valueFormatter=value_formatter_jscode)
        
    for col in angka_cols.to_list():
        gb.configure_column(f"{col}_outlier", hide=True, suppressColumnsToolPanel=True)
    for col in ['mean','std','subtotal','is_outlier']:
        gb.configure_column(col, hide=True, suppressColumnsToolPanel=True)
    for col in ['CABANG','NAMA BARANG','KATEGORI','SATUAN']:
        gb.configure_column(col, pinned='left')
    gb.configure_default_column(filter=True, sortable=True,
        flex=1,
        minWidth=100,
        enableValue=True,
        enableRowGroup=True,
        enablePivot=True
    )

    gridOptions = gb.build()
    gridOptions['autoGroupColumnDef'] = {
        'minWidth': 200,
        'pinned': 'left'
    }
    gridOptions['sideBar'] = {
        "toolPanels": ['columns', 'filters'] 
    }

    gridOptions['pivotMode'] = False
    gridOptions['pivotPanelShow'] = "always"
    if "columnDefs" in gridOptions:
        for colDef in gridOptions["columnDefs"]:
            colDef["autoWidth"] = True
 
    AgGrid(
        df_susut,
        gridOptions, update_mode=GridUpdateMode.NO_UPDATE,
        enable_enterprise_modules=enable_enterprise,
        allow_unsafe_jscode=True,custom_css=custom_css)
    if st.button('Export', key='export-susut'):
        export_with_color(df_susut, angka_cols,'Data/COM Monitoring/Output/REKAP_SUSUT.xlsx')

with tab[3]:
    kol = st.columns([1,1,4])
    with kol[0]:
        ts_trial_down = st.number_input("Threshold Lower (Z-Score):", value=1.0, min_value=1.0, step=0.1, key='threshold_trial2')
    with kol[1]:
        ts_trial_up = st.number_input("Threshold Upper (Z-Score):", value=1.0, min_value=1.0, step=0.1, key='threshold_trial')

    df_trial = df[(df['JENIS']=='TRIAL') & (df['KATEGORI'].str.upper().isin(['COM DEVIASI - RESTO', 'COM CONSUME - RESTO', 'BIAYA PACKAGING - RESTO']))].groupby(['CABANG','NAMA BARANG','KATEGORI','SATUAN'])[angka_cols].sum().reset_index().fillna(0)
    df_trial['mean'] = df_trial[angka_cols].mean(axis=1)
    df_trial['std'] = df_trial[angka_cols].std(axis=1)
    df_trial['TOTAL '] = df_trial[angka_cols].sum(axis=1)
    df_trial['subtotal'] = df_trial.groupby('NAMA BARANG')['TOTAL '].transform('sum')
    df_trial = df_trial.sort_values(['subtotal','NAMA BARANG'],ascending=[False,True]).reset_index(drop=True)
    df_trial = move_column(df_trial, col_to_move='TOTAL ', col_target='SATUAN', after=True)

    color1 = df_trial.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) < -ts_trial_down))if (x['std'] != 0 and x['SATUAN'] not in ['PCS']) else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )

    color4 = df_trial.apply(
        lambda x: ((((x[angka_cols] - x['mean']) / x['std']) > ts_trial_up)) if (x['std'] != 0)  else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )

    color2 = df_trial.apply(
        lambda x: x[angka_cols] < 0,
        axis=1, result_type='expand'
    )

    color3 = df_trial.assign(CABANG=df_trial['CABANG'].str[5:]).merge(df_bom, how='left', on=['CABANG','NAMA BARANG']).fillna(0).apply(
        lambda x: np.abs(x[angka_cols]) > 0 if x['TOTAL'] == 0 else [False]*len(angka_cols),
        axis=1, result_type='expand'
    )
    color3.columns = angka_cols

    color = color1 | color2 | color3 | color4
    color.columns = angka_cols

    for col in angka_cols:
        df_trial[f"{col}_outlier"] = color[col]

    df_trial['is_outlier'] = df_trial[[f"{col}_outlier" for col in angka_cols]].any(axis=1)
    def export_with_color(df, angka_cols, filename='output.xlsx'):
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            df[['CABANG','NAMA BARANG','KATEGORI','SATUAN','TOTAL ']+angka_cols.to_list()+['is_outlier']].to_excel(writer, index=False, sheet_name='Sheet1')
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']

            red_format = workbook.add_format({'bg_color': '#FF0000', 'font_color': 'white'})

            for row_idx, row in df.iterrows():
                for col_idx, col in enumerate(angka_cols):
                    if row[f"{col}_outlier"]:
                        worksheet.write(row_idx+1, df.columns.get_loc(col), row[col], red_format)

    outlier_trial = st.radio(
        "Hanya Tampilkan Outlier:",
        ("Iya","Tidak"), key='outlier_trial')

    if outlier_trial == "Iya": 
        df_trial = df_trial[df_trial['is_outlier']]

    gb = GridOptionsBuilder.from_dataframe(df_trial)
    for col in angka_cols:
        cell_style_jscode = JsCode(f"""
        function(params) {{
            if (params.data['{col}_outlier']) {{
                return {{'backgroundColor': 'red', 'color': 'white'}};
            }}
        }}
        """)
        gb.configure_column(col, cellStyle=cell_style_jscode)
        
        value_formatter_jscode = JsCode("""
        function(params) {
            if (params.value == null || isNaN(params.value)) return '';
            return (params.value).toLocaleString();
        }
        """)
        gb.configure_column(col, valueFormatter=value_formatter_jscode)

    for col in angka_cols.to_list():
        gb.configure_column(f"{col}_outlier", hide=True, suppressColumnsToolPanel=True)
    for col in ['mean','std','subtotal','is_outlier']:
        gb.configure_column(col, hide=True, suppressColumnsToolPanel=True)
    for col in ['CABANG','NAMA BARANG','KATEGORI','SATUAN']:
        gb.configure_column(col, pinned='left')
    gb.configure_default_column(filter=True, sortable=True,
        flex=1,
        minWidth=100,
        enableValue=True,
        enableRowGroup=True,
        enablePivot=True
    )

    gridOptions = gb.build()
    gridOptions['autoGroupColumnDef'] = {
        'minWidth': 200,
        'pinned': 'left'
    }
    gridOptions['sideBar'] = {
        "toolPanels": ['columns', 'filters']  
    }

    gridOptions['pivotMode'] = False
    gridOptions['pivotPanelShow'] = "always"
    if "columnDefs" in gridOptions:
        for colDef in gridOptions["columnDefs"]:
            colDef["autoWidth"] = True

    AgGrid(
        df_trial,
        gridOptions, update_mode=GridUpdateMode.NO_UPDATE,
        enable_enterprise_modules=enable_enterprise,
        allow_unsafe_jscode=True,custom_css=custom_css)
    if st.button('Export', key='export-trial'):
        export_with_color(df_trial, angka_cols,'Data/COM Monitoring/Output/REKAP_TRIAL.xlsx')
