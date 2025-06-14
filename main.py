import streamlit as st
import requests
import re
import io
import os
import zipfile

version = 'v2.0.1'
data = '10/06/2025'
st.set_page_config(layout="wide")
st.markdown("""
        <style>
               .block-container {
                    padding-top: 3rem;
                }
        </style>
        """, unsafe_allow_html=True)
if 'DEX.exe' not in os.listdir():
    dir_main = 'Main/'
    status = 'online'
    if 'Main' not in os.listdir():
        zip_url = "https://github.com/Analyst-FPNA/GIS-Cleaning/archive/refs/heads/main.zip"
        
        response = requests.get(zip_url)
        if response.status_code != 200:
            raise Exception(f"Gagal mengunduh ZIP: {response.status_code}")
        
        with zipfile.ZipFile(io.BytesIO(response.content)) as z:
            root_folder = z.namelist()[0].split('/')[0]  # Contoh: 'GIS-Cleaning-main'
            
            # Folder target tempat ekstraksi
            extract_root = os.path.join("Main")
        
            for member in z.namelist():
                if member.endswith("/"):
                    continue  # Lewati folder
        
                # Hapus nama folder root dari path
                relative_path = os.path.relpath(member, root_folder)
        
                # Path final tempat file disimpan
                target_path = os.path.join(extract_root, relative_path)
        
                # Buat folder jika belum ada
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
        
                # Tulis file
                with open(target_path, "wb") as f:
                    f.write(z.read(member))
else:
    dir_main = ''
    status = 'offline'

page_1 = st.Page(dir_main + "Tools/gis.py", title="GIS-Processing")
page_2 = st.Page(dir_main + "Tools/scm.py", title="SCM-Processing")
page_3 = st.Page(dir_main + "Tools/home.py", title="Home")

current_page = st.navigation(pages=[page_1,page_2,page_3], position="hidden")
pages_by_group = {
                  '🧰 Tools':[
                    {'title':'GIS-Processing','page': dir_main + 'Tools/gis.py'}, 
                    {'title':'SCM-Processing','page': dir_main + 'Tools/scm.py'}]
                   }
st.markdown(
    """
    <style>
    div[data-testid="stMarkdownContainer"] hr {
        border-color: white;
    }
    div[data-testid="stSidebarCollapseButton"] svg {
        fill: white !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #001C53;
        width:200;
    }
    
    .st-key-left .stButton button {
        text-align: left;
        justify-content: flex-start;
    }
    div[data-testid="stPopover"]>div>button{
        text-align: left;
        justify-content: flex-start;
    }
    </style>
    """,
    unsafe_allow_html=True,
  )


with st.sidebar:
    st.markdown('<h1 style="color: white; font-weight: bold;margin:0; padding:0;">DEX 🚀</h1>',unsafe_allow_html=True)
    st.markdown(f'<div style="font-size:12px ;color: white; font-weight: bold; margin:0; padding:0;">{version}</div>',unsafe_allow_html=True)

    st.markdown(' ')
    if st.button("🏠 Home", use_container_width=True, key='left'):
        st.switch_page(dir_main + "Tools/home.py")
    for group, pages in pages_by_group.items():
        with st.popover(group,use_container_width=True):
            for page in pages:
                st.page_link(
                    page["page"],
                    label=page["title"],
                    use_container_width=True
                )
    st.divider()
    try:
        # Kirim permintaan ke Google
        requests.get("https://www.google.com", timeout=3)
        url = "https://raw.githubusercontent.com/Analyst-FPnA/GIS-Cleaning/main/version.py"

        # Ambil isi file sebagai teks
        response = requests.get(url)
        file_content = response.text

        # Cari nilai variabel version (asumsikan ditulis seperti: version = "1.2.3")
        match = re.search(r'^version\s*=\s*[\'"]([^\'"]+)[\'"]', file_content, re.MULTILINE)
        remote_version = match.group(1)
        match = re.search(r'^data\s*=\s*[\'"]([^\'"]+)[\'"]', file_content, re.MULTILINE)
        data_version = match.group(1)
        if (remote_version == version) & (data_version==data):
            st.markdown('<div style="font-size:12px ;color: white; ">You are using the latest database and version</div>',unsafe_allow_html=True)
        else:
            st.markdown('<div style="font-size:12px ;color: white; ">A latest database or version is available. Please update to get the latest features</div>',unsafe_allow_html=True)
            if st.button('Update'):
                zip_url = f"https://github.com/Analyst-FPNA/GIS-Cleaning/archive/refs/heads/main.zip"

                response = requests.get(zip_url)
                if response.status_code != 200:
                    raise Exception(f"Gagal mengunduh ZIP: {response.status_code}")
                with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                    root_folder = z.namelist()[0].split('/')[0]  # contoh: 'repo-main'

                    for member in z.namelist():
                        if member.endswith("/"):
                            continue  # Lewati folder

                        # Hapus nama folder root dari path
                        relative_path = os.path.relpath(member, root_folder)

                        # Buat folder jika belum ada
                        if os.path.dirname(relative_path):
                            os.makedirs(os.path.dirname(relative_path), exist_ok=True)

                        # Simpan file ke direktori kerja
                        with open(relative_path, "wb") as f:
                            f.write(z.read(member))
            
    except (requests.ConnectionError, requests.Timeout):
        st.markdown('<div style="font-size:12px ;color: white; ">An internet connection is required to check for the latest database and version</div>',unsafe_allow_html=True)


current_page.run()
