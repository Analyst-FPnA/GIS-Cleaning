import streamlit as st
import requests
import re
import io
import os
import zipfile

version = 'v2.0.0'
st.set_page_config(layout="wide")

page_1 = st.Page("Tools/gis-cleaning.py", title="GIS-CLeaning")
page_2 = st.Page("Tools/scm-cleaning.py", title="SCM-Cleaning")
page_3 = st.Page("Tools/home.py", title="Home")

current_page = st.navigation(pages=[page_1,page_2,page_3], position="hidden")
pages_by_group = {
                  '🧰 Tools':[
                    {'title':'GIS-Cleaning','page':'Tools/gis-cleaning.py'}, 
                    {'title':'SCM-Cleaning','page':'Tools/scm-cleaning.py'}]
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
    st.markdown('<h2 style="color: white; font-weight: bold;margin:0; padding:0;">FPNA-Analyst</h2>',unsafe_allow_html=True)
    st.markdown(f'<div style="font-size:12px ;color: white; font-weight: bold; margin:0; padding:0;">{version}</div>',unsafe_allow_html=True)

    st.write(' ')
    if st.button("🏠 Home", use_container_width=True, key='left'):
        st.switch_page("Tools/home.py")
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
        if remote_version == version:
            st.markdown('<div style="font-size:12px ;color: white; ">New Version not Available</div>',unsafe_allow_html=True)
            if st.button('Update'):
                zip_url = f"https://github.com/Analyst-FPNA/GIS-Cleaning/archive/refs/heads/main.zip"

                response = requests.get(zip_url)
                if response.status_code != 200:
                    raise Exception(f"Gagal mengunduh ZIP: {response.status_code}")
                with zipfile.ZipFile(io.BytesIO(response.content)) as z:
                    root_folder = z.namelist()[0].split('/')[0]  # Misalnya: "repo-branch"

                    for member in z.namelist():
                        if member.endswith("/"):
                            continue  # Lewati folder

                        # Hapus nama folder root
                        relative_path = os.path.relpath(member, root_folder)
                        destination_path = os.path.join('..', relative_path)

                        # Buat direktori tujuan jika belum ada
                        os.makedirs(os.path.dirname(destination_path), exist_ok=True)

                        # Simpan file
                        with open(destination_path, "wb") as f:
                            f.write(z.read(member))
                
        else:
            st.markdown('<div style="font-size:12px ;color: white; ">New Version Available</div>',unsafe_allow_html=True)
            
    except (requests.ConnectionError, requests.Timeout):
        st.markdown('<div style="font-size:12px ;color: white; ">Offline</div>',unsafe_allow_html=True)


current_page.run()
