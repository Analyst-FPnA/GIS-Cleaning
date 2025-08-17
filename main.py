import streamlit as st
import requests
import re
import io
import os
import zipfile
import tempfile


with open("version.py", "r", encoding="utf-8") as f:
    file_content = f.read()

namespace = {}
exec(file_content, namespace)

version = namespace.get("version")
data= namespace.get("data")

st.set_page_config(
    page_title="DEX",
    page_icon="ikon.ico",
    layout="wide")

st.markdown("""
        <style>
               .block-container {
                    padding-top: 3rem;
                }
        </style>
        """, unsafe_allow_html=True)

st.markdown(
    """
    <style>
    div.stPopover > div>button:hover {
        background-color: #ff4b4b !important;  /* Merah saat hover */
        color: white !important;
        border: 1px solid white !important;
    }
    div.stButton > button:hover {
        background-color: #ff4b4b !important;  /* Merah saat hover */
        color: white !important;
        border: 1px solid white !important;
    }
    button[disabled], div[disabled], .stButton button:disabled {
        opacity: 1.0 !important;
        background-color: #ffffff !important;
        color: #555 !important;
        cursor: not-allowed !important;
    }
    button[data-testid="stPopoverButton"] svg {
        margin-left: auto;
        transform: translateX(4px);
    }

    div[data-testid="stSidebarCollapseButton"] svg {
        fill: white !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #001C53;
        width: 200px;
    }
    .st-key-left .stButton button {
        text-align: left;
        justify-content: flex-start;
    }
    div[data-testid="stPopover"]>div>button{
        background-color: white !important;  /* Merah saat hover */
        display: flex;
        justify-content: space-between;
        align-items: center;
        width: 100%;
    }
    </style>
    """,
    unsafe_allow_html=True)

users = {
    "admin": {"password": "admin123", "access": ["AR", "TAX"]},
    "ar_user": {"password": "aronly", "access": ["AR"]},
    "tax_user": {"password": "taxonly", "access": ["TAX"]}
}

temp_dir = tempfile.gettempdir()
TOOLS_DIR = os.path.join(temp_dir, "Tools")

page_1 = st.Page(os.path.join(TOOLS_DIR, "home.py"), title="Home")
page_2 = st.Page(os.path.join(TOOLS_DIR, "gis.py"), title="GIS-Cleaning")
page_3 = st.Page(os.path.join(TOOLS_DIR, "AR/Analytics", "1311 Checking.py"), title="1311 Checking")
page_4 = st.Page(os.path.join(TOOLS_DIR, "TAX/Processing.py"), title="Processing")

current_page = st.navigation(pages=[page_1,page_2,page_3,page_4], position="hidden")
group_labels = {
    'AR': '💵 AR',
    'TAX': '🏛️ TAX'
}

pages_by_group = {
    'AR': {
        'Analytics': [
            {'title': '1311 Checking', 'page': os.path.join(temp_dir, 'Tools/AR/Analytics/1311 Checking.py')}
        ]
    },
    'TAX': {
        'Processing': [
            {'title': 'Processing', 'page': os.path.join(temp_dir, 'Tools/TAX/Processing.py')}
        ]
    }
}
for group, subgroups in pages_by_group.items():
    for subgroup, pages in subgroups.items():
        for page in pages:
            original_path = page["page"]  # e.g., 'Tools/TAX/Processing.py'
            full_path = os.path.join(temp_dir, original_path)  # absolute path ke hasil ekstrak
            page["page"] = full_path  # replace path
            
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

# Login form
if not st.session_state.logged_in:
    st.title("🔒 Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        user = users.get(username)
        if user and user["password"] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success("Login successful")
            st.rerun()
        else:
            st.error("Invalid username or password")
    st.stop()

user_access = users[st.session_state.username]["access"]

with st.sidebar:
    st.markdown('<h1 style="color: white; font-weight: bold;margin:0; padding:0;">DEX 🚀</h1>',unsafe_allow_html=True)
    st.markdown(f'<div style="font-size:12px ;color: white; font-weight: bold; margin:0; padding:0;">{version} #TAF.{data}</div>',unsafe_allow_html=True)

    st.markdown(' ')
    if st.button("🏠 Home", use_container_width=True, key='left'):
        st.switch_page("Tools/home.py")
    st.write(' ')
    for group, subgroups in pages_by_group.items():
        label = group_labels.get(group, group)  # Ambil label dengan icon
        allowed = group in user_access
        with st.popover(label, use_container_width=True, disabled=not allowed):
            for subgroup, pages in subgroups.items():
                if len(pages) == 1:
                    page = pages[0]
                    st.page_link(
                        page["page"],
                        label=page["title"],
                        use_container_width=True,
                        disabled=not allowed
                    )
                else:
                    with st.popover(f"↳ {subgroup}", use_container_width=True, disabled=not allowed):
                        for page in pages:
                            st.page_link(
                                page["page"],
                                label=page["title"],
                                use_container_width=True,
                                disabled=not allowed
                            )

    st.markdown('<hr style="border: 1px solid white;">', unsafe_allow_html=True)


    st.write(' ')
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.rerun()
        
    
current_page.run()
