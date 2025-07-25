import pandas as pd
import glob
import streamlit as st
from io import BytesIO
import io
from PIL import Image
import os

if 'DEX.exe' not in os.listdir():
    dir_main = 'Main/'
else:
    dir_main = ''

# Membaca gambar menggunakan PIL
image = Image.open(dir_main + "etc/Pict Home.png")  # Ganti dengan path file gambarmu

col = st.columns([4,1])
with col[0]:
    st.image(image)
with col[1]:
    with st.expander("Latest Version Update Details: v2.1.5"):

        success_html = """
        <div style="
            background-color: #d4edda; 
            color: #155724; 
            border-radius: 5px; 
            font-size: 11px;
            font-weight: 600;
            border: 1px solid #c3e6cb;
        ">
        <ul style="padding-top: 10px; padding-bottom: 10px; padding-left: 12px; padding-right: 10px; margin: 0;">
            <li>Improvement of the REKAP SALES ESB & GIS module [SCM-Processing]</li>
        </ul>
        </div>
        """

        st.markdown(success_html, unsafe_allow_html=True)


        st.write(' ')
