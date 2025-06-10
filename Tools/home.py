import pandas as pd
import glob
import streamlit as st
from io import BytesIO
import io
from PIL import Image
import os

if 'Flowbit.exe' not in os.listdir():
    dir_main = 'Main/'
else:
    dir_main = ''

# Membaca gambar menggunakan PIL
image = Image.open(dir_main + "etc/Pict Home.png")  # Ganti dengan path file gambarmu

col = st.columns([4,1])
with col[0]:
    st.image(image)
with col[1]:
    with st.expander("v.2.0.1 - Updated Details:"):

        success_html = """
        <div style="
            background-color: #d4edda; 
            color: #155724; 
            border-radius: 5px; 
            font-size: 11px;
            font-weight: 600;
            border: 1px solid #c3e6cb;
        ">
        <ul style="padding-left: 10px; padding-right: 10px; margin: 0;">
            <li style="margin-bottom: 6px;">Perubahan tampilan web</li>
            <li style="margin-bottom: 6px;">Penambahan fitur pembaruan versi secara otomatis</li>
            <li style="margin-bottom: 6px;">Penambahan halaman "Home"</li>
            <li>Perbaikan bug pada modul "PENYESUAIAN IA" [SCM-Cleaning]</li>
        </ul>
        </div>
        """

        st.markdown(success_html, unsafe_allow_html=True)


        st.write(' ')
