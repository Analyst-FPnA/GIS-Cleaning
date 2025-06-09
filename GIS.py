import streamlit as st
import requests

def run_stream_script(url):
    # Mengunduh file dari GitHub
    response = requests.get(url)
    if response.status_code == 200:
        # Menjalankan file yang diunduh
        exec(response.text, globals())
    else:
        st.error(f"Failed to download file: {response.status_code}")
        
stream1_url = 'https://raw.githubusercontent.com/Analyst-FPnA/GIS-Cleaning/main/main.py'
run_stream_script(stream1_url)
