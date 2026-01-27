import streamlit as st
from streamlit_folium import st_folium
import folium
import requests
import pandas as pd
import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# --- Configuration ---
# API endpoints
GEOCODING_API_ENDPOINT = "https://maps.googleapis.com/maps/api/geocode/json"
ELEVATION_API_ENDPOINT = "https://maps.googleapis.com/maps/api/elevation/json"

# Default API Key (Securely loaded from secrets)
# When running locally, create .streamlit/secrets.toml
# When running on Streamlit Cloud, set this in the App Settings
try:
    DEFAULT_API_KEY = st.secrets["GOOGLE_MAPS_API_KEY"]
except FileNotFoundError:
    DEFAULT_API_KEY = "" # Fallback if no secrets file found
except KeyError:
    DEFAULT_API_KEY = "" # Fallback if key missing


# --- Helper Functions (Adapted from label_app.py) ---

def get_elevation(lat, lon, api_key):
    """Calls the Google Elevation API to get the altitude."""
    if not api_key:
        return "API Key Missing"
    params = {'locations': f'{lat},{lon}', 'key': api_key}
    try:
        response = requests.get(ELEVATION_API_ENDPOINT, params=params, timeout=10)
        data = response.json()
        if data['status'] == 'OK' and len(data['results']) > 0:
            return int(round(data['results'][0]['elevation']))
        else:
            return f"Error: {data.get('status')}"
    except Exception as e:
        return f"Error: {e}"

def get_google_address_for_label(lat, lon, api_key):
    """Calls Google Geocoding API and returns the formatted address."""
    if not api_key:
        return "API Key Missing"
    params = {'latlng': f'{lat},{lon}', 'key': api_key, 'language': 'ja'}
    try:
        response = requests.get(GEOCODING_API_ENDPOINT, params=params, timeout=10)
        data = response.json()
    except Exception as e:
        return f"Request Error: {e}"
        
    if data['status'] == 'OK' and len(data['results']) > 0:
        for result in data['results']:
            if 'plus_code' in result and result.get('types') == ['plus_code']:
                continue
            
            addr = result.get('formatted_address', '')
            if addr.startswith('日本、'):
                addr = addr.replace('日本、', '', 1)
            
            # Remove postal code
            space_pos = addr.find(' ')
            if space_pos != -1 and addr[:space_pos].replace('〒', '').replace('-', '').isdigit():
                return addr[space_pos+1:].strip()
            return addr.strip()
    return "Address Not Found"

def generate_label_text(lat, lon, date, method, collector, address, elevation):
    """Generates the text content for the label."""
    
    # Elevation
    if isinstance(elevation, (int, float)):
        elev_str = f"GPS({elevation}m)"
    else:
        elev_str = f"GPS({elevation})"

    # Address
    if not address or 'Error' in address or 'Missing' in address:
         addr_str = f"Address Error: {address}"
    else:
         addr_str = f"JAPAN: {address}"

    # Line 3 (Date, Method, Collector)
    line3_parts = []
    if date: line3_parts.append(str(date))
    if method: line3_parts.append(str(method))
    if collector: line3_parts.append(str(collector))
    
    line3_str = ". ".join(line3_parts)
    if line3_str and not line3_str.endswith('.'):
        line3_str += "."

    # Line 4 (Coords)
    coords_str = f"N{lat}, E{lon}"

    return f"{addr_str}\n{elev_str}\n{line3_str}\n{coords_str}"

def create_docx(label_text):
    """Creates a DOCX file with the label text."""
    doc = Document()
    # Set narrow margins for label printing if needed, or just standard
    section = doc.sections[0]
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    p = doc.add_paragraph()
    run = p.add_run(label_text)
    run.font.name = 'Arial'
    run.font.size = Pt(8) # Small font for labels
    
    # Save to memory buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- Main App ---

st.set_page_config(page_title="Specimen Label Generator", layout="wide")

st.title("🏷️ Specimen Label Generator")
st.markdown("Select a location on the map, fill in the details, and generate your label.")

# Sidebar for Settings
with st.sidebar:
    st.header("Settings")
    api_key = st.text_input("Google Maps API Key", value=DEFAULT_API_KEY, type="password")
    st.caption("Required for Address and Elevation lookup.")

# Main Layout: Map and Form
col1, col2 = st.columns([2, 1])

with col1:
    st.subheader("📍 Location Selection")
    # Default location (Japan center roughly)
    m = folium.Map(location=[36.2048, 138.2529], zoom_start=5)
    
    # Check if we have session state for markers, if strictly needed, 
    # but folium usually handles the click return well enough for simple cases.
    
    # Display Map
    output = st_folium(m, height=500, use_container_width=True)

    # Handle Map Click
    lat = None
    lon = None
    if output and output['last_clicked']:
        lat = output['last_clicked']['lat']
        lon = output['last_clicked']['lng']

with col2:
    st.subheader("📝 Label Details")
    
    # Coordinates Input (Auto-filled but editable)
    st.caption("Coordinates (Click map to auto-fill)")
    input_lat = st.number_input("Latitude", value=lat if lat else 0.0, format="%.6f", key="input_lat")
    input_lon = st.number_input("Longitude", value=lon if lon else 0.0, format="%.6f", key="input_lon")
    
    # Metadata
    collection_date = st.date_input("Collection Date", datetime.date.today())
    collector_name = st.text_input("Collector Name", value="M. Tsuchioka") # Example default
    options = ["Light trap", "Sweeping", "Beating", "Bait trap", "Hand picking", "Fit", "Malaise trap"]
    collection_method = st.selectbox("Collection Method", options + ["Other"])
    if collection_method == "Other":
        collection_method = st.text_input("Enter Method")

    generate_btn = st.button("Generate Label", type="primary", use_container_width=True)

# Generate Logic
if generate_btn:
    if not api_key:
        st.error("Please enter a Google Maps API Key in the sidebar.")
    elif input_lat == 0.0 and input_lon == 0.0:
        st.warning("Please select a location on the map or enter coordinates.")
    else:
        with st.spinner("Fetching data from Google Maps..."):
            # 1. Get Data
            address = get_google_address_for_label(input_lat, input_lon, api_key)
            elevation = get_elevation(input_lat, input_lon, api_key)
            
            # 2. Format Text
            label_text = generate_label_text(
                input_lat, input_lon, 
                collection_date.strftime('%Y.%m.%d'), # Format date
                collection_method, 
                collector_name, 
                address, 
                elevation
            )
            
            # 3. Display Result
            st.success("Label Generated!")
            st.text_area("Label Preview", value=label_text, height=150)
            
            # 4. Download Option
            docx_file = create_docx(label_text)
            st.download_button(
                label="Download Label (.docx)",
                data=docx_file,
                file_name=f"label_{datetime.date.today()}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

