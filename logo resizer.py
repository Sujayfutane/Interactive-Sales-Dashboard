import streamlit as st
from PIL import Image

# -----------------------------
# LOGO CONFIGURATION
# -----------------------------
logo_path = r"G:/Data Science Intership/Interactive Sales Dashboard/logo.jpg"  # use raw string or forward slashes
max_width = 250  # max width for display

# Load the logo
logo = Image.open(logo_path)

# Resize while keeping aspect ratio
width_percent = max_width / float(logo.size[0])
new_height = int(float(logo.size[1]) * width_percent)
logo_resized = logo.resize((max_width, new_height))

# Display logo in sidebar
st.sidebar.image(logo_resized, width=max_width)

# Optional: Display at top of main page
st.image(logo_resized, width=max_width)
