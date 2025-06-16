# Campaign+Asset Multi-Tab Output (Hero-Evolving Style)

A Streamlit app for generating campaign/asset-wise product exports from Google Sheets or Excel, with image download and background removal options.

## Features

- **Input:** Accepts a public Google Sheet link or Excel upload (hubwise tabs).
- **Product CSV:** Upload a product CSV (`MB_id`, `image_src`) to auto-fill image links for all PIDs.
- **Output:**  
  - Each unique (Campaign Name + Asset) gets its own Excel tab.
  - Each row: Hub, Focus Grid, PID1, PID2, Img1, Img2.
  - Final `All_PIDs` tab lists all unique PIDs and their corresponding image links.
- **Image Tools:**  
  - Download all unique product images as a ZIP.
  - Optional background removal (rembg) for all images.
- **Filters:**  
  - Ignores rows with both PIDs blank.
  - Ignores Asset = 'ATC' or 'ATC background'.
  - Ensures all PIDs are output as strings (no decimals).
- **Preview:**  
  - Preview any output tab and the All_PIDs list before download.

## Usage

1. Clone or download this repository.
2. Install requirements:  
