import streamlit as st
import os, glob, re
from pdf2image import convert_from_path
import pytesseract
import cv2
import numpy as np
import pandas as pd
from openpyxl import load_workbook

st.set_page_config(page_title="FBA Tracking Uploader", layout="wide")
st.title("📦 Amazon FBA Box Tracking Uploader")

# Sidebar: upload only PDFs
st.sidebar.header("Upload Files")
fba_pdf  = st.sidebar.file_uploader("Upload FBA Box Labels PDF", type=["pdf"])
ship_pdf = st.sidebar.file_uploader("Upload Shipping Labels PDF", type=["pdf"])

# Hardcoded template path (must exist in app directory)
TEMPLATE_PATH = "Amazon_Tracking_Uploads_Template.xlsx"

if st.sidebar.button("Run Extraction"):
    if not (fba_pdf and ship_pdf):
        st.error("Please upload both FBA and Shipping PDFs.")
    else:
        # Save uploaded PDFs
        tmp_dir = "./tmp"
        os.makedirs(tmp_dir, exist_ok=True)
        fba_path  = os.path.join(tmp_dir, fba_pdf.name)
        ship_path = os.path.join(tmp_dir, ship_pdf.name)
        with open(fba_path, "wb") as f: f.write(fba_pdf.read())
        with open(ship_path, "wb") as f: f.write(ship_pdf.read())

        # Derive shipment_id
        base = os.path.splitext(fba_pdf.name)[0]
        shipment_id = base.split("-")[0]
        output_xl   = f"{shipment_id}_tracking_upload.xlsx"

        # OCR-based “barcode” decoder for both FBA and UPS
        def decode_barcode(img):
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, bw = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            txt = pytesseract.image_to_string(
                bw,
                config="--psm 7 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            ).upper()
            m = re.search(r'(?:FBA[0-9A-Z]+|1Z[0-9A-Z]{8,})', txt)
            if not m:
                return ""
            raw = re.sub(r'[^0-9A-Z]', '', m.group(0))
            # Format UPS codes in 4-char groups
            if raw.startswith("1Z"):
                return " ".join(raw[i:i+4] for i in range(0, len(raw), 4))
            return raw

        # Extract box IDs and tracking numbers
        def extract_box_and_tracking(pdf_path):
            pages = convert_from_path(pdf_path, dpi=300)
            recs = []
            for i, pil_img in enumerate(pages, start=1):
                img = np.array(pil_img.convert("RGB"))
                h, w = img.shape[:2]
                # FBA crop (upper-middle)
                y1, y2 = int(h*0.2), int(h*0.45)
                fba_crop = img[y1:y2, int(w*0.05):int(w*0.95)]
                box_id = decode_barcode(fba_crop)
                # UPS crop (middle-lower)
                y3, y4 = int(h*0.5), int(h*0.8)
                ups_crop = img[y3:y4, int(w*0.05):int(w*0.95)]
                tracking = decode_barcode(ups_crop)
                recs.append({"page": i, "box_id": box_id, "tracking": tracking})
            return pd.DataFrame(recs)

        boxes_df  = extract_box_and_tracking(fba_path)[["page","box_id"]]
        tracks_df = extract_box_and_tracking(ship_path)[["page","tracking"]]
        merged    = pd.merge(boxes_df, tracks_df, on="page")

        # Fill Excel template
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        ws.cell(row=4, column=2, value=shipment_id)  # Shipment ID in B4

        start_row = 8
        for idx, row in merged.iterrows():
            r = start_row + idx
            clean_box   = row["box_id"].replace(" ", "")
            clean_track = row["tracking"].replace(" ", "")
            ws.cell(row=r, column=1, value=clean_box)
            ws.cell(row=r, column=2, value=clean_track)

        wb.save(output_xl)
        st.success(f"Generated file: {output_xl}")
        with open(output_xl, "rb") as f:
            st.download_button("Download Tracking Excel", data=f, file_name=output_xl)
