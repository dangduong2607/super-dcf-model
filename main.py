from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
import logging

# Enable logging
logging.basicConfig(level=logging.INFO)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Save consensus file
    consensus_path = "consensus.xlsx"
    with open(consensus_path, "wb") as f:
        shutil.copyfileobj(consensus.file, f)

    # Save optional profile file
    profile_path = None
    if profile:
        profile_path = "profile.xlsx"
        with open(profile_path, "wb") as f:
            shutil.copyfileobj(profile.file, f)

    # Build final output workbook
    output_path = "DCF_Model_Output.xlsx"
    build_final_excel(consensus_path, profile_path, output_path)

    # Return response as a download
    return StreamingResponse(
        open(output_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
    )

def build_final_excel(consensus_path, profile_path, output_path):
    # Step 1: Copy user's consensus as base file
    shutil.copy(consensus_path, output_path)

    # Step 2: Open base workbook for editing
    wb = load_workbook(output_path)

    # Step 3: Append DCF Model from Template.xlsx
    template_wb = load_workbook("Template.xlsx")
    if "DCF Model" in template_wb.sheetnames:
        dcf_sheet = template_wb["DCF Model"]
        copied_dcf = wb.copy_worksheet(dcf_sheet)
        copied_dcf.title = "DCF Model"

    # Step 4: Append Public Company if available
    if profile_path and os.path.exists(profile_path):
        try:
            profile_wb = load_workbook(profile_path)
            if "Public Company" in profile_wb.sheetnames:
                pub_sheet = profile_wb["Public Company"]
                copied_pub = wb.copy_worksheet(pub_sheet)
                copied_pub.title = "Public Company"
        except Exception as e:
            logging.warning(f"Skipping profile sheet due to error: {e}")

    # Step 5: Save updated workbook
    wb.save(output_path)
