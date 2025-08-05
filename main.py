from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import openpyxl
import shutil
import os

app = FastAPI()

# Enable CORS for frontend communication
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Save uploaded consensus file
    consensus_path = "consensus.xlsx"
    with open(consensus_path, "wb") as f:
        shutil.copyfileobj(consensus.file, f)

    # Save profile file if uploaded
    profile_path = None
    if profile:
        profile_path = "profile.xlsx"
        with open(profile_path, "wb") as f:
            shutil.copyfileobj(profile.file, f)

    # Generate final Excel file
    output_path = "DCF_Model_Output.xlsx"
    build_final_excel(consensus_path, profile_path, output_path)

    return FileResponse(output_path, filename="DCF_Model.xlsx")

def build_final_excel(consensus_path, profile_path, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
        # Load and write all sheets from consensus file
        consensus_wb = pd.read_excel(consensus_path, sheet_name=None, engine='openpyxl')
        for sheet, df in consensus_wb.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

        # Load and write the DCF Model sheet from Template.xlsx
        template = pd.ExcelFile("Template.xlsx", engine='openpyxl')
        dcf_model_df = template.parse("DCF Model")
        dcf_model_df.to_excel(writer, sheet_name="DCF Model", index=False)

        # Load and write Public Company sheet if profile file is present
        if profile_path and os.path.exists(profile_path):
            profile_df = pd.read_excel(profile_path, sheet_name="Public Company", engine='openpyxl')
            profile_df.to_excel(writer, sheet_name="Public Company", index=False)
