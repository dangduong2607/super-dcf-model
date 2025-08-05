from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import openpyxl
import shutil
import os

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
    # Save uploaded files
    consensus_path = "consensus.xlsx"
    with open(consensus_path, "wb") as f:
        shutil.copyfileobj(consensus.file, f)

    profile_path = None
    if profile:
        profile_path = "profile.xlsx"
        with open(profile_path, "wb") as f:
            shutil.copyfileobj(profile.file, f)

    # Generate the combined Excel file
    output_path = "DCF_Model_Output.xlsx"
    build_final_excel(consensus_path, profile_path, output_path)

    return StreamingResponse(
        open(output_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
    )

def build_final_excel(consensus_path, profile_path, output_path):
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
        # Load all sheets from consensus
        consensus_wb = pd.read_excel(consensus_path, sheet_name=None, engine="openpyxl")
        for sheet, df in consensus_wb.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

        # Add DCF Model sheet from template
        template = pd.ExcelFile("Template.xlsx", engine="openpyxl")
        dcf_model_df = template.parse("DCF Model")
        dcf_model_df.to_excel(writer, sheet_name="DCF Model", index=False)

        # Try to load Public Company sheet from profile, if uploaded
        if profile_path and os.path.exists(profile_path):
            try:
                profile_df = pd.read_excel(profile_path, sheet_name="Public Company", engine="openpyxl")
                profile_df.to_excel(writer, sheet_name="Public Company", index=False)
            except Exception as e:
                print(f"Skipping profile sheet due to error: {e}")
