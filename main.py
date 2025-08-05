from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
from tempfile import NamedTemporaryFile
from datetime import datetime

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
    try:
        # Save uploaded files temporarily
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        profile_path = None
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Generate the output file
        output_path = generate_output_file(consensus_path, profile_path)

        # Return the generated file
        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        if os.path.exists(consensus_path):
            os.remove(consensus_path)
        if profile_path and os.path.exists(profile_path):
            os.remove(profile_path)

def generate_output_file(consensus_path, profile_path):
    try:
        # Create a new workbook with just the DCF Model sheet from template
        template = load_workbook("Template.xlsx")
        
        # Create a new workbook with only the DCF Model sheet
        output_wb = load_workbook("Template.xlsx")
        
        # Remove all sheets except DCF Model
        for sheet_name in output_wb.sheetnames:
            if sheet_name != "DCF Model":
                output_wb.remove(output_wb[sheet_name])
        
        # Update valuation date in the DCF Model
        dcf_sheet = output_wb["DCF Model"]
        update_valuation_date(dcf_sheet)
        
        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        return temp_file.name
        
    except Exception as e:
        raise Exception(f"Error generating output file: {str(e)}")

def update_valuation_date(dcf_sheet):
    """Update the valuation date in the DCF model to current date"""
    for row in dcf_sheet.iter_rows():
        for cell in row:
            if cell.value == "Valuation Date":
                dcf_sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return
