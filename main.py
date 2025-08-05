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
        # Save uploaded files
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

        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model_Output.xlsx"},
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
        # Load the consensus file
        output_wb = load_workbook(consensus_path)
        
        # Add Public Company sheet if provided
        if profile_path:
            profile_wb = load_workbook(profile_path)
            if "Public Company" in profile_wb.sheetnames:
                profile_sheet = profile_wb["Public Company"]
                output_wb.create_sheet("Public Company")
                new_sheet = output_wb["Public Company"]
                
                # Copy all content from profile sheet
                for row in profile_sheet.iter_rows():
                    for cell in row:
                        new_cell = new_sheet.cell(
                            row=cell.row, 
                            column=cell.column,
                            value=cell.value
                        )
                        if cell.has_style:
                            new_cell.font = cell.font.copy()
                            new_cell.border = cell.border.copy()
                            new_cell.fill = cell.fill.copy()
                            new_cell.number_format = cell.number_format
                            new_cell.protection = cell.protection.copy()
                            new_cell.alignment = cell.alignment.copy()
        
        # Load the template and copy DCF Model sheet
        template = load_workbook("Template.xlsx")
        dcf_sheet = template["DCF Model"]
        
        # Create new DCF Model sheet in output workbook
        output_wb.create_sheet("DCF Model")
        new_dcf_sheet = output_wb["DCF Model"]
        
        # Copy all content from template's DCF Model sheet
        for row in dcf_sheet.iter_rows():
            for cell in row:
                new_cell = new_dcf_sheet.cell(
                    row=cell.row, 
                    column=cell.column,
                    value=cell.value
                )
                if cell.has_style:
                    new_cell.font = cell.font.copy()
                    new_cell.border = cell.border.copy()
                    new_cell.fill = cell.fill.copy()
                    new_cell.number_format = cell.number_format
                    new_cell.protection = cell.protection.copy()
                    new_cell.alignment = cell.alignment.copy()
        
        # Update valuation date
        update_valuation_date(new_dcf_sheet)
        
        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        return temp_file.name
        
    except Exception as e:
        raise Exception(f"Error generating output file: {str(e)}")

def update_valuation_date(sheet):
    """Update the valuation date in the DCF model to current date"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Valuation Date":
                sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return
