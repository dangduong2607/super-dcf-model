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
        # Create temporary files
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_consensus:
            shutil.copyfileobj(consensus.file, temp_consensus)
            consensus_path = temp_consensus.name

        # Load the consensus file
        consensus_wb = load_workbook(consensus_path)
        
        # If profile file is provided, add it to the workbook
        if profile:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_profile:
                shutil.copyfileobj(profile.file, temp_profile)
                profile_path = temp_profile.name
            
            profile_wb = load_workbook(profile_path)
            if "Public Company" in profile_wb.sheetnames:
                profile_sheet = profile_wb["Public Company"]
                if "Public Company" in consensus_wb.sheetnames:
                    consensus_wb.remove(consensus_wb["Public Company"])
                consensus_wb.create_sheet("Public Company")
                new_profile_sheet = consensus_wb["Public Company"]
                
                # Copy all content from profile sheet
                for row in profile_sheet.iter_rows():
                    for cell in row:
                        new_cell = new_profile_sheet.cell(
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
            
            os.unlink(profile_path)

        # Load the template
        template = load_workbook("Template.xlsx")
        
        # Remove existing DCF Model sheet if it exists
        if "DCF Model" in consensus_wb.sheetnames:
            consensus_wb.remove(consensus_wb["DCF Model"])
        
        # Copy DCF Model sheet from template to consensus workbook
        dcf_sheet = template["DCF Model"]
        new_sheet = consensus_wb.create_sheet("DCF Model")
        
        # Copy all cells including values, styles, and formulas
        for row in dcf_sheet.iter_rows():
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
        
        # Update valuation date
        update_valuation_date(new_sheet)
        
        # Save to temporary file
        temp_output = NamedTemporaryFile(delete=False, suffix=".xlsx")
        consensus_wb.save(temp_output.name)
        
        return StreamingResponse(
            open(temp_output.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model_Output.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        if os.path.exists(consensus_path):
            os.unlink(consensus_path)
        if 'temp_output' in locals() and os.path.exists(temp_output.name):
            os.unlink(temp_output.name)

def update_valuation_date(sheet):
    """Update the valuation date in the DCF model to current date"""
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and "Valuation Date" in cell.value:
                sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return
