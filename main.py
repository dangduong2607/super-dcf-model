from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
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

def copy_range(source_sheet, target_sheet, range_str="A1:T500"):
    """
    Copies cells from source_sheet to target_sheet within the specified range.
    Handles values, formulas, styles, and formatting.
    """
    # Parse range (simplified for A1:T500 case)
    start_col, start_row = 1, 1  # A1
    end_col, end_row = 20, 500   # T500

    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            # Get source cell
            source_cell = source_sheet.cell(row=row, column=col)
            
            # Get target cell and copy everything
            target_cell = target_sheet.cell(row=row, column=col, value=source_cell.value)
            
            # Copy style if exists
            if source_cell.has_style:
                target_cell.font = source_cell.font.copy()
                target_cell.border = source_cell.border.copy()
                target_cell.fill = source_cell.fill.copy()
                target_cell.number_format = source_cell.number_format
                target_cell.protection = source_cell.protection.copy()
                target_cell.alignment = source_cell.alignment.copy()

            # Copy hyperlinks if exists
            if source_cell.hyperlink:
                target_cell.hyperlink = source_cell.hyperlink
                target_cell.style = "Hyperlink"

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded consensus file
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_consensus:
            shutil.copyfileobj(consensus.file, temp_consensus)
            consensus_path = temp_consensus.name

        # Load the consensus file
        consensus_wb = load_workbook(consensus_path)
        
        # If profile file is provided, add its Public Company data
        if profile:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_profile:
                shutil.copyfileobj(profile.file, temp_profile)
                profile_path = temp_profile.name
            
            profile_wb = load_workbook(profile_path)
            if "Public Company" in profile_wb.sheetnames:
                profile_sheet = profile_wb["Public Company"]
                
                # Remove existing sheet if present
                if "Public Company" in consensus_wb.sheetnames:
                    consensus_wb.remove(consensus_wb["Public Company"])
                
                # Create new sheet and copy A1:T500
                new_profile_sheet = consensus_wb.create_sheet("Public Company")
                copy_range(profile_sheet, new_profile_sheet)
            
            os.unlink(profile_path)

        # Load the template
        template = load_workbook("Template.xlsx")
        template_dcf_sheet = template["DCF Model"]
        
        # Remove existing DCF Model sheet if it exists
        if "DCF Model" in consensus_wb.sheetnames:
            consensus_wb.remove(consensus_wb["DCF Model"])
        
        # Create new DCF Model sheet and copy A1:T500
        new_dcf_sheet = consensus_wb.create_sheet("DCF Model")
        copy_range(template_dcf_sheet, new_dcf_sheet)
        
        # Update valuation date
        update_valuation_date(new_dcf_sheet)
        
        # Save to temporary output file
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
    """Updates the valuation date in the DCF model to current date"""
    for row in sheet.iter_rows(min_row=1, max_row=500, min_col=1, max_col=20):  # A1:T500 range
        for cell in row:
            if isinstance(cell.value, str) and "Valuation Date" in cell.value:
                sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return
