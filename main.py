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

def copy_sheet_with_formatting(source_sheet, target_wb, sheet_name):
    """Copy a sheet with all formatting, formulas, dimensions, and merged cells"""
    # Create new sheet in target workbook
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy column dimensions
    for col_idx, col_dim in source_sheet.column_dimensions.items():
        new_sheet.column_dimensions[col_idx].width = col_dim.width
    
    # Copy row dimensions
    for row_idx, row_dim in source_sheet.row_dimensions.items():
        new_sheet.row_dimensions[row_idx].height = row_dim.height
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy cells with values, formulas, and formatting
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(
                row=cell.row,
                column=cell.column,
                value=cell.value
            )
            
            # Preserve formulas
            if cell.data_type == 'f':
                new_cell.value = cell.value
            
            # Copy all styling attributes
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.border = cell.border.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection.copy()
                new_cell.alignment = cell.alignment.copy()
    
    return new_sheet

def update_valuation_date(sheet):
    """Update valuation date to current date"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and "Valuation Date" in str(cell.value):
                date_cell = sheet.cell(row=cell.row, column=cell.column + 2)
                date_cell.value = datetime.now().date()
                return

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = "temp_consensus.xlsx"
    profile_path = None
    
    try:
        # Save uploaded consensus file
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)
            
        # Save profile file if provided
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Load user's consensus file - this will be our base output
        output_wb = load_workbook(consensus_path)
        
        # Load template to get perfect DCF Model sheet
        template_wb = load_workbook("Template.xlsx", data_only=False)
        dcf_sheet = template_wb["DCF Model"]
        
        # Copy the perfect DCF Model sheet to output workbook
        new_dcf_sheet = copy_sheet_with_formatting(dcf_sheet, output_wb, "DCF Model")
        
        # Update valuation date
        update_valuation_date(new_dcf_sheet)
        
        # Save combined workbook
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        
        return StreamingResponse(
            open(temp_file.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        for path in [consensus_path, profile_path]:
            if path and os.path.exists(path):
                os.remove(path)
