from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os
from tempfile import NamedTemporaryFile

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name):
    """Copy a sheet with all formatting, formulas, and dimensions"""
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy column dimensions
    for col in range(1, source_sheet.max_column + 1):
        col_letter = get_column_letter(col)
        new_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
    
    # Copy row dimensions
    for row in range(1, source_sheet.max_row + 1):
        new_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height
    
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

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = "temp_consensus.xlsx"
    profile_path = None
    temp_file_path = None
    
    try:
        # Save uploaded files
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)
            
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Load macro-enabled template as output workbook
        output_wb = load_workbook("Template.xlsm", data_only=False, keep_vba=True)
        
        # Remove existing sheets except "DCF Model"
        for sheet_name in list(output_wb.sheetnames):
            if sheet_name != "DCF Model":
                output_wb.remove(output_wb[sheet_name])
        
        # Load user's consensus file
        consensus_wb = load_workbook(consensus_path)
        
        # Copy all sheets from consensus file (except "DCF Model")
        for sheet_name in consensus_wb.sheetnames:
            if sheet_name == "DCF Model":
                continue
            source_sheet = consensus_wb[sheet_name]
            copy_sheet(source_sheet, output_wb, sheet_name)
        
        # Handle profile file if provided
        if profile_path:
            profile_wb = load_workbook(profile_path)
            for sheet_name in profile_wb.sheetnames:
                if sheet_name == "DCF Model":
                    continue
                source_sheet = profile_wb[sheet_name]
                copy_sheet(source_sheet, output_wb, sheet_name)
        
        # Save as macro-enabled workbook
        with NamedTemporaryFile(delete=False, suffix=".xlsm") as temp_file:
            output_wb.save(temp_file.name)
            temp_file_path = temp_file.name
        
        return StreamingResponse(
            open(temp_file_path, "rb"),
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsm"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        for path in [consensus_path, profile_path, temp_file_path]:
            if path and os.path.exists(path):
                os.remove(path)
