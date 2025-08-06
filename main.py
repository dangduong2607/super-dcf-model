from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
from tempfile import NamedTemporaryFile
from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name):
    """Copy a sheet from source workbook to target workbook"""
    if sheet_name not in source_sheet.sheetnames:
        return False
    
    # Create new sheet in target workbook
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Get source sheet
    source = source_sheet[sheet_name]
    
    # Copy cells
    for row in source.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = cell.font
                new_cell.border = cell.border
                new_cell.fill = cell.fill
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection
                new_cell.alignment = cell.alignment
    
    # Copy merged cells
    for merged_range in source.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy column dimensions
    for col_idx, col_dim in source.column_dimensions.items():
        new_sheet.column_dimensions[col_idx].width = col_dim.width
        new_sheet.column_dimensions[col_idx].hidden = col_dim.hidden
    
    # Copy row dimensions
    for row_idx, row_dim in source.row_dimensions.items():
        new_sheet.row_dimensions[row_idx].height = row_dim.height
        new_sheet.row_dimensions[row_idx].hidden = row_dim.hidden
    
    return True

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = None
    profile_path = None
    
    try:
        # Save consensus file temporarily
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as consensus_temp:
            shutil.copyfileobj(consensus.file, consensus_temp)
            consensus_path = consensus_temp.name

        # Save profile file temporarily if provided
        if profile:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as profile_temp:
                shutil.copyfileobj(profile.file, profile_temp)
                profile_path = profile_temp.name

        # Create output workbook
        output = BytesIO()
        
        # Load DCF template
        template = load_workbook("Template.xlsx")
        
        # Remove all sheets except DCF Model
        for sheet_name in template.sheetnames[:]:
            if sheet_name != "DCF Model":
                template.remove(template[sheet_name])
        
        # Copy user-provided sheets
        consensus_wb = load_workbook(consensus_path)
        copy_sheet(consensus_wb, template, "Consensus")
        
        if profile_path:
            profile_wb = load_workbook(profile_path)
            copy_sheet(profile_wb, template, "Public Company")
        
        # Save final workbook
        template.save(output)
        output.seek(0)
        
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing files: {str(e)}")
    
    finally:
        # Clean up temporary files
        if consensus_path and os.path.exists(consensus_path):
            os.remove(consensus_path)
        if profile_path and os.path.exists(profile_path):
            os.remove(profile_path)
