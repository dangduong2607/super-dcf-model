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
            # Handle array formulas (spilled ranges) correctly
            if cell.data_type == 'f' and cell.value.startswith('{=') and cell.value.endswith('}'):
                # Remove curly braces from array formulas
                formula = cell.value[1:-1]
                new_cell = new_sheet.cell(
                    row=cell.row,
                    column=cell.column,
                    value=formula
                )
            else:
                new_cell = new_sheet.cell(
                    row=cell.row,
                    column=cell.column,
                    value=cell.value
                )
            
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
    
    try:
        # Save uploaded files
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)
            
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Load user's consensus file
        output_wb = load_workbook(consensus_path)
        
        # Load template DCF Model
        template_wb = load_workbook("Template.xlsx", data_only=False)
        dcf_sheet = template_wb["DCF Model"]
        
        # Copy DCF Model to output workbook
        new_dcf_sheet = copy_sheet(dcf_sheet, output_wb, "DCF Model")
        
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
