from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
from tempfile import NamedTemporaryFile
import traceback

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_wb, target_wb, sheet_name):
    """Copy a sheet with all content and formatting to target workbook"""
    if sheet_name not in source_wb.sheetnames:
        return False
    
    source_sheet = source_wb[sheet_name]
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy cell values and formatting
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.border = cell.border.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.number_format = cell.number_format
                new_cell.alignment = cell.alignment.copy()
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy column dimensions
    for col in source_sheet.column_dimensions:
        target_col = new_sheet.column_dimensions[col]
        source_col = source_sheet.column_dimensions[col]
        target_col.width = source_col.width
        target_col.hidden = source_col.hidden
    
    # Copy row dimensions
    for row_idx in source_sheet.row_dimensions:
        target_row = new_sheet.row_dimensions[row_idx]
        source_row = source_sheet.row_dimensions[row_idx]
        target_row.height = source_row.height
        target_row.hidden = source_row.hidden
    
    return True

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_temp = None
    profile_temp = None
    output_temp = None
    
    try:
        # Save consensus file
        consensus_temp = NamedTemporaryFile(delete=False, suffix=".xlsx")
        shutil.copyfileobj(consensus.file, consensus_temp)
        consensus_temp.close()
        
        # Save profile file if provided
        if profile:
            profile_temp = NamedTemporaryFile(delete=False, suffix=".xlsx")
            shutil.copyfileobj(profile.file, profile_temp)
            profile_temp.close()

        # Load template and keep only DCF Model sheet
        template = load_workbook("Template.xlsx")
        for sheet_name in template.sheetnames[:]:
            if sheet_name != "DCF Model":
                del template[sheet_name]
        
        # Copy Consensus sheet
        consensus_wb = load_workbook(consensus_temp.name)
        if not copy_sheet(consensus_wb, template, "Consensus"):
            raise HTTPException(status_code=400, detail="Consensus file must contain 'Consensus' sheet")
        
        # Copy Public Company sheet if provided
        if profile_temp:
            profile_wb = load_workbook(profile_temp.name)
            if not copy_sheet(profile_wb, template, "Public Company"):
                raise HTTPException(status_code=400, detail="Profile file must contain 'Public Company' sheet")

        # Save output to temporary file
        output_temp = NamedTemporaryFile(delete=False, suffix=".xlsx")
        template.save(output_temp.name)
        
        # Return generated file
        return StreamingResponse(
            open(output_temp.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except HTTPException as he:
        raise he
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Processing error: {str(e)}")
    finally:
        # Cleanup temporary files
        for temp_file in [consensus_temp, profile_temp, output_temp]:
            if temp_file and os.path.exists(temp_file.name):
                os.unlink(temp_file.name)
