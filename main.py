from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os
import io

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name):
    """
    Copies an entire sheet from one workbook to another
    including values, formatting, formulas, and column widths
    """
    # Create new sheet in target workbook
    if sheet_name in target_wb.sheetnames:
        target_wb.remove(target_wb[sheet_name])
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy cell values, formulas, and formatting
    for row in source_sheet.iter_rows():
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
    
    # Copy column widths
    for col in range(1, source_sheet.max_column + 1):
        col_letter = get_column_letter(col)
        if col_letter in source_sheet.column_dimensions:
            new_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
    
    # Copy row heights
    for row in range(1, source_sheet.max_row + 1):
        if row in source_sheet.row_dimensions:
            new_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Create temporary files for uploaded files
    consensus_path = None
    profile_path = None
    
    try:
        # Save consensus file to temp
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_consensus:
            shutil.copyfileobj(consensus.file, temp_consensus)
            consensus_path = temp_consensus.name

        # Save profile file to temp if provided
        if profile:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_profile:
                shutil.copyfileobj(profile.file, temp_profile)
                profile_path = temp_profile.name

        # Create output in memory
        output = io.BytesIO()
        
        # Load the template
        template = load_workbook("Template.xlsx")
        
        # Create a new workbook starting with the template
        output_wb = load_workbook("Template.xlsx")
        
        # Remove all sheets except DCF Model
        for sheet_name in output_wb.sheetnames[:]:
            if sheet_name != "DCF Model":
                del output_wb[sheet_name]
        
        # Load consensus file and copy Consensus sheet
        consensus_wb = load_workbook(consensus_path)
        if "Consensus" not in consensus_wb.sheetnames:
            raise HTTPException(status_code=400, detail="Consensus sheet not found in uploaded file")
        
        copy_sheet(consensus_wb["Consensus"], output_wb, "Consensus")
        
        # If profile file is provided, copy Public Company sheet
        if profile_path:
            profile_wb = load_workbook(profile_path)
            if "Public Company" not in profile_wb.sheetnames:
                raise HTTPException(status_code=400, detail="Public Company sheet not found in profile file")
            
            copy_sheet(profile_wb["Public Company"], output_wb, "Public Company")
            profile_wb.close()
        
        consensus_wb.close()
        
        # Save to memory buffer
        output_wb.save(output)
        output.seek(0)
        
        # Return the generated file
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        if consensus_path and os.path.exists(consensus_path):
            os.unlink(consensus_path)
        if profile_path and os.path.exists(profile_path):
            os.unlink(profile_path)
