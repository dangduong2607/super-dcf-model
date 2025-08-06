from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
from tempfile import NamedTemporaryFile
from datetime import datetime
from openpyxl.utils import get_column_letter
from copy import copy

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
        # Save uploaded consensus file
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        # Save uploaded profile file if provided
        profile_path = None
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Load the consensus file (this will be our base workbook)
        output_wb = load_workbook(consensus_path)
        
        # Load the template to get the DCF Model sheet
        template = load_workbook("Template.xlsx")
        dcf_sheet = template["DCF Model"]
        
        # Create a new sheet in the output workbook
        if "DCF Model" in output_wb.sheetnames:
            output_wb.remove(output_wb["DCF Model"])
        new_sheet = output_wb.create_sheet("DCF Model")
        
        # Copy all cells with values, styles, and formulas
        copy_sheet_content(dcf_sheet, new_sheet)
        
        # Update valuation date
        update_valuation_date(new_sheet)
        
        # If profile file was provided, add its sheets to the output
        if profile_path:
            profile_wb = load_workbook(profile_path)
            for sheet_name in profile_wb.sheetnames:
                if sheet_name in output_wb.sheetnames:
                    output_wb.remove(output_wb[sheet_name])
                profile_sheet = profile_wb[sheet_name]
                new_profile_sheet = output_wb.create_sheet(sheet_name)
                copy_sheet_content(profile_sheet, new_profile_sheet)
        
        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        
        return StreamingResponse(
            open(temp_file.name, "rb"),
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

def copy_sheet_content(source_sheet, target_sheet):
    """Copy all content from source sheet to target sheet including styles, formulas, and dimensions"""
    # Copy merged cells
    for merge in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merge))
    
    # Copy column dimensions
    for col in range(1, source_sheet.max_column + 1):
        col_letter = get_column_letter(col)
        target_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
    
    # Copy row dimensions
    for row in range(1, source_sheet.max_row + 1):
        target_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height
    
    # Copy cell values and styles
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = target_sheet.cell(
                row=cell.row, 
                column=cell.column,
                value=cell.value
            )
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)
    
    # Copy sheet properties
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.page_setup = copy(source_sheet.page_setup)

def update_valuation_date(sheet):
    """Update the valuation date in the DCF model to current date"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "Valuation Date" in cell.value:
                sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return
