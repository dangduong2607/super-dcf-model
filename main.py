from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
from io import BytesIO
import os
import shutil

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name):
    """Copy a sheet with full fidelity (formats, formulas, styles)"""
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy row dimensions
    for r_idx, row_dim in source_sheet.row_dimensions.items():
        new_sheet.row_dimensions[r_idx].height = row_dim.height
    
    # Copy column dimensions
    for c_idx, col_dim in source_sheet.column_dimensions.items():
        new_sheet.column_dimensions[c_idx].width = col_dim.width
    
    # Copy cells with values, formulas, and formatting
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(
                row=cell.row, 
                column=cell.column, 
                value=cell.value
            )
            
            if cell.data_type == 'f':  # Formula
                new_cell.value = cell.value  # Preserves formula
            
            # Copy all styling properties
            new_cell.font = copy(cell.font)
            new_cell.border = copy(cell.border)
            new_cell.fill = copy(cell.fill)
            new_cell.number_format = cell.number_format
            new_cell.protection = copy(cell.protection)
            new_cell.alignment = copy(cell.alignment)
    
    # Copy sheet view options
    new_sheet.sheet_view = copy(source_sheet.sheet_view)
    return new_sheet

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Process consensus file
    consensus_bytes = await consensus.read()
    consensus_workbook = load_workbook(filename=BytesIO(consensus_bytes))
    
    # Remove existing DCF/Public Company sheets if present
    for sheet_name in ["DCF Model", "Public Company"]:
        if sheet_name in consensus_workbook.sheetnames:
            del consensus_workbook[sheet_name]
    
    # Add Public Company sheet if provided
    if profile:
        profile_bytes = await profile.read()
        profile_workbook = load_workbook(filename=BytesIO(profile_bytes))
        if "Public Company" in profile_workbook.sheetnames:
            public_sheet = profile_workbook["Public Company"]
            copy_sheet(public_sheet, consensus_workbook, "Public Company")
    
    # Add DCF Model from template
    template = load_workbook("Template.xlsx")
    dcf_sheet = template["DCF Model"]
    new_dcf_sheet = copy_sheet(dcf_sheet, consensus_workbook, "DCF Model")
    
    # Hide gridlines for DCF sheet
    new_dcf_sheet.sheet_view.showGridLines = False
    
    # Prepare output
    output = BytesIO()
    consensus_workbook.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"}
    )
