from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
import tempfile
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
    # Create temporary directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        consensus_path = os.path.join(temp_dir, "consensus.xlsx")
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        profile_path = None
        if profile:
            profile_path = os.path.join(temp_dir, "profile.xlsx")
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Generate the combined Excel file
        output_path = os.path.join(temp_dir, "DCF_Model_Output.xlsx")
        build_final_excel(consensus_path, profile_path, output_path)

        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

def copy_sheet(source_sheet, target_wb, new_sheet_name):
    """Copy a sheet with formatting, formulas, and styles to another workbook"""
    # Create new sheet in target workbook
    new_sheet = target_wb.create_sheet(new_sheet_name)
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy row dimensions (height)
    for row_idx, row_dim in source_sheet.row_dimensions.items():
        new_sheet.row_dimensions[row_idx].height = row_dim.height
    
    # Copy column dimensions (width)
    for col_idx, col_dim in source_sheet.column_dimensions.items():
        new_sheet.column_dimensions[col_idx].width = col_dim.width
    
    # Copy cell values and styles
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(
                row=cell.row, 
                column=cell.column, 
                value=cell.value
            )
            
            if cell.has_style:
                # Copy all style properties
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.alignment = copy(cell.alignment)
                new_cell.protection = copy(cell.protection)
    
    # Copy gridline setting
    new_sheet.sheet_view.showGridLines = source_sheet.sheet_view.showGridLines
    
    return new_sheet

def build_final_excel(consensus_path, profile_path, output_path):
    # Load consensus workbook as base
    base_wb = load_workbook(consensus_path)
    
    # Handle profile if provided
    if profile_path and os.path.exists(profile_path):
        profile_wb = load_workbook(profile_path)
        if "Public Company" in profile_wb.sheetnames:
            # Remove existing sheet if present
            if "Public Company" in base_wb.sheetnames:
                base_wb.remove(base_wb["Public Company"])
            
            # Copy profile sheet
            profile_sheet = profile_wb["Public Company"]
            copy_sheet(profile_sheet, base_wb, "Public Company")
    
    # Load template
    template_wb = load_workbook("Template.xlsx")
    
    # Handle DCF Model sheet
    if "DCF Model" in base_wb.sheetnames:
        base_wb.remove(base_wb["DCF Model"])
    
    # Copy DCF Model from template
    dcf_sheet = template_wb["DCF Model"]
    new_dcf_sheet = copy_sheet(dcf_sheet, base_wb, "DCF Model")
    
    # Ensure gridlines are turned off
    new_dcf_sheet.sheet_view.showGridLines = False
    
    # Save final workbook
    base_wb.save(output_path)
