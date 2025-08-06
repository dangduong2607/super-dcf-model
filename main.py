from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os
import tempfile
from copy import copy
from datetime import datetime

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name):
    """Copies a sheet with all formatting, formulas, and structure intact"""
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy row dimensions
    for row_idx, row_dim in source_sheet.row_dimensions.items():
        new_row_dim = new_sheet.row_dimensions[row_idx]
        new_row_dim.height = row_dim.height
        new_row_dim.visible = row_dim.visible
        new_row_dim.style = row_dim.style
    
    # Copy column dimensions
    for col_idx, col_dim in source_sheet.column_dimensions.items():
        new_col_dim = new_sheet.column_dimensions[col_idx]
        new_col_dim.width = col_dim.width
        new_col_dim.visible = col_dim.visible
        new_col_dim.style = col_dim.style
    
    # Copy cells with all properties
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(
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
    new_sheet.sheet_format = copy(source_sheet.sheet_format)
    new_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    new_sheet.page_setup = copy(source_sheet.page_setup)
    
    return new_sheet

def update_valuation_date(sheet):
    """Update the valuation date in the DCF model to current date"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value and "Valuation Date" in str(cell.value):
                sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return

@app.post("/upload")
async def upload(
    background_tasks: BackgroundTasks,
    consensus: UploadFile = File(...),
    profile: UploadFile = File(None)
):
    temp_files = []
    try:
        # Save consensus file
        consensus_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        shutil.copyfileobj(consensus.file, consensus_file)
        consensus_file.close()
        temp_files.append(consensus_file.name)
        
        # Save profile file if exists
        profile_file = None
        if profile:
            profile_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            shutil.copyfileobj(profile.file, profile_file)
            profile_file.close()
            temp_files.append(profile_file.name)
        
        # Load workbooks
        output_wb = load_workbook(consensus_file.name)
        
        # Add profile sheets if provided
        if profile:
            profile_wb = load_workbook(profile_file.name)
            for sheet_name in profile_wb.sheetnames:
                if sheet_name in output_wb.sheetnames:
                    continue  # Skip if sheet already exists
                sheet = profile_wb[sheet_name]
                copy_sheet(sheet, output_wb, sheet_name)
        
        # Load template and copy DCF Model
        template_wb = load_workbook("Template.xlsx")
        template_sheet = template_wb["DCF Model"]
        
        # Remove existing DCF Model if present
        if "DCF Model" in output_wb.sheetnames:
            output_wb.remove(output_wb["DCF Model"])
        
        # Copy DCF Model from template
        new_sheet = copy_sheet(template_sheet, output_wb, "DCF Model")
        update_valuation_date(new_sheet)
        
        # Save output
        output_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(output_file.name)
        output_file.close()
        temp_files.append(output_file.name)
        
        # Schedule cleanup
        background_tasks.add_task(lambda: [os.remove(f) for f in temp_files])
        
        # Return the generated file
        return StreamingResponse(
            open(output_file.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        # Cleanup on error
        for f in temp_files:
            if os.path.exists(f):
                os.remove(f)
        raise HTTPException(status_code=500, detail=f"Error processing files: {str(e)}")
