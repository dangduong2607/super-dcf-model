from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from copy import copy
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

def copy_sheet(source_sheet, target_sheet):
    """Copy all elements from source sheet to target sheet including styles and formatting"""
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_range))
    
    # Copy row dimensions
    for idx, dim in source_sheet.row_dimensions.items():
        if dim:  # Check if dimension exists
            new_dim = copy(dim)
            target_sheet.row_dimensions[idx] = new_dim
    
    # Copy column dimensions
    for col_letter, dim in source_sheet.column_dimensions.items():
        if dim:  # Check if dimension exists
            new_dim = copy(dim)
            target_sheet.column_dimensions[col_letter] = new_dim
    
    # Copy all cell values and styles
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
    
    # Copy conditional formatting
    for cf in source_sheet.conditional_formatting:
        target_sheet.conditional_formatting.add(cf)

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = None
    profile_path = None
    
    try:
        # Save uploaded files temporarily
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Generate the output file
        output_path = generate_output_file(consensus_path, profile_path)

        # Return the generated file
        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        error_trace = traceback.format_exc()
        print(f"Error occurred: {e}\n{error_trace}")
        raise HTTPException(
            status_code=500, 
            detail=f"Internal server error: {str(e)}"
        )
    
    finally:
        # Clean up temporary files
        if consensus_path and os.path.exists(consensus_path):
            os.remove(consensus_path)
        if profile_path and os.path.exists(profile_path):
            os.remove(profile_path)

def find_sheet_by_name(workbook, sheet_name):
    """Find sheet by case-insensitive name match"""
    sheet_lower = sheet_name.lower()
    for name in workbook.sheetnames:
        if name.lower() == sheet_lower:
            return workbook[name]
    return None

def generate_output_file(consensus_path, profile_path):
    try:
        # Create a new workbook
        output_wb = Workbook()
        output_wb.remove(output_wb.active)  # Remove default sheet
        
        # 1. Add DCF Model from template
        template_wb = load_workbook("Template.xlsx", data_only=False)
        dcf_sheet = find_sheet_by_name(template_wb, "DCF Model")
        if not dcf_sheet:
            raise ValueError("Template is missing 'DCF Model' sheet")
        
        new_dcf_sheet = output_wb.create_sheet("DCF Model")
        copy_sheet(dcf_sheet, new_dcf_sheet)
        
        # 2. Add Consensus sheet from user upload
        consensus_wb = load_workbook(consensus_path, data_only=False)
        consensus_sheet = find_sheet_by_name(consensus_wb, "Consensus")
        if not consensus_sheet:
            raise ValueError("Consensus file is missing 'Consensus' sheet")
        
        new_consensus_sheet = output_wb.create_sheet("Consensus")
        copy_sheet(consensus_sheet, new_consensus_sheet)
        
        # 3. Add Public Company sheet if profile was provided
        if profile_path:
            profile_wb = load_workbook(profile_path, data_only=False)
            public_company_sheet = find_sheet_by_name(profile_wb, "Public Company")
            if public_company_sheet:
                new_public_sheet = output_wb.create_sheet("Public Company")
                copy_sheet(public_company_sheet, new_public_sheet)
            else:
                print("Warning: Profile file is missing 'Public Company' sheet")
        
        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        return temp_file.name
        
    except Exception as e:
        raise Exception(f"Error generating output file: {str(e)}")
