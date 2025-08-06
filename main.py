from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from copy import copy
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

def copy_sheet(source_sheet, target_sheet):
    """Deep copy a sheet with all formatting, merged cells, and dimensions"""
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merged_range))
    
    # Copy row dimensions
    for idx, dim in source_sheet.row_dimensions.items():
        new_dim = copy(dim)
        target_sheet.row_dimensions[idx] = new_dim
    
    # Copy column dimensions
    for col_letter, dim in source_sheet.column_dimensions.items():
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

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded files temporarily
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        profile_path = None
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
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        if os.path.exists(consensus_path):
            os.remove(consensus_path)
        if profile_path and os.path.exists(profile_path):
            os.remove(profile_path)

def generate_output_file(consensus_path, profile_path):
    try:
        # Create a new workbook
        output_wb = Workbook()
        # Remove default sheet
        output_wb.remove(output_wb.active)

        # Load and copy DCF Model from template
        template_wb = load_workbook("Template.xlsx", data_only=False)
        if "DCF Model" not in template_wb.sheetnames:
            raise Exception("DCF Model sheet not found in template")
        
        template_sheet = template_wb["DCF Model"]
        new_dcf_sheet = output_wb.create_sheet("DCF Model")
        copy_sheet(template_sheet, new_dcf_sheet)

        # Load and copy Consensus sheet from user's file
        consensus_wb = load_workbook(consensus_path, data_only=False)
        if "Consensus" not in consensus_wb.sheetnames:
            raise Exception("Consensus sheet not found in uploaded file")
        
        consensus_sheet = consensus_wb["Consensus"]
        new_consensus_sheet = output_wb.create_sheet("Consensus")
        copy_sheet(consensus_sheet, new_consensus_sheet)

        # Copy Public Company sheet if provided
        if profile_path:
            profile_wb = load_workbook(profile_path, data_only=False)
            if "Public Company" not in profile_wb.sheetnames:
                raise Exception("Public Company sheet not found in profile file")
            
            profile_sheet = profile_wb["Public Company"]
            new_profile_sheet = output_wb.create_sheet("Public Company")
            copy_sheet(profile_sheet, new_profile_sheet)

        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        return temp_file.name
        
    except Exception as e:
        raise Exception(f"Error generating output file: {str(e)}")
