from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from copy import copy
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
        # Read files into memory
        consensus_data = await consensus.read()
        profile_data = await profile.read() if profile else None
        
        # Generate output in memory
        output = generate_output_file(consensus_data, profile_data)
        
        # Return the generated file
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def generate_output_file(consensus_data, profile_data):
    # Create a new workbook in memory
    output = io.BytesIO()
    
    # Create new workbook
    output_wb = Workbook()
    # Remove default sheet
    output_wb.remove(output_wb.active)

    try:
        # Load template from file system
        template_path = "Template.xlsx"
        if not os.path.exists(template_path):
            raise HTTPException(status_code=500, detail="Template file not found")
        
        template_wb = load_workbook(template_path, data_only=False)
        
        # Verify DCF Model sheet exists
        if "DCF Model" not in template_wb.sheetnames:
            raise HTTPException(status_code=500, detail="DCF Model sheet not found in template")
        
        # Copy DCF Model sheet
        template_sheet = template_wb["DCF Model"]
        new_dcf_sheet = output_wb.create_sheet("DCF Model")
        copy_sheet(template_sheet, new_dcf_sheet)

        # Load consensus data
        consensus_stream = io.BytesIO(consensus_data)
        consensus_wb = load_workbook(consensus_stream, data_only=False)
        
        # Verify Consensus sheet exists
        if "Consensus" not in consensus_wb.sheetnames:
            raise HTTPException(status_code=400, detail="Consensus sheet not found in uploaded file")
        
        # Copy Consensus sheet
        consensus_sheet = consensus_wb["Consensus"]
        new_consensus_sheet = output_wb.create_sheet("Consensus")
        copy_sheet(consensus_sheet, new_consensus_sheet)

        # Handle profile data if provided
        if profile_data:
            profile_stream = io.BytesIO(profile_data)
            profile_wb = load_workbook(profile_stream, data_only=False)
            
            # Verify Public Company sheet exists
            if "Public Company" not in profile_wb.sheetnames:
                raise HTTPException(status_code=400, detail="Public Company sheet not found in profile file")
            
            # Copy Public Company sheet
            profile_sheet = profile_wb["Public Company"]
            new_profile_sheet = output_wb.create_sheet("Public Company")
            copy_sheet(profile_sheet, new_profile_sheet)
        
        # Save to memory buffer
        output_wb.save(output)
        output.seek(0)
        
        return output

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating output: {str(e)}")
