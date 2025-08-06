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
    
    # Copy conditional formatting
    for cf in source_sheet.conditional_formatting:
        target_sheet.conditional_formatting.add(cf)

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Create a new workbook
    output_wb = Workbook()
    # Remove default sheet created by Workbook()
    output_wb.remove(output_wb.active)
    
    consensus_path = None
    profile_path = None
    output_temp = None
    
    try:
        # Save consensus file
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        # Save profile file if provided
        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Load template and get DCF Model sheet
        template_wb = load_workbook("Template.xlsx", data_only=False)
        if "DCF Model" not in template_wb.sheetnames:
            raise HTTPException(status_code=500, detail="Template is missing 'DCF Model' sheet")
        
        dcf_sheet = template_wb["DCF Model"]
        new_dcf_sheet = output_wb.create_sheet("DCF Model")
        copy_sheet(dcf_sheet, new_dcf_sheet)

        # Add Consensus sheet
        consensus_wb = load_workbook(consensus_path, data_only=False)
        if "Consensus" not in consensus_wb.sheetnames:
            raise HTTPException(status_code=400, detail="Consensus file must contain 'Consensus' sheet")
        
        consensus_sheet = consensus_wb["Consensus"]
        new_consensus_sheet = output_wb.create_sheet("Consensus")
        copy_sheet(consensus_sheet, new_consensus_sheet)

        # Add Public Company sheet if provided
        if profile_path:
            profile_wb = load_workbook(profile_path, data_only=False)
            if "Public Company" not in profile_wb.sheetnames:
                raise HTTPException(status_code=400, detail="Profile file must contain 'Public Company' sheet")
            
            profile_sheet = profile_wb["Public Company"]
            new_profile_sheet = output_wb.create_sheet("Public Company")
            copy_sheet(profile_sheet, new_profile_sheet)

        # Save output to temporary file
        output_temp = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(output_temp.name)
        
        # Return the generated file
        return StreamingResponse(
            open(output_temp.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        for path in [consensus_path, profile_path, output_temp.name if output_temp else None]:
            if path and os.path.exists(path):
                os.remove(path)
