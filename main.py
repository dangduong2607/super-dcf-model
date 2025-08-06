from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import shutil
import os
from tempfile import NamedTemporaryFile
from copy import copy

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name=None):
    """Copy a sheet from source workbook to target workbook with full formatting"""
    # Create new sheet in target workbook
    new_sheet_name = sheet_name or source_sheet.title
    new_sheet = target_wb.create_sheet(new_sheet_name)
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy row dimensions
    for idx, dim in source_sheet.row_dimensions.items():
        new_dim = copy(dim)
        new_sheet.row_dimensions[idx] = new_dim
    
    # Copy column dimensions
    for col_letter, dim in source_sheet.column_dimensions.items():
        new_dim = copy(dim)
        new_sheet.column_dimensions[col_letter] = new_dim
    
    # Copy cell values and styles
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
    
    # Copy conditional formatting
    for cf in source_sheet.conditional_formatting:
        new_sheet.conditional_formatting.add(cf)
    
    return new_sheet

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path, profile_path = None, None
    try:
        # Save uploaded files
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Generate output file
        output_path = generate_output_file(consensus_path, profile_path)

        # Streaming response with auto-cleanup
        def iterfile(file_path):
            with open(file_path, "rb") as f:
                while chunk := f.read(4096):
                    yield chunk
            # Delete after streaming completes
            os.remove(file_path)

        return StreamingResponse(
            iterfile(output_path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except KeyError as e:
        raise HTTPException(
            status_code=400, 
            detail=f"Required sheet missing: {str(e)}"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Cleanup input files
        for path in [consensus_path, profile_path]:
            if path and os.path.exists(path):
                os.remove(path)

def generate_output_file(consensus_path, profile_path):
    try:
        # Create new workbook
        output_wb = Workbook()
        output_wb.remove(output_wb.active)  # Remove default sheet
        
        # Load template and copy DCF Model sheet
        template_wb = load_workbook("Template.xlsx", data_only=False)
        copy_sheet(template_wb["DCF Model"], output_wb)
        template_wb.close()
        
        # Load consensus and copy Consensus sheet
        consensus_wb = load_workbook(consensus_path, data_only=False)
        copy_sheet(consensus_wb["Consensus"], output_wb)
        consensus_wb.close()
        
        # Load profile and copy Public Company sheet (if provided)
        if profile_path:
            profile_wb = load_workbook(profile_path, data_only=False)
            copy_sheet(profile_wb["Public Company"], output_wb)
            profile_wb.close()
        
        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        return temp_file.name
        
    except KeyError as e:
        sheet_name = str(e).strip("'")
        raise KeyError(f"The required sheet '{sheet_name}' was not found in uploaded file")
    except Exception as e:
        raise Exception(f"Error generating output file: {str(e)}")
