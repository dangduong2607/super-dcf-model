from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
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

def copy_sheet(source_sheet, target_wb, sheet_name):
    """Copy a sheet with all content and formatting to target workbook"""
    # Create new sheet
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy all cell values and formatting
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
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy column dimensions
    for col_idx, col_dim in source_sheet.column_dimensions.items():
        if col_dim.width is not None:
            new_sheet.column_dimensions[col_idx].width = col_dim.width
        new_sheet.column_dimensions[col_idx].hidden = col_dim.hidden
    
    # Copy row dimensions
    for row_idx, row_dim in source_sheet.row_dimensions.items():
        if row_dim.height is not None:
            new_sheet.row_dimensions[row_idx].height = row_dim.height
        new_sheet.row_dimensions[row_idx].hidden = row_dim.hidden

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = None
    profile_path = None
    temp_output = None
    
    try:
        # Save consensus file
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_consensus:
            shutil.copyfileobj(consensus.file, temp_consensus)
            consensus_path = temp_consensus.name

        # Save profile file if provided
        if profile:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_profile:
                shutil.copyfileobj(profile.file, temp_profile)
                profile_path = temp_profile.name

        # Create output workbook from template
        output_wb = load_workbook("Template.xlsx")
        
        # Remove all sheets except DCF Model
        for sheet_name in output_wb.sheetnames[:]:
            if sheet_name != "DCF Model":
                output_wb.remove(output_wb[sheet_name])
        
        # Add consensus sheet
        consensus_wb = load_workbook(consensus_path)
        if "Consensus" not in consensus_wb.sheetnames:
            raise HTTPException(status_code=400, detail="Consensus file must contain 'Consensus' sheet")
        copy_sheet(consensus_wb["Consensus"], output_wb, "Consensus")
        
        # Add profile sheet if provided
        if profile_path:
            profile_wb = load_workbook(profile_path)
            if "Public Company" not in profile_wb.sheetnames:
                raise HTTPException(status_code=400, detail="Profile file must contain 'Public Company' sheet")
            copy_sheet(profile_wb["Public Company"], output_wb, "Public Company")
        
        # Save to temporary file
        temp_output = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_output.name)
        
        # Return the generated file
        return StreamingResponse(
            open(temp_output.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing files: {str(e)}")
    
    finally:
        # Clean up temporary files
        for path in [consensus_path, profile_path]:
            if path and os.path.exists(path):
                os.unlink(path)
        if temp_output and os.path.exists(temp_output.name):
            os.unlink(temp_output.name)
