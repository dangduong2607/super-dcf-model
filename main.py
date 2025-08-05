from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
import shutil
import os
import tempfile
from pathlib import Path

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet_with_formatting(source_sheet, target_wb, new_sheet_name):
    """Copy a sheet while preserving all formatting and formulas"""
    # Create new sheet in target workbook
    new_sheet = target_wb.create_sheet(title=new_sheet_name)
    
    # Copy dimensions
    for row_idx, dim in source_sheet.row_dimensions.items():
        new_sheet.row_dimensions[row_idx].height = dim.height
    for col_idx, dim in source_sheet.column_dimensions.items():
        new_sheet.column_dimensions[col_idx].width = dim.width
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy cells with styles
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(
                row=cell.row,
                column=cell.column,
                value=cell.value
            )
            if cell.has_style:
                new_cell.font = cell.font
                new_cell.border = cell.border
                new_cell.fill = cell.fill
                new_cell.number_format = cell.number_format
                new_cell.alignment = cell.alignment
    
    # Copy gridline setting
    new_sheet.sheet_view.showGridLines = source_sheet.sheet_view.showGridLines
    
    return new_sheet

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Create temporary directory for processing
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save consensus file
        consensus_path = os.path.join(temp_dir, "consensus.xlsx")
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)
        
        # Save profile file if provided
        profile_path = None
        if profile:
            profile_path = os.path.join(temp_dir, "profile.xlsx")
            try:
                with open(profile_path, "wb") as f:
                    shutil.copyfileobj(profile.file, f)
                # Validate it's a real Excel file
                test_wb = load_workbook(profile_path, read_only=True)
                test_wb.close()
            except Exception as e:
                print(f"Invalid profile file: {str(e)}")
                profile_path = None  # Treat as if no profile was uploaded
        
        # Generate output path
        output_path = os.path.join(temp_dir, "DCF_Model_Output.xlsx")
        
        try:
            # Load consensus as base workbook
            base_wb = load_workbook(consensus_path)
            
            # Handle profile if valid
            if profile_path:
                profile_wb = load_workbook(profile_path)
                if "Public Company" in profile_wb.sheetnames:
                    # Remove existing if present
                    if "Public Company" in base_wb.sheetnames:
                        base_wb.remove(base_wb["Public Company"])
                    
                    # Copy with formatting
                    source_sheet = profile_wb["Public Company"]
                    copy_sheet_with_formatting(source_sheet, base_wb, "Public Company")
            
            # Handle template
            template_path = Path(__file__).parent / "Template.xlsx"
            if not template_path.exists():
                raise HTTPException(status_code=500, detail="Template file not found")
            
            template_wb = load_workbook(template_path)
            
            # Add DCF Model sheet
            if "DCF Model" in base_wb.sheetnames:
                base_wb.remove(base_wb["DCF Model"])
            
            if "DCF Model" in template_wb.sheetnames:
                dcf_sheet = template_wb["DCF Model"]
                new_dcf_sheet = copy_sheet_with_formatting(dcf_sheet, base_wb, "DCF Model")
                new_dcf_sheet.sheet_view.showGridLines = False
            else:
                raise HTTPException(status_code=500, detail="Template missing DCF Model sheet")
            
            # Save final workbook
            base_wb.save(output_path)
            
            # Return the file
            return StreamingResponse(
                open(output_path, "rb"),
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
            )
        
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Error processing files: {str(e)}")
