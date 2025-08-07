from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import shutil
import os
import tempfile
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, sheet_name):
    """Copy a sheet with all formatting, formulas, and dimensions"""
    new_sheet = target_wb.create_sheet(sheet_name)
    
    # Copy column dimensions
    for col in range(1, source_sheet.max_column + 1):
        col_letter = get_column_letter(col)
        if col_letter in source_sheet.column_dimensions:
            new_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width
    
    # Copy row dimensions
    for row in range(1, source_sheet.max_row + 1):
        if row in source_sheet.row_dimensions:
            new_sheet.row_dimensions[row].height = source_sheet.row_dimensions[row].height
    
    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_range))
    
    # Copy cells with values, formulas, and formatting
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(
                row=cell.row,
                column=cell.column,
                value=cell.value
            )
            
            # Preserve formulas
            if cell.data_type == 'f':
                new_cell.value = cell.value
            
            # Copy all styling attributes
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.border = cell.border.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection.copy()
                new_cell.alignment = cell.alignment.copy()
    
    return new_sheet

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Create a temporary directory for all files
    with tempfile.TemporaryDirectory() as temp_dir:
        consensus_path = os.path.join(temp_dir, "consensus.xlsx")
        profile_path = os.path.join(temp_dir, "profile.xlsx") if profile else None
        output_path = os.path.join(temp_dir, "output.xlsm")
        
        try:
            # Save consensus file
            with open(consensus_path, "wb") as f:
                shutil.copyfileobj(consensus.file, f)
            logger.info(f"Saved consensus file to: {consensus_path}")
            
            # Save profile file if provided
            if profile:
                with open(profile_path, "wb") as f:
                    shutil.copyfileobj(profile.file, f)
                logger.info(f"Saved profile file to: {profile_path}")
            
            # Verify template exists
            template_path = "Template.xlsm"
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Template file not found: {os.path.abspath(template_path)}")
            
            # Load template
            output_wb = load_workbook(template_path, data_only=False, keep_vba=True)
            logger.info("Loaded template workbook")
            
            # Remove existing sheets except "DCF Model"
            for sheet_name in list(output_wb.sheetnames):
                if sheet_name != "DCF Model":
                    output_wb.remove(output_wb[sheet_name])
            logger.info("Cleaned template sheets")
            
            # Process consensus file
            consensus_wb = load_workbook(consensus_path)
            for sheet_name in consensus_wb.sheetnames:
                if sheet_name == "DCF Model":
                    continue
                source_sheet = consensus_wb[sheet_name]
                copy_sheet(source_sheet, output_wb, sheet_name)
            logger.info("Processed consensus file")
            
            # Process profile file if provided
            if profile and profile_path and os.path.exists(profile_path):
                profile_wb = load_workbook(profile_path)
                for sheet_name in profile_wb.sheetnames:
                    if sheet_name == "DCF Model":
                        continue
                    source_sheet = profile_wb[sheet_name]
                    copy_sheet(source_sheet, output_wb, sheet_name)
                logger.info("Processed profile file")
            
            # Save output
            output_wb.save(output_path)
            logger.info(f"Saved output to: {output_path}")
            
            # Return response
            return StreamingResponse(
                open(output_path, "rb"),
                media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
                headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsm"},
            )
            
        except Exception as e:
            logger.error(f"Processing error: {str(e)}", exc_info=True)
            raise HTTPException(status_code=500, detail=str(e))
