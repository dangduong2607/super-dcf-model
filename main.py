from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
from tempfile import NamedTemporaryFile
from datetime import datetime

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_wb, target_sheet_name):
    """
    Copies an entire sheet from source to target workbook while preserving formulas and references
    """
    # Create new sheet in target workbook
    new_sheet = target_wb.create_sheet(target_sheet_name)
    
    # Copy all cells including values, formulas, styles, and formatting
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = new_sheet.cell(row=cell.row, column=cell.column, 
                                    value=cell.value)
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.border = cell.border.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection.copy()
                new_cell.alignment = cell.alignment.copy()
            if cell.hyperlink:
                new_cell.hyperlink = cell.hyperlink
                new_cell.style = "Hyperlink"
    
    # Copy merged cells
    for merged_cell_range in source_sheet.merged_cells.ranges:
        new_sheet.merge_cells(str(merged_cell_range))
    
    # Copy column dimensions
    for col, dimension in source_sheet.column_dimensions.items():
        new_sheet.column_dimensions[col].width = dimension.width
        new_sheet.column_dimensions[col].hidden = dimension.hidden
    
    # Copy row dimensions
    for row, dimension in source_sheet.row_dimensions.items():
        new_sheet.row_dimensions[row].height = dimension.height
        new_sheet.row_dimensions[row].hidden = dimension.hidden

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded consensus file
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_consensus:
            shutil.copyfileobj(consensus.file, temp_consensus)
            consensus_path = temp_consensus.name

        # Load the consensus file
        consensus_wb = load_workbook(consensus_path, data_only=False)  # Important: data_only=False to preserve formulas
        
        # If profile file is provided, add its Public Company data
        if profile:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_profile:
                shutil.copyfileobj(profile.file, temp_profile)
                profile_path = temp_profile.name
            
            profile_wb = load_workbook(profile_path, data_only=False)
            if "Public Company" in profile_wb.sheetnames:
                # Remove existing sheet if present
                if "Public Company" in consensus_wb.sheetnames:
                    consensus_wb.remove(consensus_wb["Public Company"])
                
                # Copy the entire Public Company sheet
                copy_sheet(profile_wb["Public Company"], consensus_wb, "Public Company")
            
            os.unlink(profile_path)

        # Load the template
        template = load_workbook("Template.xlsx", data_only=False)
        
        # Remove existing DCF Model sheet if it exists
        if "DCF Model" in consensus_wb.sheetnames:
            consensus_wb.remove(consensus_wb["DCF Model"])
        
        # Copy the entire DCF Model sheet from template
        copy_sheet(template["DCF Model"], consensus_wb, "DCF Model")
        
        # Update valuation date
        update_valuation_date(consensus_wb["DCF Model"])
        
        # Save to temporary output file
        temp_output = NamedTemporaryFile(delete=False, suffix=".xlsx")
        consensus_wb.save(temp_output.name)
        consensus_wb.close()
        
        return StreamingResponse(
            open(temp_output.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model_Output.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        if os.path.exists(consensus_path):
            os.unlink(consensus_path)
        if 'temp_output' in locals() and os.path.exists(temp_output.name):
            os.unlink(temp_output.name)

def update_valuation_date(sheet):
    """Updates the valuation date in the DCF model to current date"""
    for row in sheet.iter_rows(min_row=1, max_row=500, min_col=1, max_col=20):  # A1:T500 range
        for cell in row:
            if isinstance(cell.value, str) and "Valuation Date" in cell.value:
                sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return
