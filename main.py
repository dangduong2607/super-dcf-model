from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.formula.translate import Translator
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

def copy_sheet_content(source_sheet, target_sheet):
    """
    Copies all content from source_sheet to target_sheet including formulas, values, and formatting
    """
    # Copy merged cells
    for merge in source_sheet.merged_cells.ranges:
        target_sheet.merge_cells(str(merge))
    
    # Copy column dimensions
    for col, dimension in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[col].width = dimension.width
    
    # Copy row dimensions
    for row, dimension in source_sheet.row_dimensions.items():
        target_sheet.row_dimensions[row].height = dimension.height
    
    # Copy all cells
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = target_sheet.cell(
                row=cell.row, 
                column=cell.column,
                value=cell.value
            )
            
            # Copy style
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.border = cell.border.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.number_format = cell.number_format
                new_cell.protection = cell.protection.copy()
                new_cell.alignment = cell.alignment.copy()
            
            # Copy hyperlinks
            if cell.hyperlink:
                new_cell.hyperlink = cell.hyperlink
                new_cell.style = "Hyperlink"
            
            # Copy data validation
            if cell.data_validation:
                new_cell.data_validation = cell.data_validation.copy()

def update_formula_references(sheet, old_sheet_name, new_sheet_name):
    """
    Updates formula references to point to the new sheet name
    """
    for row in sheet.iter_rows():
        for cell in row:
            if cell.data_type == 'f' and cell.value:  # If cell contains a formula
                # Replace sheet references in formulas
                if f"'{old_sheet_name}'!" in cell.value:
                    cell.value = cell.value.replace(f"'{old_sheet_name}'!", f"'{new_sheet_name}'!")
                elif f"{old_sheet_name}!" in cell.value:
                    cell.value = cell.value.replace(f"{old_sheet_name}!", f"{new_sheet_name}!")

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded consensus file
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_consensus:
            shutil.copyfileobj(consensus.file, temp_consensus)
            consensus_path = temp_consensus.name

        # Load the consensus file
        consensus_wb = load_workbook(consensus_path)
        
        # Get the original sheet names for reference
        original_sheet_names = consensus_wb.sheetnames
        
        # If profile file is provided, add its Public Company data
        if profile:
            with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_profile:
                shutil.copyfileobj(profile.file, temp_profile)
                profile_path = temp_profile.name
            
            profile_wb = load_workbook(profile_path)
            if "Public Company" in profile_wb.sheetnames:
                profile_sheet = profile_wb["Public Company"]
                
                # Remove existing sheet if present
                if "Public Company" in consensus_wb.sheetnames:
                    consensus_wb.remove(consensus_wb["Public Company"])
                
                # Create new sheet and copy all content
                new_profile_sheet = consensus_wb.create_sheet("Public Company")
                copy_sheet_content(profile_sheet, new_profile_sheet)
            
            os.unlink(profile_path)

        # Load the template
        template = load_workbook("Template.xlsx", data_only=False)  # Important to keep formulas
        template_dcf_sheet = template["DCF Model"]
        
        # Remove existing DCF Model sheet if it exists
        if "DCF Model" in consensus_wb.sheetnames:
            consensus_wb.remove(consensus_wb["DCF Model"])
        
        # Create new DCF Model sheet and copy all content
        new_dcf_sheet = consensus_wb.create_sheet("DCF Model")
        copy_sheet_content(template_dcf_sheet, new_dcf_sheet)
        
        # Update formula references to point to the original consensus sheet name
        # Assuming the first sheet is the consensus sheet
        if original_sheet_names:
            consensus_sheet_name = original_sheet_names[0]
            update_formula_references(new_dcf_sheet, "Consensus", consensus_sheet_name)
        
        # Update valuation date
        update_valuation_date(new_dcf_sheet)
        
        # Save to temporary output file
        temp_output = NamedTemporaryFile(delete=False, suffix=".xlsx")
        consensus_wb.save(temp_output.name)
        
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
