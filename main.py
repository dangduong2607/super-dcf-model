from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy
import os
import shutil
import tempfile

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def copy_sheet(source_sheet, target_sheet):
    # Copy cells
    for row in source_sheet.iter_rows():
        for cell in row:
            new_cell = target_sheet.cell(
                row=cell.row, 
                column=cell.column,
                value=cell.value if cell.data_type != 'f' else cell.formula
            )
            
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    # Copy merged cells
    for merged_range in source_sheet.merged_cells.ranges:
        target_sheet.merged_cells.add(str(merged_range))

    # Copy column dimensions
    for col_idx, source_dim in source_sheet.column_dimensions.items():
        target_dim = target_sheet.column_dimensions[col_idx]
        target_dim.width = source_dim.width
        target_dim.hidden = source_dim.hidden

    # Copy row dimensions
    for row_idx, source_dim in source_sheet.row_dimensions.items():
        target_dim = target_sheet.row_dimensions[row_idx]
        target_dim.height = source_dim.height
        target_dim.hidden = source_dim.hidden

    # Copy array formulas
    for array_formula in source_sheet.array_formulae:
        target_sheet.array_formulae[array_formula.ref] = array_formula

    # Copy sheet properties
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.page_setup = copy(source_sheet.page_setup)

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Create temp directory
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Save uploaded files
        consensus_path = os.path.join(temp_dir, "consensus.xlsx")
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        profile_path = None
        if profile:
            profile_path = os.path.join(temp_dir, "profile.xlsx")
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Generate output path
        output_path = os.path.join(temp_dir, "DCF_Model_Output.xlsx")
        
        # Create new workbook
        new_wb = openpyxl.Workbook()
        default_sheet = new_wb.active
        new_wb.remove(default_sheet)

        # Step 1: Copy all sheets from consensus
        consensus_wb = openpyxl.load_workbook(consensus_path, data_only=False)
        for sheet_name in consensus_wb.sheetnames:
            source_sheet = consensus_wb[sheet_name]
            new_sheet = new_wb.create_sheet(sheet_name)
            copy_sheet(source_sheet, new_sheet)

        # Step 2: Handle Public Company sheet
        public_company_added = False
        # Try from profile file
        if profile_path:
            profile_wb = openpyxl.load_workbook(profile_path, data_only=False)
            if "Public Company" in profile_wb.sheetnames:
                source_sheet = profile_wb["Public Company"]
                # Remove existing if present
                if "Public Company" in new_wb.sheetnames:
                    new_wb.remove(new_wb["Public Company"])
                new_sheet = new_wb.create_sheet("Public Company")
                copy_sheet(source_sheet, new_sheet)
                public_company_added = True
        
        # Try from consensus file if not added from profile
        if not public_company_added and "Public Company" in consensus_wb.sheetnames:
            source_sheet = consensus_wb["Public Company"]
            # Remove existing if present
            if "Public Company" in new_wb.sheetnames:
                new_wb.remove(new_wb["Public Company"])
            new_sheet = new_wb.create_sheet("Public Company")
            copy_sheet(source_sheet, new_sheet)
            public_company_added = True

        # Step 3: Copy DCF Model from template
        template_wb = openpyxl.load_workbook("Template.xlsx", data_only=False)
        if "DCF Model" in template_wb.sheetnames:
            source_sheet = template_wb["DCF Model"]
            # Remove existing if present
            if "DCF Model" in new_wb.sheetnames:
                new_wb.remove(new_wb["DCF Model"])
            new_sheet = new_wb.create_sheet("DCF Model")
            copy_sheet(source_sheet, new_sheet)

        # Save the new workbook
        new_wb.save(output_path)

        # Return the generated file
        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )
        
    finally:
        # Clean up temporary files
        shutil.rmtree(temp_dir, ignore_errors=True)
