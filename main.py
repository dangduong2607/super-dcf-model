from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import openpyxl
import shutil
import os
from copy import copy

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    # Save uploaded files
    consensus_path = "consensus.xlsx"
    with open(consensus_path, "wb") as f:
        shutil.copyfileobj(consensus.file, f)

    profile_path = None
    if profile:
        profile_path = "profile.xlsx"
        with open(profile_path, "wb") as f:
            shutil.copyfileobj(profile.file, f)

    # Generate the combined Excel file
    output_path = "DCF_Model_Output.xlsx"
    build_final_excel(consensus_path, profile_path, output_path)

    return StreamingResponse(
        open(output_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
    )

def build_final_excel(consensus_path, profile_path, output_path):
    # Helper function to copy sheets with full fidelity
    def copy_sheet(source_sheet, target_wb, sheet_name):
        new_sheet = target_wb.create_sheet(sheet_name)
        
        # Copy cells with values and formatting
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
        
        # Copy column dimensions
        for col_letter, col_dim in source_sheet.column_dimensions.items():
            new_col_dim = new_sheet.column_dimensions[col_letter]
            new_col_dim.width = col_dim.width
            new_col_dim.hidden = col_dim.hidden
        
        # Copy row dimensions
        for row_idx, row_dim in source_sheet.row_dimensions.items():
            new_row_dim = new_sheet.row_dimensions[row_idx]
            new_row_dim.height = row_dim.height
            new_row_dim.hidden = row_dim.hidden
        
        # Copy merged cells
        for merged_range in source_sheet.merged_cells.ranges:
            new_sheet.merge_cells(str(merged_range))
        
        return new_sheet

    # Start with consensus workbook as base
    wb = openpyxl.load_workbook(consensus_path)

    # Add Public Company sheet from profile if available
    if profile_path and os.path.exists(profile_path):
        profile_wb = openpyxl.load_workbook(profile_path)
        if "Public Company" in profile_wb.sheetnames:
            # Remove existing sheet if present
            if "Public Company" in wb.sheetnames:
                del wb["Public Company"]
            source_sheet = profile_wb["Public Company"]
            copy_sheet(source_sheet, wb, "Public Company")

    # Add DCF Model from template
    template_wb = openpyxl.load_workbook("Template.xlsx")
    if "DCF Model" in template_wb.sheetnames:
        # Remove existing sheet if present
        if "DCF Model" in wb.sheetnames:
            del wb["DCF Model"]
        
        source_sheet = template_wb["DCF Model"]
        dcf_sheet = copy_sheet(source_sheet, wb, "DCF Model")
        
        # Turn off gridlines for cleaner look
        dcf_sheet.sheet_view.showGridLines = False

    # Save the final workbook
    wb.save(output_path)
