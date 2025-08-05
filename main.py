from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
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

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = "consensus.xlsx"
    with open(consensus_path, "wb") as f:
        shutil.copyfileobj(consensus.file, f)

    profile_path = None
    if profile:
        profile_path = "profile.xlsx"
        with open(profile_path, "wb") as f:
            shutil.copyfileobj(profile.file, f)

    output_stream = build_final_excel(consensus_path, profile_path)
    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
    )

def copy_sheet_contents(source, target):
    # Copy cell values, styles, and formulas
    for row in source.iter_rows():
        for cell in row:
            new_cell = target.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell._style = cell._style
            if cell.hyperlink:
                new_cell.hyperlink = cell.hyperlink
            if cell.comment:
                new_cell.comment = cell.comment

    # Copy column widths
    for col_letter, dim in source.column_dimensions.items():
        target.column_dimensions[col_letter].width = dim.width

    # Copy row heights
    for row_num, dim in source.row_dimensions.items():
        target.row_dimensions[row_num].height = dim.height

    # Copy merged cells
    for merged_range in source.merged_cells.ranges:
        target.merge_cells(str(merged_range))

def build_final_excel(consensus_path, profile_path):
    # Load user's uploaded consensus workbook
    wb = load_workbook(consensus_path)

    # Load template and DCF Model sheet
    template_wb = load_workbook("Template.xlsx", data_only=False)
    dcf_sheet_template = template_wb["DCF Model"]

    # Create a new sheet and copy DCF Model contents into it
    if "DCF Model" in wb.sheetnames:
        del wb["DCF Model"]
    new_dcf = wb.create_sheet("DCF Model")
    copy_sheet_contents(dcf_sheet_template, new_dcf)

    # If profile provided, copy Public Company sheet
    if profile_path and os.path.exists(profile_path):
        try:
            profile_wb = load_workbook(profile_path, data_only=False)
            if "Public Company" in profile_wb.sheetnames:
                if "Public Company" in wb.sheetnames:
                    del wb["Public Company"]
                new_profile = wb.create_sheet("Public Company")
                copy_sheet_contents(profile_wb["Public Company"], new_profile)
        except Exception as e:
            print(f"Skipping profile copy due to error: {e}")

    # Save result to in-memory stream
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
