from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
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

def copy_range(source, target, cell_range: str):
    min_col = ord(cell_range[0]) - ord('A') + 1
    max_col = ord(cell_range[3]) - ord('A') + 1
    min_row = int(cell_range[1:cell_range.find(':')].lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    max_row = int(cell_range[cell_range.find(':') + 1:].lstrip('ABCDEFGHIJKLMNOPQRSTUVWXYZ'))

    # Copy cells and styles
    for row in source[cell_range]:
        for cell in row:
            new_cell = target.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell._style = cell._style
            if cell.hyperlink:
                new_cell.hyperlink = cell.hyperlink
            if cell.comment:
                new_cell.comment = cell.comment

    # Column widths
    for col_letter in [chr(c) for c in range(ord('A'), ord('T') + 1)]:
        if col_letter in source.column_dimensions:
            target.column_dimensions[col_letter].width = source.column_dimensions[col_letter].width

    # Row heights
    for row_num in range(1, 501):
        if row_num in source.row_dimensions:
            target.row_dimensions[row_num].height = source.row_dimensions[row_num].height

    # Merged cells within A1:T500
    for merged_range in source.merged_cells.ranges:
        bounds = merged_range.bounds  # (min_col, min_row, max_col, max_row)
        if bounds[0] >= 1 and bounds[1] >= 1 and bounds[2] <= 20 and bounds[3] <= 500:
            target.merge_cells(str(merged_range))

def build_final_excel(consensus_path, profile_path):
    wb = load_workbook(consensus_path)

    # Copy DCF Model from Template.xlsx
    template_wb = load_workbook("Template.xlsx", data_only=False)
    if "DCF Model" in wb.sheetnames:
        del wb["DCF Model"]
    new_dcf = wb.create_sheet("DCF Model")
    copy_range(template_wb["DCF Model"], new_dcf, "A1:T500")

    # Copy Public Company if profile provided
    if profile_path and os.path.exists(profile_path):
        try:
            profile_wb = load_workbook(profile_path, data_only=False)
            if "Public Company" in profile_wb.sheetnames:
                if "Public Company" in wb.sheetnames:
                    del wb["Public Company"]
                new_profile = wb.create_sheet("Public Company")
                copy_range(profile_wb["Public Company"], new_profile, "A1:T500")
        except Exception as e:
            print(f"Skipping Public Company copy due to error: {e}")

    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
