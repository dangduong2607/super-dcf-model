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

def copy_range(source, target, start_cell="A1", end_cell="T500"):
    for row in source[start_cell:end_cell]:
        for cell in row:
            new_cell = target.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell._style = cell._style
            if cell.hyperlink:
                new_cell.hyperlink = cell.hyperlink
            if cell.comment:
                new_cell.comment = cell.comment

def build_final_excel(consensus_path, profile_path):
    # Step 1: Load base workbook from user
    wb = load_workbook(consensus_path)

    # Step 2: Load Template.xlsx and copy DCF Model
    template_wb = load_workbook("Template.xlsx", data_only=False)
    if "DCF Model" in wb.sheetnames:
        del wb["DCF Model"]
    dcf_sheet_template = template_wb["DCF Model"]
    dcf_sheet_new = wb.create_sheet("DCF Model")
    copy_range(dcf_sheet_template, dcf_sheet_new)

    # Step 3: Copy Public Company if profile provided
    if profile_path and os.path.exists(profile_path):
        try:
            profile_wb = load_workbook(profile_path, data_only=False)
            if "Public Company" in profile_wb.sheetnames:
                if "Public Company" in wb.sheetnames:
                    del wb["Public Company"]
                new_pub_sheet = wb.create_sheet("Public Company")
                copy_range(profile_wb["Public Company"], new_pub_sheet)
        except Exception as e:
            print(f"Error copying Public Company sheet: {e}")

    # Step 4: Stream output
    output_stream = io.BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    return output_stream
