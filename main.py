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

@app.post("/upload")
async def upload(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded consensus file
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        # Load the consensus file
        consensus_wb = load_workbook(consensus_path)
        
        # Load the template
        template = load_workbook("Template.xlsx")
        
        # Copy DCF Model sheet from template to consensus workbook
        dcf_sheet = template["DCF Model"]
        consensus_wb.create_sheet("DCF Model")
        new_sheet = consensus_wb["DCF Model"]
        
        # Copy all cells including values, styles, and formulas
        for row in dcf_sheet.iter_rows():
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
        
        # Update valuation date
        update_valuation_date(new_sheet)
        
        # Save to temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        consensus_wb.save(temp_file.name)
        
        return StreamingResponse(
            open(temp_file.name, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model_Output.xlsx"},
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        if os.path.exists(consensus_path):
            os.remove(consensus_path)

def update_valuation_date(sheet):
    """Update the valuation date in the DCF model to current date"""
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Valuation Date":
                sheet.cell(row=cell.row, column=cell.column + 2).value = datetime.now().date()
                return
