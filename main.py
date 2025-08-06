from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
import shutil
import os
from tempfile import NamedTemporaryFile

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
        # Save consensus file temporarily
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        # Load the template and create output workbook
        output_wb = load_workbook("Template.xlsx")
        
        # Keep only DCF Model sheet
        for sheet_name in output_wb.sheetnames[:]:  # Make a copy of sheetnames list
            if sheet_name != "DCF Model":
                output_wb.remove(output_wb[sheet_name])
        
        # Add Consensus sheet from user's file
        consensus_wb = load_workbook(consensus_path)
        if "Consensus" in consensus_wb.sheetnames:
            consensus_sheet = consensus_wb["Consensus"]
            new_sheet = output_wb.create_sheet("Consensus")
            
            # Simple cell copy (values only)
            for row in consensus_sheet.iter_rows():
                for cell in row:
                    new_sheet.cell(
                        row=cell.row, 
                        column=cell.column,
                        value=cell.value
                    )
        else:
            raise HTTPException(status_code=400, detail="Consensus file must contain 'Consensus' sheet")

        # Save to temporary file
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            output_wb.save(temp_file.name)
            output_path = temp_file.name

        # Return the generated file
        return StreamingResponse(
            open(output_path, "rb"),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except HTTPException as he:
        raise he
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Clean up temporary files
        if os.path.exists(consensus_path):
            os.remove(consensus_path)
        if 'output_path' in locals() and os.path.exists(output_path):
            os.remove(output_path)
