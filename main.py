from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
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

def cleanup_files(*file_paths):
    """Clean up temporary files."""
    for path in file_paths:
        if path and os.path.exists(path):
            os.remove(path)

def generate_merged_solution(consensus_path, public_company_path=None):
    """Generates a new Excel file by combining data from uploaded and backend files."""
    try:
        # Define paths to the local backend files
        template_dcf_model_path = "Template.xlsx - DCF Model.csv"
        
        # Read the CSV files into pandas DataFrames
        dcf_model_df = pd.read_csv(template_dcf_model_path)
        consensus_upload_df = pd.read_csv(consensus_path)
        
        # Create a new workbook
        merged_wb = Workbook()

        # Write DCF Model data to a new sheet
        ws_dcf = merged_wb.create_sheet("DCF Model", 0)
        for r in dataframe_to_rows(dcf_model_df, index=False, header=True):
            ws_dcf.append(r)
        
        # Write Consensus data from the uploaded file to a new sheet
        ws_consensus = merged_wb.create_sheet("Consensus", 1)
        for r in dataframe_to_rows(consensus_upload_df, index=False, header=True):
            ws_consensus.append(r)

        # Handle the optional public company file
        if public_company_path and os.path.exists(public_company_path):
            public_company_df = pd.read_csv(public_company_path)
            ws_public = merged_wb.create_sheet("Public Company", 2)
            for r in dataframe_to_rows(public_company_df, index=False, header=True):
                ws_public.append(r)

        # Remove the default sheet created by openpyxl
        if "Sheet" in merged_wb.sheetnames:
            merged_wb.remove(merged_wb["Sheet"])
        
        # Save to a temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        merged_wb.save(temp_file.name)
        temp_file.close()
        
        return temp_file.name
        
    except FileNotFoundError as e:
        raise HTTPException(status_code=500, detail=f"Required file not found on the server: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error generating merged solution: {str(e)}")

@app.post("/upload")
async def upload(
    background_tasks: BackgroundTasks, 
    consensus: UploadFile = File(...),
    public_company: UploadFile = File(None)
):
    consensus_path = None
    public_company_path = None
    output_path = None

    try:
        # Save uploaded files temporarily
        consensus_path = "temp_consensus_upload.csv"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)
        
        if public_company:
            public_company_path = "temp_public_company.csv"
            with open(public_company_path, "wb") as f:
                shutil.copyfileobj(public_company.file, f)

        # Generate the output file
        output_path = generate_merged_solution(consensus_path, public_company_path)

        # Add cleanup to run in the background after the response is sent
        cleanup_files_list = [consensus_path, output_path]
        if public_company_path:
            cleanup_files_list.append(public_company_path)
        background_tasks.add_task(cleanup_files, *cleanup_files_list)

        # Return the generated file using a generator to stream its content
        def file_iterator():
            with open(output_path, "rb") as f:
                yield from f

        return StreamingResponse(
            file_iterator(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=Merged Solution.xlsx"},
        )

    except HTTPException as e:
        cleanup_files(consensus_path, public_company_path, output_path)
        raise e
    except Exception as e:
        cleanup_files(consensus_path, public_company_path, output_path)
        raise HTTPException(status_code=500, detail=str(e))
