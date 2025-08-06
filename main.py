from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
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

def cleanup_files(*file_paths):
    """Clean up temporary files."""
    for path in file_paths:
        if path and os.path.exists(path):
            os.remove(path)

def generate_output_file(consensus_path, profile_path):
    """Generates a new Excel file from the uploaded data and a template."""
    try:
        # Load the template
        output_wb = load_workbook("Template.xlsx")

        # Your logic to process data from consensus_path and profile_path and
        # populate the output_wb would go here.
        # For this example, we will assume you only need to remove other sheets.

        # Remove all sheets except DCF Model
        for sheet_name in list(output_wb.sheetnames):
            if sheet_name != "DCF Model":
                output_wb.remove(output_wb[sheet_name])
        
        # Save to a temporary file
        temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
        output_wb.save(temp_file.name)
        temp_file.close()
        return temp_file.name
        
    except Exception as e:
        raise Exception(f"Error generating output file: {str(e)}")

@app.post("/upload")
async def upload(background_tasks: BackgroundTasks, consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    consensus_path = None
    profile_path = None
    output_path = None

    try:
        # Save uploaded files temporarily
        consensus_path = "temp_consensus.xlsx"
        with open(consensus_path, "wb") as f:
            shutil.copyfileobj(consensus.file, f)

        if profile:
            profile_path = "temp_profile.xlsx"
            with open(profile_path, "wb") as f:
                shutil.copyfileobj(profile.file, f)

        # Generate the output file
        output_path = generate_output_file(consensus_path, profile_path)

        # Add cleanup to run in the background after the response is sent
        background_tasks.add_task(cleanup_files, consensus_path, profile_path, output_path)

        # Return the generated file using a generator to stream its content
        def file_iterator():
            with open(output_path, "rb") as f:
                yield from f

        return StreamingResponse(
            file_iterator(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=DCF_Model.xlsx"},
        )

    except Exception as e:
        cleanup_files(consensus_path, profile_path, output_path)
        raise HTTPException(status_code=500, detail=str(e))
