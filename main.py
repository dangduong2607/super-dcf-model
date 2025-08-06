from fastapi import FastAPI, UploadFile, File, Form, Request, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import openpyxl
from openpyxl import load_workbook
import tempfile
import os
from datetime import datetime
import shutil

app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def process_dcf_model(consensus_file, profile_file=None):
    # Create a temporary directory to work in
    with tempfile.TemporaryDirectory() as temp_dir:
        # Load the template
        template_path = os.path.join(os.path.dirname(__file__), "Template.xlsx")
        template_wb = load_workbook(template_path)
        
        # Load the consensus data
        consensus_wb = load_workbook(consensus_file)
        
        # Copy data from Consensus sheet to template
        if 'Consensus' in consensus_wb.sheetnames:
            consensus_sheet = consensus_wb['Consensus']
            template_consensus_sheet = template_wb['Consensus']
            
            # Copy all data from uploaded consensus to template
            for row in consensus_sheet.iter_rows(values_only=True):
                template_consensus_sheet.append(row)
        
        # If profile file is provided, load and process it
        if profile_file:
            profile_wb = load_workbook(profile_file)
            if 'Public Company' in profile_wb.sheetnames:
                profile_sheet = profile_wb['Public Company']
                template_profile_sheet = template_wb['Public Company']
                
                # Copy all data from uploaded profile to template
                for row in profile_sheet.iter_rows(values_only=True):
                    template_profile_sheet.append(row)
        
        # Save the modified template to a temporary file
        output_path = os.path.join(temp_dir, "DCF_Model.xlsx")
        template_wb.save(output_path)
        
        return output_path

@app.post("/upload")
async def upload_files(consensus: UploadFile = File(...), profile: UploadFile = File(None)):
    try:
        # Save uploaded files temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as consensus_temp:
            shutil.copyfileobj(consensus.file, consensus_temp)
            consensus_temp_path = consensus_temp.name
        
        profile_temp_path = None
        if profile:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as profile_temp:
                shutil.copyfileobj(profile.file, profile_temp)
                profile_temp_path = profile_temp.name
        
        # Process the files
        output_path = process_dcf_model(consensus_temp_path, profile_temp_path)
        
        # Clean up temporary files
        os.unlink(consensus_temp_path)
        if profile_temp_path:
            os.unlink(profile_temp_path)
        
        # Return the generated file
        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="DCF_Model.xlsx"
        )
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
