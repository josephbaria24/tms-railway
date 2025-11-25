# main.py (Railway FastAPI Service)
from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Optional
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
import requests
from io import BytesIO
import os

app = FastAPI(title="Excel Export Service")
                    
# CORS middleware - allow your Vercel domain
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with your Vercel domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class Trainee(BaseModel):
    id: str
    first_name: str
    last_name: str
    middle_initial: Optional[str] = None
    suffix: Optional[str] = None
    gender: Optional[str] = None
    age: Optional[int] = None
    company_name: Optional[str] = None
    company_position: Optional[str] = None
    company_city: Optional[str] = None
    company_region: Optional[str] = None
    company_industry: Optional[str] = None
    total_workers: Optional[int] = None
    company_email: Optional[str] = None
    email: Optional[str] = None
    phone_number: Optional[str] = None
    company_landline: Optional[str] = None
    picture_2x2_url: Optional[str] = None
    schedule_id: str
    certificate_number: Optional[str] = None

class ExportRequest(BaseModel):
    trainees: List[Trainee]
    courseName: str
    trainingDates: str
    scheduleId: str
    proxyUrl: Optional[str] = None

def get_template_filename(trainee_count: int) -> str:
    """Determine which template to use based on trainee count"""
    if trainee_count > 250:
        return 'Directory-300.xlsx'
    elif trainee_count > 200:
        return 'Directory-250.xlsx'
    elif trainee_count > 150:
        return 'Directory-200.xlsx'
    elif trainee_count > 100:
        return 'Directory-150.xlsx'
    elif trainee_count > 50:
        return 'Directory-100.xlsx'
    elif trainee_count > 30:
        return 'Directory-0.xlsx'
    elif trainee_count > 25:
        return 'Directory-1.xlsx'
    elif trainee_count > 20:
        return 'Directory-2.xlsx'
    elif trainee_count > 15:
        return 'Directory-3.xlsx'
    elif trainee_count > 10:
        return 'Directory-4.xlsx'
    else:
        return 'Directory-5.xlsx'

@app.get("/")
def read_root() -> dict[str, str]:
    return {
        "service": "Excel Export Service",
        "status": "running",
        "version": "1.0.0"
    }

@app.get("/health")
def health_check() -> dict[str, str]:
    return {"status": "healthy"}

@app.post("/export-excel")
def export_excel(request: ExportRequest) -> StreamingResponse:
    try:
        trainees = request.trainees
        course_name = request.courseName
        training_dates = request.trainingDates
        schedule_id = request.scheduleId
        proxy_url = request.proxyUrl

        if not trainees:
            raise HTTPException(
                status_code=400,
                detail="At least one trainee is required to generate the Excel file.",
            )

        if not schedule_id or len(schedule_id) < 4:
            raise HTTPException(
                status_code=400,
                detail="scheduleId is required and must be at least 4 characters long.",
            )

        trainee_count = len(trainees)
        
        # Get template filename
        template_filename = get_template_filename(trainee_count)
        template_path = f"templates/{template_filename}"
        
        print(f"Loading template: {template_path} for {trainee_count} trainees")
        
        # Check if template exists
        if not os.path.exists(template_path):
            raise HTTPException(
                status_code=404,
                detail=f"Template not found: {template_filename}"
            )
        
        # Load workbook
        wb = load_workbook(template_path)
        ws = wb['Directory of Participants']
        
        # Fill header information
        ws['C10'] = course_name
        ws['C11'] = training_dates
        
        # Count gender
        male_count = sum(1 for t in trainees if t.gender and t.gender.lower() == 'male')
        female_count = sum(1 for t in trainees if t.gender and t.gender.lower() == 'female')
        
        # Participant rows start from row 15
        start_row = 15
        
        for i, trainee in enumerate(trainees):
            row_num = start_row + i
            
            # Fill data
            ws[f'A{row_num}'] = i + 1
            # ws[f'B{row_num}'] = trainee.certificate_number or ''  # Commented out
            ws[f'C{row_num}'] = (trainee.last_name or '').upper()
            ws[f'D{row_num}'] = (trainee.first_name or '').upper()
            
            middle_initial = trainee.middle_initial or ''
            ws[f'E{row_num}'] = middle_initial[0].upper() if middle_initial else ''
            
            ws[f'F{row_num}'] = (trainee.suffix or '').upper()
            ws[f'G{row_num}'] = trainee.gender or ''
            ws[f'H{row_num}'] = trainee.age or ''
            ws[f'I{row_num}'] = trainee.company_name or ''
            ws[f'J{row_num}'] = trainee.company_position or ''
            ws[f'K{row_num}'] = trainee.company_city or ''
            ws[f'L{row_num}'] = trainee.company_region or ''
            ws[f'M{row_num}'] = trainee.company_industry or ''
            ws[f'N{row_num}'] = trainee.total_workers or ''
            ws[f'O{row_num}'] = trainee.company_email or ''
            ws[f'P{row_num}'] = trainee.email or ''
            ws[f'Q{row_num}'] = trainee.phone_number or ''
            ws[f'R{row_num}'] = trainee.company_landline or ''
            ws[f'T{row_num}'] = 'Online Training'
            ws[f'U{row_num}'] = f"#{trainee.schedule_id[-4:]}"
            
            # Add image if picture URL exists
            if trainee.picture_2x2_url:
                try:
                    # Use proxy URL if provided
                    if proxy_url:
                        img_url = f"{proxy_url}?url={trainee.picture_2x2_url}"
                    else:
                        img_url = trainee.picture_2x2_url
                    
                    print(f"Fetching image {i+1}: {img_url}")
                    
                    response = requests.get(img_url, timeout=10)
                    if response.status_code == 200:
                        img = OpenpyxlImage(BytesIO(response.content))
                        
                        # Set image size to 96x96 pixels
                        img.width = 96
                        img.height = 96
                        
                        # Add image to cell S
                        ws.add_image(img, f'S{row_num}')
                        print(f"✓ Image {i+1} added successfully")
                    else:
                        print(f"✗ Failed to fetch image {i+1}: HTTP {response.status_code}")
                except Exception as e:
                    print(f"✗ Error adding image {i+1}: {str(e)}")
        
        # Update Training Database sheet if it exists
        try:
            db_ws = wb['Training Database']
            db_ws['A11'] = 1
            db_ws['B11'] = course_name
            db_ws['C11'] = training_dates
            db_ws['D11'] = f"#{schedule_id[-4:]}"
            db_ws['E11'] = trainee_count
            db_ws['F11'] = male_count
            db_ws['G11'] = female_count
            db_ws['M11'] = 'Online Training'
        except Exception as e:
            print(f"Training Database sheet not found: {str(e)}")
        
        # Save to BytesIO
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        print(f"✓ Excel file generated successfully")
        
        # Return file as streaming response
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=Originals{schedule_id[-4:]}.xlsx"
            }
        )
    
    except Exception as e:
        print(f"Error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)