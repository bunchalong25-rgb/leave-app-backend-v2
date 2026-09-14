import os
import json
import uuid
from datetime import datetime
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional, List, Dict
from excel_processor import create_sample_template, process_excel_template

app = FastAPI(title="Staff Leave Management API")

# Enable CORS for Netlify and all origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_FILE = os.path.join(os.path.dirname(__file__), "db.json")
UPLOADS_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOADS_DIR, exist_ok=True)
SAMPLE_TEMPLATE_PATH = os.path.join(UPLOADS_DIR, "sample_template.xlsx")

# Ensure sample template exists
if not os.path.exists(SAMPLE_TEMPLATE_PATH):
    create_sample_template(SAMPLE_TEMPLATE_PATH)

# Initial DB seed
INITIAL_DB = {
    "brands": [
        {"id": "1", "name": "Brand Alpha (สาขา สยาม)"},
        {"id": "2", "name": "Brand Beta (สาขา ชิดลม)"},
        {"id": "3", "name": "Brand Gamma (สาขา พารากอน)"},
        {"id": "4", "name": "Brand Delta (สาขา ไอคอนสยาม)"}
    ],
    "employees": [
        {"id": "emp_1", "name": "สมชาย สายดี", "brand": "Brand Alpha (สาขา สยาม)", "defaultRole": "staff"},
        {"id": "emp_2", "name": "วิภาวี มั่นคง", "brand": "Brand Alpha (สาขา สยาม)", "defaultRole": "staff"},
        {"id": "emp_3", "name": "อภิสิทธิ์ ขยันทำ", "brand": "Brand Beta (สาขา ชิดลม)", "defaultRole": "staff"},
        {"id": "emp_4", "name": "นภา เพลินตา", "brand": "Brand Beta (สาขา ชิดลม)", "defaultRole": "staff"},
        {"id": "emp_5", "name": "กิตติพงษ์ ยอดเยี่ยม", "brand": "Brand Delta (สาขา ไอคอนสยาม)", "defaultRole": "staff"},
        {"id": "emp_6", "name": "ดารินทร์ สุขใจ (หัวหน้า)", "brand": "Brand Alpha (สาขา สยาม)", "defaultRole": "admin"}
    ],
    "users": [
        {
            "id": "usr_admin",
            "lineUserId": "U_ADMIN_DEMO",
            "displayName": "ดารินทร์ สุขใจ (Admin)",
            "fullName": "ดารินทร์ สุขใจ (หัวหน้า)",
            "brand": "Brand Alpha (สาขา สยาม)",
            "role": "admin",
            "hasCompletedOnboarding": True,
            "createdAt": datetime.utcnow().isoformat()
        },
        {
            "id": "usr_staff1",
            "lineUserId": "U_STAFF_1",
            "displayName": "สมชาย (Staff)",
            "fullName": "สมชาย สายดี",
            "brand": "Brand Alpha (สาขา สยาม)",
            "role": "staff",
            "hasCompletedOnboarding": True,
            "createdAt": datetime.utcnow().isoformat()
        }
    ],
    "cutoffs": [
        {
            "id": "cut_2026_09_1",
            "yearMonth": "2026-09",
            "period": "1-15",
            "cutoffDatetime": "2026-09-25T23:59:59.000Z",
            "updatedAt": datetime.utcnow().isoformat()
        },
        {
            "id": "cut_2026_09_2",
            "yearMonth": "2026-09",
            "period": "16-end",
            "cutoffDatetime": "2026-09-30T23:59:59.000Z",
            "updatedAt": datetime.utcnow().isoformat()
        }
    ],
    "leaveRecords": [
        {
            "id": "lr_1",
            "userId": "usr_staff1",
            "userName": "สมชาย สายดี",
            "brand": "Brand Alpha (สาขา สยาม)",
            "yearMonth": "2026-09",
            "period": "1-15",
            "date": "2026-09-01",
            "dayNumber": 1,
            "code": "W",
            "updatedAt": datetime.utcnow().isoformat()
        },
        {
            "id": "lr_2",
            "userId": "usr_staff1",
            "userName": "สมชาย สายดี",
            "brand": "Brand Alpha (สาขา สยาม)",
            "yearMonth": "2026-09",
            "period": "1-15",
            "date": "2026-09-02",
            "dayNumber": 2,
            "code": "C",
            "updatedAt": datetime.utcnow().isoformat()
        }
    ]
}

def read_db():
    if not os.path.exists(DB_FILE):
        write_db(INITIAL_DB)
        return INITIAL_DB
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        write_db(INITIAL_DB)
        return INITIAL_DB

def write_db(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# Models
class OnboardingRequest(BaseModel):
    lineUserId: Optional[str] = None
    displayName: Optional[str] = None
    fullName: str
    brand: str
    role: Optional[str] = "staff"
    inviteCode: Optional[str] = None

class MockLoginRequest(BaseModel):
    role: str
    lineUserId: Optional[str] = None

class CutoffRequest(BaseModel):
    yearMonth: str
    period: str
    cutoffDatetime: str

class LeaveSubmitRequest(BaseModel):
    userId: str
    userName: str
    brand: str
    yearMonth: str
    period: str
    dateCodeMap: Dict[str, Optional[str]]

class ExcelProcessRequest(BaseModel):
    yearMonth: str
    period: Optional[str] = "1-15"

class BrandRequest(BaseModel):
    name: str

class EmployeeRequest(BaseModel):
    name: str
    brand: str
    defaultRole: Optional[str] = "staff"

# Routes
@app.get("/")
def read_root():
    return {"message": "Staff Leave Management API is online", "status": "ok"}

@app.get("/api/brands")
def get_brands():
    db = read_db()
    return {"brands": db.get("brands", [])}

@app.post("/api/brands")
def add_brand(req: BrandRequest):
    db = read_db()
    new_b = {"id": f"b_{int(datetime.utcnow().timestamp())}", "name": req.name}
    db["brands"].append(new_b)
    write_db(db)
    return {"success": True, "brand": new_b}

@app.get("/api/employees")
def get_employees():
    db = read_db()
    return {"employees": db.get("employees", [])}

@app.post("/api/employees")
def add_employee(req: EmployeeRequest):
    db = read_db()
    new_emp = {
        "id": f"emp_{int(datetime.utcnow().timestamp())}",
        "name": req.name,
        "brand": req.brand,
        "defaultRole": req.defaultRole or "staff"
    }
    db["employees"].append(new_emp)
    write_db(db)
    return {"success": True, "employee": new_emp}

@app.delete("/api/employees/{emp_id}")
def delete_employee(emp_id: str):
    db = read_db()
    db["employees"] = [e for e in db.get("employees", []) if e["id"] != emp_id]
    write_db(db)
    return {"success": True}

@app.post("/api/auth/register-onboarding")
def register_onboarding(req: OnboardingRequest):
    if req.role == "admin":
        invite_code_env = os.getenv("ADMIN_INVITE_CODE", "ADMIN2026")
        if not req.inviteCode or req.inviteCode.strip().upper() != invite_code_env:
            raise HTTPException(status_code=403, detail="รหัสผ่านพิเศษ (Invite Code) สำหรับหัวหน้าแผนกไม่ถูกต้อง")
    
    db = read_db()
    user = {
        "id": f"usr_{int(datetime.utcnow().timestamp())}",
        "lineUserId": req.lineUserId or f"LINE_{uuid.uuid4().hex[:8]}",
        "displayName": req.displayName or req.fullName,
        "fullName": req.fullName,
        "brand": req.brand,
        "role": req.role or "staff",
        "hasCompletedOnboarding": True,
        "createdAt": datetime.utcnow().isoformat()
    }
    db["users"].append(user)
    write_db(db)
    return {"success": True, "user": user}

@app.post("/api/auth/login-mock")
def login_mock(req: MockLoginRequest):
    db = read_db()
    users = db.get("users", [])
    matched = next((u for u in users if u.get("role") == req.role), None)
    if not matched:
        matched = users[0] if users else None
    return {"success": True, "user": matched}

@app.get("/api/cutoffs")
def get_cutoffs():
    db = read_db()
    return {"cutoffs": db.get("cutoffs", [])}

@app.post("/api/cutoffs")
def set_cutoff(req: CutoffRequest):
    db = read_db()
    cutoffs = db.get("cutoffs", [])
    found = False
    for c in cutoffs:
        if c.get("yearMonth") == req.yearMonth and c.get("period") == req.period:
            c["cutoffDatetime"] = req.cutoffDatetime
            c["updatedAt"] = datetime.utcnow().isoformat()
            found = True
            break
    if not found:
        cutoffs.append({
            "id": f"cut_{int(datetime.utcnow().timestamp())}",
            "yearMonth": req.yearMonth,
            "period": req.period,
            "cutoffDatetime": req.cutoffDatetime,
            "updatedAt": datetime.utcnow().isoformat()
        })
    db["cutoffs"] = cutoffs
    write_db(db)
    return {"success": True, "cutoffs": cutoffs}

@app.get("/api/leaves")
def get_leaves(yearMonth: str = "2026-09", period: Optional[str] = None):
    db = read_db()
    records = db.get("leaveRecords", [])
    filtered = [r for r in records if r.get("yearMonth") == yearMonth and (not period or r.get("period") == period)]
    return {"records": filtered}

@app.post("/api/leaves")
def save_leaves(req: LeaveSubmitRequest):
    db = read_db()
    cutoffs = db.get("cutoffs", [])
    current_cutoff = next((c for c in cutoffs if c.get("yearMonth") == req.yearMonth and c.get("period") == req.period), None)

    if current_cutoff:
        cutoff_dt = datetime.fromisoformat(current_cutoff["cutoffDatetime"].replace("Z", "+00:00"))
        now_dt = datetime.utcnow()
        if now_dt > cutoff_dt.replace(tzinfo=None):
            raise HTTPException(status_code=403, detail="ระบบปิดรับการบันทึกข้อมูลสำหรับรอบนี้แล้ว ไม่สามารถแก้ไขได้")

    # Remove existing for user/period
    records = [
        r for r in db.get("leaveRecords", [])
        if not (r.get("userId") == req.userId and r.get("yearMonth") == req.yearMonth and r.get("period") == req.period)
    ]

    for date_str, code in req.dateCodeMap.items():
        if code:
            day_num = int(date_str.split("-")[2])
            records.append({
                "id": f"lr_{int(datetime.utcnow().timestamp())}_{uuid.uuid4().hex[:4]}",
                "userId": req.userId,
                "userName": req.userName,
                "brand": req.brand,
                "yearMonth": req.yearMonth,
                "period": req.period,
                "date": date_str,
                "dayNumber": day_num,
                "code": code,
                "updatedAt": datetime.utcnow().isoformat()
            })

    db["leaveRecords"] = records
    write_db(db)
    return {"success": True, "message": "บันทึกตารางวันหยุดสำเร็จ"}

active_template_file = SAMPLE_TEMPLATE_PATH

@app.post("/api/excel/upload-template")
async def upload_template(templateFile: UploadFile = File(...)):
    global active_template_file
    file_path = os.path.join(UPLOADS_DIR, f"custom_{templateFile.filename}")
    content = await templateFile.read()
    with open(file_path, "wb") as f:
        f.write(content)
    active_template_file = file_path
    return {"success": True, "message": "อัปโหลดไฟล์ Excel Template สำเร็จ", "filename": templateFile.filename}

@app.post("/api/excel/process")
def process_excel(req: ExcelProcessRequest):
    global active_template_file
    db = read_db()
    records = [
        r for r in db.get("leaveRecords", [])
        if r.get("yearMonth") == req.yearMonth and (not req.period or r.get("period") == req.period)
    ]

    try:
        excel_bytes, updated_count = process_excel_template(
            template_path=active_template_file,
            leave_records=records,
            period=req.period
        )

        filename = f"Leave_Schedule_{req.yearMonth}_{req.period or 'All'}.xlsx"
        headers = {
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Updated-Cells-Count": str(updated_count),
            "Access-Control-Expose-Headers": "Content-Disposition, X-Updated-Cells-Count"
        }

        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
