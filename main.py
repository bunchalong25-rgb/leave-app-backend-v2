import os
import json
import uuid
import io
import urllib.parse
from datetime import datetime
from typing import Optional, List, Dict
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Response, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse
from pydantic import BaseModel
import httpx
import pymongo
from excel_processor import create_sample_template, process_excel_template, process_master_data_excel

app = FastAPI(title="Staff Leave Management API")

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

if not os.path.exists(SAMPLE_TEMPLATE_PATH):
    create_sample_template(SAMPLE_TEMPLATE_PATH)

INITIAL_DB = {
    "brands": [],
    "departments": [],
    "employees": [],
    "users": [],
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
    "leaveRecords": []
}

MONGO_URI = os.getenv("MONGO_URI")
if MONGO_URI:
    mongo_client = pymongo.MongoClient(MONGO_URI)
    mongo_db = mongo_client["leave_app_db"]
    mongo_collection = mongo_db["main_data"]
    mongo_reset_col = mongo_db["reset_requests"]
else:
    mongo_client = None
    mongo_db = None
    mongo_collection = None
    mongo_reset_col = None

def read_db():
    if mongo_client:
        try:
            doc = mongo_collection.find_one({"_id": "main_db"})
            if not doc:
                mongo_collection.insert_one({"_id": "main_db", **INITIAL_DB})
                return INITIAL_DB
            return doc
        except Exception as e:
            print("MongoDB Read Error:", e)
            return INITIAL_DB
            
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
    if mongo_client:
        try:
            data_to_save = {k: v for k, v in data.items() if k != '_id'}
            mongo_collection.update_one(
                {"_id": "main_db"}, 
                {"$set": data_to_save}, 
                upsert=True
            )
        except Exception as e:
            print("MongoDB Write Error:", e)
    else:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

# Request Models
class LineLoginRequest(BaseModel):
    lineUserId: str
    displayName: Optional[str] = None

class OnboardingRequest(BaseModel):
    lineUserId: str
    displayName: Optional[str] = None
    fullName: str
    brand: str
    department: Optional[str] = None
    role: str # 'staff' or 'admin'
    supervisorId: Optional[str] = None
    supervisorName: Optional[str] = None
    inviteCode: Optional[str] = None

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
    department: Optional[str] = None

class BrandRequest(BaseModel):
    name: str

class EmployeeRequest(BaseModel):
    name: str
    brand: str

class ResetRequestModel(BaseModel):
    line_user_id: Optional[str] = None
    lineUserId: Optional[str] = None
    name: Optional[str] = None
    fullName: Optional[str] = None
    brand: Optional[str] = None
    department: Optional[str] = None
    userId: Optional[str] = None
    reason: str

class ApproveResetModel(BaseModel):
    request_id: Optional[str] = None
    requestId: Optional[str] = None
    line_user_id: Optional[str] = None
    lineUserId: Optional[str] = None
    userId: Optional[str] = None

class RejectResetModel(BaseModel):
    requestId: str
    reason: Optional[str] = None

# Base Configs
FRONTEND_URL = os.getenv("FRONTEND_URL", "https://creative-macaron-98004d.netlify.app")
LINE_CHANNEL_ID = os.getenv("LINE_CHANNEL_ID", "")
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET", "")
LINE_CALLBACK_URL = os.getenv("LINE_CALLBACK_URL", "https://leave-app-backend-v2.onrender.com/api/auth/line/callback")

# Routes
@app.get("/")
def read_root():
    return {"message": "Staff Leave Management API is online", "status": "ok"}

# ---------------------------------------------------------
# REAL LINE LOGIN OAUTH 2.0 FLOW
# ---------------------------------------------------------

@app.get("/api/auth/line/login")
def line_oauth_login():
    channel_id = LINE_CHANNEL_ID
    if not channel_id:
        # Fallback if channel ID env variable not yet configured
        return RedirectResponse(url=f"{FRONTEND_URL}/?error=LINE_CHANNEL_ID_NOT_CONFIGURED")

    state = uuid.uuid4().hex[:12]
    params = {
        "response_type": "code",
        "client_id": channel_id,
        "redirect_uri": LINE_CALLBACK_URL,
        "state": state,
        "scope": "profile openid"
    }

    line_authorize_url = f"https://access.line.me/oauth2/v2.1/authorize?{urllib.parse.urlencode(params)}"
    return RedirectResponse(url=line_authorize_url)


@app.get("/api/auth/line/callback")
async def line_oauth_callback(code: Optional[str] = None, state: Optional[str] = None, error: Optional[str] = None):
    if error or not code:
        return RedirectResponse(url=f"{FRONTEND_URL}/?error={error or 'NO_CODE'}")

    try:
        # 1. Exchange Auth Code for Access Token
        token_url = "https://api.line.me/oauth2/v2.1/token"
        token_data = {
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": LINE_CALLBACK_URL,
            "client_id": LINE_CHANNEL_ID,
            "client_secret": LINE_CHANNEL_SECRET,
        }

        async with httpx.AsyncClient() as client:
            token_res = await client.post(
                token_url,
                data=token_data,
                headers={"Content-Type": "application/x-www-form-urlencoded"}
            )

            if token_res.status_code != 200:
                print("LINE Token exchange failed:", token_res.text)
                return RedirectResponse(url=f"{FRONTEND_URL}/?error=TOKEN_EXCHANGE_FAILED")

            tokens = token_res.json()
            access_token = tokens.get("access_token")

            # 2. Get User Profile from LINE API
            profile_res = await client.get(
                "https://api.line.me/v2/profile",
                headers={"Authorization": f"Bearer {access_token}"}
            )

            if profile_res.status_code != 200:
                print("LINE Profile fetch failed:", profile_res.text)
                return RedirectResponse(url=f"{FRONTEND_URL}/?error=PROFILE_FETCH_FAILED")

            profile = profile_res.json()
            line_user_id = profile.get("userId")
            display_name = profile.get("displayName", "")
            picture_url = profile.get("pictureUrl", "")

            # 3. Check DB if User exists
            db = read_db()
            users = db.get("users", [])
            existing_user = next((u for u in users if u.get("lineUserId") == line_user_id), None)

            # 4. Redirect back to Frontend Netlify with User info Query Params
            if existing_user:
                user_json = urllib.parse.quote(json.dumps(existing_user))
                return RedirectResponse(url=f"{FRONTEND_URL}/?exists=true&user={user_json}")
            else:
                encoded_name = urllib.parse.quote(display_name)
                encoded_pic = urllib.parse.quote(picture_url)
                return RedirectResponse(
                    url=f"{FRONTEND_URL}/?exists=false&lineUserId={line_user_id}&displayName={encoded_name}&pictureUrl={encoded_pic}"
                )

    except Exception as e:
        print("Error in LINE OAuth Callback:", str(e))
        return RedirectResponse(url=f"{FRONTEND_URL}/?error=INTERNAL_CALLBACK_ERROR")


# ---------------------------------------------------------
# AUTHENTICATION & ONBOARDING API
# ---------------------------------------------------------

@app.post("/api/auth/login-line")
def login_line(req: LineLoginRequest):
    db = read_db()
    users = db.get("users", [])
    matched = next((u for u in users if u.get("lineUserId") == req.lineUserId), None)
    if matched:
        return {"exists": True, "user": matched}
    else:
        return {"exists": False, "lineUserId": req.lineUserId, "displayName": req.displayName}

@app.get("/api/onboarding/options")
def get_onboarding_options():
    db = read_db()
    users = db.get("users", [])
    supervisors = [
        {"id": u.get("id"), "fullName": u.get("fullName") or u.get("displayName")}
        for u in users
        if u.get("role") == "admin" and "ดารินทร์ สุขใจ" not in (u.get("fullName") or "")
    ]
    brands = db.get("brands", [])
    employees = db.get("employees", [])
    
    # ดึงรายชื่อแผนกจริงที่ไม่ซ้ำกันจากฐานข้อมูลและพนักงาน
    dept_set = set()
    for e in employees:
        d = e.get("department")
        if d and isinstance(d, str) and d.strip():
            dept_set.add(d.strip())
    for d in db.get("departments", []):
        if d and isinstance(d, str) and d.strip():
            dept_set.add(d.strip())
    departments = sorted(list(dept_set))

    claimed_names = set(u.get("fullName") for u in users if u.get("role") == "staff")
    available_employees = [e for e in employees if e.get("name") not in claimed_names]

    return {
        "supervisors": supervisors,
        "brands": brands,
        "departments": departments,
        "masterEmployees": available_employees,
        "employees": available_employees,
        "hasMasterData": len(employees) > 0
    }

@app.post("/api/auth/register-onboarding")
def register_onboarding(req: OnboardingRequest):
    if req.role == "admin":
        invite_code_env = os.getenv("ADMIN_INVITE_CODE", "ADMIN2026")
        if not req.inviteCode or req.inviteCode.strip().upper() != invite_code_env:
            raise HTTPException(status_code=403, detail="รหัสผ่านพิเศษ (Invite Code) สำหรับหัวหน้าแผนกไม่ถูกต้อง")

    db = read_db()
    user = {
        "id": f"usr_{int(datetime.utcnow().timestamp())}",
        "lineUserId": req.lineUserId,
        "displayName": req.displayName or req.fullName,
        "fullName": req.fullName,
        "department": req.department,
        "brand": req.brand,
        "role": req.role,
        "supervisorId": req.supervisorId,
        "supervisorName": req.supervisorName,
        "hasCompletedOnboarding": True,
        "createdAt": datetime.utcnow().isoformat()
    }
    db["users"].append(user)
    write_db(db)
    return {"success": True, "user": user}

# ---------------------------------------------------------
# ACCOUNT RESET REQUESTS API (FastAPI + MongoDB)
# ---------------------------------------------------------

@app.post("/api/request-reset")
def request_reset(req: ResetRequestModel):
    """
    POST /api/request-reset
    รับข้อมูล: line_user_id (หรือ lineUserId), name (หรือ fullName), brand, reason
    บันทึกคำขอลงใน collection reset_requests (สถานะ pending)
    """
    line_id = req.line_user_id or req.lineUserId
    name = req.name or req.fullName or "พนักงาน"
    brand = req.brand or "-"
    reason = req.reason.strip() if req.reason else ""

    if not line_id:
        raise HTTPException(status_code=400, detail="กรุณาระบุ line_user_id หรือ lineUserId")
    if not reason:
        raise HTTPException(status_code=400, detail="กรุณาระบุเหตุผลในการขอรีเซ็ตบัญชี")

    now_iso = datetime.utcnow().isoformat()
    req_id = f"rst_{int(datetime.utcnow().timestamp())}_{uuid.uuid4().hex[:6]}"

    doc = {
        "_id": req_id,
        "id": req_id,
        "line_user_id": line_id,
        "lineUserId": line_id,
        "name": name,
        "fullName": name,
        "brand": brand,
        "department": req.department or "-",
        "userId": req.userId,
        "reason": reason,
        "status": "pending",
        "createdAt": now_iso,
        "updatedAt": now_iso
    }

    # 1. บันทึกลงใน MongoDB collection reset_requests (ถ้าเชื่อมต่อ MongoDB อยู่)
    if mongo_client and mongo_reset_col is not None:
        try:
            existing = mongo_reset_col.find_one({
                "$or": [{"line_user_id": line_id}, {"lineUserId": line_id}],
                "status": "pending"
            })
            if existing:
                mongo_reset_col.update_one(
                    {"_id": existing["_id"]},
                    {"$set": {
                        "name": name,
                        "fullName": name,
                        "brand": brand,
                        "reason": reason,
                        "updatedAt": now_iso
                    }}
                )
                doc = {**existing, "name": name, "fullName": name, "brand": brand, "reason": reason, "updatedAt": now_iso}
                doc["_id"] = str(doc["_id"])
            else:
                mongo_reset_col.insert_one(doc)
                doc["_id"] = str(doc["_id"])
        except Exception as e:
            print("MongoDB insert reset_requests error:", e)

    # 2. ซิงค์กับ JSON DB เผื่อกรณีรันแบบ Local / Offline
    db = read_db()
    if "resetRequests" not in db:
        db["resetRequests"] = []

    existing_local = next(
        (r for r in db["resetRequests"] if r.get("status") == "pending" and (
            r.get("line_user_id") == line_id or r.get("lineUserId") == line_id or (req.userId and r.get("userId") == req.userId)
        )),
        None
    )
    if existing_local:
        existing_local["name"] = name
        existing_local["fullName"] = name
        existing_local["brand"] = brand
        existing_local["reason"] = reason
        existing_local["updatedAt"] = now_iso
        doc = existing_local
    else:
        db["resetRequests"].append(doc)

    write_db(db)

    return {
        "success": True,
        "message": "ส่งคำขอรีเซ็ตบัญชีเรียบร้อยแล้ว (สถานะ pending) กรุณารอหัวหน้างานอนุมัติ",
        "request": doc
    }

@app.get("/api/reset-requests")
def get_pending_reset_requests():
    """
    GET /api/reset-requests
    ดึงรายการคำขอรีเซ็ตทั้งหมดที่สถานะเป็น pending ส่งกลับไปให้หน้า Admin Dashboard
    """
    if mongo_client and mongo_reset_col is not None:
        try:
            cursor = mongo_reset_col.find({"status": "pending"}).sort("createdAt", -1)
            requests = []
            for doc in cursor:
                doc["_id"] = str(doc["_id"])
                requests.append(doc)
            return {
                "success": True,
                "total": len(requests),
                "pendingCount": len(requests),
                "requests": requests
            }
        except Exception as e:
            print("MongoDB fetch reset_requests error:", e)

    # Fallback to local DB
    db = read_db()
    requests = db.get("resetRequests", [])
    pending = [r for r in requests if r.get("status") == "pending"]
    sorted_reqs = sorted(pending, key=lambda x: x.get("createdAt", ""), reverse=True)
    return {
        "success": True,
        "total": len(sorted_reqs),
        "pendingCount": len(sorted_reqs),
        "requests": sorted_reqs
    }

@app.get("/api/admin/reset-requests")
def get_admin_reset_requests():
    """
    GET /api/admin/reset-requests
    ดึงรายการคำขอทั้งหมดสำหรับหน้า Admin Dashboard
    """
    if mongo_client and mongo_reset_col is not None:
        try:
            cursor = mongo_reset_col.find({}).sort("createdAt", -1)
            requests = []
            for doc in cursor:
                doc["_id"] = str(doc["_id"])
                requests.append(doc)
            pending_count = sum(1 for r in requests if r.get("status") == "pending")
            return {
                "success": True,
                "total": len(requests),
                "pendingCount": pending_count,
                "requests": requests
            }
        except Exception as e:
            print("MongoDB fetch admin reset requests error:", e)

    db = read_db()
    requests = db.get("resetRequests", [])
    sorted_reqs = sorted(requests, key=lambda x: x.get("createdAt", ""), reverse=True)
    pending_count = sum(1 for r in sorted_reqs if r.get("status") == "pending")
    return {
        "success": True,
        "total": len(sorted_reqs),
        "pendingCount": pending_count,
        "requests": sorted_reqs
    }

@app.get("/api/reset-status")
def get_reset_status(lineUserId: Optional[str] = None, line_user_id: Optional[str] = None, userId: Optional[str] = None):
    line_id = line_user_id or lineUserId

    if mongo_client and mongo_reset_col is not None and line_id:
        try:
            matched = list(mongo_reset_col.find({
                "$or": [{"line_user_id": line_id}, {"lineUserId": line_id}]
            }).sort("createdAt", -1))
            if matched:
                latest = matched[0]
                latest["_id"] = str(latest["_id"])
                return {
                    "hasPending": latest.get("status") == "pending",
                    "status": latest.get("status"),
                    "request": latest
                }
        except Exception as e:
            print("MongoDB get_reset_status error:", e)

    db = read_db()
    requests = db.get("resetRequests", [])
    matched = [r for r in requests if (line_id and (r.get("lineUserId") == line_id or r.get("line_user_id") == line_id)) or (userId and r.get("userId") == userId)]
    if not matched:
        return {"hasPending": False, "status": "none", "request": None}

    latest = sorted(matched, key=lambda x: x.get("createdAt", ""), reverse=True)[0]
    return {
        "hasPending": latest.get("status") == "pending",
        "status": latest.get("status"),
        "request": latest
    }

@app.post("/api/approve-reset")
@app.post("/api/admin/approve-reset")
def approve_reset(req: ApproveResetModel):
    """
    POST /api/approve-reset
    รับค่า request_id หรือ line_user_id
    เมื่อหัวหน้ากดอนุมัติ ให้ลบหรือรีเซ็ตข้อมูลการลงทะเบียนของ LINE ID นั้นในตารางผู้ใช้
    อัปเดตสถานะคำขอใน reset_requests เป็น approved
    """
    req_id = req.request_id or req.requestId
    line_id = req.line_user_id or req.lineUserId
    user_id = req.userId

    if not req_id and not line_id and not user_id:
        raise HTTPException(status_code=400, detail="กรุณาระบุ request_id หรือ line_user_id")

    now_iso = datetime.utcnow().isoformat()
    emp_name = None

    # 1. ดำเนินการใน MongoDB (ถ้าเชื่อมต่ออยู่)
    if mongo_client and mongo_reset_col is not None:
        try:
            query = {}
            if req_id:
                query = {"$or": [{"_id": req_id}, {"id": req_id}]}
            elif line_id:
                query = {"$or": [{"line_user_id": line_id}, {"lineUserId": line_id}], "status": "pending"}

            target_doc = mongo_reset_col.find_one(query)
            if target_doc:
                line_id = target_doc.get("line_user_id") or target_doc.get("lineUserId") or line_id
                emp_name = target_doc.get("name") or target_doc.get("fullName")

                # อัปเดตสถานะคำขอใน reset_requests เป็น approved
                mongo_reset_col.update_one(
                    {"_id": target_doc["_id"]},
                    {"$set": {
                        "status": "approved",
                        "approvedAt": now_iso,
                        "updatedAt": now_iso
                    }}
                )

            # ลบหรือรีเซ็ตข้อมูลการลงทะเบียนของ LINE ID นั้นในตารางผู้ใช้
            if line_id:
                # ลบออกจาก main_data (collection หลัก)
                mongo_collection.update_one(
                    {"_id": "main_db"},
                    {"$pull": {
                        "users": {
                            "$or": [
                                {"lineUserId": line_id},
                                {"line_user_id": line_id}
                            ]
                        }
                    }}
                )
                # ลบออกจาก collection users (ถ้ามี)
                try:
                    mongo_db["users"].delete_many({
                        "$or": [
                            {"lineUserId": line_id},
                            {"line_user_id": line_id}
                        ]
                    })
                except Exception:
                    pass
        except Exception as e:
            print("MongoDB approve_reset error:", e)

    # 2. ซิงค์กับ JSON DB เผื่อกรณีรันแบบ Local / Offline
    db = read_db()
    users = db.get("users", [])
    for u in users:
        if (line_id and (u.get("lineUserId") == line_id or u.get("line_user_id") == line_id)) or (user_id and u.get("id") == user_id):
            if not emp_name:
                emp_name = u.get("fullName") or u.get("name")
            break

    # ล้างข้อมูลการลงทะเบียนของผู้ใช้เพื่อให้สามารถเข้าสู่ระบบและลงทะเบียนใหม่ได้
    db["users"] = [
        u for u in users
        if not (
            (line_id and (u.get("lineUserId") == line_id or u.get("line_user_id") == line_id)) or
            (user_id and u.get("id") == user_id)
        )
    ]

    # อัปเดตสถานะใน resetRequests
    requests = db.get("resetRequests", [])
    for r in requests:
        is_match = False
        if req_id and (r.get("id") == req_id or r.get("_id") == req_id):
            is_match = True
        elif line_id and (r.get("lineUserId") == line_id or r.get("line_user_id") == line_id) and r.get("status") == "pending":
            is_match = True
        elif user_id and r.get("userId") == user_id and r.get("status") == "pending":
            is_match = True

        if is_match:
            r["status"] = "approved"
            r["approvedAt"] = now_iso
            r["updatedAt"] = now_iso

    write_db(db)

    return {
        "success": True,
        "message": f"อนุมัติการรีเซ็ตบัญชีของ {emp_name or 'พนักงาน'} เรียบร้อยแล้ว (สถานะ approved) พนักงานสามารถลงทะเบียนใหม่ได้ทันที"
    }

@app.post("/api/admin/reject-reset")
def reject_reset(req: RejectResetModel):
    now_iso = datetime.utcnow().isoformat()
    if mongo_client and mongo_reset_col is not None:
        try:
            mongo_reset_col.update_one(
                {"$or": [{"_id": req.requestId}, {"id": req.requestId}]},
                {"$set": {
                    "status": "rejected",
                    "rejectReason": req.reason or "ไม่อนุมัติ",
                    "rejectedAt": now_iso,
                    "updatedAt": now_iso
                }}
            )
        except Exception as e:
            print("MongoDB reject_reset error:", e)

    db = read_db()
    requests = db.get("resetRequests", [])
    target = next((r for r in requests if r.get("id") == req.requestId or r.get("_id") == req.requestId), None)
    if not target:
        raise HTTPException(status_code=404, detail="ไม่พบรายการคำขอ")

    target["status"] = "rejected"
    target["rejectReason"] = req.reason or "ไม่อนุมัติ"
    target["rejectedAt"] = now_iso
    target["updatedAt"] = now_iso
    write_db(db)
    return {"success": True, "message": "ปฏิเสธคำขอรีเซ็ตเรียบร้อยแล้ว"}

# ---------------------------------------------------------
# ADMIN MASTER DATA UPLOAD & STATUS DASHBOARD
# ---------------------------------------------------------

@app.post("/api/admin/upload-master-data")
async def upload_master_data(
    file: UploadFile = File(...),
    minStaffThreshold: Optional[int] = Form(2)
):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="กรุณาอัปโหลดไฟล์ Excel (.xlsx)")

    content = await file.read()
    threshold = minStaffThreshold if minStaffThreshold is not None else 2
    res = process_master_data_excel(content, min_staff_threshold=threshold)

    db = read_db()
    existing_employees = db.get("employees", [])
    existing_brands = db.get("brands", [])

    added_emp_count = 0
    added_brand_count = 0

    for idx, emp in enumerate(res.get("employees", [])):
        emp_name = emp.get("name")
        brand_name = emp.get("brand")
        dept_name = (emp.get("department") or "").strip()
        requires_shift = emp.get("requires_shift_selection", False)

        if emp_name:
            found_emp = next((e for e in existing_employees if e.get("name") == emp_name), None)
            if not found_emp:
                existing_employees.append({
                    "id": f"emp_{int(datetime.utcnow().timestamp())}_{idx}",
                    "name": emp_name,
                    "brand": brand_name or "General",
                    "department": dept_name,
                    "requiresShiftSelection": requires_shift
                })
                added_emp_count += 1
            else:
                found_emp["brand"] = brand_name
                if dept_name:
                    found_emp["department"] = dept_name
                found_emp["requiresShiftSelection"] = requires_shift

        if brand_name:
            if not any(b.get("name") == brand_name for b in existing_brands):
                existing_brands.append({
                    "id": f"b_{int(datetime.utcnow().timestamp())}_{idx}",
                    "name": brand_name
                })
                added_brand_count += 1

    # ซิงค์รายชื่อแผนกทั้งหมดลงฐานข้อมูล (เรียงตามลำดับที่ตรวจพบ)
    dept_list = []
    for d in res.get("departments", []):
        if d and isinstance(d, str) and d.strip() and d.strip() not in dept_list:
            dept_list.append(d.strip())
    for emp_item in existing_employees:
        d = emp_item.get("department")
        if d and isinstance(d, str) and d.strip() and d.strip() not in dept_list:
            dept_list.append(d.strip())
    for d in db.get("departments", []):
        if d and isinstance(d, str) and d.strip() and d.strip() not in dept_list:
            dept_list.append(d.strip())
    db["departments"] = dept_list

    db["employees"] = existing_employees
    db["brands"] = existing_brands
    db["remarks"] = res.get("remarks", [])
    db["brandCounts"] = res.get("brand_counts", {})
    write_db(db)

    total_emp = res.get("total_employees", len(existing_employees))

    return {
        "success": True,
        "message": f"นำเข้าข้อมูลสำเร็จ: พบพนักงานทั้งหมด {total_emp} คน (เพิ่มใหม่ {added_emp_count} คน), เพิ่มแบรนด์ใหม่ {added_brand_count} แบรนด์, {len(db['departments'])} แผนก",
        "totalEmployees": total_emp,
        "addedEmployees": added_emp_count,
        "addedBrands": added_brand_count,
        "departments": db["departments"],
        "brandCounts": res.get("brand_counts", {}),
        "remarks": res.get("remarks", [])
    }

@app.get("/api/admin/submission-status")
def get_submission_status(
    yearMonth: str = "2026-09",
    period: str = "1-15",
    supervisorId: Optional[str] = None,
    department: Optional[str] = None
):
    db = read_db()
    users = db.get("users", [])
    leave_records = db.get("leaveRecords", [])

    staff_users = [u for u in users if u.get("role") == "staff"]
    if supervisorId:
        staff_users = [u for u in staff_users if u.get("supervisorId") == supervisorId]
    if department and department != "ALL":
        staff_users = [u for u in staff_users if (u.get("department") or "").strip() == department.strip()]

    submitted_user_ids = set(
        r.get("userId") for r in leave_records
        if r.get("yearMonth") == yearMonth and r.get("period") == period
    )

    submitted_list = []
    pending_list = []

    for u in staff_users:
        info = {
            "userId": u.get("id"),
            "fullName": u.get("fullName"),
            "brand": u.get("brand"),
            "department": u.get("department") or "-",
            "supervisorName": u.get("supervisorName", "-")
        }
        if u.get("id") in submitted_user_ids:
            submitted_list.append(info)
        else:
            pending_list.append(info)

    return {
        "totalStaff": len(staff_users),
        "submittedCount": len(submitted_list),
        "pendingCount": len(pending_list),
        "submittedList": submitted_list,
        "pendingList": pending_list
    }

# ---------------------------------------------------------
# CUTOFFS & LEAVE RECORDS
# ---------------------------------------------------------

@app.get("/api/departments")
def get_departments():
    db = read_db()
    dept_list = []
    for d in db.get("departments", []):
        if d and isinstance(d, str) and d.strip() and d.strip() not in dept_list:
            dept_list.append(d.strip())
    for e in db.get("employees", []):
        d = e.get("department")
        if d and isinstance(d, str) and d.strip() and d.strip() not in dept_list:
            dept_list.append(d.strip())
    return {"departments": dept_list}

@app.get("/api/brands")
def get_brands():
    db = read_db()
    return {"brands": db.get("brands", [])}

@app.get("/api/employees")
def get_employees():
    db = read_db()
    dept_list = []
    for d in db.get("departments", []):
        if d and isinstance(d, str) and d.strip() and d.strip() not in dept_list:
            dept_list.append(d.strip())
    for e in db.get("employees", []):
        d = e.get("department")
        if d and isinstance(d, str) and d.strip() and d.strip() not in dept_list:
            dept_list.append(d.strip())
    return {
        "employees": db.get("employees", []),
        "departments": dept_list
    }

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
    users = db.get("users", [])
    employees = db.get("employees", [])

    allowed_names = set()
    if req.department:
        for u in users:
            if u.get("department") == req.department:
                if u.get("fullName"):
                    allowed_names.add(u.get("fullName").strip())
        for e in employees:
            if e.get("department") == req.department:
                if e.get("name"):
                    allowed_names.add(e.get("name").strip())

    records = [
        r for r in db.get("leaveRecords", [])
        if r.get("yearMonth") == req.yearMonth 
        and (not req.period or r.get("period") == req.period)
        and (not req.department or (r.get("userName") and r.get("userName").strip() in allowed_names))
    ]

    try:
        excel_bytes, updated_count = process_excel_template(
            template_path=active_template_file,
            leave_records=records,
            period=req.period
        )

        dept_suffix = f"_{req.department}" if req.department else ""
        filename = f"Leave_Schedule_{req.yearMonth}_{req.period or 'All'}{dept_suffix}.xlsx"
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
