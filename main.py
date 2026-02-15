from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
import tempfile
from pathlib import Path
from datetime import datetime, date
import glob
from pydantic import BaseModel
from typing import Optional

from resume_generator import create_resume_from_file, preview_resume_content

# Supabase client
from supabase import create_client, Client

app = FastAPI(
    title="Professional Resume Generator API",
    description="Upload your Excel template and get a professional ATS-friendly resume",
    version="1.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://resumedj.com",
        "https://www.resumedj.com",
        "https://deweslibb.github.io",  # For test environment
        "http://localhost:8000"  # Keep for local testing
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Initialize Supabase
SUPABASE_URL = os.getenv('SUPABASE_URL', 'https://wfsszxypjzzpebjhujuu.supabase.co')
SUPABASE_SECRET_KEY = os.getenv('SUPABASE_SECRET_KEY')

supabase: Client = None
if SUPABASE_SECRET_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_SECRET_KEY)

TEMPLATE_PATH = Path("Resume_Creator.xlsx")

# Store file paths temporarily
file_storage = {}

# ==================== PYDANTIC MODELS ====================

class UsageCheck(BaseModel):
    user_id: str

class UsageRecord(BaseModel):
    user_id: str

class CheckoutSessionRequest(BaseModel):
    user_id: str
    email: str
    price_id: str

class StripeWebhookEvent(BaseModel):
    type: str
    data: dict

# ==================== HEALTH & INFO ENDPOINTS ====================

@app.get("/")
async def root():
    return {
        "message": "Professional Resume Generator API",
        "version": "1.0.0",
        "endpoints": {
            "GET /template": "Download the Excel template",
            "POST /generate": "Upload filled Excel and get resume",
            "GET /health": "Health check",
            "POST /check-usage": "Check user generation limits",
            "POST /record-usage": "Record a generation",
            "POST /create-checkout-session": "Create Stripe checkout session",
            "POST /webhook/stripe": "Handle Stripe webhook events"
        }
    }

@app.get("/health")
async def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

# ==================== TEMPLATE ENDPOINT ====================

@app.get("/template")
async def download_template():
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=404, detail="Template file not found")

    return FileResponse(
        path=str(TEMPLATE_PATH),
        filename="Resume_Creator.xlsx",
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={
            "Content-Disposition": 'attachment; filename="Resume_Creator.xlsx"',
            "X-Content-Type-Options": "nosniff"
        }
    )

# ==================== USAGE TRACKING ENDPOINTS ====================

@app.post("/check-usage")
async def check_usage(usage: UsageCheck):
    """Check if user can generate a resume"""
    if not supabase:
        # If Supabase not configured, allow (for testing)
        return {"can_generate": True, "reason": "unlimited"}
    
    try:
        # Get user usage
        response = supabase.table('user_usage').select('*').eq('user_id', usage.user_id).execute()
        
        if not response.data:
            # New user - can generate
            return {"can_generate": True, "reason": "new_user"}
        
        user_data = response.data[0]
        
        # Pro users have unlimited
        if user_data.get('subscription_status') == 'pro':
            return {"can_generate": True, "reason": "pro"}
        
        # Check daily limit
        if user_data.get('generations_today', 0) >= 1:
            return {
                "can_generate": False,
                "reason": "daily_limit",
                "message": "Daily limit reached (1 per day). Upgrade to Pro for unlimited!"
            }
        
        # Check monthly limit
        if user_data.get('generations_month', 0) >= 5:
            return {
                "can_generate": False,
                "reason": "monthly_limit",
                "message": "Monthly limit reached (5 per month). Upgrade to Pro for unlimited!"
            }
        
        return {"can_generate": True, "reason": "within_limits"}
        
    except Exception as e:
        print(f"Usage check error: {e}")
        # On error, allow (fail open)
        return {"can_generate": True, "reason": "error_fallback"}

@app.post("/record-usage")
async def record_usage(usage: UsageRecord):
    """Record a resume generation"""
    if not supabase:
        return {"success": True, "message": "Supabase not configured"}
    
    try:
        today = date.today().isoformat()
        
        # Get existing usage
        response = supabase.table('user_usage').select('*').eq('user_id', usage.user_id).execute()
        
        if response.data:
            # Update existing
            user_data = response.data[0]
            last_date = user_data.get('last_generation_date')
            is_new_day = last_date != today
            
            supabase.table('user_usage').update({
                'generations_today': 1 if is_new_day else user_data.get('generations_today', 0) + 1,
                'generations_month': user_data.get('generations_month', 0) + 1,
                'last_generation_date': today
            }).eq('user_id', usage.user_id).execute()
        else:
            # Create new
            supabase.table('user_usage').insert({
                'user_id': usage.user_id,
                'generations_today': 1,
                'generations_month': 1,
                'last_generation_date': today,
                'subscription_status': 'free'
            }).execute()
        
        return {"success": True}
        
    except Exception as e:
        print(f"Record usage error: {e}")
        return {"success": False, "error": str(e)}

# ==================== RESUME GENERATION ENDPOINTS ====================

@app.post("/generate")
async def generate_resume(file: UploadFile = File(...), format: str = "pdf"):
    """Generate resume and return download links"""
    if not file.filename.endswith(('.xlsx', '.xlsm')):
        raise HTTPException(status_code=400, detail="File must be Excel")

    temp_dir = tempfile.mkdtemp()

    try:
        temp_excel = Path(temp_dir) / "input.xlsx"
        with open(temp_excel, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        result = create_resume_from_file(
            excel_path=str(temp_excel),
            output_dir=temp_dir,
            generate_pdf=(format in ["pdf", "both"]),
            generate_docx=(format in ["docx", "both"])
        )

        if not result["success"]:
            raise HTTPException(status_code=500, detail=result.get('error'))

        # Store file paths with unique IDs
        file_id = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        file_storage[file_id] = {
            "pdf": result.get('pdf_path'),
            "docx": result.get('docx_path'),
            "name": result['name']
        }

        # Return file info for frontend to download
        response = {
            "success": True,
            "name": result['name'],
            "file_id": file_id,
            "files": {}
        }

        if result.get('pdf_path') and Path(result['pdf_path']).exists():
            response["files"]["pdf"] = f"/download/{file_id}/pdf"
        if result.get('docx_path') and Path(result['docx_path']).exists():
            response["files"]["docx"] = f"/download/{file_id}/docx"

        return JSONResponse(response)

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download/{file_id}/{file_type}")
async def download_file(file_id: str, file_type: str):
    """Download individual file with headers to prevent Windows blocking"""

    if file_id not in file_storage:
        raise HTTPException(status_code=404, detail="File not found or expired")

    files = file_storage[file_id]
    file_path = files.get(file_type)

    if not file_path or not Path(file_path).exists():
        raise HTTPException(status_code=404, detail=f"{file_type.upper()} file not available")

    filename = Path(file_path).name

    # Set proper media type
    if file_type == "pdf":
        media_type = 'application/pdf'
    else:
        media_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

    # Headers to prevent Windows SmartScreen blocking
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type=media_type,
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Content-Type-Options": "nosniff",
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "Expires": "0"
        }
    )

@app.post("/generate-preview")
async def generate_preview(file: UploadFile = File(...)):
    """Preview the uploaded Excel file without generating resume"""
    if not file.filename.endswith(('.xlsx', '.xlsm')):
        raise HTTPException(status_code=400, detail="File must be Excel")

    temp_dir = tempfile.mkdtemp()

    try:
        temp_excel = Path(temp_dir) / "input.xlsx"
        with open(temp_excel, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        preview_data = preview_resume_content(str(temp_excel))
        return JSONResponse(content=preview_data)

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error previewing resume: {str(e)}")

    finally:
        if Path(temp_excel).exists():
            os.remove(temp_excel)
        if Path(temp_dir).exists():
            os.rmdir(temp_dir)

# ==================== STRIPE INTEGRATION ====================

@app.post("/create-checkout-session")
async def create_checkout_session(request: CheckoutSessionRequest):
    """Create Stripe checkout session for Pro subscription"""
    try:
        import stripe

        # Get Stripe secret key from environment
        stripe.api_key = os.getenv('STRIPE_SECRET_KEY')

        if not stripe.api_key:
            raise HTTPException(status_code=500, detail="Stripe not configured")

        # Create checkout session
        checkout_session = stripe.checkout.Session.create(
            customer_email=request.email,
            client_reference_id=request.user_id,
            payment_method_types=['card'],
            line_items=[{
                'price': request.price_id,
                'quantity': 1,
            }],
            mode='subscription',
            success_url='https://www.resumedj.com/?checkout=success',
            cancel_url='https://www.resumedj.com/?checkout=cancelled',
            metadata={
                'user_id': request.user_id
            }
        )

        return JSONResponse({
            'sessionId': checkout_session.id
        })

    except ImportError:
        raise HTTPException(status_code=500, detail="Stripe library not installed")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/webhook/stripe")
async def stripe_webhook(request: Request):
    """Handle Stripe webhook events"""
    try:
        import stripe

        # Get environment variables
        stripe.api_key = os.getenv('STRIPE_SECRET_KEY')
        webhook_secret = os.getenv('STRIPE_WEBHOOK_SECRET')

        if not all([stripe.api_key, webhook_secret, supabase]):
            raise HTTPException(status_code=500, detail="Webhook not configured")

        # Get request body and signature
        payload = await request.body()
        sig_header = request.headers.get('stripe-signature')

        try:
            event = stripe.Webhook.construct_event(
                payload, sig_header, webhook_secret
            )
        except ValueError:
            raise HTTPException(status_code=400, detail="Invalid payload")
        except stripe.error.SignatureVerificationError:
            raise HTTPException(status_code=400, detail="Invalid signature")

        # Handle the event
        if event['type'] == 'checkout.session.completed':
            session = event['data']['object']
            user_id = session.get('metadata', {}).get('user_id') or session.get('client_reference_id')

            if user_id:
                # Update user to Pro
                supabase.table('user_usage').update({
                    'subscription_status': 'pro',
                    'stripe_customer_id': session.get('customer'),
                    'stripe_subscription_id': session.get('subscription')
                }).eq('user_id', user_id).execute()

        elif event['type'] == 'customer.subscription.deleted':
            subscription = event['data']['object']

            # Find user by subscription ID and downgrade to free
            result = supabase.table('user_usage').update({
                'subscription_status': 'free'
            }).eq('stripe_subscription_id', subscription['id']).execute()

        return JSONResponse({'status': 'success'})

    except ImportError:
        raise HTTPException(status_code=500, detail="Required libraries not installed")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
