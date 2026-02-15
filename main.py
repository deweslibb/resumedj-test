from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import shutil
import os
import tempfile
from pathlib import Path
from datetime import datetime
import glob
from pydantic import BaseModel
from typing import Optional

from resume_generator import create_resume_from_file, preview_resume_content

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

TEMPLATE_PATH = Path("Resume_Creator.xlsx")

# Store file paths temporarily
file_storage = {}

# Pydantic models for request/response
class CheckoutSessionRequest(BaseModel):
    user_id: str
    email: str
    price_id: str

class StripeWebhookEvent(BaseModel):
    type: str
    data: dict

@app.get("/")
async def root():
    return {
        "message": "Professional Resume Generator API",
        "version": "1.0.0",
        "endpoints": {
            "GET /template": "Download the Excel template",
            "POST /generate": "Upload filled Excel and get resume",
            "GET /health": "Health check",
            "POST /create-checkout-session": "Create Stripe checkout session",
            "POST /webhook/stripe": "Handle Stripe webhook events"
        }
    }

@app.get("/health")
async def health_check():
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

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

# ====================
# STRIPE INTEGRATION
# ====================

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
        from supabase import create_client, Client

        # Get environment variables
        stripe.api_key = os.getenv('STRIPE_SECRET_KEY')
        webhook_secret = os.getenv('STRIPE_WEBHOOK_SECRET')
        supabase_url = os.getenv('SUPABASE_URL')
        supabase_key = os.getenv('SUPABASE_SERVICE_KEY')

        if not all([stripe.api_key, webhook_secret, supabase_url, supabase_key]):
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

        # Initialize Supabase
        supabase: Client = create_client(supabase_url, supabase_key)

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