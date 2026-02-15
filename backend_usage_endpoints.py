# ADD THESE ENDPOINTS TO YOUR main.py

from fastapi import HTTPException
from pydantic import BaseModel
from datetime import datetime, date

# Add Supabase client
import os
from supabase import create_client, Client

# Initialize Supabase (add environment variable SUPABASE_SECRET_KEY to Render)
SUPABASE_URL = os.getenv('SUPABASE_URL', 'https://wfsszxypjzzpebjhujuu.supabase.co')
SUPABASE_SECRET_KEY = os.getenv('SUPABASE_SECRET_KEY')  # Get from Supabase settings

# Only initialize if secret key exists
supabase: Client = None
if SUPABASE_SECRET_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_SECRET_KEY)

class UsageCheck(BaseModel):
    user_id: str

class UsageRecord(BaseModel):
    user_id: str

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

@app.post("/webhook/stripe")
async def stripe_webhook(request: Request):
    """Handle Stripe webhooks for subscription updates"""
    # TODO: Implement Stripe webhook handling
    # This will update subscription_status to 'pro' when user subscribes
    pass
