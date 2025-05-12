import requests
import logging
from datetime import datetime
import pandas as pd
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# إعداد اللوق
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# API
API_KEY = "K-CBC9E418-8984-4176-A285-0F33D699285A"  
# إعداد SharePoint
SHAREPOINT_SITE = "https://maaden.sharepoint.com/:x:/s/IMC_MaterialHandlingTeam"
SHAREPOINT_DOC_LIB = "Shared Documents"  # المكتبة
EXCEL_FILE_NAME = "TrackingSheet.xlsx"  # اسم الملف

USERNAME = "AlsharifR@maaden.com.sa"  # عدّل هنا
PASSWORD = "Magnesite@MH$1245"  # الأفضل وضعه في Secret

# تحميل وتحديث ملف الإكسل
def get_excel_file(ctx):
    response = ctx.web.get_file_by_server_relative_url(f"/sites/your-site-name/{SHAREPOINT_DOC_LIB}/{EXCEL_FILE_NAME}").download().execute_query()
    return pd.read_excel(io.BytesIO(response.content))

# رفع الملف المحدّث
def upload_excel_file(ctx, df):
    excel_stream = io.BytesIO()
    df.to_excel(excel_stream, index=False)
    excel_stream.seek(0)
    ctx.web.get_folder_by_server_relative_url(f"/sites/your-site-name/{SHAREPOINT_DOC_LIB}").upload_file(EXCEL_FILE_NAME, excel_stream).execute_query()

# الاتصال بـ API
def track_shipment(number):
    url = "https://tracking.searates.com/tracking"
    headers = {"Content-Type": "application/json"}
    body = {"api_key": API_KEY, "number": number}
    
    try:
        response = requests.get(url, json=body, headers=headers, timeout=30)
        if response.status_code == 200:
            logger.info(f"Response: {response.text}")
            data = response.json().get("data", {})
            meta = data.get("metadata", {})
            route = data.get("route", {})
            locs = data.get("locations", [])

            pol_idx = route.get("pol", {}).get("location", 0)
            pod_idx = route.get("pod", {}).get("location", 0)

            pol = locs[pol_idx - 1].get("name") if pol_idx and len(locs) >= pol_idx else ""
            pod = locs[pod_idx - 1].get("name") if pod_idx and len(locs) >= pod_idx else ""
            status = meta.get("status", "Unknown")
            last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")
            return pol, pod, status, last_updated
        else:
            return "", "", "Not Found", datetime.now().strftime("%Y-%m-%d %H:%M")
    except Exception as e:
        logger.error(f"Error for {number}: {e}")
        return "", "", "Error", datetime.now().strftime("%Y-%m-%d %H:%M")

# المعالجة الرئيسية
def update_tracking():
    ctx_auth = AuthenticationContext(SHAREPOINT_SITE)
    if not ctx_auth.acquire_token_for_user(USERNAME, PASSWORD):
        logger.error("Authentication failed.")
        return

    ctx = ClientContext(SHAREPOINT_SITE, ctx_auth)
    df = get_excel_file(ctx)

    # تأكد من الأعمدة
    for col in ["POL", "POD", "Status", "LastUpdated"]:
        if col not in df.columns:
            df[col] = ""

    for i in range(len(df)):
        container = str(df.loc[i, "ContainsNumber"]).strip() if not pd.isna(df.loc[i, "ContainsNumber"]) else ""
        booking = str(df.loc[i, "BookingNumber"]).strip() if not pd.isna(df.loc[i, "BookingNumber"]) else ""
        tracking_number = booking if booking else container

        if not tracking_number:
            logger.info(f"Row {i}: No tracking number. Skipped.")
            continue

        pol, pod, status, last_updated = track_shipment(tracking_number)
        df.at[i, "POL"] = pol
        df.at[i, "POD"] = pod
        df.at[i, "Status"] = status
        df.at[i, "LastUpdated"] = last_updated

    upload_excel_file(ctx, df)
    logger.info("SharePoint Excel updated successfully.")

# تشغيل
update_tracking()







