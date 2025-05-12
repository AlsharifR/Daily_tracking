import requests
This Message is from an External Domain (outside Ma’aden) if this Email is Suspicious, please Click on "Report Phish" button. 
import os
import io
import logging
from datetime import datetime
import pandas as pd
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext

# إعداد اللوق
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# إعدادات API
API_KEY = "K-CBC9E418-8984-4176-A285-0F33D699285A" 

# إعدادات SharePoint
SHAREPOINT_SITE = "https://maaden.sharepoint.com/sites/IMC_MaterialHandlingTeam"
SHAREPOINT_DOC_LIB = "Shared Documents"
EXCEL_FILE_NAME = "TrackingSheet.xlsx"

# بيانات الدخول من GitHub Secrets
USERNAME = os.environ["SHAREPOINT_USERNAME"]
PASSWORD = os.environ["SHAREPOINT_PASSWORD"]

def track_shipment(number_to_track):
    url = "https://tracking.searates.com/tracking"
    headers = {"Content-Type": "application/json"}
    payload = {"api_key": API_KEY, "number": number_to_track}

    try:
        response = requests.get(url, json=payload, headers=headers, timeout=30)
        logger.info(f"\n--- Response code: {response.status_code} for {number_to_track} ---")
        logger.info(f"Full API response:\n{response.text}")

        if response.status_code == 200:
            data = response.json().get("data", {})
            metadata = data.get("metadata", {})
            route = data.get("route", {})
            locations = data.get("locations", [])

            pol_num = route.get("pol", {}).get("location", 0)
            pod_num = route.get("pod", {}).get("location", 0)

            pol = locations[pol_num - 1].get("name") if pol_num and len(locations) >= pol_num else ""
            pod = locations[pod_num - 1].get("name") if pod_num and len(locations) >= pod_num else ""
            status = metadata.get("status", "Unknown")
            last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")

            logger.info(f"{number_to_track} | POL: {pol} | POD: {pod} | Status: {status}")
            return pol, pod, status, last_updated
        else:
            return "", "", "Not Found", datetime.now().strftime("%Y-%m-%d %H:%M")

    except Exception as e:
        logger.error(f"Exception for {number_to_track}: {e}")
        return "", "", "Error", datetime.now().strftime("%Y-%m-%d %H:%M")

def get_excel_file(ctx):
    logger.info("Downloading Excel from SharePoint...")
    response = ctx.web.get_file_by_server_relative_url(
        f"/sites/IMC_MaterialHandlingTeam/{SHAREPOINT_DOC_LIB}/{EXCEL_FILE_NAME}"
    ).download().execute_query()
    return pd.read_excel(io.BytesIO(response.content))

def upload_excel_file(ctx, df):
    logger.info("Uploading updated Excel to SharePoint...")
    excel_stream = io.BytesIO()
    df.to_excel(excel_stream, index=False)
    excel_stream.seek(0)
    ctx.web.get_folder_by_server_relative_url(
        f"/sites/IMC_MaterialHandlingTeam/{SHAREPOINT_DOC_LIB}"
    ).upload_file(EXCEL_FILE_NAME, excel_stream).execute_query()

def update_tracking():
    ctx_auth = AuthenticationContext(SHAREPOINT_SITE)
    if not ctx_auth.acquire_token_for_user(USERNAME, PASSWORD):
        logger.error("Authentication failed.")
        return

    ctx = ClientContext(SHAREPOINT_SITE, ctx_auth)
    df = get_excel_file(ctx)

    logger.info(f"Columns detected: {df.columns.tolist()}")

    # تأكد من الأعمدة الأساسية
    for col in ["POL", "POD", "Status", "LastUpdated"]:
        if col not in df.columns:
            df[col] = ""

    for i in range(len(df)):
        container = str(df.loc[i, "ContainsNumber"]).strip() if not pd.isna(df.loc[i, "ContainsNumber"]) else ""
        booking = str(df.loc[i, "BookingNumber"]).strip() if not pd.isna(df.loc[i, "BookingNumber"]) else ""
        tracking_number = booking if booking else container

        if not tracking_number:
            logger.info(f"Row {i}: No tracking number, skipped.")
            continue

        pol, pod, status, last_updated = track_shipment(tracking_number)
        df.at[i, "POL"] = pol
        df.at[i, "POD"] = pod
        df.at[i, "Status"] = status
        df.at[i, "LastUpdated"] = last_updated

    upload_excel_file(ctx, df)
    logger.info("Done! SharePoint Excel updated successfully.")

# تشغيل البرنامج
update_tracking()







