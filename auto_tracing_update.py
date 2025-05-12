import requests
import logging
import pandas as pd
from datetime import datetime


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


API_KEY = "K-CBC9E418-8984-4176-A285-0F33D699285A"

def track_shipment(number_to_track):
    url = "https://tracking.searates.com/tracking"
    headers = {"Content-Type": "application/json"}
    payload = {"api_key": API_KEY, "number": number_to_track}

    try:
        response = requests.get(url, json=payload, headers=headers, timeout=30)
        logger.info(f"Response code: {response.status_code}")

        if response.status_code == 200:
            data = response.json().get("data", {})

            metadata = data.get("metadata", {})
            route = data.get("route", {})
            locations = data.get("locations", [])

            pol_num = route.get("pol", {}).get("location", 0)
            pod_num = route.get("pod", {}).get("location", 0)

            pol = locations[pol_num - 1].get("name") if pol_num > 0 and len(locations) >= pol_num else ""
            pod = locations[pod_num - 1].get("name") if pod_num > 0 and len(locations) >= pod_num else ""
            status = metadata.get("status", "Unknown")
            last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")

            logger.info(f"{number_to_track} | POL: {pol} | POD: {pod} | Status: {status}")
            return pol, pod, status, last_updated
        else:
            logger.warning(f"Failed for {number_to_track}")
            return "", "", "Not Found", datetime.now().strftime("%Y-%m-%d %H:%M")

    except Exception as e:
        logger.error(f"Exception for {number_to_track}: {e}")
        return "", "", "Error", datetime.now().strftime("%Y-%m-%d %H:%M")

def update_excel(file_path):
    df = pd.read_excel(file_path)

    # تأكد من الأعمدة الأساسية موجودة
    for col in ["POL", "POD", "Status", "LastUpdated"]:
        if col not in df.columns:
            df[col] = ""

    for i in range(len(df)):
        container = str(df.loc[i, "ContainsNumber"]).strip() if not pd.isna(df.loc[i, "ContainsNumber"]) else ""
        booking = str(df.loc[i, "BookingNumber "]).strip() if not pd.isna(df.loc[i, "BookingNumber "]) else ""

        tracking_number = booking if booking else container
        if not tracking_number:
            logger.info(f"Row {i}: No tracking number found, skipping.")
            continue

        pol, pod, status, last_updated = track_shipment(tracking_number)
        df.at[i, "POL"] = pol
        df.at[i, "POD"] = pod
        df.at[i, "Status"] = status
        df.at[i, "LastUpdated"] = last_updated

    df.to_excel(file_path, index=False)
    logger.info("Excel file updated successfully.")


update_excel("TrackingSheet.xlsx")



# تشغيل الكود
update_excel("TrackingSheet.xlsx")


