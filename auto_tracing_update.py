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

    response = requests.get(url, json=payload, headers=headers)

    if response.status_code == 200:
        try:
            logger.info(f"Raw response: {response.text}")
            data = response.json().get("data", {})
            locations = data.get("locations", [])
            route = data.get("route", {})
            pol_num = route.get("pol", {}).get("location", 0)
            pod_num = route.get("pod", {}).get("location", 0)

            pol = locations[pol_num - 1].get("name") if pol_num > 0 and len(locations) >= pol_num else ""
            pod = locations[pod_num - 1].get("name") if pod_num > 0 and len(locations) >= pod_num else ""
            status = data.get("metadata", {}).get("status", "Unknown")
            last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")

            logger.info(f"{number_to_track} | POL: {pol} | POD: {pod} | Status: {status}")
            return pol, pod, status, last_updated

        except Exception as e:
            logger.info(f"Error parsing response for {number_to_track}: {e}")
            return "", "", "Error parsing", datetime.now().strftime("%Y-%m-%d %H:%M")
    else:
        logger.info(f"Failed to track {number_to_track}. Status code: {response.status_code}")
        return "", "", "Not Found", datetime.now().strftime("%Y-%m-%d %H:%M")

def update_excel(file_path):
    df = pd.read_excel(file_path)

    if "ContainsNumber" not in df.columns:
        raise Exception("Missing column: 'ContainsNumber'")

    for i in range(len(df)):
        number = df.loc[i, "ContainsNumber"]
        pol, pod, status, last_updated = track_shipment(number)
        df.at[i, "POL"] = pol
        df.at[i, "POD"] = pod
        df.at[i, "Status"] = status
        df.at[i, "LastUpdated"] = last_updated

    df.to_excel(file_path, index=False)
    logger.info("Excel file updated.")

# تشغيل الكود
update_excel("TrackingSheet.xlsx")


