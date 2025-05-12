import requests
import logging
import pandas as pd
from datetime import datetime

# Ø¥Ø¹Ø¯Ø§Ø¯ Ø§Ù„Ù„ÙˆÙ‚
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Ù…ÙØªØ§Ø¬ API
API_KEY = "K-CBC9E418-8984-4176-A285-0F33D699285A"

def track_shipment(number_to_track):
    url = f"https://tracking.searates.com/tracking"
    url_header = {"Content-Type": "application/json"}
    url_body = {"api_key": API_KEY, "number": number_to_track}

    response = requests.get(url, json=url_body, headers=url_header)

    if response.status_code == 200:
        try:
            data = response.json().get("data")
            locations = data.get("locations")
            pol_num = data.get("route").get("pol").get("location")
            pod_num = data.get("route").get("pod").get("location")

            pol = locations[pol_num - 1].get("name")
            pod = locations[pod_num - 1].get("name")
            status = data.get("metadata").get("status")
            last_updated = datetime.now().strftime("%Y-%m-%d %H:%M")

            logger.info(f"{number_to_track} | POL: {pol} | POD: {pod} | Status: {status}")
            return pol, pod, status, last_updated

        except Exception as e:
            logger.info(f"Error parsing response for {number_to_track}: {e}")
            return "", "", "Error parsing", datetime.now().strftime("%Y-%m-%d %H:%M")

    else:
        logger.info(f"Failed to track {number_to_track}")
        return "", "", "Not Found", datetime.now().strftime("%Y-%m-%d %H:%M")

def update_excel(file_path):
    df = pd.read_excel(file_path)

    for i in range(len(df)):
        number = df.loc[i, "ContainsNumber"]
        pol, pod, status, last_updated = track_shipment(number)
        df.at[i, "POL"] = pol
        df.at[i, "POD"] = pod
        df.at[i, "Status"] = status
        df.at[i, "LastUpdated"] = last_updated

    df.to_excel(file_path, index=False)
    logger.info("Excel file updated.")

# Ø´ØºÙ„ Ø§Ù„ÙƒÙˆØ¯
update_excel("TrackingSheet.xlsx")
