import json
import os
import time
import urllib.parse
import urllib.request

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# ======================
# CONFIG
# ======================
SPREADSHEET_ID = "1jZRnRVneEVqjwjWGNanOJkXyZvVTniWmqDwjzVUmwNk"
SHEET_NAME = "Sheet1"
TIMEZONE = "Europe/Lisbon"

FORECAST_URL = "https://api.open-meteo.com/v1/forecast"

# FIXED, VERIFIED COORDINATES (MAINLAND PORTUGAL)
CONCELHOS = {
    "Sobral de Monte Agraço": (39.019, -9.150),
    "Torres Vedras": (39.091, -9.258),
    "Lourinhã": (39.243, -9.312),
    "Caldas da Rainha": (39.403, -9.138),
    "Cadaval": (39.244, -9.106),
    "Bombarral": (39.267, -9.157),
    "Peniche": (39.355, -9.381),
    "Óbidos": (39.360, -9.157),
}

HEADERS = [
    "date",
    "concelho",
    "t_min_c",
    "t_max_c",
    "wind_max_kmh",
    "wind_max_dir",
    "wind_second_max_kmh",
    "wind_second_max_dir",
]

# ======================
# GOOGLE SHEETS
# ======================
def get_sheets_service():
    creds = Credentials.from_service_account_info(
        json.loads(os.environ["GOOGLE_CREDENTIALS"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def read_sheet(service):
    res = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
    ).execute()
    return res.get("values", [])


def write_sheet(service, values):
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
        valueInputOption="RAW",
        body={"values": values},
    ).execute()


# ======================
# WEATHER HELPERS
# ======================
def _fetch_json(url, retries=3):
    req = urllib.request.Request(url, headers={"User-Agent": "ipma-weather/1.0"})
    for i in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=15) as r:
                return json.load(r)
        except Exception:
            if i == retries - 1:
                raise
            time.sleep(2)


def _degrees_to_compass(deg):
    dirs = [
        "N","NNE","NE","ENE","E","ESE","SE","SSE",
        "S","SSW","SW","WSW","W","WNW","NW","NNW",
    ]
    return dirs[int((deg + 11.25) / 22.5) % 16]


def _forecast_today(lat, lon):
    params = {
        "latitude": lat,
        "longitude": lon,
        "daily": "temperature_2m_max,temperature_2m_min",
        "hourly": "windspeed_10m,winddirection_10m",
        "forecast_days": 1,
        "timezone": TIMEZONE,
        "windspeed_unit": "kmh",
    }
    url = f"{FORECAST_URL}?{urllib.parse.urlencode(params)}"
    data = _fetch_json(url)

    daily = data["daily"]
    hourly = data["hourly"]

    speeds = hourly["windspeed_10m"]
    directions = hourly["winddirection_10m"]

    max_speed = max(speeds)
    max_dirs = {
        _degrees_to_compass(d)
        for s, d in zip(speeds, directions)
        if s == max_speed
    }

    unique_speeds = sorted(set(speeds), reverse=True)
    second_speed = unique_speeds[1] if len(unique_speeds) > 1 else None
    second_dirs = {
        _degrees_to_compass(d)
        for s, d in zip(speeds, directions)
        if second_speed and s == second_speed
    }

    return {
        "date": daily["time"][0],
        "t_min_c": round(daily["temperature_2m_min"][0], 1),
        "t_max_c": round(daily["temperature_2m_max"][0], 1),
        "wind_max_kmh": round(max_speed, 1),
        "wind_max_dir": ",".join(sorted(max_dirs)),
        "wind_second_max_kmh": round(second_speed, 1) if second_speed else "",
        "wind_second_max_dir": ",".join(sorted(second_dirs)),
    }


# ======================
# MAIN
# ======================
def main():
    service = get_sheets_service()
    existing = read_sheet(service) or [HEADERS]

    today_rows = []
    run_date = None

    for name, (lat, lon) in CONCELHOS.items():
        forecast = _forecast_today(lat, lon)
        run_date = forecast["date"]

        today_rows.append([
            forecast["date"],
            name,
            forecast["t_min_c"],
            forecast["t_max_c"],
            forecast["wind_max_kmh"],
            forecast["wind_max_dir"],
            forecast["wind_second_max_kmh"],
            forecast["wind_second_max_dir"],
        ])

    # remove today's existing block
    new_sheet = [HEADERS]
    skip = False
    for row in existing[1:]:
        if row and row[0] == run_date:
            skip = True
            continue
        if skip and row and row[0]:
            skip = False
        if not skip:
            new_sheet.append(row)

    if len(new_sheet) > 1:
        new_sheet.append([""] * len(HEADERS))
    new_sheet.extend(today_rows)

    write_sheet(service, new_sheet)
    print("✅ Data appended with correct locations")


if __name__ == "__main__":
    main()
