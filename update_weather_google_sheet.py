import json
import os
import time
import unicodedata
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

GEOCODE_URL = "https://geocoding-api.open-meteo.com/v1/search"
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"

CONCELHOS = [
    "Sobral de Monte Agraco",
    "Torres Vedras",
    "Lourinha",
    "Caldas da Rainha",
    "Cadaval",
    "Bombarral",
    "Peniche",
    "Obidos",
]

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
    creds_info = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(
        creds_info,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def write_sheet(service, values):
    sheet = service.spreadsheets()

    # Clear existing content
    sheet.values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
    ).execute()

    # Write new content
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
        valueInputOption="RAW",
        body={"values": values},
    ).execute()


# ======================
# WEATHER HELPERS
# ======================
def _normalize(text):
    return "".join(
        ch
        for ch in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(ch)
    ).lower()


def _fetch_json(url):
    req = urllib.request.Request(
        url, headers={"User-Agent": "ipma-weather/1.0"}
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def _geocode(name):
    params = {
        "name": name,
        "count": 5,
        "language": "pt",
        "format": "json",
        "country": "PT",
    }
    url = f"{GEOCODE_URL}?{urllib.parse.urlencode(params)}"
    data = _fetch_json(url)
    results = data.get("results") or []
    return results[0] if results else None


def _degrees_to_compass(deg):
    dirs = [
        "N", "NNE", "NE", "ENE",
        "E", "ESE", "SE", "SSE",
        "S", "SSW", "SW", "WSW",
        "W", "WNW", "NW", "NNW",
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
    max_dirs = [
        _degrees_to_compass(d)
        for s, d in zip(speeds, directions)
        if s == max_speed
    ]

    unique_dirs = []
    for d in max_dirs:
        if d not in unique_dirs:
            unique_dirs.append(d)

    return {
        "date": daily["time"][0],
        "t_min_c": round(daily["temperature_2m_min"][0], 1),
        "t_max_c": round(daily["temperature_2m_max"][0], 1),
        "wind_max_kmh": round(max_speed, 1),
        "wind_max_dir": ",".join(unique_dirs),
    }


# ======================
# MAIN
# ======================
def main():
    service = get_sheets_service()

    values = [HEADERS]

    for concelho in CONCELHOS:
        geo = _geocode(concelho)
        if not geo:
            continue

        forecast = _forecast_today(
            geo["latitude"],
            geo["longitude"],
        )

        values.append([
            forecast["date"],
            geo["name"],
            forecast["t_min_c"],
            forecast["t_max_c"],
            forecast["wind_max_kmh"],
            forecast["wind_max_dir"],
            "",
            "",
        ])

        time.sleep(0.2)

    write_sheet(service, values)
    print("Google Sheet updated successfully")


if __name__ == "__main__":
    main()
