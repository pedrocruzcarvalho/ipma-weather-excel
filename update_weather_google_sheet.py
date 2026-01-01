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


def read_sheet(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=SHEET_NAME,
    ).execute()
    return result.get("values", [])


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
# HELPERS
# ======================
def _normalize(text):
    return "".join(
        c for c in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(c)
    ).lower()


def _fetch_json(url, retries=3, timeout=15):
    req = urllib.request.Request(
        url, headers={"User-Agent": "ipma-weather-ci/1.0"}
    )
    for i in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return json.load(r)
        except Exception:
            if i == retries - 1:
                raise
            time.sleep(2)


def _geocode(name):
    params = {
        "name": name,
        "count": 10,
        "language": "pt",
        "format": "json",
        "country": "PT",
    }
    url = f"{GEOCODE_URL}?{urllib.parse.urlencode(params)}"
    data = _fetch_json(url)
    results = data.get("results") or []

    target = _normalize(name)

    def score(item):
        score = 0
        if _normalize(item.get("name", "")) == target:
            score += 3
        if item.get("country_code") == "PT":
            score += 2
        elevation = item.get("elevation")
        if elevation is not None:
            score += max(0, 100 - abs(elevation))
        return score

    results.sort(key=score, reverse=True)
    return results[0] if results else None


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
    max_dirs = {_degrees_to_compass(d) for s, d in zip(speeds, directions) if s == max_speed}

    unique_speeds = sorted(set(speeds), reverse=True)
    second_speed = unique_speeds[1] if len(unique_speeds) > 1 else None
    second_dirs = (
        {_degrees_to_compass(d) for s, d in zip(speeds, directions) if s == second_speed}
        if second_speed else set()
    )

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
    existing = read_sheet(service)

    if not existing:
        existing = [HEADERS]

    today_rows = []
    run_date = None

    for concelho in CONCELHOS:
        geo = _geocode(concelho)
        forecast = _forecast_today(geo["latitude"], geo["longitude"])
        run_date = forecast["date"]

        today_rows.append([
            forecast["date"],
            geo["name"],
            forecast["t_min_c"],
            forecast["t_max_c"],
            forecast["wind_max_kmh"],
            forecast["wind_max_dir"],
            forecast["wind_second_max_kmh"],
            forecast["wind_second_max_dir"],
        ])

    # remove existing block for today
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

    # append empty line + today
    if len(new_sheet) > 1:
        new_sheet.append([""] * len(HEADERS))
    new_sheet.extend(today_rows)

    write_sheet(service, new_sheet)
    print("✅ Appended new day correctly")


if __name__ == "__main__":
    main()
