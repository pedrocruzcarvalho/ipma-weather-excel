import json
import os
import re
import sys
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from xml.sax.saxutils import escape as xml_escape
from xml.etree import ElementTree

GEOCODE_URL = "https://geocoding-api.open-meteo.com/v1/search"
FORECAST_URL = "https://api.open-meteo.com/v1/forecast"
TIMEZONE = "Europe/Lisbon"
TIMEOUT_SECS = 30

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


def _normalize(text):
    normalized = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in normalized if not unicodedata.combining(ch)).lower()


def _fetch_json(url):
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "ipma-concelhos/1.0"},
    )
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECS) as response:
        return json.load(response)


def _geocode(concelho):
    params = {
        "name": concelho,
        "count": 5,
        "language": "pt",
        "format": "json",
        "country": "PT",
    }
    url = "{}?{}".format(GEOCODE_URL, urllib.parse.urlencode(params))
    data = _fetch_json(url)
    results = data.get("results") or []
    if not results:
        return None

    target = _normalize(concelho)

    def score(item):
        name_norm = _normalize(item.get("name", ""))
        admin2_norm = _normalize(item.get("admin2", ""))
        admin1_norm = _normalize(item.get("admin1", ""))
        points = 0
        if name_norm == target:
            points += 3
        if admin2_norm == target:
            points += 2
        if admin1_norm == target:
            points += 1
        return points

    results = [item for item in results if item.get("country_code") == "PT"] or results
    results.sort(key=score, reverse=True)
    return results[0]


def _round_value(value):
    if value is None:
        return None
    try:
        return round(float(value), 1)
    except (TypeError, ValueError):
        return None


def _degrees_to_compass(degrees):
    if degrees is None:
        return ""
    try:
        degrees = float(degrees)
    except (TypeError, ValueError):
        return ""
    directions = [
        "N",
        "NNE",
        "NE",
        "ENE",
        "E",
        "ESE",
        "SE",
        "SSE",
        "S",
        "SSW",
        "SW",
        "WSW",
        "W",
        "WNW",
        "NW",
        "NNW",
    ]
    index = int((degrees + 11.25) / 22.5) % 16
    return directions[index]


def _unique_in_order(items):
    seen = set()
    ordered = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            ordered.append(item)
    return ordered


def _compute_wind_stats(hourly, target_date):
    times = hourly.get("time") or []
    speeds = hourly.get("windspeed_10m") or []
    directions = hourly.get("winddirection_10m") or []
    if not times or not speeds or not directions:
        return None, [], None, []

    entries = []
    for time_value, speed, direction in zip(times, speeds, directions):
        if target_date and not time_value.startswith(target_date):
            continue
        speed_value = _round_value(speed)
        if speed_value is None or direction is None:
            continue
        entries.append((speed_value, direction))

    if not entries:
        return None, [], None, []

    speed_values = [speed for speed, _ in entries]
    max_speed = max(speed_values)
    max_dirs = [
        _degrees_to_compass(direction)
        for speed, direction in entries
        if speed == max_speed
    ]
    max_dirs = _unique_in_order(max_dirs)

    unique_speeds = sorted(set(speed_values), reverse=True)
    second_speed = unique_speeds[1] if len(unique_speeds) > 1 else None
    second_dirs = []
    if second_speed is not None:
        second_dirs = [
            _degrees_to_compass(direction)
            for speed, direction in entries
            if speed == second_speed
        ]
        second_dirs = _unique_in_order(second_dirs)

    return max_speed, max_dirs, second_speed, second_dirs


def _forecast_today(latitude, longitude):
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "daily": "temperature_2m_max,temperature_2m_min",
        "hourly": "windspeed_10m,winddirection_10m",
        "timezone": TIMEZONE,
        "forecast_days": 1,
        "windspeed_unit": "kmh",
    }
    url = "{}?{}".format(FORECAST_URL, urllib.parse.urlencode(params))
    data = _fetch_json(url)
    daily = data.get("daily") or {}
    hourly = data.get("hourly") or {}
    target_date = (daily.get("time") or [None])[0]
    max_speed, max_dirs, second_speed, second_dirs = _compute_wind_stats(
        hourly, target_date
    )
    return {
        "date": target_date,
        "t_max_c": _round_value((daily.get("temperature_2m_max") or [None])[0]),
        "t_min_c": _round_value((daily.get("temperature_2m_min") or [None])[0]),
        "wind_max_kmh": max_speed,
        "wind_max_dirs": max_dirs,
        "wind_second_max_kmh": second_speed,
        "wind_second_max_dirs": second_dirs,
    }


def _column_letter(index):
    label = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        label = chr(65 + remainder) + label
    return label


def _format_number(value):
    if isinstance(value, float):
        return "{:.1f}".format(value)
    return str(value)


def _cell_xml(value, row_idx, col_idx):
    ref = "{}{}".format(_column_letter(col_idx), row_idx)
    if value is None or value == "":
        return '<c r="{}" />'.format(ref)
    if isinstance(value, bool):
        text = "TRUE" if value else "FALSE"
        return '<c r="{}" t="inlineStr"><is><t>{}</t></is></c>'.format(ref, text)
    if isinstance(value, (int, float)):
        return '<c r="{}"><v>{}</v></c>'.format(ref, _format_number(value))
    text = xml_escape(str(value))
    return '<c r="{}" t="inlineStr"><is><t>{}</t></is></c>'.format(ref, text)


def _write_xlsx(headers, rows, output_path):
    sheet_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "  <sheetData>",
    ]
    for row_idx, row in enumerate([headers] + rows, start=1):
        sheet_lines.append('    <row r="{}">'.format(row_idx))
        for col_idx, value in enumerate(row, start=1):
            sheet_lines.append("      {}".format(_cell_xml(value, row_idx, col_idx)))
        sheet_lines.append("    </row>")
    sheet_lines.extend(["  </sheetData>", "</worksheet>"])
    sheet_xml = "\n".join(sheet_lines)

    content_types = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""
    workbook = """<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Results" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""
    styles = """<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as handle:
        handle.writestr("[Content_Types].xml", content_types)
        handle.writestr("_rels/.rels", rels)
        handle.writestr("xl/workbook.xml", workbook)
        handle.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        handle.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        handle.writestr("xl/styles.xml", styles)


def _join_directions(values):
    if not values:
        return ""
    return ",".join(values)


def _column_index(column_letters):
    index = 0
    for ch in column_letters:
        if not ("A" <= ch <= "Z"):
            break
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index


def _read_shared_strings(zip_handle):
    try:
        data = zip_handle.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ElementTree.fromstring(data)
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    shared = []
    for si in root.findall("main:si", ns):
        parts = []
        for text_node in si.findall(".//main:t", ns):
            parts.append(text_node.text or "")
        shared.append("".join(parts))
    return shared


def _parse_cell_number(text):
    try:
        value = float(text)
    except (TypeError, ValueError):
        return text
    if value.is_integer():
        return int(value)
    return value


def _get_cell_value(cell, shared_strings, ns):
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        text_node = cell.find("main:is/main:t", ns)
        if text_node is not None:
            return text_node.text or ""
        rich_text = cell.findall("main:is/main:r/main:t", ns)
        return "".join(node.text or "" for node in rich_text)
    if cell_type == "s":
        value_node = cell.find("main:v", ns)
        if value_node is None or value_node.text is None:
            return ""
        try:
            index = int(value_node.text)
        except ValueError:
            return ""
        if 0 <= index < len(shared_strings):
            return shared_strings[index]
        return ""
    value_node = cell.find("main:v", ns)
    if value_node is None or value_node.text is None:
        return ""
    if cell_type == "str":
        return value_node.text
    return _parse_cell_number(value_node.text)


def _read_xlsx_rows(path):
    if not os.path.exists(path):
        return []
    with zipfile.ZipFile(path, "r") as handle:
        try:
            sheet_data = handle.read("xl/worksheets/sheet1.xml")
        except KeyError:
            return []
        shared_strings = _read_shared_strings(handle)
    root = ElementTree.fromstring(sheet_data)
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows = []
    for row in root.findall("main:sheetData/main:row", ns):
        row_values = []
        for cell in row.findall("main:c", ns):
            ref = cell.get("r") or ""
            match = re.match(r"([A-Z]+)", ref)
            if not match:
                continue
            col_index = _column_index(match.group(1))
            if col_index <= 0:
                continue
            while len(row_values) < col_index:
                row_values.append("")
            row_values[col_index - 1] = _get_cell_value(cell, shared_strings, ns)
        rows.append(row_values)
    return rows


def _pad_row(row, size):
    row = list(row)
    if len(row) < size:
        row.extend([""] * (size - len(row)))
    return row[:size]


def _is_header_row(row, headers):
    if not row:
        return False
    header_norm = [header.strip().lower() for header in headers]
    row_norm = [str(value).strip().lower() for value in row[: len(header_norm)]]
    return row_norm == header_norm


def _normalize_existing_rows(rows, headers):
    if not rows:
        return rows
    header_norm = [str(value).strip().lower() for value in rows[0]]
    try:
        date_index = header_norm.index("date")
    except ValueError:
        return rows
    if date_index == 0:
        return rows
    for row in rows:
        for cell in row[:date_index]:
            if cell not in ("", None):
                return rows
    normalized = []
    for row in rows:
        normalized.append(row[date_index:])
    return normalized


def _is_empty_row(row):
    return all(cell in ("", None) for cell in row)


def _split_day_blocks(rows):
    blocks = []
    current_rows = []
    current_date = None
    for row in rows:
        if _is_empty_row(row):
            if current_rows:
                blocks.append((current_date, current_rows))
                current_rows = []
                current_date = None
            continue
        row_date = row[0] if row else ""
        if row_date not in ("", None):
            if current_rows and current_date not in ("", None) and row_date != current_date:
                blocks.append((current_date, current_rows))
                current_rows = []
            current_date = row_date
        if current_date is None:
            current_date = row_date
        current_rows.append(row)
    if current_rows:
        blocks.append((current_date, current_rows))
    return blocks


def _ensure_block_date(rows, date_value):
    if date_value in ("", None):
        return rows
    normalized = []
    for row in rows:
        updated = list(row)
        if not updated:
            continue
        updated[0] = date_value
        normalized.append(updated)
    return normalized


def main():
    rows = []
    errors = []
    for concelho in CONCELHOS:
        row = {
            "date": "",
            "concelho": concelho,
            "t_min_c": None,
            "t_max_c": None,
            "wind_max_kmh": None,
            "wind_max_dir": "",
            "wind_second_max_kmh": None,
            "wind_second_max_dir": "",
        }
        try:
            geocode = _geocode(concelho)
            if not geocode:
                errors.append("{}: geocode_not_found".format(concelho))
                rows.append(row)
                continue

            row["concelho"] = geocode.get("name") or concelho

            forecast = _forecast_today(geocode.get("latitude"), geocode.get("longitude"))
            row["date"] = forecast["date"]
            row["t_min_c"] = forecast["t_min_c"]
            row["t_max_c"] = forecast["t_max_c"]
            row["wind_max_kmh"] = forecast["wind_max_kmh"]
            row["wind_max_dir"] = _join_directions(forecast["wind_max_dirs"])
            row["wind_second_max_kmh"] = forecast["wind_second_max_kmh"]
            row["wind_second_max_dir"] = _join_directions(
                forecast["wind_second_max_dirs"]
            )
        except (urllib.error.URLError, json.JSONDecodeError, KeyError) as exc:
            errors.append("{}: {}".format(concelho, exc))
        rows.append(row)
        time.sleep(0.2)

    output_path = "results.xlsx"
    headers = [
        "date",
        "concelho",
        "t_min_c",
        "t_max_c",
        "wind_max_kmh",
        "wind_max_dir",
        "wind_second_max_kmh",
        "wind_second_max_dir",
    ]
    data_rows = []
    run_date = ""
    for row in rows:
        if not run_date and row["date"]:
            run_date = row["date"]
        data_rows.append(
            [
                row["date"],
                row["concelho"],
                row["t_min_c"],
                row["t_max_c"],
                row["wind_max_kmh"],
                row["wind_max_dir"],
                row["wind_second_max_kmh"],
                row["wind_second_max_dir"],
            ]
        )

    existing_rows = _read_xlsx_rows(output_path)
    existing_rows = _normalize_existing_rows(existing_rows, headers)
    if existing_rows and _is_header_row(existing_rows[0], headers):
        existing_data_rows = existing_rows[1:]
    else:
        existing_data_rows = existing_rows
    existing_data_rows = [
        _pad_row(row, len(headers)) for row in existing_data_rows
    ]

    blocks = _split_day_blocks(existing_data_rows)
    ordered_dates = []
    block_map = {}
    for date_value, block_rows in blocks:
        if date_value in block_map:
            continue
        ordered_dates.append(date_value)
        block_map[date_value] = block_rows

    if run_date in block_map:
        block_map[run_date] = data_rows
    else:
        ordered_dates.append(run_date)
        block_map[run_date] = data_rows

    final_rows = []
    empty_row = [""] * len(headers)
    for date_value in ordered_dates:
        block_rows = block_map.get(date_value, [])
        block_rows = _ensure_block_date(block_rows, date_value)
        for row in block_rows:
            final_rows.append(_pad_row(row, len(headers)))
        final_rows.append(empty_row)

    _write_xlsx(headers, final_rows, output_path)

    print("Wrote {} rows to {}".format(len(rows), output_path))
    for error in errors:
        print(error)


if __name__ == "__main__":
    sys.exit(main())
