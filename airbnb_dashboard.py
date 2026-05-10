"""Generate an interactive Airbnb dashboard from Excel_cleaned.xlsx.

Run:
    python3 airbnb_dashboard.py

The script writes airbnb_dashboard.html next to this file. It only depends on
pandas and the Excel reader already used by the project.
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Excel_cleaned.xlsx"
OUTPUT_FILE = BASE_DIR / "airbnb_dashboard.html"


def clean_currency(series: pd.Series) -> pd.Series:
    """Convert values like '$193' into numeric values."""
    return pd.to_numeric(
        series.astype(str).str.replace(r"[$,]", "", regex=True).str.strip(),
        errors="coerce",
    )


def load_data() -> pd.DataFrame:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Could not find {EXCEL_FILE}")

    df = pd.read_excel(EXCEL_FILE)
    df = df.rename(columns={column: column.strip() for column in df.columns})

    if "SERVICE FEE" in df:
        df["SERVICE FEE($)"] = clean_currency(df["SERVICE FEE"])
    else:
        df["SERVICE FEE($)"] = 0

    df["TOTAL COST($)"] = df["PRICE($)"] + df["SERVICE FEE($)"].fillna(0)
    df["REVIEW INTENSITY"] = df["NUMBER OF REVIEWS"] / df["MINIMUM NIGHTS"].clip(lower=1)

    return df


def unique_values(df: pd.DataFrame, column: str) -> list[str]:
    return sorted(value for value in df[column].dropna().astype(str).unique().tolist() if value)


def to_records(df: pd.DataFrame) -> list[dict[str, object]]:
    dashboard_columns = [
        "NAME",
        "HOST NAME",
        "NEIGHBOURHOOD GROUP",
        "NEIGHBOURHOOD",
        "ROOM TYPE",
        "CANCELLATION_POLICY",
        "HOST_IDENTITY_VERIFIED",
        "INSTANT_BOOKABLE",
        "CONSTRUCTION YEAR",
        "PRICE($)",
        "SERVICE FEE($)",
        "TOTAL COST($)",
        "MINIMUM NIGHTS",
        "NUMBER OF REVIEWS",
        "REVIEW INTENSITY",
        "LAT",
        "LONG",
    ]
    export = df[[column for column in dashboard_columns if column in df]].copy()
    export = export.dropna(subset=["PRICE($)", "LAT", "LONG"])
    export = export.fillna("")
    return export.to_dict(orient="records")


def build_html(df: pd.DataFrame) -> str:
    records = to_records(df)
    options = {
        "groups": unique_values(df, "NEIGHBOURHOOD GROUP"),
        "roomTypes": unique_values(df, "ROOM TYPE"),
        "policies": unique_values(df, "CANCELLATION_POLICY"),
        "verified": unique_values(df, "HOST_IDENTITY_VERIFIED"),
        "years": [
            int(df["CONSTRUCTION YEAR"].min()),
            int(df["CONSTRUCTION YEAR"].max()),
        ],
        "prices": [
            int(df["PRICE($)"].min()),
            int(df["PRICE($)"].quantile(0.99)),
        ],
    }

    payload = json.dumps(records, separators=(",", ":"))
    option_payload = json.dumps(options, separators=(",", ":"))

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Airbnb Listings Dashboard</title>
  <style>
    :root {{
      --bg: #f7f6f2;
      --panel: #ffffff;
      --ink: #202225;
      --muted: #696f78;
      --line: #d9ddd7;
      --accent: #d94747;
      --accent-2: #197278;
      --accent-3: #f2a541;
      --shadow: 0 12px 32px rgba(32, 34, 37, 0.09);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: var(--bg);
      color: var(--ink);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      letter-spacing: 0;
    }}
    header {{
      background: #202225;
      color: white;
      padding: 24px clamp(16px, 4vw, 44px);
      border-bottom: 5px solid var(--accent);
    }}
    h1 {{ margin: 0; font-size: clamp(28px, 4vw, 44px); line-height: 1.05; }}
    header p {{ margin: 10px 0 0; color: #d8ddd9; max-width: 900px; }}
    main {{
      display: grid;
      grid-template-columns: minmax(260px, 320px) minmax(0, 1fr);
      gap: 18px;
      padding: 18px clamp(16px, 4vw, 44px) 36px;
    }}
    aside, section, .metric, .table-wrap {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
    }}
    aside {{ align-self: start; padding: 16px; position: sticky; top: 14px; }}
    .filters {{ display: grid; gap: 12px; }}
    label {{ display: grid; gap: 6px; font-size: 13px; color: var(--muted); font-weight: 700; }}
    select, input, button {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: white;
      color: var(--ink);
      min-height: 38px;
      padding: 8px 10px;
      font: inherit;
    }}
    button {{
      background: var(--ink);
      color: white;
      border-color: var(--ink);
      cursor: pointer;
      font-weight: 800;
    }}
    .content {{ display: grid; gap: 18px; min-width: 0; }}
    .metrics {{ display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }}
    .metric {{ padding: 16px; min-height: 102px; }}
    .metric span {{ display: block; color: var(--muted); font-size: 13px; font-weight: 800; }}
    .metric strong {{ display: block; margin-top: 8px; font-size: clamp(24px, 3vw, 34px); line-height: 1; }}
    .grid {{ display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; }}
    section {{ padding: 16px; min-width: 0; }}
    h2 {{ margin: 0 0 12px; font-size: 17px; }}
    canvas {{ display: block; width: 100%; height: 320px; }}
    #map {{ height: 420px; }}
    .table-wrap {{ overflow: hidden; }}
    .table-head {{ display: flex; justify-content: space-between; gap: 12px; align-items: center; padding: 14px 16px; border-bottom: 1px solid var(--line); }}
    .table-scroll {{ overflow: auto; max-height: 430px; }}
    table {{ border-collapse: collapse; width: 100%; min-width: 980px; font-size: 13px; }}
    th, td {{ text-align: left; border-bottom: 1px solid var(--line); padding: 9px 10px; white-space: nowrap; }}
    th {{ background: #eef1ed; position: sticky; top: 0; z-index: 1; }}
    .empty {{ padding: 28px; color: var(--muted); }}
    @media (max-width: 980px) {{
      main {{ grid-template-columns: 1fr; }}
      aside {{ position: static; }}
      .metrics, .grid {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header>
    <h1>Airbnb Listings Dashboard</h1>
    <p>Explore pricing, inventory, reviews, booking settings, and location patterns from Excel_cleaned.xlsx.</p>
  </header>
  <main>
    <aside>
      <div class="filters">
        <label>Neighbourhood group<select id="group"></select></label>
        <label>Room type<select id="room"></select></label>
        <label>Cancellation policy<select id="policy"></select></label>
        <label>Host verified<select id="verified"></select></label>
        <label>Min construction year<input id="year" type="number"></label>
        <label>Max price<input id="price" type="number" min="0" step="25"></label>
        <button id="reset">Reset filters</button>
      </div>
    </aside>
    <div class="content">
      <div class="metrics">
        <div class="metric"><span>Listings</span><strong id="mListings">0</strong></div>
        <div class="metric"><span>Average price</span><strong id="mPrice">$0</strong></div>
        <div class="metric"><span>Average reviews</span><strong id="mReviews">0</strong></div>
        <div class="metric"><span>Instant bookable</span><strong id="mInstant">0%</strong></div>
      </div>
      <div class="grid">
        <section><h2>Average Price by Neighbourhood Group</h2><canvas id="groupChart"></canvas></section>
        <section><h2>Room Type Mix</h2><canvas id="roomChart"></canvas></section>
        <section><h2>Top Neighbourhoods by Listings</h2><canvas id="neighbourhoodChart"></canvas></section>
        <section><h2>Price Distribution</h2><canvas id="priceChart"></canvas></section>
      </div>
      <section><h2>Listing Locations</h2><canvas id="map"></canvas></section>
      <div class="table-wrap">
        <div class="table-head">
          <h2>Filtered Listings</h2>
          <button id="download" style="max-width: 190px;">Download CSV</button>
        </div>
        <div class="table-scroll" id="table"></div>
      </div>
    </div>
  </main>
  <script>
    const DATA = {payload};
    const OPTIONS = {option_payload};
    const colors = ["#d94747", "#197278", "#f2a541", "#5b6c5d", "#6a4c93", "#2f4858", "#c77d32", "#277da1"];
    const money = new Intl.NumberFormat("en-US", {{ style: "currency", currency: "USD", maximumFractionDigits: 0 }});
    const number = new Intl.NumberFormat("en-US");

    const $ = (id) => document.getElementById(id);
    const controls = ["group", "room", "policy", "verified", "year", "price"].map($);

    function fillSelect(id, values) {{
      const select = $(id);
      select.innerHTML = '<option value="">All</option>' + values.map((value) => `<option>${{value}}</option>`).join("");
    }}

    function initControls() {{
      fillSelect("group", OPTIONS.groups);
      fillSelect("room", OPTIONS.roomTypes);
      fillSelect("policy", OPTIONS.policies);
      fillSelect("verified", OPTIONS.verified);
      $("year").value = OPTIONS.years[0];
      $("price").value = OPTIONS.prices[1];
      controls.forEach((control) => control.addEventListener("input", render));
      $("reset").addEventListener("click", () => {{
        ["group", "room", "policy", "verified"].forEach((id) => $(id).value = "");
        $("year").value = OPTIONS.years[0];
        $("price").value = OPTIONS.prices[1];
        render();
      }});
      $("download").addEventListener("click", downloadCsv);
    }}

    function filteredData() {{
      const group = $("group").value;
      const room = $("room").value;
      const policy = $("policy").value;
      const verified = $("verified").value;
      const year = Number($("year").value || OPTIONS.years[0]);
      const price = Number($("price").value || OPTIONS.prices[1]);
      return DATA.filter((row) =>
        (!group || row["NEIGHBOURHOOD GROUP"] === group) &&
        (!room || row["ROOM TYPE"] === room) &&
        (!policy || row["CANCELLATION_POLICY"] === policy) &&
        (!verified || row["HOST_IDENTITY_VERIFIED"] === verified) &&
        Number(row["CONSTRUCTION YEAR"]) >= year &&
        Number(row["PRICE($)"]) <= price
      );
    }}

    function aggregate(rows, key, value, mode = "count", limit = 10) {{
      const map = new Map();
      rows.forEach((row) => {{
        const label = row[key] || "Unknown";
        if (!map.has(label)) map.set(label, {{ label, count: 0, sum: 0 }});
        const item = map.get(label);
        item.count += 1;
        item.sum += Number(row[value] || 0);
      }});
      return [...map.values()]
        .map((item) => ({{ label: item.label, value: mode === "avg" ? item.sum / item.count : item.count }}))
        .sort((a, b) => b.value - a.value)
        .slice(0, limit);
    }}

    function drawBar(canvasId, items, formatter = (v) => number.format(v)) {{
      const canvas = $(canvasId);
      const rect = canvas.getBoundingClientRect();
      canvas.width = rect.width * devicePixelRatio;
      canvas.height = rect.height * devicePixelRatio;
      const ctx = canvas.getContext("2d");
      ctx.scale(devicePixelRatio, devicePixelRatio);
      ctx.clearRect(0, 0, rect.width, rect.height);
      if (!items.length) return drawEmpty(ctx, rect);
      const left = 150, right = 18, top = 10, rowH = Math.max(24, (rect.height - 24) / items.length);
      const max = Math.max(...items.map((item) => item.value), 1);
      ctx.font = "12px system-ui";
      items.forEach((item, index) => {{
        const y = top + index * rowH;
        const width = (rect.width - left - right) * item.value / max;
        ctx.fillStyle = "#59615c";
        ctx.fillText(String(item.label).slice(0, 22), 0, y + rowH * 0.62);
        ctx.fillStyle = colors[index % colors.length];
        ctx.fillRect(left, y + 5, width, Math.max(10, rowH - 10));
        ctx.fillStyle = "#202225";
        ctx.fillText(formatter(item.value), left + width + 6, y + rowH * 0.62);
      }});
    }}

    function drawDistribution(canvasId, rows) {{
      const values = rows.map((row) => Number(row["PRICE($)"])).filter(Number.isFinite);
      const cap = OPTIONS.prices[1];
      const bins = Array.from({{ length: 12 }}, (_, i) => ({{ label: `$${{Math.round(i * cap / 12)}}`, value: 0 }}));
      values.forEach((value) => bins[Math.min(11, Math.floor(value / cap * 12))].value += 1);
      drawBar(canvasId, bins, (v) => number.format(v));
    }}

    function drawMap(rows) {{
      const canvas = $("map");
      const rect = canvas.getBoundingClientRect();
      canvas.width = rect.width * devicePixelRatio;
      canvas.height = rect.height * devicePixelRatio;
      const ctx = canvas.getContext("2d");
      ctx.scale(devicePixelRatio, devicePixelRatio);
      ctx.clearRect(0, 0, rect.width, rect.height);
      ctx.fillStyle = "#eef1ed";
      ctx.fillRect(0, 0, rect.width, rect.height);
      const points = rows.filter((row) => Number.isFinite(Number(row.LAT)) && Number.isFinite(Number(row.LONG)));
      if (!points.length) return drawEmpty(ctx, rect);
      const lats = points.map((row) => Number(row.LAT));
      const longs = points.map((row) => Number(row.LONG));
      const minLat = Math.min(...lats), maxLat = Math.max(...lats);
      const minLong = Math.min(...longs), maxLong = Math.max(...longs);
      ctx.globalAlpha = 0.22;
      ctx.fillStyle = "#d94747";
      points.slice(0, 16000).forEach((row) => {{
        const x = ((Number(row.LONG) - minLong) / (maxLong - minLong || 1)) * (rect.width - 24) + 12;
        const y = rect.height - (((Number(row.LAT) - minLat) / (maxLat - minLat || 1)) * (rect.height - 24) + 12);
        ctx.beginPath();
        ctx.arc(x, y, 2.1, 0, Math.PI * 2);
        ctx.fill();
      }});
      ctx.globalAlpha = 1;
    }}

    function drawEmpty(ctx, rect) {{
      ctx.fillStyle = "#696f78";
      ctx.font = "14px system-ui";
      ctx.fillText("No listings match the current filters.", 18, rect.height / 2);
    }}

    function updateMetrics(rows) {{
      const count = rows.length;
      const avgPrice = count ? rows.reduce((sum, row) => sum + Number(row["PRICE($)"] || 0), 0) / count : 0;
      const avgReviews = count ? rows.reduce((sum, row) => sum + Number(row["NUMBER OF REVIEWS"] || 0), 0) / count : 0;
      const instant = count ? rows.filter((row) => Number(row["INSTANT_BOOKABLE"]) === 1).length / count * 100 : 0;
      $("mListings").textContent = number.format(count);
      $("mPrice").textContent = money.format(avgPrice);
      $("mReviews").textContent = avgReviews.toFixed(1);
      $("mInstant").textContent = `${{instant.toFixed(0)}}%`;
    }}

    function renderTable(rows) {{
      const sample = rows.slice(0, 250);
      if (!sample.length) {{
        $("table").innerHTML = '<div class="empty">No listings match the current filters.</div>';
        return;
      }}
      const columns = ["NAME", "NEIGHBOURHOOD GROUP", "NEIGHBOURHOOD", "ROOM TYPE", "PRICE($)", "SERVICE FEE($)", "MINIMUM NIGHTS", "NUMBER OF REVIEWS"];
      const head = columns.map((column) => `<th>${{column}}</th>`).join("");
      const body = sample.map((row) => `<tr>${{columns.map((column) => `<td>${{formatCell(column, row[column])}}</td>`).join("")}}</tr>`).join("");
      $("table").innerHTML = `<table><thead><tr>${{head}}</tr></thead><tbody>${{body}}</tbody></table>`;
    }}

    function formatCell(column, value) {{
      if (["PRICE($)", "SERVICE FEE($)"].includes(column)) return money.format(Number(value || 0));
      return String(value ?? "").replace(/[&<>"']/g, (char) => ({{ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }}[char]));
    }}

    function downloadCsv() {{
      const rows = filteredData();
      const columns = ["NAME", "HOST NAME", "NEIGHBOURHOOD GROUP", "NEIGHBOURHOOD", "ROOM TYPE", "PRICE($)", "SERVICE FEE($)", "MINIMUM NIGHTS", "NUMBER OF REVIEWS"];
      const csv = [
        columns.join(","),
        ...rows.map((row) => columns.map((column) => `"${{String(row[column] ?? "").replaceAll('"', '""')}}"`).join(","))
      ].join("\\n");
      const blob = new Blob([csv], {{ type: "text/csv" }});
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "filtered_airbnb_listings.csv";
      link.click();
      URL.revokeObjectURL(url);
    }}

    function render() {{
      const rows = filteredData();
      updateMetrics(rows);
      drawBar("groupChart", aggregate(rows, "NEIGHBOURHOOD GROUP", "PRICE($)", "avg"), money.format);
      drawBar("roomChart", aggregate(rows, "ROOM TYPE", "PRICE($)", "count"), (v) => number.format(v));
      drawBar("neighbourhoodChart", aggregate(rows, "NEIGHBOURHOOD", "PRICE($)", "count", 12), (v) => number.format(v));
      drawDistribution("priceChart", rows);
      drawMap(rows);
      renderTable(rows);
    }}

    window.addEventListener("resize", render);
    initControls();
    render();
  </script>
</body>
</html>"""


def main() -> None:
    df = load_data()
    OUTPUT_FILE.write_text(build_html(df), encoding="utf-8")
    print(f"Dashboard written to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
