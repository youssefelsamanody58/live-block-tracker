import asyncio
import datetime as dt
import json
import os
import re
import smtplib
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from playwright.async_api import async_playwright


DATA_DIR = Path("data")
REPORTS_DIR = Path("reports")
UNITS_PATH = Path("units.json")

BASELINE_PATH = DATA_DIR / "blocked_dates_baseline.csv"
BATCH_STATE_PATH = DATA_DIR / "live_batch_state.json"
NEW_BLOCKS_HISTORY_PATH = DATA_DIR / "new_blocks_history.csv"

TODAY_STR = dt.date.today().isoformat()
TIMESTAMP_STR = dt.datetime.utcnow().strftime("%Y-%m-%d_%H-%M-%S")

NEW_BLOCKS_ALERT_XLSX = REPORTS_DIR / f"new_blocks_alert_{TIMESTAMP_STR}.xlsx"
FAILED_UNITS_XLSX = REPORTS_DIR / f"failed_units_{TIMESTAMP_STR}.xlsx"

MONTHS_TO_COLLECT = int(os.getenv("MONTHS_TO_COLLECT") or 6)
BATCH_SIZE = int(os.getenv("BATCH_SIZE") or 20)

SMTP_HOST = os.getenv("SMTP_HOST") or ""
SMTP_PORT = int(os.getenv("SMTP_PORT") or 587)
SMTP_USERNAME = os.getenv("SMTP_USERNAME") or ""
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD") or ""
EMAIL_FROM = os.getenv("EMAIL_FROM") or SMTP_USERNAME
EMAIL_TO = os.getenv("EMAIL_TO") or ""
EMAIL_SUBJECT_PREFIX = os.getenv("EMAIL_SUBJECT_PREFIX") or "Birdnest Live Block Alert"


def ensure_dirs():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)


def load_units():
    with open(UNITS_PATH, "r", encoding="utf-8") as f:
        units = json.load(f)

    cleaned = []
    seen = set()

    for row in units:
        unit_id = str(row.get("unit_id", "")).strip()

        if not unit_id or unit_id in seen:
            continue

        seen.add(unit_id)

        cleaned.append({
            "unit_id": unit_id,
            "url": f"https://www.birdnestlife.com/unit/{unit_id}"
        })

    return cleaned


def load_batch_state(total_units):
    if not BATCH_STATE_PATH.exists():
        return {"next_start_index": 0}

    try:
        with open(BATCH_STATE_PATH, "r", encoding="utf-8") as f:
            state = json.load(f)
    except Exception:
        state = {"next_start_index": 0}

    next_start_index = int(state.get("next_start_index", 0))

    if total_units <= 0:
        next_start_index = 0
    elif next_start_index >= total_units:
        next_start_index = 0

    return {"next_start_index": next_start_index}


def save_batch_state(next_start_index):
    with open(BATCH_STATE_PATH, "w", encoding="utf-8") as f:
        json.dump({"next_start_index": next_start_index}, f, indent=2)


def get_next_batch(units):
    total_units = len(units)
    state = load_batch_state(total_units)

    start = state["next_start_index"]
    end = min(start + BATCH_SIZE, total_units)

    batch = units[start:end]

    if end >= total_units:
        next_start = 0
    else:
        next_start = end

    save_batch_state(next_start)

    return batch, start, end, total_units, next_start


def parse_date(aria):
    if not aria:
        return None

    aria = aria.replace("Today, ", "")
    cleaned = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", aria)

    try:
        parsed = dt.datetime.strptime(cleaned, "%A, %B %d, %Y")
        return parsed.strftime("%Y-%m-%d")
    except Exception:
        return None


async def scrape_unit_blocked_dates(url, unit_id, months_to_collect=6):
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--single-process",
            ],
        )

        page = await browser.new_page(viewport={"width": 1800, "height": 2200})

        await page.goto(url, wait_until="domcontentloaded", timeout=120000)
        await page.wait_for_timeout(5000)

        try:
            open_buttons = page.get_by_role("button", name="Select a date")
            await open_buttons.first.wait_for(state="visible", timeout=10000)
            await open_buttons.first.click()
            await page.wait_for_timeout(2000)
        except Exception:
            print(f"No calendar found for unit {unit_id}. Skipping.")
            await browser.close()
            return []

        deduped = {}

        for i in range(months_to_collect):
            items = await page.eval_on_selector_all(
                "button[aria-label]",
                """
                (buttons) => buttons.map(btn => {
                    const aria = btn.getAttribute('aria-label') || '';
                    const priceSpan = btn.querySelector('span.flex.flex-col > span');
                    const priceText = priceSpan?.textContent?.trim() || '';
                    const disabled = btn.disabled || false;

                    return {
                        aria,
                        priceText,
                        disabled
                    };
                })
                """
            )

            for item in items:
                date = parse_date(item["aria"])
                if not date:
                    continue

                blocked = bool(item["disabled"])
                status = "blocked" if blocked else "available"

                price_text = str(item["priceText"]).strip()
                price_egp = int(price_text) if price_text.isdigit() else None

                deduped[date] = {
                    "unit_id": str(unit_id),
                    "date": date,
                    "status": status,
                    "blocked": blocked,
                    "price_egp": price_egp,
                }

            if i < months_to_collect - 1:
                try:
                    next_button = page.get_by_role("button", name="Go to the Next Month")
                    await next_button.wait_for(state="visible", timeout=5000)
                    await next_button.click(force=True)
                    await page.wait_for_timeout(1500)
                except Exception:
                    print(f"Could not go to next month for unit {unit_id}. Stopping this unit early.")
                    break

        await browser.close()

        return sorted(deduped.values(), key=lambda x: x["date"])


async def scrape_batch(batch_units):
    all_data = []
    failed_units = []

    for index, unit in enumerate(batch_units, start=1):
        unit_id = unit["unit_id"]
        url = unit["url"]

        try:
            print(f"[{index}/{len(batch_units)}] Scraping unit {unit_id}")

            rows = await scrape_unit_blocked_dates(
                url=url,
                unit_id=unit_id,
                months_to_collect=MONTHS_TO_COLLECT,
            )

            all_data.extend(rows)
            print(f"Rows scraped for {unit_id}: {len(rows)}")

        except Exception as e:
            print(f"Failed for {unit_id}: {e}")
            failed_units.append({
                "unit_id": unit_id,
                "url": url,
                "error": str(e),
            })

    df = pd.DataFrame(all_data)

    if df.empty:
        df = pd.DataFrame(columns=["unit_id", "date", "status", "blocked", "price_egp"])
    else:
        df["unit_id"] = df["unit_id"].astype(str)
        df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%Y-%m-%d")
        df["blocked"] = df["blocked"].astype(bool)

        df = (
            df.dropna(subset=["unit_id", "date"])
              .drop_duplicates(subset=["unit_id", "date"], keep="last")
              .sort_values(["unit_id", "date"])
              .reset_index(drop=True)
        )

    failed_df = pd.DataFrame(failed_units)

    return df, failed_df


def load_baseline():
    if not BASELINE_PATH.exists():
        return pd.DataFrame(columns=["unit_id", "date", "status", "blocked", "price_egp"])

    df = pd.read_csv(BASELINE_PATH)

    required_cols = {"unit_id", "date", "status", "blocked"}
    if not required_cols.issubset(df.columns):
        return pd.DataFrame(columns=["unit_id", "date", "status", "blocked", "price_egp"])

    df["unit_id"] = df["unit_id"].astype(str)
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    df["blocked"] = df["blocked"].astype(bool)

    df = (
        df.dropna(subset=["unit_id", "date"])
          .drop_duplicates(subset=["unit_id", "date"], keep="last")
          .sort_values(["unit_id", "date"])
          .reset_index(drop=True)
    )

    return df


def update_baseline(df_baseline, df_batch_today):
    if df_baseline.empty:
        updated = df_batch_today.copy()
    else:
        batch_keys = df_batch_today[["unit_id", "date"]].copy()
        batch_keys["replace_today"] = True

        baseline_without_batch = df_baseline.merge(
            batch_keys,
            on=["unit_id", "date"],
            how="left"
        )

        baseline_without_batch = baseline_without_batch[
            baseline_without_batch["replace_today"].isna()
        ].drop(columns=["replace_today"])

        updated = pd.concat([baseline_without_batch, df_batch_today], ignore_index=True)

    updated = (
        updated.drop_duplicates(subset=["unit_id", "date"], keep="last")
               .sort_values(["unit_id", "date"])
               .reset_index(drop=True)
    )

    updated.to_csv(BASELINE_PATH, index=False)

    return updated


def group_ranges(df):
    results = []

    if df.empty:
        return pd.DataFrame(columns=["unit_id", "start_date", "end_date", "status", "days_count"])

    work = df.copy()
    work["date"] = pd.to_datetime(work["date"])
    work = work.sort_values(["unit_id", "date"]).reset_index(drop=True)

    for unit_id, group in work.groupby("unit_id", sort=True):
        group = group.sort_values("date").reset_index(drop=True)

        start_date = group.loc[0, "date"]
        prev_date = group.loc[0, "date"]
        prev_status = group.loc[0, "status"]

        for i in range(1, len(group)):
            current_date = group.loc[i, "date"]
            current_status = group.loc[i, "status"]

            is_next_day = (current_date - prev_date).days == 1

            if current_status != prev_status or not is_next_day:
                results.append({
                    "unit_id": unit_id,
                    "start_date": start_date.strftime("%Y-%m-%d"),
                    "end_date": prev_date.strftime("%Y-%m-%d"),
                    "status": prev_status,
                    "days_count": (prev_date - start_date).days + 1,
                })

                start_date = current_date
                prev_status = current_status

            prev_date = current_date

        results.append({
            "unit_id": unit_id,
            "start_date": start_date.strftime("%Y-%m-%d"),
            "end_date": prev_date.strftime("%Y-%m-%d"),
            "status": prev_status,
            "days_count": (prev_date - start_date).days + 1,
        })

    return pd.DataFrame(results)


def build_new_blocks(df_batch_today, df_baseline):
    today_blocked = df_batch_today[df_batch_today["blocked"] == True].copy()

    if today_blocked.empty:
        return pd.DataFrame(columns=["unit_id", "start_date", "end_date", "status", "days_count"])

    if df_baseline.empty:
        new_blocks_daily = today_blocked.copy()
    else:
        baseline_blocked_keys = df_baseline[df_baseline["blocked"] == True][["unit_id", "date"]].copy()
        baseline_blocked_keys["was_blocked_before"] = True

        merged = today_blocked.merge(
            baseline_blocked_keys,
            on=["unit_id", "date"],
            how="left",
        )

        new_blocks_daily = merged[merged["was_blocked_before"].isna()].copy()
        new_blocks_daily = new_blocks_daily.drop(columns=["was_blocked_before"])

    if new_blocks_daily.empty:
        return pd.DataFrame(columns=["unit_id", "start_date", "end_date", "status", "days_count"])

    new_blocks_daily["status"] = "new_block"

    grouped = group_ranges(new_blocks_daily[["unit_id", "date", "status", "blocked", "price_egp"]])

    return grouped


def update_new_blocks_history(new_blocks):
    history_columns = [
        "detected_at_utc",
        "detected_date",
        "unit_id",
        "start_date",
        "end_date",
        "status",
        "days_count",
    ]

    if new_blocks.empty:
        if not NEW_BLOCKS_HISTORY_PATH.exists():
            pd.DataFrame(columns=history_columns).to_csv(NEW_BLOCKS_HISTORY_PATH, index=False)
        return

    new_rows = new_blocks.copy()
    new_rows["detected_at_utc"] = dt.datetime.utcnow().isoformat(timespec="seconds")
    new_rows["detected_date"] = TODAY_STR

    new_rows = new_rows[history_columns]

    if NEW_BLOCKS_HISTORY_PATH.exists():
        history = pd.read_csv(NEW_BLOCKS_HISTORY_PATH)
    else:
        history = pd.DataFrame(columns=history_columns)

    combined = pd.concat([history, new_rows], ignore_index=True)

    combined = (
        combined.drop_duplicates(
            subset=["unit_id", "start_date", "end_date", "status"],
            keep="first",
        )
        .sort_values(["detected_at_utc", "unit_id", "start_date"])
        .reset_index(drop=True)
    )

    combined.to_csv(NEW_BLOCKS_HISTORY_PATH, index=False)


def export_alert_report(new_blocks, failed_units):
    if new_blocks.empty and failed_units.empty:
        return

    with pd.ExcelWriter(NEW_BLOCKS_ALERT_XLSX, engine="openpyxl") as writer:
        if not new_blocks.empty:
            new_blocks.to_excel(writer, sheet_name="new_blocks", index=False)
        if not failed_units.empty:
            failed_units.to_excel(writer, sheet_name="failed_units", index=False)


def attach_file(msg, path):
    with open(path, "rb") as f:
        data = f.read()

    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )


def build_email_body(new_blocks, failed_units, batch_start, batch_end, total_units, next_start):
    lines = [
        f"Birdnest LIVE block tracker alert: {TIMESTAMP_STR} UTC",
        "",
        f"Checked batch: units {batch_start + 1} to {batch_end} of {total_units}",
        f"Next batch starts at index: {next_start + 1 if total_units else 0}",
        f"New block ranges detected: {len(new_blocks)}",
        f"Failed units: {len(failed_units)}",
        "",
    ]

    if not new_blocks.empty:
        lines.append("New blocked ranges:")
        lines.append("")

        for unit_id, group in new_blocks.groupby("unit_id", sort=True):
            lines.append(f"Unit {unit_id}")
            for _, row in group.sort_values("start_date").iterrows():
                lines.append(
                    f"  - {row['start_date']} to {row['end_date']} "
                    f"({int(row['days_count'])} days)"
                )
            lines.append("")

    if not failed_units.empty:
        lines.append("Failed units:")
        for _, row in failed_units.iterrows():
            lines.append(f"  - {row['unit_id']}: {row['error']}")

    return "\n".join(lines)


def send_email_if_needed(body, new_blocks, failed_units):
    if new_blocks.empty and failed_units.empty:
        print("No new blocks and no failed units. No email sent.")
        return

    if not all([SMTP_HOST, SMTP_USERNAME, SMTP_PASSWORD, EMAIL_FROM, EMAIL_TO]):
        print("Missing email settings. Skipping email.")
        return

    msg = EmailMessage()
    msg["Subject"] = f"{EMAIL_SUBJECT_PREFIX} - {TIMESTAMP_STR} UTC"
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.set_content(body)

    if NEW_BLOCKS_ALERT_XLSX.exists():
        attach_file(msg, NEW_BLOCKS_ALERT_XLSX)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USERNAME, SMTP_PASSWORD)
        server.send_message(msg)

    print("Alert email sent.")


async def main():
    ensure_dirs()

    units = load_units()

    if not units:
        print("No units found.")
        return

    batch_units, batch_start, batch_end, total_units, next_start = get_next_batch(units)

    print(f"Checking batch {batch_start + 1}-{batch_end} of {total_units}")
    print(f"Batch size: {len(batch_units)}")

    df_baseline = load_baseline()

    df_batch_today, failed_units = await scrape_batch(batch_units)

    new_blocks = build_new_blocks(df_batch_today, df_baseline)

    update_new_blocks_history(new_blocks)

    export_alert_report(new_blocks, failed_units)

    update_baseline(df_baseline, df_batch_today)

    email_body = build_email_body(
        new_blocks=new_blocks,
        failed_units=failed_units,
        batch_start=batch_start,
        batch_end=batch_end,
        total_units=total_units,
        next_start=next_start,
    )

    send_email_if_needed(email_body, new_blocks, failed_units)

    print("Done.")
    print(f"Rows scraped this batch: {len(df_batch_today)}")
    print(f"New block ranges: {len(new_blocks)}")
    print(f"Failed units: {len(failed_units)}")


if __name__ == "__main__":
    asyncio.run(main())
