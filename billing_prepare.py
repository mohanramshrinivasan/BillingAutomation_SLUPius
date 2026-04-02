# ==========================================================
# BILLING AUTOMATION SCRIPT
# ==========================================================
# What this script does:
# 1. Reads config values from JSON
# 2. Reads patron/item data from Excel
# 3. Groups multiple rows for the same email into one email
# 4. Chooses the correct invoice stage (1, 2, 3, or final)
# 5. Builds one combined Outlook draft per person
# 6. Marks the Outlook draft as High Importance
# 7. Updates the correct invoice date in Excel
# 8. Changes Run from Yes to Done
# 9. Writes logs to a log file
# ==========================================================

import sys
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd


# ----------------------------------------------------------
# Tee class
# Purpose: write terminal output to both screen and log file
# ----------------------------------------------------------
class Tee:
    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.original_stdout = sys.stdout

    def write(self, message):
        self.original_stdout.write(message)
        self.file_obj.write(message)

    def flush(self):
        self.original_stdout.flush()
        self.file_obj.flush()


# ----------------------------------------------------------
# Project folder paths
# ----------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = DATA_DIR / "templates"
INPUT_DIR = DATA_DIR / "input"
EXCEL_DIR = DATA_DIR / "excel"
LOGS_DIR = BASE_DIR / "logs"

CONFIG_FILE = INPUT_DIR / "input_data.json"
EXCEL_FILE = EXCEL_DIR / "billing_data.xlsx"


# ----------------------------------------------------------
# Helper functions
# ----------------------------------------------------------
def timestamp() -> str:
    return datetime.now().strftime("%m%d%y-%H%M%S")


def run_date() -> str:
    return datetime.now().strftime("%m/%d/%Y")


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def load_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_template(path: Path) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def filled(value) -> bool:
    return not pd.isna(value) and str(value).strip() != ""


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def format_name(name) -> str:
    name = clean_text(name)
    if not name:
        return ""
    if "," in name:
        last, first = name.split(",", 1)
        return f"{first.strip().title()} {last.strip().title()}"
    return name.title()


def format_money(value) -> str:
    try:
        return f"{float(value):.2f}"
    except Exception:
        return "0.00"


def money_float(value) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(text)).strip("_")


# ----------------------------------------------------------
# Determine invoice stage for one row
# ----------------------------------------------------------
def get_stage(row: pd.Series, config: dict) -> int:
    d1 = row.get(config["invoice_date_1_column"])
    d2 = row.get(config["invoice_date_2_column"])
    d3 = row.get(config["invoice_date_3_column"])
    d4 = row.get(config["invoice_date_4_column"])

    if not filled(d1):
        return 1
    if filled(d1) and not filled(d2):
        return 2
    if filled(d1) and filled(d2) and not filled(d3):
        return 3
    if filled(d1) and filled(d2) and filled(d3) and not filled(d4):
        return 4
    return 0


def get_stage_column(stage: int, config: dict) -> str:
    if stage == 1:
        return config["invoice_date_1_column"]
    if stage == 2:
        return config["invoice_date_2_column"]
    if stage == 3:
        return config["invoice_date_3_column"]
    if stage == 4:
        return config["invoice_date_4_column"]
    return ""


# ----------------------------------------------------------
# Build grouped email content for one person/email
# ----------------------------------------------------------
def build_group_context(group_df: pd.DataFrame, config: dict) -> dict:
    first_row = group_df.iloc[0]

    item_lines_list = []
    item_details_list = []
    total_fines = 0.0

    for _, row in group_df.iterrows():
        title = clean_text(row.get(config["title_column"], ""))
        barcode = clean_text(row.get(config["barcode_column"], ""))
        call_number = clean_text(row.get(config["call_number_column"], ""))
        fines_value = money_float(row.get(config["fines_column"], 0))

        total_fines += fines_value

        item_lines_list.append(f"{title} - ${format_money(fines_value)}")

        item_details_list.append(
            f"Title/Author: {title}\n"
            f"Call Number: {call_number}\n"
            f"Barcode: {barcode}\n"
            f"Replacement Cost: ${format_money(fines_value)}"
        )

    item_lines = "\n".join(item_lines_list)
    item_details = "\n\n".join(item_details_list)

    return {
        "Preferred_Name": format_name(first_row.get(config["name_column"], "")),
        "Title": clean_text(first_row.get(config["title_column"], "")),
        "Barcode": clean_text(first_row.get(config["barcode_column"], "")),
        "Call_Number": clean_text(first_row.get(config["call_number_column"], "")),
        "Fines": format_money(first_row.get(config["fines_column"], 0)),
        "Item_Lines": item_lines,
        "Item_Details": item_details,
        "Total_Fines": format_money(total_fines),
        "From_Email": clean_text(config.get("from_email", "")),
        "Today": run_date(),
    }


# ----------------------------------------------------------
# Prepare Outlook draft on Windows
# This creates the draft, marks it High Importance,
# saves it in Drafts, and opens it for review
# ----------------------------------------------------------
def outlook_prepare_windows(to_email: str, subject: str, body: str, from_email: str = "") -> str:
    import win32com.client  # type: ignore

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = to_email
    mail.Subject = subject
    mail.Body = body

    # Mark draft as High Importance
    # Outlook values:
    # 0 = Low
    # 1 = Normal
    # 2 = High
    mail.Importance = 2

    # Try to use the configured sending account
    if from_email:
        session = outlook.Session
        for account in session.Accounts:
            try:
                if str(account.SmtpAddress).lower() == from_email.lower():
                    mail.SendUsingAccount = account
                    break
            except Exception:
                pass

        # Useful for shared mailbox / on behalf of scenarios
        try:
            mail.SentOnBehalfOfName = from_email
        except Exception:
            pass

    # Save to Outlook Drafts and open for manual review
    mail.Save()
    mail.Display()

    return "Outlook draft prepared, displayed, and marked High Importance"


# ----------------------------------------------------------
# Main script
# ----------------------------------------------------------
def main():
    ensure_dir(LOGS_DIR)

    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"Missing config file: {CONFIG_FILE}")

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Missing Excel file: {EXCEL_FILE}")

    config = load_json(CONFIG_FILE)

    for key, value in config.items():
        if isinstance(value, str):
            config[key] = value.strip()

    run_stamp = timestamp()
    log_file = LOGS_DIR / f"logs-{run_stamp}.txt"

    log_f = open(log_file, "w", encoding="utf-8")
    sys.stdout = Tee(log_f)
    sys.stderr = Tee(log_f)

    try:
        print(f"Run started: {run_stamp}")
        print(f"Workbook: {EXCEL_FILE}")

        template_map = {
            1: {
                "subject": config["template_1_subject"],
                "body": load_template(TEMPLATES_DIR / config["template_1_file"]),
            },
            2: {
                "subject": config["template_2_subject"],
                "body": load_template(TEMPLATES_DIR / config["template_2_file"]),
            },
            3: {
                "subject": config["template_3_subject"],
                "body": load_template(TEMPLATES_DIR / config["template_3_file"]),
            },
            4: {
                "subject": config["template_4_subject"],
                "body": load_template(TEMPLATES_DIR / config["template_4_file"]),
            },
        }

        df = pd.read_excel(EXCEL_FILE, sheet_name=config["sheet_name"])
        df.columns = df.columns.str.strip()

        # Make invoice date columns text-safe
        date_columns = [
            config["invoice_date_1_column"],
            config["invoice_date_2_column"],
            config["invoice_date_3_column"],
            config["invoice_date_4_column"],
        ]

        for col in date_columns:
            if col in df.columns:
                df[col] = df[col].astype("string")

        if len(df.columns) < 2:
            raise ValueError("Excel must have at least 2 columns. Column 2 should be Run.")

        run_column = df.columns[1]
        df[run_column] = df[run_column].astype("string")

        processed_count = 0
        skip_count = 0
        today_value = run_date()

        # Keep only rows where Run = Yes
        eligible_df = df[df[run_column].fillna("").astype(str).str.strip().str.lower() == "yes"].copy()

        # Group rows by recipient email
        grouped = eligible_df.groupby(config["to_column"], dropna=False)

        for to_email, group_df in grouped:
            to_email = clean_text(to_email)

            if not to_email:
                print("Skipped one group because TO email is blank.")
                skip_count += len(group_df)
                continue

            # Check all invoice stages inside the group
            stages = group_df.apply(lambda row: get_stage(row, config), axis=1).tolist()
            unique_stages = set(stages)

            if unique_stages == {0}:
                print(
                    f"Email {to_email}: Cannot draft email since all invoice date fields are already filled. "
                    f"Please check the invoice dates."
                )
                skip_count += len(group_df)
                continue

            # For safety, all grouped rows must be in the same stage
            valid_stages = {s for s in unique_stages if s != 0}
            if len(valid_stages) != 1:
                print(
                    f"Email {to_email}: skipped because rows have mixed invoice stages {sorted(unique_stages)}. "
                    f"Please clean Excel data."
                )
                skip_count += len(group_df)
                continue

            stage = valid_stages.pop()
            template_info = template_map[stage]
            subject = template_info["subject"]
            from_email = clean_text(config.get("from_email", ""))

            # Build grouped email text
            context = build_group_context(group_df, config)

            try:
                body = template_info["body"].format(**context)
            except KeyError as e:
                print(f"Email {to_email}: template placeholder missing: {e}")
                skip_count += len(group_df)
                continue

            try:
                action_result = outlook_prepare_windows(
                    to_email=to_email,
                    subject=subject,
                    body=body,
                    from_email=from_email
                )
            except Exception as e:
                print(f"Email {to_email}: Outlook draft preparation failed: {e}")
                skip_count += len(group_df)
                continue

            stage_column = get_stage_column(stage, config)

            # Update all rows in this group
            for idx in group_df.index:
                if stage_column:
                    df.at[idx, stage_column] = today_value
                df.at[idx, run_column] = "Done"

            processed_count += len(group_df)

            print(
                f"Email {to_email}: {action_result} | "
                f"stage={stage} | updated {stage_column} to {today_value} | "
                f"updated {run_column} to Done | rows combined={len(group_df)} | "
                f"total=${context['Total_Fines']}"
            )

        # Save updated Excel data back
        with pd.ExcelWriter(
            EXCEL_FILE,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name=config["sheet_name"], index=False)

        print(f"Run completed. Processed rows: {processed_count}. Rows skipped: {skip_count}.")
        print(f"Log file: {log_file}")

    finally:
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        log_f.close()


# ----------------------------------------------------------
# Script start point
# ----------------------------------------------------------
if __name__ == "__main__":
    main()