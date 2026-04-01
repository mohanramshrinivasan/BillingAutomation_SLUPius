import sys
import json
import re
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pandas as pd


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


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = DATA_DIR / "templates"
INPUT_DIR = DATA_DIR / "input"
EXCEL_DIR = DATA_DIR / "excel"
OUTPUT_DIR = BASE_DIR / "output"
LOGS_DIR = BASE_DIR / "logs"

CONFIG_FILE = INPUT_DIR / "input_data.json"
EXCEL_FILE = EXCEL_DIR / "billing_data.xlsx"


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


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", str(text)).strip("_")


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


def build_context(row: pd.Series, config: dict) -> dict:
    return {
        "Preferred_Name": format_name(row.get(config["name_column"], "")),
        "Title": clean_text(row.get(config["title_column"], "")),
        "Barcode": clean_text(row.get(config["barcode_column"], "")),
        "Call_Number": clean_text(row.get(config["call_number_column"], "")),
        "Fines": format_money(row.get(config["fines_column"], 0)),
        "From_Email": clean_text(config.get("from_email", "")),
        "Today": run_date(),
    }


def main():
    ensure_dir(OUTPUT_DIR)
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

        if len(df.columns) < 2:
            raise ValueError("Excel must have at least 2 columns. Column 2 should be Run.")

        run_column = df.columns[1]

        draft_count = 0
        skip_count = 0
        today_value = run_date()

        for idx, row in df.iterrows():
            row_number = idx + 2
            run_value = clean_text(row.get(run_column, ""))

            if run_value.lower() != "yes":
                print(f"Row {row_number}: skipped because Run is not Yes.")
                skip_count += 1
                continue

            stage = get_stage(row, config)

            if stage == 0:
                print(
                    f"Row {row_number}: Cannot draft email since all invoice date fields are already filled. "
                    f"Please check the invoice dates."
                )
                skip_count += 1
                continue

            to_email = clean_text(row.get(config["to_column"], ""))
            if not to_email:
                print(f"Row {row_number}: skipped because TO email is blank.")
                skip_count += 1
                continue

            template_info = template_map[stage]
            context = build_context(row, config)

            try:
                body = template_info["body"].format(**context)
            except KeyError as e:
                print(f"Row {row_number}: template placeholder missing: {e}")
                skip_count += 1
                continue

            subject = template_info["subject"]

            msg = EmailMessage()
            msg["From"] = clean_text(config["from_email"])
            msg["To"] = to_email

            reply_to = clean_text(config.get("reply_to", ""))
            if reply_to:
                msg["Reply-To"] = reply_to

            msg["Subject"] = subject
            msg.set_content(body)

            person_name = context["Preferred_Name"] or "Unknown"
            email_stamp = timestamp()
            file_name = f"{safe_filename(person_name)}-{email_stamp}.eml"
            output_file = OUTPUT_DIR / file_name

            with open(output_file, "wb") as f:
                f.write(bytes(msg))

            stage_column = get_stage_column(stage, config)
            if stage_column:
                df.at[idx, stage_column] = today_value

            df.at[idx, run_column] = "Done"

            draft_count += 1
            print(
                f"Row {row_number}: draft created -> {output_file.name} | "
                f"stage={stage} | updated {stage_column} to {today_value} | "
                f"updated {run_column} to Done | to={to_email}"
            )

        with pd.ExcelWriter(
            EXCEL_FILE,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            df.to_excel(writer, sheet_name=config["sheet_name"], index=False)

        print(f"Run completed. Drafts created: {draft_count}. Rows skipped: {skip_count}.")
        print(f"Output directory: {OUTPUT_DIR}")
        print(f"Log file: {log_file}")

    finally:
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        log_f.close()


if __name__ == "__main__":
    main()