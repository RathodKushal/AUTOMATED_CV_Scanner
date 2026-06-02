import os
import time
import gspread
import smtplib
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ==========================================
# Load all config from .env file
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(SCRIPT_DIR, ".env"))

def get_env_var(name, default=None):
    """Gets an env var and cleans it of any accidental quotes or spaces."""
    val = os.getenv(name, default)
    if val:
        return val.strip().strip("'").strip('"')
    return val

# Google Sheets
CREDENTIALS_FILE        = get_env_var("CREDENTIALS_FILE", os.path.join(SCRIPT_DIR, "Credentials.json"))
SHEET_URL_OR_ID         = get_env_var("GOOGLE_SHEET_ID")
RESULTS_SHEET_NAME      = get_env_var("RESULTS_SHEET_NAME", "Ai Responses")

# HR Interview Sheet (the tab HR fills after interviews)
HR_INTERVIEW_SHEET_NAME = get_env_var("HR_INTERVIEW_SHEET_NAME", "HR Interview Results")
HR_CANDIDATE_ID_COLUMN  = get_env_var("HR_CANDIDATE_ID_COLUMN",  "Candidate ID")
HR_RESULT_COLUMN        = get_env_var("HR_RESULT_COLUMN",         "Interview Result")
HR_INTERVIEWER_COLUMN   = get_env_var("HR_INTERVIEWER_COLUMN",    "Interviewer Name")
HR_INTERVIEW_DATE_COLUMN= get_env_var("HR_INTERVIEW_DATE_COLUMN", "Interview Date")

# Email
SENDER_EMAIL    = get_env_var("SENDER_EMAIL")
SENDER_PASSWORD = get_env_var("SENDER_PASSWORD")
SMTP_SERVER     = get_env_var("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT       = int(get_env_var("SMTP_PORT", "587"))

# Polling interval (seconds)
POLL_INTERVAL   = int(get_env_var("POLL_INTERVAL", "60"))

# Column positions in the "Ai Responses" sheet (1-indexed)
COL_ID               = 1
COL_NAME             = 2
COL_EMAIL            = 3
COL_RECOMMENDATION   = 8
COL_JOB_ROLE         = 10
COL_INTERVIEW_RESULT = 23
COL_INTERVIEWER      = 24
COL_INTERVIEW_DATE   = 25


# ==========================================
# Google Sheets Authentication
# ==========================================
def get_sheets_client():
    """Authenticates and returns a gspread client."""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scopes)
        gc = gspread.authorize(creds)
        print("Authenticated with Google successfully.")
        return gc
    except Exception as e:
        print(f"ERROR - Authentication failed: {e}")
        return None


# ==========================================
# Email Sender
# ==========================================
def send_email(to_email, subject, body):
    """Sends a plain-text email via SMTP."""
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        print("WARNING - Email credentials missing in .env. Skipping email.")
        return False
    try:
        msg = MIMEMultipart()
        msg['From']    = SENDER_EMAIL
        msg['To']      = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"ERROR - Failed to send email to {to_email}: {e}")
        return False


# ==========================================
# Sheet Setup
# ==========================================
def ensure_hr_interview_sheet(spreadsheet):
    """
    Opens the HR Interview tab. If it does not exist, creates it with headers.
    HR fills this tab (manually or via Google Form) after each interview.
    """
    try:
        try:
            hr_sheet = spreadsheet.worksheet(HR_INTERVIEW_SHEET_NAME)
            print(f"Tab '{HR_INTERVIEW_SHEET_NAME}' found.")
        except gspread.WorksheetNotFound:
            print(f"Tab '{HR_INTERVIEW_SHEET_NAME}' not found. Creating it...")
            hr_sheet = spreadsheet.add_worksheet(
                title=HR_INTERVIEW_SHEET_NAME, rows="200", cols="10"
            )

        # Only write headers if the sheet is completely empty
        existing_headers = hr_sheet.row_values(1)
        if not existing_headers:
            headers = [
                HR_CANDIDATE_ID_COLUMN,
                HR_RESULT_COLUMN,
                HR_INTERVIEWER_COLUMN,
                HR_INTERVIEW_DATE_COLUMN,
                "Status",
                "Processed At"
            ]
            hr_sheet.append_row(headers)
            print(f"Headers created in '{HR_INTERVIEW_SHEET_NAME}'.")

        return hr_sheet
    except Exception as e:
        print(f"ERROR - Could not prepare HR Interview sheet: {e}")
        return None


# ==========================================
# Helpers
# ==========================================
def find_candidate_row_by_id(results_sheet, candidate_id):
    """
    Searches the 'Ai Responses' sheet for the row whose ID column matches
    candidate_id. Returns the 1-indexed row number, or None if not found.
    """
    try:
        all_ids = results_sheet.col_values(COL_ID)
        candidate_id_str = str(candidate_id).strip()
        for row_idx, cell_value in enumerate(all_ids):
            if row_idx == 0:
                continue  # skip header
            if str(cell_value).strip() == candidate_id_str:
                return row_idx + 1
        return None
    except Exception as e:
        print(f"ERROR - Searching for candidate ID {candidate_id}: {e}")
        return None


def fuzzy_get(record, target_key):
    """
    Case-insensitive, whitespace-tolerant dict lookup.
    Returns the value if found, otherwise empty string.
    """
    target = target_key.strip().lower()
    for key, val in record.items():
        if str(key).strip().lower() == target:
            return str(val).strip()
    return ""


def fuzzy_col_index(headers, target_header):
    """
    Case-insensitive, whitespace-tolerant header to 1-indexed column number.
    Returns None if not found.
    """
    target = target_header.strip().lower()
    for idx, h in enumerate(headers):
        if str(h).strip().lower() == target:
            return idx + 1
    return None


# ==========================================
# Core Processing Logic
# ==========================================
def process_hr_results(hr_sheet, results_sheet):
    """
    Reads every unprocessed row in the HR Interview tab.
    - If result is Pass/Selected  -> writes 'Selected' to Ai Responses, sends selection email.
    - If result is Fail           -> leaves Interview Result blank, sends rejection email.
    - Marks each processed HR row as 'Processed' to avoid duplicate emails.
    """
    all_rows = hr_sheet.get_all_values()
    if not all_rows:
        return 0

    headers = all_rows[0]

    # Show actual column names every poll so mismatches are visible
    print(f"   [DEBUG] Columns in '{HR_INTERVIEW_SHEET_NAME}': {headers}")
    print(f"   [DEBUG] Looking for -> ID:'{HR_CANDIDATE_ID_COLUMN}'  "
          f"Result:'{HR_RESULT_COLUMN}'  "
          f"Interviewer:'{HR_INTERVIEWER_COLUMN}'  "
          f"Date:'{HR_INTERVIEW_DATE_COLUMN}'")

    status_col_idx       = fuzzy_col_index(headers, "Status")
    processed_at_col_idx = fuzzy_col_index(headers, "Processed At")

    data_rows = all_rows[1:]
    processed_count = 0

    # Show how many rows with actual data exist
    non_empty_rows = [r for r in data_rows if any(c.strip() for c in r)]
    print(f"   [DEBUG] Total data rows (non-empty): {len(non_empty_rows)}")
    if not non_empty_rows:
        print("   [DEBUG] WARNING - The HR sheet has no data rows.")
        print("   [DEBUG] Either HR has not submitted the form yet, or")
        print(f"  [DEBUG] the form responses are going to a different tab (see tab list above).")
        return 0

    for i, row in enumerate(data_rows):
        hr_row_number = i + 2

        # Skip completely empty rows
        if not any(c.strip() for c in row):
            continue

        record = dict(zip(headers, row))

        # Show raw row data
        print(f"   [DEBUG] Row {hr_row_number} data: {dict(record)}")

        # Skip already-processed rows
        row_status = fuzzy_get(record, "Status")
        if row_status.lower() == "processed":
            print(f"   [DEBUG] Row {hr_row_number}: Already processed. Skipping.")
            continue

        # Read values
        candidate_id   = fuzzy_get(record, HR_CANDIDATE_ID_COLUMN)
        raw_result     = fuzzy_get(record, HR_RESULT_COLUMN)
        interviewer    = fuzzy_get(record, HR_INTERVIEWER_COLUMN)
        interview_date = fuzzy_get(record, HR_INTERVIEW_DATE_COLUMN)

        print(f"   [DEBUG] Extracted -> candidate_id='{candidate_id}'  "
              f"result='{raw_result}'  interviewer='{interviewer}'  date='{interview_date}'")

        if not candidate_id:
            print(f"   WARNING - Row {hr_row_number}: Candidate ID is empty. "
                  f"HR must enter the exact ID from the 'Ai Responses' sheet.")
            continue
        if not raw_result:
            print(f"   WARNING - Row {hr_row_number}: Candidate ID '{candidate_id}' found "
                  f"but Interview Result is empty. Skipping.")
            continue

        print(f"\nProcessing HR row {hr_row_number} -> Candidate ID: {candidate_id}, Result: {raw_result}")

        # Find the candidate row in Ai Responses by ID
        candidate_row = find_candidate_row_by_id(results_sheet, candidate_id)
        if not candidate_row:
            print(f"   WARNING - Candidate ID '{candidate_id}' not found in '{RESULTS_SHEET_NAME}'.")
            print(f"   WARNING - Make sure the ID in the HR form exactly matches column A of '{RESULTS_SHEET_NAME}'.")
            continue

        # Fetch candidate details
        candidate_name  = results_sheet.cell(candidate_row, COL_NAME).value     or "Candidate"
        candidate_email = results_sheet.cell(candidate_row, COL_EMAIL).value    or ""
        job_role        = results_sheet.cell(candidate_row, COL_JOB_ROLE).value or "the applied role"

        if not candidate_email:
            print(f"   WARNING - No email found for {candidate_name}. Email will be skipped.")

        is_pass = raw_result.lower() in ("pass", "passed", "selected", "yes")

        # Write back to Ai Responses
        if is_pass:
            results_sheet.update_cell(candidate_row, COL_INTERVIEW_RESULT, "Selected")
            results_sheet.update_cell(candidate_row, COL_INTERVIEWER,      interviewer)
            results_sheet.update_cell(candidate_row, COL_INTERVIEW_DATE,   interview_date)
            print(f"   SUCCESS - '{candidate_name}' marked as Selected in '{RESULTS_SHEET_NAME}'.")
        else:
            # Write "Failed" so the result is clearly recorded in the sheet
            results_sheet.update_cell(candidate_row, COL_INTERVIEW_RESULT, "Failed")
            results_sheet.update_cell(candidate_row, COL_INTERVIEWER,      interviewer)
            results_sheet.update_cell(candidate_row, COL_INTERVIEW_DATE,   interview_date)
            print(f"   INFO - '{candidate_name}' did not pass. Interview Result left blank in sheet.")

        # Send email to candidate
        if candidate_email:
            if is_pass:
                subject = f"Congratulations! You are Selected - {job_role}"
                body = (
                    f"Dear {candidate_name},\n\n"
                    f"We are pleased to inform you that you have successfully passed your interview "
                    f"for the {job_role} position.\n\n"
                    f"Your performance during the interview was impressive, and we look forward to "
                    f"welcoming you to our team. Our HR team will contact you shortly with further "
                    f"details regarding your offer letter and joining formalities.\n\n"
                    f"Interview Details:\n"
                    f"  Date        : {interview_date or 'N/A'}\n"
                    f"  Interviewer : {interviewer or 'N/A'}\n\n"
                    f"Congratulations once again. We look forward to having you on board.\n\n"
                    f"Best regards,\n"
                    f"Recruitment Team"
                )
                if send_email(candidate_email, subject, body):
                    print(f"   Selection email sent to {candidate_name} ({candidate_email}).")
            else:
                subject = f"Update Regarding Your Interview for {job_role}"
                body = (
                    f"Dear {candidate_name},\n\n"
                    f"Thank you for attending the interview for the {job_role} position "
                    f"and for the effort you invested in the process.\n\n"
                    f"After careful evaluation, we regret to inform you that we will not be "
                    f"moving forward with your application at this time. This was a highly "
                    f"competitive process and we encourage you not to be disheartened.\n\n"
                    f"We appreciate your interest in our organization and wish you the very best "
                    f"in your future endeavors. We hope to see you again in the future.\n\n"
                    f"Best regards,\n"
                    f"Recruitment Team"
                )
                if send_email(candidate_email, subject, body):
                    print(f"   Rejection email sent to {candidate_name} ({candidate_email}).")

        # Mark HR row as Processed
        if status_col_idx:
            hr_sheet.update_cell(hr_row_number, status_col_idx, "Processed")
        else:
            new_status_col = len(headers) + 1
            hr_sheet.update_cell(1, new_status_col, "Status")
            hr_sheet.update_cell(hr_row_number, new_status_col, "Processed")
            headers.append("Status")
            status_col_idx = new_status_col

        if processed_at_col_idx:
            hr_sheet.update_cell(
                hr_row_number, processed_at_col_idx,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )

        processed_count += 1
        time.sleep(1)

    return processed_count


# ==========================================
# Main Entry Point
# ==========================================
def main():
    if not SHEET_URL_OR_ID:
        print("ERROR - GOOGLE_SHEET_ID is missing in .env. Exiting.")
        return

    gc = get_sheets_client()
    if not gc:
        return

    # Open the spreadsheet
    try:
        print(f"Opening spreadsheet (ID: {SHEET_URL_OR_ID})...")
        if "http" in SHEET_URL_OR_ID:
            spreadsheet = gc.open_by_url(SHEET_URL_OR_ID)
        else:
            spreadsheet = gc.open_by_key(SHEET_URL_OR_ID)
        print("Spreadsheet opened successfully.")
    except Exception as e:
        print(f"ERROR - Could not open spreadsheet: {e}")
        return

    # ------------------------------------------------------------------
    # STARTUP: List all tabs so we can see where form responses actually go
    # ------------------------------------------------------------------
    print("\n" + "=" * 60)
    print("ALL TABS IN THIS SPREADSHEET:")
    print("=" * 60)
    for ws in spreadsheet.worksheets():
        all_rows = ws.get_all_values()
        headers  = all_rows[0] if all_rows else []
        data_cnt = len([r for r in all_rows[1:] if any(c.strip() for c in r)]) if len(all_rows) > 1 else 0
        marker   = "  <-- HR.py is reading THIS tab" if ws.title == HR_INTERVIEW_SHEET_NAME else ""
        print(f"\n  Tab  : \"{ws.title}\"{marker}")
        print(f"  Rows : {data_cnt}")
        print(f"  Cols : {headers}")
    print("\n" + "=" * 60)
    print(f"HR_INTERVIEW_SHEET_NAME in .env = \"{HR_INTERVIEW_SHEET_NAME}\"")
    print("If your HR form data is in a different tab, update HR_INTERVIEW_SHEET_NAME in .env")
    print("=" * 60 + "\n")
    # ------------------------------------------------------------------

    hr_sheet = ensure_hr_interview_sheet(spreadsheet)
    if not hr_sheet:
        print("ERROR - Could not open or create the HR Interview sheet. Exiting.")
        return

    try:
        results_sheet = spreadsheet.worksheet(RESULTS_SHEET_NAME)
        print(f"Connected to '{RESULTS_SHEET_NAME}' sheet.")
    except gspread.WorksheetNotFound:
        print(f"ERROR - Sheet '{RESULTS_SHEET_NAME}' not found. Run main.py first.")
        return

    print(f"\nHR Interview Result Processor started.")
    print(f"Polling '{HR_INTERVIEW_SHEET_NAME}' every {POLL_INTERVAL} seconds.")
    print(f"Press Ctrl+C to stop.\n")

    while True:
        try:
            count = process_hr_results(hr_sheet, results_sheet)
            if count > 0:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Processed {count} interview result(s).")
            else:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] No new HR results to process.")
        except Exception as e:
            print(f"WARNING - {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nHR processor stopped.")
