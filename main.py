import os  # to access the .env file
import time  # to pause the script
import tempfile # to create temporary files
import re # to use regular expressions
import io # to use input/output operations
import gspread    # to access google sheet
import PyPDF2   # to read pdf files
import json # to use json data
from groq import Groq # to use groq api 
from oauth2client.service_account import ServiceAccountCredentials # to use oauth2client for google sheets
from googleapiclient.discovery import build # to use googleapis
from google.oauth2 import service_account # to use googleapis
from googleapiclient.http import MediaIoBaseDownload # to use googleapis
import google.generativeai as genai # to use gemini api
from dotenv import load_dotenv # to load .env file
from datetime import datetime # to get the current time
import smtplib # to use smtplib for sending emails
from email.mime.text import MIMEText # to use mimedtext for sending emails
from email.mime.multipart import MIMEMultipart # to use mimemultipart for sending emails
from email.mime.base import MIMEBase # to use mimedbase for sending emails
from email import encoders # to use encoders for sending emails

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

CREDENTIALS_FILE   = get_env_var("CREDENTIALS_FILE", os.path.join(SCRIPT_DIR, "Credentials.json"))

SHEET_URL_OR_ID    = get_env_var("GOOGLE_SHEET_ID")
RESULTS_SHEET_NAME = get_env_var("RESULTS_SHEET_NAME", "Ai Responses")
SOURCE_SHEET_NAME  = get_env_var("SOURCE_SHEET_NAME")  # Tab with candidate CV form responses
DRIVE_URL_COLUMN   = get_env_var("DRIVE_URL_COLUMN", "Upload Your Resume")
JOB_ROLE_COLUMN    = get_env_var("JOB_ROLE_COLUMN", "Job Role")
NAME_COLUMN        = get_env_var("NAME_COLUMN", "Name")
EMAIL_COLUMN       = get_env_var("EMAIL_COLUMN", "Email Address")
PHONE_COLUMN       = get_env_var("PHONE_COLUMN", "Phone Number")
STATUS_COLUMN      = get_env_var("STATUS_COLUMN", "Status")
GROQ_API_KEY       = get_env_var("GROQ_API_KEY")
GROQ_MODEL         = get_env_var("GROQ_MODEL", "llama-3.1-8b-instant")
GROQ_MODEL_BACKUP  = get_env_var("GROQ_MODEL_BACKUP", "llama-3.1-70b-versatile")
POLL_INTERVAL      = int(get_env_var("POLL_INTERVAL", "10"))

# Gemini Configuration (Backup)
GEMINI_API_KEY     = get_env_var("GEMINI_API_KEY")
GEMINI_MODEL       = get_env_var("GEMINI_MODEL", "gemini-1.5-flash")

# Email Configuration
SENDER_EMAIL     = get_env_var("SENDER_EMAIL")
SENDER_PASSWORD  = get_env_var("SENDER_PASSWORD")
HR_EMAIL         = get_env_var("HR_EMAIL")
SMTP_SERVER      = get_env_var("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT        = int(get_env_var("SMTP_PORT", "587"))
SHORTLISTED_FOLDER_ID = get_env_var("SHORTLISTED_FOLDER_ID")

# ==========================================
# Your Custom AI Prompt
# ==========================================
USER_PROMPT_TEMPLATE = """
You are a senior HR analyst and resume intelligence engine.
Your task: parse a raw resume, evaluate it against a specific job role, and return a COMPLETE structured analysis — regardless of whether the candidate is suitable or not.

## Important 
1. FIND THE CERTIFICATE (IF THERE) FROM THIS

══ JSON OUTPUT CONTRACT — HIGHEST PRIORITY ══
Your entire response must be one valid JSON object and nothing else.

RULES THAT MUST NEVER BE BROKEN:
  1. Start your response with {{ and end with }}. No text before or after.
  2. Every single key in the schema below MUST be present in your output.
  3. Never skip, rename, or reorder any key.
  4. Never return plain text, an apology, an explanation, or a markdown block.
  5. If the resume text is empty, garbled, or unreadable → still return the full JSON schema
     with null / [] for unparseable fields, score: 0, recommendation: "Not Recommended",
     and score_reason: "Resume text could not be parsed."
  6. String values: use null if genuinely absent. Never use "" as a placeholder.
  7. Array values: use [] if empty. Never use null for an array field.
  8. Number values: must be actual numbers (0, 1.5, 75). Never wrap in quotes.
  9. No trailing commas. No comments inside JSON. No extra keys beyond the schema.
  10. Validate your JSON mentally before outputting — every bracket and quote must close.
  11. Every time you run this or read this you need to generate a unquie id for each and every cv you read or how many time you executed and if it is number increment it every time it must be unqiue every time it must be different every time it should not be same . 

══ DOMAIN CLASSIFICATION GATE — RUN BEFORE ANY SCORING ══

STEP 1 — Identify the candidate's PRIMARY domain from their job titles, education, and core experience.
STEP 2 — Identify the target role's PRIMARY domain from the Job Role input.
STEP 3 — Ask: "Are these the same professional domain?"

Domain examples (these do NOT overlap):
  • Law / Legal          → lawyer, advocate, legal counsel, paralegal, LLB
  • Medicine             → doctor, surgeon, nurse, pharmacist, MBBS
  • Engineering / Tech   → software engineer, data scientist, DevOps, network engineer, CS degree
  • Finance / Accounting → accountant, financial analyst, CA, auditor, banker
  • Education            → teacher, professor, curriculum designer
  • Design / Creative    → graphic designer, UX/UI, illustrator, animator
  • Hospitality          → chef, hotel manager, restaurant supervisor
  • Construction / Civil → civil engineer, site supervisor, architect (non-software)
  • Sales / Marketing    → sales executive, digital marketer, brand manager
  • HR / Admin           → HR manager, recruiter, office administrator

DOMAIN MISMATCH = the candidate's domain and the job role's domain are clearly different fields.
Examples of mismatch:
  • Computer engineer applying for Lawyer → MISMATCH
  • Chef applying for Software Developer  → MISMATCH
  • Doctor applying for Financial Analyst → MISMATCH
  • Graphic designer applying for Nurse   → MISMATCH

══ WHAT TO DO WHEN DOMAIN IS MISMATCHED ══
IMPORTANT: Even mismatched candidates must receive a FULL, COMPLETE response.
Do NOT reduce the output. Do NOT skip fields. Fill EVERYTHING as follows:

  SCORE:
  → Score reflects resume quality (completeness, clarity, professionalism).
  → Do NOT artificially lower the score just because it is a domain mismatch. A highly impressive CV in the wrong field should still get a high score (e.g. 70-95) based on its raw quality.

  RECOMMENDATION:
  → Must be exactly "Not Recommended". No exceptions.

  SCORE_REASON:
  → Must contain 3 sentences:
      Sentence 1: State the domain mismatch explicitly.
        Example: "Candidate holds a Computer Engineering background and has applied for a
        Lawyer role — these are fundamentally different professional domains."
      Sentence 2: State what the role actually requires.
        Example: "A lawyer role requires legal qualifications, bar certification, knowledge
        of statutes, and courtroom experience."
      Sentence 3: State what the candidate has instead.
        Example: "The candidate's profile consists of programming, software development,
        and technical skills which are not applicable to legal practice."

  REQUIRED_SKILLS_FOR_ROLE:
  → List 5–10 skills genuinely required by the TARGET job role.
  → Base this on the job role input — not the candidate's background.
  → Example for "Lawyer": ["Legal research", "Contract drafting", "Litigation",
    "Knowledge of civil/criminal law", "Bar certification", "Client counselling",
    "Case preparation", "Court representation"]

  MISSING_SKILLS:
  → Compare required_skills_for_role against the candidate's extracted skills.
  → missing_skills = every required skill NOT found in the candidate's skill list.
  → For a fully mismatched candidate, this will typically equal all required skills.
  → Never leave this empty for a mismatched candidate.

  SKILL_MATCH_PERCENTAGE:
  → Calculate honestly: round((matched required skills / total required skills) × 100).
  → Soft skills (communication, teamwork) do NOT count as domain matches.
  → For a fully mismatched candidate this will typically be 0–10.

  ALL OTHER FIELDS (name, email, phone, experience, education, skills, etc.):
  → Extract normally from the resume. These fields are about the candidate, not the role.
  → Always populate them fully. A mismatch does not mean we skip personal data.

══ SCORING RULES (domain-aligned candidates) ══
Score 0–100 as an integer. Evaluate these five dimensions equally:
  1. Completeness      — all resume sections present and filled
  2. Clarity           — professional language, logical structure
  3. Skill relevance   — candidate skills vs role requirements
  4. Experience depth  — years, seniority, impact, measurable results
  5. Professionalism   — tone, consistency, no red flags

══ RECOMMENDATION RULES (domain-aligned candidates) ══
  score >= 70  →  "Recommended"
  score <= 69  →  "Not Recommended"
  domain mismatch (any score) → "Not Recommended"  ← overrides all thresholds

══ FIELD-LEVEL RULES (all candidates) ══
  name, email, phone, location, linkedin
    → Extract exactly as written. null if not present.

  summary
    → 2–3 sentences, third person, professional tone.
    → Describe who the candidate is and what they bring — not the job role.
    → Never null. Write "Insufficient information to generate a summary." if resume is empty.

  total_experience_years
    → Sum all role durations. Round to 1 decimal. 0 if no experience found.

  skills
    → Extract ALL skills mentioned: technical tools, frameworks, soft skills, domain knowledge.
    → Even for mismatched candidates, list what they actually have.

  experience[]
    → company, title, duration (string e.g. "Jan 2021 – Mar 2023"), years (number e.g. 2.2),
      responsibilities (array of strings, each a brief bullet point).
    → If no experience found: use [].

  education[]
    → degree, field, institution, year — all strings. null if missing.
    → If no education found: use [].

  certifications, languages
    → Arrays of strings. [] if none found.

  projects[]
    → name and description for each. [] if none found.

  score
    → Integer 0–100. Always present. Never null.
 

  score_reason
    → Always 3 sentences. Never null. Never empty.
    → Mismatched: follow the 3-sentence mismatch template above.
    → Aligned: cite specific strengths, gaps, and justification for the score given.

  recommendation
    → Exactly one of: "Recommended" | "Consider" | "Not Recommended"
    → Never null. Never a custom string.

  required_skills_for_role
    → 5–10 skills the TARGET JOB ROLE requires. Based on the role, not the candidate.
    → Always populated. Never [].

  missing_skills
    → Required skills not found in candidate's skills list.
    → Always populated for mismatched candidates. [] only if candidate meets all requirements.

  skill_match_percentage
    → Integer 0–100. round((matched / total required) × 100).
    → Soft skills alone cannot produce a match above 15 for mismatched candidates.

══ FINAL SELF-CHECK BEFORE OUTPUT ══
Before writing your response, verify:
  ☑ Does my output start with {{ and end with }}?
  ☑ Are all schema keys present?
  ☑ Are all brackets and quotes properly closed?
  ☑ Is score a plain integer (not a string)?
  ☑ Is recommendation one of the three exact allowed values?
  ☑ Is required_skills_for_role populated (never [])?
  ☑ Is missing_skills accurate and populated for any gaps?
  ☑ Is score_reason exactly 3 sentences? 
  ☑ For a mismatched candidate: is recommendation "Not Recommended" (regardless of how high the score is)?
  ☑ Is there ANY text outside the JSON? (If yes → remove it.)
  ☑ Is "ID" set to the exact integer from the Execution ID input? (If no → fix it.)

══ JSON SCHEMA ══
{{
  "ID": null,
  "name": null,
  "email": null,
  "phone": null,
  "location": null,
  "linkedin": null,
  "summary": null,
  "total_experience_years": 0,
  "skills": [],
  "experience": [
    {{
      "company": null,
      "title": null,
      "duration": null,
      "years": 0,
      "responsibilities": []
    }}
  ],
  "education": [
    {{
      "degree": null,
      "field": null,
      "institution": null,
      "year": null
    }}
  ],
  "certifications": [],
  "languages": [],
  "projects": [
    {{
      "name": null,
      "description": null
    }}
  ],
  "score": 0,
  "score_reason": null,
  "recommendation": null,
  "required_skills_for_role": [],
  "missing_skills": [],
  "skill_match_percentage": 0
}}

══ INPUTS ══
Execution ID: {execution_id}
Job Role: {job_role}
Resume Text: 
{cv_text}
"""
# ==========================================


def get_google_services():
    """Authenticates and returns Google Sheets and Drive service objects."""
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.readonly',
        'https://www.googleapis.com/auth/drive' # Full access for moving files
    ]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scopes)
        gc = gspread.authorize(creds)
        drive_creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
        drive_service = build('drive', 'v3', credentials=drive_creds)
        print("Authenticated with Google successfully.")
        return gc, drive_service
    except Exception as e:
        print(f"Error authenticating with Google: {e}") 
        return None, None   


def extract_file_id_from_url(url):
    """Extracts a Google Drive File ID from a sharing URL."""
    match = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    match_open = re.search(r'id=([a-zA-Z0-9_-]+)', url)
    if match: return match.group(1)
    elif match_open: return match_open.group(1) 
    return None


def download_file_from_drive(drive_service, file_id, output_path):
    """Downloads a file from Google Drive."""
    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.FileIO(output_path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return True
    except Exception as e:
        print(f"Error downloading file: {e}")
        return False


def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF."""
    text = ""
    try:
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted: text += extracted + "\n"
        return text.strip()
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return ""


def process_with_groq(cv_text, job_role, execution_id, model_to_use=None):
    """Sends CV text to Groq and expects a JSON response."""
    model = model_to_use or GROQ_MODEL
    try:
        client = Groq(api_key=GROQ_API_KEY)
        full_prompt = USER_PROMPT_TEMPLATE.format(
            execution_id=execution_id,
            job_role=job_role,
            cv_text=cv_text
        )
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a specialized JSON-only HR assistant. Output ONLY valid JSON."},
                {"role": "user",   "content": full_prompt}
            ],
            temperature=0.1,
            response_format={"type": "json_object"} 
        )
        content = response.choices[0].message.content
        return json.loads(content) 
    except Exception as e:
        print(f"Warning: Error with model {model}: {e}")
        return None


def process_with_gemini(cv_text, job_role, execution_id):
    """Fallback: Sends CV text to Google Gemini and returns JSON."""
    if not GEMINI_API_KEY:
        print("Warning: Gemini API Key missing. Skipping backup.")
        return None
        
    # Try standard models in order
    models_to_try = [GEMINI_MODEL, "gemini-1.5-flash", "gemini-pro"]
    
    for model_name in models_to_try:
        try:
            print(f"   Trying Gemini model: {model_name}...")
            genai.configure(api_key=GEMINI_API_KEY)
            model = genai.GenerativeModel(model_name)
            
            full_prompt = USER_PROMPT_TEMPLATE.format(
                execution_id=execution_id,
                job_role=job_role,
                cv_text=cv_text
            )
            
            response = model.generate_content(
                full_prompt,
                generation_config={"response_mime_type": "application/json"}
            )
            return json.loads(response.text)
        except:
            continue # Try next model if this one fails
            
    return None


def send_email(to_email, subject, body, attachment_path=None):
    """Sends a professional email using SMTP with optional attachment."""
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        print("⚠️ Email credentials missing in .env. Skipping email.")
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        if attachment_path and os.path.exists(attachment_path):
            filename = os.path.basename(attachment_path)
            with open(attachment_path, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {filename}")
            msg.attach(part)

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        return False


def move_file_in_drive(drive_service, file_id, new_folder_id):
    """Moves a file in Google Drive to a new folder."""
    if not new_folder_id:
        return False
    try:
        # Retrieve the existing parents to remove them
        file = drive_service.files().get(fileId=file_id, fields='parents').execute()
        previous_parents = ",".join(file.get('parents'))
        # Move the file to the new folder
        file = drive_service.files().update(fileId=file_id,
                                            addParents=new_folder_id,
                                            removeParents=previous_parents,
                                            fields='id, parents').execute()
        return True
    except Exception as e:
        print(f"Failed to move file in Drive: {e}")
        return False


def handle_notifications(candidate_data, job_role, cv_file_id=None, local_pdf_path=None):
    """Triggers specific emails and Drive actions based on AI recommendation."""
    name = candidate_data.get("name")
    email = candidate_data.get("email")
    rec = candidate_data.get("recommendation", "")
    score = candidate_data.get("score", 0)

    if rec == "Recommended":
        # 1. Email to Candidate
        subject_can = f"Interview Invitation: {job_role} Role"
        body_can = f"Dear {name},\n\nCongratulations! Your profile has been shortlisted for the {job_role} position.Your skills, attitude, and enthusiasm stood out throughout the process, and we are confident you will be a valuable addition to our organization. Our HR team will contact you shortly to schedule an interview. \n Please Book your Interview Slot from Here \n https://calendly.com/ggyantalks/interview \n\nBest regards,\nRecruitment Team"
        send_email(email, subject_can, body_can)
        print(f"Interview invite sent to {name}.")

        # 2. Email to HR (with CV attachment)
        if HR_EMAIL:
            subject_hr = f"New Shortlisted Candidate: {name}"
            body_hr = f"A new candidate has been shortlisted for the {job_role} role.\n\nName: {name}\nScore: {score}/100\nEmail: {email}\n\nPlease check the AI Responses sheet for full details. The CV is attached."
            send_email(HR_EMAIL, subject_hr, body_hr, attachment_path=local_pdf_path)
            print(f"HR notification (with CV) sent for {name}.")

        # 3. Move CV in Google Drive
        if cv_file_id and SHORTLISTED_FOLDER_ID:
            if move_file_in_drive(drive_service, cv_file_id, SHORTLISTED_FOLDER_ID):
                print(f"CV moved to Shortlisted Folder.")

    elif rec == "Not Recommended":
        subject_can = f"Update regarding your application for {job_role}"
        body_can = f"Dear {name},\n\nThank you for your interest in the {job_role} position. After reviewing your profile, we regret to inform you that we will not be moving forward with your application at this time. We wish you the best in your career search.\n\nBest regards,\nRecruitment Team"
        send_email(email, subject_can, body_can)
        print(f"Rejection email sent to {name}.")


def ensure_results_sheet(spreadsheet):
    """Ensures the Results tab exists and has headers for ALL fields."""
    try:
        try:
            results_sheet = spreadsheet.worksheet(RESULTS_SHEET_NAME)
        except gspread.WorksheetNotFound:
            print(f"Creating new tab: '{RESULTS_SHEET_NAME}'")
            results_sheet = spreadsheet.add_worksheet(title=RESULTS_SHEET_NAME, rows="100", cols="25")
        
        headers = results_sheet.row_values(1)
        if not headers:
            # Expanded headers to include almost every field from the AI JSON
            expected_headers = [
                "ID", "Name", "Email", "Phone", "Location", "LinkedIn", "Score", "Recommendation", 
                "Match %", "Job Role", "Summary", "Experience Years", "Skills", 
                "Missing Skills", "Required Skills", "Score Reason", 
                "Full Experience", "Full Education", "Certifications", "Languages", "Projects",
                "Processed At",
                "Interview Result", "Interviewer", "Interview Date"
            ]
            results_sheet.append_row(expected_headers)
            print(f"Full headers created in '{RESULTS_SHEET_NAME}'")
        
        return results_sheet
    except Exception as e:
        print(f"Error preparing results sheet: {e}")
        return None


def format_array_of_objects(arr):
    """Helper to convert complex AI objects to a readable string."""
    if not arr or not isinstance(arr, list): return ""
    lines = []
    for obj in arr:
        if isinstance(obj, dict):
            # Format experience/education into a readable line
            parts = [f"{v}" for k, v in obj.items() if v]
            lines.append(" | ".join(parts))
        else:
            lines.append(str(obj))
    return "\n".join(lines)


def process_sheet(source_sheet, results_sheet):
    """Main processing logic."""
    records = source_sheet.get_all_records()
    source_headers = source_sheet.row_values(1)

    # DEBUG: Print all headers to help the user match .env variables
    print(f"Detected Columns in Sheet: {', '.join(source_headers)}")

    def find_val_in_record(record, target_header):
        """Helper to find a value even if there are slight spelling/space differences."""
        if not target_header: return None
        target = target_header.strip().lower()
        for key, val in record.items():
            if key.strip().lower() == target:
                return val
        return None

    # Find Status column using fuzzy matching (handles trailing spaces, case differences)
    status_col_index = None
    for idx, h in enumerate(source_headers):
        if h.strip().lower() == STATUS_COLUMN.strip().lower():
            status_col_index = idx + 1
            break

    if status_col_index is None:
        # Status column does not exist — add it
        status_col_index = len(source_headers) + 1
        source_sheet.update_cell(1, status_col_index, STATUS_COLUMN)
        source_headers.append(STATUS_COLUMN)
        print(f"Status column added at position {status_col_index}.")
    else:
        print(f"Status column found at position {status_col_index}.")

    if not any(h.strip().lower() == DRIVE_URL_COLUMN.strip().lower() for h in source_headers):
        print(f"Column '{DRIVE_URL_COLUMN}' not found. Check .env")
        return 0

    processed_count = 0
    for i, record in enumerate(records):
        row_number = i + 2

        # Read status safely - always a string, never None
        raw_status = find_val_in_record(record, STATUS_COLUMN)
        status = str(raw_status).strip() if raw_status is not None else ""

        # Skip rows already done or being handled by another instance
        if status == "Processed" or status == "Processing" or status.startswith("Error"):
            continue

        print(f"\nProcessing submission at row {row_number}...")

        # --- LOCK THE ROW IMMEDIATELY ---
        # Write "Processing" before doing any work so that if multiple instances
        # of main.py are running, they will skip this row and not send duplicate emails.
        source_sheet.update_cell(row_number, status_col_index, "Processing")

        cv_url = find_val_in_record(record, DRIVE_URL_COLUMN)
        job_role = find_val_in_record(record, JOB_ROLE_COLUMN) or "General"

        # Get data directly from the form row
        form_name = find_val_in_record(record, NAME_COLUMN) or "Unknown"
        form_email = find_val_in_record(record, EMAIL_COLUMN) or "N/A"
        form_phone = find_val_in_record(record, PHONE_COLUMN) or "N/A"

        file_id = extract_file_id_from_url(cv_url)
        if not file_id:
            source_sheet.update_cell(row_number, status_col_index, "Error - Invalid Link")
            continue

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            temp_pdf_path = tmp.name

        if download_file_from_drive(drive_service, file_id, temp_pdf_path):
            cv_text = extract_text_from_pdf(temp_pdf_path)
            
            if cv_text:
                # --- PREVENT 413 ERROR (Request too large) ---
                # Truncate to ~8,000 characters (very safe limit for free tier)
                if len(cv_text) > 8000:
                    print(f"Warning: CV is very large ({len(cv_text)} chars). Truncating to 8,000 chars for AI...")
                    cv_text = cv_text[:8000] + "\n[...Text truncated for length...]"

                execution_id = int(time.time() + row_number)
                print(f"AI Analysis for: {job_role} (Model: {GROQ_MODEL})...")
                ai_data = process_with_groq(cv_text, job_role, execution_id)

                # --- FALLBACK TO GEMINI ---
                if not ai_data:
                    print(f"Groq failed. Switching to Gemini Backup...")
                    ai_data = process_with_gemini(cv_text, job_role, execution_id)

                if ai_data:
                    # --- AI MISTAKE OVERRIDE ---
                    # Sometimes the smaller AI models forget the rules and mark a low score as "Recommended".
                    # This code forces the recommendation to align with the score strictly.
                    try:
                        score_val = int(ai_data.get("score", 0))
                        if score_val < 70 and ai_data.get("recommendation") == "Recommended":
                            print(f"Warning: AI gave score {score_val} but said Recommended. Overriding to Not Recommended.")
                            ai_data["recommendation"] = "Not Recommended"
                    except (ValueError, TypeError):
                        pass

                    # Map JSON to sheet columns
                    row_data = [
                        ai_data.get("ID"),
                        form_name,   # Use name from Form
                        form_email,  # Use email from Form
                        form_phone,  # Use phone from Form
                        ai_data.get("location"),
                        ai_data.get("linkedin"),
                        ai_data.get("score"),
                        ai_data.get("recommendation"),
                        ai_data.get("skill_match_percentage"),
                        job_role,
                        ai_data.get("summary"),
                        ai_data.get("total_experience_years"),
                        ", ".join(ai_data.get("skills", [])),
                        ", ".join(ai_data.get("missing_skills", [])),
                        ", ".join(ai_data.get("required_skills_for_role", [])),
                        ai_data.get("score_reason"),
                        format_array_of_objects(ai_data.get("experience")),
                        format_array_of_objects(ai_data.get("education")),
                        ", ".join(ai_data.get("certifications", [])),
                        ", ".join(ai_data.get("languages", [])),
                        format_array_of_objects(ai_data.get("projects")),
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "",  # Interview Result  — filled later by HR.py
                        "",  # Interviewer      — filled later by HR.py
                        ""   # Interview Date   — filled later by HR.py
                    ]
                    results_sheet.append_row(row_data)
                    source_sheet.update_cell(row_number, status_col_index, "Processed")
                    print(f"All information added to '{RESULTS_SHEET_NAME}'")
                    
                    # --- Send Automated Emails ---
                    print(f"Sending automated emails and managing files...")
                    # Inject data for the notification handler
                    ai_data['name'] = form_name
                    ai_data['email'] = form_email
                    handle_notifications(
                        ai_data, 
                        job_role, 
                        cv_file_id=file_id, 
                        local_pdf_path=temp_pdf_path
                    )
                    
                    if os.path.exists(temp_pdf_path):
                        os.remove(temp_pdf_path)  # Clean up after email is sent
                    
                    processed_count += 1
                else:
                    source_sheet.update_cell(row_number, status_col_index, "Error - AI Failed")
                    if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
            else:
                source_sheet.update_cell(row_number, status_col_index, "Error - No Text")
                if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
        else:
            source_sheet.update_cell(row_number, status_col_index, "Error - Download Failed")
            if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
        
        time.sleep(2)
    return processed_count


def main():
    if not GROQ_API_KEY or not SHEET_URL_OR_ID:
        print("Check .env for API Key and Sheet ID.")
        return

    global drive_service
    gc, drive_service = get_google_services()
    if not gc: return

    # --- Open the Google Spreadsheet ---
    try:
        print(f"Opening Spreadsheet (ID/URL: {SHEET_URL_OR_ID})...")
        if "http" in SHEET_URL_OR_ID:
            spreadsheet = gc.open_by_url(SHEET_URL_OR_ID)
        else:
            spreadsheet = gc.open_by_key(SHEET_URL_OR_ID)

        # ------------------------------------------------------------------
        # STARTUP: List all tabs so misconfiguration is immediately visible
        # ------------------------------------------------------------------
        print("\n" + "=" * 60)
        print("ALL TABS IN THIS SPREADSHEET:")
        print("=" * 60)
        for ws in spreadsheet.worksheets():
            all_rows = ws.get_all_values()
            hdrs     = all_rows[0] if all_rows else []
            d_cnt    = len([r for r in all_rows[1:] if any(c.strip() for c in r)]) if len(all_rows) > 1 else 0
            marker   = "  <-- main.py will read THIS tab" if ws.title == SOURCE_SHEET_NAME else ""
            if not SOURCE_SHEET_NAME and ws == spreadsheet.sheet1:
                marker = "  <-- main.py will read THIS tab (sheet1 fallback)"
            print(f"\n  Tab  : \"{ws.title}\"{marker}")
            print(f"  Rows : {d_cnt}")
            print(f"  Cols : {hdrs}")
        print("\n" + "=" * 60)
        if SOURCE_SHEET_NAME:
            print(f"SOURCE_SHEET_NAME in .env = \"{SOURCE_SHEET_NAME}\"")
        else:
            print("SOURCE_SHEET_NAME not set in .env — using sheet1 (first tab).")
            print("If the candidate CV form uses a different tab, set SOURCE_SHEET_NAME in .env.")
        print("=" * 60 + "\n")
        # ------------------------------------------------------------------

        # Connect to the correct source sheet
        if SOURCE_SHEET_NAME:
            try:
                source_sheet = spreadsheet.worksheet(SOURCE_SHEET_NAME)
                print(f"Source sheet connected: '{SOURCE_SHEET_NAME}'")
            except gspread.WorksheetNotFound:
                print(f"ERROR - Tab '{SOURCE_SHEET_NAME}' not found. Check SOURCE_SHEET_NAME in .env.")
                print("Available tabs are listed above. Falling back to sheet1.")
                source_sheet = spreadsheet.sheet1
        else:
            source_sheet = spreadsheet.sheet1
            print(f"Source sheet connected: sheet1 (first tab).")

        # Prepare the results sheet
        results_sheet = ensure_results_sheet(spreadsheet)
        if not results_sheet:
            print("Failed to create or open the Results sheet.")
            return

    except Exception as e:
        print(f"Error opening spreadsheet: {e}")
        return

    print(f"\nPolling every {POLL_INTERVAL}s. Press Ctrl+C to stop.\n")

    while True:
        try:
            count = process_sheet(source_sheet, results_sheet)
            if count > 0: print(f"[{datetime.now().strftime('%H:%M:%S')}] Done.")
        except Exception as e:
            print(f"Warning: Error: {e}")
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped.")
