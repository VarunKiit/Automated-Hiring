# Automated-Hiring
Automates Outlook email parsing, candidate info extraction, and interview scheduling using LLaMA 3.1. Features include fetching emails, extracting data from attachments, detecting missing fields, sending requests, generating timetables, proposing candidate schedules, and saving everything to Excel with a Jupyter widgets UI.
Here’s a detailed description of your **Outlook & Interview Scheduling Automation** code suitable for a GitHub README or documentation:

---

# Outlook Email Parser & Interview Scheduling Automation

This Python project provides a fully automated system for fetching emails from Outlook, extracting candidate details, maintaining an Excel-based tracker, and scheduling interviews using **LLM-based reasoning**. It combines Microsoft Outlook COM automation, Excel handling via `pandas`, and OpenAI/LLama 3.1 LLMs for parsing and scheduling.

---

## Features

1. **Outlook Email Integration**

   * Fetches emails from a specific mailbox folder (e.g., `"your_folder"`) with optional sender and subject filters.
   * Parses both email body and attachments (`.pdf`, `.docx`, `.txt`) to extract textual content.
   * Uses `win32com.client` to interact with Outlook.

2. **Candidate Information Extraction**

   * Leverages LLM (LLaMA 3.1) to parse emails and extract structured candidate information in JSON format.
   * Extracted fields include:

     * `candidate_name`, `phone_number`, `email`
     * `notice_period`, `last_working_date`
     * `current_location`, `preferred_location`
     * `total_experience`, `appian_experience`
     * `certification`, `current_ctc`, `expected_ctc`
   * Handles missing fields by setting them to `null` and allows sending a **missing details request email**.

3. **Excel Tracking**

   * Saves candidate data into `candidate_tracker.xlsx`.
   * Saves confirmed interview schedules into `interview_schedule.xlsx`.
   * Automatically appends to existing Excel files or creates new ones if they do not exist.

4. **Missing Info Handling**

   * Automatically detects which required fields are missing.
   * Sends a polite **email request** to the candidate for missing details.

5. **Interview Timetable Generation**

   * Generates a complete interview timetable based on:

     * Panel member availability
     * Slot duration
     * Strict scheduling constraints (no overlaps, 2 interviewers per slot, parallel interviews allowed only if no conflicts)
   * Uses LLM reasoning to propose optimal timetable.
   * Timetable output is in **JSON format** for easy parsing and modifications.

6. **Timetable Updates**

   * LLM-assisted updates for timetable modifications.
   * Users can request specific changes (e.g., "Nathan only 11 to 12") and LLM returns a valid updated timetable.

7. **Candidate Scheduling**

   * Assigns a candidate to a suitable interview slot based on their preferred time and panel availability.
   * LLM ensures all scheduling constraints are respected.
   * Proposed schedule is presented as JSON and can be **accepted and saved** to Excel.

8. **Interactive UI (Jupyter Widgets)**

   * Two main tabs:

     * **Update Tracker**: Fetch, parse, and save candidate data.
     * **Schedule Interview**: Generate timetables, apply changes, propose schedules, and save confirmations.
   * Input widgets:

     * Text boxes, text areas, dropdowns, buttons, and output areas.
   * Clear separation of outputs for emails, parsed candidate info, missing fields, timetable view, and proposals.

---

## Libraries & Tools Used

* **Python Core**: `os`, `io`, `re`, `json`, `datetime`, `typing`
* **Data Handling**: `pandas`, `numpy`
* **Email & Office**: `win32com.client` (Outlook automation), `python-docx`, `PyPDF2` or `pypdf`
* **LLM Integration**: `OpenAI` / LLaMA 3.1 Instruct for parsing & scheduling
* **UI**: `ipywidgets` for interactive Jupyter interface
* **Validation & Models**: `pydantic` for conversation state modeling
* **Environment Variables**: `dotenv` to manage API keys

---

## Workflow Overview

### 1. Update Tracker Tab

1. Enter sender/subject filters.
2. Fetch emails from Outlook.
3. Parse each email with LLM for candidate information.
4. Review parsed candidate info.
5. Save to Excel or request missing details.

### 2. Schedule Interview Tab

1. Enter panel availability and slot duration.
2. Generate initial timetable using LLM.
3. Apply any manual change requests.
4. Enter candidate name, email, and preferred time.
5. Propose a schedule using LLM.
6. Accept and save the schedule to Excel.

---

## Data Flow

```
Outlook Email → Email Body & Attachments → LLM Candidate Parser → ConversationState → Excel Tracker → Panel Availability → LLM Timetable → Candidate Schedule → Excel Schedule
```

---

## Key Classes & Functions

* `ConversationState`: Stores the state of the current workflow, including candidate info, timetable, and schedules.
* `search_emails(sender, subject)`: Fetch emails from Outlook folder with optional filters.
* `get_email_body_and_attachments(message)`: Extracts text from email and attachments.
* `parse_candidate_info(email_text)`: Uses LLM to extract structured candidate info.
* `calc_missing_fields(candidate_data)`: Computes missing required fields.
* `send_missing_details_email(...)`: Sends an email requesting missing information.
* `propose_timetable_llm(panel_timings, slot_duration)`: Generates initial interview timetable using LLM.
* `update_timetable_with_change(state, change_req)`: Updates timetable as per user-requested changes.
* `propose_candidate_schedule(state)`: Assigns a candidate to an interview slot respecting all constraints.
* `save_candidate_to_excel(state)` / `save_schedule_to_excel(state)`: Persist data to Excel files.

---

## Notes

* The system assumes the Outlook mailbox and folder names are correctly set.
* LLM responses are parsed from JSON code fences, with error handling for invalid JSON.
* Interactive interface is designed for **Jupyter Notebook or JupyterLab**.

---

This project is ideal for **HR automation, recruitment tracking, and intelligent interview scheduling**, reducing manual effort and leveraging LLMs for precise parsing and scheduling.

---
**Note** : 
* I have created this as part of POC, therefore have made it a button style interface, further changes can be applied by incorporating Agentc AI architecture to create a chat interface that replaces buttons style with followup questions by adding a conversational agent in the current code. You also have the flexibility of replacing prompts and better LLM Models.  
* In the current code, I have made features of candidate as static. There is also an option to make it dynamic to provide flexibility to user in creating his/her desired features.

Do you want me to create that next?
