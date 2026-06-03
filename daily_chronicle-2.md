# Daily Chronicle Planning Document (Implementation-Based)

**Tool:** Cabit  
**Feature Name:** Daily Chronicle  
**Document Purpose:** Detailed specification of the Daily Chronicle reporting function as currently implemented in the Android application.

---

## 1. Overview
The **Daily Chronicle** is an automated reporting function within Cabit. It aggregates user activity and system modifications that occurred since midnight (00:00:00) of the current day into a comprehensive, chronological report.

## 2. Process & Workflow

The Daily Chronicle function executes a multi-step process to compile activity data:

### Step 1: Document & Spreadsheet Creations
* **Target:** All files created since midnight of the current day.
* **Action:** Scans Google Drive for new files.
* **Exclusions:** 
    * Files located in the designated `logs` folder (handled in Step 2).
    * Core system files (`TasksAndLists`, `Calendar Events`, `Preferences`, `LogCategories`).
* **Data Captured:**
  * Time of creation.
  * File Title.
  * File Type (labeled as "Document" or "Spreadsheet").
  * Link to the file.

### Step 2: Spreadsheet Updates (Logging)
* **Target:** The designated `logs` folder.
* **Action:** Identifies any spreadsheets updated today. 
* **Data Captured:**
  * Parses each updated spreadsheet for rows matching today's date.
  * Utilizes a robust parser (`parseDateTimeAndText`) supporting `MM-dd-yyyy HH:mm:ss` (and variations with `/` or `.`).
  * Records the exact timestamp and the body text of the log entry.
  * Formats as: `Log:[Category] - [Body]`.

### Step 3: Tasks and Lists Parsing
* **Target:** The `TasksAndLists` system file.
* **Action:** Parses all categories and items for activity matching today's date.
* **Data Captured:**
  * **Created:** Task creation date matches today.
  * **Updated (Pending):** Task move-to-pending date matches today.
  * **Completed:** Task completion date matches today.
* **Formatting Rule:** Each action (creation, update, completion) is recorded as a distinct chronological entry.

### Step 4: Calendar Parsing
* **Target:** The `Calendar Events` system file.
* **Action:** Parses the JSON array of events for status changes occurring today.
* **Data Captured:**
  * **Created:** `created_at` timestamp matches today.
  * **Completed:** `completed_at` timestamp matches today.
  * **Cancelled:** `cancelled_at` timestamp matches today.

### Step 5: Advanced Document Parsing (Deep Scan)
* **Target:** Modified Google Documents (`application/vnd.google-apps.document`) that are not handled by previous steps.
* **Action:** Performs a "deep scan" of the document content.
* **Data Captured:**
  * **Location Notes:** Extracts address information from "Location-" titled documents.
  * **Linked Notes:** Identifies associations to specific Task/List items.
  * **Timestamped Entries:** Searches for internal "Date: YYYY-MM-DD HH:mm" patterns within the document text to identify specific recorded events.

---

## 3. Presentation & Interaction

* **Visual Report:** Rendered in a dark-themed WebView (`#121212` background).
    * **Time:** Highlighted in Green (`#00FF00`) and sorted chronologically (24-hour normalized).
    * **Description:** Includes clickable Blue (`#00AAFF`) links to the source files on Google Drive.
* **Clipboard Integration:** A "Copy to Clipboard" button provides a plain-text version of the entire chronicle, including timestamps and descriptions, for easy sharing or external archiving.

---
*End of specification.*
