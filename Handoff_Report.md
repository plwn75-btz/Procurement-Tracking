# Engineering Handoff Report: Procurement Tracking Dashboard & Weekly Cut-off Upload Center

**Project Name:** Procurement Tracking Dashboard (Zawtika Z1F & Aung Sinkhla ASK)  
**Date:** October 09, 2026 (Updated from July 28, 2026 baseline)  
**Status:** Cut `20261002` Fully Integrated, Logic Verified, & Operational  
**Active Cut Files:**  
- `Attachment 1 - Procurement Plan-Z1F_20261002.xlsx` (142 packages, 24 stages, P/F/A in Col L)  
- `Attachment 2-Procurement Plan-ASK-20261002 R3.xlsx` (98 packages, 24 stages, P/F/A in Col K)  
- **Total Packages Ingested:** 240 packages  
**Primary Tech Stack:** Python 3.11 (`server.py`), Vanilla JavaScript (`app.js`), Modern CSS3 (`styles.css`), HTML5 (`index.html`)  

---

## 1. System Architecture Overview

The **Procurement Tracking Dashboard** is a high-performance, self-contained web application designed to track delays, milestones, float erosion, and procurement stages across two major offshore development projects:
- **Z1F (Zawtika Project Development - Topside & Jacket - WP1)**
- **ASK (Aung Sinkhla Project Development - Overall Cycle - WP2)**

### Architecture Design:
1. **Backend Service (`server.py`):**
   - Built with pure Python `http.server` (`DashboardHandler`), ensuring lightweight deployment with zero heavy WSGI framework dependencies (no Flask/Django).
   - Uses `pandas` and `openpyxl` to parse complex multi-sheet Excel workbooks (`extract_all_data()`).
   - Serves static web assets with zero-cache headers (`must-revalidate`) and exposes RESTful API endpoints:
     - `GET /api/data`: Returns full pre-parsed JSON analytics (summary KPIs, stage distribution, active items, lookahead alerts).
     - `GET /api/refresh`: Forces thread-safe re-parsing of Excel spreadsheets from disk.
     - `POST /api/upload`: Handles multipart form uploads for Z1F and ASK weekly cut-off Excel files.

2. **Frontend SPA (`index.html`, `styles.css`, `app.js`):**
   - Single-page application styled with a sleek dark glassmorphism aesthetic (`#0a0e1a` background, subtle borders, cyan/blue gradients).
   - Dynamic interactivity via `app.js`: real-time search filtering, stage filtering tabs, lookahead date windows (`Today -> +7 Days`), and interactive file drag-and-drop.

---

## 2. Core Features & Capabilities Implemented

### 2.1. Executive Dashboard & Visual Analytics
- **Summary KPI Cards:** Real-time metrics across **Total Procurement Items** (240), **Completed**, **On Track**, **At Risk**, and **Delayed**.
- **Interactive Stage Filter Bar:** Filter procurement packages by engineering stage (e.g., *PR / RFQ Preparation*, *ITB / Bidding*, *Bid Evaluation & Recommendation*, *PO / Contract Award*, *Vendor Engineering*, *Fabrication & Delivery*).
- **Lookahead & Delay Highlights:** Automatically flags packages with negative float (`Delay Days > 0`) in red (`var(--status-delayed)`) and highlights upcoming milestones due within the next 7 days (`getLookaheadEnd()`).

### 2.2. Weekly Cut-off Upload Center (`📁 Upload Center`)
- **Direct Web-Based Uploads:** Engineers click **📁 Upload Center** in the header to open a modal dialog for uploading weekly Excel workbooks without requiring command-line or FTP access.
- **Drag-and-Drop Auto-Detection:** Users can drop spreadsheet files directly onto the cloud drop zone. The frontend (`handleDroppedFiles`) scans filenames for `Z1F`/`Topside`/`Jacket` or `ASK`/`Pipeline`/`Overall` to route uploads automatically.
- **Dedicated Project Cards:** Individual cards (`WP-1 Topside & Jacket (Z1F)` and `WP-2 Overall Cycle (ASK)`) display the currently loaded source filename (`src-z1f`, `src-ask`) and allow explicit manual file selection.

### 2.3. Expanded Detail View & Table Layout Precision
- **Full-Width Package Name Banner (`.pkg-name-box`):** The `Package Name` metadata box inside the expanded view (`.detail-grid`) spans across all columns (`grid-column: 1 / -1;`), providing generous horizontal width for long 80+ character titles without squeezing adjacent metadata cards (`RFQ No.`, `MR No.`, `Priority`).
- **Strict Left-Alignment & Scoped Selector Isolation:** Outer data table (`.data-table`) rules are strictly scoped using direct child selectors (`>`) so they never leak into nested inner elements (`tr.detail-row td`). All detail cards (`.detail-item`) and inner stage table headers/cells (`.stage-table thead th`, `.stage-table tbody td`) explicitly enforce `text-align: left !important;`. Column headers like **DELAY (DAYS)** align precisely above their numerical data (`34d`, `11d`).
- **Multi-Line Word Wrapping & Overflow Resilience:** Detail boxes and values (`.detail-item .value`) enforce `white-space: normal !important; overflow: visible !important; overflow-wrap: anywhere !important; word-break: break-word !important;`. Long identifiers wrap cleanly onto additional lines and naturally expand the card height, ensuring 100% of text is visible with zero clipping or truncation.

### 2.4. Dynamic File Naming & Chronological Sorting (`YYYYMMDD`)
- **Automated Standard Naming:** When a file is uploaded, the backend extracts the date from the original filename (`re.search(r'(20\d{6})', orig_filename)`) or defaults to the current upload date.
- Files are saved as:
  - `Attachment 1-Procurement Plan-Z1F - <YYYYMMDD>.xlsx` (or raw format `Attachment 1 - Procurement Plan-Z1F_20261002.xlsx`)
  - `Attachment 2-Procurement Plan-ASK - <YYYYMMDD>.xlsx` (or raw format `Attachment 2-Procurement Plan-ASK-20261002 R3.xlsx`)
- **Smart Date Sorting (`get_file_sort_key`):** The data extractor automatically scans `BASE_DIR`, extracts the `YYYYMMDD` date suffix, and selects the newest weekly file (`max(files, key=get_file_sort_key)`).

### 2.5. Package Float Monitoring & Dynamic Overdue PO Baseline
- **Actual Float vs. Dynamic Forecast Float:**
  - **1. Actual Float (`Act` Badge):** For awarded packages where `Actual PO Date` exists:
    $$\text{Actual Delivery Date} = \text{Actual PO Date} + \text{Delivery Duration (Days)}$$
    $$\text{Actual Float} = \text{Plan ROS Date} - \text{Actual Delivery Date}$$
  - **2. Dynamic Forecast Float for Overdue POs (`Fcst*` Badge):** When a PO has not been awarded and the planned/forecast PO date has passed ($\text{Forecast PO Date} < \mathbf{Today}$), static Excel dates create false optimism. The system dynamically sets the PO award baseline to $\mathbf{Today}$:
    $$\text{Earliest Delivery Date} = \mathbf{Today} + \text{Delivery Duration (Days)}$$
    $$\mathbf{Dynamic\ Forecast\ Float} = \text{Plan ROS Date} - (\mathbf{Today} + \text{Delivery Duration (Days)})$$
    This accurately reflects real-world project risk by eroding float day-by-day until the PO is actually awarded.
  - **3. Standard Forecast Float (`Fcst` Badge):** When unissued forecast PO dates remain in the future ($\ge \mathbf{Today}$), standard forecast float is displayed:
    $$\text{Forecast Float} = \text{Plan ROS Date} - (\text{Forecast PO Date} + \text{Delivery Duration})$$
- **Three-Tier Color Thresholds:**
  - 🔴 **Red** (`Float < 0`): Critical delay / projected delivery is after the Required on Site (ROS) date.
  - 🟡 **Yellow** (`0 <= Float < 21 days`): Tight buffer / cushion is under 3 weeks (21 days).
  - 🟢 **Green** (`Float >= 21 days`): Healthy float.

---

## 3. Advanced Schedule Logic & Milestone Verification (October 2026 Release)

### 3.1. Downstream Unreached Stage Guard (FAT & Ready for Shipment)
- **Problem Statement:** In Cut `20261002`, packages like `Check Valve (Flange End)` (`RFQ-PIP-052`) showed false delays of **`288 days`** on `FAT` and **`274 days`** on `Ready for Shipment` because source Excel files contained stale late-2025/early-2026 forecast dates, while their baseline plan dates were scheduled for **February 2027**.
- **Solution:** In `computePackageStatus(pkg)` (`app.js`), any stage occurring sequentially *after* the package's active stage (`stageIdx > currentStageIdx`) that has a future planned date (`plan > today`) is guarded against historical forecast typos. It is locked to **`● Upcoming`** with `delayDays = 0`, completely preventing upstream packages from reporting premature downstream inspection delays.

### 3.2. Dual-Status Alignment: Latest Completed Execution vs. Active Forecast Slip
- **Latest Completed Milestone Status:** Main table package status (`STATUS` column) maps to the latest completed milestone (`latestCompletedStage`). When `VD Submission` finishes early or on time, the package displays **`● On Track`** (Green).
- **Active Stage Forecast Slip Tracking:** When the active milestone has slipped past plan (`forecast > plan`), the dashboard displays:
  - An amber **`+44d Fcst Slip`** pill in the main table `Current Stage` column.
  - A dedicated **`● Forecast Slip`** status badge in the stage detail table and visual pipeline.
- **Max Delay Preservation:** The `Max Delay` column (yellow tabular figures) is retained to show maximum historical or active delay (e.g., `19d` from `TBE Approved`) without downstream distortion.

### 3.3. Lookahead Due Horizon (+Nd) & Clean Integer Durations
- **Lookahead Remaining Days Horizon:** For milestones flagged as `● Due Soon` (`atrisk`), the Delay column represents days until milestone forecast from today:
  $$\text{Days Remaining} = \text{Forecast Date} - \mathbf{Today}$$
  For `Piggable Wye` (`RFQ-PLR-087`, ASK), `PO Issued` (forecast `13 Oct 26` vs Today `9 Oct 26`) correctly displays **`+4d`**.
- **Integer Duration Formatting:** Delivery durations/lead times are cleanly rounded using `Math.round()` in frontend rendering and `int(round(float(val)))` in `server.py`, ensuring lead time displays cleanly as **`90d`** (eliminating floating-point decimal artifacts like `90.00000000000003d`).

---

## 4. Directory Structure & File Inventory

```text
Procurement Tracking/
├── server.py                                      # Main Python HTTP server, API endpoints, & Excel data parser
├── index.html                                     # Dashboard UI structure, asset versioning (?v=20261009_3)
├── styles.css                                     # Glassmorphism design tokens, badges, .fcst-pill, .fcst-slip
├── app.js                                         # Client-side state, unreached stage guard, dual-status tracking
├── requirements.txt                               # Minimal Python dependencies (pandas, openpyxl)
├── render.yaml                                    # Render cloud deployment specification (PORT=10000)
├── .gitignore                                     # Excludes temporary Excel lock files (~$*.xlsx) & dev scripts
├── Lesson_Learn.md                                # Comprehensive technical lessons learned (sections 2.1 - 2.13)
├── Handoff_Report.md                              # Complete system architecture and handoff report
├── README.md                                      # Project guide, quick start, API specs, and operation rules
├── Attachment 1 - Procurement Plan-Z1F_20261002.xlsx      # Active Z1F cut (142 packages)
└── Attachment 2-Procurement Plan-ASK-20261002 R3.xlsx     # Active ASK cut (98 packages)
```

---

## 5. Cloud Deployment & Operations Guide (Render)

### 5.1. Render Configuration (`render.yaml`)
The project is configured for one-click deployment on Render as a Web Service:
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `python server.py`
- **Environment Variables:** `PORT=10000`, `PYTHON_VERSION=3.11.0`

### 5.2. Static Asset Caching Prevention (Crucial Operational Note)
To prevent CDNs and browsers from caching old CSS/JS files after an update:
1. `server.py` includes custom headers inside `DashboardHandler.end_headers()`:
   ```python
   if any(path_clean.endswith(ext) for ext in (".html", ".css", ".js")):
       self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
   ```
2. `index.html` references versioned assets (currently `styles.css?v=20261009_3` and `app.js?v=20261009_3`). **Always bump the `?v=` version string in `index.html` whenever you modify CSS or JS.**

---

## 6. Operational Policies & Project Directives

1. **No Power BI Execution Required:**
   - The application runs exclusively as a standalone Python web service and browser SPA.
   - Files in `PowerBI_Solution/` are preserved for historical reference but are not to be executed or automated.
2. **Manual Git Synchronization:**
   - Automated git commits and git pushes by AI agents are strictly forbidden per user instruction.
   - All version control commits and pushes to GitHub are performed manually by the repository owner.
3. **Weekly Cut Ingestion Protocol:**
   - Drop new Excel workbooks into the project root directory or upload them via the **📁 Upload Center** modal.
   - Verify package counts and data extraction via CLI:
     ```bash
     python -c "from server import extract_all_data; d = extract_all_data(); print('Z1F:', d['projects']['Z1F']['package_count'], 'ASK:', d['projects']['ASK']['package_count'])"
     ```

---

## 7. Verification & Testing Checklist for Future Engineers

Before handing off or deploying updates, verify the following:
- [x] **Syntax Verification:** Run `python -m py_compile server.py` to ensure clean Python syntax.
- [x] **File Selection Verification:** Ensure `server.find_excel_files()` identifies `Attachment 1 - Procurement Plan-Z1F_20261002.xlsx` and `Attachment 2-Procurement Plan-ASK-20261002 R3.xlsx`.
- [x] **Package Ingestion:** Ensure combined packages equal 240 (Z1F: 142, ASK: 98).
- [x] **Unreached Stage Guard Verification:** Verify `Check Valve (Flange End)` has `FAT` and `Ready for Shipment` marked as `● Upcoming` with `0d` delay.
- [x] **Execution vs. Forecast Status Verification:** Verify `Check Valve` main table status is `● On Track` (green), current stage shows amber `+44d Fcst Slip`, and stage table shows `● Forecast Slip`.
- [x] **Lookahead Horizon & Duration Precision:** Verify `Piggable Wye` shows Lead Time as `90d` and `PO Issued` shows `● Due Soon` with `+4d` delay.
- [x] **Local Server Execution:** Start `python server.py` and verify all tabs and views on `http://localhost:8080`.
