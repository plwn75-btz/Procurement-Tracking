# Engineering Handoff Report: Procurement Tracking Dashboard & Weekly Cut-off Upload Center

**Project Name:** Procurement Tracking Dashboard (Zawtika Z1F & Aung Sinkhla ASK)  
**Date:** July 28, 2026  
**Status:** Completed, Verified, & Cloud-Ready  
**Primary Tech Stack:** Python 3.11 (`server.py`), Vanilla JavaScript (`app.js`), Modern CSS3 (`styles.css`), HTML5 (`index.html`)  

---

## 1. System Architecture Overview

The **Procurement Tracking Dashboard** is a high-performance, self-contained web application designed to track delays, milestones, and procurement stages across two major offshore development projects:
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
- **Summary KPI Cards:** Real-time metrics across **Total Procurement Items**, **Completed**, **On Track**, **At Risk**, and **Delayed**.
- **Interactive Stage Filter Bar:** Filter procurement packages by engineering stage (e.g., *PR / RFQ Preparation*, *ITB / Bidding*, *Bid Evaluation & Recommendation*, *PO / Contract Award*, *Vendor Engineering*, *Fabrication & Delivery*).
- **Lookahead & Delay Highlights:** Automatically flags packages with negative float (`Delay Days > 0`) in red (`var(--status-delayed)`) and highlights upcoming milestones due within the next 7 days (`getLookaheadEnd()`).

### 2.2. Weekly Cut-off Upload Center (`📁 Upload Center`)
- **Direct Web-Based Uploads:** Engineers click **📁 Upload Center** in the header to open a modal dialog for uploading weekly Excel workbooks without requiring command-line or FTP access.
- **Drag-and-Drop Auto-Detection:** Users can drop spreadsheet files directly onto the cloud drop zone. The frontend (`handleDroppedFiles`) scans filenames for `Z1F`/`Topside`/`Jacket` or `ASK`/`Pipeline`/`Overall` to route uploads automatically.
- **Dedicated Project Cards:** Individual cards (`WP-1 Topside & Jacket (Z1F)` and `WP-2 Overall Cycle (ASK)`) display the currently loaded source filename (`src-z1f`, `src-ask`) and allow explicit manual file selection.

### 2.3. Expanded Detail View & Table Layout Precision
- **Full-Width Package Name Banner (`.pkg-name-box`):** The `Package Name` metadata box inside the expanded view (`.detail-grid`) spans across all columns (`grid-column: 1 / -1;`), providing generous horizontal width for long 80+ character titles without squeezing adjacent metadata cards (`RFQ No.`, `MR No.`, `Priority`).
- **Strict Left-Alignment & Scoped Selector Isolation:** Outer data table (`.data-table`) rules are strictly scoped using direct child selectors (`>`) so they never leak into nested inner elements (`tr.detail-row td`). All detail cards (`.detail-item`) and inner stage table headers/cells (`.stage-table thead th`, `.stage-table tbody td`) explicitly enforce `text-align: left !important;`. Column headers like **DELAY (DAYS)** align precisely above their numerical data (`34d`, `11d`).
- **Multi-Line Word Wrapping & Overflow Resilience:** Detail boxes and values (`.detail-item .value`) enforce `white-space: normal !important; overflow: visible !important; overflow-wrap: anywhere !important; word-break: break-word !important;`. Long identifiers (`MM-ZTK-1F-ZWP21-STR-MR-001`) wrap cleanly onto additional lines and naturally expand the card height, ensuring 100% of text is visible with zero clipping or truncation.

### 2.4. Dynamic File Naming & Chronological Sorting (`YYYYMMDD`)
- **Automated Standard Naming:** When a file is uploaded, the backend extracts the date from the original filename (`re.search(r'(20\d{6})', orig_filename)`) or defaults to the current upload date.
- Files are saved as:
  - `Attachment 1-Procurement Plan-Z1F - <YYYYMMDD>.xlsx`
  - `Attachment 2-Procurement Plan-ASK - <YYYYMMDD>.xlsx`
- **Smart Date Sorting (`get_file_sort_key`):** The data extractor automatically scans `BASE_DIR`, extracts the `YYYYMMDD` date suffix, and selects the newest weekly file (`max(files, key=get_file_sort_key)`).

### 2.5. Package Float Monitoring & Dynamic Overdue PO Baseline
- **Engineering Logic & Excel Formula Mapping:**
  - **Z1F Mapping:** Column 40 (`Float `), Column 39 (`ROS Date`), Column 38 (`Delivery Date`), Column 37 (`days`), Column 36 (`weeks`). Source Excel formula: `=AM13-AL14` ($\text{ROS Date} - \text{Delivery Date}$).
  - **ASK Mapping:** Column 7/44 (`Float`), Column 40 (`ROS date`), Column 39 (`ETA date`), Column 38 (`Lead Time`), Column 37 (`weeks`). Source Excel formula: `=AN20-AM20` ($\text{ROS Date} - \text{ETA Date}$).
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
  - 🟢 **Green** (`Float >= 21 days`): Healthy schedule cushion ($\ge 3$ weeks).
- **Dedicated Table Column & Sorting:**
  - Dedicated `FLOAT` column in the primary procurement table with colored badges (`+45d Act`, `+6d Fcst`, `-12d Fcst*`).
  - Added clickable `<th>` table header sorting by float (`sortBy('float')`).
  - Added `Sort: Lowest Float (Critical First)` dropdown option to prioritize packages with negative/tight floats at the top.
- **Expanded Detail Float Card:**
  - Detail card displays the float badge, status badge, delivery lead time, estimated delivery date, ROS date, and an overdue PO warning note when the dynamic baseline is active.

---

## 3. Directory Structure & File Inventory

```text
Procurement Tracking/
├── server.py                     # Main Python HTTP server, API endpoints, & Excel data parser (float extraction)
├── index.html                    # Dashboard UI structure, Float column, Upload Center modal HTML
├── styles.css                    # Glassmorphism design tokens, animations, float badge & card styling
├── app.js                        # Client-side state, dynamic float calculation, sorting, & drag-and-drop
├── requirements.txt              # Minimal Python dependencies (pandas, openpyxl)
├── render.yaml                   # Render cloud deployment specification (PORT=10000)
├── .gitignore                    # Excludes temporary Excel lock files (~$*.xlsx) & dev scripts
├── Lesson_Learn.md               # Technical takeaways, caching gotchas, float logic, & resolutions
├── Handoff_Report.md             # Complete system architecture and handoff report
├── Attachment 1-Procurement Plan-Z1F - 20260711.xlsx # Z1F weekly source file
└── Attachment 2-Procurement Plan-ASK-20260726 R1.xlsx # ASK weekly source file (with 'New Plan' support)
```

---

## 4. Cloud Deployment & Operations Guide (Render)

### 4.1. Render Configuration (`render.yaml`)
The project is configured for one-click deployment on Render as a Web Service:
- **Build Command:** `pip install -r requirements.txt`
- **Start Command:** `python server.py`
- **Environment Variables:** `PORT=10000`, `PYTHON_VERSION=3.11.0`

### 4.2. Static Asset Caching Prevention (Crucial Operational Note)
To prevent CDNs and browsers from caching old CSS/JS files after a new deployment:
1. `server.py` includes custom headers inside `DashboardHandler.end_headers()`:
   ```python
   if any(path_clean.endswith(ext) for ext in (".html", ".css", ".js")):
       self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
   ```
2. `index.html` references `styles.css?v=20260714_1` and `app.js?v=20260714_1`. **Whenever you make major CSS or JavaScript changes in the future, bump the `?v=` version number in `index.html`** before pushing to Git.

---

## 5. Verification & Testing Checklist for Future Engineers

Before merging changes or deploying updates, verify the following:
- [ ] **Syntax Verification:** Run `python -m py_compile server.py` to ensure clean Python syntax.
- [ ] **Sorting Verification:** Run `python -c "import server; print(server.find_excel_files())"` to verify that `server.py` correctly identifies and selects the latest `YYYYMMDD` files for both Z1F and ASK.
- [ ] **Local Server Test:** Start `python server.py` and open `http://localhost:8000`.
- [ ] **Upload Center UI:** Click **📁 Upload Center** in the top right header. Verify the dark glassmorphism modal opens cleanly.
- [ ] **File Upload Execution:** Upload a sample `.xlsx` or `.xlsm` file via drag-and-drop or card button. Check that the toast notification shows `Successfully uploaded Attachment X-...` and that the dashboard figures refresh immediately without a page reload.
- [ ] **Data Extraction Validation:** If packages show as `0` after upload, verify if the source Excel file has changed its Plan indicator (e.g., from "Plan" to "New Plan", as occurred in `Attachment 2-Procurement Plan-ASK-20260726 R1.xlsx`) and update `pfa_values` in `server.py` accordingly.
- [ ] **Package Count Verification:** You can verify the extractor using the CLI: `python -c "from server import extract_all_data; data = extract_all_data(); print('Z1F Packages:', data['projects']['Z1F'].get('package_count')); print('ASK Packages:', data['projects']['ASK'].get('package_count'))"`
- [ ] **Float Calculation & Schedule Validation:** Verify that `float_days`, `ros_date`, and `delivery_duration_days` are extracted: `python -c "from server import extract_all_data; d = extract_all_data(); print('Z1F pkg 0 float:', d['projects']['Z1F']['packages'][0].get('float_days'), 'ros:', d['projects']['Z1F']['packages'][0].get('ros_date'))"`. Check that the dashboard displays the **Float** column, with red badges for negative float (< 0), yellow (< 21d), green (>= 21d), and that sorting by Lowest Float elevates critical path packages.

---

## 6. Maintenance & Future Roadmap

1. **Persistent Cloud Storage:** Render ephemeral containers reset local disk files when restarted or re-deployed. If long-term persistence of weekly Excel files is required across container restarts, mount a Render Persistent Disk volume to `BASE_DIR` or add an S3/Azure Blob upload hook inside `send_api_upload()`.
2. **Additional Project Streams:** To add a third project (e.g., `WP3 - Subsea`), simply:
   - Add a `WP3` card in `index.html` modal.
   - Extend `handleDroppedFiles` keywords (`SUBSEA`) in `app.js`.
   - Add `WP3` regex matching inside `send_api_upload()` and `find_excel_files()` in `server.py`.
3. **PowerBI Integration:** The PowerBI solution (`PowerBI_Solution/`) can continue reading from the exact same standardized `Attachment X-Procurement Plan-<Proj> - YYYYMMDD.xlsx` files using the existing M queries (`03_Fact_Procurement_ASK.m`, etc.).
