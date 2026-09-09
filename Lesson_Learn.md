# Lessons Learned: Procurement Tracking Dashboard & Weekly Cut-off Upload Center

**Document Date:** July 28, 2026  
**Project:** Procurement Tracking Dashboard (Zawtika Z1F & Aung Sinkhla ASK)  
**Author:** Engineering Team  

---

## 1. Executive Summary

During the implementation and cloud deployment (Render) of the **Procurement Tracking Dashboard** and its interactive **Weekly Cut-off Upload Center**, several technical challenges and edge cases were encountered and successfully resolved. This document captures the key engineering lessons learned across full-stack architecture, pure Python HTTP server customization, dynamic file handling, and browser/CDN static asset caching to guide future development and maintenance.

---

## 2. Key Lessons & Technical Takeaways

### 2.1. Cloud Deployment & Static Asset Caching (`SimpleHTTPRequestHandler` vs. Browser/CDN Cache)
- **The Challenge:** After pushing updated frontend code (`index.html`, `styles.css`, `app.js`) to Git and deploying to Render, the newly added **📁 Upload Center** button rendered without styling (white background, black text) and clicked with no action (`ReferenceError: openUploadModal is not defined`).
- **Root Cause:** By default, Python's `http.server.SimpleHTTPRequestHandler` serves static files (`.css`, `.js`, `.html`) without explicit anti-caching or versioning headers (`Cache-Control: no-cache`). Cloud proxies (Render CDN) and client browsers aggressively cache CSS/JS files. Consequently, after deployment, browsers continued executing stale cached `app.js` and `styles.css` files that lacked the newly added modal classes and functions.
- **The Solution & Best Practice:**
  1. **Server-Side Cache Control:** Override `end_headers(self)` inside `DashboardHandler` (`server.py`) to explicitly send cache-prevention headers (`Cache-Control: no-cache, no-store, must-revalidate`, `Pragma: no-cache`, `Expires: 0`) for `.html`, `.css`, and `.js` requests.
  2. **Clean Query Parameter Stripping:** In `do_GET(self)`, strip query strings (`self.path.split("?")[0]`) before routing or delegating to `super().do_GET()` so version parameters don't cause `404 Not Found` errors when mapping to local filesystem paths.
  3. **Frontend Cache-Busting Versioning:** Always append build or date version query parameters to `<link>` and `<script>` tags in `index.html` (`styles.css?v=20260712`, `app.js?v=20260712`) whenever structural changes are made.

### 2.2. Dynamic Date-Based File Naming (`YYYYMMDD`) & Multi-Project Sorting
- **The Challenge:** Source Excel files from project teams arrive with varying date suffixes (e.g., `Attachment 1-Procurement Plan-Z1F - 20260703.xlsx` or `Attachment 2-Procurement Plan-ASK - 20260711.xlsx`). The backend needed to automatically identify whether an uploaded file belonged to **Z1F** or **ASK** and store it with a standardized, date-ordered naming scheme so the dashboard always ingests the newest weekly cut-off file.
- **The Solution & Best Practice:**
  1. **Robust Date Extraction (`server.py`):** Implemented regex extraction (`re.search(r'(20\d{6})', orig_filename)` fallback to `20\d{2}[-_.]?\d{2}[-_.]?\d{2}`) to preserve original cut-off dates from user filenames when present, falling back to `datetime.datetime.now().strftime("%Y%m%d")` when no date is supplied.
  2. **Standardized Target Naming:** All uploaded spreadsheets are written to `BASE_DIR` with rigid prefixes:
     - `Attachment 1-Procurement Plan-Z1F - <YYYYMMDD>.xlsx`
     - `Attachment 2-Procurement Plan-ASK - <YYYYMMDD>.xlsx`
  3. **Lexicographical Sort Key:** In `get_file_sort_key(filename)`, regex `re.search(r'(\d{8})', fn)` extracts the 8-digit date string, ensuring reliable chronological sorting where `20260711` properly supersedes `20260703`.

### 2.3. Pure Python `multipart/form-data` Parsing Without External WSGI Dependencies
- **The Challenge:** To maintain lightweight zero-dependency deployment (or minimal requirements (`pandas`, `openpyxl`)), the backend HTTP server (`server.py`) does not use Flask or FastAPI. Handling multi-megabyte `multipart/form-data` file uploads directly inside `BaseHTTPRequestHandler.do_POST()` requires careful byte-level parsing.
- **The Solution & Best Practice:**
  - Extracted the MIME boundary from `self.headers.get("Content-Type")` (`re.search(r'boundary=([^\s;]+)', content_type)`).
  - Split incoming raw bytes (`self.rfile.read(content_length)`) by `--boundary`.
  - Safely separated part headers (`\r\n\r\n` or `\n\n`) from binary spreadsheet bytes (`file_bytes`), decoding part headers to identify `name="project"` and `filename="..."`.
  - Added a strict `50MB` content-length check (`content_length > 50 * 1024 * 1024`) to protect against memory exhaustion or denial-of-service during large uploads.

### 2.4. Real-Time In-Memory Cache Invalidation upon Upload
- **The Challenge:** Previously, uploading new files required restarting the server or waiting for a manual `/api/refresh` trigger before dashboard metrics updated.
- **The Solution & Best Practice:**
  - Synchronously call `extract_all_data()` directly inside `send_api_upload(self)` immediately upon writing the spreadsheet to disk.
  - Return the freshly computed JSON payload directly inside the upload API response (`{"success": true, "filename": "...", "data": extract_all_data()}`).
  - In `app.js`, when `fetch('/api/upload')` resolves successfully, directly assign `appState.data = result.data; renderDashboard();` to immediately reflect the new KPIs, stage counts, and filename indicators (`src-z1f`, `src-ask`) without an extra round-trip HTTP fetch.

### 2.5. Drag-and-Drop Auto-Detection & UX Resilience
- **The Challenge:** Users frequently drop files into the drop zone without specifying whether the file is Z1F (Topside & Jacket) or ASK (Overall Pipeline Cycle).
- **The Solution & Best Practice:**
  - Implemented heuristic keyword matching in `handleDroppedFiles(files)` checking `file.name.toUpperCase()`:
    - Keywords `Z1F`, `TOPSIDE`, `JACKET` → Assign as `Z1F`.
    - Keywords `ASK`, `PIPELINE`, `OVERALL` → Assign as `ASK`.
  - If a filename is completely ambiguous (e.g., `procurement_plan.xlsx`), gracefully intercept the upload and display a clear toast notification (`showToast('Could not auto-detect Z1F or ASK from filename. Please use Upload Specific File card.', 'error')`) guiding the user to click the dedicated card buttons.

### 2.6. CSS Specificity & Table Styling Inheritance (`td:last-child` Leakage into Nested Tables)
- **The Challenge:** When clicking expanded detail rows (`tr.detail-row`), the column title row (`stage-table thead th`) for `Delay (days)` was left-aligned while the data column was right-aligned, making titles and data out of line. Furthermore, all card labels (`RFQ NO.`, `MR NO.`, `PRIORITY`, etc.) and values inside the detail grid were pushed to the right-hand side (`text-align: right`).
- **Root Cause:** The outer data table (`.data-table`) used broad descendant CSS selectors (`.data-table tbody tr td:last-child` and `.data-table tbody tr td`) instead of direct child selectors (`>`). Because `<tr class="detail-row">` and its `<td colspan="7">` container are the last child cell of the outer table body, that container cell received `text-align: right`, `white-space: nowrap`, and `overflow: hidden`. Since `text-align` and `white-space` inherit down to all child `div`s, `.detail-grid`, `.detail-item`, and `.stage-table` cells inherited right-hand text alignment and no-wrapping constraints from the parent table cell.
- **The Solution & Best Practice:**
  1. **Strict Direct Child Scoping (`>`):** Always scope table structure CSS with direct child selectors (`.data-table > thead > tr > th`, `.data-table > tbody > tr > td`, `.data-table > tbody > tr > td:last-child`) so rules never leak into nested child tables or detail components.
  2. **Explicit Reset on Container Cell (`tr.detail-row > td`):** Explicitly override inherited styles on detail row containers (`text-align: left !important; white-space: normal !important; overflow: visible !important;`).
  3. **Explicit Element Alignment & `<colgroup>` Sizing:** For nested detail tables (`.stage-table`), enforce explicit `table-layout: fixed`, balanced `<colgroup>` percentages (`28%`, `15%`, `15%`, `15%`, `15%`, `12%`), and explicit `text-align: left;` on both headers and body cells (`thead th:last-child, tbody td:last-child`) so headers and data align uniformly.

### 2.7. CSS Grid Overflow & Word-Wrapping Inside Detail Cards (`overflow-wrap: anywhere`)
- **The Challenge:** Long procurement package names (often 80+ characters) or long MR numbers (`MM-ZTK-1F-ZWP21-STR-MR-001`) cut off horizontally (`...` or clipped right at `-`) instead of wrapping neatly to the next line within their detail cards.
- **Root Cause:** CSS Grid (`minmax(200px, 1fr)`) combined with `overflow: hidden` and inherited `white-space: nowrap` forced text strings containing hyphens or underscores to stay on a single line or break awkwardly right after `-`. If the string after the hyphen exceeded the remaining inner card width, `overflow: hidden` clipped it.
- **The Solution & Best Practice:**
  1. **Grid Spanning for Long Banners:** Give wide metadata boxes (`Package Name`) a dedicated class (`.pkg-name-box`) with `grid-column: 1 / -1;` so they span the full container width instead of being squeezed into a narrow 200px column.
  2. **Aggressive Wrapping (`overflow-wrap: anywhere`):** Apply `overflow-wrap: anywhere !important; word-break: break-word !important; white-space: normal !important; overflow: visible !important;` to `.detail-item .value`. This ensures any long identifier or code breaks cleanly at the container boundary and naturally expands the card height without hiding or clipping any characters.

### 2.8. Handling Variations in P/F/A Terminology ("New Plan")
- **The Challenge:** After an update, the uploaded ASK Excel file changed the terminology for the Plan rows from "Plan" to "New Plan". The backend extractor strictly mapped "P" and "PLAN" to the internal "Plan" state, causing it to skip all rows labeled "New Plan". As a result, the dashboard displayed 0 packages for the ASK project.
- **Root Cause:** The `pfa_values` dictionary in `server.py` mapped a hardcoded list of acceptable string variations (e.g., "P", "PLAN", "F", "FORECAST"). It lacked flexibility for unannounced formatting changes made by the planning engineers in the source Excel files.
- **The Solution & Best Practice:**
  - Added `"NEW PLAN": "Plan"` to the `pfa_values` dictionaries in both `Z1F_CONFIG` and `ASK_CONFIG`.
  - This allows the data extractor to successfully identify the start of package groups in the new `"Attachment 2-Procurement Plan-ASK-20260726 R1.xlsx"` file which started using "New Plan" instead of "Plan".
  - **Testing & Verification:** Ran `python -c "from server import extract_all_data; data = extract_all_data(); print('Z1F Packages:', data['projects']['Z1F'].get('package_count')); print('ASK Packages:', data['projects']['ASK'].get('package_count'))"` to confirm parsing. Successfully parsed 115 packages for Z1F and 100 packages for ASK.
  - Moving forward, if the dashboard reads 0 packages after an upload or if any future Excel files change the wording again (e.g., "Original Plan", "Revised Plan"), we will simply need to add them to the `pfa_values` dictionary mapping in `server.py` just like we did for "New Plan".

### 2.9. Stale Static Float vs. Dynamic Reality (Using `Today` as Baseline for Overdue POs)
- **The Challenge:** In static Excel spreadsheets, the Forecast Float formula calculates delivery dates from a stale forecast PO date (e.g., `2026-07-08`), showing packages with positive float (e.g. `+6 days`, yellow) even when the PO award is several weeks overdue and unissued. This created a false sense of security, misrepresenting critical schedule risks to project management.
- **Root Cause:** Static Excel formulas do not dynamically advance the PO award baseline when forecast milestone dates elapse into the past without an actual PO issuance.
- **The Solution & Best Practice:**
  1. **Badge Nomenclature & Operational Semantics:**
     - `Act` (Actual): PO has been formally awarded. Contractual delivery is calculated from the actual PO issuance date.
     - `Fcst*` (Dynamic Forecast - Overdue Alert): PO has not been awarded, and the forecast date has elapsed ($\text{Forecast PO} < \mathbf{Today}$). The award baseline is dynamically pinned to $\mathbf{Today}$.
     - `Fcst` (Standard Forecast): PO has not been awarded, but the forecast award date remains in the future ($\text{Forecast PO} \ge \mathbf{Today}$).
  2. **Post-Award State (`Actual PO` exists):** Compute true actual float:
     $$\text{Actual Delivery Date} = \text{Actual PO Date} + \text{Delivery Duration (Days)}$$
     $$\text{Actual Float} = \text{Plan ROS Date} - \text{Actual Delivery Date}$$
  3. **Unissued Overdue State (`Forecast PO < Today`):** When a PO has not been awarded and its forecast award date is in the past, dynamically reset the baseline to **Today**:
     $$\text{Dynamic Delivery Date} = \mathbf{Today} + \text{Delivery Duration (Days)}$$
     $$\mathbf{Dynamic\ Forecast\ Float} = \text{Plan ROS Date} - (\mathbf{Today} + \text{Delivery Duration (Days)})$$
     This guarantees that float degrades by exactly 1 day for each day the PO award is delayed past its forecast date.
  4. **Formula Sign Convention:** Always compute $\text{ROS Date} - \text{Delivery Date}$ so that late deliveries yield negative float (`< 0` &rarr; Red 🔴), tight buffers yield small positive float (`< 21d` &rarr; Yellow 🟡), and safe buffers yield high float (`>= 21d` &rarr; Green 🟢).
  5. **Unit Resilience (Weeks vs. Days):** Source spreadsheets alternate between weeks (Col 36/37) and days (Col 37/38). The parser checks for explicit day duration, falls back to $(\text{weeks} \times 7)$, and falls back to explicit delivery/ETA dates if duration numbers are blank.
  6. **Table Sorting Prioritization:** Provide a dedicated `Sort: Lowest Float (Critical First)` option in the UI so project controls can immediately filter and identify packages where overdue PO awards are threatening the offshore installation schedule.

---

## 3. Summary of Pitfalls vs. Resolutions Table

| # | Pitfall encountered | Technical Cause | Engineering Resolution |
|---|---|---|---|
| 1 | Upload button white & non-functional after Git push | Stale CSS/JS cached by browser & Render CDN | Added `end_headers()` cache headers (`must-revalidate`) + `?v=YYYYMMDD` to `<link>` and `<script>` in `index.html`. |
| 2 | `SimpleHTTPRequestHandler` 404 error on query URLs | `self.path` containing `?v=...` failing path translation | Added `path_clean = self.path.split("?")[0]` before routing or delegating to `super().do_GET()`. |
| 3 | File overwrite / ambiguity across weekly updates | Static file names replacing previous week's history | Enforced `Attachment X-Procurement Plan-<Proj> - YYYYMMDD.xlsx` dynamic naming format. |
| 4 | Uploading `.xlsx` without immediate dashboard update | Data cache locked/stale until manual page reload | Integrated `extract_all_data()` execution directly inside `/api/upload` endpoint handler. |
| 5 | Detail cards right-aligned & stage headers misaligned | Outer `.data-table` rules (`td:last-child`) leaked into nested detail cell | Scoped outer rules with direct child selectors (`>`) and set explicit `text-align: left !important` on `.detail-row td`. |
| 6 | Long MR numbers & package names cut off (`MR-...`) | Inherited `white-space: nowrap` + `overflow: hidden` in grid card cells | Added `overflow-wrap: anywhere !important; white-space: normal !important; overflow: visible !important;` to detail values and spanned package name full width (`grid-column: 1 / -1`). |
| 7 | Dashboard showing 0 packages after file upload | Source Excel changed terminology from "Plan" to "New Plan" (e.g., `Attachment 2-...-20260726 R1.xlsx`) | Added `"NEW PLAN": "Plan"` to the `pfa_values` mapping in `server.py` and verified package counts via CLI. |
| 8 | Stale positive float shown when PO is overdue | Static Excel formula calculated from passed PO date | Automatically set baseline to `Today` when `Forecast PO < Today` and `Actual PO` is blank, dynamically eroding float day-by-day. |
| 9 | Inconsistent column names & whitespace (`Float `, `Lead Time`, `days`) | Trailing spaces in headers or different project naming conventions | Implemented case-insensitive, stripped alias detection in `auto_detect_columns` in `server.py` with multi-key fallbacks. |

---

## 4. Recommendations for Future Development

1. **Automated End-to-End Cypress / Playwright Tests:** Add automated browser testing for file drag-and-drop and modal interactions to catch styling or script reference breaks prior to cloud deployment.
2. **Persistent Storage Mounts on Cloud Platforms:** Note that ephemeral container hosting on Render (`render.yaml`) resets disk space upon container restarts. If permanent multi-year historical storage of Excel files is required across restarts, attach a persistent disk volume on Render or integrate cloud object storage (AWS S3 / Google Cloud Storage).
3. **Structured Logging:** Keep `logging.basicConfig(level=logging.INFO)` in `server.py` and output timestamped audit logs whenever a file is uploaded (`[UPLOAD] User uploaded Attachment 1-... at 2026-07-12 18:30:00`).
