/**
 * Procurement Tracking Dashboard
 * Tracks delays across procurement stages from Z1F and ASK Excel files.
 * 
 * Delay Logic:
 *   - If stage is "Bidder List Approval" and Actual is blank → IGNORED (not counted as delay)
 *   - If Actual date is blank AND Forecast date has passed → DELAYED
 *   - If Forecast date is within 1-week lookahead → AT RISK (due soon)
 *   - If Actual date exists → COMPLETED
 *   - Otherwise → UPCOMING
 */

// ─── State ──────────────────────────────────────────────────────────────────

let appState = {
    data: null,
    activeProject: 'ALL',  // 'Z1F', 'ASK', 'ALL'
    searchQuery: '',
    statusFilter: 'all',   // 'all', 'delayed', 'atrisk', 'ontrack', 'completed'
    expandedRows: new Set(),
    sortField: 'delay',
    sortDir: 'desc',
};

// ─── Constants ──────────────────────────────────────────────────────────────

function getToday() {
    const d = new Date();
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function getLookaheadEnd() {
    const d = new Date();
    d.setDate(d.getDate() + 7);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// ─── Data Loading ───────────────────────────────────────────────────────────

async function loadData() {
    showLoading(true);
    try {
        const resp = await fetch('/api/data');
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        appState.data = await resp.json();
        const urlProj = new URLSearchParams(window.location.search).get('project');
        if (urlProj && ['ALL', 'Z1F', 'ASK'].includes(urlProj.toUpperCase())) {
            appState.activeProject = urlProj.toUpperCase();
            document.querySelectorAll('.tab-btn').forEach(btn => {
                btn.classList.toggle('active', btn.dataset.project === appState.activeProject);
            });
        }
        renderDashboard();
        showToast('Data loaded successfully', 'success');
    } catch (err) {
        console.error('Failed to load data:', err);
        showToast('Failed to load data: ' + err.message, 'error');
    } finally {
        showLoading(false);
    }
}

// ─── Delay Computation ─────────────────────────────────────────────────────

function computePackageStatus(pkg) {
    const today = getToday();
    const lookaheadEnd = getLookaheadEnd();
    
    let totalStages = 0;
    let completedStages = 0;
    let delayedStages = [];
    let atRiskStages = [];
    let maxDelay = 0;
    let currentStage = null;
    let currentStageStatus = 'upcoming';

    for (const stage of pkg.stages) {
        if (!stage.forecast && !stage.plan && !stage.actual) continue;

        const isBidderList = stage.name.toLowerCase().includes('bidder list approval');
        const isBidClosing = stage.name.toLowerCase().includes('bid closing date');

        const stageInfo = {
            name: stage.name,
            plan: stage.plan,
            forecast: stage.forecast,
            actual: stage.actual,
            status: 'upcoming',
            delayDays: 0,
        };

        if (stage.actual) {
            totalStages++;
            // Stage completed
            stageInfo.status = 'completed';
            completedStages++;

            // Check if actual was delayed vs plan
            if (stage.plan && stage.actual > stage.plan) {
                const delay = dateDiffDays(stage.plan, stage.actual);
                stageInfo.delayDays = delay;
            }
        } else if (isBidderList || isBidClosing) {
            // User requirement: Actual blank in Bidder List Approval or Bid Closing Date should be ignored and does not mean delay
            stageInfo.status = 'ignored';
        } else if (stage.forecast) {
            totalStages++;
            // Stage not completed, check forecast
            if (stage.forecast < today) {
                // Forecast date has passed but no actual → DELAYED
                stageInfo.status = 'delayed';
                stageInfo.delayDays = dateDiffDays(stage.forecast, today);
                delayedStages.push(stageInfo);
                if (stageInfo.delayDays > maxDelay) maxDelay = stageInfo.delayDays;
            } else if (stage.forecast <= lookaheadEnd) {
                // Forecast within 1-week lookahead → AT RISK / DUE SOON
                stageInfo.status = 'atrisk';
                atRiskStages.push(stageInfo);
            } else {
                // Future date
                stageInfo.status = 'upcoming';
            }

            // First non-completed stage = current stage
            if (!currentStage) {
                currentStage = stageInfo;
                currentStageStatus = stageInfo.status;
            }
        } else if (stage.plan) {
            totalStages++;
            // Only plan date, no forecast or actual
            if (stage.plan < today) {
                // Plan date passed, no forecast or actual → treat as delayed
                stageInfo.status = 'delayed';
                stageInfo.delayDays = dateDiffDays(stage.plan, today);
                delayedStages.push(stageInfo);
                if (stageInfo.delayDays > maxDelay) maxDelay = stageInfo.delayDays;
            } else if (stage.plan <= lookaheadEnd) {
                stageInfo.status = 'atrisk';
                atRiskStages.push(stageInfo);
            }

            if (!currentStage) {
                currentStage = stageInfo;
                currentStageStatus = stageInfo.status;
            }
        }

        stage._computed = stageInfo;
    }

    // Overall package status
    let overallStatus;
    if (delayedStages.length > 0) {
        overallStatus = 'delayed';
    } else if (atRiskStages.length > 0) {
        overallStatus = 'atrisk';
    } else if (completedStages === totalStages && totalStages > 0) {
        overallStatus = 'completed';
    } else {
        overallStatus = 'ontrack';
    }

    return {
        overallStatus,
        currentStage: currentStage ? currentStage.name : '—',
        currentStageStatus,
        totalStages,
        completedStages,
        delayedCount: delayedStages.length,
        atRiskCount: atRiskStages.length,
        maxDelay,
        delayedStages,
        atRiskStages,
    };
}

function dateDiffDays(dateStr1, dateStr2) {
    const d1 = new Date(dateStr1);
    const d2 = new Date(dateStr2);
    return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
}

function formatDate(dateStr) {
    if (!dateStr) return '—';
    const d = new Date(dateStr);
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear().toString().slice(-2)}`;
}

// ─── Rendering ──────────────────────────────────────────────────────────────

function renderDashboard() {
    if (!appState.data) return;

    const allPackages = getAllPackages();
    const filtered = filterPackages(allPackages);

    renderStats(allPackages);
    renderTable(filtered);
    renderLastUpdated();
}

function getAllPackages() {
    const result = [];
    const projects = appState.data.projects || {};

    for (const [projectName, projectData] of Object.entries(projects)) {
        if (projectData.error) continue;
        if (appState.activeProject !== 'ALL' && appState.activeProject !== projectName) continue;

        for (const pkg of (projectData.packages || [])) {
            const status = computePackageStatus(pkg);
            result.push({
                ...pkg,
                project: projectName,
                _status: status,
            });
        }
    }

    // Sort
    result.sort((a, b) => {
        if (appState.sortField === 'delay') {
            const diff = b._status.maxDelay - a._status.maxDelay;
            return appState.sortDir === 'desc' ? diff : -diff;
        }
        if (appState.sortField === 'name') {
            const diff = a.package_name.localeCompare(b.package_name);
            return appState.sortDir === 'desc' ? -diff : diff;
        }
        if (appState.sortField === 'status') {
            const order = { delayed: 0, atrisk: 1, ontrack: 2, completed: 3 };
            const diff = (order[a._status.overallStatus] || 4) - (order[b._status.overallStatus] || 4);
            return appState.sortDir === 'asc' ? diff : -diff;
        }
        return 0;
    });

    return result;
}

function filterPackages(packages) {
    return packages.filter(pkg => {
        // Search filter
        if (appState.searchQuery) {
            const q = appState.searchQuery.toLowerCase();
            const match = pkg.package_name.toLowerCase().includes(q) ||
                          pkg.rfq_no.toLowerCase().includes(q) ||
                          (pkg.mr_no || '').toLowerCase().includes(q);
            if (!match) return false;
        }

        // Status filter
        if (appState.statusFilter !== 'all') {
            if (pkg._status.overallStatus !== appState.statusFilter) return false;
        }

        return true;
    });
}

function renderStats(allPackages) {
    const stats = {
        total: allPackages.length,
        delayed: allPackages.filter(p => p._status.overallStatus === 'delayed').length,
        atrisk: allPackages.filter(p => p._status.overallStatus === 'atrisk').length,
        ontrack: allPackages.filter(p => p._status.overallStatus === 'ontrack').length,
        completed: allPackages.filter(p => p._status.overallStatus === 'completed').length,
    };

    document.getElementById('stat-total').textContent = stats.total;
    document.getElementById('stat-delayed').textContent = stats.delayed;
    document.getElementById('stat-atrisk').textContent = stats.atrisk;
    document.getElementById('stat-ontrack').textContent = stats.ontrack;
    document.getElementById('stat-completed').textContent = stats.completed;

    // Desc
    document.getElementById('stat-total-desc').textContent = `Across ${appState.activeProject === 'ALL' ? 'both projects' : appState.activeProject}`;
    document.getElementById('stat-delayed-desc').textContent = `Stages overdue (no actual date)`;
    document.getElementById('stat-atrisk-desc').textContent = `Due within 7-day lookahead`;
    document.getElementById('stat-ontrack-desc').textContent = `Progressing on schedule`;
}

function renderTable(packages) {
    const tbody = document.getElementById('table-body');
    const countEl = document.getElementById('table-count');

    countEl.textContent = `${packages.length} packages`;

    if (packages.length === 0) {
        tbody.innerHTML = `
            <tr><td colspan="7">
                <div class="empty-state">
                    <div class="empty-icon">📋</div>
                    <div class="empty-title">No packages found</div>
                    <div>Adjust filters or search to see results</div>
                </div>
            </td></tr>`;
        return;
    }

    let html = '';
    for (const pkg of packages) {
        const s = pkg._status;
        const isExpanded = appState.expandedRows.has(pkg.rfq_no + pkg.project);
        const rowClass = s.overallStatus === 'delayed' ? 'delayed-row' : 
                         s.overallStatus === 'atrisk' ? 'atrisk-row' : '';

        // Build delayed stages text
        let delayText = '—';
        if (s.delayedStages.length > 0) {
            delayText = s.delayedStages.map(ds => ds.name).join(', ');
        } else if (s.atRiskStages.length > 0) {
            delayText = s.atRiskStages.map(ds => ds.name).join(', ');
        }

        // Truncate delay text
        if (delayText.length > 50) {
            delayText = delayText.substring(0, 47) + '...';
        }

        html += `
        <tr class="${rowClass}" onclick="toggleExpand('${pkg.rfq_no + pkg.project}')" id="row-${css_safe(pkg.rfq_no + pkg.project)}">
            <td>
                <span class="expand-icon ${isExpanded ? 'expanded' : ''}">▶</span>
            </td>
            <td style="color: var(--text-primary); font-weight: 500;">${escHtml(pkg.package_name)}</td>
            <td>${escHtml(pkg.rfq_no)}</td>
            <td><span class="status-badge ${s.overallStatus}"><span class="status-dot"></span>${statusLabel(s.overallStatus)}</span></td>
            <td>${escHtml(s.currentStage)}</td>
            <td title="${escHtml(delayText)}">${escHtml(delayText)}</td>
            <td>${s.maxDelay > 0 ? `<span class="delay-value negative">${s.maxDelay}d</span>` : '<span class="delay-value zero">—</span>'}</td>
        </tr>
        <tr class="detail-row ${isExpanded ? 'expanded' : ''}" id="detail-${css_safe(pkg.rfq_no + pkg.project)}">
            <td colspan="7">
                <div class="detail-content">
                    ${renderDetailContent(pkg)}
                </div>
            </td>
        </tr>`;
    }

    tbody.innerHTML = html;
}

function renderDetailContent(pkg) {
    const s = pkg._status;

    // Info grid
    let infoHtml = `
    <div class="detail-grid">
        <div class="detail-item pkg-name-box"><div class="label">Package Name</div><div class="value">${escHtml(pkg.package_name)}</div></div>
        <div class="detail-item"><div class="label">RFQ No.</div><div class="value">${escHtml(pkg.rfq_no)}</div></div>
        ${pkg.mr_no ? `<div class="detail-item"><div class="label">MR No.</div><div class="value">${escHtml(pkg.mr_no)}</div></div>` : ''}
        <div class="detail-item"><div class="label">Priority</div><div class="value">${escHtml(pkg.priority)}</div></div>
        <div class="detail-item"><div class="label">Long Lead Item</div><div class="value">${escHtml(pkg.lli)}</div></div>
        <div class="detail-item"><div class="label">Project</div><div class="value">${pkg.project}</div></div>
        <div class="detail-item"><div class="label">Progress</div><div class="value">${s.completedStages}/${s.totalStages} stages</div></div>
    </div>`;

    // Pipeline visualization
    let pipelineHtml = '<div class="pipeline">';
    for (const stage of pkg.stages) {
        if (!stage._computed) continue;
        const c = stage._computed;
        const dateToShow = c.actual || c.forecast || c.plan || '';

        pipelineHtml += `
        <div class="pipeline-stage stage-${c.status}" title="${c.name}\nPlan: ${formatDate(c.plan)}\nForecast: ${formatDate(c.forecast)}\nActual: ${formatDate(c.actual)}${c.delayDays > 0 ? '\nDelay: ' + c.delayDays + ' days' : ''}">
            <div class="stage-name">${abbreviateStage(c.name)}</div>
            <div class="stage-date">${formatDate(dateToShow)}</div>
            ${c.delayDays > 0 && (c.status === 'delayed' || c.status === 'atrisk') ? `<div class="stage-delay-badge">${c.delayDays}d</div>` : ''}
        </div>`;
    }
    pipelineHtml += '</div>';

    // Stage detail table
    let stageTableHtml = `
    <table class="stage-table">
        <colgroup>
            <col style="width: 28%;">
            <col style="width: 15%;">
            <col style="width: 15%;">
            <col style="width: 15%;">
            <col style="width: 15%;">
            <col style="width: 12%;">
        </colgroup>
        <thead>
            <tr>
                <th>Stage</th>
                <th>Plan Date</th>
                <th>Forecast Date</th>
                <th>Actual Date</th>
                <th>Status</th>
                <th>Delay (days)</th>
            </tr>
        </thead>
        <tbody>`;

    for (const stage of pkg.stages) {
        if (!stage._computed) continue;
        const c = stage._computed;
        const rowCls = c.status === 'delayed' ? 'stage-delayed-row' : 
                       c.status === 'atrisk' ? 'stage-atrisk-row' : '';

        stageTableHtml += `
            <tr class="${rowCls}">
                <td style="color: var(--text-primary); font-weight: 500;">${escHtml(c.name)}</td>
                <td>${formatDate(c.plan)}</td>
                <td>${formatDate(c.forecast)}</td>
                <td>${formatDate(c.actual)}</td>
                <td><span class="status-badge ${c.status}"><span class="status-dot"></span>${statusLabel(c.status)}</span></td>
                <td>${c.delayDays > 0 ? `<span class="delay-value negative">${c.delayDays}d</span>` : c.delayDays === 0 ? '—' : `<span class="delay-value positive">${Math.abs(c.delayDays)}d ahead</span>`}</td>
            </tr>`;
    }
    stageTableHtml += '</tbody></table>';

    return infoHtml + pipelineHtml + stageTableHtml;
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function statusLabel(status) {
    const labels = {
        delayed: 'Delayed',
        atrisk: 'Due Soon',
        ontrack: 'On Track',
        completed: 'Completed',
        upcoming: 'Upcoming',
        ignored: 'Ignored',
    };
    return labels[status] || status;
}

function abbreviateStage(name) {
    const abbrevs = {
        'Bidder List Approval': 'BLA',
        'MR Issued (IFR)': 'MR Iss.',
        'MR IFA Issued': 'MR IFA',
        'MR IFA Approved': 'MR IFA Apr',
        'MR Issued': 'MR Iss.',
        'MR Approved': 'MR Apr.',
        'RFQ Issued': 'RFQ Iss.',
        'RFQ Approved': 'RFQ Apr.',
        'CFT Issuance': 'CFT',
        'Bid Closing Date': 'Bid Close',
        'TBE Issued': 'TBE Iss.',
        'TBE Approved': 'TBE Apr.',
        'PO Issued': 'PO Iss.',
        'Vendor Acknowledgement': 'VA',
        'KOM': 'KOM',
        'VDB Submission': 'VDB',
        'VD Submission': 'VD',
        'Approval of Key VP': 'Key VP',
        'Approval of Key VD': 'Key VD',
        'Main Material Arrived': 'Mat. Arr.',
        'Delivery of Major Materials': 'Mat. Del.',
        'PIM': 'PIM',
        'Start Production': 'Prod.',
        'FAT': 'FAT',
        'Ready for Shipment': 'Ship',
        'Receipt at Worksite': 'RaW',
        'Punch List Clearance': 'PLC',
        'Final Documentation': 'FDoc',
    };
    return abbrevs[name] || name.substring(0, 8);
}

function escHtml(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function css_safe(str) {
    return String(str).replace(/[^a-zA-Z0-9-_]/g, '_');
}

// ─── UI Actions ─────────────────────────────────────────────────────────────

function toggleExpand(id) {
    const safeId = css_safe(id);
    const detailRow = document.getElementById('detail-' + safeId);
    const mainRow = document.getElementById('row-' + safeId);

    if (appState.expandedRows.has(id)) {
        appState.expandedRows.delete(id);
        detailRow?.classList.remove('expanded');
        mainRow?.querySelector('.expand-icon')?.classList.remove('expanded');
    } else {
        appState.expandedRows.add(id);
        detailRow?.classList.add('expanded');
        mainRow?.querySelector('.expand-icon')?.classList.add('expanded');
    }
}

function setActiveProject(project) {
    appState.activeProject = project;
    appState.expandedRows.clear();

    // Update tab UI
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.project === project);
    });

    renderDashboard();
}

function setStatusFilter(status) {
    appState.statusFilter = status;
    renderDashboard();
}

function setSearchQuery(query) {
    appState.searchQuery = query;
    renderDashboard();
}

function sortBy(field) {
    if (appState.sortField === field) {
        appState.sortDir = appState.sortDir === 'desc' ? 'asc' : 'desc';
    } else {
        appState.sortField = field;
        appState.sortDir = 'desc';
    }
    renderDashboard();
}

async function refreshData() {
    const btn = document.getElementById('btn-update');
    btn.classList.add('loading');
    showLoading(true);
    try {
        const resp = await fetch('/api/refresh');
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        appState.data = await resp.json();
        renderDashboard();
        showToast('Data refreshed from Excel files', 'success');
    } catch (err) {
        console.error('Failed to refresh data:', err);
        showToast('Failed to refresh: ' + err.message, 'error');
    } finally {
        showLoading(false);
        btn.classList.remove('loading');
    }
}

// ─── UI Utilities ───────────────────────────────────────────────────────────

function showLoading(show) {
    const el = document.getElementById('loading-overlay');
    if (el) el.style.display = show ? 'flex' : 'none';
    const table = document.getElementById('table-container');
    if (table) table.style.display = show ? 'none' : 'block';
}

function showToast(message, type = 'success') {
    let toast = document.getElementById('toast');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'toast';
        toast.className = 'toast';
        document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.className = `toast ${type}`;
    setTimeout(() => toast.classList.add('show'), 10);
    setTimeout(() => toast.classList.remove('show'), 3000);
}

function renderLastUpdated() {
    const el = document.getElementById('last-updated');
    if (el && appState.data) {
        const genAt = new Date(appState.data.generated_at);
        el.textContent = `Last updated: ${genAt.toLocaleString()}`;
    }
    if (appState.data && appState.data.projects) {
        const z1f = appState.data.projects.Z1F?.filename || '—';
        const ask = appState.data.projects.ASK?.filename || '—';
        const z1fEl = document.getElementById('src-z1f');
        const askEl = document.getElementById('src-ask');
        const modalZ1f = document.getElementById('modal-z1f-file');
        const modalAsk = document.getElementById('modal-ask-file');
        if (z1fEl) z1fEl.textContent = z1f;
        if (askEl) askEl.textContent = ask;
        if (modalZ1f) modalZ1f.textContent = z1f;
        if (modalAsk) modalAsk.textContent = ask;
    }
}

// ─── Date display ───────────────────────────────────────────────────────────

function updateDateDisplay() {
    const todayEl = document.getElementById('display-today');
    const lookaheadEl = document.getElementById('display-lookahead');
    if (todayEl) todayEl.textContent = formatDate(getToday());
    if (lookaheadEl) lookaheadEl.textContent = formatDate(getLookaheadEnd());
}

// ─── Upload Center Modal & Drag & Drop ──────────────────────────────────────

function openUploadModal() {
    const overlay = document.getElementById('upload-modal-overlay');
    if (overlay) overlay.classList.add('open');
    setupDragAndDrop();
}

function closeUploadModal() {
    const overlay = document.getElementById('upload-modal-overlay');
    if (overlay) overlay.classList.remove('open');
}

function closeModalOnOverlay(e) {
    if (e.target.id === 'upload-modal-overlay') {
        closeUploadModal();
    }
}

function setupDragAndDrop() {
    const dropZone = document.getElementById('drop-zone');
    if (!dropZone || dropZone.dataset.initialized) return;
    dropZone.dataset.initialized = 'true';

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files && files.length > 0) {
            handleDroppedFiles(files);
        }
    }, false);
}

function handleDroppedFiles(files) {
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        let projectType = '';
        const nameUpper = file.name.toUpperCase();
        if (nameUpper.includes('Z1F') || nameUpper.includes('TOPSIDE') || nameUpper.includes('JACKET')) {
            projectType = 'Z1F';
        } else if (nameUpper.includes('ASK') || nameUpper.includes('PIPELINE') || nameUpper.includes('OVERALL')) {
            projectType = 'ASK';
        } else {
            showToast('Could not auto-detect Z1F or ASK from filename. Please use Upload Specific File card.', 'error');
            continue;
        }
        uploadFile(file, projectType);
    }
}

function handleGlobalFileInput(e) {
    const files = e.target.files;
    if (files && files.length > 0) {
        handleDroppedFiles(files);
        e.target.value = '';
    }
}

function handleSpecificFileInput(e, projectType) {
    const files = e.target.files;
    if (files && files.length > 0) {
        uploadFile(files[0], projectType);
        e.target.value = '';
    }
}

async function uploadFile(file, projectType) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xlsm')) {
        showToast('Only .xlsx and .xlsm files are supported', 'error');
        return;
    }

    showToast(`Uploading ${projectType} file: ${file.name}...`, 'success');
    showLoading(true);

    try {
        const formData = new FormData();
        formData.append('project', projectType);
        formData.append('file', file, file.name);

        const resp = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        if (!resp.ok) {
            const errJson = await resp.json().catch(() => ({ error: `HTTP ${resp.status}` }));
            throw new Error(errJson.error || `HTTP ${resp.status}`);
        }

        const result = await resp.json();
        appState.data = result.data;
        renderDashboard();
        showToast(`Successfully uploaded ${result.filename}`, 'success');
    } catch (err) {
        console.error('Upload failed:', err);
        showToast(`Failed to upload: ${err.message}`, 'error');
    } finally {
        showLoading(false);
    }
}

// ─── Initialize ─────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
    updateDateDisplay();
    loadData();

    // Search input debounce
    let searchTimeout;
    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            clearTimeout(searchTimeout);
            searchTimeout = setTimeout(() => setSearchQuery(e.target.value), 200);
        });
    }
});
