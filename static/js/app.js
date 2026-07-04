/* ==========================================================================
   CGD Analytics Dashboard -- Frontend JavaScript
   Handles navigation, data fetching, table rendering, and model execution.
   ========================================================================== */

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let currentInputTab = 'city';
let currentInputData = [];

// Chart name -> human-readable title mapping
const CHART_TITLES = {
    'cumulative_cashflow_top3.png': 'Cumulative Cash Flow — Top 3 Priority Zones',
    'demand_forecast_stacked_bar.png': 'PNG Demand Forecast by Customer Segment',
    'demand_growth_curve.png': 'Total PNG Demand Growth Curve',
    'zone_irr_comparison.png': 'Zone-wise IRR Comparison',
    'zone_priority_ranking.png': 'Zone Expansion Priority Ranking',
    'zone_score_breakdown.png': 'Zone Score Breakdown by Factor',
};

// ---------------------------------------------------------------------------
// Navigation
// ---------------------------------------------------------------------------
function navigateTo(sectionId) {
    // Deactivate all sections & nav items
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

    // Activate target
    const section = document.getElementById('section-' + sectionId);
    const navItem = document.getElementById('nav-' + sectionId);
    if (section) section.classList.add('active');
    if (navItem) navItem.classList.add('active');

    // Update page title
    const titles = { dashboard: 'Dashboard', charts: 'Charts', data: 'Data Tables', inputs: 'Edit Inputs' };
    document.getElementById('pageTitle').textContent = titles[sectionId] || 'Dashboard';

    // Close sidebar on mobile
    document.getElementById('sidebar').classList.remove('open');

    // Lazy-load section data
    if (sectionId === 'charts') loadCharts();
    if (sectionId === 'inputs') loadInputTab(currentInputTab);
}

// Sidebar nav click handlers
document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        navigateTo(item.dataset.section);
    });
});

function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('open');
}

// ---------------------------------------------------------------------------
// API Helpers
// ---------------------------------------------------------------------------
async function fetchJSON(url) {
    try {
        const res = await fetch(url);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        return await res.json();
    } catch (err) {
        console.error('Fetch error:', url, err);
        return null;
    }
}

// ---------------------------------------------------------------------------
// KPIs
// ---------------------------------------------------------------------------
async function loadKPIs() {
    const kpis = await fetchJSON('/api/kpis');
    if (!kpis) return;

    setKPI('kpi-demand', kpis.year10_demand != null ? Number(kpis.year10_demand).toLocaleString() : '—');
    setKPI('kpi-growth', kpis.demand_growth_pct != null ? kpis.demand_growth_pct + '%' : '—');
    setKPI('kpi-capex', kpis.total_capex_cr != null ? '₹' + kpis.total_capex_cr : '—');
    setKPI('kpi-irr', kpis.avg_irr != null ? kpis.avg_irr + '%' : '—');
    setKPI('kpi-viable', kpis.viable_zones || '—');
    setKPI('kpi-npv', kpis.total_npv_cr != null ? '₹' + kpis.total_npv_cr : '—');
}

function setKPI(id, value) {
    const el = document.getElementById(id);
    if (el) {
        el.textContent = value;
        el.style.animation = 'none';
        el.offsetHeight; // reflow
        el.style.animation = 'fadeIn 0.5s ease';
    }
}

// ---------------------------------------------------------------------------
// Dashboard Preview Charts
// ---------------------------------------------------------------------------
async function loadPreviewCharts() {
    const charts = await fetchJSON('/api/charts');
    if (!charts) return;

    const preview1 = document.getElementById('previewChart1');
    const preview2 = document.getElementById('previewChart2');

    if (charts.includes('demand_forecast_stacked_bar.png')) {
        preview1.src = '/charts/demand_forecast_stacked_bar.png?' + Date.now();
    }
    if (charts.includes('zone_irr_comparison.png')) {
        preview2.src = '/charts/zone_irr_comparison.png?' + Date.now();
    }
}

// ---------------------------------------------------------------------------
// Dashboard Financial Table
// ---------------------------------------------------------------------------
async function loadDashFinancialTable() {
    const data = await fetchJSON('/api/outputs/financial');
    if (!data || data.length === 0) {
        document.getElementById('dashFinancialTable').innerHTML =
            '<p class="placeholder-text">No financial data yet. Run the model first.</p>';
        return;
    }
    document.getElementById('dashFinancialTable').innerHTML = buildReadOnlyTable(data);
}

// ---------------------------------------------------------------------------
// Charts Gallery
// ---------------------------------------------------------------------------
async function loadCharts() {
    const charts = await fetchJSON('/api/charts');
    const container = document.getElementById('chartGallery');

    if (!charts || charts.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No charts available. Run the model to generate them.</p>';
        return;
    }

    container.innerHTML = charts.map(filename => {
        const title = CHART_TITLES[filename] || filename.replace(/_/g, ' ').replace('.png', '');
        return `
            <div class="chart-card">
                <div class="chart-card-header"><h4>${title}</h4></div>
                <div class="chart-card-body">
                    <img src="/charts/${filename}?${Date.now()}" alt="${title}" loading="lazy">
                </div>
            </div>
        `;
    }).join('');
}

// ---------------------------------------------------------------------------
// Data Tables
// ---------------------------------------------------------------------------
function switchDataTab(tab, btnEl) {
    document.querySelectorAll('#section-data .tab').forEach(t => t.classList.remove('active'));
    if (btnEl) btnEl.classList.add('active');

    const container = document.getElementById('dataTableContainer');
    container.innerHTML = '<p class="placeholder-text">Loading...</p>';

    const endpoints = {
        demand: '/api/outputs/demand',
        zoneRank: '/api/outputs/zones',
        financial: '/api/outputs/financial',
    };

    fetchJSON(endpoints[tab]).then(data => {
        if (!data || data.length === 0) {
            container.innerHTML = '<p class="placeholder-text">No data available. Run the model first.</p>';
            return;
        }
        container.innerHTML = buildReadOnlyTable(data);
    });
}

function buildReadOnlyTable(data) {
    if (!data || data.length === 0) return '<p class="placeholder-text">No data.</p>';

    const keys = Object.keys(data[0]);
    let html = '<table><thead><tr>';
    keys.forEach(k => {
        const label = k.replace(/_/g, ' ');
        html += `<th>${label}</th>`;
    });
    html += '</tr></thead><tbody>';

    data.forEach(row => {
        html += '<tr>';
        keys.forEach(k => {
            let val = row[k];
            // Format numbers
            if (typeof val === 'number') {
                if (Number.isInteger(val)) {
                    val = val.toLocaleString();
                } else {
                    val = val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                }
            }
            // Add badges for Viability / Recommendation columns
            if (k === 'Viability') {
                if (val === 'Financially Viable') val = `<span class="badge badge-green">${val}</span>`;
                else if (val === 'Marginal') val = `<span class="badge badge-amber">${val}</span>`;
                else val = `<span class="badge badge-red">${val}</span>`;
            }
            if (k === 'Recommendation') {
                if (val && val.includes('Year 1')) val = `<span class="badge badge-green">${val}</span>`;
                else if (val && val.includes('Year 2')) val = `<span class="badge badge-amber">${val}</span>`;
                else if (val && val.includes('Year 3')) val = `<span class="badge badge-red">${val}</span>`;
            }
            html += `<td>${val != null ? val : ''}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
}

// ---------------------------------------------------------------------------
// Edit Inputs
// ---------------------------------------------------------------------------
function switchInputTab(tab, btnEl) {
    document.querySelectorAll('#section-inputs .tab').forEach(t => t.classList.remove('active'));
    if (btnEl) btnEl.classList.add('active');
    currentInputTab = tab;

    const titles = { city: 'City Parameters', zones: 'Zone Data', financial: 'Financial Parameters' };
    document.getElementById('inputTableTitle').textContent = titles[tab] || '';

    loadInputTab(tab);
}

async function loadInputTab(tab) {
    const container = document.getElementById('inputTableContainer');
    container.innerHTML = '<p class="placeholder-text">Loading...</p>';

    const endpoints = { city: '/api/inputs/city', zones: '/api/inputs/zones', financial: '/api/inputs/financial' };
    const data = await fetchJSON(endpoints[tab]);

    if (!data || data.length === 0) {
        container.innerHTML = '<p class="placeholder-text">No input data found.</p>';
        return;
    }

    currentInputData = JSON.parse(JSON.stringify(data)); // deep copy
    container.innerHTML = buildEditableTable(data, tab);
}

function buildEditableTable(data, tab) {
    if (!data || data.length === 0) return '<p class="placeholder-text">No data.</p>';

    const keys = Object.keys(data[0]);
    let html = '<table><thead><tr>';
    keys.forEach(k => {
        const label = k.replace(/_/g, ' ');
        html += `<th>${label}</th>`;
    });
    html += '</tr></thead><tbody>';

    data.forEach((row, rowIdx) => {
        html += '<tr>';
        keys.forEach((k, colIdx) => {
            const val = row[k] != null ? row[k] : '';
            html += `<td><input type="text" value="${val}" data-row="${rowIdx}" data-key="${k}" onchange="updateInputData(this)"></td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
}

function updateInputData(inputEl) {
    const rowIdx = parseInt(inputEl.dataset.row);
    const key = inputEl.dataset.key;
    let val = inputEl.value.trim();

    // Try numeric conversion
    const num = Number(val);
    if (val !== '' && !isNaN(num)) {
        val = num;
    }

    if (currentInputData[rowIdx]) {
        currentInputData[rowIdx][key] = val;
    }
}

async function saveCurrentInput() {
    const endpoints = {
        city: '/api/inputs/city',
        zones: '/api/inputs/zones',
        financial: '/api/inputs/financial',
    };

    const url = endpoints[currentInputTab];
    if (!url) return;

    try {
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(currentInputData),
        });
        const result = await res.json();

        if (result.status === 'ok') {
            showToast('success', 'Input data saved successfully!');
        } else {
            showToast('error', 'Save failed: ' + (result.message || 'Unknown error'));
        }
    } catch (err) {
        showToast('error', 'Save failed: ' + err.message);
    }
}

// ---------------------------------------------------------------------------
// Run Model
// ---------------------------------------------------------------------------
async function runModel() {
    const modal = document.getElementById('runModal');
    const statusEl = document.getElementById('runStatus');
    const outputEl = document.getElementById('runOutput');
    const closeBtn = document.getElementById('modalCloseBtn');
    const runBtn = document.getElementById('runModelBtn');
    const statusBadge = document.getElementById('statusBadge');

    // Open modal
    modal.classList.add('visible');
    statusEl.className = 'run-status';
    statusEl.innerHTML = '<div class="spinner"></div><p>Running model... this may take a moment.</p>';
    outputEl.textContent = '';
    closeBtn.style.display = 'none';
    runBtn.classList.add('running');

    // Update status badge
    statusBadge.classList.add('running');
    statusBadge.querySelector('span:last-child').textContent = 'Running...';

    try {
        const res = await fetch('/api/run', { method: 'POST' });
        const result = await res.json();

        if (result.status === 'ok') {
            statusEl.className = 'run-status success';
            statusEl.innerHTML = '<p>✓ Model completed successfully!</p>';
            outputEl.textContent = result.stdout || '(No output)';
            showToast('success', 'Model execution completed successfully!');
        } else {
            statusEl.className = 'run-status error';
            statusEl.innerHTML = '<p>✗ Model execution encountered errors.</p>';
            outputEl.textContent = (result.stderr || '') + '\n' + (result.stdout || '') + '\n' + (result.message || '');
            showToast('error', 'Model execution failed. Check the output for details.');
        }
    } catch (err) {
        statusEl.className = 'run-status error';
        statusEl.innerHTML = '<p>✗ Network error.</p>';
        outputEl.textContent = err.message;
        showToast('error', 'Network error: ' + err.message);
    }

    closeBtn.style.display = 'inline-block';
    runBtn.classList.remove('running');
    statusBadge.classList.remove('running');
    statusBadge.querySelector('span:last-child').textContent = 'Ready';
}

function closeModal() {
    document.getElementById('runModal').classList.remove('visible');
    // Refresh all dashboard data after model run
    refreshAll();
}

// ---------------------------------------------------------------------------
// Toast Notifications
// ---------------------------------------------------------------------------
function showToast(type, message) {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    container.appendChild(toast);

    // Auto-remove after animation
    setTimeout(() => {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
    }, 3500);
}

// ---------------------------------------------------------------------------
// Refresh All Data
// ---------------------------------------------------------------------------
function refreshAll() {
    loadKPIs();
    loadPreviewCharts();
    loadDashFinancialTable();

    // Refresh current open section
    const activeSection = document.querySelector('.section.active');
    if (activeSection) {
        if (activeSection.id === 'section-charts') loadCharts();
        if (activeSection.id === 'section-data') {
            const activeTab = document.querySelector('#section-data .tab.active');
            if (activeTab) activeTab.click();
        }
    }
}

// ---------------------------------------------------------------------------
// Initialisation
// ---------------------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {
    loadKPIs();
    loadPreviewCharts();
    loadDashFinancialTable();
    switchDataTab('demand', document.querySelector('#section-data .tab.active'));
});
