// State Management
const state = {
    rawData: [],
    missingData: [],
    completeness: {
        oerComplete: [],
        fruitComplete: []
    },
    processedData: [],
    filters: {
        psm: 'all',
        region: 'all',
        estate: 'all',
        lmm: 'all',
        startDate: '',
        endDate: ''
    },
    charts: {}
};

// DOM Elements
const elements = {
    dashboardContent: document.getElementById('dashboard-content'),
    fileStatus: document.getElementById('file-status'),
    psmFilter: document.getElementById('psm-filter'),
    regionFilter: document.getElementById('region-filter'),
    estateFilter: document.getElementById('estate-filter'),
    lmmFilter: document.getElementById('lmm-filter'),
    startDateFilter: document.getElementById('start-date-filter'),
    endDateFilter: document.getElementById('end-date-filter'),
    tabBtns: document.querySelectorAll('.tab-btn'),
    tabContents: document.querySelectorAll('.tab-content'),
    kpi: {
        oerBefore: document.getElementById('kpi-oer-before'),
        oerAfter: document.getElementById('kpi-oer-after'),
        oerGain: document.getElementById('kpi-oer-gain'),
        oerMax: document.getElementById('kpi-oer-max'),
        oerMin: document.getElementById('kpi-oer-min'),
        fruitSource: document.getElementById('kpi-fruit-source'),
        lmmStatus: document.getElementById('kpi-lmm-status'),
        missingData: document.getElementById('kpi-missing-data')
    }
};

// Embedded workbook fallback helpers
const EMBEDDED_DEFAULT_WORKBOOK_ID = 'embedded-default-workbook';

function getEmbeddedWorkbookBase64() {
    const element = document.getElementById(EMBEDDED_DEFAULT_WORKBOOK_ID);
    if (!element) return '';
    return element.textContent.replace(/\s+/g, '');
}

function buildWorkbookFromEmbeddedData() {
    const base64 = getEmbeddedWorkbookBase64();
    if (!base64) throw new Error('Embedded workbook payload is missing');
    const binaryString = atob(base64);
    const buffer = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
        buffer[i] = binaryString.charCodeAt(i);
    }
    return XLSX.read(buffer, { type: 'array' });
}

function fallbackToEmbeddedWorkbook(error) {
    console.warn('Falling back to embedded workbook data.', error);
    try {
        const workbook = buildWorkbookFromEmbeddedData();
        processWorkbook(workbook);
        setFileStatus('Embedded workbook loaded (offline)', 'success');
    } catch (embeddedError) {
        console.error('Unable to load embedded workbook:', embeddedError);
        setFileStatus('Failed to load default workbook. Please upload a file.', 'error');
        alert('Failed to load default workbook. Please upload your workbook to continue.');
    }
}

// Auto-load default workbook on page load
document.addEventListener('DOMContentLoaded', preloadDefaultWorkbook);

// Event Listeners
elements.psmFilter.addEventListener('change', (e) => updateFilters('psm', e.target.value));
elements.regionFilter.addEventListener('change', (e) => updateFilters('region', e.target.value));
elements.estateFilter.addEventListener('change', (e) => updateFilters('estate', e.target.value));
elements.lmmFilter.addEventListener('change', (e) => updateFilters('lmm', e.target.value));
elements.startDateFilter.addEventListener('change', (e) => updateFilters('startDate', e.target.value));
elements.endDateFilter.addEventListener('change', (e) => updateFilters('endDate', e.target.value));

elements.tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        elements.tabBtns.forEach(b => {
            b.classList.remove('active', 'bg-white', 'shadow-sm', 'text-slate-700');
            b.classList.add('text-slate-500');
        });
        btn.classList.add('active', 'bg-white', 'shadow-sm', 'text-slate-700');
        btn.classList.remove('text-slate-500');

        const tabId = btn.dataset.tab;
        elements.tabContents.forEach(content => {
            content.classList.add('hidden');
            if (content.id === `tab-${tabId}`) {
                content.classList.remove('hidden');
            }
        });
    });
});

// File Handling
async function preloadDefaultWorkbook() {
    setFileStatus('Loading default workbook...', 'loading');
    try {
        const response = await fetch('data/default.xlsx');
        if (!response.ok) throw new Error(`HTTP ${response.status}`);

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        processWorkbook(workbook);
    } catch (error) {
        fallbackToEmbeddedWorkbook(error);
    }
}

function setFileStatus(message, type = 'info') {
    if (!elements.fileStatus) return;
    elements.fileStatus.classList.remove('hidden');
    const icon = elements.fileStatus.querySelector('i');
    if (icon) {
        if (type === 'success') {
            icon.setAttribute('data-lucide', 'file-check');
            icon.className = 'h-4 w-4 text-green-500';
        } else if (type === 'error') {
            icon.setAttribute('data-lucide', 'alert-circle');
            icon.className = 'h-4 w-4 text-red-500';
        } else {
            icon.setAttribute('data-lucide', 'loader-2');
            icon.className = 'h-4 w-4 text-primary animate-spin';
        }
        if (window.lucide && typeof window.lucide.createIcons === 'function') {
            window.lucide.createIcons();
        }
    }
    const span = elements.fileStatus.querySelector('span');
    if (span) span.textContent = message;
    else elements.fileStatus.textContent = message;
}

function showDashboard() {
    if (elements.dashboardContent) elements.dashboardContent.classList.remove('hidden');
}

async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        processWorkbook(workbook);
    };
    reader.readAsArrayBuffer(file);
}

function processWorkbook(workbook) {
    try {
        const sheetNames = workbook.SheetNames;
        const requiredSheets = ['OER Data Before HFC', 'OER Data After HFC', 'Fruit Mix %'];

        const missingSheets = requiredSheets.filter(s => !sheetNames.includes(s));
        if (missingSheets.length > 0) {
            alert(`Missing sheets: ${missingSheets.join(', ')}. Please check your file.`);
            return;
        }

        let rawOerBefore = XLSX.utils.sheet_to_json(workbook.Sheets['OER Data Before HFC'], { defval: null });
        let rawOerAfter = XLSX.utils.sheet_to_json(workbook.Sheets['OER Data After HFC'], { defval: null });
        let rawFruitMix = XLSX.utils.sheet_to_json(workbook.Sheets['Fruit Mix %'], { defval: null });

        // Filter out SNKM
        rawOerBefore = rawOerBefore.filter(r => r['Estate Code'] !== 'SNKM');
        rawOerAfter = rawOerAfter.filter(r => r['Estate Code'] !== 'SNKM');
        rawFruitMix = rawFruitMix.filter(r => r['Estate Code'] !== 'SNKM');

        const completeness = findIncompleteMills(rawOerBefore, rawOerAfter, rawFruitMix);
        state.missingData = completeness.incomplete;
        state.completeness = completeness.completeness;

        // Process Mapping Sheet
        let mappingData = {};
        if (sheetNames.includes('Mapping')) {
            // Mapping sheet headers are in row 2 (index 1)
            const rawMapping = XLSX.utils.sheet_to_json(workbook.Sheets['Mapping'], { range: 1 });
            rawMapping.forEach(row => {
                if (row['Estate Code']) {
                    mappingData[row['Estate Code']] = {
                        psm: row['PSM'] || 'Unknown',
                        region: row['Region'] || 'Unknown'
                    };
                }
            });
        }

        const oerBefore = transformData(rawOerBefore, 'oer_before');
        const oerAfter = transformData(rawOerAfter, 'oer_after');
        const fruitMix = transformData(rawFruitMix, 'fruit_mix_pct', true);

        state.rawData = mergeDatasets(oerBefore, oerAfter, fruitMix, mappingData);

        initializeFilters();
        updateDashboard();

        showDashboard();
        setFileStatus('Data Loaded', 'success');

    } catch (error) {
        console.error("Error processing file:", error);
        alert(`Error processing file: ${error.message}. Please ensure the format is correct.`);
    }
}

function transformData(data, valueKey, hasCategoryColumn = false) {
    const result = [];

    data.forEach(row => {
        const estate = row['Estate Code'];
        const lmm = row['LMM CPO'] || 'Unknown';
        const category = hasCategoryColumn ? row['Fruit Mix'] : null;

        Object.keys(row).forEach(key => {
            const dateObj = parseDate(key);

            if (dateObj) {
                const year = dateObj.getFullYear();
                const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                const dateStr = `${year}-${month}-01`;

                result.push({
                    estate,
                    lmm,
                    date: dateStr,
                    year: String(year),
                    month: month,
                    category,
                    [valueKey]: parseValue(row[key])
                });
            }
        });
    });
    return result;
}

function parseValue(val) {
    if (val == null || val === '') return 0;
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
        val = val.trim();
        if (val.endsWith('%')) {
            const parsed = parseFloat(val);
            return isNaN(parsed) ? 0 : parsed / 100;
        }
        const parsed = parseFloat(val);
        return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
}

function parseDate(input) {
    if (!input) return null;

    if (!isNaN(input) && Number(input) > 40000 && Number(input) < 60000) {
        const date = new Date(Math.round((Number(input) - 25569) * 86400 * 1000));
        return date;
    }

    const mmmYyRegex = /^([A-Za-z]{3})-(\d{2,4})$/;
    const match = String(input).match(mmmYyRegex);
    if (match) {
        const monthStr = match[1];
        let yearStr = match[2];
        if (yearStr.length === 2) yearStr = '20' + yearStr;

        const monthIndex = new Date(`${monthStr} 1, 2000`).getMonth();
        if (!isNaN(monthIndex)) {
            return new Date(Number(yearStr), monthIndex, 1);
        }
    }

    if (String(input).includes('Code') || String(input).includes('Mix') || String(input).includes('LMM')) return null;

    const date = new Date(input);
    if (!isNaN(date.getTime()) && date.getFullYear() > 1990 && date.getFullYear() < 2050) {
        return date;
    }

    return null;
}

function mergeDatasets(oerBefore, oerAfter, fruitMix, mappingData = {}) {
    const merged = {};
    const getKey = (estate, date) => `${estate}|${date}`;

    oerBefore.forEach(item => {
        const key = getKey(item.estate, item.date);
        if (!merged[key]) {
            const mapInfo = mappingData[item.estate] || { psm: 'Unknown', region: 'Unknown' };
            merged[key] = {
                estate: item.estate,
                date: item.date,
                year: item.year,
                month: item.month,
                psm: mapInfo.psm,
                region: mapInfo.region
            };
        }
        merged[key].oer_before = item.oer_before;
        merged[key].lmm = item.lmm;
    });

    oerAfter.forEach(item => {
        const key = getKey(item.estate, item.date);
        if (!merged[key]) {
            const mapInfo = mappingData[item.estate] || { psm: 'Unknown', region: 'Unknown' };
            merged[key] = {
                estate: item.estate,
                date: item.date,
                year: item.year,
                month: item.month,
                psm: mapInfo.psm,
                region: mapInfo.region
            };
        }
        merged[key].oer_after = item.oer_after;
        if (!merged[key].lmm || merged[key].lmm === 'Unknown') merged[key].lmm = item.lmm;
    });

    fruitMix.forEach(item => {
        const key = getKey(item.estate, item.date);
        if (!merged[key]) {
            const mapInfo = mappingData[item.estate] || { psm: 'Unknown', region: 'Unknown' };
            merged[key] = {
                estate: item.estate,
                date: item.date,
                year: item.year,
                month: item.month,
                lmm: 'Unknown',
                psm: mapInfo.psm,
                region: mapInfo.region
            };
        }

        if (item.category) {
            const cat = String(item.category).toLowerCase();
            if (cat.includes('inti')) merged[key].fruit_inti = item.fruit_mix_pct;
            else if (cat.includes('plasma')) merged[key].fruit_plasma = item.fruit_mix_pct;
            else if (cat.includes('3p') || cat.includes('third') || cat.includes('external')) merged[key].fruit_3p = item.fruit_mix_pct;
        }
    });

    return Object.values(merged);
}

function findIncompleteMills(rawOerBefore, rawOerAfter, rawFruitMix) {
    const getEstates = (rows) => new Set(
        rows
            .map(r => r['Estate Code'])
            .filter(code => code && String(code).trim() !== '')
            .map(code => String(code).trim())
    );

    const countBlankCells = (rows) => {
        const blanks = {};
        rows.forEach(row => {
            const estate = row['Estate Code'];
            if (!estate) return;
            const estateKey = String(estate).trim();
            Object.keys(row).forEach(key => {
                if (parseDate(key)) {
                    const val = row[key];
                    const isBlank = val === null || val === undefined || val === '' || (typeof val === 'string' && val.trim() === '');
                    if (isBlank) {
                        blanks[estateKey] = (blanks[estateKey] || 0) + 1;
                    }
                }
            });
        });
        return blanks;
    };

    const estatesBefore = getEstates(rawOerBefore || []);
    const estatesAfter = getEstates(rawOerAfter || []);
    const estatesFruit = getEstates(rawFruitMix || []);
    const blanksBefore = countBlankCells(rawOerBefore || []);
    const blanksAfter = countBlankCells(rawOerAfter || []);
    const blanksFruit = countBlankCells(rawFruitMix || []);

    const allEstates = new Set([...estatesBefore, ...estatesAfter, ...estatesFruit]);
    const incomplete = [];
    const completeOer = [];
    const completeFruit = [];

    allEstates.forEach(estate => {
        const missing = [];
        if (!estatesBefore.has(estate)) missing.push('Pre-LMM OER');
        if (!estatesAfter.has(estate)) missing.push('Post-LMM OER');
        if (blanksBefore[estate]) missing.push(`Pre-LMM OER: ${blanksBefore[estate]} blank month${blanksBefore[estate] > 1 ? 's' : ''}`);
        if (blanksAfter[estate]) missing.push(`Post-LMM OER: ${blanksAfter[estate]} blank month${blanksAfter[estate] > 1 ? 's' : ''}`);
        if (!estatesFruit.has(estate)) missing.push('Fruit Mix %');
        if (blanksFruit[estate]) missing.push(`Fruit Mix %: ${blanksFruit[estate]} blank month${blanksFruit[estate] > 1 ? 's' : ''}`);

        const oerComplete = estatesBefore.has(estate) && estatesAfter.has(estate) && !blanksBefore[estate] && !blanksAfter[estate];
        const fruitComplete = estatesFruit.has(estate) && !blanksFruit[estate];

        if (oerComplete) completeOer.push(estate);
        if (fruitComplete) completeFruit.push(estate);

        if (missing.length) {
            incomplete.push({
                estate,
                missing,
                details: {
                    oerComplete,
                    fruitComplete,
                    blanks: {
                        before: blanksBefore[estate] || 0,
                        after: blanksAfter[estate] || 0,
                        fruit: blanksFruit[estate] || 0
                    },
                    hasBefore: estatesBefore.has(estate),
                    hasAfter: estatesAfter.has(estate),
                    hasFruit: estatesFruit.has(estate)
                }
            });
        }
    });

    return {
        incomplete: incomplete.sort((a, b) => a.estate.localeCompare(b.estate)),
        completeness: {
            oerComplete: completeOer,
            fruitComplete: completeFruit
        }
    };
}

function initializeFilters() {
    // Populate PSM
    const psms = [...new Set(state.rawData.map(d => d.psm))].filter(Boolean).sort();
    const psmSelect = elements.psmFilter;
    psmSelect.innerHTML = '<option value="all">All PSMs</option>';
    psms.forEach(p => {
        const option = document.createElement('option');
        option.value = p;
        option.textContent = p;
        psmSelect.appendChild(option);
    });

    // Populate Regions (initially all)
    const regions = [...new Set(state.rawData.map(d => d.region))].filter(Boolean).sort();
    const regionSelect = elements.regionFilter;
    regionSelect.innerHTML = '<option value="all">All Regions</option>';
    regions.forEach(r => {
        const option = document.createElement('option');
        option.value = r;
        option.textContent = r;
        regionSelect.appendChild(option);
    });

    // Populate Estates (initially all)
    const estates = [...new Set(state.rawData.map(d => d.estate))].sort();
    const estateSelect = elements.estateFilter;
    estateSelect.innerHTML = '<option value="all">All Mills</option>';
    estates.forEach(e => {
        const option = document.createElement('option');
        option.value = e;
        option.textContent = e;
        estateSelect.appendChild(option);
    });

    // Set default date range (full range)
    const dates = state.rawData.map(d => d.date).sort();
    if (dates.length > 0) {
        const minDate = dates[0].substring(0, 7); // YYYY-MM
        const maxDate = dates[dates.length - 1].substring(0, 7);

        elements.startDateFilter.value = minDate;
        elements.endDateFilter.value = maxDate;
        state.filters.startDate = minDate;
        state.filters.endDate = maxDate;
    }
}

function updateFilters(type, value) {
    state.filters[type] = value;

    // Cascading Logic
    if (type === 'psm') {
        // Filter Regions based on PSM
        const relevantData = value === 'all'
            ? state.rawData
            : state.rawData.filter(d => d.psm === value);

        const validRegions = [...new Set(relevantData.map(d => d.region))].filter(Boolean).sort();

        // Update Region Dropdown
        const regionSelect = elements.regionFilter;
        const currentRegion = regionSelect.value;
        regionSelect.innerHTML = '<option value="all">All Regions</option>';
        validRegions.forEach(r => {
            const option = document.createElement('option');
            option.value = r;
            option.textContent = r;
            regionSelect.appendChild(option);
        });

        // Reset Region if not valid, or keep if valid
        if (value !== 'all' && !validRegions.includes(currentRegion)) {
            state.filters.region = 'all';
            regionSelect.value = 'all';
        }

        // Also update estates
        updateEstateFilter();
    }

    if (type === 'region') {
        updateEstateFilter();
    }

    updateDashboard();
}

function updateEstateFilter() {
    // Filter Estates based on PSM and Region
    let relevantData = state.rawData;
    if (state.filters.psm !== 'all') {
        relevantData = relevantData.filter(d => d.psm === state.filters.psm);
    }
    if (state.filters.region !== 'all') {
        relevantData = relevantData.filter(d => d.region === state.filters.region);
    }

    const validEstates = [...new Set(relevantData.map(d => d.estate))].sort();
    const estateSelect = elements.estateFilter;
    const currentEstate = estateSelect.value;

    estateSelect.innerHTML = '<option value="all">All Mills</option>';
    validEstates.forEach(e => {
        const option = document.createElement('option');
        option.value = e;
        option.textContent = e;
        estateSelect.appendChild(option);
    });

    if (state.filters.estate !== 'all' && !validEstates.includes(currentEstate)) {
        state.filters.estate = 'all';
        estateSelect.value = 'all';
    }
}

function updateDashboard() {
    // Filter Data
    let filtered = state.rawData.filter(d => {
        const matchPsm = state.filters.psm === 'all' || d.psm === state.filters.psm;
        const matchRegion = state.filters.region === 'all' || d.region === state.filters.region;
        const matchEstate = state.filters.estate === 'all' || d.estate === state.filters.estate;
        const matchLmm = state.filters.lmm === 'all' || d.lmm === state.filters.lmm;

        // Date Range Filter
        const dateKey = d.date.substring(0, 7); // YYYY-MM
        const matchDate = (!state.filters.startDate || dateKey >= state.filters.startDate) &&
            (!state.filters.endDate || dateKey <= state.filters.endDate);

        return matchPsm && matchRegion && matchEstate && matchLmm && matchDate;
    });

    filtered.sort((a, b) => new Date(a.date) - new Date(b.date));

    const completeOerSet = new Set(state.completeness?.oerComplete || []);
    const completeFruitSet = new Set(state.completeness?.fruitComplete || []);

    const filteredOerComplete = filtered.filter(d => completeOerSet.has(d.estate));
    const filteredFruitComplete = filtered.filter(d => completeFruitSet.has(d.estate));

    renderMissingDataCard();
    updateKPIs(filteredOerComplete);
    renderCharts(filteredFruitComplete);
    renderOverviewTable(filtered);
    renderRankingTable(filtered);
    renderAnalysis(filteredFruitComplete);
}

function updateKPIs(data) {
    // Calculate Averages
    const validOerBefore = data.filter(d => d.oer_before != null && d.oer_before !== 0);
    const avgOerBefore = validOerBefore.reduce((sum, d) => sum + d.oer_before, 0) / (validOerBefore.length || 1);

    const validOerAfter = data.filter(d => d.oer_after != null && d.oer_after !== 0);
    const avgOerAfter = validOerAfter.reduce((sum, d) => sum + d.oer_after, 0) / (validOerAfter.length || 1);

    // Calculate Max/Min OER (After HFC)
    let maxOer = 0;
    let minOer = 0;

    if (validOerAfter.length > 0) {
        maxOer = Math.max(...validOerAfter.map(d => d.oer_after));
        minOer = Math.min(...validOerAfter.map(d => d.oer_after));
    }

    // OER is already in %, no need to multiply by 100
    elements.kpi.oerBefore.textContent = validOerBefore.length ? avgOerBefore.toFixed(2) + '%' : '-';
    elements.kpi.oerAfter.textContent = validOerAfter.length ? avgOerAfter.toFixed(2) + '%' : '-';

    const gain = (validOerAfter.length ? avgOerAfter : 0) - (validOerBefore.length ? avgOerBefore : 0);
    elements.kpi.oerGain.textContent = (gain > 0 ? '+' : '') + (validOerAfter.length && validOerBefore.length ? gain.toFixed(2) + '%' : '-');
    elements.kpi.oerGain.className = `text-2xl font-bold mt-1 ${gain >= 0 ? 'text-green-600' : 'text-red-600'}`;

    elements.kpi.oerMax.textContent = validOerAfter.length ? maxOer.toFixed(2) + '%' : '-';
    if (elements.kpi.oerMin) {
        elements.kpi.oerMin.textContent = validOerAfter.length ? minOer.toFixed(2) + '%' : '-';
    }

    // Dominant Fruit Source
    let sumInti = 0, sumPlasma = 0, sum3p = 0;
    data.forEach(d => {
        sumInti += d.fruit_inti || 0;
        sumPlasma += d.fruit_plasma || 0;
        sum3p += d.fruit_3p || 0;
    });

    let dominant = 'Inti';
    let maxVal = sumInti;
    if (sumPlasma > maxVal) { dominant = 'Plasma'; maxVal = sumPlasma; }
    if (sum3p > maxVal) { dominant = '3P'; maxVal = sum3p; }

    elements.kpi.fruitSource.textContent = dominant;

    // LMM Status KPI
    const lmmStatuses = [...new Set(data.map(d => d.lmm))];
    if (lmmStatuses.length === 1) {
        elements.kpi.lmmStatus.textContent = lmmStatuses[0];
    } else if (lmmStatuses.length > 1) {
        elements.kpi.lmmStatus.textContent = 'Mixed';
    } else {
        elements.kpi.lmmStatus.textContent = '-';
    }
}

function renderMissingDataCard() {
    const container = elements.kpi.missingData;
    if (!container) return;

    const missing = state.missingData || [];

    if (!state.rawData.length) {
        container.innerHTML = '<p class="text-sm text-slate-500">Upload the workbook to check completeness.</p>';
        return;
    }

    if (missing.length === 0) {
        container.innerHTML = `
            <p class="text-lg font-bold text-green-700">Data Complete</p>
            <div class="mt-3 text-xs text-slate-500 space-y-1">
                <p>✅ Mills without complete OER data are excluded from Avg OER (Before/After).</p>
                <p>✅ Months with OER value 0.00 are treated as non-operational and excluded from averages.</p>
                <p>✅ Mills without complete Fruit Mix data are excluded from the fruit-mix analysis.</p>
            </div>
        `;
        return;
    }

    const listItems = missing.map(item => `
        <li class="flex items-start justify-between bg-slate-50 px-3 py-2 rounded-lg border border-slate-100">
            <div>
                <p class="text-sm font-semibold text-slate-800">${item.estate}</p>
                <p class="text-xs text-slate-500">Missing: ${item.missing.join(', ')}</p>
            </div>
        </li>
    `).join('');

    container.innerHTML = `
        <p class="text-sm text-slate-600 mb-2">${missing.length} mill${missing.length > 1 ? 's' : ''} need data completion</p>
        <ul class="space-y-2 max-h-32 overflow-y-auto">${listItems}</ul>
        <div class="mt-3 text-xs text-slate-500 space-y-1">
            <p>✅ Mills without complete OER data are excluded from Avg OER (Before/After).</p>
            <p>✅ Months with OER value 0.00 are treated as non-operational and excluded from averages.</p>
            <p>✅ Mills without complete Fruit Mix data are excluded from the fruit-mix analysis.</p>
        </div>
    `;
}

function renderCharts(data) {
    if (data.length === 0) return;

    const dateMap = {};
    data.forEach(d => {
        if (!dateMap[d.date]) {
            dateMap[d.date] = {
                date: d.date,
                oer_before_sum: 0, oer_before_count: 0,
                oer_after_sum: 0, oer_after_count: 0,
                inti_sum: 0, plasma_sum: 0, p3_sum: 0, count: 0
            };
        }
        if (d.oer_before != null && d.oer_before !== 0) {
            dateMap[d.date].oer_before_sum += d.oer_before;
            dateMap[d.date].oer_before_count++;
        }
        if (d.oer_after != null && d.oer_after !== 0) {
            dateMap[d.date].oer_after_sum += d.oer_after;
            dateMap[d.date].oer_after_count++;
        }
        dateMap[d.date].inti_sum += d.fruit_inti || 0;
        dateMap[d.date].plasma_sum += d.fruit_plasma || 0;
        dateMap[d.date].p3_sum += d.fruit_3p || 0;
        dateMap[d.date].count++;
    });

    const dates = Object.keys(dateMap).sort((a, b) => new Date(a) - new Date(b));
    if (dates.length === 0) return;

    const minDate = new Date(dates[0]);
    const maxDate = new Date(dates[dates.length - 1]);
    const allMonths = [];

    let currentDate = new Date(minDate);
    currentDate.setDate(1);

    while (currentDate <= maxDate) {
        const year = currentDate.getFullYear();
        const month = String(currentDate.getMonth() + 1).padStart(2, '0');
        const dateStr = `${year}-${month}-01`;

        const matchKey = dates.find(d => d.startsWith(`${year}-${month}`));

        allMonths.push({
            label: currentDate.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }),
            data: matchKey ? dateMap[matchKey] : null
        });

        currentDate.setMonth(currentDate.getMonth() + 1);
    }

    const labels = allMonths.map(m => m.label);

    const ctxOer = document.getElementById('oerTrendChart').getContext('2d');
    if (state.charts.oer) state.charts.oer.destroy();

    state.charts.oer = new Chart(ctxOer, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'OER Pre-LMM',
                    // OER is already %, no multiply by 100
                    data: allMonths.map(m => m.data && m.data.oer_before_count ? (m.data.oer_before_sum / m.data.oer_before_count) : null),
                    borderColor: '#94a3b8',
                    backgroundColor: '#94a3b8',
                    tension: 0.3,
                    pointRadius: 2,
                    spanGaps: true
                },
                {
                    label: 'OER Post-LMM',
                    data: allMonths.map(m => m.data && m.data.oer_after_count ? (m.data.oer_after_sum / m.data.oer_after_count) : null),
                    borderColor: '#0f766e',
                    backgroundColor: '#0f766e',
                    tension: 0.3,
                    pointRadius: 2,
                    spanGaps: true
                },
                {
                    label: 'Target 23%',
                    data: allMonths.map(() => 23),
                    borderColor: '#ef4444',
                    borderDash: [6, 4],
                    fill: false,
                    pointRadius: 0,
                    spanGaps: true
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            scales: {
                y: {
                    min: 15,
                    max: 26,
                    title: { display: true, text: 'OER (%)' },
                    ticks: { stepSize: 1 }
                }
            }
        }
    });

    const ctxFruit = document.getElementById('fruitMixChart').getContext('2d');
    if (state.charts.fruit) state.charts.fruit.destroy();

    state.charts.fruit = new Chart(ctxFruit, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Inti',
                    // Fruit Mix is 0-1, so multiply by 100
                    data: allMonths.map(m => m.data && m.data.count ? (m.data.inti_sum / m.data.count) * 100 : null),
                    backgroundColor: '#0f766e',
                    stack: 'Stack 0'
                },
                {
                    label: 'Plasma',
                    data: allMonths.map(m => m.data && m.data.count ? (m.data.plasma_sum / m.data.count) * 100 : null),
                    backgroundColor: '#f59e0b',
                    stack: 'Stack 0'
                },
                {
                    label: '3P',
                    data: allMonths.map(m => m.data && m.data.count ? (m.data.p3_sum / m.data.count) * 100 : null),
                    backgroundColor: '#94a3b8',
                    stack: 'Stack 0'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            scales: {
                y: { stacked: true, max: 100, title: { display: true, text: 'Percentage (%)' } },
                x: { stacked: true }
            }
        }
    });
}

function renderAnalysis(data) {
    if (data.length === 0) return;

    const toDate = (value) => {
        if (!value) return null;
        const date = new Date(value);
        return Number.isNaN(date.getTime()) ? null : date;
    };

    const last12MonthsData = (() => {
        const dates = data.map(d => toDate(d.date)).filter(Boolean);
        if (!dates.length) return [];

        const maxTime = Math.max(...dates.map(d => d.getTime()));
        const maxDate = new Date(maxTime);
        const minDate = new Date(maxDate);
        minDate.setMonth(minDate.getMonth() - 11);
        minDate.setDate(1);

        return data.filter(d => {
            const dt = toDate(d.date);
            return dt && dt >= minDate && dt <= maxDate;
        });
    })();

    // Calculate Average OER (After HFC) for the last 12 months
    const validOer = last12MonthsData.filter(d => d.oer_after != null && d.oer_after !== 0);
    const avgOer = validOer.length
        ? validOer.reduce((sum, d) => sum + d.oer_after, 0) / validOer.length
        : 0;

    // Update Text
    const currentOerText = document.getElementById('analysis-current-oer');
    const statusText = document.getElementById('analysis-status');

    if (currentOerText) currentOerText.textContent = `Current: ${avgOer.toFixed(2)}% OER`;

    let status = 'Poor';
    if (avgOer >= 19) status = 'Average';
    if (avgOer >= 21) status = 'Good';
    if (avgOer >= 23) status = 'Excellent';
    if (avgOer >= 25) status = 'Exceptional';

    if (statusText) statusText.textContent = `Status: ${status}`;

    const matchesFilters = (record) => {
        if (!record) return false;
        const matchPsm = state.filters.psm === 'all' || record.psm === state.filters.psm;
        const matchRegion = state.filters.region === 'all' || record.region === state.filters.region;
        const matchEstate = state.filters.estate === 'all' || record.estate === state.filters.estate;
        const matchLmm = state.filters.lmm === 'all' || record.lmm === state.filters.lmm;

        const dateKey = record.date ? record.date.substring(0, 7) : '';
        const matchDate = (!state.filters.startDate || dateKey >= state.filters.startDate) &&
            (!state.filters.endDate || dateKey <= state.filters.endDate);

        return matchPsm && matchRegion && matchEstate && matchLmm && matchDate;
    };

    const focusData = (state.rawData || []).filter(matchesFilters);
    const monthlyStats = {};

    focusData.forEach(record => {
        if (record.oer_after == null) return;
        const monthKey = record.date ? record.date.substring(0, 7) : '';
        if (!monthKey) return;

        if (!monthlyStats[monthKey]) {
            monthlyStats[monthKey] = { sum: 0, count: 0 };
        }
        monthlyStats[monthKey].sum += record.oer_after;
        monthlyStats[monthKey].count += 1;
    });

    const monthEntries = Object.entries(monthlyStats)
        .map(([month, stats]) => ({
            month,
            avg: stats.sum / stats.count,
            count: stats.count
        }));

    const peakMonthEntry = monthEntries.reduce((best, entry) => {
        if (!best || entry.avg > best.avg) return entry;
        return best;
    }, null);

    const historicalPeakAvg = peakMonthEntry ? peakMonthEntry.avg : null;

    const hasCurrentOer = validOer.length > 0;
    const gapValue = hasCurrentOer && historicalPeakAvg != null ? historicalPeakAvg - avgOer : null;

    const updateText = (id, text) => {
        const el = document.getElementById(id);
        if (el) el.textContent = text;
    };

    updateText('oer-gap-avg', hasCurrentOer ? `${avgOer.toFixed(2)}%` : '-');
    const formatMonthYear = (monthKey) => {
        if (!monthKey) return '';
        const dateObj = new Date(`${monthKey}-01`);
        if (Number.isNaN(dateObj.getTime())) return monthKey;
        return dateObj.toLocaleString('en-US', { month: 'short', year: 'numeric' });
    };

    updateText('oer-gap-peak', historicalPeakAvg != null ? `${historicalPeakAvg.toFixed(2)}%` : '-');
    updateText('oer-gap-difference', gapValue != null ? `${gapValue.toFixed(2)}%` : '-');
    const peakNote = peakMonthEntry
        ? `Peak month: ${formatMonthYear(peakMonthEntry.month)} (avg of ${peakMonthEntry.count} records)`
        : 'No peak data';
    updateText('oer-gap-peak-note', peakNote);

    // Chart Configuration with extended line
    const ctx = document.getElementById('benchmarkChart').getContext('2d');
    if (state.charts.benchmark) state.charts.benchmark.destroy();

    const labels = ['Poor', 'Average', 'Good', 'Excellent', 'Exceptional'];

    // Floating Bars Data: [min, max]
    const benchmarkData = [
        [17, 19], [19, 21], [21, 23], [23, 25], [25, 27]
    ];

    // Extended line data - TRUE END-TO-END across entire chart
    const lineData = [
        { x: -1.5, y: avgOer },  // Far left edge
        { x: 0, y: avgOer },
        { x: 1, y: avgOer },
        { x: 2, y: avgOer },
        { x: 3, y: avgOer },
        { x: 4, y: avgOer },
        { x: 5.5, y: avgOer }    // Far right edge
    ];

    state.charts.benchmark = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    type: 'line',
                    label: 'Your Performance',
                    data: lineData,
                    borderColor: '#ef4444',
                    borderWidth: 3,
                    pointRadius: 0,
                    pointHoverRadius: 0,
                    tension: 0,
                    order: 0
                },
                {
                    label: 'Benchmark Ranges',
                    data: benchmarkData,
                    backgroundColor: ['#ef4444', '#facc15', '#22c55e', '#2563eb', '#9333ea'],
                    borderWidth: 0,
                    barPercentage: 0.9,
                    categoryPercentage: 0.9,
                    order: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: true },
                tooltip: {
                    callbacks: {
                        label: function (context) {
                            if (context.dataset.type === 'line') {
                                return `Your OER: ${context.parsed.y.toFixed(2)}%`;
                            }
                            const val = context.raw;
                            return `${context.label}: ${val[0]}% - ${val[1]}%`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    min: 16,
                    max: 28,
                    title: { display: true, text: 'Oil Extraction Rate (%)' },
                    grid: { color: '#f1f5f9' }
                },
                x: {
                    grid: { display: false }
                }
            }
        }
    });


    // Correlation Analysis
    const correlations = renderCorrelationAnalysis(data);

    // Stability Analysis
    const stability = renderStabilityAnalysis(data);

    // OER Fluctuation (Last 12 Months)
    renderOerFluctuation(data);

    // Combined Conclusion
    if (correlations && stability) {
        renderCombinedConclusion(data, correlations, stability);
    }

    // AI Recommendations
    renderAIRecommendations(data, avgOer, status);
}

// Helper: Calculate Pearson Correlation
function calculateCorrelation(x, y) {
    const n = x.length;
    if (n === 0) return 0;

    const sumX = x.reduce((a, b) => a + b, 0);
    const sumY = y.reduce((a, b) => a + b, 0);
    const sumXY = x.reduce((sum, xi, i) => sum + xi * y[i], 0);
    const sumX2 = x.reduce((sum, xi) => sum + xi * xi, 0);
    const sumY2 = y.reduce((sum, yi) => sum + yi * yi, 0);

    const numerator = (n * sumXY) - (sumX * sumY);
    const denominator = Math.sqrt(((n * sumX2) - (sumX * sumX)) * ((n * sumY2) - (sumY * sumY)));

    return denominator === 0 ? 0 : numerator / denominator;
}

// Correlation Analysis Rendering (12-Month Rolling Window)
function renderCorrelationAnalysis(data) {
    const validData = data.filter(d => d.oer_after != null && d.fruit_inti != null && d.fruit_plasma != null && d.fruit_3p != null);
    if (validData.length < 3) return;

    // Sort by date descending and take only the last 12 months
    const sortedData = validData.sort((a, b) => new Date(b.date) - new Date(a.date));

    // Get unique months and take the last 12
    const uniqueMonths = [...new Set(sortedData.map(d => d.date.substring(0, 7)))];
    const last12Months = uniqueMonths.slice(0, 12);

    // Filter data to only include the last 12 months
    const last12MonthsData = sortedData.filter(d => last12Months.includes(d.date.substring(0, 7)));

    // Always use last 12 months from latest month
    const dataToUse = last12MonthsData;
    const periodLabel = last12Months.length < 12 ? `${last12Months.length}-Month Period` : '12-Month Rolling';
    const dataPointCount = last12MonthsData.length;

    const oerValues = dataToUse.map(d => d.oer_after);
    const intiValues = dataToUse.map(d => (d.fruit_inti || 0) * 100);
    const plasmaValues = dataToUse.map(d => (d.fruit_plasma || 0) * 100);
    const p3Values = dataToUse.map(d => (d.fruit_3p || 0) * 100);

    const corrInti = calculateCorrelation(intiValues, oerValues);
    const corrPlasma = calculateCorrelation(plasmaValues, oerValues);
    const corr3P = calculateCorrelation(p3Values, oerValues);

    // Simple impact estimation
    const impactInti = corrInti * 0.8;
    const impactPlasma = corrPlasma * 0.8;
    const impact3P = corr3P * 0.8;

    const getStrengthLabel = (corr) => {
        const abs = Math.abs(corr);
        if (abs > 0.5) return 'Strong';
        if (abs > 0.3) return 'Moderate';
        return 'Weak';
    };

    const getStrengthColor = (corr) => {
        const abs = Math.abs(corr);
        if (abs > 0.5) return 'text-green-600';
        if (abs > 0.3) return 'text-yellow-600';
        return 'text-red-600';
    };

    const buildRow = (source, corr, impact) => {
        const positive = corr >= 0;
        const badgeColor = positive ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-700';
        return `
            <tr>
                <td class="px-4 py-3 border-b border-slate-200">${source}</td>
                <td class="px-4 py-3 border-b border-slate-200">
                    <span class="inline-flex px-3 py-1 text-[11px] font-semibold uppercase tracking-wide rounded-full ${badgeColor}">
                        ${positive ? 'Positive' : 'Negative'}
                    </span>
                </td>
                <td class="px-4 py-3 border-b border-slate-200">
                    <span class="font-medium ${getStrengthColor(corr)}">
                        ${corr > 0 ? '+' : ''}${corr.toFixed(2)}
                    </span>
                    <span class="text-xs text-slate-500 ml-2">${getStrengthLabel(corr)}</span>
                </td>
                <td class="px-4 py-3 border-b border-slate-200 ${impact > 0 ? 'text-green-600' : 'text-red-600'} font-medium">
                    ${impact > 0 ? '+' : ''}${impact.toFixed(2)}%
                </td>
            </tr>
        `;
    };

    const correlationHTML = `
        ${buildRow('Inti', corrInti, impactInti)}
        ${buildRow('Plasma', corrPlasma, impactPlasma)}
        ${buildRow('3P (Third Party)', corr3P, impact3P)}
    `;

    const tableBody = document.getElementById('correlation-table-body');
    if (tableBody) tableBody.innerHTML = correlationHTML;

    const periodIndicator = document.getElementById('correlation-period-indicator');
    if (periodIndicator) {
        periodIndicator.textContent = `📊 Analysis Period: ${periodLabel} (${dataPointCount} data points)`;
    }

    // Correlation-Based Recommendations
    let correlationRecommendations = '';

    if (corrInti < 0) {
        correlationRecommendations += `
            <div class="p-4 bg-red-50 border-l-4 border-red-500 rounded">
                <h5 class="font-semibold text-red-800 mb-2">⚠️ Inti (Internal Fruit) Quality Alert</h5>
                <p class="text-sm text-red-700 mb-2">
                    <strong>Issue:</strong> Negative correlation detected (${corrInti.toFixed(2)}). Increasing Inti proportion is associated with <strong>lower</strong> OER.
                </p>
                <p class="text-sm text-red-700 mb-2">
                    <strong>Root Cause:</strong> Likely quality issues with internal fruit supply.
                </p>
                <p class="text-sm text-red-700">
                    <strong>Action Required:</strong>
                </p>
                <ul class="text-sm text-red-700 list-disc list-inside ml-2">
                    <li>Inspect ripeness standards (target 80-90% ripe bunches)</li>
                    <li>Check for fruit damage during harvest and transportation</li>
                    <li>Review internal estate agronomic practices</li>
                    <li>Verify fruit age (freshness from harvest to processing)</li>
                    <li>Assess bunch composition (loose fruit percentage)</li>
                </ul>
            </div>
        `;
    } else if (corrInti > 0.3) {
        correlationRecommendations += `
            <div class="p-4 bg-green-50 border-l-4 border-green-500 rounded">
                <h5 class="font-semibold text-green-800 mb-2">✅ Inti (Internal Fruit) Performing Well</h5>
                <p class="text-sm text-green-700">
                    <strong>Good correlation (+${corrInti.toFixed(2)}):</strong> Your internal fruit quality is contributing positively to OER. 
                    Maintain current harvest and quality standards.
                </p>
            </div>
        `;
    }

    if (corr3P < -0.3) {
        correlationRecommendations += `
            <div class="p-4 bg-yellow-50 border-l-4 border-yellow-500 rounded mt-3">
                <h5 class="font-semibold text-yellow-800 mb-2">⚠️ Third-Party (3P) Fruit Quality Concern</h5>
                <p class="text-sm text-yellow-700 mb-2">
                    <strong>Issue:</strong> Strong negative correlation (${corr3P.toFixed(2)}). Higher 3P proportion correlates with lower OER.
                </p>
                <p class="text-sm text-yellow-700">
                    <strong>Recommendation:</strong> Implement stricter quality acceptance criteria for third-party suppliers. 
                    Consider reducing 3P dependency and increasing Inti/Plasma proportion.
                </p>
            </div>
        `;
    }

    // Generate Correlation Conclusion
    let correlationConclusion = '';
    if (corrInti < 0) {
        correlationConclusion = `
            <p class="mb-2"><strong class="text-red-600">⚠️ Negative Inti Correlation:</strong> Internal fruit quality is negatively impacting OER. This is a critical issue as Inti should be the highest quality source.</p>
        `;
    } else if (corrInti > 0.5) {
        correlationConclusion = `
            <p class="mb-2"><strong class="text-green-600">✅ Strong Positive Inti Correlation:</strong> Internal fruit is driving OER performance. Maintain current estate practices.</p>
        `;
    }

    if (corr3P < -0.3) {
        correlationConclusion += `
            <p><strong class="text-yellow-600">⚠️ 3P Quality Drag:</strong> Third-party fruit is significantly reducing overall OER. Quality control measures are needed.</p>
        `;
    }

    if (!correlationConclusion) {
        correlationConclusion = `<p class="text-slate-500">No significant negative correlations detected. Fruit mix impact on OER appears stable.</p>`;
    }

    const conclusionContainer = document.getElementById('correlation-conclusion');
    if (conclusionContainer) conclusionContainer.innerHTML = correlationConclusion;

    // Return correlation data for combined conclusion
    return { corrInti, corrPlasma, corr3P };
}

// Fruit Mix Stability Analysis (12-Month Rolling Window)
function renderStabilityAnalysis(data) {
    // --- 12-Month Rolling Window Logic ---
    const sortedData = [...data].sort((a, b) => new Date(b.date) - new Date(a.date));
    const uniqueMonths = [...new Set(sortedData.map(d => d.date.substring(0, 7)))].slice(0, 12);
    const analysisData = sortedData.filter(d => uniqueMonths.includes(d.date.substring(0, 7)));

    const indicatorEl = document.getElementById('stability-period-indicator');
    if (indicatorEl) {
        const periodText = uniqueMonths.length < 12 ? `${uniqueMonths.length}-Month Period` : '12-Month Rolling';
        indicatorEl.textContent = `📊 Analysis Period: ${periodText} (${analysisData.length} data points)`;
    }

    // Calculate statistics for each fruit source
    // Inti
    const intiValues = analysisData.filter(d => d.fruit_inti != null && d.fruit_inti > 0)
        .map(d => d.fruit_inti * 100); // Convert to percentage
    const meanInti = intiValues.length > 0 ? intiValues.reduce((a, b) => a + b, 0) / intiValues.length : 0;
    const varianceInti = intiValues.length > 0 ? intiValues.reduce((a, b) => a + Math.pow(b - meanInti, 2), 0) / intiValues.length : 0;
    const sdInti = Math.sqrt(varianceInti);
    const cvInti = meanInti > 0 ? (sdInti / meanInti) * 100 : 0;

    // Plasma
    const plasmaValues = analysisData.filter(d => d.fruit_plasma != null && d.fruit_plasma > 0)
        .map(d => d.fruit_plasma * 100);
    const meanPlasma = plasmaValues.length > 0 ? plasmaValues.reduce((a, b) => a + b, 0) / plasmaValues.length : 0;
    const variancePlasma = plasmaValues.length > 0 ? plasmaValues.reduce((a, b) => a + Math.pow(b - meanPlasma, 2), 0) / plasmaValues.length : 0;
    const sdPlasma = Math.sqrt(variancePlasma);
    const cvPlasma = meanPlasma > 0 ? (sdPlasma / meanPlasma) * 100 : 0;

    // 3P (Third Party)
    const p3Values = analysisData.filter(d => d.fruit_3p != null && d.fruit_3p > 0)
        .map(d => d.fruit_3p * 100);
    const mean3P = p3Values.length > 0 ? p3Values.reduce((a, b) => a + b, 0) / p3Values.length : 0;
    const variance3P = p3Values.length > 0 ? p3Values.reduce((a, b) => a + Math.pow(b - mean3P, 2), 0) / p3Values.length : 0;
    const sd3P = Math.sqrt(variance3P);
    const cv3P = mean3P > 0 ? (sd3P / mean3P) * 100 : 0;

    const getStabilityLevel = (cv) => {
        if (cv < 10) return { level: 'Excellent', color: 'text-green-600', bgColor: 'bg-green-50' };
        if (cv < 20) return { level: 'Good', color: 'text-blue-600', bgColor: 'bg-blue-50' };
        if (cv < 30) return { level: 'Moderate', color: 'text-yellow-600', bgColor: 'bg-yellow-50' };
        return { level: 'Poor', color: 'text-red-600', bgColor: 'bg-red-50' };
    };

    const stabilityInti = getStabilityLevel(cvInti);
    const stabilityPlasma = getStabilityLevel(cvPlasma);
    const stability3P = getStabilityLevel(cv3P);

    const stabilityHTML = `
        <tr>
            <td class="px-4 py-3 border-b border-slate-200">Inti</td>
            <td class="px-4 py-3 border-b border-slate-200">${meanInti.toFixed(1)}%</td>
            <td class="px-4 py-3 border-b border-slate-200">${sdInti.toFixed(2)}%</td>
            <td class="px-4 py-3 border-b border-slate-200 font-medium">${cvInti.toFixed(1)}%</td>
            <td class="px-4 py-3 border-b border-slate-200">
                <span class="px-2 py-1 text-xs font-semibold rounded ${stabilityInti.bgColor} ${stabilityInti.color}">
                    ${stabilityInti.level}
                </span>
            </td>
        </tr>
        <tr>
            <td class="px-4 py-3 border-b border-slate-200">Plasma</td>
            <td class="px-4 py-3 border-b border-slate-200">${meanPlasma.toFixed(1)}%</td>
            <td class="px-4 py-3 border-b border-slate-200">${sdPlasma.toFixed(2)}%</td>
            <td class="px-4 py-3 border-b border-slate-200 font-medium">${cvPlasma.toFixed(1)}%</td>
            <td class="px-4 py-3 border-b border-slate-200">
                <span class="px-2 py-1 text-xs font-semibold rounded ${stabilityPlasma.bgColor} ${stabilityPlasma.color}">
                    ${stabilityPlasma.level}
                </span>
            </td>
        </tr>
        <tr>
            <td class="px-4 py-3">3P (Third Party)</td>
            <td class="px-4 py-3">${mean3P.toFixed(1)}%</td>
            <td class="px-4 py-3">${sd3P.toFixed(2)}%</td>
            <td class="px-4 py-3 font-medium">${cv3P.toFixed(1)}%</td>
            <td class="px-4 py-3">
                <span class="px-2 py-1 text-xs font-semibold rounded ${stability3P.bgColor} ${stability3P.color}">
                    ${stability3P.level}
                </span>
            </td>
        </tr>
    `;

    const tableBody = document.getElementById('stability-table-body');
    if (tableBody) tableBody.innerHTML = stabilityHTML;

    // Generate Stability Conclusion
    let stabilityConclusion = '';
    const avgCV = (cvInti + cvPlasma + cv3P) / 3;

    if (avgCV < 15) {
        stabilityConclusion = `<p><strong class="text-green-600">✅ Excellent Stability:</strong> Fruit sourcing is consistent (Avg CV: ${avgCV.toFixed(1)}%). Predictable supply supports stable mill operations.</p>`;
    } else if (avgCV < 25) {
        stabilityConclusion = `<p><strong class="text-blue-600">ℹ️ Good Stability:</strong> Moderate consistency (Avg CV: ${avgCV.toFixed(1)}%). Some seasonal fluctuation exists but is manageable.</p>`;
    } else {
        stabilityConclusion = `<p><strong class="text-red-600">⚠️ High Variability:</strong> Significant supply fluctuations detected (Avg CV: ${avgCV.toFixed(1)}%). This makes OER prediction difficult.</p>`;
    }

    if (cvInti > 25) {
        stabilityConclusion += `<p class="mt-2"><strong class="text-red-600">⚠️ Inti Instability:</strong> Internal fruit supply is highly variable (${cvInti.toFixed(1)}% CV). Review harvesting schedules.</p>`;
    }

    const conclusionContainer = document.getElementById('stability-conclusion');
    if (conclusionContainer) conclusionContainer.innerHTML = stabilityConclusion;

    // Return stability data for combined conclusion
    return { cvInti, cvPlasma, cv3P, avgCV };
}

// OER Fluctuation (Last 12 Months)
function renderOerFluctuation(data) {
    const chartEl = document.getElementById('oerFluctuationChart');
    const sdEl = document.getElementById('oer-fluctuation-sd');
    const cvEl = document.getElementById('oer-fluctuation-cv');
    const insightEl = document.getElementById('oer-fluctuation-insight');
    const periodEl = document.getElementById('oer-fluctuation-period');

    if (!chartEl || !sdEl || !cvEl || !insightEl || !periodEl) return;

    const valid = data
        .filter(d => d.oer_after != null && d.oer_after !== 0 && d.date)
        .sort((a, b) => new Date(a.date) - new Date(b.date));

    const resetCards = (message) => {
        sdEl.textContent = '-';
        cvEl.textContent = '-';
        insightEl.className = 'mt-1 px-3 py-2 rounded-lg text-sm font-medium bg-slate-100 text-slate-700';
        insightEl.textContent = message;
    };

    if (valid.length === 0) {
        periodEl.textContent = 'No data';
        resetCards('No data available for volatility check.');
        if (state.charts.oerFluctuation) {
            state.charts.oerFluctuation.destroy();
            delete state.charts.oerFluctuation;
        }
        return;
    }

    const monthBuckets = {};
    valid.forEach(d => {
        const key = d.date.substring(0, 7);
        if (!monthBuckets[key]) monthBuckets[key] = { sum: 0, count: 0 };
        monthBuckets[key].sum += d.oer_after;
        monthBuckets[key].count++;
    });

    // Build monthly averages for all mills (unfiltered)
    const allValid = (state.rawData || [])
        .filter(d => d.oer_after != null && d.oer_after !== 0 && d.date);
    const allMonthBuckets = {};
    allValid.forEach(d => {
        const key = d.date.substring(0, 7);
        if (!allMonthBuckets[key]) allMonthBuckets[key] = { sum: 0, count: 0 };
        allMonthBuckets[key].sum += d.oer_after;
        allMonthBuckets[key].count++;
    });

    const monthKeys = Object.keys(monthBuckets).sort();
    if (monthKeys.length === 0) {
        periodEl.textContent = 'No data';
        resetCards('No data available for volatility check.');
        return;
    }

    const lastMonthKey = monthKeys[monthKeys.length - 1];
    const lastDate = new Date(`${lastMonthKey}-01T00:00:00`);
    const startDate = new Date(lastDate);
    startDate.setMonth(startDate.getMonth() - 11);

    const filteredKeys = monthKeys.filter(key => {
        const dt = new Date(`${key}-01T00:00:00`);
        return dt >= startDate;
    });

    const monthData = filteredKeys.map(key => {
        const bucket = monthBuckets[key];
        const avg = bucket.sum / bucket.count;
        const dateObj = new Date(`${key}-01T00:00:00`);
        return {
            key,
            label: dateObj.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }),
            value: avg
        };
    });

    if (monthData.length === 0) {
        periodEl.textContent = 'No recent data';
        resetCards('Not enough recent months to analyze.');
        return;
    }

    periodEl.textContent = `${monthData[0].label} - ${monthData[monthData.length - 1].label}`;

    const values = monthData.map(m => m.value);
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const variance = values.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (values.length || 1);
    const sd = Math.sqrt(variance);
    const cv = mean > 0 ? (sd / mean) * 100 : 0;

    sdEl.textContent = `${sd.toFixed(2)}%`;
    cvEl.textContent = `${cv.toFixed(1)}%`;

    let status = 'Stable';
    let bgClass = 'bg-green-50';
    let textClass = 'text-green-700';
    let detail = 'Low month-to-month swing; OER is steady.';

    if (cv >= 6) {
        status = 'Volatile';
        bgClass = 'bg-red-50';
        textClass = 'text-red-700';
        detail = 'High month-to-month fluctuation; investigate operational consistency.';
    } else if (cv >= 3) {
        status = 'Moderate';
        bgClass = 'bg-yellow-50';
        textClass = 'text-yellow-700';
        detail = 'Some variability present; monitor for drift and tighten controls.';
    }

    insightEl.className = `mt-1 px-3 py-2 rounded-lg text-sm font-semibold ${bgClass} ${textClass}`;
    insightEl.textContent = `${status} OER: ${detail}`;

    const ctx = chartEl.getContext('2d');
    if (state.charts.oerFluctuation) state.charts.oerFluctuation.destroy();

    const avgLine = monthData.map(() => mean);

    const datasets = [
        {
            label: 'Avg OER (After HFC)',
            data: monthData.map(m => m.value),
            borderColor: '#0f766e',
            backgroundColor: 'rgba(15, 118, 110, 0.1)',
            tension: 0.3,
            fill: false,
            pointRadius: 3
        },
        {
            label: 'Recent Avg OER',
            data: avgLine,
            borderColor: '#9ca3af',
            borderWidth: 2,
            borderDash: [6, 6],
            pointRadius: 0,
            fill: false
        }
    ];

    const allMonthLine = monthData.map(m => {
        const bucket = allMonthBuckets[m.key];
        if (!bucket) return null;
        return bucket.sum / bucket.count;
    });
    if (allMonthLine.some(v => v !== null)) {
        datasets.push({
            label: 'All Mills Avg OER (Monthly)',
            data: allMonthLine,
            borderColor: '#ef4444',
            borderWidth: 2,
            borderDash: [4, 2],
            pointRadius: 0,
            fill: false,
            spanGaps: true
        });
    }

    state.charts.oerFluctuation = new Chart(ctx, {
        type: 'line',
        data: {
            labels: monthData.map(m => m.label),
            datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: true },
                tooltip: {
                    callbacks: {
                        label: (context) => `${context.parsed.y.toFixed(2)}%`
                    }
                }
            },
            scales: {
                y: {
                    min: 15,
                    max: 26,
                    title: { display: true, text: 'OER (%)' },
                    ticks: { stepSize: 1 }
                },
                x: { title: { display: true, text: 'Period' } }
            }
        }
    });
}

// Combined Conclusion: Analyze correlations and stability to identify 5 key issues
function renderCombinedConclusion(data, correlations, stability) {
    const container = document.getElementById('combined-conclusion-body');
    if (!container) return;

    const issues = [];
    const { corrInti = 0, corrPlasma = 0, corr3P = 0 } = correlations || {};
    const { cvInti = 0, cvPlasma = 0, cv3P = 0, avgCV = 0 } = stability || {};

    // Issue 1: Negative Inti Correlation
    if (corrInti < 0) {
        issues.push({
            title: '⚠️ Negative Internal Fruit Correlation',
            problem: `Internal fruit (Inti) shows a negative correlation with OER (${corrInti.toFixed(2)}). This means increasing Inti proportion is associated with <strong>lower</strong> OER, which is abnormal.`,
            resolution: `<strong>Immediate Actions:</strong>
                <ul class="list-disc ml-6 mt-2 space-y-1">
                    <li>Inspect ripeness standards at internal estates (target 80-90% ripe bunches)</li>
                    <li>Check for fruit damage during harvest and transportation</li>
                    <li>Review harvesting timing - ensure fruits are processed within 24 hours</li>
                    <li>Assess bunch quality and loose fruit percentage</li>
                    <li>Conduct estate-level quality audits to identify root causes</li>
                </ul>`,
            severity: 'high'
        });
    }

    // Issue 2: High 3P Negative Correlation
    if (corr3P < -0.3) {
        issues.push({
            title: '⚠️ Third-Party Fruit Quality Concerns',
            problem: `Third-party (3P) fruit shows strong negative correlation (${corr3P.toFixed(2)}). Higher 3P proportion significantly reduces OER.`,
            resolution: `<strong>Recommended Strategies:</strong>
                <ul class="list-disc ml-6 mt-2 space-y-1">
                    <li>Implement stricter quality acceptance criteria for 3P suppliers</li>
                    <li>Conduct supplier quality audits and provide training</li>
                    <li>Introduce quality-based pricing incentives</li>
                    <li>Consider reducing dependency on low-quality 3P sources</li>
                    <li>Increase Inti/Plasma proportion where possible</li>
                </ul>`,
            severity: 'high'
        });
    }

    // Issue 3: High CV in Any Source
    const highCVSources = [];
    if (cvInti > 30) highCVSources.push(`Inti (${cvInti.toFixed(1)}% CV)`);
    if (cvPlasma > 30) highCVSources.push(`Plasma (${cvPlasma.toFixed(1)}% CV)`);
    if (cv3P > 30) highCVSources.push(`3P (${cv3P.toFixed(1)}% CV)`);

    if (highCVSources.length > 0) {
        issues.push({
            title: '📊 High Supply Variability Detected',
            problem: `The following sources show unstable supply patterns: <strong>${highCVSources.join(', ')}</strong>. High CV (>30%) indicates unpredictable sourcing.`,
            resolution: `<strong>Stabilization Measures:</strong>
                <ul class="list-disc ml-6 mt-2 space-y-1">
                    <li>Establish long-term supply contracts with fixed proportions</li>
                    <li>Implement buffer inventory strategies</li>
                    <li>Review estate harvesting schedules and crop planning</li>
                    <li>For 3P sources: negotiate more consistent delivery schedules</li>
                    <li>Monitor monthly to detect early signs of instability</li>
                </ul>`,
            severity: 'medium'
        });
    }

    // Issue 4: Overall Fruit Mix Instability
    if (avgCV > 25) {
        issues.push({
            title: '🔄 Inconsistent Fruit Mix Patterns',
            problem: `Average CV across all sources is ${avgCV.toFixed(1)}%, indicating significant month-to-month fluctuations in fruit mix composition.`,
            resolution: `<strong>Strategic Planning:</strong>
                <ul class="list-disc ml-6 mt-2 space-y-1">
                    <li>Develop annual sourcing plan with target proportions</li>
                    <li>Coordinate with estate managers for predictable harvest schedules</li>
                    <li>Implement demand forecasting to anticipate supply needs</li>
                    <li>Create contingency plans for supply disruptions</li>
                    <li>Review and adjust sourcing strategy quarterly</li>
                </ul>`,
            severity: 'medium'
        });
    }

    // Issue 5: Weak Positive Correlations
    if (corrInti >= 0 && corrInti < 0.3 && corrPlasma >= 0 && corrPlasma < 0.3) {
        issues.push({
            title: 'ℹ️ Weak Correlation - Optimization Opportunity',
            problem: `Both Inti (${corrInti.toFixed(2)}) and Plasma (${corrPlasma.toFixed(2)}) show weak positive correlations. While not negative, there's room for improvement.`,
            resolution: `<strong>Optimization Steps:</strong>
                <ul class="list-disc ml-6 mt-2 space-y-1">
                    <li>Analyze quality differences between high-OER and low-OER periods</li>
                    <li>Standardize fruit acceptance criteria across all estates</li>
                    <li>Invest in agronomic practices to improve fruit quality</li>
                    <li>Conduct training on optimal harvesting practices</li>
                    <li>Track fruit quality metrics alongside OER for deeper insights</li>
                </ul>`,
            severity: 'low'
        });
    }

    // If no issues, show positive message
    if (issues.length === 0) {
        container.innerHTML = `
            <tr>
                <td colspan="3" class="px-4 py-6 text-center bg-green-50 text-green-800 rounded-lg">
                    <h4 class="font-bold text-lg mb-2">✅ Excellent Performance!</h4>
                    <p class="text-green-700">
                        No major concerns detected. Your fruit sourcing shows positive correlations with OER and stable supply patterns.
                        Continue monitoring to maintain this performance.
                    </p>
                </td>
            </tr>
        `;
        return;
    }

    //Limit to top 5 issues by severity
    const prioritizedIssues = issues
        .sort((a, b) => {
            const severityOrder = { high: 0, medium: 1, low: 2 };
            return severityOrder[a.severity] - severityOrder[b.severity];
        })
        .slice(0, 5);

    // Render issues as table rows
    let htmlContent = '';
    prioritizedIssues.forEach((issue, index) => {
        const severityClass = issue.severity === 'high' ? 'text-red-600 bg-red-50' :
            issue.severity === 'medium' ? 'text-yellow-600 bg-yellow-50' : 'text-blue-600 bg-blue-50';

        const severityLabel = issue.severity.charAt(0).toUpperCase() + issue.severity.slice(1);

        htmlContent += `
            <tr class="align-top">
                <td class="px-4 py-3 font-semibold text-slate-500">${index + 1}</td>
                <td class="px-4 py-3">
                    <div class="mb-1">
                        <span class="inline-block px-2 py-0.5 text-xs font-bold rounded mb-1 ${severityClass}">${severityLabel} Priority</span>
                    </div>
                    <h4 class="font-bold text-slate-800 mb-1">${issue.title}</h4>
                    <div class="text-sm text-slate-600">
                        <strong>Problem:</strong> ${issue.problem}
                    </div>
                </td>
                <td class="px-4 py-3">
                    <div class="text-sm text-slate-600">
                        ${issue.resolution}
                    </div>
                </td>
            </tr>
        `;
    });

    container.innerHTML = htmlContent;
}


function renderOverviewTable(data) {
    const tableBody = document.getElementById('overview-table-body');
    if (!tableBody) return;

    if (data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="6" class="px-4 py-3 text-center text-slate-500">No data available</td></tr>';
        return;
    }

    // Group by Estate
    const estateData = {};
    data.forEach(d => {
        if (!estateData[d.estate]) {
            estateData[d.estate] = {
                psm: d.psm,
                region: d.region,
                estate: d.estate,
                lmm: d.lmm,
                oer_after_sum: 0,
                oer_after_count: 0,
                fruit_inti: 0,
                fruit_plasma: 0,
                fruit_3p: 0,
                fruit_count: 0
            };
        }
        if (d.oer_after != null) {
            estateData[d.estate].oer_after_sum += d.oer_after;
            estateData[d.estate].oer_after_count++;
        }
        // Accumulate fruit mix data
        if (d.fruit_inti != null || d.fruit_plasma != null || d.fruit_3p != null) {
            estateData[d.estate].fruit_inti += d.fruit_inti || 0;
            estateData[d.estate].fruit_plasma += d.fruit_plasma || 0;
            estateData[d.estate].fruit_3p += d.fruit_3p || 0;
            estateData[d.estate].fruit_count++;
        }
    });

    // Sort by avg OER descending (highest to lowest)
    const rows = Object.values(estateData).sort((a, b) => {
        const avgA = a.oer_after_count ? (a.oer_after_sum / a.oer_after_count) : 0;
        const avgB = b.oer_after_count ? (b.oer_after_sum / b.oer_after_count) : 0;
        return avgB - avgA; // Descending order
    });

    tableBody.innerHTML = rows.map(row => {
        const avgOerAfter = row.oer_after_count ? (row.oer_after_sum / row.oer_after_count) : 0;

        // Calculate dominant fruit source
        let dominantFruit = '-';
        if (row.fruit_count > 0) {
            const fruits = [
                { name: 'Inti', value: row.fruit_inti },
                { name: 'Plasma', value: row.fruit_plasma },
                { name: '3P', value: row.fruit_3p }
            ];
            const max = fruits.reduce((prev, current) => (prev.value > current.value) ? prev : current);
            dominantFruit = max.value > 0 ? max.name : '-';
        }

        return `
            <tr class="hover:bg-slate-50 transition-colors">
                <td class="px-4 py-3 border-b border-slate-100">${row.psm || '-'}</td>
                <td class="px-4 py-3 border-b border-slate-100">${row.region || '-'}</td>
                <td class="px-4 py-3 border-b border-slate-100 font-medium text-slate-700">${row.estate}</td>
                <td class="px-4 py-3 border-b border-slate-100 text-xs">
                    <span class="px-2 py-1 rounded-full ${row.lmm === 'LMM' ? 'bg-teal-50 text-teal-700' : 'bg-slate-100 text-slate-600'}">
                        ${row.lmm}
                    </span>
                </td>
                <td class="px-4 py-3 border-b border-slate-100 text-right font-medium">${avgOerAfter > 0 ? avgOerAfter.toFixed(2) + '%' : '-'}</td>
                <td class="px-4 py-3 border-b border-slate-100 text-center">${dominantFruit}</td>
            </tr>
        `;
    }).join('');
}

function getRollingWindow(records, months = 12) {
    if (!records.length) return [];
    const entries = records.filter(r => r && r.date);
    if (!entries.length) return [];

    const sorted = [...entries].sort((a, b) => new Date(b.date) - new Date(a.date));
    const uniqueMonths = [...new Set(sorted.map(d => d.date.substring(0, 7)))].slice(0, months);
    if (!uniqueMonths.length) return [];

    return sorted.filter(d => uniqueMonths.includes(d.date.substring(0, 7)));
}

function renderRankingTable(data) {
    const tableBody = document.getElementById('ranking-table-body');
    const stabilityBody = document.getElementById('stability-ranking-body');
    if (!tableBody) return;

    if (data.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="3" class="px-4 py-3 text-center text-slate-500">No data available for the selected filters.</td></tr>';
        if (stabilityBody) {
            stabilityBody.innerHTML = '<tr><td colspan="3" class="px-4 py-3 text-center text-slate-500">No data available for the selected filters.</td></tr>';
        }
        return;
    }

    const estateStats = {};
    data.forEach(d => {
        if (!d.estate || d.oer_after == null || d.oer_after === 0) return;
        if (!estateStats[d.estate]) {
            estateStats[d.estate] = { sum: 0, count: 0 };
        }
        estateStats[d.estate].sum += d.oer_after;
        estateStats[d.estate].count++;
    });

    const rows = Object.entries(estateStats)
        .map(([estate, stats]) => ({
            estate,
            avg: stats.count ? (stats.sum / stats.count) : 0
        }))
        .sort((a, b) => b.avg - a.avg);

    if (rows.length === 0) {
        tableBody.innerHTML = '<tr><td colspan="3" class="px-4 py-3 text-center text-slate-500">No valid OER readings found.</td></tr>';
        return;
    }

    tableBody.innerHTML = rows.map((row, index) => `
        <tr class="border-b border-slate-100 text-xs text-slate-600">
            <td class="px-4 py-3">${index + 1}</td>
            <td class="px-4 py-3 font-semibold text-slate-800">${row.estate}</td>
            <td class="px-4 py-3 text-right font-bold text-slate-900">${row.avg.toFixed(2)}%</td>
        </tr>
    `).join('');

    if (!stabilityBody) return;

    const windowedData = getRollingWindow(data);
    if (windowedData.length === 0) {
        stabilityBody.innerHTML = '<tr><td colspan="3" class="px-4 py-3 text-center text-slate-500">No recent data for stability scoring.</td></tr>';
        return;
    }

    const stabilityStats = {};
    windowedData.forEach(d => {
        if (!d.estate || d.oer_after == null) return;
        if (!stabilityStats[d.estate]) {
            stabilityStats[d.estate] = { values: [] };
        }
        stabilityStats[d.estate].values.push(d.oer_after);
    });

    const stabilityRows = Object.entries(stabilityStats)
        .map(([estate, stats]) => {
            const values = stats.values;
            if (!values.length) return null;
            const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
            const variance = values.reduce((sum, value) => sum + Math.pow(value - mean, 2), 0) / values.length;
            const sd = Math.sqrt(variance);
            const cv = mean > 0 ? (sd / mean) * 100 : 0;
            const score = 0.5 * sd + 0.5 * cv;
            return { estate, score };
        })
        .filter(Boolean)
        .sort((a, b) => a.score - b.score || a.estate.localeCompare(b.estate));

    if (stabilityRows.length === 0) {
        stabilityBody.innerHTML = '<tr><td colspan="3" class="px-4 py-3 text-center text-slate-500">No valid stability data available.</td></tr>';
        return;
    }

    stabilityBody.innerHTML = stabilityRows.map((row, index) => `
        <tr class="border-b border-slate-100 text-xs text-slate-600">
            <td class="px-4 py-3">${index + 1}</td>
            <td class="px-4 py-3 font-semibold text-slate-800">${row.estate}</td>
            <td class="px-4 py-3 text-right font-bold text-slate-900">${row.score.toFixed(2)}</td>
        </tr>
    `).join('');

    const bestAverage = rows[0]?.estate || 'N/A';
    const worstAverage = rows[rows.length - 1]?.estate || 'N/A';
    const bestStability = stabilityRows[0]?.estate;
    const worstStability = stabilityRows[stabilityRows.length - 1]?.estate;

    const bestOverall = bestStability || bestAverage;
    const worstOverall = worstAverage || worstStability || 'N/A';

    const insightsEl = document.getElementById('ranking-ai-insights-text');
    if (insightsEl) {
        const positiveName = bestOverall || 'N/A';
        const negativeName = worstOverall || 'N/A';
        insightsEl.innerHTML = `
            <ul class="space-y-1 list-disc list-inside">
                <li><strong>${positiveName}</strong> shows the most consistent leadership across average OER and volatility.</li>
                <li><strong>${negativeName}</strong> is the biggest opportunity for improvement, with lower average OER and/or higher variability.</li>
            </ul>
        `;
    }
}

function renderAIRecommendations(data, avgOer, status) {
    // Performance Summary
    const summaryContainer = document.getElementById('performance-summary');
    if (summaryContainer) {
        let summaryHTML = `
            <div class="flex items-center justify-between">
                <div>
                    <h4 class="font-semibold text-slate-800 mb-1">Current Performance: ${avgOer.toFixed(2)}% OER</h4>
                    <p class="text-sm text-slate-600">Status: <span class="font-medium">${status}</span></p>
                </div>
                <div class="text-right">
                    <p class="text-xs text-slate-500">Based on ${data.length} data points</p>
                </div>
            </div>
        `;
        summaryContainer.innerHTML = summaryHTML;
    }

    // AI Recommendations
    const recBody = document.getElementById('ai-recommendations-body');
    if (!recBody) return;

    let recommendations = '';

    // Performance-based recommendations
    if (avgOer < 19) {
        recommendations += `
            <div class="p-4 bg-red-50 border-l-4 border-red-500 rounded mb-3">
                <h5 class="font-semibold text-red-800 mb-2">🚨 Critical: Below Industry Average</h5>
                <p class="text-sm text-red-700 mb-2">Your OER (${avgOer.toFixed(2)}%) is below industry standards. Immediate action required.</p>
                <ul class="text-sm text-red-700 list-disc list-inside space-y-1">
                    <li>Conduct comprehensive mill audit (sterilization, pressing, clarification)</li>
                    <li>Review fruit quality standards and ripeness criteria</li>
                    <li>Check equipment maintenance schedules and efficiency</li>
                    <li>Analyze process losses at each stage</li>
                </ul>
            </div>
        `;
    } else if (avgOer < 21) {
        recommendations += `
            <div class="p-4 bg-yellow-50 border-l-4 border-yellow-500 rounded mb-3">
                <h5 class="font-semibold text-yellow-800 mb-2">⚠️ Room for Improvement</h5>
                <p class="text-sm text-yellow-700 mb-2">Your OER (${avgOer.toFixed(2)}%) is average. Target 21%+ for better performance.</p>
                <ul class="text-sm text-yellow-700 list-disc list-inside space-y-1">
                    <li>Optimize sterilization parameters (temperature, pressure, time)</li>
                    <li>Improve fruit ripeness selection (target 80-90% ripe)</li>
                    <li>Review pressing efficiency and screw press settings</li>
                </ul>
            </div>
        `;
    } else if (avgOer < 23) {
        recommendations += `
            <div class="p-4 bg-blue-50 border-l-4 border-blue-500 rounded mb-3">
                <h5 class="font-semibold text-blue-800 mb-2">✅ Good Performance</h5>
                <p class="text-sm text-blue-700 mb-2">Your OER (${avgOer.toFixed(2)}%) is good. Focus on consistency and incremental gains.</p>
                <ul class="text-sm text-blue-700 list-disc list-inside space-y-1">
                    <li>Maintain current best practices</li>
                    <li>Fine-tune process parameters for 1-2% improvement</li>
                    <li>Monitor fruit quality consistency</li>
                </ul>
            </div>
        `;
    } else {
        recommendations += `
            <div class="p-4 bg-green-50 border-l-4 border-green-500 rounded mb-3">
                <h5 class="font-semibold text-green-800 mb-2">🏆 Excellent Performance</h5>
                <p class="text-sm text-green-700 mb-2">Your OER (${avgOer.toFixed(2)}%) is excellent. Focus on maintaining this level.</p>
                <ul class="text-sm text-green-700 list-disc list-inside space-y-1">
                    <li>Document current best practices</li>
                    <li>Share knowledge across estates</li>
                    <li>Monitor for any degradation in performance</li>
                </ul>
            </div>
        `;
    }

    recBody.innerHTML = recommendations;
    const performerData = renderMillPerformanceTables(data);
    if (performerData) {
        renderKeyInsights(performerData.topStats, performerData.bottomStats);
    }
}

const AI_PERFORMANCE_TARGET_OER = 23;
const AI_PERFORMANCE_MAX_ROWS = 5;

function renderMillPerformanceTables(data) {
    const recentData = getRecentMonthsRecords(data, 12);
    const stats = buildEstateOerStats(recentData);
    const topStats = renderPerformersTable(stats, 'ai-top-performers-body', (a, b) => b.avgOer - a.avgOer);
    const bottomStats = renderPerformersTable(stats, 'ai-bottom-performers-body', (a, b) => a.avgOer - b.avgOer);
    return { topStats, bottomStats };
}

function getRecentMonthsRecords(records, months = 12) {
    const monthKeys = [...new Set(
        records
            .map(r => (r.date ? r.date.substring(0, 7) : null))
            .filter(Boolean)
    )]
        .sort((a, b) => new Date(`${a}-01`) - new Date(`${b}-01`));

    if (!monthKeys.length) return [];
    const selected = monthKeys.slice(-months);
    return records.filter(r => {
        if (!r.date) return false;
        return selected.includes(r.date.substring(0, 7));
    });
}

function buildEstateOerStats(records) {
    const map = {};
    records.forEach(rec => {
        if (!rec || !rec.estate || rec.oer_after == null || rec.oer_after === 0) return;
        const key = rec.estate;
        if (!map[key]) {
            map[key] = {
                estate: rec.estate,
                oerSum: 0,
                oerCount: 0,
                maxOer: -Infinity,
                fruitMonths: 0,
                fruitTotals: { inti: 0, plasma: 0, p3: 0 }
            };
        }
        const entry = map[key];
        entry.oerSum += rec.oer_after;
        entry.oerCount += 1;
        if (rec.oer_after > entry.maxOer) entry.maxOer = rec.oer_after;

        const hasFruit = rec.fruit_inti != null || rec.fruit_plasma != null || rec.fruit_3p != null;
        if (hasFruit) {
            entry.fruitMonths += 1;
            entry.fruitTotals.inti += rec.fruit_inti || 0;
            entry.fruitTotals.plasma += rec.fruit_plasma || 0;
            entry.fruitTotals.p3 += rec.fruit_3p || 0;
        }
    });

    return Object.values(map).map(entry => {
        if (!entry.oerCount) return null;
        const avgOer = entry.oerSum / entry.oerCount;
        const maxOer = entry.maxOer === -Infinity ? null : entry.maxOer;
        const fruitMonths = entry.fruitMonths;
        const avgInti = fruitMonths ? entry.fruitTotals.inti / fruitMonths : 0;
        const avgPlasma = fruitMonths ? entry.fruitTotals.plasma / fruitMonths : 0;
        const avg3P = fruitMonths ? entry.fruitTotals.p3 / fruitMonths : 0;

        const fruits = [
            { name: 'Inti', value: avgInti },
            { name: 'Plasma', value: avgPlasma },
            { name: '3P', value: avg3P }
        ];
        const dominant = fruits.reduce((prev, current) => (current.value > prev.value ? current : prev));
        const dominantFruit = dominant.value > 0 ? dominant.name : '-';
        const dominantPct = dominant.value > 0 ? dominant.value * 100 : null;

        const avgIntiPct = avgInti * 100;
        const avgPlasmaPct = avgPlasma * 100;
        const avg3PPct = avg3P * 100;

        return {
            estate: entry.estate,
            avgOer,
            gapToTarget: avgOer - AI_PERFORMANCE_TARGET_OER,
            gapToTopMonth: maxOer != null ? maxOer - avgOer : null,
            dominantFruit,
            dominantFruitPct: dominantPct,
            avgIntiPct,
            avgPlasmaPct,
            avg3PPct
        };
    }).filter(Boolean);
}

function renderPerformersTable(stats, elementId, comparator, limit = AI_PERFORMANCE_MAX_ROWS) {
    const container = document.getElementById(elementId);
    if (!container) return [];

    if (!stats.length) {
        container.innerHTML = '<tr><td colspan="7" class="px-3 py-6 text-center text-slate-500">Insufficient data in the selected period.</td></tr>';
        return [];
    }

    const sorted = [...stats].sort(comparator).slice(0, limit);

    container.innerHTML = sorted.map((row, index) => {
        const avgOerText = `${row.avgOer.toFixed(2)}%`;
        const targetGap = formatGap(row.gapToTarget);
        const topGap = formatGap(row.gapToTopMonth);
        const dominantPct = row.dominantFruitPct != null ? `${row.dominantFruitPct.toFixed(1)}%` : '-';

        return `
            <tr class="border-b border-slate-100 text-[11px] text-slate-600">
                <td class="px-3 py-2 font-semibold text-slate-700">${index + 1}</td>
                <td class="px-3 py-2 font-semibold text-slate-800">${row.estate}</td>
                <td class="px-3 py-2 text-right font-bold text-slate-900">${avgOerText}</td>
                <td class="px-3 py-2 text-right font-medium text-slate-700">${targetGap}</td>
                <td class="px-3 py-2">${row.dominantFruit}</td>
                <td class="px-3 py-2 text-right font-semibold text-slate-900">${dominantPct}</td>
                <td class="px-3 py-2 text-right font-medium text-slate-700">${topGap}</td>
            </tr>
        `;
    }).join('');

    return sorted;
}

function renderKeyInsights(topStats = [], bottomStats = []) {
    const container = document.getElementById('ai-key-insights-body');
    if (!container) return;

    if (!topStats.length || !bottomStats.length) {
        container.innerHTML = '<p class="text-sm text-slate-500">Key insights populate after the workbook is processed.</p>';
        return;
    }

    const topAvg = calculateAverageShares(topStats);
    const bottomAvg = calculateAverageShares(bottomStats);
    const insights = [];

    const intiDelta = topAvg.inti - bottomAvg.inti;
    if (intiDelta >= 2) {
        insights.push(`Top performers maintain roughly ${topAvg.inti.toFixed(1)}% Inti vs ${bottomAvg.inti.toFixed(1)}% across the bottom group, reinforcing that higher Inti share tracks with stronger OER.`);
    }

    const plasmaDelta = bottomAvg.plasma - topAvg.plasma;
    if (plasmaDelta >= 2) {
        insights.push(`Bottom mills rely on ${bottomAvg.plasma.toFixed(1)}% Plasma compared to ${topAvg.plasma.toFixed(1)}% for leaders, suggesting Plasma quality or proportion can suppress OER when elevated.`);
    }

    const thirdPartyDelta = bottomAvg.p3 - topAvg.p3;
    if (thirdPartyDelta >= 2) {
        insights.push(`Higher 3P exposure (${bottomAvg.p3.toFixed(1)}% vs ${topAvg.p3.toFixed(1)}%) aligns with lower OER, so clamp down on 3P quality/quantity if you need gains.`);
    }

    if (!insights.length) {
        insights.push('Top and bottom performers currently present similar fruit mix patterns; dig into processing losses or fruit quality signals for the next insight layer.');
    }

    container.innerHTML = `
        <ul class="space-y-2 list-disc list-inside text-sm text-slate-600">
            ${insights.map(item => `<li>${item}</li>`).join('')}
        </ul>
    `;
}

function calculateAverageShares(stats) {
    if (!stats.length) return { inti: 0, plasma: 0, p3: 0 };
    const aggregates = stats.reduce((acc, row) => {
        acc.inti += typeof row.avgIntiPct === 'number' ? row.avgIntiPct : 0;
        acc.plasma += typeof row.avgPlasmaPct === 'number' ? row.avgPlasmaPct : 0;
        acc.p3 += typeof row.avg3PPct === 'number' ? row.avg3PPct : 0;
        return acc;
    }, { inti: 0, plasma: 0, p3: 0 });

    const count = stats.length;
    return {
        inti: aggregates.inti / count,
        plasma: aggregates.plasma / count,
        p3: aggregates.p3 / count
    };
}

function formatGap(value) {
    if (value == null || Number.isNaN(value)) return '-';
    const formatted = value.toFixed(2);
    return `${value >= 0 ? '+' : ''}${formatted}%`;
}
