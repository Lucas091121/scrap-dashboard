// ==================== GLOBAL VARIABLES ====================
let allRecords = [];
const dellBlue = '#0076CE';
const dellDarkBlue = '#00447C';
const dellMediumBlue = '#0096D6';
const dellLightBlue = '#62B5E5';
const dellSkyBlue = '#A7D4E8';
const dellNavy = '#004B87';
const bluesPalette = [dellDarkBlue, dellBlue, dellMediumBlue, dellLightBlue, dellSkyBlue, dellNavy, '#0085CA', '#B8D4E3', '#3A8FBF', '#1A6DAA'];

let chartMostRecovered, chartHighestValue, chartQtyPlatform, chartValuePlatform, chartTrend;

// ==================== MONTH MAPPING ====================
const monthMap = {
    'Jan':'January', 'Feb':'February', 'Mar':'March', 'Apr':'April',
    'May':'May', 'Jun':'June', 'Jul':'July', 'Aug':'August',
    'Sep':'September', 'Oct':'October', 'Nov':'November', 'Dec':'December',
    'January':'January', 'February':'February', 'March':'March', 'April':'April',
    'June':'June', 'July':'July', 'August':'August',
    'September':'September', 'October':'October', 'November':'November', 'December':'December'
};

const quarterMap = {
    'Q1':'Qtr 1', 'Q2':'Qtr 2', 'Q3':'Qtr 3', 'Q4':'Qtr 4',
    'Qtr 1':'Qtr 1', 'Qtr 2':'Qtr 2', 'Qtr 3':'Qtr 3', 'Qtr 4':'Qtr 4'
};

// ==================== LOAD EXCEL FILE ====================
async function loadExcelData() {
    try {
        const response = await fetch('data/scrap-data.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        // Read first sheet
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        console.log('Raw data from Excel:', jsonData);
        console.log('First row:', jsonData[0]);
        console.log('Column names:', Object.keys(jsonData[0]));

        // Convert to our format
        allRecords = jsonData.map(row => {
            // Find the right column names (flexible matching)
            const keys = Object.keys(row);

            // Get Quarter
            const quarterKey = keys.find(k => k.toLowerCase().includes('quarter') || k === 'Quarter') || keys[0];
            let quarter = String(row[quarterKey] || '');
            quarter = quarterMap[quarter] || quarter;

            // Get Date and parse month/day
            const dateKey = keys.find(k => k.toLowerCase().includes('data') || k.toLowerCase().includes('date')) || keys[1];
            let dateVal = row[dateKey];
            let month = '';
            let day = 0;

            if (dateVal) {
                const dateStr = String(dateVal);
                // Try format "13-Mar" or "Mar-13"
                const parts = dateStr.split('-');
                if (parts.length === 2) {
                    const num = parseInt(parts[0]);
                    if (isNaN(num)) {
                        // Format: "Mar-13"
                        month = monthMap[parts[0]] || parts[0];
                        day = parseInt(parts[1]) || 0;
                    } else {
                        // Format: "13-Mar"
                        day = num;
                        month = monthMap[parts[1]] || parts[1];
                    }
                } else if (typeof dateVal === 'number') {
                    // Excel serial date
                    const excelDate = XLSX.SSF.parse_date_code(dateVal);
                    if (excelDate) {
                        const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
                        month = monthNames[excelDate.m - 1] || '';
                        day = excelDate.d || 0;
                    }
                }
            }

            // Get PPID
            const ppidKey = keys.find(k => k.toLowerCase().includes('ppid')) || keys[2];
            const ppid = String(row[ppidKey] || '-');

            // Get Item
            const itemKey = keys.find(k => k.toLowerCase() === 'item') || keys[3];
            const item = String(row[itemKey] || '');

            // Get Item Name
            const itemNameKey = keys.find(k => k.toLowerCase().includes('item name') || k.toLowerCase().includes('itemname') || k.toLowerCase().includes('item_name')) || keys[4];
            const itemName = String(row[itemNameKey] || '');

            // Get QTY
            const qtyKey = keys.find(k => k.toLowerCase().includes('qty') || k.toLowerCase().includes('quantity')) || keys[5];
            const qty = parseInt(row[qtyKey]) || 0;

            // Get Platform
            const platformKey = keys.find(k => k.toLowerCase().includes('plataforma') || k.toLowerCase().includes('platform')) || keys[6];
            const platform = String(row[platformKey] || '');

            // Get Cost
            const costKey = keys.find(k => k.toLowerCase().includes('cost') || k.toLowerCase().includes('scrap cost')) || keys[keys.length - 1];
            let cost = row[costKey];
            if (typeof cost === 'string') {
                cost = parseFloat(cost.replace('$', '').replace(',', '')) || 0;
            } else {
                cost = parseFloat(cost) || 0;
            }

            return { quarter, month, day, platform, ppid, item, itemName, qty, cost };
        }).filter(r => r.item && r.item !== '' && r.item !== 'undefined');

        console.log('Converted records:', allRecords.length);
        console.log('Sample record:', allRecords[0]);

        // Initialize dashboard
        populateFilters();
        updateDashboard();

        // Update status
        document.querySelector('.meta-date').textContent = 'Data loaded: ' + allRecords.length + ' records • ' + new Date().toLocaleString();

    } catch (error) {
        console.error('Error loading Excel:', error);
        document.querySelector('.meta-date').textContent = 'Error loading data. Check console (F12).';
    }
}

// ==================== HELPERS ====================
function groupBy(arr, key, valueKey, aggregate) {
    const map = {};
    arr.forEach(r => {
        const k = r[key];
        if (!map[k]) map[k] = 0;
        map[k] += r[valueKey];
    });
    const sorted = Object.entries(map).sort((a, b) => b[1] - a[1]);
    if (aggregate === 'top4') return sorted.slice(0, 4);
    return sorted;
}

function getMonthOrder(m) {
    const order = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    return order.indexOf(m);
}

function groupByMonth(arr) {
    const map = {};
    arr.forEach(r => {
        if (r.month && r.month !== '') {
            if (!map[r.month]) map[r.month] = 0;
            map[r.month] += r.cost;
        }
    });
    return Object.entries(map).sort((a, b) => getMonthOrder(a[0]) - getMonthOrder(b[0]));
}

// ==================== FILTER DATA ====================
function getFilteredData() {
    const q = document.getElementById('quarterFilter').value;
    const p = document.getElementById('platformFilter').value;
    return allRecords.filter(r => {
        if (q !== 'all' && r.quarter !== q) return false;
        if (p !== 'all' && r.platform !== p) return false;
        return true;
    });
}

// ==================== UPDATE KPIs ====================
function updateKPIs(data) {
    const totalCost = data.reduce((s, r) => s + r.cost, 0);
    const totalQty = data.reduce((s, r) => s + r.qty, 0);
    const uniqueItems = new Set(data.map(r => r.item)).size;
    const avgValue = totalQty > 0 ? totalCost / totalQty : 0;

    document.getElementById('kpiRecoveryValue').textContent = '$' + totalCost.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
    document.getElementById('kpiRecoveredParts').textContent = totalQty.toLocaleString('en-US');
    document.getElementById('kpiUniqueMaterials').textContent = uniqueItems;
    document.getElementById('kpiAvgValue').textContent = '$' + avgValue.toFixed(2);
}

// ==================== UPDATE TABLE ====================
function updateTable(data) {
    const tbody = document.getElementById('tableBody');
    let html = '';
    data.forEach(r => {
        html += '<tr>';
        html += '<td>2026</td>';
        html += '<td>' + r.quarter + '</td>';
        html += '<td>' + r.month + '</td>';
        html += '<td>' + r.day + '</td>';
        html += '<td>' + r.platform + '</td>';
        html += '<td>' + r.ppid + '</td>';
        html += '<td>' + r.item + '</td>';
        html += '<td>' + r.itemName + '</td>';
        html += '<td>' + r.qty + '</td>';
        html += '<td>$' + r.cost.toFixed(2) + '</td>';
        html += '</tr>';
    });
    tbody.innerHTML = html;

    const totalQty = data.reduce((s, r) => s + r.qty, 0);
    const totalCost = data.reduce((s, r) => s + r.cost, 0);
    const tfoot = document.querySelector('.data-table tfoot');
    tfoot.innerHTML = '<tr><td colspan="8"><strong>Total</strong></td><td><strong>' + totalQty.toLocaleString('en-US') + '</strong></td><td><strong>$' + totalCost.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}) + '</strong></td></tr>';
}

// ==================== CHART OPTIONS ====================
const barOptions = (dollarSign) => ({
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: true,
    plugins: { legend: { display: false } },
    scales: {
        x: {
            grid: { color: '#f0f0f0' },
            ticks: { font: { size: 11 }, callback: function(v) { return dollarSign ? '$'+v : v; } }
        },
        y: { grid: { display: false }, ticks: { font: { size: 11, weight: '600' } } }
    }
});

// ==================== BUILD CHARTS ====================
function buildCharts(data) {
    if (chartMostRecovered) chartMostRecovered.destroy();
    if (chartHighestValue) chartHighestValue.destroy();
    if (chartQtyPlatform) chartQtyPlatform.destroy();
    if (chartValuePlatform) chartValuePlatform.destroy();
    if (chartTrend) chartTrend.destroy();

    const topItems = groupBy(data, 'item', 'qty', 'top4');
    chartMostRecovered = new Chart(document.getElementById('chartMostRecovered').getContext('2d'), {
        type: 'bar',
        data: { labels: topItems.map(i => i[0]), datasets: [{ data: topItems.map(i => i[1]), backgroundColor: dellBlue, borderRadius: 4 }] },
        options: barOptions(false)
    });

    const topValue = groupBy(data, 'item', 'cost', 'top4');
    chartHighestValue = new Chart(document.getElementById('chartHighestValue').getContext('2d'), {
        type: 'bar',
        data: { labels: topValue.map(i => i[0]), datasets: [{ data: topValue.map(i => parseFloat(i[1].toFixed(2))), backgroundColor: dellBlue, borderRadius: 4 }] },
        options: barOptions(true)
    });

    const topPlatQty = groupBy(data, 'platform', 'qty', 'top4');
    chartQtyPlatform = new Chart(document.getElementById('chartQtyPlatform').getContext('2d'), {
        type: 'bar',
        data: { labels: topPlatQty.map(i => i[0]), datasets: [{ data: topPlatQty.map(i => i[1]), backgroundColor: dellBlue, borderRadius: 4 }] },
        options: barOptions(false)
    });

    const platValue = groupBy(data, 'platform', 'cost');
    chartValuePlatform = new Chart(document.getElementById('chartValuePlatform').getContext('2d'), {
        type: 'doughnut',
        data: {
            labels: platValue.map(i => i[0]),
            datasets: [{ data: platValue.map(i => parseFloat(i[1].toFixed(2))), backgroundColor: bluesPalette.slice(0, platValue.length), borderWidth: 2, borderColor: '#fff' }]
        },
        options: {
            responsive: true, maintainAspectRatio: true,
            plugins: {
                legend: { position: 'bottom', labels: { font: { size: 10 }, padding: 8, usePointStyle: true, pointStyle: 'circle' } },
                tooltip: { callbacks: { label: function(ctx) { return ctx.label + ': $' + ctx.parsed.toFixed(2); } } }
            }
        }
    });

    const monthly = groupByMonth(data);
    chartTrend = new Chart(document.getElementById('chartTrend').getContext('2d'), {
        type: 'line',
        data: {
            labels: monthly.map(i => i[0]),
            datasets: [{ data: monthly.map(i => parseFloat(i[1].toFixed(2))), borderColor: dellBlue, backgroundColor: 'rgba(0,118,206,0.1)', fill: true, tension: 0.3, pointBackgroundColor: dellBlue, pointBorderColor: '#fff', pointBorderWidth: 2, pointRadius: 6, pointHoverRadius: 8 }]
        },
        options: {
            responsive: true, maintainAspectRatio: true,
            plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(ctx) { return '$' + ctx.parsed.y; } } } },
            scales: {
                x: { grid: { display: false }, ticks: { font: { size: 11 } } },
                y: { grid: { color: '#f0f0f0' }, ticks: { font: { size: 11 }, callback: function(v) { return '$'+v; } } }
            }
        }
    });
}

// ==================== POPULATE FILTERS ====================
function populateFilters() {
    const quarters = [...new Set(allRecords.map(r => r.quarter))].filter(q => q).sort();
    const platforms = [...new Set(allRecords.map(r => r.platform))].filter(p => p).sort();

    const qSelect = document.getElementById('quarterFilter');
    qSelect.innerHTML = '<option value="all">All</option>';
    quarters.forEach(q => { qSelect.innerHTML += '<option value="' + q + '">' + q + '</option>'; });

    const pSelect = document.getElementById('platformFilter');
    pSelect.innerHTML = '<option value="all">All</option>';
    platforms.forEach(p => { pSelect.innerHTML += '<option value="' + p + '">' + p + '</option>'; });
}

// ==================== UPDATE ALL ====================
function updateDashboard() {
    const data = getFilteredData();
    updateKPIs(data);
    buildCharts(data);
    updateTable(data);
}

// ==================== EVENT LISTENERS ====================
document.getElementById('quarterFilter').addEventListener('change', updateDashboard);
document.getElementById('platformFilter').addEventListener('change', updateDashboard);

// ==================== INIT ====================
loadExcelData();
