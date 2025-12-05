// Note: app.js was corrupted during edits. 
// The file needs to be restored from backup or from a working version.
// The renderOverviewTable function should:
// 1. Group data by estate
// 2. Calculate average OER for each estate
// 3. Calculate dominant fruit source (Inti/Plasma/3P)
// 4. Sort by average OER descending (highest first)
// 5. Display columns: PSM, Region, Estate Code, LMM, Avg OER (After), Dominant Fruit Source

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
