let chartPie = null, chartBar = null;
const state = {
    allData: [], filteredData: [],
    categories: ['Instalacion de D2Express', 'Venta', 'Salas Cerradas', 'Folios dentro de SLA', 'CC', 'Servicio tecnico pedidos enviado y entregado', 'Servicio tecnico revisiones preventivas', 'Almacen', 'Sin clasificar'],
    colors: ['#FFB3BA', '#FFDFBA', '#FFFFBA', '#BAFFC9', '#BAE1FF', '#D1BAFF', '#FFBAF2', '#E2F0CB', '#999999']
};

const setNow = () => {
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    document.getElementById('analysisDate').value = now.toISOString().slice(0, 16);
};
setNow();

document.getElementById('excelFile').addEventListener('change', e => {
    document.getElementById('fileNameDisplay').textContent = e.target.files[0]?.name || "Elegir archivo...";
});

const filterSelect = document.getElementById('categoryFilter');
state.categories.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = opt.textContent = cat;
    filterSelect.appendChild(opt);
});

document.getElementById('analyzeBtn').addEventListener('click', () => {
    const file = document.getElementById('excelFile').files[0];
    if (!file) return alert("Selecciona un Excel");
    const reader = new FileReader();
    reader.onload = e => {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        processData(json);
    };
    reader.readAsArrayBuffer(file);
});

function processData(rows) {
    const refDate = new Date(document.getElementById('analysisDate').value);
    const delCols = ['INICIO REPARACION', 'FIN REPARACION', 'DIAS TECNICO', 'HORAS TECNICO'];
    state.allData = rows.map(row => {
        const cat = classifyRow(row, refDate);
        const clean = {};
        Object.keys(row).forEach(k => {
            if (!delCols.includes(k.toUpperCase()) && String(row[k]).trim() !== "") clean[k] = row[k];
        });
        return { ...clean, CategoriaFinal: cat };
    });
    state.filteredData = [...state.allData];
    updateUI();
}

function classifyRow(row, refDate) {
    const n = v => String(v || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    const obs = n(row['Observacion']), sub = n(row['subfalla']), gar = n(row['Garantia']), ped = n(row['Pedidos Central']);
    if (obs.includes("d2express")) return 'Instalacion de D2Express';
    if (sub.includes("venta") || sub.includes("cotizacion") || gar === "venta") return 'Venta';
    if (sub.includes("sala cerrada")) return 'Salas Cerradas';
    if (sub.includes("preventiva")) return 'Servicio tecnico revisiones preventivas';
    if (gar.includes("renta") || gar === "" || gar.includes("datos")) {
        if (!ped) {
            const diff = (refDate - new Date(row['Inicio Folio'])) / 3600000;
            return diff < 24 ? 'Folios dentro de SLA' : 'CC';
        }
        if (ped.includes("proceso") || ped.includes("parcial") || ped.includes("pendiente")) return 'Almacen';
        if (ped.includes("enviado") || ped.includes("entregado")) return 'Servicio tecnico pedidos enviado y entregado';
        if (ped.includes("cancelado") || ped.includes("no visto") || ped.includes("espera")) return 'CC';
    }
    return 'Sin clasificar';
}

function updateUI() {
    renderCards(); renderCharts(); renderTable();
    document.getElementById('resultsSection').classList.remove('hidden');
    document.getElementById('downloadBtn').disabled = false;
    document.getElementById('downloadFilteredBtn').disabled = false;
}

function renderCards() {
    const container = document.getElementById('summaryCards');
    container.innerHTML = "";
    const counts = {};
    state.categories.forEach(c => counts[c] = state.allData.filter(r => r.CategoriaFinal === c).length);
    state.categories.forEach(cat => {
        const card = document.createElement('div');
        card.className = 'card';
        card.innerHTML = `<div class="label">${cat}</div><div class="value">${counts[cat]}</div>`;
        card.onclick = () => { document.getElementById('categoryFilter').value = cat; applyFilters(); };
        container.appendChild(card);
    });
    const totalCard = document.createElement('div');
    totalCard.className = 'card total-card';
    totalCard.innerHTML = `<div class="label">Total de folios al corte</div><div class="value">${state.allData.length}</div>`;
    totalCard.onclick = () => { document.getElementById('categoryFilter').value = ""; applyFilters(); };
    container.appendChild(totalCard);
}

function renderCharts() {
    const data = state.categories.map(c => state.allData.filter(r => r.CategoriaFinal === c).length);
    if (chartPie) chartPie.destroy(); if (chartBar) chartBar.destroy();
    chartPie = new Chart(document.getElementById('pieChart'), {
        type: 'doughnut',
        data: { labels: state.categories, datasets: [{ data, backgroundColor: state.colors, borderColor: '#111', borderWidth: 2 }] },
        options: { maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#fff' } } } }
    });
    chartBar = new Chart(document.getElementById('barChart'), {
        type: 'bar',
        data: { labels: state.categories, datasets: [{ data, backgroundColor: state.colors }] },
        options: { maintainAspectRatio: false, scales: { y: { ticks: { color: '#fff' } }, x: { ticks: { color: '#fff' } } }, plugins: { legend: { display: false } } }
    });
}

function renderTable() {
    const tbody = document.getElementById('detailBody'), thead = document.getElementById('tableHeaders');
    tbody.innerHTML = ""; if (!state.filteredData.length) return;
    const cols = Object.keys(state.filteredData[0]);
    thead.innerHTML = `<th>#</th>` + cols.map(c => `<th>${c}</th>`).join("");
    state.filteredData.forEach((r, i) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${i+1}</td>` + cols.map(c => c === 'CategoriaFinal' ? `<td><span class="badge">${r[c]}</span></td>` : `<td>${r[c]}</td>`).join("");
        tbody.appendChild(tr);
    });
}

function applyFilters() {
    const cat = document.getElementById('categoryFilter').value.toLowerCase();
    const search = document.getElementById('searchInput').value.toLowerCase();
    state.filteredData = state.allData.filter(r => {
        const mCat = !cat || r.CategoriaFinal.toLowerCase() === cat;
        const mSearch = !search || Object.values(r).some(v => String(v).toLowerCase().includes(search));
        return mCat && mSearch;
    });
    renderTable();
}

document.getElementById('categoryFilter').addEventListener('change', applyFilters);
document.getElementById('searchInput').addEventListener('input', applyFilters);

const modal = document.getElementById('dividirModal');
document.getElementById('downloadFilteredBtn').addEventListener('click', () => modal.classList.remove('hidden'));
document.getElementById('closeModal').addEventListener('click', () => modal.classList.add('hidden'));

document.getElementById('numPersonas').addEventListener('change', (e) => {
    const n = parseInt(e.target.value);
    const cont = document.getElementById('nombresInputs');
    cont.innerHTML = "";
    for(let i=1; i<=n; i++) cont.innerHTML += `<input type="text" placeholder="Nombre Supervisor ${i}" id="p${i}" class="p-input">`;
});

document.getElementById('confirmExport').addEventListener('click', () => {
    const n = parseInt(document.getElementById('numPersonas').value);
    const personas = [];
    for(let i=1; i<=n; i++) personas.push(document.getElementById(`p${i}`).value || `Supervisor ${i}`);
    const excluded = ['venta', 'instalacion de d2express', 'salas cerradas', 'almacen'];
    const filtered = state.allData.filter(r => !excluded.includes(r.CategoriaFinal.toLowerCase()));
    const dividedData = filtered.map((row, index) => {
        const personaIndex = index % personas.length;
        const newRow = { SUPERVISOR: personas[personaIndex], FOLIO: row['Folio'] || "", "OBSERVACIONES CC": "" };
        Object.keys(row).forEach(key => { if(key.toLowerCase() !== 'folio') newRow[key] = row[key]; });
        return newRow;
    });
    exportXls(dividedData, "Reporte_Zitro_Operativo");
    modal.classList.add('hidden');
});

document.getElementById('downloadBtn').addEventListener('click', () => exportXls(state.allData, "Reporte_Zitro_Completo"));

// FUNCIÓN DE EXCEL INTELIGENTE Y AUTOMATIZADA
function exportXls(dataArray, fileNameBase) {
    const includeMonthly = document.getElementById('checkMensual').checked;
    const wb = XLSX.utils.book_new();
    let summary = [];

    if (includeMonthly) {
        // DETECCIÓN DINÁMICA DE AÑOS
        const currentYear = new Date().getFullYear();
        const lastYear = currentYear - 1;
        const yearShort = currentYear.toString().slice(-2);
        
        const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
        const dynamicMonths = monthNames.map(m => `${m}-${yearShort}`);
        
        // ORDEN DE COLUMNAS FORZADO (FOLIOS -> No. -> Año Anterior -> Meses Año Actual)
        const headerOrder = ["FOLIOS", "No.", lastYear.toString(), ...dynamicMonths];

        summary = state.categories.map(cat => {
            const catData = dataArray.filter(r => r.CategoriaFinal === cat);
            const row = {};
            row["FOLIOS"] = cat;
            row["No."] = catData.length;
            
            // Cálculo dinámico para el año anterior
            row[lastYear.toString()] = catData.filter(r => {
                const d = new Date(r['Inicio Folio']);
                return d.getFullYear() === lastYear;
            }).length;

            // Cálculo dinámico para los meses del año actual
            dynamicMonths.forEach((m, i) => {
                row[m] = catData.filter(r => {
                    const d = new Date(r['Inicio Folio']);
                    return d.getFullYear() === currentYear && d.getMonth() === i;
                }).length;
            });
            return row;
        });

        // FILA DE TOTALES DINÁMICA
        const totalRow = {};
        totalRow["FOLIOS"] = "TOTAL";
        totalRow["No."] = dataArray.length;
        totalRow[lastYear.toString()] = dataArray.filter(r => new Date(r['Inicio Folio']).getFullYear() === lastYear).length;
        
        dynamicMonths.forEach((m, i) => {
            totalRow[m] = dataArray.filter(r => {
                const d = new Date(r['Inicio Folio']);
                return d.getFullYear() === currentYear && d.getMonth() === i;
            }).length;
        });
        summary.push(totalRow);

        const wsSummary = XLSX.utils.json_to_sheet(summary, { header: headerOrder });
        XLSX.utils.book_append_sheet(wb, wsSummary, "RESUMEN");

    } else {
        summary = state.categories.map(c => ({ "FOLIOS": c, "No.": dataArray.filter(r => r.CategoriaFinal === c).length }));
        summary.push({ "FOLIOS": "TOTAL", "No.": dataArray.length });
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "RESUMEN");
    }

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dataArray), "DETALLE");
    XLSX.writeFile(wb, `${fileNameBase}_${new Date().toLocaleDateString().replace(/\//g,'-')}.xlsx`);
}
