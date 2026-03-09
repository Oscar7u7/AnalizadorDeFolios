let chartPie = null;
let chartBar = null;

const state = {
    allData: [],
    filteredData: [],
    categories: [
        'Instalacion de D2Express', 'Venta', 'Salas Cerradas', 
        'Folios dentro de SLA', 'CC', 'Servicio tecnico pedidos enviado y entregado',
        'Servicio tecnico revisiones preventivas', 'Almacen', 'Sin clasificar'
    ],
    colors: ['#FFB3BA', '#FFDFBA', '#FFFFBA', '#BAFFC9', '#BAE1FF', '#D1BAFF', '#FFBAF2', '#E2F0CB', '#999999']
};

// Función para poner fecha y hora actual automáticamente
function setCurrentDateTime() {
    const now = new Date();
    // Ajuste de zona horaria local para el input datetime-local
    const offset = now.getTimezoneOffset() * 60000;
    const localISOTime = (new Date(now - offset)).toISOString().slice(0, 16);
    document.getElementById('analysisDate').value = localISOTime;
}

setCurrentDateTime();

// Mostrar nombre del archivo al elegirlo
document.getElementById('excelFile').addEventListener('change', function(e) {
    const fileName = e.target.files[0] ? e.target.files[0].name : "Elegir archivo...";
    document.getElementById('fileNameDisplay').textContent = fileName;
});

// Llenar el select de categorías
const select = document.getElementById('categoryFilter');
state.categories.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = cat;
    opt.textContent = cat;
    select.appendChild(opt);
});

document.getElementById('analyzeBtn').addEventListener('click', () => {
    const fileInput = document.getElementById('excelFile');
    if (!fileInput.files[0]) return alert("Por favor, selecciona un archivo Excel.");
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
        processData(json);
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
});

function processData(rows) {
    const refDate = new Date(document.getElementById('analysisDate').value);
    const forbiddenCols = ['INICIO REPARACION', 'FIN REPARACION', 'DIAS TECNICO', 'HORAS TECNICO'];

    state.allData = rows.map(row => {
        const cat = classifyRow(row, refDate);
        const cleanRow = {};
        Object.keys(row).forEach(c => {
            if (!forbiddenCols.includes(c.toUpperCase()) && String(row[c]).trim() !== "") {
                cleanRow[c] = row[c];
            }
        });
        return { ...cleanRow, CategoriaFinal: cat };
    });

    state.filteredData = [...state.allData];
    updateUI();
}

function classifyRow(row, refDate) {
    const norm = (v) => String(v || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    const obs = norm(row['Observacion']);
    const sub = norm(row['subfalla']);
    const gar = norm(row['Garantia']);
    const ped = norm(row['Pedidos Central']);

    if (obs.includes("d2express")) return 'Instalacion de D2Express';
    if (sub.includes("venta") || sub.includes("cotizacion") || gar === "venta") return 'Venta';
    if (sub.includes("sala cerrada")) return 'Salas Cerradas';
    if (sub.includes("preventiva")) return 'Servicio tecnico revisiones preventivas';

    if (gar.includes("renta") || gar === "" || gar.includes("datos")) {
        if (!ped) {
            const inicio = new Date(row['Inicio Folio']);
            const diff = (refDate - inicio) / (1000 * 60 * 60);
            return diff < 24 ? 'Folios dentro de SLA' : 'CC';
        }
        if (ped.includes("proceso") || ped.includes("parcial") || ped.includes("pendiente")) return 'Almacen';
        if (ped.includes("enviado") || ped.includes("entregado")) return 'Servicio tecnico pedidos enviado y entregado';
        if (ped.includes("cancelado") || ped.includes("no visto") || ped.includes("espera")) return 'CC';
    }
    return 'Sin clasificar';
}

function updateUI() {
    renderCards();
    renderCharts();
    renderTable();
    document.getElementById('resultsSection').classList.remove('hidden');
    document.getElementById('downloadBtn').disabled = false;
    document.getElementById('statusBox').textContent = "Análisis completado con éxito.";
}

function renderCards() {
    const container = document.getElementById('summaryCards');
    container.innerHTML = "";
    const counts = {};
    state.categories.forEach(c => counts[c] = 0);
    state.allData.forEach(r => counts[r.CategoriaFinal]++);

    state.categories.forEach(cat => {
        const card = document.createElement('div');
        card.className = 'card';
        card.innerHTML = `<div class="label">${cat}</div><div class="value">${counts[cat]}</div>`;
        card.onclick = () => {
            document.getElementById('categoryFilter').value = cat;
            applyFilters();
        };
        container.appendChild(card);
    });
}

function renderCharts() {
    const countsMap = {};
    state.categories.forEach(c => countsMap[c] = state.allData.filter(r => r.CategoriaFinal === c).length);
    const dataValues = state.categories.map(c => countsMap[c]);

    if (chartPie) chartPie.destroy();
    if (chartBar) chartBar.destroy();

    chartPie = new Chart(document.getElementById('pieChart'), {
        type: 'doughnut',
        data: {
            labels: state.categories,
            datasets: [{ data: dataValues, backgroundColor: state.colors, borderColor: '#111', borderWidth: 2 }]
        },
        options: {
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: 'white', font: { size: 11 } } } }
        }
    });

    chartBar = new Chart(document.getElementById('barChart'), {
        type: 'bar',
        data: {
            labels: state.categories,
            datasets: [{ label: 'Folios', data: dataValues, backgroundColor: state.colors, borderRadius: 5 }]
        },
        options: {
            maintainAspectRatio: false,
            scales: { 
                y: { ticks: { color: 'white' }, grid: { color: '#333' } },
                x: { ticks: { color: 'white', font: { size: 9 } }, grid: { display: false } }
            },
            plugins: { legend: { display: false } }
        }
    });
}

function renderTable() {
    const tbody = document.getElementById('detailBody');
    const thead = document.getElementById('tableHeaders');
    tbody.innerHTML = "";
    if (state.filteredData.length === 0) return;

    const dataCols = Object.keys(state.filteredData[0]);
    thead.innerHTML = `<th>#</th>` + dataCols.map(c => `<th>${c}</th>`).join("");

    state.filteredData.forEach((row, index) => {
        const tr = document.createElement('tr');
        let rowHtml = `<td>${index + 1}</td>`;
        rowHtml += dataCols.map(c => c === 'CategoriaFinal' ? `<td><span class="badge">${row[c]}</span></td>` : `<td>${row[c]}</td>`).join("");
        tr.innerHTML = rowHtml;
        tbody.appendChild(tr);
    });
}

function applyFilters() {
    const cat = document.getElementById('categoryFilter').value.toLowerCase();
    const search = document.getElementById('searchInput').value.toLowerCase();

    state.filteredData = state.allData.filter(r => {
        const matchesCat = cat === "" || r.CategoriaFinal.toLowerCase() === cat;
        const matchesSearch = search === "" || Object.values(r).some(v => String(v).toLowerCase().includes(search));
        return matchesCat && matchesSearch;
    });
    renderTable();
}

document.getElementById('categoryFilter').addEventListener('change', applyFilters);
document.getElementById('searchInput').addEventListener('input', applyFilters);

document.getElementById('downloadBtn').addEventListener('click', () => {
    const counts = {};
    state.categories.forEach(c => counts[c] = state.allData.filter(r => r.CategoriaFinal === c).length);
    const summary = state.categories.map(c => ({ "FOLIOS": c, "No.": counts[c] }));
    summary.push({ "FOLIOS": "TOTAL", "No.": state.allData.length });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "RESUMEN");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.allData), "DETALLE");
    XLSX.writeFile(wb, "Reporte_Zitro_Final.xlsx");
});