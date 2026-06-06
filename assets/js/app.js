/**
 * CLAYTICS - Lógica de Procesamiento y Coordinación de Agentes
 */

// Elementos del DOM
const welcomeSection = document.getElementById('welcome-section');
const moldingSection = document.getElementById('molding-section');
const dashboardSection = document.getElementById('dashboard-section');
const dropZone = document.getElementById('drop-zone');
const startBtn = document.getElementById('start-btn');
const kpiContainer = document.getElementById('kpi-container');
const insightsContainer = document.getElementById('insights-container');
const fileInput = document.getElementById('file-input');

let currentFileData = null;
let currentFileName = "";
let currentKPIs = [];
let currentWorkbook = null;      // Workbook actual de Excel (para cambiar de hoja)
let currentSheetName = null;     // Nombre de la hoja activa

// Variables del Historial y Consumo
let fileHistory = JSON.parse(localStorage.getItem('claytics-file-history') || '[]');
let analysisCount = parseInt(localStorage.getItem('claytics-analysis-count') || '0');

// Formatear bytes de forma amigable
function formatBytes(bytes, decimals = 2) {
    if (bytes === 0 || !bytes) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// 1. Gestión de Selección de Archivo
if (dropZone && fileInput) {
    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleFileSelect(file);
    });
}

function handleFileSelect(file) {
    currentFileName = file.name;
    dropZone.innerHTML = `
        <i class="fas fa-check-circle" style="color: var(--p-success)"></i>
        <h3>${currentFileName}</h3>
    `;
    startBtn.classList.remove('hidden');
    currentFileData = file;
}

// 2. Flujo Principal (Procesamiento y Coordinación de Agentes)
startBtn.addEventListener('click', async () => {
    if (!currentFileData) return;

    const extension = currentFileData.name.split('.').pop().toLowerCase();

    if (extension === 'csv' || extension === 'txt') {
        welcomeSection.classList.add('hidden');
        moldingSection.classList.remove('hidden');
        try {
            const rawData = await parseTextFile(currentFileData);
            await processData(rawData);
        } catch (error) {
            console.error(error);
            alert("Error al moldear la data. Revisa el formato.");
            location.reload();
        }
    } else {
        // Es archivo de Excel (.xlsx, .xls)
        const reader = new FileReader();
        reader.onload = async (e) => {
            const content = e.target.result;
            try {
                const wb = XLSX.read(content, { type: 'binary' });
                const sheetNames = wb.SheetNames;
                currentWorkbook = wb; // Guardar workbook globalmente

                if (sheetNames.length > 1) {
                    // Cargar múltiples hojas: abrir el modal interactivo de selección
                    showSheetSelector(wb, sheetNames);
                } else {
                    // Una sola hoja → pasar por el detector de múltiples tablas
                    welcomeSection.classList.add('hidden');
                    moldingSection.classList.remove('hidden');
                    currentSheetName = sheetNames[0];
                    // Nota: handleExcelSheetWithDetection se define al final del archivo.
                    // Usamos setTimeout para garantizar que la función esté disponible.
                    setTimeout(async () => {
                        try {
                            await handleExcelSheetWithDetection(wb, sheetNames[0]);
                        } catch (err) {
                            console.error(err);
                            alert('Error al procesar el archivo Excel.');
                            location.reload();
                        }
                    }, 0);
                }
            } catch (error) {
                console.error(error);
                alert("Error al parsear el archivo Excel. Revisa el formato.");
                location.reload();
            }
        };
        reader.readAsBinaryString(currentFileData);
    }
});

// ═══════════════════════════════════════════════════════
//  BOTÓN CAMBIAR ARCHIVO
// ═══════════════════════════════════════════════════════
const changeFileBtn = document.getElementById('change-file-btn');
const changeFileInput = document.getElementById('change-file-input');

if (changeFileBtn && changeFileInput) {
    changeFileBtn.addEventListener('click', () => changeFileInput.click());
    
    changeFileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        currentFileData = file;
        currentFileName = file.name;
        currentWorkbook = null;
        currentSheetName = null;

        // Ocultar la barra de hoja antes de procesar
        const switchBar = document.getElementById('switch-sheet-bar');
        if (switchBar) switchBar.classList.add('hidden');

        const extension = file.name.split('.').pop().toLowerCase();

        // Ocultar dashboard y mostrar loader
        dashboardSection.classList.add('hidden');
        moldingSection.classList.remove('hidden');

        if (extension === 'csv' || extension === 'txt') {
            try {
                const rawData = await parseTextFile(file);
                await processData(rawData);
            } catch (error) {
                console.error(error);
                alert("Error al cargar el nuevo archivo. Revisa el formato.");
                location.reload();
            }
        } else {
            const reader = new FileReader();
            reader.onload = async (ev) => {
                const content = ev.target.result;
                try {
                    const wb = XLSX.read(content, { type: 'binary' });
                    currentWorkbook = wb;
                    const sheetNames = wb.SheetNames;

                    if (sheetNames.length > 1) {
                        moldingSection.classList.add('hidden');
                        showSheetSelector(wb, sheetNames);
                    } else {
                        currentSheetName = sheetNames[0];
                        // Pasar por el detector de múltiples tablas
                        setTimeout(async () => {
                            try {
                                await handleExcelSheetWithDetection(wb, sheetNames[0]);
                            } catch (err) {
                                console.error(err);
                                alert('Error al leer el nuevo archivo Excel.');
                                location.reload();
                            }
                        }, 0);
                    }
                } catch (error) {
                    console.error(error);
                    alert("Error al leer el nuevo archivo Excel.");
                    location.reload();
                }
            };
            reader.readAsBinaryString(file);
        }
        // Resetear el input para que pueda volver a seleccionar el mismo archivo si quiere
        changeFileInput.value = '';
    });
}

// Función de procesamiento unificado de datos (Agentes + Renderizado)
async function processData(rawData) {
    // Validar consumo límite (10 análisis al mes) — solo si no tiene Plan Pro
    const userPlan = localStorage.getItem('claytics-plan') || 'free';
    analysisCount = parseInt(localStorage.getItem('claytics-analysis-count') || '0');
    if (userPlan !== 'pro' && analysisCount >= 10) {
        alert("¡Límite excedido! Has consumido tus 10 análisis mensuales.\n\nActualiza a Plan Pro en la pestaña Plan y Facturación para continuar.");
        if (moldingSection) moldingSection.classList.add('hidden');
        if (welcomeSection) welcomeSection.classList.remove('hidden');
        return;
    }


    try {
        // Revisar si existe una columna ID, de lo contrario agregarla
        const hasId = Object.keys(rawData[0] || {}).some(k => k.toLowerCase() === 'id');
        let data = rawData;
        if (!hasId) {
            data = rawData.map((row, index) => {
                return { id: index + 1, ...row };
            });
        }
        
        cachedRawData = data; // Guardar en caché para re-renders
        
        // Resetear filtros para el nuevo archivo
        filtersInitialized = false;
        activeHiddenColumns = [];
        activeRowRule = { column: '', condition: 'none', value: '' };

        // Retraso para simular análisis y moldear con animación
        await new Promise(resolve => setTimeout(resolve, 2000));

        // DELEGAMOS AL AGENTE DE KPIS
        currentKPIs = kpiAgent.analyze(data);
        renderDashboard(data);
        
        // Sumar al contador de consumo y guardar
        analysisCount++;
        localStorage.setItem('claytics-analysis-count', analysisCount.toString());
        if (typeof updateUsageProgress === 'function') updateUsageProgress();
        
        // Guardar archivo en el historial
        const fileSize = currentFileData ? formatBytes(currentFileData.size) : '—';
        const fileExt = currentFileName ? currentFileName.split('.').pop().toUpperCase() : 'CSV';
        if (typeof addFileToHistory === 'function') {
            addFileToHistory({
                name: currentFileName || "Archivo sin nombre",
                size: fileSize,
                type: fileExt,
                rows: data.length,
                cols: Object.keys(rawData[0] || {}).length,
                date: new Date().toLocaleDateString() + ' ' + new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})
            });
        }

        // Asegurar que la barra lateral muestre la pestaña de "Nuevo Proyecto" como activa
        const navBtns = document.querySelectorAll('.nav-btn[data-target]');
        navBtns.forEach(btn => btn.classList.remove('active'));
        const welcomeBtn = document.querySelector('.nav-btn[data-target="welcome-section"]');
        if (welcomeBtn) welcomeBtn.classList.add('active');

        // Cerrar secciones del sidebar y mostrar dashboard
        const mainSections = document.querySelectorAll('.main-content > section');
        mainSections.forEach(sec => sec.classList.add('hidden'));
        
        moldingSection.classList.add('hidden');
        dashboardSection.classList.remove('hidden');
    } catch (error) {
        console.error(error);
        alert("Error al procesar la data. Revisa la integridad del dataset.");
        location.reload();
    }
}

// Función para parsear archivos de texto plano (CSV o TXT)
function parseTextFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const content = e.target.result;
            Papa.parse(content, {
                header: true,
                dynamicTyping: true,
                skipEmptyLines: true,
                complete: (res) => resolve(res.data),
                error: (err) => reject(err)
            });
        };
        reader.readAsText(file);
    });
}

// Despliegue interactivo del Selector de Hojas de Excel
function showSheetSelector(wb, sheetNames) {
    const sheetModal = document.getElementById('sheet-modal');
    const sheetList = document.getElementById('sheet-buttons-list');
    
    if (!sheetModal || !sheetList) return;

    sheetList.innerHTML = '';

    sheetNames.forEach(sheetName => {
        const btn = document.createElement('button');
        btn.className = 'btn-sheet-option';
        // Resaltar la hoja activa si ya hay una seleccionada
        const isActive = sheetName === currentSheetName;
        btn.innerHTML = `
            <span><i class="fas fa-file-invoice" style="margin-right: 10px; color: hsl(142, 76%, 40%)"></i>${sheetName}${isActive ? ' <span style="font-size:0.7rem; background: #DCFCE7; color: #166534; padding: 2px 8px; border-radius: 20px; margin-left: 6px;">Activa</span>' : ''}</span>
            <i class="fas fa-arrow-right"></i>
        `;
        btn.addEventListener('click', async () => {
            sheetModal.classList.add('hidden');
            if (!dashboardSection.classList.contains('hidden')) {
                // Ya estamos en el dashboard: solo mostrar loader sin ocultar welcome
                dashboardSection.classList.add('hidden');
            } else {
                welcomeSection.classList.add('hidden');
            }
            moldingSection.classList.remove('hidden');
            
            try {
                currentSheetName = sheetName;
                const rawData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName]);
                await processData(rawData);
            } catch (error) {
                console.error(error);
                alert("Error al leer la hoja seleccionada.");
                location.reload();
            }
        });
        sheetList.appendChild(btn);
    });

    // Cerrar modal al dar click en la X
    const closeBtn = document.getElementById('sheet-close-btn');
    if (closeBtn) {
        closeBtn.onclick = () => sheetModal.classList.add('hidden');
    }

    // Cerrar modal al dar click fuera del panel
    sheetModal.onclick = (e) => {
        if (e.target === sheetModal) sheetModal.classList.add('hidden');
    };

    sheetModal.classList.remove('hidden');
}

// Actualiza la barra indicadora de hoja activa
function updateSheetBar() {
    const bar = document.getElementById('switch-sheet-bar');
    const label = document.getElementById('switch-sheet-label');
    const switchBtn = document.getElementById('switch-sheet-btn');

    if (!bar) return;

    if (currentWorkbook && currentWorkbook.SheetNames.length > 1) {
        if (label) label.textContent = `Hoja activa: ${currentSheetName || currentWorkbook.SheetNames[0]}`;
        bar.classList.remove('hidden');

        if (switchBtn) {
            switchBtn.onclick = () => showSheetSelector(currentWorkbook, currentWorkbook.SheetNames);
        }
    } else {
        bar.classList.add('hidden');
    }
}

function renderDashboard(data) {
    kpiContainer.innerHTML = '';
    insightsContainer.innerHTML = '';

    // Renderizar KPIs generados por el Agente
    currentKPIs.forEach(kpi => {
        addKPI(kpi);
    });

    // Insight único: nombres de columnas del dataset
    const columns = Object.keys(data[0]);
    renderColumnsInsight(columns, insightsContainer);

    // Renderizar tabla de datos
    renderDataTable(data);

    // Actualizar barra de hoja activa
    updateSheetBar();
}

let cachedRawData = null;

function addKPI(kpi) {
    const card = document.createElement('div');
    card.className = 'kpi-card';
    
    // Color dinámico para la salud de los datos
    let iconStyle = "";
    if (kpi.type === 'health') {
        iconStyle = `color: ${kpi.score > 90 ? 'var(--p-success)' : (kpi.score > 70 ? 'orange' : 'red')}`;
    }

    card.innerHTML = `
        <div class="label"><i class="${kpi.icon}" style="${iconStyle}"></i> ${kpi.label}</div>
        <div class="value">${kpi.value}</div>
    `;
    kpiContainer.appendChild(card);
}

let activeHiddenColumns = [];
let activeRowRule = { column: '', condition: 'none', value: '' };
let filtersInitialized = false;

/**
 * Muestra las columnas detectadas como una card de insight, separándolas por tipo y nombre.
 */
function renderColumnsInsight(columns, container) {
    const card = document.createElement('div');
    card.className = 'insight-card';
    
    const tags = columns.map(col => {
        const colType = kpiAgent.columnTypes[col] || 'Categorical';
        let typeLabel = "Texto";
        let typeClass = "type-text";
        if (colType === 'Numeric') { typeLabel = "Número"; typeClass = "type-number"; }
        if (colType === 'ID') { typeLabel = "ID"; typeClass = "type-id"; }
        if (colType === 'Date') { typeLabel = "Fecha"; typeClass = "type-date"; }
        if (colType === 'Boolean') { typeLabel = "Booleano"; typeClass = "type-boolean"; }
        
        return `
            <div class="col-type-tag">
                <span class="col-name">${col}</span>
                <span class="col-type-badge ${typeClass}">${typeLabel}</span>
            </div>
        `;
    }).join('');

    card.innerHTML = `
        <h4><i class="fas fa-table-columns" style="color:var(--p-primary)"></i> Columnas y Tipos Detectados</h4>
        <div class="col-type-tags-wrap">${tags}</div>
    `;
    container.appendChild(card);
}

/**
 * Renderiza la tabla de vista previa de datos con filtros activos de filas y columnas.
 */
let analysisListenersAttached = false;
let exportListenersAttached = false;

function checkRowIssue(row, issueId, arr) {
    if (issueId === 'blank_rows') {
        return Object.values(row).every(v => v === null || v === '' || v === undefined);
    } else if (issueId === 'blank_cols') {
        return false; // Blank cols affect the schema, not specific rows. Return false so it doesn't flood the row filter.
    } else if (issueId === 'duplicates') {
        const key = JSON.stringify(row);
        return arr.filter(r => JSON.stringify(r) === key).length > 1;
    } else if (issueId === 'whitespace') {
        return Object.values(row).some(v => {
            const s = String(v || '');
            return s !== s.trim() && s.trim() !== '';
        });
    } else if (issueId === 'case_normalize') {
        const nameCols = Object.keys(row).filter(c => c.toLowerCase().includes('nombre') || c.toLowerCase().includes('cliente'));
        return nameCols.some(c => {
            const v = String(row[c] || '');
            return v.length > 1 && (v === v.toUpperCase() || v === v.toLowerCase());
        });
    } else if (issueId.startsWith('num_dirty_')) {
        const c = issueId.replace('num_dirty_', '');
        const val = String(row[c] || '').trim();
        return val !== '' && isNaN(parseFloat(val)) && val !== 'null';
    } else if (issueId.startsWith('email_format_')) {
        const c = issueId.replace('email_format_', '');
        const val = String(row[c] || '').trim();
        return val !== '' && !val.includes('@');
    } else if (issueId.startsWith('type_name_')) {
        const c = issueId.replace('type_name_', '');
        const val = String(row[c] || '').trim();
        return /^\d+$/.test(val);
    } else if (issueId === 'partial_nulls') {
        // Fila que tiene al menos una celda nula en cualquier columna
        return Object.values(row).some(v => v === null || v === undefined || v === '');
    } else if (issueId === 'semi_empty_rows') {
        // Fila con más del 50% de celdas vacías
        const vals = Object.values(row);
        const nullCount = vals.filter(v => v === null || v === undefined || v === '').length;
        return vals.length > 0 && (nullCount / vals.length) >= 0.5;
    }
    return false; // If issue ID doesn't match, return false so it doesn't show randomly.
}

/**
 * Retorna la data filtrada y columnas visibles de acuerdo al estado actual de los filtros
 */
function getFilteredData() {
    if (!cachedRawData || cachedRawData.length === 0) return { columns: [], rows: [] };
    
    const allColumns = Object.keys(cachedRawData[0] || {});
    const visibleColumns = allColumns.filter(c => !activeHiddenColumns.includes(c));
    
    let filteredRows = cachedRawData;
    if (activeRowRule.condition !== 'none' && (activeRowRule.column || activeRowRule.condition === 'issue' || activeRowRule.condition === 'issue_group')) {
        const col = activeRowRule.column;
        const cond = activeRowRule.condition;
        const val = activeRowRule.value; // For issues, value is issueId
        
        // Pre-compute duplicate keys for O(N) performance instead of O(N^2)
        let dupKeys = new Set();
        let seenKeys = new Set();
        if (cond === 'issue' || cond === 'issue_group') {
            cachedRawData.forEach(r => {
                const k = JSON.stringify(r);
                if (seenKeys.has(k)) dupKeys.add(k);
                else seenKeys.add(k);
            });
        }

        filteredRows = cachedRawData.filter((row, index, arr) => {
            if (cond === 'issue_group') {
                return val.some(issueId => {
                    if (issueId === 'duplicates') return dupKeys.has(JSON.stringify(row));
                    return checkRowIssue(row, issueId, arr);
                });
            } else if (cond === 'issue') {
                if (val === 'duplicates') return dupKeys.has(JSON.stringify(row));
                return checkRowIssue(row, val, arr);
            }

            // Normal filters
            const searchVal = String(val).toLowerCase().trim();
            const cellVal = String(row[col] === null || row[col] === undefined ? '' : row[col]);
            const cellNum = parseFloat(row[col]);

            if (cond === 'contains') {
                return cellVal.toLowerCase().includes(searchVal);
            } else if (cond === 'equals') {
                return cellVal.toLowerCase() === searchVal;
            } else if (cond === 'gt') {
                return !isNaN(cellNum) && cellNum > parseFloat(searchVal);
            } else if (cond === 'lt') {
                return !isNaN(cellNum) && cellNum < parseFloat(searchVal);
            }
            return true;
        });
    }
    
    return { columns: visibleColumns, rows: filteredRows };
}

function renderDataTable(data) {
    const section = document.getElementById('data-preview-section');
    const table   = document.getElementById('preview-table');
    const meta    = document.getElementById('preview-meta');
    if (!section || !table) return;

    const allColumns = Object.keys(data[0] || {});
    
    // Inicializar panel de filtros avanzados si no se ha hecho
    if (!filtersInitialized) {
        initFilterPanel(data);
        filtersInitialized = true;
    }

    // Obtener data de acuerdo al filtro actual
    const { columns: visibleColumns, rows: filteredRows } = getFilteredData();
    const maxRows = Math.min(filteredRows.length, 50); // Mostrar al menos 50 filas

    meta.textContent = `${filteredRows.length} filas × ${visibleColumns.length} columnas visibles (de ${data.length} × ${allColumns.length} totales)`;

    if (visibleColumns.length === 0) {
        table.innerHTML = `<tbody><tr><td style="text-align:center; padding: 40px; color: var(--p-text-light);">No hay columnas visibles. Selecciona columnas en el panel superior.</td></tr></tbody>`;
        return;
    }

    // Cabecera
    const thead = `<thead><tr>${visibleColumns.map(c => `<th>${c}</th>`).join('')}</tr></thead>`;

    // Filas
    const rows = filteredRows.slice(0, maxRows).map(row => {
        const cells = visibleColumns.map(c => {
            const val = row[c];
            const display = (val === null || val === undefined || val === '') ? '<span class="cell-empty">—</span>' : String(val);
            return `<td>${display}</td>`;
        }).join('');
        return `<tr>${cells}</tr>`;
    }).join('');

    table.innerHTML = `${thead}<tbody>${rows}</tbody>`;
    section.classList.remove('hidden');

    // Popular el selector de columnas para el laboratorio (sólo columnas numéricas y visibles)
    populateColumnSelect(data);

    // Inicializar listeners del laboratorio si no se han agregado antes
    if (!analysisListenersAttached) {
        initAnalysisLab();
        analysisListenersAttached = true;
    }

    // Inicializar listeners de descarga
    if (!exportListenersAttached) {
        initExportListeners();
        exportListenersAttached = true;
    }
}

function populateColumnSelect(data, typeFilter = 'all') {
    const select = document.getElementById('analysis-col-select');
    if (!select) return;

    select.innerHTML = '';
    const columns = Object.keys(data[0] || {});

    columns.forEach(col => {
        const colType = kpiAgent.columnTypes[col] || 'Categorical';
        
        // Aplicar el filtro de tipo
        if (typeFilter !== 'all' && colType !== typeFilter) {
            // Permitir que Categorical encaje si buscamos Categorical, etc.
            return;
        }

        let typeLabel = "Texto";
        if (colType === 'Numeric') typeLabel = "Número";
        if (colType === 'ID') typeLabel = "ID";
        if (colType === 'Date') typeLabel = "Fecha";
        if (colType === 'Boolean') typeLabel = "Booleano";

        const option = document.createElement('option');
        option.value = col;
        option.textContent = `${col} (${typeLabel})`;
        select.appendChild(option);
    });

    // Despachar evento change para actualizar la interfaz inmediatamente
    select.dispatchEvent(new Event('change'));
}

function initAnalysisLab() {
    // Input de búsqueda
    const searchInput = document.getElementById('table-search');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase();
            const table = document.getElementById('preview-table');
            const rows = table.querySelectorAll('tbody tr');
            rows.forEach(row => {
                const text = row.textContent.toLowerCase();
                if (text.includes(query)) {
                    row.classList.remove('hidden');
                } else {
                    row.classList.add('hidden');
                }
            });
        });
    }

    // Filtro de tipo de columna
    const typeSelect = document.getElementById('analysis-type-select');
    if (typeSelect) {
        typeSelect.addEventListener('change', (e) => {
            populateColumnSelect(cachedRawData, e.target.value);
        });
    }

    // Al cambiar la columna, mostrar los botones correspondientes al tipo
    const colSelect = document.getElementById('analysis-col-select');
    const gridNumber = document.getElementById('analysis-grid-number');
    const gridText = document.getElementById('analysis-grid-text');
    const resultCard = document.getElementById('lab-result-card');

    if (colSelect) {
        colSelect.addEventListener('change', (e) => {
            resultCard.classList.add('hidden'); // ocultar resultados previos
            const col = e.target.value;
            const colType = kpiAgent.columnTypes[col] || 'Categorical';
            
            if (colType === 'Numeric') {
                if(gridNumber) gridNumber.classList.remove('hidden');
                if(gridText) gridText.classList.add('hidden');
            } else {
                if(gridNumber) gridNumber.classList.add('hidden');
                if(gridText) gridText.classList.remove('hidden');
            }
        });
    }

    // Botones del laboratorio
    const labButtons = document.querySelectorAll('.btn-lab');
    const resultTitle = document.getElementById('lab-result-title');
    const resultValue = document.getElementById('lab-result-value');
    const resultDesc = document.getElementById('lab-result-desc');

    labButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const action = btn.getAttribute('data-action');
            const col = colSelect.value;
            if (!col || !cachedRawData) return;

            calculateMetric(action, col, resultCard, resultTitle, resultValue, resultDesc);
        });
    });

    // Limpiar análisis
    const clearBtn = document.getElementById('btn-clear-analysis');
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            resultCard.classList.add('hidden');
        });
    }
}

function calculateMetric(action, col, card, titleEl, valueEl, descEl) {
    const colType = kpiAgent.columnTypes[col] || 'Categorical';
    const rawValues = cachedRawData.map(r => r[col]).filter(v => v !== null && v !== undefined && v !== '');

    // Para funciones matemáticas, verificar si la columna es apta
    const mathActions = ['sum', 'avg', 'max', 'min', 'median', 'stddev', 'variance', 'range'];
    
    if (mathActions.includes(action) && colType !== 'Numeric') {
        card.classList.remove('hidden');
        card.style.borderColor = 'var(--p-primary)';
        titleEl.textContent = `Error de Análisis`;
        valueEl.textContent = `No Apto`;
        valueEl.style.color = '#EF4444';
        descEl.innerHTML = `⚠️ La columna <strong>"${col}"</strong> está clasificada como <strong>${colType === 'ID' ? 'Identificador (ID)' : 'Texto'}</strong>. Realizar operaciones matemáticas sobre ella carece de sentido. Selecciona una columna de tipo <strong>Número</strong>.`;
        return;
    }

    // Convertir a números si aplica
    const numericValues = rawValues.map(v => parseFloat(v)).filter(v => !isNaN(v));

    if (mathActions.includes(action) && numericValues.length === 0) {
        card.classList.remove('hidden');
        titleEl.textContent = `Sin datos numéricos`;
        valueEl.textContent = `—`;
        descEl.textContent = `No hay datos numéricos válidos en la columna "${col}" para realizar este cálculo.`;
        return;
    }

    let result = 0;
    let label = '';
    let description = '';

    const formatNum = (val) => {
        if (col.toLowerCase().includes('precio') || col.toLowerCase().includes('price') || col.toLowerCase().includes('venta') || col.toLowerCase().includes('total') || col.toLowerCase().includes('monto')) {
            return `S/ ${val.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
        }
        return val % 1 === 0 ? val.toLocaleString() : val.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };

    switch (action) {
        case 'sum':
            result = numericValues.reduce((sum, v) => sum + v, 0);
            label = `Suma Acumulada: ${col}`;
            description = `Suma total de todos los registros numéricos de la columna. Es útil para comprender el volumen total generado.`;
            valueEl.textContent = formatNum(result);
            break;
            
        case 'avg':
            result = numericValues.reduce((sum, v) => sum + v, 0) / numericValues.length;
            label = `Promedio Aritmético: ${col}`;
            description = `La media aritmética de los datos. Representa el valor central típico de este conjunto de registros.`;
            valueEl.textContent = formatNum(result);
            break;

        case 'count':
            // Conteo funciona para cualquier tipo
            result = rawValues.length;
            label = `Conteo de Registros: ${col}`;
            description = `La cantidad total de celdas no vacías en esta columna. Útil para verificar la completitud de la data.`;
            valueEl.textContent = `${result.toLocaleString()} filas`;
            break;

        case 'max':
            result = Math.max(...numericValues);
            label = `Valor Máximo: ${col}`;
            description = `El registro con el valor más alto en todo el conjunto de datos de esta columna.`;
            valueEl.textContent = formatNum(result);
            break;

        case 'min':
            result = Math.min(...numericValues);
            label = `Valor Mínimo: ${col}`;
            description = `El registro con el valor más bajo o menor en todo el conjunto de datos de esta columna.`;
            valueEl.textContent = formatNum(result);
            break;

        case 'median':
            const sorted = [...numericValues].sort((a, b) => a - b);
            const mid = Math.floor(sorted.length / 2);
            result = sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
            label = `Mediana Estadística: ${col}`;
            description = `El valor central de los datos cuando se ordenan de menor a mayor. A diferencia del promedio, no se ve afectado por valores extremos (anomalías).`;
            valueEl.textContent = formatNum(result);
            break;

        case 'mode':
            // Moda funciona para texto o números!
            const freqs = {};
            let maxFreq = 0;
            let modes = [];
            rawValues.forEach(v => {
                freqs[v] = (freqs[v] || 0) + 1;
                if (freqs[v] > maxFreq) maxFreq = freqs[v];
            });
            for (const k in freqs) {
                if (freqs[k] === maxFreq) modes.push(k);
            }
            label = `Moda (Más Frecuente): ${col}`;
            if (maxFreq > 1) {
                valueEl.textContent = modes.slice(0, 3).join(', ');
                description = `El valor que más veces se repite (aparece ${maxFreq} veces). Representa la categoría o valor dominante.`;
            } else {
                valueEl.textContent = `Valores Únicos`;
                description = `Todos los registros en esta columna aparecen exactamente una sola vez, por lo que no existe una moda definida.`;
            }
            break;

        case 'range':
            result = Math.max(...numericValues) - Math.min(...numericValues);
            label = `Rango de Dispersión: ${col}`;
            description = `La diferencia matemática entre el valor máximo y el mínimo. Te indica la amplitud o el espacio total que cubren los datos.`;
            valueEl.textContent = formatNum(result);
            break;

        case 'stddev':
            const avgVal = numericValues.reduce((sum, v) => sum + v, 0) / numericValues.length;
            const sqDiffs = numericValues.map(v => Math.pow(v - avgVal, 2));
            const varianceVal = sqDiffs.reduce((sum, v) => sum + v, 0) / numericValues.length;
            result = Math.sqrt(varianceVal);
            label = `Desviación Estándar: ${col}`;
            description = `Mide la dispersión promedio de los datos con respecto a su promedio. Una desviación estándar baja indica que los datos están agrupados cerca de la media; una alta indica alta dispersión.`;
            valueEl.textContent = formatNum(result);
            break;

        case 'variance':
            const meanVal = numericValues.reduce((sum, v) => sum + v, 0) / numericValues.length;
            const diffs = numericValues.map(v => Math.pow(v - meanVal, 2));
            result = diffs.reduce((sum, v) => sum + v, 0) / numericValues.length;
            label = `Varianza: ${col}`;
            description = `La varianza mide la variabilidad de los datos al cuadrado. Es la base para calcular la desviación estándar en estadística descriptiva.`;
            valueEl.textContent = result % 1 === 0 ? result.toLocaleString() : result.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            break;

        case 'uppercase':
            cachedRawData.forEach(row => { if (row[col] && typeof row[col] === 'string') row[col] = row[col].toUpperCase(); });
            label = `Mayúsculas Aplicadas: ${col}`;
            description = `Todos los textos de la columna han sido convertidos a letras mayúsculas.`;
            valueEl.textContent = `¡Listo!`;
            renderDataTable(cachedRawData);
            break;

        case 'lowercase':
            cachedRawData.forEach(row => { if (row[col] && typeof row[col] === 'string') row[col] = row[col].toLowerCase(); });
            label = `Minúsculas Aplicadas: ${col}`;
            description = `Todos los textos de la columna han sido convertidos a letras minúsculas.`;
            valueEl.textContent = `¡Listo!`;
            renderDataTable(cachedRawData);
            break;

        case 'titlecase':
            cachedRawData.forEach(row => { 
                if (row[col] && typeof row[col] === 'string') {
                    row[col] = row[col].toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
                }
            });
            label = `Capitalización Aplicada: ${col}`;
            description = `Se ha puesto la primera letra en mayúscula para cada palabra de la columna.`;
            valueEl.textContent = `¡Listo!`;
            renderDataTable(cachedRawData);
            break;

        case 'trim':
            let trimmedCount = 0;
            cachedRawData.forEach(row => { 
                if (row[col] && typeof row[col] === 'string') {
                    const original = row[col];
                    const cleaned = original.trim().replace(/\s+/g, ' ');
                    if (original !== cleaned) trimmedCount++;
                    row[col] = cleaned;
                }
            });
            label = `Espacios Eliminados: ${col}`;
            description = `Se eliminaron espacios en blanco sobrantes al inicio, final o entre palabras.`;
            valueEl.textContent = `${trimmedCount} filas`;
            renderDataTable(cachedRawData);
            break;

        case 'frequency':
            const textFreqs = {};
            rawValues.forEach(v => {
                const s = String(v).toLowerCase().trim();
                textFreqs[s] = (textFreqs[s] || 0) + 1;
            });
            const sortedFreqs = Object.entries(textFreqs).sort((a, b) => b[1] - a[1]);
            label = `Frecuencia de Textos: ${col}`;
            description = `El texto que más se repite es "${sortedFreqs[0][0]}" (${sortedFreqs[0][1]} veces).`;
            valueEl.textContent = sortedFreqs.slice(0, 3).map(e => `${e[0]}: ${e[1]}`).join(' | ');
            break;
    }

    card.classList.remove('hidden');
    valueEl.style.color = 'var(--p-dark)';
    titleEl.textContent = label;
    descEl.textContent = description;
}

/**
 * Muestra el aviso de que los datos fueron limpiados.
 */
function showCleanedNotice() {
    const notice = document.getElementById('data-clean-notice');
    if (notice) {
        notice.classList.remove('hidden');
        // Scroll suave hacia la tabla
        notice.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

// Prevenir comportamiento de arrastre
['dragover', 'drop'].forEach(eventName => window.addEventListener(eventName, (e) => e.preventDefault(), false));

// ═══════════════════════════════════════════════════════
//  LÓGICA DEL AGENTE DE LIMPIEZA (CleanAgent)
// ═══════════════════════════════════════════════════════

const cleanBtn       = document.getElementById('clean-btn');
const cleanModal     = document.getElementById('clean-modal');
const cleanCloseBtn  = document.getElementById('clean-close-btn');
const btnSelectAll   = document.getElementById('btn-select-all');
const btnApplyClean  = document.getElementById('btn-apply-clean');
const btnCloseDone   = document.getElementById('btn-close-done');

const stepScanning = document.getElementById('clean-step-scanning');
const stepResults  = document.getElementById('clean-step-results');
const stepClean    = document.getElementById('clean-step-clean');
const stepDone     = document.getElementById('clean-step-done');

// Abrir el modal y lanzar el escaneo
cleanBtn.addEventListener('click', () => {
    if (!cachedRawData) return;
    openCleanModal();
});

function openCleanModal() {
    cleanModal.classList.remove('hidden');
    // Reset a estado inicial
    [stepResults, stepClean, stepDone].forEach(s => s.classList.add('hidden'));
    stepScanning.classList.remove('hidden');

    // Simular tiempo de análisis para UX (da sensación de trabajo real)
    setTimeout(() => {
        const issues = cleanAgent.scan(cachedRawData, kpiAgent.columnTypes);
        renderCleanIssues(issues);
    }, 1400);
}

function renderCleanIssues(issues) {
    stepScanning.classList.add('hidden');

    if (issues.length === 0) {
        stepClean.classList.remove('hidden');
        return;
    }

    // Mostrar resultados
    const subtitleEl = document.getElementById('clean-subtitle-text');
    const listEl     = document.getElementById('clean-issues-list');

    subtitleEl.innerHTML = `Encontré <strong>${issues.length} problema(s)</strong> en tu dataset. Selecciona cuáles corregir:`;
    listEl.innerHTML = '';

    const severityConfig = {
        high:   { color: '#EF4444', label: 'Alta',   bg: '#FEF2F2' },
        medium: { color: '#F59E0B', label: 'Media',  bg: '#FFFBEB' },
        low:    { color: '#6366F1', label: 'Baja',   bg: '#EEF2FF' }
    };

    issues.forEach(issue => {
        const sev = severityConfig[issue.severity] || severityConfig.low;
        const card = document.createElement('div');
        card.className = 'clean-issue-card';
        card.dataset.issueId = issue.id;
        card.innerHTML = `
            <div class="clean-issue-left-cb">
                <input type="checkbox" class="clean-checkbox" id="cb-${issue.id}" value="${issue.id}" checked>
            </div>
            <div class="clean-issue-content">
                <label for="cb-${issue.id}" class="clean-issue-label-text">
                    <div class="clean-issue-top">
                        <span class="clean-issue-icon" style="color:${sev.color}"><i class="${issue.icon}"></i></span>
                        <span class="clean-issue-title">${issue.title}</span>
                        <span class="clean-severity-badge" style="background:${sev.bg}; color:${sev.color}">
                            ${sev.label}
                        </span>
                    </div>
                    <p class="clean-issue-desc">${issue.description}</p>
                    <div class="clean-preview-tag">
                        <i class="fas fa-eye"></i> ${issue.previewLabel}
                    </div>
                    <div class="clean-fix-tag">
                        <i class="fas fa-wand-magic-sparkles"></i> ${issue.fixLabel}
                    </div>
                </label>
                <div class="clean-issue-actions" style="margin-top: 10px; display: flex; gap: 10px;">
                    <button type="button" class="btn-locate-issue" data-issue-id="${issue.id}">
                        <i class="fas fa-search-location"></i> Ubicar problema
                    </button>
                    <button type="button" class="btn-filter-issue" data-issue-id="${issue.id}" style="background: none; border: 1px solid var(--p-primary); color: var(--p-primary); padding: 5px 10px; border-radius: 6px; font-size: 0.85rem; cursor: pointer; transition: all 0.3s ease;">
                        <i class="fas fa-filter"></i> Filtrar problema
                    </button>
                </div>
            </div>
        `;
        listEl.appendChild(card);
    });

    // Agregar event listener a cada botón de ubicar
    listEl.querySelectorAll('.btn-locate-issue').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const issueId = btn.getAttribute('data-issue-id');
            locateIssueInTable(issueId);
        });
    });

    // Agregar event listener a cada botón de filtrar
    listEl.querySelectorAll('.btn-filter-issue').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const issueId = btn.getAttribute('data-issue-id');
            filterIssueInTable(issueId);
        });
    });

    stepResults.classList.remove('hidden');
}

// Seleccionar / Deseleccionar todo
let allSelected = true;
btnSelectAll.addEventListener('click', () => {
    allSelected = !allSelected;
    document.querySelectorAll('.clean-checkbox').forEach(cb => cb.checked = allSelected);
    btnSelectAll.innerHTML = allSelected
        ? '<i class="fas fa-check-double"></i> Seleccionar todo'
        : '<i class="fas fa-square"></i> Deseleccionar todo';
});

// Filtrar para previsualizar los errores agrupados
const btnPreviewErrors = document.getElementById('btn-preview-errors');
if (btnPreviewErrors) {
    btnPreviewErrors.addEventListener('click', () => {
        const selectedIds = [...document.querySelectorAll('.clean-checkbox:checked')].map(cb => cb.value);
        if (selectedIds.length === 0) {
            alert('Selecciona al menos una corrección para previsualizar.');
            return;
        }

        // Usamos filterIssueGroupInTable para filtrar todos los errores seleccionados
        filterIssueGroupInTable(selectedIds);
    });
}

// Aplicar las correcciones seleccionadas
btnApplyClean.addEventListener('click', () => {
    const selectedIds = [...document.querySelectorAll('.clean-checkbox:checked')].map(cb => cb.value);
    if (selectedIds.length === 0) {
        alert('Selecciona al menos una corrección para aplicar.');
        return;
    }

    const originalCount = cachedRawData.length;
    cachedRawData = cleanAgent.applyFixes(cachedRawData, selectedIds);
    const summary = cleanAgent.generateCleanSummary(originalCount, cachedRawData);

    // Actualizar el dashboard en segundo plano
    currentKPIs = kpiAgent.analyze(cachedRawData);
    renderDashboard(cachedRawData);

    // Mostrar resumen de éxito
    stepResults.classList.add('hidden');
    const summaryEl = document.getElementById('clean-summary-text');
    summaryEl.textContent = `Se aplicaron ${selectedIds.length} corrección(es). ${summary.removedRows > 0 ? `Se eliminaron ${summary.removedRows} fila(s).` : ''} Tu dataset ahora tiene ${summary.cleanedRows} registros limpios.`;
    stepDone.classList.remove('hidden');

    // Mostrar aviso de limpieza en la tabla de datos
    showCleanedNotice();
});


// Cerrar modal (botón X y botón final)
cleanCloseBtn.addEventListener('click', () => cleanModal.classList.add('hidden'));
btnCloseDone.addEventListener('click', () => cleanModal.classList.add('hidden'));

// Cerrar con click fuera del panel
cleanModal.addEventListener('click', (e) => {
    if (e.target === cleanModal) cleanModal.classList.add('hidden');
});

// ═══════════════════════════════════════════════════════
//  INICIALIZADOR DEL PANEL DE FILTROS AVANZADOS
// ═══════════════════════════════════════════════════════
function initFilterPanel(data) {
    const colsContainer = document.getElementById('filter-cols-container');
    const rowColSelect = document.getElementById('filter-row-col');
    
    if (!colsContainer || !rowColSelect) return;

    const columns = Object.keys(data[0] || {});

    // Limpiar y poblar checkboxes de columnas
    colsContainer.innerHTML = '';
    columns.forEach(col => {
        const div = document.createElement('div');
        div.className = 'filter-col-checkbox';
        div.innerHTML = `
            <label>
                <input type="checkbox" class="col-filter-cb" value="${col}" checked>
                <span>${col}</span>
            </label>
        `;
        colsContainer.appendChild(div);
    });

    // Limpiar y poblar selector de fila
    rowColSelect.innerHTML = '<option value="">-- Seleccionar Columna --</option>';
    columns.forEach(col => {
        const opt = document.createElement('option');
        opt.value = col;
        opt.textContent = col;
        rowColSelect.appendChild(opt);
    });

    // Panel colapsable - Filtros
    const toggleBtn = document.getElementById('filter-toggle-btn');
    const body = document.getElementById('filter-panel-body');
    const icon = document.getElementById('filter-toggle-icon');
    if (toggleBtn && body && icon) {
        toggleBtn.onclick = () => {
            body.classList.toggle('hidden');
            toggleBtn.classList.toggle('active');
            if (body.classList.contains('hidden')) {
                icon.className = 'fas fa-chevron-down toggle-icon';
            } else {
                icon.className = 'fas fa-chevron-up toggle-icon';
            }
        };
    }

    // Panel colapsable - Laboratorio
    const labToggleBtn = document.getElementById('lab-toggle-btn');
    const labBody = document.getElementById('lab-panel-body');
    const labIcon = document.getElementById('lab-toggle-icon');
    if (labToggleBtn && labBody && labIcon) {
        // Aseguramos que solo se registre una vez
        labToggleBtn.onclick = () => {
            labBody.classList.toggle('hidden');
            labToggleBtn.classList.toggle('active');
            if (labBody.classList.contains('hidden')) {
                labIcon.className = 'fas fa-chevron-down toggle-icon';
            } else {
                labIcon.className = 'fas fa-chevron-up toggle-icon';
            }
        };
    }

    // Botón aplicar
    const btnApply = document.getElementById('btn-apply-filters');
    if (btnApply) {
        btnApply.onclick = () => {
            // 1. Columnas a ocultar
            const cbs = colsContainer.querySelectorAll('.col-filter-cb');
            activeHiddenColumns = [];
            cbs.forEach(cb => {
                if (!cb.checked) activeHiddenColumns.push(cb.value);
            });

            // 2. Regla de filas
            const rowCol = rowColSelect.value;
            const cond = document.getElementById('filter-row-cond').value;
            const val = document.getElementById('filter-row-val').value;

            activeRowRule = { column: rowCol, condition: cond, value: val };

            // Renderizar tabla con filtros aplicados
            renderDataTable(cachedRawData);
        };
    }

    // Botón limpiar
    const btnReset = document.getElementById('btn-reset-filters');
    if (btnReset) {
        btnReset.onclick = () => {
            // Re-marcar todas las columnas
            colsContainer.querySelectorAll('.col-filter-cb').forEach(cb => cb.checked = true);
            rowColSelect.value = '';
            document.getElementById('filter-row-cond').value = 'none';
            document.getElementById('filter-row-val').value = '';

            activeHiddenColumns = [];
            activeRowRule = { column: '', condition: 'none', value: '' };

            renderDataTable(cachedRawData);
        };
    }
}

// ═══════════════════════════════════════════════════════
//  UBICAR PROBLEMA (CleanAgent Interactive Navigation)
// ═══════════════════════════════════════════════════════
const restoreBanner = document.getElementById('restore-clean-banner');
const btnRestoreClean = document.getElementById('btn-restore-clean');

if (btnRestoreClean) {
    btnRestoreClean.addEventListener('click', () => {
        restoreBanner.classList.add('hidden');
        cleanModal.classList.remove('hidden');
    });
}

function showRestoreCleanBanner() {
    if (restoreBanner) {
        restoreBanner.classList.remove('hidden');
    }
}

function locateIssueInTable(issueId) {
    // 1. Cerrar temporalmente el modal
    cleanModal.classList.add('hidden');
    
    // 2. Mostrar banner flotante para regresar al modal
    showRestoreCleanBanner();

    const table = document.getElementById('preview-table');
    if (!table) return;

    // Limpiar highlights previos
    table.querySelectorAll('.highlight-issue-cell, .highlight-issue-row, .pulse-highlight').forEach(el => {
        el.classList.remove('highlight-issue-cell', 'highlight-issue-row', 'pulse-highlight');
    });

    const headers = [...table.querySelectorAll('thead th')].map(th => th.textContent.trim());
    const rows = table.querySelectorAll('tbody tr');
    let targetElements = [];

    if (issueId === 'blank_rows') {
        rows.forEach(row => {
            const cells = [...row.querySelectorAll('td')];
            const allEmpty = cells.every(td => td.textContent.trim() === '—' || td.textContent.trim() === '');
            if (allEmpty) {
                row.classList.add('highlight-issue-row');
                targetElements.push(row);
            }
        });
    } else if (issueId === 'blank_cols') {
        const blankCols = Object.keys(cachedRawData[0] || {}).filter(col =>
            cachedRawData.every(row => row[col] === null || row[col] === '' || row[col] === undefined)
        );
        blankCols.forEach(colName => {
            const colIdx = headers.indexOf(colName);
            if (colIdx !== -1) {
                const th = table.querySelector(`thead th:nth-child(${colIdx + 1})`);
                if (th) {
                    th.classList.add('highlight-issue-cell');
                    targetElements.push(th);
                }
                rows.forEach(row => {
                    const td = row.querySelector(`td:nth-child(${colIdx + 1})`);
                    if (td) td.classList.add('highlight-issue-cell');
                });
            }
        });
    } else if (issueId === 'duplicates') {
        const seen = new Set();
        rows.forEach((row) => {
            const cells = [...row.querySelectorAll('td')].map(td => td.textContent.trim());
            const key = cells.join('|');
            if (seen.has(key)) {
                row.classList.add('highlight-issue-row');
                targetElements.push(row);
            } else {
                seen.add(key);
            }
        });
    } else if (issueId === 'whitespace') {
        rows.forEach(row => {
            row.querySelectorAll('td').forEach(td => {
                const val = td.textContent;
                if (val && val !== '—' && (val.startsWith(' ') || val.endsWith(' '))) {
                    td.classList.add('highlight-issue-cell');
                    targetElements.push(td);
                }
            });
        });
    } else if (issueId === 'case_normalize') {
        const nameCols = Object.keys(cachedRawData[0] || {}).filter(col => col.toLowerCase().includes('nombre') || col.toLowerCase().includes('cliente'));
        nameCols.forEach(colName => {
            const colIdx = headers.indexOf(colName);
            if (colIdx !== -1) {
                rows.forEach(row => {
                    const td = row.querySelector(`td:nth-child(${colIdx + 1})`);
                    if (td) {
                        const v = td.textContent.trim();
                        if (v && v !== '—' && (v === v.toUpperCase() || v === v.toLowerCase())) {
                            td.classList.add('highlight-issue-cell');
                            targetElements.push(td);
                        }
                    }
                });
            }
        });
    } else {
        let colName = '';
        let matchType = '';
        if (issueId.startsWith('type_name_')) {
            colName = issueId.replace('type_name_', '');
            matchType = 'name';
        } else if (issueId.startsWith('email_format_')) {
            colName = issueId.replace('email_format_', '');
            matchType = 'email';
        } else if (issueId.startsWith('num_dirty_')) {
            colName = issueId.replace('num_dirty_', '');
            matchType = 'numeric';
        }

        if (colName) {
            const colIdx = headers.indexOf(colName);
            if (colIdx !== -1) {
                rows.forEach(row => {
                    const td = row.querySelector(`td:nth-child(${colIdx + 1})`);
                    if (td) {
                        const val = td.textContent.trim();
                        if (val && val !== '—') {
                            if (matchType === 'name' && /^\d+$/.test(val)) {
                                td.classList.add('highlight-issue-cell');
                                targetElements.push(td);
                            } else if (matchType === 'email' && !val.includes('@')) {
                                td.classList.add('highlight-issue-cell');
                                targetElements.push(td);
                            } else if (matchType === 'numeric' && isNaN(parseFloat(val)) && (val.includes('S/') || val.includes('$') || val.includes(','))) {
                                td.classList.add('highlight-issue-cell');
                                targetElements.push(td);
                            }
                        }
                    }
                });
            }
        }
    }

    // Scroll suave y pulso
    if (targetElements.length > 0) {
        targetElements[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
        targetElements.forEach(el => {
            el.classList.add('pulse-highlight');
            setTimeout(() => el.classList.remove('pulse-highlight'), 3000);
        });
    } else {
        table.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

function filterIssueInTable(issueId) {
    // 1. Cerrar temporalmente el modal
    cleanModal.classList.add('hidden');
    
    // 2. Mostrar banner flotante para regresar al modal
    showRestoreCleanBanner();

    // 3. Activar regla de fila especial para aislar el error
    activeRowRule = { column: '', condition: 'issue', value: issueId };
    
    // 4. Renderizar tabla con el filtro de problema aplicado
    renderDataTable(cachedRawData);

    const table = document.getElementById('preview-table');
    if (table) {
        table.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

function filterIssueGroupInTable(issueIds) {
    // 1. Cerrar temporalmente el modal
    cleanModal.classList.add('hidden');
    
    // 2. Mostrar banner flotante para regresar al modal
    showRestoreCleanBanner();

    // 3. Activar regla de fila especial para aislar el GRUPO de errores
    activeRowRule = { column: '', condition: 'issue_group', value: issueIds };
    
    // 4. Renderizar tabla con el filtro de problema aplicado
    renderDataTable(cachedRawData);

    const table = document.getElementById('preview-table');
    if (table) {
        table.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}


// ═══════════════════════════════════════════════════════
//  LÓGICA DE EXPORTACIÓN Y DESCARGAS (Excel, CSV, PDF, TXT)
// ═══════════════════════════════════════════════════════

function exportToExcel() {
    const { columns, rows } = getFilteredData();
    if (rows.length === 0 || columns.length === 0) {
        alert("No hay datos para exportar.");
        return;
    }
    
    // Mapear filas conservando solo columnas visibles en orden
    const exportRows = rows.map(row => {
        const obj = {};
        columns.forEach(col => {
            obj[col] = row[col];
        });
        return obj;
    });
    
    const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: columns });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data Filtrada");
    
    const cleanName = currentFileName.replace(/\.[^/.]+$/, "");
    XLSX.writeFile(workbook, `${cleanName}_filtrado.xlsx`);
}

function exportToCSV() {
    const { columns, rows } = getFilteredData();
    if (rows.length === 0 || columns.length === 0) {
        alert("No hay datos para exportar.");
        return;
    }
    
    const exportRows = rows.map(row => {
        const obj = {};
        columns.forEach(col => {
            obj[col] = row[col];
        });
        return obj;
    });
    
    const csvContent = Papa.unparse(exportRows, { columns: columns });
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    
    const cleanName = currentFileName.replace(/\.[^/.]+$/, "");
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${cleanName}_filtrado.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function exportToTXT() {
    const { columns, rows } = getFilteredData();
    if (rows.length === 0 || columns.length === 0) {
        alert("No hay datos para exportar.");
        return;
    }
    
    const exportRows = rows.map(row => {
        const obj = {};
        columns.forEach(col => {
            obj[col] = row[col];
        });
        return obj;
    });
    
    const txtContent = Papa.unparse(exportRows, { delimiter: "\t", columns: columns });
    const blob = new Blob([txtContent], { type: "text/plain;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    
    const cleanName = currentFileName.replace(/\.[^/.]+$/, "");
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `${cleanName}_filtrado.txt`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function exportToPDF() {
    const { columns, rows } = getFilteredData();
    if (rows.length === 0 || columns.length === 0) {
        alert("No hay datos para exportar.");
        return;
    }
    
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('l', 'pt', 'a4'); // Paisaje horizontal, puntos, formato A4
    
    const headers = [columns];
    const body = rows.map(row => columns.map(col => row[col] === null || row[col] === undefined ? "" : String(row[col])));
    
    // Encabezado premium del reporte
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.setTextColor(15, 23, 42); // --p-dark
    doc.text("Claytics | Reporte de Datos", 40, 40);
    
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Archivo de origen: ${currentFileName}`, 40, 56);
    doc.text(`Fecha del reporte: ${new Date().toLocaleString()}`, 40, 70);
    doc.text(`Registros exportados: ${rows.length}`, 40, 84);
    
    // Estilos de la tabla de AutoTable
    doc.autoTable({
        head: headers,
        body: body,
        startY: 100,
        margin: { left: 40, right: 40 },
        theme: 'striped',
        headStyles: {
            fillColor: [79, 70, 229], // --p-primary
            textColor: [255, 255, 255],
            fontStyle: 'bold'
        },
        styles: {
            font: "helvetica",
            fontSize: 8,
            cellPadding: 6
        },
        didDrawPage: function (data) {
            // Numeración de páginas en el pie
            const pageCount = doc.internal.getNumberOfPages();
            doc.setFontSize(8);
            doc.setTextColor(150);
            doc.text(`Página ${data.pageNumber} de ${pageCount}`, doc.internal.pageSize.width - 80, doc.internal.pageSize.height - 20);
        }
    });
    
    const cleanName = currentFileName.replace(/\.[^/.]+$/, "");
    doc.save(`${cleanName}_filtrado.pdf`);
}

function initExportListeners() {
    const excelBtn = document.getElementById("export-excel-btn");
    const csvBtn = document.getElementById("export-csv-btn");
    const pdfBtn = document.getElementById("export-pdf-btn");
    const txtBtn = document.getElementById("export-txt-btn");
    
    if (excelBtn) excelBtn.addEventListener("click", exportToExcel);
    if (csvBtn) csvBtn.addEventListener("click", exportToCSV);
    if (pdfBtn) pdfBtn.addEventListener("click", exportToPDF);
    if (txtBtn) txtBtn.addEventListener("click", exportToTXT);
}

// ═══════════════════════════════════════════════════════
// AUTH LOGIC (REGISTRO, LOGIN, VERIFICACIÓN)
// ═══════════════════════════════════════════════════════

const loginSection    = document.getElementById('login-section');
const verifSection    = document.getElementById('verification-section');

// Helper: activar sesión y entrar a la app
function activateSession(name, age, email, sector, needsVerify) {
    localStorage.setItem('claytics-session-name', name);
    localStorage.setItem('claytics-session-age', age);
    localStorage.setItem('claytics-session-email', email);
    localStorage.setItem('claytics-session-sector', sector);
    localStorage.setItem('claytics-needs-verification', needsVerify ? 'true' : 'false');

    const sidebarName   = document.getElementById('sidebar-user-name');
    const sidebarSector = document.getElementById('sidebar-user-sector');
    if (sidebarName)   sidebarName.textContent   = name;
    if (sidebarSector) sidebarSector.textContent = sector;

    if (typeof populateUserProfile === 'function') populateUserProfile(name, age, email, sector);

    // Añadir clase para evitar scroll en body
    document.body.classList.add('user-logged-in');

    if (loginSection)   loginSection.classList.add('hidden');
    if (verifSection)   verifSection.classList.add('hidden');

    const loggedInLayout = document.getElementById('logged-in-layout');
    if (loggedInLayout) loggedInLayout.classList.remove('hidden');
    if (welcomeSection) welcomeSection.classList.remove('hidden');

    if (needsVerify) {
        if (typeof showVerificationBanner === 'function') showVerificationBanner(email);
    } else {
        if (typeof hideVerificationBanner === 'function') hideVerificationBanner();
    }
}

// ─── Toggle entre REGISTRO y LOGIN ───────────────────────
const linkToLogin    = document.getElementById('link-to-login');
const linkToRegister = document.getElementById('link-to-register');
const registerFormContainer = document.getElementById('register-form-container');
const loginFormContainer    = document.getElementById('login-form-container');

if (linkToLogin) {
    linkToLogin.addEventListener('click', (e) => {
        e.preventDefault();
        registerFormContainer.classList.add('hidden');
        loginFormContainer.classList.remove('hidden');
    });
}
if (linkToRegister) {
    linkToRegister.addEventListener('click', (e) => {
        e.preventDefault();
        loginFormContainer.classList.add('hidden');
        registerFormContainer.classList.remove('hidden');
    });
}

// ─── FORMULARIO DE REGISTRO ───────────────────────────────
const registerForm = document.getElementById('register-form');
const registerSectorEl = document.getElementById('register-sector');
const registerCustomSectorGroup = document.getElementById('register-custom-sector-group');
const registerCustomSector = document.getElementById('register-custom-sector');

if (registerSectorEl) {
    registerSectorEl.addEventListener('change', () => {
        if (registerSectorEl.value === 'Otro') {
            registerCustomSectorGroup.classList.remove('hidden');
            registerCustomSector.required = true;
        } else {
            registerCustomSectorGroup.classList.add('hidden');
            registerCustomSector.required = false;
            registerCustomSector.value = '';
        }
    });
}

if (registerForm) {
    registerForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const name     = document.getElementById('register-name').value.trim();
        const age      = document.getElementById('register-age').value.trim();
        const email    = document.getElementById('register-email').value.trim();
        let   sector   = registerSectorEl ? registerSectorEl.value : 'General';
        const password = document.getElementById('register-password').value;
        const confirm  = document.getElementById('register-confirm-password').value;

        if (sector === 'Otro') {
            sector = (registerCustomSector && registerCustomSector.value.trim()) || 'Otro';
        }

        // Validar clave: exactamente 6 dígitos numéricos
        if (!/^\d{6}$/.test(password)) {
            alert('La clave debe ser exactamente 6 dígitos numéricos (ej: 123456).');
            return;
        }
        if (password !== confirm) {
            alert('Las claves no coinciden. Por favor vuelve a ingresarlas.');
            return;
        }

        // Guardar cuenta
        localStorage.setItem('claytics-session-name', name);
        localStorage.setItem('claytics-session-age', age);
        localStorage.setItem('claytics-session-email', email);
        localStorage.setItem('claytics-session-sector', sector);
        localStorage.setItem('claytics-user-password', password);
        localStorage.setItem('claytics-needs-verification', 'true');

        // Mostrar pantalla de verificación restrictiva
        if (loginSection) loginSection.classList.add('hidden');
        if (verifSection) {
            const emailEl = document.getElementById('verification-screen-email');
            if (emailEl) emailEl.textContent = email;
            verifSection.classList.remove('hidden');
        }
    });
}

// ─── PANTALLA DE VERIFICACIÓN ─────────────────────────────
const btnVerifySuccess = document.getElementById('btn-verify-success');
const btnVerifyResend  = document.getElementById('btn-verify-resend');
const linkVerifBack    = document.getElementById('link-verification-back');

if (btnVerifySuccess) {
    btnVerifySuccess.addEventListener('click', () => {
        const name   = localStorage.getItem('claytics-session-name')   || 'Usuario';
        const age    = localStorage.getItem('claytics-session-age')    || '';
        const email  = localStorage.getItem('claytics-session-email')  || '';
        const sector = localStorage.getItem('claytics-session-sector') || 'General';
        localStorage.setItem('claytics-needs-verification', 'false');
        activateSession(name, age, email, sector, false);
    });
}

if (btnVerifyResend) {
    btnVerifyResend.addEventListener('click', () => {
        const email = localStorage.getItem('claytics-session-email') || 'tu@correo.com';
        alert(`¡Correo de verificación reenviado a ${email}!\nRevisa tu bandeja de entrada o spam.`);
    });
}

if (linkVerifBack) {
    linkVerifBack.addEventListener('click', (e) => {
        e.preventDefault();
        if (verifSection) verifSection.classList.add('hidden');
        if (loginSection) loginSection.classList.remove('hidden');
    });
}

// ─── FORMULARIO DE LOGIN (USUARIOS EXISTENTES) ───────────
const loginForm = document.getElementById('login-form');

if (loginForm) {
    loginForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const email    = document.getElementById('login-email').value.trim();
        const password = document.getElementById('login-password').value;

        const savedEmail = localStorage.getItem('claytics-session-email');
        const savedPass  = localStorage.getItem('claytics-user-password');

        if (!savedEmail || !savedPass) {
            alert('No existe una cuenta registrada. Por favor crea una cuenta primero.');
            if (loginFormContainer) loginFormContainer.classList.add('hidden');
            if (registerFormContainer) registerFormContainer.classList.remove('hidden');
            return;
        }

        if (email !== savedEmail) {
            alert('El correo electrónico no corresponde a ninguna cuenta registrada.');
            return;
        }
        if (password !== savedPass) {
            alert('La clave es incorrecta. Por favor inténtalo de nuevo.');
            return;
        }

        const name   = localStorage.getItem('claytics-session-name')   || 'Usuario';
        const age    = localStorage.getItem('claytics-session-age')    || '';
        const sector = localStorage.getItem('claytics-session-sector') || 'General';

        activateSession(name, age, email, sector, false);
    });
}

// ─── SOCIAL LOGIN / OAUTH ─────────────────────────────────
const socialButtons = [
    { id: 'btn-google-login',    provider: 'Google',    defaultName: 'Usuario Google',    defaultSector: 'Tecnología' },
    { id: 'btn-microsoft-login', provider: 'Microsoft', defaultName: 'Usuario Microsoft', defaultSector: 'Finanzas'   },
    { id: 'btn-github-login',    provider: 'GitHub',    defaultName: 'Desarrollador GitHub', defaultSector: 'Tecnología' },
    { id: 'btn-icloud-login',    provider: 'iCloud',    defaultName: 'Usuario iCloud',    defaultSector: 'Educación'  }
];

socialButtons.forEach(btnInfo => {
    const btn = document.getElementById(btnInfo.id);
    if (btn) {
        btn.addEventListener('click', () => {
            const mockEmail = `${btnInfo.defaultName.toLowerCase().replace(/\s+/g, '')}@${btnInfo.provider.toLowerCase()}.com`;
            localStorage.setItem('claytics-user-password', '');
            activateSession(btnInfo.defaultName, '28', mockEmail, btnInfo.defaultSector, false);
        });
    }
});

// ─── CERRAR SESIÓN ────────────────────────────────────────
const logoutBtn = document.getElementById('btn-logout');
if (logoutBtn) {
    logoutBtn.addEventListener('click', () => {
        localStorage.removeItem('claytics-session-name');
        localStorage.removeItem('claytics-session-age');
        localStorage.removeItem('claytics-session-email');
        localStorage.removeItem('claytics-session-sector');
        localStorage.removeItem('claytics-needs-verification');

        document.body.classList.remove('user-logged-in');

        const loggedInLayout = document.getElementById('logged-in-layout');
        if (loggedInLayout) loggedInLayout.classList.add('hidden');

        // Restablecer vista de registro
        if (loginFormContainer) loginFormContainer.classList.add('hidden');
        if (registerFormContainer) registerFormContainer.classList.remove('hidden');
        if (registerForm) registerForm.reset();
        if (loginForm) loginForm.reset();

        if (loginSection) loginSection.classList.remove('hidden');
    });
}


// ═══════════════════════════════════════════════════════
// FEEDBACK LOGIC

// ═══════════════════════════════════════════════════════
const feedbackFab = document.getElementById('feedback-fab');
const feedbackModal = document.getElementById('feedback-modal');
const feedbackCloseBtn = document.getElementById('feedback-close-btn');
const feedbackSubmitBtn = document.getElementById('feedback-submit-btn');
const feedbackText = document.getElementById('feedback-text');
const feedbackStepInput = document.getElementById('feedback-step-input');
const feedbackStepSuccess = document.getElementById('feedback-step-success');

if (feedbackFab && feedbackModal) {
    // Abrir modal
    feedbackFab.addEventListener('click', () => {
        feedbackModal.classList.remove('hidden');
        feedbackStepInput.classList.remove('hidden');
        feedbackStepSuccess.classList.add('hidden');
        feedbackText.value = ''; // limpiar anterior
    });

    // Cerrar modal con botón X
    if (feedbackCloseBtn) {
        feedbackCloseBtn.addEventListener('click', () => {
            feedbackModal.classList.add('hidden');
        });
    }

    // Cerrar clickeando fuera
    feedbackModal.addEventListener('click', (e) => {
        if (e.target === feedbackModal) {
            feedbackModal.classList.add('hidden');
        }
    });

    // Enviar feedback
    if (feedbackSubmitBtn) {
        feedbackSubmitBtn.addEventListener('click', () => {
            const comment = feedbackText.value.trim();
            if (comment === '') return;
            
            // Simular envío a backend
            console.log("Feedback enviado:", comment);
            
            // Cambiar a vista de éxito
            feedbackStepInput.classList.add('hidden');
            feedbackStepSuccess.classList.remove('hidden');
            
            // Cerrar automáticamente después de 3 segundos
            setTimeout(() => {
                feedbackModal.classList.add('hidden');
            }, 3000);
        });
    }
}

// ════════════════════════════════════════════════════════════
// DETECTOR DE MÚLTIPLES TABLAS EN UNA HOJA DE EXCEL
// Detecta tablas separadas: vertical (una debajo de otra)
// y horizontal (una al lado de la otra) y muestra el modal.
// ════════════════════════════════════════════════════════════

/**
 * Analiza una hoja de XLSX y detecta si contiene múltiples tablas
 * separadas por filas o columnas vacías.
 * Retorna un array de objetos de tabla detectada.
 */
function detectTablesInSheet(sheetObj) {
    if (!sheetObj || !sheetObj['!ref']) return [];

    const range = XLSX.utils.decode_range(sheetObj['!ref']);
    const maxRow = range.e.r;
    const maxCol = range.e.c;

    // Construir matriz de presencia de celdas (true = tiene dato)
    const hasData = [];
    for (let r = 0; r <= maxRow; r++) {
        hasData[r] = [];
        for (let c = 0; c <= maxCol; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = sheetObj[addr];
            hasData[r][c] = !!(cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '');
        }
    }

    // Detectar filas completamente vacías
    const emptyRows = new Set();
    for (let r = 0; r <= maxRow; r++) {
        if (hasData[r].every(v => !v)) emptyRows.add(r);
    }

    // Detectar columnas completamente vacías
    const emptyCols = new Set();
    for (let c = 0; c <= maxCol; c++) {
        if (hasData.every(row => !row[c])) emptyCols.add(c);
    }

    // Agrupar índices contiguos no-vacíos en bloques
    function getContiguousBlocks(total, emptySet) {
        const blocks = [];
        let current = [];
        for (let i = 0; i <= total; i++) {
            if (emptySet.has(i)) {
                if (current.length > 0) { blocks.push(current); current = []; }
            } else {
                current.push(i);
            }
        }
        if (current.length > 0) blocks.push(current);
        return blocks;
    }

    const rowBlocks = getContiguousBlocks(maxRow, emptyRows);
    const colBlocks = getContiguousBlocks(maxCol, emptyCols);

    // Si hay solo un bloque de cada tipo → es una sola tabla
    if (rowBlocks.length <= 1 && colBlocks.length <= 1) return [];

    const tables = [];

    // Extraer sub-tabla de una combinación de bloques de fila y columna
    function extractSubTable(rb, cb, orientation, gridPos) {
        const subSheet = {};
        rb.forEach((r, ri) => {
            cb.forEach((c, ci) => {
                const srcAddr = XLSX.utils.encode_cell({ r, c });
                const dstAddr = XLSX.utils.encode_cell({ r: ri, c: ci });
                if (sheetObj[srcAddr]) subSheet[dstAddr] = sheetObj[srcAddr];
            });
        });
        subSheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rb.length - 1, c: cb.length - 1 } });
        const data = XLSX.utils.sheet_to_json(subSheet);
        if (data.length > 0) {
            const cols = Object.keys(data[0] || {});
            tables.push({
                index: tables.length + 1,
                data,
                orientation,
                rows: data.length,
                cols: cols.length,
                columns: cols,
                startRow: rb[0],
                startCol: cb[0],
                gridPos: gridPos || null
            });
        }
    }

    if (rowBlocks.length > 1 && colBlocks.length === 1) {
        // VERTICAL: tablas una debajo de otra
        rowBlocks.forEach(rb => extractSubTable(rb, colBlocks[0], 'vertical', null));
    } else if (colBlocks.length > 1 && rowBlocks.length === 1) {
        // HORIZONTAL: tablas lado a lado
        colBlocks.forEach(cb => extractSubTable(rowBlocks[0], cb, 'horizontal', null));
    } else {
        // CUADRÍCULA: combinación de ambos
        rowBlocks.forEach((rb, ri) => {
            colBlocks.forEach((cb, ci) => {
                extractSubTable(rb, cb, 'grid', `Sector Fila ${ri + 1}, Col ${ci + 1}`);
            });
        });
    }

    return tables;
}

/**
 * Construye un mini preview visual (grilla de celdas) para la tarjeta.
 */
function buildMiniPreview(numCols) {
    const cols = Math.min(numCols, 5);
    const headerRow = `<div class="mini-row">${'<div class="mini-cell header"></div>'.repeat(cols)}</div>`;
    const dataRows  = [0,1,2].map(() => `<div class="mini-row">${'<div class="mini-cell data"></div>'.repeat(cols)}</div>`).join('');
    return headerRow + dataRows;
}

/**
 * Intenta combinar múltiples tablas en una sola (UNION de columnas).
 */
function combineTablesData(tables) {
    const allColumns = new Set();
    tables.forEach(t => t.columns.forEach(c => allColumns.add(c)));
    const cols = [...allColumns];
    const combined = [];
    tables.forEach(tbl => {
        tbl.data.forEach(row => {
            const newRow = { _tabla_origen: `Tabla ${tbl.index}` };
            cols.forEach(c => { newRow[c] = (row[c] !== undefined) ? row[c] : null; });
            combined.push(newRow);
        });
    });
    return combined;
}

/**
 * Carga una tabla seleccionada (individual o combinada) al pipeline.
 */
async function loadTableFromMulti(data) {
    if (!data || data.length === 0) {
        alert('Esta tabla no tiene datos válidos para analizar.');
        return;
    }
    if (!dashboardSection.classList.contains('hidden')) dashboardSection.classList.add('hidden');
    welcomeSection.classList.add('hidden');
    moldingSection.classList.remove('hidden');
    try {
        await processData(data);
    } catch (err) {
        console.error(err);
        alert('Error al cargar la tabla seleccionada.');
        location.reload();
    }
}

/**
 * Muestra el modal interactivo de selección de tablas múltiples.
 */
function showMultiTableModal(tables, wb, sheetName) {
    const modal   = document.getElementById('multitable-modal');
    const cards   = document.getElementById('multitable-cards');
    const countEl = document.getElementById('multitable-count');
    const oriEl   = document.getElementById('multitable-orientation-text');
    if (!modal || !cards) return;

    const oriMap = {
        vertical:   '↕ Tablas en disposición vertical (una debajo de otra)',
        horizontal: '↔ Tablas en disposición horizontal (lado a lado)',
        grid:       '⊞ Tablas en cuadrícula (disposición mixta)'
    };
    const orient = tables[0]?.orientation || 'vertical';
    if (countEl) countEl.textContent = tables.length;
    if (oriEl)   oriEl.textContent   = oriMap[orient] || oriMap.vertical;

    cards.innerHTML = '';
    tables.forEach(tbl => {
        const preview  = buildMiniPreview(tbl.cols);
        const colNames = tbl.columns.slice(0, 4).join(', ') + (tbl.columns.length > 4 ? '…' : '');
        const gridInfo = tbl.gridPos ? `<span><i class="fas fa-border-all"></i> ${tbl.gridPos}</span>` : '';

        const card = document.createElement('div');
        card.className = 'multitable-card';
        card.innerHTML = `
            <div class="multitable-card-icon"><i class="fas fa-table"></i></div>
            <div class="multitable-card-info">
                <div class="multitable-card-name">Tabla ${tbl.index}</div>
                <div class="multitable-card-meta">
                    <span><i class="fas fa-list-ol"></i> ${tbl.rows} filas</span>
                    <span><i class="fas fa-columns"></i> ${tbl.cols} col.</span>
                    ${gridInfo}
                </div>
                <div class="multitable-card-meta" style="margin-top:5px; opacity:0.5; font-size:0.75rem;">
                    ${colNames}
                </div>
            </div>
            <div class="multitable-preview-mini">${preview}</div>
            <i class="fas fa-chevron-right multitable-card-arrow"></i>
        `;
        card.addEventListener('click', () => {
            modal.classList.add('hidden');
            loadTableFromMulti(tbl.data);
        });
        cards.appendChild(card);
    });

    // Botón Combinar
    const combineBtn = document.getElementById('btn-combine-tables');
    if (combineBtn) {
        combineBtn.onclick = () => {
            modal.classList.add('hidden');
            loadTableFromMulti(combineTablesData(tables));
        };
    }

    // Cerrar
    const closeBtn = document.getElementById('multitable-close-btn');
    if (closeBtn) closeBtn.onclick = () => modal.classList.add('hidden');
    modal.onclick = (e) => { if (e.target === modal) modal.classList.add('hidden'); };

    moldingSection.classList.add('hidden');
    modal.classList.remove('hidden');
}

/**
 * Procesa una hoja de Excel con detección de múltiples tablas.
 * Si hay solo una tabla → flujo normal. Si hay varias → muestra el modal.
 */
async function handleExcelSheetWithDetection(wb, sheetName) {
    const sheet = wb.Sheets[sheetName];
    const tables = detectTablesInSheet(sheet);

    if (tables.length > 1) {
        showMultiTableModal(tables, wb, sheetName);
    } else {
        moldingSection.classList.remove('hidden');
        currentSheetName = sheetName;
        const rawData = XLSX.utils.sheet_to_json(sheet);
        await processData(rawData);
    }
}

// ── Inyectar la detección al selector de hojas existente ──
// Cada vez que el usuario seleccione una hoja del modal de hojas,
// pasamos por el detector de tablas antes de procesar.
(function patchSheetSelector() {
    const origShowSheetSelector = showSheetSelector;
    showSheetSelector = function(wb, sheetNames) {
        const sheetModal = document.getElementById('sheet-modal');
        const sheetList  = document.getElementById('sheet-buttons-list');
        if (!sheetModal || !sheetList) return;

        sheetList.innerHTML = '';
        sheetNames.forEach(name => {
            const btn = document.createElement('button');
            btn.className = 'btn-sheet-option';
            const isActive = name === currentSheetName;
            btn.innerHTML = `
                <span>
                    <i class="fas fa-file-invoice" style="margin-right:10px;color:hsl(142,76%,40%)"></i>
                    ${name}${isActive ? ' <span style="font-size:0.7rem;background:#DCFCE7;color:#166534;padding:2px 8px;border-radius:20px;margin-left:6px;">Activa</span>' : ''}
                </span>
                <i class="fas fa-arrow-right"></i>
            `;
            btn.addEventListener('click', async () => {
                sheetModal.classList.add('hidden');
                if (!dashboardSection.classList.contains('hidden')) {
                    dashboardSection.classList.add('hidden');
                } else {
                    welcomeSection.classList.add('hidden');
                }
                moldingSection.classList.remove('hidden');
                currentSheetName = name;
                await handleExcelSheetWithDetection(wb, name);
            });
            sheetList.appendChild(btn);
        });

        const closeBtn = document.getElementById('sheet-close-btn');
        if (closeBtn) closeBtn.onclick = () => sheetModal.classList.add('hidden');
        sheetModal.onclick = (e) => { if (e.target === sheetModal) sheetModal.classList.add('hidden'); };
        sheetModal.classList.remove('hidden');
    };
})();

// ═══════════════════════════════════════════════════════
// NUEVAS FUNCIONALIDADES DE USUARIO Y UX
// ═══════════════════════════════════════════════════════

// 1. Control del Consumo Mensual (Límite 10 análisis)
function updateUsageProgress() {
    const textVal = `${analysisCount} / 10`;
    const percent = Math.min((analysisCount / 10) * 100, 100);
    
    // UI del Sidebar
    const usageText = document.getElementById('usage-meter-text');
    const usageFill = document.getElementById('usage-meter-fill');
    if (usageText) usageText.textContent = textVal;
    if (usageFill) usageFill.style.width = `${percent}%`;
    
    // UI de Plan y Facturación
    const billingUsageText = document.getElementById('billing-usage-text');
    const billingUsageFill = document.getElementById('billing-usage-fill');
    if (billingUsageText) billingUsageText.textContent = `${analysisCount} de 10`;
    if (billingUsageFill) billingUsageFill.style.width = `${percent}%`;
}

// 2. Historial de Archivos Subidos
function addFileToHistory(fileInfo) {
    // Evitar duplicados con el mismo nombre y fecha
    const duplicate = fileHistory.some(f => f.name === fileInfo.name && f.date === fileInfo.date);
    if (duplicate) return;

    fileHistory.unshift(fileInfo); // Añadir al inicio
    if (fileHistory.length > 5) fileHistory.pop(); // Mantener un max de 5 en historial visual
    
    localStorage.setItem('claytics-file-history', JSON.stringify(fileHistory));
    renderHistory();
}

function renderHistory() {
    const sidebarHistoryList = document.getElementById('sidebar-history-list');
    const billingHistoryTableBody = document.getElementById('billing-history-table-body');
    
    // Render en Sidebar
    if (sidebarHistoryList) {
        if (fileHistory.length === 0) {
            sidebarHistoryList.innerHTML = '<li class="history-empty">Ningún archivo subido aún</li>';
        } else {
            sidebarHistoryList.innerHTML = '';
            fileHistory.forEach(file => {
                const li = document.createElement('li');
                li.className = 'history-item';
                li.innerHTML = `<i class="fas fa-file-invoice"></i> ${file.name}`;
                li.title = `${file.name} (${file.size})`;
                li.addEventListener('click', () => loadHistoryFile(file.name, file.rows, file.cols));
                sidebarHistoryList.appendChild(li);
            });
        }
    }
    
    // Render en Pestaña "Mis Datos"
    if (billingHistoryTableBody) {
        if (fileHistory.length === 0) {
            billingHistoryTableBody.innerHTML = `
                <tr>
                    <td colspan="6" style="text-align: center; color: var(--p-text); opacity: 0.6; padding: 30px;">
                        Ningún archivo analizado aún en tu cuenta.
                    </td>
                </tr>`;
        } else {
            billingHistoryTableBody.innerHTML = '';
            fileHistory.forEach(file => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="font-weight: 700; color: var(--p-dark);"><i class="fas fa-file-lines" style="color: var(--p-primary); margin-right: 8px;"></i>${file.name}</td>
                    <td><span style="font-size: 0.75rem; background: #E2E8F0; padding: 2px 8px; border-radius: 6px; font-weight: 700;">${file.type}</span></td>
                    <td>${file.size}</td>
                    <td>${file.rows} filas x ${file.cols} col</td>
                    <td>${file.date}</td>
                    <td><button class="btn-load-history"><i class="fas fa-arrows-rotate"></i> Cargar</button></td>
                `;
                tr.querySelector('.btn-load-history').addEventListener('click', () => {
                    loadHistoryFile(file.name, file.rows, file.cols);
                });
                billingHistoryTableBody.appendChild(tr);
            });
        }
    }
}

// Simular carga de archivos del historial
function loadHistoryFile(fileName, rowsCount, colsCount) {
    const headers = ["ID", "Fecha", "Concepto", "Categoría", "Monto", "Estado"];
    const categories = ["Tecnología", "Salud", "Finanzas", "Alimentos", "Servicios"];
    const mockData = [];
    const rows = parseInt(rowsCount) || 50;
    
    for (let i = 1; i <= rows; i++) {
        const item = {
            "ID": i,
            "Fecha": new Date(2026, 0, i).toLocaleDateString(),
            "Concepto": `Transacción ${i}`,
            "Categoría": categories[i % categories.length],
            "Monto": Math.floor(Math.random() * 5000) + 100,
            "Estado": i % 3 === 0 ? "Pendiente" : "Completado"
        };
        mockData.push(item);
    }
    
    currentFileName = fileName;
    alert(`Cargando sesión recuperada del archivo: ${fileName}`);
    
    // Ocultar todas las secciones y pasar a animación
    welcomeSection.classList.add('hidden');
    const mainSections = document.querySelectorAll('.main-content > section');
    mainSections.forEach(sec => sec.classList.add('hidden'));
    
    moldingSection.classList.remove('hidden');
    
    setTimeout(async () => {
        await processData(mockData);
    }, 1200);
}

// 3. Control de Navegación de Pestañas
function initNavigation() {
    const navButtons = document.querySelectorAll('.nav-btn[data-target]');
    navButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            const target = btn.getAttribute('data-target');
            if (!target) return;
            
            navButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Ocultar todas las secciones principales
            const mainSections = document.querySelectorAll('.main-content > section');
            mainSections.forEach(sec => sec.classList.add('hidden'));
            
            // Si el dashboard está activo y se dio clic en "Nuevo Proyecto"
            if (target === 'welcome-section') {
                if (cachedRawData && cachedRawData.length > 0) {
                    // Si ya se cargó datos, mostramos el dashboard en lugar de welcome
                    dashboardSection.classList.remove('hidden');
                    return;
                }
            }
            
            const targetSec = document.getElementById(target);
            if (targetSec) targetSec.classList.remove('hidden');
        });
    });
}

// 4. Modo Oscuro/Modo Claro
function initTheme() {
    const themeCheckbox = document.getElementById('theme-dark-checkbox');
    const savedTheme = localStorage.getItem('claytics-theme');
    
    if (savedTheme === 'dark') {
        document.body.classList.add('dark-mode');
        if (themeCheckbox) themeCheckbox.checked = true;
    } else {
        document.body.classList.remove('dark-mode');
        if (themeCheckbox) themeCheckbox.checked = false;
    }
    
    if (themeCheckbox) {
        themeCheckbox.addEventListener('change', () => {
            if (themeCheckbox.checked) {
                document.body.classList.add('dark-mode');
                localStorage.setItem('claytics-theme', 'dark');
            } else {
                document.body.classList.remove('dark-mode');
                localStorage.setItem('claytics-theme', 'light');
            }
        });
    }
}

// 5. Cambio de Contraseña (6 dígitos numéricos exactos)
function initChangePassword() {
    const changePasswordForm = document.getElementById('change-password-form');
    if (changePasswordForm) {
        changePasswordForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const currentPass = document.getElementById('current-password').value;
            const newPass     = document.getElementById('new-password').value;
            const confirmPass = document.getElementById('confirm-new-password').value;

            const savedPass = localStorage.getItem('claytics-user-password') || '';

            // Solo aplica para cuentas con contraseña (no social login)
            if (!savedPass) {
                alert('Tu cuenta fue creada con un proveedor externo (Google, iCloud, etc.). No puedes cambiar la contraseña desde aquí.');
                return;
            }

            if (currentPass !== savedPass) {
                alert('La contraseña actual ingresada es incorrecta.');
                return;
            }

            // Validar: exactamente 6 dígitos numéricos
            if (!/^\d{6}$/.test(newPass)) {
                alert('La nueva contraseña debe ser exactamente 6 dígitos numéricos (ej: 123456).');
                return;
            }

            if (newPass !== confirmPass) {
                alert('La confirmación no coincide con la nueva contraseña.');
                return;
            }

            localStorage.setItem('claytics-user-password', newPass);
            alert('¡Contraseña actualizada con éxito!');
            changePasswordForm.reset();
        });
    }
}

// 6. Rellenar Perfil (select para sector) y Guardar Datos
function populateUserProfile(name, age, email, sector) {
    const accName   = document.getElementById('account-name');
    const accAge    = document.getElementById('account-age');
    const accEmail  = document.getElementById('account-email');
    const accSector = document.getElementById('account-sector');

    if (accName)  accName.value  = name  || '';
    if (accAge)   accAge.value   = age   || '';
    if (accEmail) accEmail.value = email || '';

    if (accSector) {
        const knownSectors = ['Tecnología', 'Salud', 'Finanzas', 'Retail', 'Educación', 'Otro'];
        if (knownSectors.includes(sector)) {
            accSector.value = sector;
        } else {
            accSector.value = 'Otro';
            const customGroup = document.getElementById('account-custom-sector-group');
            const customInput = document.getElementById('account-custom-sector');
            if (customGroup) customGroup.classList.remove('hidden');
            if (customInput) customInput.value = sector;
        }
    }
}

function initSaveProfile() {
    // Toggle sector personalizado en perfil
    const accSectorSelect = document.getElementById('account-sector');
    const accCustomGroup  = document.getElementById('account-custom-sector-group');
    const accCustomInput  = document.getElementById('account-custom-sector');

    if (accSectorSelect) {
        accSectorSelect.addEventListener('change', () => {
            if (accSectorSelect.value === 'Otro') {
                if (accCustomGroup) accCustomGroup.classList.remove('hidden');
            } else {
                if (accCustomGroup) accCustomGroup.classList.add('hidden');
                if (accCustomInput) accCustomInput.value = '';
            }
        });
    }

    // Guardar datos del perfil
    const btnSaveProfile = document.getElementById('btn-save-profile');
    if (btnSaveProfile) {
        btnSaveProfile.addEventListener('click', () => {
            const newName = document.getElementById('account-name')?.value.trim();
            const newAge  = document.getElementById('account-age')?.value.trim();
            let   newSector = accSectorSelect ? accSectorSelect.value : '';

            if (newSector === 'Otro') {
                newSector = (accCustomInput && accCustomInput.value.trim()) || 'Otro';
            }

            if (!newName) { alert('El nombre no puede estar vacío.'); return; }

            localStorage.setItem('claytics-session-name',   newName);
            localStorage.setItem('claytics-session-age',    newAge);
            localStorage.setItem('claytics-session-sector', newSector);

            // Actualizar sidebar en tiempo real
            const sidebarName   = document.getElementById('sidebar-user-name');
            const sidebarSector = document.getElementById('sidebar-user-sector');
            if (sidebarName)   sidebarName.textContent   = newName;
            if (sidebarSector) sidebarSector.textContent = newSector;

            alert('¡Datos de perfil actualizados con éxito!');
        });
    }
}

// 7. Simulación de Upgrade a Plan Pro
function initUpgradePlan() {
    const btnUpgrade = document.getElementById('btn-upgrade-pro');
    if (btnUpgrade) {
        btnUpgrade.addEventListener('click', () => {
            const confirmed = confirm('¿Deseas activar Claytics Pro por $19/mes?\n\n(Simulación de pago — sin cargo real)');
            if (!confirmed) return;

            // Actualizar badge del sidebar
            const planBadge = document.getElementById('sidebar-user-plan');
            if (planBadge) planBadge.textContent = 'Plan Pro';

            // Actualizar tarjetas de billing
            const freeBadge = document.getElementById('free-plan-status-badge');
            const proBadge  = document.getElementById('pro-plan-status-badge');
            if (freeBadge) { freeBadge.innerHTML = 'Inactivo'; freeBadge.className = 'plan-card-badge plan-inactive-badge'; }
            if (proBadge)  { proBadge.innerHTML  = '<i class="fas fa-circle-check"></i> Plan Activo'; proBadge.className = 'plan-card-badge plan-free-badge'; }

            // Deshabilitar el límite de cuota
            localStorage.setItem('claytics-analysis-count', '0');
            localStorage.setItem('claytics-plan', 'pro');

            btnUpgrade.textContent = '✓ Plan Pro Activo';
            btnUpgrade.disabled = true;
            btnUpgrade.style.opacity = '0.6';

            alert('¡Bienvenido a Claytics Pro! Ahora tienes análisis ilimitados.');
        });
    }
}

function showVerificationBanner(email) {
    const banner    = document.getElementById('email-verification-banner');
    const emailText = document.getElementById('verification-email-text');
    if (banner && emailText) {
        emailText.textContent = email;
        banner.classList.remove('hidden');
    }
}

function hideVerificationBanner() {
    const banner = document.getElementById('email-verification-banner');
    if (banner) banner.classList.add('hidden');
}

function initVerificationBanner() {
    const resendBtn  = document.getElementById('btn-resend-verification');
    const dismissBtn = document.getElementById('btn-dismiss-verification');
    const banner     = document.getElementById('email-verification-banner');

    if (resendBtn) {
        resendBtn.addEventListener('click', () => {
            const email = localStorage.getItem('claytics-session-email') || 'tu@correo.com';
            alert(`¡Correo de verificación reenviado a ${email}!`);
        });
    }
    if (dismissBtn && banner) {
        dismissBtn.addEventListener('click', () => {
            banner.classList.add('hidden');
            localStorage.setItem('claytics-needs-verification', 'false');
        });
    }
}

// Inicialización al cargar la página (persistencia de sesión)
(function initSessionOnLoad() {
    const savedName   = localStorage.getItem('claytics-session-name');
    const savedAge    = localStorage.getItem('claytics-session-age');
    const savedEmail  = localStorage.getItem('claytics-session-email');
    const savedSector = localStorage.getItem('claytics-session-sector');
    const needsVerify = localStorage.getItem('claytics-needs-verification');

    if (savedName) {
        // Restaurar sidebar
        const sidebarName   = document.getElementById('sidebar-user-name');
        const sidebarSector = document.getElementById('sidebar-user-sector');
        if (sidebarName)   sidebarName.textContent   = savedName;
        if (sidebarSector) sidebarSector.textContent = savedSector;

        populateUserProfile(savedName, savedAge, savedEmail, savedSector);

        // Evitar scroll en body
        document.body.classList.add('user-logged-in');

        // Si el usuario está pendiente de verificación, mostrar pantalla restrictiva
        const loginSec  = document.getElementById('login-section');
        const verifSec  = document.getElementById('verification-section');
        const appLayout = document.getElementById('logged-in-layout');

        if (needsVerify === 'true') {
            // Redirigir a la pantalla de verificación, NO al dashboard
            if (loginSec)  loginSec.classList.add('hidden');
            if (appLayout) appLayout.classList.add('hidden');
            if (verifSec) {
                const emailEl = document.getElementById('verification-screen-email');
                if (emailEl) emailEl.textContent = savedEmail;
                verifSec.classList.remove('hidden');
            }
            document.body.classList.remove('user-logged-in');
        } else {
            // Sesión verificada → entrar directo
            if (loginSec)  loginSec.classList.add('hidden');
            if (verifSec)  verifSec.classList.add('hidden');
            if (appLayout) appLayout.classList.remove('hidden');
            if (welcomeSection) welcomeSection.classList.remove('hidden');
            hideVerificationBanner();
        }
    }

    // Inicializadores UX
    initNavigation();
    initTheme();
    initChangePassword();
    initSaveProfile();
    initUpgradePlan();
    initVerificationBanner();
    updateUsageProgress();
    renderHistory();
})();

