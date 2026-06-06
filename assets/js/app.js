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

let userCharts = [];
let chartCounter = 0;
let activeChartFilter = null;

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
                    // Procesar la única hoja directamente
                    welcomeSection.classList.add('hidden');
                    moldingSection.classList.remove('hidden');
                    currentSheetName = sheetNames[0];
                    const rawData = XLSX.utils.sheet_to_json(wb.Sheets[sheetNames[0]]);
                    await processData(rawData);
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
                        const rawData = XLSX.utils.sheet_to_json(wb.Sheets[sheetNames[0]]);
                        await processData(rawData);
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
        if (kpiAgent) {
            kpiAgent.manualColumnTypes = {};
        }

        // Retraso para simular análisis y moldear con animación
        await new Promise(resolve => setTimeout(resolve, 2000));

        // DELEGAMOS AL AGENTE DE KPIS
        currentKPIs = kpiAgent.analyze(data);
        renderDashboard(data);
        
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
    if (insightsContainer) insightsContainer.innerHTML = '';

    // Renderizar KPIs generados por el Agente
    currentKPIs.forEach(kpi => {
        addKPI(kpi);
    });

    // Insight único: nombres de columnas del dataset
    const columns = Object.keys(data[0]);
    if (insightsContainer) renderColumnsInsight(columns, insightsContainer);

    // Renderizar tabla de datos
    renderDataTable(data);

    // Mostrar y popular Constructor de Gráficos
    initChartBuilder(data);

    // Renderizar Gráficas
    renderAllUserCharts();

    // Actualizar barra de hoja activa
    updateSheetBar();

    // Recalcular KPIs personalizados con la nueva data
    if (typeof refreshCustomKpiValues === 'function') refreshCustomKpiValues();
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

    // --- Badges estáticos por columna ---
    const TYPE_META = {
        'Numeric':     { label: 'Número',   cls: 'type-number',  icon: 'fa-hashtag' },
        'Categorical': { label: 'Texto',    cls: 'type-text',    icon: 'fa-font' },
        'Date':        { label: 'Fecha',    cls: 'type-date',    icon: 'fa-calendar-alt' },
        'Boolean':     { label: 'Booleano', cls: 'type-boolean', icon: 'fa-toggle-on' },
        'ID':          { label: 'ID',       cls: 'type-id',      icon: 'fa-fingerprint' },
    };

    const tags = columns.map(col => {
        const colType = kpiAgent.columnTypes[col] || 'Categorical';
        const meta = TYPE_META[colType] || TYPE_META['Categorical'];
        return `
            <div class="col-type-tag">
                <span class="col-name">${col}</span>
                <span class="col-type-badge ${meta.cls}">
                    <i class="fas ${meta.icon}"></i> ${meta.label}
                </span>
            </div>`;
    }).join('');

    // --- Panel único de cambio de tipo ---
    const columnOptions = columns.map(col => {
        const ct = kpiAgent.columnTypes[col] || 'Categorical';
        const m  = TYPE_META[ct] || TYPE_META['Categorical'];
        return `<option value="${col}" data-current="${ct}">${col} (${m.label})</option>`;
    }).join('');

    card.innerHTML = `
        <h4><i class="fas fa-table-columns" style="color:var(--p-primary)"></i> Columnas y Tipos Detectados</h4>
        <div class="col-type-tags-wrap">${tags}</div>

        <div class="type-changer-panel">
            <div class="type-changer-header">
                <i class="fas fa-exchange-alt"></i>
                <span>Cambiar tipo de dato</span>
            </div>
            <div class="type-changer-row">
                <div class="type-changer-field">
                    <label class="type-changer-label">Columna</label>
                    <div class="type-changer-select-wrap">
                        <select id="tc-col-select" class="type-changer-select">
                            ${columnOptions}
                        </select>
                    </div>
                </div>
                <div class="type-changer-arrow"><i class="fas fa-arrow-right"></i></div>
                <div class="type-changer-field">
                    <label class="type-changer-label">Nuevo tipo</label>
                    <div class="type-changer-select-wrap">
                        <select id="tc-type-select" class="type-changer-select"></select>
                    </div>
                </div>
                <button id="tc-apply-btn" class="type-changer-btn">
                    <i class="fas fa-check"></i> Aplicar
                </button>
            </div>
        </div>
    `;
    container.appendChild(card);

    // --- Lógica del panel ---
    const colSelect  = card.querySelector('#tc-col-select');
    const typeSelect = card.querySelector('#tc-type-select');
    const applyBtn   = card.querySelector('#tc-apply-btn');


    const TYPE_OPTIONS = [
        { value: 'Numeric',     label: 'Número',    icon: 'fa-hashtag' },
        { value: 'Categorical', label: 'Texto',     icon: 'fa-font' },
        { value: 'Date',        label: 'Fecha',     icon: 'fa-calendar-alt' },
        { value: 'Boolean',     label: 'Booleano',  icon: 'fa-toggle-on' },
        { value: 'ID',          label: 'ID',        icon: 'fa-fingerprint' },
    ];

    // Sugerencias inteligentes según tipo actual
    const SMART_SUGGESTIONS = {
        'Numeric':     ['Categorical', 'Boolean', 'ID'],
        'Categorical': ['Numeric', 'Date', 'Boolean', 'ID'],
        'Date':        ['Categorical', 'Numeric'],
        'Boolean':     ['Numeric', 'Categorical'],
        'ID':          ['Categorical', 'Numeric'],
    };

    function populateTypeOptions(currentType) {
        const suggestions = SMART_SUGGESTIONS[currentType] || [];
        typeSelect.innerHTML = TYPE_OPTIONS
            .filter(t => t.value !== currentType)
            .map(t => {
                const isSuggested = suggestions.includes(t.value);
                return `<option value="${t.value}" ${isSuggested ? 'data-suggested="true"' : ''}>
                    ${isSuggested ? '⭐ ' : ''}${t.label}
                </option>`;
            }).join('');
        // Seleccionar la primera sugerencia por defecto
        const firstSuggested = typeSelect.querySelector('[data-suggested="true"]');
        if (firstSuggested) firstSuggested.selected = true;
    }

    colSelect.addEventListener('change', () => {
        const currentType = colSelect.selectedOptions[0]?.dataset.current || 'Categorical';
        populateTypeOptions(currentType);
    });

    applyBtn.addEventListener('click', () => {
        const col     = colSelect.value;
        const newType = typeSelect.value;
        if (!col || !newType) return;

        applyBtn.classList.add('applying');
        applyBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Aplicando…';
        setTimeout(() => {
            changeColumnDataType(col, newType);
            applyBtn.classList.remove('applying');
            applyBtn.innerHTML = '<i class="fas fa-check"></i> Aplicar';
        }, 200);
    });

    // Inicializar con la primera columna seleccionada
    if (columns.length > 0) {
        const initType = kpiAgent.columnTypes[columns[0]] || 'Categorical';
        populateTypeOptions(initType);
    }
}

/**
 * Cambia el tipo de dato de una columna, convierte los valores de forma inteligente,
 * persiste el tipo manual en el agente y refresca el Dashboard y el Lab de Análisis.
 */
function changeColumnDataType(col, newType) {
    if (!cachedRawData || cachedRawData.length === 0) return;

    // 1. Persistir el tipo manual en el agente para que sobreviva re-análisis
    kpiAgent.manualColumnTypes[col] = newType;
    kpiAgent.columnTypes[col] = newType;

    // 2. Convertir la data real de forma inteligente
    cachedRawData.forEach(row => {
        let val = row[col];
        if (val === null || val === undefined || val === '') {
            row[col] = null;
            return;
        }

        if (newType === 'Numeric') {
            // Limpiar moneda (S/, $, €, etc.), espacios y comas de miles
            let cleanVal = String(val).replace(/[^\d.\-]/g, '');
            let parsed = parseFloat(cleanVal);
            row[col] = isNaN(parsed) ? 0 : parsed;
        }
        else if (newType === 'Categorical') {
            row[col] = String(val).trim();
        }
        else if (newType === 'Boolean') {
            const strVal = String(val).toLowerCase().trim();
            row[col] = !['false', '0', '', 'no', 'falso', 'null', 'undefined'].includes(strVal);
        }
        else if (newType === 'Date') {
            let date = new Date(val);
            if (!isNaN(date.getTime())) {
                row[col] = date.toISOString().split('T')[0];
            } else {
                row[col] = String(val).trim();
            }
        }
        else if (newType === 'ID') {
            row[col] = String(val).trim();
        }
    });

    // 3. Recalcular KPIs y refrescar el Dashboard completo
    currentKPIs = kpiAgent.analyze(cachedRawData);
    renderDashboard(cachedRawData);
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

    // Poblar selectores de Columna Calculada (solo numéricas)
    const calcColA = document.getElementById('calc-col-a');
    const calcColB = document.getElementById('calc-col-b');
    if (calcColA && calcColB) {
        calcColA.innerHTML = '<option value="">-- Seleccionar Columna --</option>';
        calcColB.innerHTML = '<option value="">-- Seleccionar Columna --</option>';
        columns.forEach(col => {
            if (kpiAgent.columnTypes[col] === 'Numeric') {
                calcColA.innerHTML += `<option value="${col}">${col}</option>`;
                calcColB.innerHTML += `<option value="${col}">${col}</option>`;
            }
        });
    }

    // Inicializar lógica del panel de calculadora si no se ha hecho
    initCalcPanel();

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
    const pdfSimpleBtn = document.getElementById("export-pdf-simple-btn");
    const pdfFullBtn = document.getElementById("export-pdf-full-btn");
    const txtBtn = document.getElementById("export-txt-btn");
    
    if (excelBtn) excelBtn.addEventListener("click", exportToExcel);
    if (csvBtn) csvBtn.addEventListener("click", exportToCSV);
    if (pdfSimpleBtn) pdfSimpleBtn.addEventListener("click", exportToPDF);
    if (pdfFullBtn) pdfFullBtn.addEventListener("click", exportToFullPDF);
    if (txtBtn) txtBtn.addEventListener("click", exportToTXT);
}

// ═══════════════════════════════════════════════════════
//  NUEVAS FUNCIONES: GRÁFICAS, COLUMNA CALCULADA, REPORTE PDF
// ═══════════════════════════════════════════════════════

let builderInitialized = false;

function initChartBuilder(data) {
    const builderSection = document.getElementById('chart-builder-section');
    if (!builderSection) return;
    builderSection.classList.remove('hidden');

    if (builderInitialized) return;
    builderInitialized = true;

    const colXSelect = document.getElementById('builder-col-x');
    const colYSelect = document.getElementById('builder-col-y');
    const columns = Object.keys(data[0] || {});

    colXSelect.innerHTML = '<option value="">-- Seleccionar X --</option>';
    colYSelect.innerHTML = '<option value="">-- Seleccionar Y --</option>';

    columns.forEach(col => {
        if (kpiAgent.columnTypes[col] === 'Categorical' || kpiAgent.columnTypes[col] === 'Date' || kpiAgent.columnTypes[col] === 'Boolean') {
            colXSelect.innerHTML += `<option value="${col}">${col}</option>`;
        }
        if (kpiAgent.columnTypes[col] === 'Numeric') {
            colYSelect.innerHTML += `<option value="${col}">${col}</option>`;
        }
    });

    // Añadir gráfico por defecto si no hay ninguno
    if (userCharts.length === 0) {
        let defaultCat = null;
        let defaultNum = null;
        for (let col of columns) {
            if (!defaultCat && (kpiAgent.columnTypes[col] === 'Categorical' || kpiAgent.columnTypes[col] === 'Date')) defaultCat = col;
            if (!defaultNum && kpiAgent.columnTypes[col] === 'Numeric') defaultNum = col;
            if (defaultCat && defaultNum) break;
        }
        if (defaultCat && defaultNum) {
            addChartConfig('bar', defaultCat, defaultNum);
            addChartConfig('doughnut', defaultCat, defaultNum);
        }
    }

    document.getElementById('btn-add-chart').addEventListener('click', () => {
        const type = document.getElementById('builder-type').value;
        const colX = document.getElementById('builder-col-x').value;
        const colY = document.getElementById('builder-col-y').value;

        if (!colX || !colY) {
            alert('Por favor, selecciona las columnas para el Eje X y el Eje Y.');
            return;
        }

        addChartConfig(type, colX, colY);
        renderAllUserCharts();
    });

    document.getElementById('btn-clear-chart-filter').addEventListener('click', () => {
        activeChartFilter = null;
        document.getElementById('active-chart-filter-badge').classList.add('hidden');
        renderAllUserCharts(); 
    });
}

function addChartConfig(type, colX, colY) {
    chartCounter++;
    userCharts.push({
        id: `custom-chart-${chartCounter}`,
        type: type,
        colX: colX,
        colY: colY,
        instance: null
    });
}

// Hacer removeChartConfig global para que el onClick del HTML funcione
window.removeChartConfig = function(id) {
    userCharts = userCharts.filter(c => c.id !== id);
    renderAllUserCharts();
};

function handleChartClick(event, elements, chart, chartConfig) {
    if (!elements || elements.length === 0) return;
    
    const elementIndex = elements[0].index;
    const label = chart.data.labels[elementIndex];
    
    // Establecer el filtro global interactivo
    activeChartFilter = { col: chartConfig.colX, val: label };
    
    // Mostrar Badge
    const badge = document.getElementById('active-chart-filter-badge');
    const badgeText = document.getElementById('active-filter-text');
    badgeText.textContent = `${chartConfig.colX} = ${label}`;
    badge.classList.remove('hidden');

    // Re-renderizar todo el dashboard basado en este filtro
    renderAllUserCharts();
}

function getActiveChartData() {
    let data = cachedRawData;
    if (activeChartFilter) {
        data = data.filter(row => String(row[activeChartFilter.col]).trim() === String(activeChartFilter.val).trim());
    }
    return data;
}

function renderAllUserCharts() {
    const container = document.getElementById('charts-container');
    if (!container) return;
    
    if (userCharts.length === 0) {
        container.classList.add('hidden');
        return;
    }
    container.classList.remove('hidden');

    // Obtener data (filtrada o completa)
    const currentData = getActiveChartData();

    // Actualizar KPIs si se están viendo gráficos filtrados
    const kpisToRender = kpiAgent.analyze(currentData);
    const kpiContainer = document.getElementById('kpi-container');
    if (kpiContainer) {
        kpiContainer.innerHTML = kpisToRender.map(kpi => `
            <div class="kpi-card">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <span class="label">${kpi.label}</span>
                    <i class="fas ${kpi.icon}"></i>
                </div>
                <span class="value">${kpi.value}</span>
                ${activeChartFilter ? `<span style="font-size: 0.7rem; color: #EF4444;">Filtrado por: ${activeChartFilter.val}</span>` : ''}
            </div>
        `).join('');
    }

    // Limpiar el contenedor de HTML
    container.innerHTML = '';

    const bgColors = [
        '#4F46E5', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', 
        '#06B6D4', '#EC4899', '#14B8A6', '#F97316', '#6366F1',
        '#3B82F6', '#84CC16', '#F43F5E', '#A855F7', '#14B8A6'
    ];
    const bgColorSingle = 'rgba(79, 70, 229, 0.8)';

    userCharts.forEach(chartConfig => {
        // 1. Agregar el HTML del Canvas
        const card = document.createElement('div');
        card.className = 'chart-card';
        card.innerHTML = `
            <button class="btn-remove-chart" onclick="removeChartConfig('${chartConfig.id}')" title="Eliminar gráfico"><i class="fas fa-trash"></i></button>
            <div style="width: 100%; height: 300px;"><canvas id="${chartConfig.id}"></canvas></div>
        `;
        container.appendChild(card);

        // 2. Procesar Data
        const aggregated = {};
        currentData.forEach(row => {
            let key = row[chartConfig.colX];
            let val = parseFloat(row[chartConfig.colY]);
            if (key !== undefined && key !== null && !isNaN(val)) {
                key = String(key).trim();
                if (key.length > 25) key = key.substring(0, 25) + '...';
                aggregated[key] = (aggregated[key] || 0) + val;
            }
        });

        const sortedData = Object.entries(aggregated).sort((a, b) => b[1] - a[1]);
        const maxItems = (chartConfig.type === 'pie' || chartConfig.type === 'doughnut' || chartConfig.type === 'polarArea') ? 10 : 15;
        const slicedData = sortedData.slice(0, maxItems);
            
        const labels = slicedData.map(d => d[0]);
        const values = slicedData.map(d => d[1]);

        // 3. Crear Configuración de Chart.js
        const ctx = document.getElementById(chartConfig.id).getContext('2d');
        
        let datasets = [];
        if (chartConfig.type === 'bar' || chartConfig.type === 'line') {
            datasets = [{
                label: `Total ${chartConfig.colY}`,
                data: values,
                backgroundColor: chartConfig.type === 'bar' ? bgColorSingle : 'rgba(79, 70, 229, 0.2)',
                borderColor: '#4F46E5',
                borderWidth: chartConfig.type === 'line' ? 2 : 0,
                borderRadius: chartConfig.type === 'bar' ? 4 : 0,
                fill: chartConfig.type === 'line'
            }];
        } else {
            datasets = [{
                data: values,
                backgroundColor: bgColors.slice(0, values.length)
            }];
        }

        const newChart = new Chart(ctx, {
            type: chartConfig.type,
            data: { labels: labels, datasets: datasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                onClick: (event, elements, chart) => handleChartClick(event, elements, chart, chartConfig),
                plugins: {
                    legend: { 
                        display: (chartConfig.type !== 'bar' && chartConfig.type !== 'line'),
                        position: 'right'
                    },
                    title: { 
                        display: true, 
                        text: `${chartConfig.colY} por ${chartConfig.colX}`, 
                        font: { size: 14 } 
                    }
                }
            }
        });

        chartConfig.instance = newChart;
    });
}

let calcPanelInitialized = false;

function initCalcPanel() {
    if (calcPanelInitialized) return;
    calcPanelInitialized = true;
    
    const toggleBtn = document.getElementById('calc-toggle-btn');
    const body = document.getElementById('calc-panel-body');
    const icon = document.getElementById('calc-toggle-icon');
    
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

    const btnCreateCol = document.getElementById('btn-create-col');
    if (btnCreateCol) {
        btnCreateCol.addEventListener('click', () => {
            const newName = document.getElementById('calc-new-name').value.trim();
            const colA = document.getElementById('calc-col-a').value;
            const operator = document.getElementById('calc-operator').value;
            const colB = document.getElementById('calc-col-b').value;
            const errorMsg = document.getElementById('calc-error-msg');
            
            if (!newName || !colA || !colB) {
                errorMsg.classList.remove('hidden');
                return;
            }
            errorMsg.classList.add('hidden');

            // Calcular en cachedRawData
            cachedRawData.forEach(row => {
                const valA = parseFloat(row[colA]);
                const valB = parseFloat(row[colB]);
                let result = null;
                
                if (!isNaN(valA) && !isNaN(valB)) {
                    if (operator === '+') result = valA + valB;
                    else if (operator === '-') result = valA - valB;
                    else if (operator === '*') result = valA * valB;
                    else if (operator === '/') result = valB !== 0 ? valA / valB : 0;
                }
                
                row[newName] = result;
            });

            // Forzar que la nueva columna sea numérica
            kpiAgent.columnTypes[newName] = 'Numeric';
            kpiAgent.manualColumnTypes[newName] = 'Numeric';
            
            // Re-analizar KPIs y re-renderizar todo
            currentKPIs = kpiAgent.analyze(cachedRawData);
            renderDashboard(cachedRawData);
            
            // Reset fields
            document.getElementById('calc-new-name').value = '';
            body.classList.add('hidden');
            icon.className = 'fas fa-chevron-down toggle-icon';
            
            // Scroll a la tabla para ver la nueva columna
            document.getElementById('preview-table').scrollIntoView({ behavior: 'smooth', block: 'center' });
        });
    }
}

async function exportToFullPDF() {
    if (!currentKPIs || currentKPIs.length === 0) {
        alert("No hay datos analizados para exportar.");
        return;
    }

    const btn = document.getElementById('export-pdf-full-btn');
    const originalText = btn.innerHTML;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Generando...';
    btn.disabled = true;

    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        const pageW = doc.internal.pageSize.getWidth();   // 210mm
        const pageH = doc.internal.pageSize.getHeight();  // 297mm
        const margin = 14;
        const contentW = pageW - margin * 2;

        // ─── PALETA DE COLORES FIJA ───────────────────────────────────────────
        const C = {
            primary:   [79,  70,  229],   // indigo
            accent:    [139, 92,  246],   // violeta
            success:   [34,  197, 94],    // verde
            warning:   [234, 179, 8],     // amarillo
            dark:      [15,  23,  42],    // fondo oscuro
            mid:       [71,  85,  105],   // gris medio
            light:     [241, 245, 249],   // fondo claro
            white:     [255, 255, 255],
            border:    [203, 213, 225],
        };

        let curY = margin;

        // ─── FUNCIÓN AUXILIAR: nueva página si es necesario ──────────────────
        function checkPage(neededHeight) {
            if (curY + neededHeight > pageH - margin) {
                doc.addPage();
                curY = margin;
            }
        }

        // ─── ENCABEZADO ───────────────────────────────────────────────────────
        // Franja superior degradada simulada con rectángulos superpuestos
        doc.setFillColor(...C.primary);
        doc.rect(0, 0, pageW, 28, 'F');
        doc.setFillColor(...C.accent);
        doc.rect(pageW - 60, 0, 60, 28, 'F');

        // Logo / Título
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(18);
        doc.setTextColor(...C.white);
        doc.text('Claytics', margin, 12);
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(9);
        doc.text('Reporte Especializado', margin, 20);

        // Metadata a la derecha
        doc.setFontSize(8);
        doc.setTextColor(200, 200, 255);
        const fecha = new Date().toLocaleString('es-MX');
        doc.text(fecha, pageW - margin, 12, { align: 'right' });
        doc.text(`Archivo: ${currentFileName}`, pageW - margin, 20, { align: 'right' });

        curY = 36;

        // ─── SECCIÓN KPIs ─────────────────────────────────────────────────────
        // Título de sección
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(11);
        doc.setTextColor(...C.primary);
        doc.text('▌ Indicadores Clave (KPIs)', margin, curY);
        curY += 4;

        // Línea separadora
        doc.setDrawColor(...C.primary);
        doc.setLineWidth(0.5);
        doc.line(margin, curY, pageW - margin, curY);
        curY += 4;

        // Dibujar KPI cards: 3 por fila
        const colsPerRow = 3;
        const cardGap = 4;
        const cardW = (contentW - cardGap * (colsPerRow - 1)) / colsPerRow;
        const cardH = 18;

        const KPI_COLORS = [C.primary, C.accent, [6,182,212], [16,185,129], [245,158,11], [239,68,68]];

        currentKPIs.forEach((kpi, i) => {
            const col = i % colsPerRow;
            const row = Math.floor(i / colsPerRow);

            if (col === 0) checkPage(cardH + 2);

            const x = margin + col * (cardW + cardGap);
            const y = curY + row * (cardH + cardGap);
            const color = KPI_COLORS[i % KPI_COLORS.length];

            // Fondo de tarjeta
            doc.setFillColor(...C.white);
            doc.roundedRect(x, y, cardW, cardH, 2, 2, 'F');

            // Barra lateral de color
            doc.setFillColor(...color);
            doc.roundedRect(x, y, 3, cardH, 1, 1, 'F');

            // Valor (grande)
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(11);
            doc.setTextColor(...C.dark);
            const valStr = String(kpi.value).substring(0, 20);
            doc.text(valStr, x + 7, y + 7);

            // Label (pequeño)
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(7);
            doc.setTextColor(...C.mid);
            const labelStr = String(kpi.label).substring(0, 35);
            doc.text(labelStr, x + 7, y + 13);
        });

        const kpiRows = Math.ceil(currentKPIs.length / colsPerRow);
        curY += kpiRows * (cardH + cardGap) + 6;

        // ─── SECCIÓN TIPOS DE DATOS POR COLUMNA ──────────────────────────────
        checkPage(20);

        doc.setFont('helvetica', 'bold');
        doc.setFontSize(11);
        doc.setTextColor(...C.primary);
        doc.text('▌ Tipos de Datos por Columna', margin, curY);
        curY += 4;
        doc.setDrawColor(...C.primary);
        doc.setLineWidth(0.5);
        doc.line(margin, curY, pageW - margin, curY);
        curY += 4;

        // Colores por tipo
        const TYPE_COLORS = {
            'Numeric':     { bg: [219,234,254], text: [29,78,216],   label: 'Número'   },
            'Categorical': { bg: [237,233,254], text: [109,40,217],  label: 'Texto'    },
            'Date':        { bg: [209,250,229], text: [21,128,61],   label: 'Fecha'    },
            'Boolean':     { bg: [254,243,199], text: [161,98,7],    label: 'Booleano' },
            'ID':          { bg: [226,232,240], text: [71,85,105],   label: 'ID'       },
        };

        if (kpiAgent && kpiAgent.columnTypes) {
            const colEntries = Object.entries(kpiAgent.columnTypes);

            const tableHead = [['Columna', 'Tipo de Dato']];
            const tableBody = colEntries.map(([col, type]) => {
                const meta = TYPE_COLORS[type] || TYPE_COLORS['Categorical'];
                return [col, meta.label];
            });

            checkPage(30);
            doc.autoTable({
                head: tableHead,
                body: tableBody,
                startY: curY,
                margin: { left: margin, right: margin },
                theme: 'grid',
                headStyles: {
                    fillColor: C.primary,
                    textColor: C.white,
                    fontStyle: 'bold',
                    fontSize: 9,
                    cellPadding: 4,
                },
                columnStyles: {
                    0: { cellWidth: contentW * 0.6, fontSize: 8, textColor: C.dark },
                    1: { cellWidth: contentW * 0.4, fontSize: 8, fontStyle: 'bold' }
                },
                bodyStyles: {
                    fontSize: 8,
                    cellPadding: 3,
                    textColor: C.dark,
                },
                alternateRowStyles: {
                    fillColor: C.light,
                },
                didParseCell: function(data) {
                    if (data.section === 'body' && data.column.index === 1) {
                        const typeVal = colEntries[data.row.index]?.[1] || 'Categorical';
                        const meta = TYPE_COLORS[typeVal] || TYPE_COLORS['Categorical'];
                        data.cell.styles.fillColor = meta.bg;
                        data.cell.styles.textColor = meta.text;
                    }
                },
                didDrawPage: function(data) {
                    // Pie de página
                    doc.setFontSize(7);
                    doc.setTextColor(150, 150, 150);
                    doc.text(
                        `Página ${data.pageNumber} • Claytics`,
                        pageW / 2,
                        pageH - 6,
                        { align: 'center' }
                    );
                }
            });
            curY = doc.lastAutoTable.finalY + 8;
        }

        // ─── SECCIÓN DASHBOARDS / GRÁFICAS ───────────────────────────────────
        const chartCanvases = document.querySelectorAll('#charts-container canvas');
        if (chartCanvases.length > 0) {
            checkPage(20);

            doc.setFont('helvetica', 'bold');
            doc.setFontSize(11);
            doc.setTextColor(...C.primary);
            doc.text('▌ Dashboards', margin, curY);
            curY += 4;
            doc.setDrawColor(...C.primary);
            doc.setLineWidth(0.5);
            doc.line(margin, curY, pageW - margin, curY);
            curY += 5;

            // Dos gráficas por fila
            const chartW = (contentW - 6) / 2;
            const chartH = chartW * 0.65; // relación de aspecto

            let chartCol = 0;
            for (const canvasEl of chartCanvases) {
                checkPage(chartH + 4);

                const imgData = canvasEl.toDataURL('image/png', 1.0);
                const x = margin + chartCol * (chartW + 6);

                // Fondo blanco de la gráfica
                doc.setFillColor(...C.white);
                doc.roundedRect(x, curY, chartW, chartH, 2, 2, 'F');
                doc.setDrawColor(...C.border);
                doc.setLineWidth(0.3);
                doc.roundedRect(x, curY, chartW, chartH, 2, 2, 'S');

                // Insertar imagen del canvas
                doc.addImage(imgData, 'PNG', x + 1, curY + 1, chartW - 2, chartH - 2);

                chartCol++;
                if (chartCol >= 2) {
                    chartCol = 0;
                    curY += chartH + 5;
                }
            }
            if (chartCol > 0) curY += chartH + 5;
        }

        // Pie de última página
        doc.setFontSize(7);
        doc.setTextColor(150, 150, 150);
        doc.text(
            `Página ${doc.internal.getNumberOfPages()} • Claytics`,
            pageW / 2,
            pageH - 6,
            { align: 'center' }
        );

        const cleanName = currentFileName.replace(/\.[^/.]+$/, "");
        doc.save(`${cleanName}_Reporte_Especializado.pdf`);

    } catch (error) {
        console.error("Error generando PDF especializado:", error);
        alert("Ocurrió un error generando el reporte especializado.");
    } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
}

// LOGIN LOGIC
// ═══════════════════════════════════════════════════════
const loginForm = document.getElementById('login-form');
const loginSection = document.getElementById('login-section');

if (loginForm && loginSection) {
    loginForm.addEventListener('submit', (e) => {
        e.preventDefault();
        
        // Simular inicio de sesión guardando datos básicos en memoria (o localStorage)
        const name = document.getElementById('login-name').value;
        const age = document.getElementById('login-age').value;
        const email = document.getElementById('login-email').value;
        const sector = document.getElementById('login-sector').value;
        
        console.log(`User logged in: ${name}, ${age}, ${email}, ${sector}`);
        
        // Actualizar UI del Sidebar
        const sidebarName = document.getElementById('sidebar-user-name');
        const sidebarSector = document.getElementById('sidebar-user-sector');
        if (sidebarName) sidebarName.textContent = name || 'Usuario';
        if (sidebarSector) sidebarSector.textContent = sector || 'General';

        // Ocultar login y mostrar layout principal
        loginSection.classList.add('hidden');
        const loggedInLayout = document.getElementById('logged-in-layout');
        if (loggedInLayout) loggedInLayout.classList.remove('hidden');
        welcomeSection.classList.remove('hidden');
    });
}

// Botón de Cerrar Sesión
const logoutBtn = document.getElementById('btn-logout');
if (logoutBtn) {
    logoutBtn.addEventListener('click', () => {
        const loggedInLayout = document.getElementById('logged-in-layout');
        if (loggedInLayout) loggedInLayout.classList.add('hidden');
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
