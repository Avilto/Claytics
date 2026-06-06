/**
 * CLAYTICS — KPI Agent & Custom KPI Manager
 * Allows users to create KPIs with custom formulas like SUM({Col}), AVG({Col}), etc.
 * Also suggests relevant KPIs based on the current dataset's numeric columns.
 */

// ──────────────────────────────────────────────
//  State
// ──────────────────────────────────────────────
let customKPIs = [];        // { id, name, icon, formula, value }
let kpiEditId  = null;      // null = new, otherwise = editing id
let kpiSelectedIcon = 'fas fa-chart-line';

// ──────────────────────────────────────────────
//  Formula Engine
//  Supported: SUM({Col}), AVG({Col}), COUNT({Col}),
//             MAX({Col}), MIN({Col}), plain math: {A} * {B}
// ──────────────────────────────────────────────
function evaluateKpiFormula(formula, data) {
    if (!data || data.length === 0) return { ok: false, msg: 'Sin datos cargados.' };

    // Extract all {ColName} references
    const colRefs = [...formula.matchAll(/\{([^}]+)\}/g)].map(m => m[1]);
    if (colRefs.length === 0) return { ok: false, msg: 'Incluye al menos una columna entre llaves: {Columna}' };

    const columns = Object.keys(data[0] || {});
    for (const ref of colRefs) {
        if (!columns.includes(ref)) {
            return { ok: false, msg: `La columna "${ref}" no existe en el dataset.` };
        }
    }

    // Check for aggregation functions: SUM, AVG, COUNT, MAX, MIN
    const fnMatch = formula.match(/^(SUM|AVG|COUNT|MAX|MIN|MEDIAN)\(\{([^}]+)\}\)$/i);
    if (fnMatch) {
        const fn  = fnMatch[1].toUpperCase();
        const col = fnMatch[2];
        const values = data.map(r => parseFloat(r[col])).filter(v => !isNaN(v));

        if (fn !== 'COUNT' && values.length === 0) {
            return { ok: false, msg: `La columna "${col}" no tiene valores numéricos.` };
        }

        let result;
        switch (fn) {
            case 'SUM':    result = values.reduce((a, b) => a + b, 0); break;
            case 'AVG':    result = values.reduce((a, b) => a + b, 0) / values.length; break;
            case 'COUNT':  result = data.filter(r => r[col] !== null && r[col] !== undefined && r[col] !== '').length; break;
            case 'MAX':    result = Math.max(...values); break;
            case 'MIN':    result = Math.min(...values); break;
            case 'MEDIAN': {
                const sorted = [...values].sort((a, b) => a - b);
                const mid = Math.floor(sorted.length / 2);
                result = sorted.length % 2 !== 0 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
                break;
            }
        }
        return { ok: true, value: formatKpiValue(result, fn) };
    }

    // Row-level expression: e.g. {Precio} * {Cantidad}
    // Evaluate per row and sum
    try {
        let jsFormula = formula;
        const colNames = [...formula.matchAll(/\{([^}]+)\}/g)].map(m => m[1]);

        // Check all referenced columns are numeric (best-effort)
        const rowValues = data.map(row => {
            let expr = jsFormula;
            for (const col of colNames) {
                const val = parseFloat(row[col]);
                if (isNaN(val)) return NaN;
                // Escape column name for regex
                expr = expr.replace(new RegExp('\\{' + escapeRegex(col) + '\\}', 'g'), val);
            }
            try {
                // Safe eval: only allow numeric operators
                if (!/^[\d\s\+\-\*\/\.\(\)]+$/.test(expr)) return NaN;
                return Function('"use strict"; return (' + expr + ')')();
            } catch { return NaN; }
        }).filter(v => !isNaN(v));

        if (rowValues.length === 0) return { ok: false, msg: 'La expresión no produjo valores numéricos.' };
        const total = rowValues.reduce((a, b) => a + b, 0);
        return { ok: true, value: formatKpiValue(total, 'SUM') };
    } catch (e) {
        return { ok: false, msg: 'Error en la fórmula: ' + e.message };
    }
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function formatKpiValue(val, fn) {
    if (fn === 'COUNT') return val.toLocaleString('es-PE');
    if (typeof val !== 'number' || isNaN(val)) return '—';
    if (Math.abs(val) >= 1_000_000) return (val / 1_000_000).toFixed(2) + 'M';
    if (Math.abs(val) >= 1_000)     return val.toLocaleString('es-PE', { maximumFractionDigits: 2 });
    return val.toLocaleString('es-PE', { maximumFractionDigits: 4 });
}

// ──────────────────────────────────────────────
//  Rendering
// ──────────────────────────────────────────────
function renderCustomKpis() {
    const grid  = document.getElementById('custom-kpi-grid');
    const badge = document.getElementById('kpi-count-badge');
    if (!grid) return;

    badge && (badge.textContent = customKPIs.length);

    if (customKPIs.length === 0) {
        grid.innerHTML = `
            <div class="kpi-empty-state">
                <i class="fas fa-tachometer-alt"></i>
                <p>Aún no tienes KPIs personalizados.<br>
                Haz clic en <strong>Nuevo KPI</strong> o pide sugerencias a la IA.</p>
            </div>`;
        return;
    }

    grid.innerHTML = customKPIs.map(kpi => `
        <div class="custom-kpi-card" id="ckpi-${kpi.id}">
            <div class="ckpi-icon"><i class="${kpi.icon}"></i></div>
            <div class="ckpi-body">
                <div class="ckpi-label">${kpi.name}</div>
                <div class="ckpi-value">${kpi.value ?? '—'}</div>
                <div class="ckpi-formula">${kpi.formula}</div>
            </div>
            <div class="ckpi-actions">
                <button class="ckpi-btn-edit" onclick="openKpiForm('${kpi.id}')" title="Editar">
                    <i class="fas fa-pen"></i>
                </button>
                <button class="ckpi-btn-delete" onclick="deleteKpi('${kpi.id}')" title="Eliminar">
                    <i class="fas fa-trash"></i>
                </button>
            </div>
        </div>`).join('');
}

// ──────────────────────────────────────────────
//  Formula column pills (clickable shortcuts)
// ──────────────────────────────────────────────
function updateFormulaCols() {
    const container = document.getElementById('kpi-formula-cols');
    if (!container || !cachedRawData || cachedRawData.length === 0) return;

    const numCols = Object.keys(cachedRawData[0] || {}).filter(col => {
        const type = (window.kpiAgent && kpiAgent.columnTypes[col]) || 'Categorical';
        return type === 'Numeric';
    });

    if (numCols.length === 0) {
        container.innerHTML = '';
        return;
    }

    container.innerHTML = `
        <div class="kpi-cols-hint">Columnas numéricas disponibles (clic para insertar):</div>
        <div class="kpi-cols-pills">
            ${numCols.map(c => `<button class="kpi-col-pill" onclick="insertColRef('${c}')">{${c}}</button>`).join('')}
        </div>
        <div class="kpi-fns-hint">Funciones:</div>
        <div class="kpi-cols-pills">
            ${['SUM', 'AVG', 'COUNT', 'MAX', 'MIN', 'MEDIAN'].map(fn =>
                `<button class="kpi-fn-pill" onclick="insertFn('${fn}')">${fn}()</button>`
            ).join('')}
        </div>`;
}

function insertColRef(colName) {
    const input = document.getElementById('kpi-formula-input');
    if (!input) return;
    const pos   = input.selectionStart || input.value.length;
    const before = input.value.slice(0, pos);
    const after  = input.value.slice(pos);
    input.value  = before + `{${colName}}` + after;
    input.focus();
}

function insertFn(fn) {
    const input = document.getElementById('kpi-formula-input');
    if (!input) return;
    const pos   = input.selectionStart || input.value.length;
    const before = input.value.slice(0, pos);
    const after  = input.value.slice(pos);
    input.value  = before + `${fn}({})` + after;
    // Place cursor inside {}
    const newPos = pos + fn.length + 2;
    input.focus();
    input.setSelectionRange(newPos, newPos);
}

// ──────────────────────────────────────────────
//  Form Open / Close
// ──────────────────────────────────────────────
function openKpiForm(editId = null) {
    const formCard = document.getElementById('kpi-form-card');
    const titleEl  = document.getElementById('kpi-form-title');
    const nameIn   = document.getElementById('kpi-name-input');
    const formulaIn= document.getElementById('kpi-formula-input');
    const errEl    = document.getElementById('kpi-form-error');

    kpiEditId = editId;
    errEl && errEl.classList.add('hidden');

    if (editId) {
        const kpi = customKPIs.find(k => k.id === editId);
        if (!kpi) return;
        titleEl.innerHTML  = '<i class="fas fa-pen"></i> Editar KPI';
        nameIn.value       = kpi.name;
        formulaIn.value    = kpi.formula;
        kpiSelectedIcon    = kpi.icon;
        // Sync icon picker
        document.querySelectorAll('.kpi-icon-opt').forEach(b => {
            b.classList.toggle('active', b.dataset.icon === kpi.icon);
        });
    } else {
        titleEl.innerHTML  = '<i class="fas fa-edit"></i> Nuevo KPI';
        nameIn.value       = '';
        formulaIn.value    = '';
        kpiSelectedIcon    = 'fas fa-chart-line';
        document.querySelectorAll('.kpi-icon-opt').forEach((b, i) => b.classList.toggle('active', i === 0));
    }

    updateFormulaCols();
    formCard && formCard.classList.remove('hidden');
    nameIn && nameIn.focus();
}

function closeKpiForm() {
    const formCard = document.getElementById('kpi-form-card');
    formCard && formCard.classList.add('hidden');
    kpiEditId = null;
}

// ──────────────────────────────────────────────
//  Save KPI
// ──────────────────────────────────────────────
function saveKpi() {
    const nameIn   = document.getElementById('kpi-name-input');
    const formulaIn= document.getElementById('kpi-formula-input');
    const errEl    = document.getElementById('kpi-form-error');

    const name    = nameIn.value.trim();
    const formula = formulaIn.value.trim();

    if (!name) {
        showKpiError('Por favor ingresa un nombre para el KPI.');
        return;
    }
    if (!formula) {
        showKpiError('Por favor ingresa una fórmula.');
        return;
    }

    const result = evaluateKpiFormula(formula, cachedRawData);
    if (!result.ok) {
        showKpiError(result.msg);
        return;
    }

    if (kpiEditId) {
        // Edit existing
        const idx = customKPIs.findIndex(k => k.id === kpiEditId);
        if (idx > -1) {
            customKPIs[idx] = { ...customKPIs[idx], name, icon: kpiSelectedIcon, formula, value: result.value };
        }
    } else {
        // New
        customKPIs.push({
            id: 'kpi_' + Date.now(),
            name,
            icon: kpiSelectedIcon,
            formula,
            value: result.value
        });
    }

    closeKpiForm();
    renderCustomKpis();
}

function showKpiError(msg) {
    const errEl = document.getElementById('kpi-form-error');
    if (!errEl) return;
    errEl.textContent = msg;
    errEl.classList.remove('hidden');
}

// ──────────────────────────────────────────────
//  Delete KPI
// ──────────────────────────────────────────────
function deleteKpi(id) {
    customKPIs = customKPIs.filter(k => k.id !== id);
    renderCustomKpis();
    const formCard = document.getElementById('kpi-form-card');
    if (kpiEditId === id) closeKpiForm();
}

// ──────────────────────────────────────────────
//  AI Suggestions
// ──────────────────────────────────────────────
function suggestKpis() {
    const panel    = document.getElementById('kpi-suggestions-panel');
    const listEl   = document.getElementById('kpi-suggestions-list');
    if (!panel || !listEl) return;

    if (!cachedRawData || cachedRawData.length === 0) {
        alert('Carga un archivo primero para obtener sugerencias.');
        return;
    }

    const suggestions = generateKpiSuggestions(cachedRawData);

    if (suggestions.length === 0) {
        listEl.innerHTML = `<p style="padding:16px; color: var(--p-text-light);">No se encontraron columnas numéricas para sugerir KPIs.</p>`;
    } else {
        listEl.innerHTML = suggestions.map((s, i) => `
            <div class="kpi-suggestion-item" id="kpi-sug-${i}">
                <div class="kpi-sug-icon"><i class="${s.icon}"></i></div>
                <div class="kpi-sug-body">
                    <div class="kpi-sug-name">${s.name}</div>
                    <div class="kpi-sug-formula">${s.formula}</div>
                    <div class="kpi-sug-desc">${s.description}</div>
                </div>
                <button class="kpi-sug-add-btn" onclick="addSuggestedKpi(${i})">
                    <i class="fas fa-plus"></i> Agregar
                </button>
            </div>`).join('');
    }

    // Store suggestions temporarily on window for access from onclick
    window._kpiSuggestions = suggestions;
    panel.classList.remove('hidden');
}

function generateKpiSuggestions(data) {
    if (!data || data.length === 0) return [];

    const columns = Object.keys(data[0] || {});
    const numCols = columns.filter(col => {
        const type = (window.kpiAgent && kpiAgent.columnTypes[col]) || guessType(data, col);
        return type === 'Numeric';
    });

    const suggestions = [];
    const iconPool = ['fas fa-chart-line', 'fas fa-dollar-sign', 'fas fa-users', 'fas fa-box', 'fas fa-percent', 'fas fa-star', 'fas fa-arrow-up', 'fas fa-calculator'];

    numCols.forEach((col, idx) => {
        const icon = iconPool[idx % iconPool.length];
        const lower = col.toLowerCase();

        // Determine likely metric type from column name heuristics
        const isRevenue  = /venta|ingreso|revenue|sale|monto|precio|costo|gasto|pago/i.test(lower);
        const isQty      = /cantidad|qty|units|unidad|piezas/i.test(lower);
        const isScore    = /score|rating|calificacion|puntos/i.test(lower);

        if (isRevenue) {
            suggestions.push({
                name: `Total ${col}`,
                formula: `SUM({${col}})`,
                icon: 'fas fa-dollar-sign',
                description: `Suma total de todos los valores en "${col}". Ideal para medir ingresos o egresos acumulados.`
            });
            suggestions.push({
                name: `Promedio ${col}`,
                formula: `AVG({${col}})`,
                icon: 'fas fa-chart-line',
                description: `Valor promedio de "${col}". Útil para analizar el ticket promedio.`
            });
        } else if (isQty) {
            suggestions.push({
                name: `Total ${col}`,
                formula: `SUM({${col}})`,
                icon: 'fas fa-box',
                description: `Suma total de unidades en "${col}".`
            });
            suggestions.push({
                name: `Promedio ${col}`,
                formula: `AVG({${col}})`,
                icon: 'fas fa-hashtag',
                description: `Cantidad promedio por registro en "${col}".`
            });
        } else if (isScore) {
            suggestions.push({
                name: `Score Promedio ${col}`,
                formula: `AVG({${col}})`,
                icon: 'fas fa-star',
                description: `Puntuación media en "${col}".`
            });
            suggestions.push({
                name: `Score Máximo ${col}`,
                formula: `MAX({${col}})`,
                icon: 'fas fa-arrow-up',
                description: `Puntuación máxima registrada en "${col}".`
            });
        } else {
            // Generic numeric column
            suggestions.push({
                name: `Suma de ${col}`,
                formula: `SUM({${col}})`,
                icon,
                description: `Suma acumulada de todos los registros en "${col}".`
            });
            suggestions.push({
                name: `Máximo ${col}`,
                formula: `MAX({${col}})`,
                icon: 'fas fa-arrow-up',
                description: `Valor más alto encontrado en la columna "${col}".`
            });
        }
    });

    // Add a "Record Count" suggestion always
    if (data.length > 0) {
        const firstNumCol = numCols[0];
        if (firstNumCol) {
            suggestions.push({
                name: 'Total de Registros',
                formula: `COUNT({${firstNumCol}})`,
                icon: 'fas fa-list-ol',
                description: 'Conteo total de filas con datos válidos en el dataset.'
            });
        }
    }

    return suggestions.slice(0, 8); // Return max 8 suggestions
}

function guessType(data, col) {
    const sample = data.slice(0, 20).map(r => r[col]).filter(v => v !== null && v !== undefined && v !== '');
    const numCount = sample.filter(v => !isNaN(parseFloat(v))).length;
    return numCount >= sample.length * 0.7 ? 'Numeric' : 'Categorical';
}

function addSuggestedKpi(index) {
    const suggestions = window._kpiSuggestions || [];
    const s = suggestions[index];
    if (!s) return;

    const result = evaluateKpiFormula(s.formula, cachedRawData);
    customKPIs.push({
        id: 'kpi_' + Date.now() + '_' + index,
        name: s.name,
        icon: s.icon,
        formula: s.formula,
        value: result.ok ? result.value : '—'
    });

    // Visual feedback: mark as added
    const btn = document.getElementById(`kpi-sug-${index}`);
    if (btn) {
        btn.querySelector('.kpi-sug-add-btn').innerHTML = '<i class="fas fa-check"></i> Agregado';
        btn.querySelector('.kpi-sug-add-btn').disabled = true;
        btn.classList.add('kpi-sug-added');
    }

    renderCustomKpis();
}

// ──────────────────────────────────────────────
//  Re-evaluate all custom KPIs when data changes
// ──────────────────────────────────────────────
function refreshCustomKpiValues() {
    if (!cachedRawData || customKPIs.length === 0) return;
    customKPIs = customKPIs.map(kpi => {
        const result = evaluateKpiFormula(kpi.formula, cachedRawData);
        return { ...kpi, value: result.ok ? result.value : '—' };
    });
    renderCustomKpis();
}

// ──────────────────────────────────────────────
//  Init (called after DOM ready)
// ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    // New KPI button
    const btnNewKpi = document.getElementById('btn-new-kpi');
    if (btnNewKpi) btnNewKpi.addEventListener('click', () => openKpiForm(null));

    // Close form button
    const btnClose = document.getElementById('kpi-form-close');
    if (btnClose) btnClose.addEventListener('click', closeKpiForm);

    // Save KPI button
    const btnSave = document.getElementById('btn-save-kpi');
    if (btnSave) btnSave.addEventListener('click', saveKpi);

    // Suggest KPIs
    const btnSuggest = document.getElementById('suggest-kpi-btn');
    if (btnSuggest) btnSuggest.addEventListener('click', suggestKpis);

    // Close suggestions panel
    const btnCloseSug = document.getElementById('kpi-suggestions-close');
    if (btnCloseSug) btnCloseSug.addEventListener('click', () => {
        document.getElementById('kpi-suggestions-panel')?.classList.add('hidden');
    });

    // Icon picker
    document.querySelectorAll('.kpi-icon-opt').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.kpi-icon-opt').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            kpiSelectedIcon = btn.dataset.icon;
        });
    });

    // Initial empty render
    renderCustomKpis();
});
