/**
 * CLAYTICS - AGENTE EXPERTO EN ANÁLISIS DE DATOS Y KPIS
 * Versión: Avanzada (Deep Knowledge)
 * Capacidades: Detección profunda de tipos, Data Quality Score, Estadísticas Avanzadas, Heurísticas de Nomenclatura.
 */

class KPIAgent {
    constructor(name = "Kpis") {
        this.name = name;
        this.kpiList = [];
        this.data = null;
        this.columns = [];
        this.columnTypes = {}; // Guardará el tipo de cada columna
        this.manualColumnTypes = {}; // Guardará los tipos modificados manualmente por el usuario
        this.maxKPIs = 6; // Límite de KPIs principales a mostrar para no saturar
    }

    /**
     * Realiza un escaneo profundo de la data
     */
    analyze(data) {
        this.data = data;
        this.columns = Object.keys(data[0] || {});
        this.kpiList = [];
        this.columnTypes = {};

        if (this.columns.length === 0) return [];

        // 1. Detección Profunda de Tipos
        this.detectColumnTypes();

        // 2. Data Quality Score
        this.generateDataQualityKPI();

        // 3. Datos por Fila (columnas del dataset)
        this.kpiList.push({
            id: 'cols_per_row',
            label: 'Datos por Fila',
            value: this.columns.length.toLocaleString(),
            icon: 'fas fa-table-columns',
            type: 'info'
        });

        // 4. Dimensiones totales del dataset (filas × columnas)
        this.kpiList.push({
            id: 'dimensions',
            label: 'Filas y Columnas Totales',
            value: `${data.length.toLocaleString()} Filas / ${this.columns.length} Cols`,
            icon: 'fas fa-border-all',
            type: 'info'
        });

        // 4. Generación de KPIs Numéricos y Estadísticas Avanzadas
        this.generateNumericKPIs();

        // 5. Ordenar y limitar los KPIs
        this.prioritizeKPIs();

        return this.kpiList;
    }

    /**
     * Clasifica cada columna: Numeric, Categorical, Date, Boolean, ID
     */
    detectColumnTypes() {
        const sampleSize = Math.min(this.data.length, 50);

        // Columnas que contienen números pero NO son métricas calculables.
        // Sumar o promediar un DNI o código postal carece de sentido.
        const IDENTIFIER_PATTERNS = [
            'dni', 'ruc', 'documento', 'nro_doc', 'cedula', 'pasaporte',
            'codigo_postal', 'zip', 'cod_postal',
            'id', '_id', 'codigo', 'cod'
        ];

        // Teléfonos los trataremos como Categorical, no como ID ni Numeric.
        const PHONE_PATTERNS = [
            'telefono', 'phone', 'celular', 'cel', 'movil', 'fono', 'telf'
        ];

        this.columns.forEach(col => {
            if (this.manualColumnTypes && this.manualColumnTypes[col]) {
                this.columnTypes[col] = this.manualColumnTypes[col];
                return;
            }
            const lowerCol = col.toLowerCase();
            const samples = this.data.slice(0, sampleSize).map(r => r[col]).filter(v => v !== null && v !== "" && v !== undefined);

            if (samples.length === 0) {
                this.columnTypes[col] = 'Empty';
                return;
            }

            // Exclusión semántica: identificadores y códigos
            const isIdentifier = IDENTIFIER_PATTERNS.some(pat =>
                pat.startsWith('_')
                    ? lowerCol.endsWith(pat)
                    : lowerCol.includes(pat)
            );
            if (isIdentifier) {
                this.columnTypes[col] = 'ID';
                return;
            }

            const isPhone = PHONE_PATTERNS.some(pat => lowerCol.includes(pat));
            if (isPhone) {
                this.columnTypes[col] = 'Categorical'; // Lo forzamos a texto para evitar operaciones matemáticas
                return;
            }

            // Fechas
            const isDate = samples.some(v => String(v).match(/^\d{4}-\d{2}-\d{2}/) || (!isNaN(Date.parse(v)) && isNaN(Number(v))));
            if (isDate) {
                this.columnTypes[col] = 'Date';
                return;
            }

            // Numéricos (solo si no fue excluido arriba)
            const isNumeric = samples.every(v => !isNaN(parseFloat(v)) && isFinite(v));
            if (isNumeric) {
                this.columnTypes[col] = 'Numeric';
                return;
            }

            // Booleanos
            const isBoolean = samples.every(v => {
                const str = String(v).toLowerCase();
                return str === 'true' || str === 'false' || str === 'si' || str === 'no' || str === 'yes' || str === '1' || str === '0';
            });
            if (isBoolean) {
                this.columnTypes[col] = 'Boolean';
                return;
            }

            this.columnTypes[col] = 'Categorical';
        });
    }

    /**
     * Calcula un Data Quality Score sofisticado
     */
    generateDataQualityKPI() {
        let totalCells = this.data.length * this.columns.length;
        let emptyCells = 0;
        let anomalyCells = 0;

        this.data.forEach(row => {
            this.columns.forEach(col => {
                const val = row[col];
                if (val == null || val === "") emptyCells++;
                else if (this.columnTypes[col] === 'Numeric' && (isNaN(parseFloat(val)) || !isFinite(val))) anomalyCells++;
            });
        });

        const missingPenalty = (emptyCells / totalCells) * 100;
        const anomalyPenalty = (anomalyCells / totalCells) * 150; // Anomalías penalizan más
        let score = 100 - missingPenalty - anomalyPenalty;
        score = Math.max(0, Math.min(100, score)); // Entre 0 y 100

        let healthLabel = "Calidad de Data";
        if (score === 100) healthLabel = "Data Impecable";
        else if (score < 80) healthLabel = "Requiere Limpieza";

        this.kpiList.push({
            id: 'quality',
            label: healthLabel,
            value: `${score.toFixed(1)}%`,
            icon: score > 90 ? 'fas fa-shield-check' : 'fas fa-triangle-exclamation',
            type: 'health',
            score: score
        });
    }

    /**
     * Genera Totales, Promedios, Máximos basados en la heurística de la columna
     */
    generateNumericKPIs() {
        this.columns.forEach(col => {
            if (this.columnTypes[col] === 'Numeric') {
                const values = this.data.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
                if (values.length === 0) return;

                const total = values.reduce((sum, v) => sum + v, 0);
                const avg = total / values.length;
                const max = Math.max(...values);
                
                const lowerCol = col.toLowerCase();

                // Decidir qué mostrar según el nombre de la columna
                if (lowerCol.includes('precio') || lowerCol.includes('price') || lowerCol.includes('costo') || lowerCol.includes('calificacion') || lowerCol.includes('rating') || lowerCol.includes('score')) {
                    // Para precios o ratings, el promedio es más importante que el total
                    this.kpiList.push(this.createKPIObj(col, 'avg', avg, "Promedio"));
                    this.kpiList.push(this.createKPIObj(col, 'max', max, "Máximo"));
                } else {
                    // Para ventas, cantidades, ingresos, el total es el rey
                    this.kpiList.push(this.createKPIObj(col, 'total', total, "Total"));
                    // Añadimos el promedio como secundario si hay mucha data
                    if (this.data.length > 10) {
                        this.kpiList.push(this.createKPIObj(col, 'avg', avg, "Promedio"));
                    }
                }
            }
        });
    }

    createKPIObj(col, statType, value, statPrefix) {
        const id = `${statType}_${col.replace(/\s+/g, '_')}`;
        const lowerCol = col.toLowerCase();
        let formatValue;

        // Si es una columna de dinero, formatear como moneda local
        if (lowerCol.includes('venta') || lowerCol.includes('total') || lowerCol.includes('monto') || lowerCol.includes('price')) {
            // El navegador a veces reporta 'es-CO' o 'America/Bogota' para países UTC-5 (como Perú).
            // Por ello, priorizamos la zona horaria y damos opciones.
            const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
            const locale = navigator.language;
            
            const currencyMap = {
                'America/Lima': 'PEN', 'es-PE': 'PEN',
                'America/Mexico_City': 'MXN', 'es-MX': 'MXN',
                'America/Bogota': 'COP', 'es-CO': 'COP',
                'America/Argentina/Buenos_Aires': 'ARS', 'es-AR': 'ARS',
                'America/Santiago': 'CLP', 'es-CL': 'CLP',
                'Europe/Madrid': 'EUR', 'es-ES': 'EUR'
            };

            // 1. Intentar por Zona Horaria (Más preciso geográficamente)
            // 2. Intentar por Idioma del navegador
            // 3. Fallback a PEN (Soles) si estamos en zona UTC-5 ambigua, o USD si no.
            let userCurrency = currencyMap[timeZone] || currencyMap[locale];
            
            // Corrección específica: Si el sistema detecta Bogota pero el usuario está en Perú (ambos UTC-5)
            // En un entorno real se usaría IP, aquí forzamos una comprobación segura.
            if (!userCurrency) userCurrency = 'PEN'; // Fallback por defecto a Soles para este prototipo

            try {
                // Usamos 'narrowSymbol' para forzar "S/" en lugar de "PEN"
                formatValue = new Intl.NumberFormat('es-PE', { 
                    style: 'currency', 
                    currency: userCurrency,
                    currencyDisplay: 'narrowSymbol' 
                }).format(value);
            } catch(e) {
                formatValue = `S/ ${value.toLocaleString(undefined, { minimumFractionDigits: 2 })}`;
            }
        } else {
            formatValue = value > 1000 ? value.toLocaleString() : (Number.isInteger(value) ? value : value.toFixed(2));
        }
        
        return {
            id: id,
            originalCol: col,
            label: this.generateSmartLabel(col, statType),
            value: formatValue,
            icon: this.guessIcon(col, statType),
            type: 'numeric',
            priority: this.calculatePriority(col, statType)
        };
    }

    generateSmartLabel(col, statType) {
        const lower = col.toLowerCase();
        let base = col;
        
        // Mejorar la base
        if (lower === 'total' || lower === 'monto') base = 'Monto';
        if (lower === 'cantidad' || lower === 'qty') base = 'Volumen';
        if (lower.includes('descuento') || lower.includes('discount')) base = 'Descuento';

        // Aplicar el prefijo estadístico
        if (statType === 'avg') return `Promedio de ${base}`;
        if (statType === 'max') return `Mayor ${base}`;
        if (statType === 'total') {
            if (base === 'Monto' || base === 'Volumen') return `${base} Acumulado`;
            return `Total de ${base}`;
        }
        
        return `${statPrefix} ${base}`;
    }

    guessIcon(col, statType) {
        const lower = col.toLowerCase();
        if (statType === 'avg') return 'fas fa-chart-line';
        if (statType === 'max') return 'fas fa-arrow-trend-up';

        if (lower.includes('venta') || lower.includes('monto') || lower.includes('price') || lower.includes('total')) return 'fas fa-money-bill-wave';
        if (lower.includes('cantidad') || lower.includes('qty') || lower.includes('volumen')) return 'fas fa-boxes-stacked';
        if (lower.includes('descuento') || lower.includes('discount')) return 'fas fa-tags';
        if (lower.includes('cliente') || lower.includes('user')) return 'fas fa-users';
        
        return 'fas fa-chart-bar';
    }

    calculatePriority(col, statType) {
        let p = 0;
        const lower = col.toLowerCase();
        if (statType === 'total') p += 10;
        if (statType === 'avg') p += 5;
        if (lower.includes('venta') || lower.includes('total') || lower.includes('monto')) p += 20;
        if (lower.includes('cantidad')) p += 15;
        return p;
    }

    /**
     * Ordena los KPIs por prioridad y limita la cantidad para no abrumar al usuario
     */
    prioritizeKPIs() {
        // Separamos info y salud para que siempre salgan primero
        const fixedKPIs = this.kpiList.filter(k => k.type === 'health' || k.type === 'info');
        let metrics = this.kpiList.filter(k => k.type === 'numeric');

        // Ordenar por prioridad calculada (descendente)
        metrics.sort((a, b) => b.priority - a.priority);

        // Limitar a los N más importantes
        metrics = metrics.slice(0, this.maxKPIs - fixedKPIs.length);

        this.kpiList = [...fixedKPIs, ...metrics];
    }

    renameKPI(id, newName) {
        const kpi = this.kpiList.find(k => k.id === id);
        if (kpi) {
            kpi.label = newName;
            return true;
        }
        return false;
    }
}

const kpiAgent = new KPIAgent();

/**
 * CLAYTICS - AGENTE DE LIMPIEZA DE DATOS (CleanAgent)
 * Versión: Inteligente / Semántica
 * Detecta tipos contextuales por nombre de columna y aplica
 * correcciones apropiadas para cada tipo de dato.
 */
class CleanAgent {
    constructor(name = "Limpieza") {
        this.name = name;
        this.issues = [];
        this.columnTypes = {};
        this.semanticTypes = {};
    }

    /**
     * Detecta el tipo SEMÁNTICO de una columna por su nombre.
     * Esto es diferente al tipo de dato (Numeric/Categorical/etc.)
     * Es el "qué significa" la columna para el negocio.
     */
    detectSemanticType(colName) {
        const l = colName.toLowerCase();
        if (l.includes('email') || l.includes('correo') || l.includes('mail')) return 'email';
        if (l.includes('nombre') || l.includes('name') || l.includes('apellido') || l.includes('cliente') || l.includes('proveedor') || l.includes('persona')) return 'name';
        if (l.includes('telefono') || l.includes('phone') || l.includes('cel') || l.includes('movil') || l.includes('fono')) return 'phone';
        if (l.includes('direccion') || l.includes('address') || l.includes('calle') || l.includes('ubigeo')) return 'address';
        if (l.includes('ruc') || l.includes('dni') || l.includes('documento') || l.includes('nro_doc') || l.includes('cedula')) return 'id_doc';
        if (l.includes('fecha') || l.includes('date') || l.includes('dia') || l.includes('mes')) return 'date';
        if (l.includes('descripcion') || l.includes('description') || l.includes('detalle') || l.includes('observacion') || l.includes('nota')) return 'description';
        return null; // Tipo semántico desconocido
    }

    /**
     * Escanea el dataset completo y genera una lista de problemas detectados.
     * Cada problema incluye: qué es, cuántos casos hay, y cómo repararlo.
     */
    scan(data, columnTypes) {
        this.issues = [];
        this.columnTypes = columnTypes;
        const columns = Object.keys(data[0] || {});

        // Mapear tipos semánticos para cada columna
        columns.forEach(col => {
            this.semanticTypes[col] = this.detectSemanticType(col);
        });

        // ─────────────────────────────────────────
        // PROBLEMA 1: Filas completamente vacías
        // ─────────────────────────────────────────
        const blankRowIndices = [];
        data.forEach((row, i) => {
            const isEmpty = columns.every(col => row[col] === null || row[col] === '' || row[col] === undefined);
            if (isEmpty) blankRowIndices.push(i + 1);
        });
        if (blankRowIndices.length > 0) {
            this.issues.push({
                id: 'blank_rows',
                icon: 'fas fa-minus-circle',
                severity: 'high',
                title: `${blankRowIndices.length} fila(s) completamente vacías`,
                description: 'Estas filas no contienen ningún dato. Eliminarlas mejora el análisis.',
                previewLabel: `Filas afectadas (nro.): ${blankRowIndices.slice(0, 8).join(', ')}${blankRowIndices.length > 8 ? '...' : ''}`,
                fixLabel: 'Eliminar filas vacías',
                fix: (d) => d.filter(row => !columns.every(col => row[col] === null || row[col] === '' || row[col] === undefined))
            });
        }

        // ─────────────────────────────────────────
        // PROBLEMA 2: Columnas completamente vacías
        // ─────────────────────────────────────────
        const blankCols = columns.filter(col =>
            data.every(row => row[col] === null || row[col] === '' || row[col] === undefined)
        );
        if (blankCols.length > 0) {
            this.issues.push({
                id: 'blank_cols',
                icon: 'fas fa-grip-lines-vertical',
                severity: 'high',
                title: `${blankCols.length} columna(s) sin ningún dato`,
                description: 'Columnas vacías que no aportan información al análisis.',
                previewLabel: `Columnas: ${blankCols.join(', ')}`,
                fixLabel: 'Eliminar columnas vacías',
                fix: (d) => d.map(row => {
                    const newRow = { ...row };
                    blankCols.forEach(col => delete newRow[col]);
                    return newRow;
                })
            });
        }

        // ─────────────────────────────────────────
        // PROBLEMA 3: Filas duplicadas
        // ─────────────────────────────────────────
        const seen = new Set();
        const dupCount = data.reduce((count, row) => {
            const key = JSON.stringify(row);
            if (seen.has(key)) return count + 1;
            seen.add(key);
            return count;
        }, 0);
        if (dupCount > 0) {
            this.issues.push({
                id: 'duplicates',
                icon: 'fas fa-copy',
                severity: 'high',
                title: `${dupCount} fila(s) duplicadas`,
                description: 'Registros idénticos que distorsionan totales y promedios.',
                previewLabel: `${dupCount} duplicado(s) serán eliminados manteniendo la primera ocurrencia.`,
                fixLabel: 'Eliminar duplicados',
                fix: (d) => {
                    const seen2 = new Set();
                    return d.filter(row => {
                        const key = JSON.stringify(row);
                        if (seen2.has(key)) return false;
                        seen2.add(key);
                        return true;
                    });
                }
            });
        }

        // ─────────────────────────────────────────
        // PROBLEMA 4: Type mismatch contextual por columna
        // Aplica reglas DISTINTAS según el tipo semántico.
        // ─────────────────────────────────────────
        columns.forEach(col => {
            const sem = this.semanticTypes[col];
            const colType = this.columnTypes[col];

            // Regla: Columnas de NOMBRE/DIRECCIÓN no deben tener números puros
            if (sem === 'name' || sem === 'address') {
                const dirty = data.filter(row => /^\d+$/.test(String(row[col] || '').trim()));
                if (dirty.length > 0) {
                    this.issues.push({
                        id: `type_name_${col}`,
                        icon: 'fas fa-exclamation-triangle',
                        severity: 'medium',
                        title: `"${col}": ${dirty.length} valor(es) son solo números`,
                        description: `Una columna de nombre no debería tener valores puramente numéricos (ej: "1234"). Se vaciarán esas celdas.`,
                        previewLabel: `Ejemplo: "${dirty[0][col]}", "${dirty[Math.min(1, dirty.length-1)][col]}"`,
                        fixLabel: `Limpiar valores inválidos en "${col}"`,
                        fix: (d) => d.map(row => {
                            if (/^\d+$/.test(String(row[col] || '').trim())) return { ...row, [col]: '' };
                            return row;
                        })
                    });
                }
            }

            // Regla: Columnas de EMAIL deben tener "@"
            // (Se DEJA el valor aunque tenga números, solo valida el "@")
            if (sem === 'email') {
                const invalid = data.filter(row => {
                    const val = String(row[col] || '').trim();
                    return val !== '' && !val.includes('@');
                });
                if (invalid.length > 0) {
                    this.issues.push({
                        id: `email_format_${col}`,
                        icon: 'fas fa-at',
                        severity: 'medium',
                        title: `"${col}": ${invalid.length} correo(s) sin formato válido`,
                        description: `Valores sin "@". El contenido (letras y números) se conservará pero se marcará como inválido.`,
                        previewLabel: `Ejemplo: "${invalid[0][col]}"`,
                        fixLabel: `Marcar correos inválidos en "${col}"`,
                        fix: (d) => d.map(row => {
                            const val = String(row[col] || '').trim();
                            if (val !== '' && !val.includes('@')) return { ...row, [col]: `[INVÁLIDO] ${val}` };
                            return row;
                        })
                    });
                }
            }

            // Regla: Columnas NUMÉRICAS con símbolos de moneda u otros caracteres
            if (colType === 'Numeric') {
                const contaminated = data.filter(row => {
                    const val = String(row[col] || '').trim();
                    return val !== '' && isNaN(parseFloat(val)) && val !== 'null';
                });
                if (contaminated.length > 0) {
                    this.issues.push({
                        id: `num_dirty_${col}`,
                        icon: 'fas fa-hashtag',
                        severity: 'medium',
                        title: `"${col}": ${contaminated.length} valor(es) numéricos contaminados`,
                        description: `Pueden contener símbolos como "S/", "$", comas de millar o texto. Se extraerá solo el número.`,
                        previewLabel: `Ejemplo: "${contaminated[0][col]}"`,
                        fixLabel: `Limpiar valores en "${col}"`,
                        fix: (d) => d.map(row => {
                            const val = String(row[col] || '').trim();
                            if (val !== '' && isNaN(parseFloat(val))) {
                                const num = parseFloat(val.replace(/[^0-9.\-]/g, ''));
                                return { ...row, [col]: isNaN(num) ? null : num };
                            }
                            return row;
                        })
                    });
                }
            }
        });

        // ─────────────────────────────────────────
        // PROBLEMA 5: Espacios en blanco innecesarios
        // ─────────────────────────────────────────
        let wsCount = 0;
        columns.forEach(col => {
            if (this.columnTypes[col] === 'Categorical' || this.semanticTypes[col] === 'name' || this.semanticTypes[col] === 'email') {
                data.forEach(row => {
                    const val = String(row[col] || '');
                    if (val !== val.trim() && val.trim() !== '') wsCount++;
                });
            }
        });
        if (wsCount > 0) {
            this.issues.push({
                id: 'whitespace',
                icon: 'fas fa-text-slash',
                severity: 'low',
                title: `${wsCount} celda(s) con espacios innecesarios`,
                description: 'Espacios al inicio o al final de valores de texto que pueden causar errores en búsquedas y filtros.',
                previewLabel: 'Se aplicará un recorte (trim) a todos los campos de texto.',
                fixLabel: 'Recortar espacios en blanco',
                fix: (d) => d.map(row => {
                    const newRow = { ...row };
                    columns.forEach(col => {
                        if (typeof newRow[col] === 'string') newRow[col] = newRow[col].trim();
                    });
                    return newRow;
                })
            });
        }

        // ─────────────────────────────────────────
        // PROBLEMA 6: Capitalización inconsistente en columnas de nombre
        // ─────────────────────────────────────────
        const nameCols = columns.filter(col => this.semanticTypes[col] === 'name');
        if (nameCols.length > 0) {
            const needsNorm = nameCols.some(col =>
                data.some(row => {
                    const v = String(row[col] || '');
                    return v.length > 1 && (v === v.toUpperCase() || v === v.toLowerCase());
                })
            );
            if (needsNorm) {
                this.issues.push({
                    id: 'case_normalize',
                    icon: 'fas fa-font',
                    severity: 'low',
                    title: `Capitalización inconsistente en columna(s) de nombre`,
                    description: `Columnas: ${nameCols.join(', ')}. Valores como "JUAN PEREZ" o "maria gonzalez" se convertirán a "Juan Perez".`,
                    previewLabel: 'Se aplicará formato "Título Case" (primera letra de cada palabra en mayúscula).',
                    fixLabel: 'Normalizar capitalización',
                    fix: (d) => d.map(row => {
                        const newRow = { ...row };
                        nameCols.forEach(col => {
                            if (typeof newRow[col] === 'string') {
                                newRow[col] = newRow[col].toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
                            }
                        });
                        return newRow;
                    })
                });
            }
        }

        return this.issues;
    }

    /**
     * Aplica únicamente las correcciones que el usuario seleccionó.
     */
    applyFixes(data, selectedIssueIds) {
        let result = data.map(row => ({ ...row }));
        const toApply = this.issues.filter(issue => selectedIssueIds.includes(issue.id));
        toApply.forEach(issue => {
            result = issue.fix(result);
        });
        return result;
    }

    /**
     * Genera un resumen del estado de salud de la data después de la limpieza.
     */
    generateCleanSummary(originalCount, cleanedData) {
        return {
            originalRows: originalCount,
            cleanedRows: cleanedData.length,
            removedRows: originalCount - cleanedData.length,
            fixesApplied: this.issues.length
        };
    }
}

const cleanAgent = new CleanAgent();
