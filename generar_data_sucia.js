/**
 * Generador de Data Sucia para probar el CleanAgent de Claytics
 * Ejecutar: node generar_data_sucia.js
 */

const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

// Instalar xlsx si no está disponible
try {
    require.resolve('xlsx');
} catch (e) {
    console.log('Instalando dependencia xlsx...');
    execSync('npm install xlsx', { stdio: 'inherit' });
}

const XLSX = require('xlsx');

// ══════════════════════════════════════════════
//  DATA CON PROBLEMAS INTENCIONALES
//  Cada fila tiene comentario del error que tiene
// ══════════════════════════════════════════════
const data = [
    // Fila limpia de referencia
    { nombre_cliente: 'María García',     correo: 'maria.garcia@gmail.com',  telefono: '987654321', monto_venta: 1500,       categoria: 'Retail',       fecha_compra: '2024-01-15' },
    { nombre_cliente: 'Carlos López',     correo: 'carlos.lopez@outlook.com',telefono: '912345678', monto_venta: 2300,       categoria: 'Mayorista',    fecha_compra: '2024-01-18' },

    // PROBLEMA: Nombre en MAYÚSCULAS (capitalización)
    { nombre_cliente: 'JUAN PEREZ',       correo: 'juan.perez@hotmail.com',  telefono: '956789012', monto_venta: 850,        categoria: 'Retail',       fecha_compra: '2024-01-20' },
    { nombre_cliente: 'ANA TORRES',       correo: 'ana.torres@gmail.com',    telefono: '934567890', monto_venta: 1200,       categoria: 'Corporativo',  fecha_compra: '2024-01-22' },

    // PROBLEMA: Nombre en minúsculas (capitalización)
    { nombre_cliente: 'rosa mendoza',     correo: 'rosa.mendoza@gmail.com',  telefono: '978901234', monto_venta: 670,        categoria: 'Retail',       fecha_compra: '2024-01-25' },

    // PROBLEMA: Número puro en columna nombre (type mismatch)
    { nombre_cliente: '123456',           correo: 'error@empresa.com',       telefono: '945678901', monto_venta: 320,        categoria: 'Mayorista',    fecha_compra: '2024-01-28' },
    { nombre_cliente: '9999',             correo: 'otro@empresa.com',        telefono: '967890123', monto_venta: 510,        categoria: 'Retail',       fecha_compra: '2024-02-01' },

    // PROBLEMA: Correo sin "@" (email inválido, pero se deja el contenido)
    { nombre_cliente: 'Pedro Castillo',   correo: 'pedro.castillo.empresa',  telefono: '923456789', monto_venta: 1800,       categoria: 'Corporativo',  fecha_compra: '2024-02-03' },
    { nombre_cliente: 'Lucía Vargas',     correo: 'lucia_vargas_sin_arroba', telefono: '989012345', monto_venta: 750,        categoria: 'Retail',       fecha_compra: '2024-02-05' },

    // PROBLEMA: Monto con símbolo de moneda (numérico contaminado)
    { nombre_cliente: 'Roberto Silva',    correo: 'roberto@empresa.pe',      telefono: '901234567', monto_venta: 'S/ 2,450', categoria: 'Mayorista',    fecha_compra: '2024-02-08' },
    { nombre_cliente: 'Carmen Huanca',   correo: 'carmen@negocio.com',      telefono: '956781234', monto_venta: '$1.350',   categoria: 'Retail',       fecha_compra: '2024-02-10' },

    // PROBLEMA: Espacios en blanco al inicio/final de texto
    { nombre_cliente: '  Elena Quispe  ', correo: '  elena@correo.com  ',   telefono: '934561234', monto_venta: 980,        categoria: '  Retail  ',   fecha_compra: '2024-02-12' },
    { nombre_cliente: 'Diego Mamani  ',   correo: 'diego@empresa.com',      telefono: '978901230', monto_venta: 1450,       categoria: 'Mayorista',    fecha_compra: '2024-02-15' },

    // FILA LIMPIA
    { nombre_cliente: 'Sofía Ramos',      correo: 'sofia.ramos@gmail.com',   telefono: '967891230', monto_venta: 2100,       categoria: 'Corporativo',  fecha_compra: '2024-02-18' },

    // PROBLEMA: FILA DUPLICADA (igual a la primera fila)
    { nombre_cliente: 'María García',     correo: 'maria.garcia@gmail.com',  telefono: '987654321', monto_venta: 1500,       categoria: 'Retail',       fecha_compra: '2024-01-15' },

    // FILA LIMPIA
    { nombre_cliente: 'Andrés Cáceres',   correo: 'andres.caceres@pe.com',   telefono: '923451234', monto_venta: 3200,       categoria: 'Corporativo',  fecha_compra: '2024-02-20' },

    // PROBLEMA: FILA COMPLETAMENTE VACÍA
    { nombre_cliente: '',                 correo: '',                        telefono: '',           monto_venta: '',         categoria: '',             fecha_compra: '' },

    // FILA LIMPIA
    { nombre_cliente: 'Patricia Flores',  correo: 'patricia@empresa.pe',     telefono: '945671234', monto_venta: 890,        categoria: 'Retail',       fecha_compra: '2024-02-22' },

    // PROBLEMA: FILA DUPLICADA (igual a la segunda)
    { nombre_cliente: 'Carlos López',     correo: 'carlos.lopez@outlook.com',telefono: '912345678', monto_venta: 2300,       categoria: 'Mayorista',    fecha_compra: '2024-01-18' },

    // PROBLEMA: Monto con texto "N/D"
    { nombre_cliente: 'Julio Paredes',    correo: 'julio@empresa.com',       telefono: '978905678', monto_venta: 'N/D',      categoria: 'Mayorista',    fecha_compra: '2024-02-25' },

    // FILA LIMPIA
    { nombre_cliente: 'Sandra Medina',    correo: 'sandra.medina@gmail.com', telefono: '956784321', monto_venta: 1700,       categoria: 'Retail',       fecha_compra: '2024-02-28' },

    // PROBLEMA: Nombre en minúsculas + espacio extra
    { nombre_cliente: ' gabriel torres',  correo: 'gabriel@correo.com',      telefono: '934562345', monto_venta: 920,        categoria: 'Retail',       fecha_compra: '2024-03-01' },

    // PROBLEMA: SEGUNDA FILA COMPLETAMENTE VACÍA
    { nombre_cliente: null,               correo: null,                       telefono: null,        monto_venta: null,       categoria: null,           fecha_compra: null },

    // FILA LIMPIA
    { nombre_cliente: 'Valeria Cruz',     correo: 'valeria.cruz@empresa.pe', telefono: '967892345', monto_venta: 2800,       categoria: 'Corporativo',  fecha_compra: '2024-03-05' },

    // PROBLEMA: Número puro en nombre
    { nombre_cliente: '00147',            correo: 'cliente147@empresa.com',  telefono: '923453456', monto_venta: 430,        categoria: 'Mayorista',    fecha_compra: '2024-03-08' },

    // FILA LIMPIA
    { nombre_cliente: 'Fernando Salas',   correo: 'fernando.salas@gmail.com',telefono: '945672345', monto_venta: 1350,       categoria: 'Retail',       fecha_compra: '2024-03-10' },
];

// Duplicar la data de prueba con algunas variaciones para llegar a 75 filas
const baseDataLength = data.length;
for (let i = 0; i < 45; i++) {
    const original = data[i % baseDataLength];
    const copy = { ...original };
    if (copy.nombre_cliente && !copy.nombre_cliente.startsWith(' ')) {
        copy.nombre_cliente = copy.nombre_cliente + ` (${i})`;
    }
    if (copy.correo && copy.correo.includes('@')) {
        const parts = copy.correo.split('@');
        copy.correo = parts[0] + `${i}@` + parts[1];
    }
    if (typeof copy.monto_venta === 'number') {
        copy.monto_venta = copy.monto_venta + (i * 10);
    }
    data.push(copy);
}

// Crear el workbook con una columna vacía (para simular columna vacía)
const dataConColumnaVacia = data.map(row => ({
    nombre_cliente: row.nombre_cliente,
    correo:         row.correo,
    columna_vacia:  '',           // ← COLUMNA COMPLETAMENTE VACÍA
    telefono:       row.telefono,
    monto_venta:    row.monto_venta,
    categoria:      row.categoria,
    fecha_compra:   row.fecha_compra,
}));

// Crear workbook y worksheet
const wb = XLSX.utils.book_new();
const ws = XLSX.utils.json_to_sheet(dataConColumnaVacia);

// Estilo de ancho de columnas
ws['!cols'] = [
    { wch: 22 }, // nombre_cliente
    { wch: 30 }, // correo
    { wch: 18 }, // columna_vacia
    { wch: 14 }, // telefono
    { wch: 14 }, // monto_venta
    { wch: 14 }, // categoria
    { wch: 14 }, // fecha_compra
];

XLSX.utils.book_append_sheet(wb, ws, 'Ventas_2024');

const outputPath = path.join(__dirname, 'data_sucia_test.xlsx');
XLSX.writeFile(wb, outputPath);

console.log('\n✅ Archivo generado: data_sucia_test.xlsx');
console.log('\n📋 Resumen de problemas incluidos:');
console.log('   🔴 2 filas completamente vacías');
console.log('   🔴 2 filas duplicadas');
console.log('   🟡 2 nombres con números puros (123456, 9999, 00147)');
console.log('   🟡 2 correos sin @ (email inválido)');
console.log('   🟡 2 montos con símbolo de moneda (S/, $)');
console.log('   🟡 1 monto con texto "N/D"');
console.log('   🟢 3 nombres en MAYÚSCULAS');
console.log('   🟢 3 nombres en minúsculas');
console.log('   🟢 3 celdas con espacios en blanco al inicio/final');
console.log('   🔴 1 columna completamente vacía (columna_vacia)');
console.log('\n📁 Ubícalo en: c:\\Users\\aless\\Desktop\\Claytics\\data_sucia_test.xlsx\n');
