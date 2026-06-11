// =============================================================
// comercialEngine.js
// Motor de datos para la vista "Resumen Comercial"
// Lee DataF (real 2026), Data 2025 (histórico) y PPTO
// Exporta: processComercialWorkbook, renderResumenComercial
// =============================================================

import * as XLSX from 'xlsx';

// ------------------------------------------------------------------
// CONSTANTES Y ESTRUCTURA DE ÁRBOL EXACTA CON MATCHERS INSENSIBLES
// ------------------------------------------------------------------

function isBonRow(r) {
  if (!r) return false;
  const f = String(r.familia || '').trim().toUpperCase();
  const a = String(r.agrupacion || '').trim().toUpperCase();
  const d = String(r.descProd || '').trim().toUpperCase();
  return f.includes('BON') || a.includes('BON') || d.includes('BON') || f.includes('CCN') || a.includes('CCN') || d.includes('CCN');
}

const cleanStr = (str) => String(str || '').replace(/[\r\n\t\s]+/g, '').toUpperCase();

const cleanA = (str) => String(str || '').toUpperCase().replace(/\s+/g, '');

function matchKeywords(str, ...kws) {
  if (!str) return false;
  const s = String(str).toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  return kws.every(kw => {
    const kwUpper = String(kw).toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    return s.includes(kwUpper);
  });
}

// 1. DICCIONARIO DE MAPEO EXACTO (La única fuente de verdad)
const CATEGORIAS_EXACTAS = {
    // BOTELLAS Y BOTELLON
    "APA BOTELLA 0.5 LTS (x20)": "APA_BOTELLA_0_5_LTS_x20",
    "APA BOTELLA 1.5 LTS (x12)": "APA_BOTELLA_1_5_LTS_x12",
    "APA OTROS": "APA_BOTELLA_0_5LTS_x12",
    "APA BOTELLA 0.5LTS (x12)": "APA_BOTELLA_0_5LTS_x12",
    "APA BOTELLA 0.5 LTS (x12)": "APA_BOTELLA_0_5LTS_x12",
    "APA BOTELLON 5 Gls.": "BT5_Total",
    "APA BOTELLON 18.9 LTS (x1)": "BT5_Total",
    "APA BOTELLON 18.9 LTS (x1": "BT5_Total",
    "APA BOTELLON 18.9 LTS": "BT5_Total",

    // EVP Y EMPAQUES ESPECIALES
    "APA FUNDITA 8 Oz. (x60)": "APA_TETRA_PACK_0_5_LTS_x18",
    "BOTELLA 5LTS (x4)": "MAQUILA_BOTELLA_5LTS_x4",
    "APA BOTELLA 5LTS (x4)": "APA_BOTELLA_5LTS_x4",
    "100% RPET 0.5LTS (x12)": "MAQUILA_100_RPET_0_5LTS_x12",
    "APA 100% RPET 0.5LTS (x12)": "APA_100_RPET_0_5LTS_x12",
    "VASO 250 ML (x24)": "APA_BOTELLA_8_VASOS_1_89_LTS_x6",
    "APA BOTELLA 8 VASOS 1.89 LTS (x6)": "APA_BOTELLA_8_VASOS_1_89_LTS_x6",
    "APA SPORT 0.71 LTS (x12)": "APA_SPORT_0_71_LTS_x12",
    "APA TETRA PACK 0.5 LTS (x18)": "APA_TETRA_PACK_0_5_LTS_x18",

    // MAQUILA
    "MAQUILA AGUA 0.5 LTS (x20)": "MAQUILA_AGUA_0_5_LTS_x20",
    "MAQUILA AGUA 1.5 L TS (x12)": "MAQUILA_AGUA_1_5_L_TS_x12",
    "MAQUILA AGUA 1.5 LTS (x12)": "MAQUILA_AGUA_1_5_L_TS_x12",

    // HIDRACTIVE + (Consolidación solicitada)
    "PA H+ 0.5 LTS (x12)": "HIDRACTIVE_PLUS",
    "PA H+ 0.71 LTS (x12)": "HIDRACTIVE_PLUS",
    "HIDRACTIVE +": "HIDRACTIVE_PLUS",
    "SURTIDO HIDRACTIVE 24 OZ": "HIDRACTIVE_PLUS",
    "HIDRACTIVE 24 OZ": "HIDRACTIVE_PLUS",

    // BEBIDAS Y OTROS
    "ALOE": "PA_SABOR_0_5_LTS_x12",
    "PA SABOR 0.5 LTS (x12)": "PA_SABOR_0_5_LTS_x12",
    "ZUMOS CCN": "BON_Total",
    "MARCAS PRIVADAS": "BON_Total",
    "BON": "BON_Total"
};

// 2. NORMALIZADOR (Elimina espacios ocultos, saltos, acentos y homogeneiza mayúsculas)
const normalizarEstricto = (str) => String(str || '')
  .toUpperCase()
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g, "")
  .replace(/[\r\n\t\s]+/g, '');

// 3. COMPILACIÓN EN MEMORIA
const MAPA_COINCIDENCIA = {};
Object.keys(CATEGORIAS_EXACTAS).forEach(key => {
    MAPA_COINCIDENCIA[normalizarEstricto(key)] = CATEGORIAS_EXACTAS[key];
});

export function classifyRow(r) {
  if (!r) return null;
  
  // Extraemos el valor de la agrupación CEO probando las distintas propiedades conocidas
  let val = '';
  if (r.agrupacionCeo !== undefined && r.agrupacionCeo !== null) {
    val = String(r.agrupacionCeo);
  } else if (r['Agrupacion CEO'] !== undefined && r['Agrupacion CEO'] !== null) {
    val = String(r['Agrupacion CEO']);
  } else if (r['AGRUPACION CEO'] !== undefined && r['AGRUPACION CEO'] !== null) {
    val = String(r['AGRUPACION CEO']);
  } else if (r.agrupacion !== undefined && r.agrupacion !== null) {
    val = String(r.agrupacion);
  } else {
    // Escaneo genérico de llaves por si acaso
    for (let key in r) {
      const uKey = key.trim().toUpperCase();
      if (uKey === 'AGRUPACION CEO' || uKey === 'AGRUPACIÓN CEO' || uKey === 'AGRUPACION_CEO') {
        val = String(r[key]);
        break;
      }
    }
  }

  const agrupLimpia = normalizarEstricto(val);

  // Mapeo Comercial 100% Exacto (SUMIF) por AGRUPACION CEO
  if (MAPA_COINCIDENCIA[agrupLimpia]) {
    return MAPA_COINCIDENCIA[agrupLimpia];
  }

  return null;
}

const ARBOL_COMERCIAL = [
  { 
    id: 'BT5_Total', 
    label: 'BOTELLÓN', 
    type: 'main', 
    matcher: (r) => classifyRow(r) === 'BT5_Total'
  },
  
  { id: 'BOTELLAS_Total', label: 'BOTELLAS', type: 'main' },
    { 
      id: 'APA_BOTELLA_0_5_LTS_x20', 
      parent: 'BOTELLAS_Total', 
      label: 'APA BOTELLA 0.5 LTS (x20)', 
      type: 'item',
      matcher: (r) => classifyRow(r) === 'APA_BOTELLA_0_5_LTS_x20'
    },
    { 
      id: 'APA_BOTELLA_1_5_LTS_x12', 
      parent: 'BOTELLAS_Total', 
      label: 'APA BOTELLA 1.5 LTS (x12)', 
      type: 'item',
      matcher: (r) => classifyRow(r) === 'APA_BOTELLA_1_5_LTS_x12'
    },
    { id: 'APA_OTROS', parent: 'BOTELLAS_Total', label: 'APA OTROS', type: 'sub' },
      { 
        id: 'APA_100_RPET_0_5LTS_x12', 
        parent: 'APA_OTROS', 
        label: 'APA 100% RPET 0.5LTS (x12)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'APA_100_RPET_0_5LTS_x12'
      },
      { 
        id: 'APA_BOTELLA_0_5LTS_x12', 
        parent: 'APA_OTROS', 
        label: 'APA BOTELLA 0.5LTS (x12)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'APA_BOTELLA_0_5LTS_x12'
      },
      { 
        id: 'APA_BOTELLA_5LTS_x4', 
        parent: 'APA_OTROS', 
        label: 'APA BOTELLA 5LTS (x4)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'APA_BOTELLA_5LTS_x4'
      },
      { 
        id: 'APA_BOTELLA_8_VASOS_1_89_LTS_x6', 
        parent: 'APA_OTROS', 
        label: 'APA BOTELLA 8 VASOS 1.89 LTS (x6)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'APA_BOTELLA_8_VASOS_1_89_LTS_x6'
      },
      { 
        id: 'APA_SPORT_0_71_LTS_x12', 
        parent: 'APA_OTROS', 
        label: 'APA SPORT 0.71 LTS (x12)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'APA_SPORT_0_71_LTS_x12'
      },
      { 
        id: 'APA_TETRA_PACK_0_5_LTS_x18', 
        parent: 'APA_OTROS', 
        label: 'APA TETRA PACK 0.5 LTS (x18)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'APA_TETRA_PACK_0_5_LTS_x18'
      },
      
    { 
      id: 'MAQUILA_AGUA_0_5_LTS_x20', 
      parent: 'BOTELLAS_Total', 
      label: 'MAQUILA AGUA 0.5 LTS (x20)', 
      type: 'sub',
      matcher: (r) => classifyRow(r) === 'MAQUILA_AGUA_0_5_LTS_x20'
    },
    { 
      id: 'MAQUILA_AGUA_1_5_L_TS_x12', 
      parent: 'BOTELLAS_Total', 
      label: 'MAQUILA AGUA 1.5 L TS (x12)', 
      type: 'sub',
      matcher: (r) => classifyRow(r) === 'MAQUILA_AGUA_1_5_L_TS_x12'
    },
    { id: 'MAQUILA_OTROS', parent: 'BOTELLAS_Total', label: 'MAQUILA OTROS', type: 'sub' },
      { 
        id: 'MAQUILA_100_RPET_0_5LTS_x12', 
        parent: 'MAQUILA_OTROS', 
        label: '100% RPET 0.5LTS (x12)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'MAQUILA_100_RPET_0_5LTS_x12'
      },
      { 
        id: 'MAQUILA_BOTELLA_5LTS_x4', 
        parent: 'MAQUILA_OTROS', 
        label: 'BOTELLA 5LTS (x4)', 
        type: 'item',
        matcher: (r) => classifyRow(r) === 'MAQUILA_BOTELLA_5LTS_x4'
      },
      {
        id: 'MAQUILA_AGUA_OTROS',
        parent: 'MAQUILA_OTROS',
        label: 'MAQUILA AGUA OTROS',
        type: 'item',
        matcher: (r) => classifyRow(r) === 'MAQUILA_AGUA_OTROS'
      },
      
  { id: 'BEBIDAS_Total', parent: 'BOTELLAS_Total', label: 'BEBIDAS', type: 'sub' },
    { 
      id: 'PA_SABOR_0_5_LTS_x12', 
      parent: 'BEBIDAS_Total', 
      label: 'PA SABOR 0.5 LTS (x12)', 
      type: 'item',
      matcher: (r) => classifyRow(r) === 'PA_SABOR_0_5_LTS_x12'
    },
    { 
      id: 'HIDRACTIVE_PLUS', 
      parent: 'BEBIDAS_Total', 
      label: 'HIDRACTIVE +', 
      type: 'item',
      matcher: (r) => classifyRow(r) === 'HIDRACTIVE_PLUS'
    },
    
  { 
    id: 'BON_Total', 
    label: 'BON', 
    type: 'main',
    matcher: (r) => classifyRow(r) === 'BON_Total'
  }
];

// Cache de estado en módulo
let comercialRawData = null; // { dataF: [], data2025: [], ppto: { vol: [], vta: [] } }

// ------------------------------------------------------------------
// HELPERS ROBUSTOS Y CONTROL DE DATOS
// ------------------------------------------------------------------

function cleanString(val) {
  if (val === undefined || val === null) return '';
  return String(val)
    .normalize("NFD")
    .toUpperCase()
    .replace(/[\u0300-\u036f]/g, "") // Remueve tildes
    .replace(/[^A-Z0-9]/g, "") // Conserva únicamente letras y números
    .trim();
}

function findColumnIndex(headers, keywords) {
  if (!headers) return -1;
  const cleanKeywords = keywords.map(kw => cleanString(kw));
  
  // 1. Coincidencia exacta de términos limpios (ej: "CANTVENDIDA" frente a "CANTVENDIDA")
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === undefined || headers[i] === null) continue;
    const hClean = cleanString(headers[i]);
    if (cleanKeywords.includes(hClean)) {
      return i;
    }
  }
  
  // 2. Coincidencia parcial si falla la coincidencia exacta
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === undefined || headers[i] === null) continue;
    const hClean = cleanString(headers[i]);
    if (cleanKeywords.some(kw => hClean.includes(kw) || kw.includes(hClean))) {
      return i;
    }
  }
  
  return -1;
}

function normalizeText(text) {
  if (!text) return '';
  return text.toString()
    .toUpperCase()
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, ""); // Remueve tildes
}

function excelDateToMonth(val) {
  if (val instanceof Date) return val.getMonth() + 1;
  
  if (val !== undefined && val !== null) {
    const trimmed = String(val).trim();
    const numVal = parseInt(trimmed, 10);
    if (!isNaN(numVal) && numVal >= 1 && numVal <= 12 && (trimmed.length === 1 || trimmed.length === 2)) {
      return numVal;
    }
  }

  const num = Number(val);
  if (!isNaN(num) && num > 25569) { // serial Excel
    const d = new Date(Math.round((num - 25569) * 86400 * 1000));
    return d.getUTCMonth() + 1;
  }
  
  if (val !== undefined && val !== null) {
    const textK = normalizeText(val);
    if (typeof textK !== 'string') return null;
    
    // Si viene en formato "2026-03" o "2026/03" o "03-2026"
    const isYYYYMM = /^\d{4}[-/]\d{2}/.test(textK);
    const isMMYYYY = /^\d{2}[-/]\d{4}/.test(textK);
    if (isYYYYMM) {
      return parseInt(textK.split(/[-/]/)[1], 10);
    }
    if (isMMYYYY) {
      return parseInt(textK.split(/[-/]/)[0], 10);
    }

    // Buscador por nombre de mes
    const monthsMap = {
      'ENE': 1, 'FEB': 2, 'MAR': 3, 'ABR': 4, 'MAY': 5, 'JUN': 6,
      'JUL': 7, 'AGO': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DIC': 12,
      'JAN': 1, 'APR': 4, 'AUG': 8, 'DEC': 12
    };
    for (const key in monthsMap) {
      if (textK.includes(key)) return monthsMap[key];
    }

    // Intentar Parse general de fecha
    const d = new Date(val);
    if (!isNaN(d.getTime())) return d.getMonth() + 1;
  }
  
  return null;
}

function safeNum(v) {
  if (v === undefined || v === null || v === '') return 0;
  if (typeof v === 'number') return v;
  const cleanStr = String(v).replace(/,/g, '').replace(/\$/g, '').trim();
  const n = parseFloat(cleanStr);
  return isNaN(n) ? 0 : n;
}

// Sumar columnas PPTO para rango de meses usando el array robusto monthValues
function sumPptoCols(entry, mesDesde, mesHasta) {
  let total = 0;
  for (let m = mesDesde; m <= mesHasta; m++) {
    total += safeNum(entry.monthValues[m - 1]);
  }
  return total;
}

// ------------------------------------------------------------------
// PARSER ADAPTATIVO: hoja DataF (Real 2026)
// ------------------------------------------------------------------
function parseDataF(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 2) return [];

  // Localizar fila del encabezado inteligente buscando palabras clave (incluyendo agrupación)
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const r = rows[i];
    if (r && r.some(cell => {
      if (cell === null || cell === undefined) return false;
      const s = String(cell).toUpperCase().trim();
      return s.includes('FAMILIA') || s.includes('PERIODO') || s.includes('CANTIDAD') || s.includes('DESC PRODUCT') || s.includes('DESC PROD') || s.includes('DESC PRODUCT.') || s.includes('DESC PROD CF') || s.includes('CATEGORIA') || s.includes('CATEGORIAS') || s.includes('CATEGORIAS CF') || s.includes('AGRUPACION') || s.includes('AGRUPACIÓN');
    })) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = rows[headerRowIdx];
  if (!headers) return [];

  // Localizar columnas 100% dinámicamente sin fallbacks forzados
  let idxAgrupacion = findColumnIndex(headers, ['AGRUPACION CEO', 'AGRUPACIÓN CEO']);
  let idxFamilia = findColumnIndex(headers, ['FAMILIA', 'FAMILY', 'LINEA', 'LÍNEA']);
  let idxCantidad = findColumnIndex(headers, ['CANTIDAD VENDIDA', 'CANT. VENDIDA', 'CANT VENDIDA', 'CANT. VEND', 'CANT VEND', 'CANT VEND.', 'CANTIDAD', 'CANT.', 'CANT', 'VOLUMEN', 'VOL', 'QTY', 'QUANTITY']);
  let idxIngreso = findColumnIndex(headers, ['INGRESO TOTAL', 'INGRESO', 'VENTAS', 'VENTA', 'REVENUE', 'INGRESO NETO', 'MONTO', 'IMPORTE', 'VALOR', 'INGRESO NETO DOP', 'VENTA NETA', 'INGRESO NETO (DOP)', 'VENTAS NETAS']);
  let idxPeriod = findColumnIndex(headers, ['PERIODO', 'PERIOD', 'MES', 'MONTH', 'FECHA', 'PERIODOS', 'MESES']);
  let idxDescProd = findColumnIndex(headers, ['DESC PROD CF', 'DESC PRODUCT.', 'DESC PRODUCT', 'DESC PROD', 'DESCRIPCION', 'DESC PRODUCTO', 'DESC_PRODUCT', 'PRODUCTO', 'DESCRIPCIÓN', 'DESC. PRODUCTO', 'DESCRIPCIÓN DEL PRODUCTO', 'DESCRIPCION PRODUCTO']);

  console.log('🔍 [comercialEngine] Columnas detectadas dinámicamente en DataF (2026):', {
    periodo: idxPeriod,
    descProd: idxDescProd,
    agrupacion: idxAgrupacion,
    familia: idxFamilia,
    cantidad: idxCantidad,
    ingreso: idxIngreso
  });

  const result = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    
    const periodVal = (idxPeriod >= 0 && idxPeriod < r.length) ? r[idxPeriod] : null;
    const descProdVal = (idxDescProd >= 0 && idxDescProd < r.length) ? r[idxDescProd] : null;
    const agrupacionVal = (idxAgrupacion >= 0 && idxAgrupacion < r.length) ? r[idxAgrupacion] : '';
    const familiaVal = (idxFamilia >= 0 && idxFamilia < r.length) ? r[idxFamilia] : '';
    const cantidadVal = (idxCantidad >= 0 && idxCantidad < r.length) ? r[idxCantidad] : 0;
    const ingresoVal = (idxIngreso >= 0 && idxIngreso < r.length) ? r[idxIngreso] : 0;

    if (periodVal === null || periodVal === undefined) continue;
    const mes = excelDateToMonth(periodVal);
    if (!mes) continue;

    const descProd = descProdVal ? String(descProdVal).trim() : '';

    result.push({
      source: '2026',
      mes,
      descProd,
      agrupacion: agrupacionVal ? String(agrupacionVal).trim() : '',
      familia: familiaVal ? String(familiaVal).trim() : '',
      cantidad: safeNum(cantidadVal),
      ingreso: safeNum(ingresoVal),
    });
  }

  console.log(`📊 [comercialEngine] DataF procesada exitosamente. Filas: ${result.length}`);
  if (result.length > 0) {
    const uniqueAgrp = Array.from(new Set(result.map(r => r.agrupacion))).slice(0, 30);
    console.log("🔍 [comercialEngine] Primeras 5 filas en DataF:", result.slice(0, 5));
    console.log("🔍 [comercialEngine] Valores de 'agrupacion' únicos (primeros 30) en DataF:", uniqueAgrp);
  }
  return result;
}

// ------------------------------------------------------------------
// PARSER ADAPTATIVO: hoja Data 2025 (Histórico)
// ------------------------------------------------------------------
function parseData2025(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 2) return [];

  // Localizar fila del encabezado inteligente buscando palabras clave (incluyendo agrupación)
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const r = rows[i];
    if (r && r.some(cell => {
      if (cell === null || cell === undefined) return false;
      const s = String(cell).toUpperCase().trim();
      return s.includes('FAMILIA') || s.includes('PERIODO') || s.includes('CANTIDAD') || s.includes('DESC PRODUCT') || s.includes('DESC PROD') || s.includes('DESC PRODUCT.') || s.includes('DESC PROD CF') || s.includes('CATEGORIA') || s.includes('CATEGORIAS') || s.includes('CATEGORIAS CF') || s.includes('AGRUPACION') || s.includes('AGRUPACIÓN');
    })) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = rows[headerRowIdx];
  if (!headers) return [];

  // Localizar columnas 100% dinámicamente sin fallbacks forzados
  let idxAgrupacion = findColumnIndex(headers, ['AGRUPACION CEO', 'AGRUPACIÓN CEO']);
  let idxFamilia = findColumnIndex(headers, ['FAMILIA', 'FAMILY', 'LINEA', 'LÍNEA']);
  let idxDescProd = findColumnIndex(headers, ['DESC PROD CF', 'DESC PRODUCT.', 'DESC PRODUCT', 'DESC PROD', 'DESCRIPCION', 'DESC PRODUCTO', 'DESC_PRODUCT', 'PRODUCTO', 'DESCRIPCIÓN', 'DESC. PRODUCTO', 'DESCRIPCIÓN DEL PRODUCTO', 'DESCRIPCION PRODUCTO']);
  let idxPeriod = findColumnIndex(headers, ['PERIODO', 'PERIOD', 'MES', 'MONTH', 'FECHA', 'PERIODOS', 'MESES']);
  let idxCantidad = findColumnIndex(headers, ['CANTIDAD VENDIDA', 'CANT. VENDIDA', 'CANT VENDIDA', 'CANT. VEND', 'CANT VEND', 'CANT VEND.', 'CANTIDAD', 'CANT.', 'CANT', 'VOLUMEN', 'VOL', 'QTY', 'QUANTITY']);
  let idxIngreso = findColumnIndex(headers, ['INGRESO TOTAL', 'INGRESO', 'VENTAS', 'VENTA', 'REVENUE', 'INGRESO NETO', 'MONTO', 'IMPORTE', 'VALOR', 'INGRESO NETO DOP', 'VENTA NETA', 'INGRESO NETO (DOP)', 'VENTAS NETAS']);

  console.log('🔍 [comercialEngine] Columnas detectadas dinámicamente en Data 2025:', {
    periodo: idxPeriod,
    descProd: idxDescProd,
    agrupacion: idxAgrupacion,
    familia: idxFamilia,
    cantidad: idxCantidad,
    ingreso: idxIngreso
  });

  const result = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;

    const periodVal = (idxPeriod >= 0 && idxPeriod < r.length) ? r[idxPeriod] : null;
    const descProdVal = (idxDescProd >= 0 && idxDescProd < r.length) ? r[idxDescProd] : null;
    const agrupacionVal = (idxAgrupacion >= 0 && idxAgrupacion < r.length) ? r[idxAgrupacion] : '';
    const familiaVal = (idxFamilia >= 0 && idxFamilia < r.length) ? r[idxFamilia] : '';
    const cantidadVal = (idxCantidad >= 0 && idxCantidad < r.length) ? r[idxCantidad] : 0;
    const ingresoVal = (idxIngreso >= 0 && idxIngreso < r.length) ? r[idxIngreso] : 0;

    if (periodVal === null || periodVal === undefined) continue;
    
    // El mes puede venir directamente como un entero (1-12) o una fecha
    let mes = parseInt(periodVal);
    if (isNaN(mes) || mes < 1 || mes > 12) {
      mes = excelDateToMonth(periodVal);
    }
    if (!mes) continue;

    const descProd = descProdVal ? String(descProdVal).trim() : '';

    result.push({
      source: '2025',
      mes,
      descProd,
      agrupacion: agrupacionVal ? String(agrupacionVal).trim() : '',
      familia: familiaVal ? String(familiaVal).trim() : '',
      cantidad: safeNum(cantidadVal),
      ingreso: safeNum(ingresoVal),
    });
  }

  console.log(`📊 [comercialEngine] Data 2025 procesada exitosamente. Filas: ${result.length}`);
  if (result.length > 0) {
    const uniqueAgrp = Array.from(new Set(result.map(r => r.agrupacion))).slice(0, 30);
    console.log("🔍 [comercialEngine] Primeras 5 filas en Data 2025:", result.slice(0, 5));
    console.log("🔍 [comercialEngine] Valores de 'agrupacion' únicos (primeros 30) en Data 2025:", uniqueAgrp);
  }
  return result;
}

// ------------------------------------------------------------------
// PARSER ADAPTATIVO: hoja PPTO (Presupuesto)
// ------------------------------------------------------------------
function parsePPTO(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: 0 });
  if (rows.length < 2) return { vol: [], vta: [] };

  // Localizar fila del encabezado inteligente buscando palabras clave
  let headerRowIdx = 1;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const r = rows[i];
    if (r && r.some(cell => {
      const s = String(cell).toUpperCase().trim();
      return s.includes('TIPO') || s.includes('DESC MATERIAL') || s.includes('DESC PRODUCT') || s.includes('DESCRIPCION');
    })) {
      headerRowIdx = i;
      break;
    }
  }

  const headers = rows[headerRowIdx];
  if (!headers) return { vol: [], vta: [] };

  const idxTipo = findColumnIndex(headers, ['TIPO', 'TYPE', 'METRICA', 'MÉTRICA']);
  const idxDescProd = findColumnIndex(headers, ['DESC PRODUCT.', 'DESC PROD', 'DESCRIPCION', 'DESC PRODUCTO', 'DESC_PRODUCT', 'PRODUCTO', 'DESCRIPCIÓN']);
  let idxAgrupacion = findColumnIndex(headers, ['AGRUPACION CEO', 'AGRUPACIÓN CEO']);

  // Mapear dinámicamente las columnas de los 12 meses
  const monthCols = Array(12).fill(-1);
  const monthLabels = [
    ['1', 'ENE', 'ENERO', 'JAN', 'JANUARY'],
    ['2', 'FEB', 'FEBRERO', 'FEB', 'FEBRUARY'],
    ['3', 'MAR', 'MARZO', 'MAR', 'MARCH'],
    ['4', 'ABR', 'ABRIL', 'APR', 'APRIL'],
    ['5', 'MAY', 'MAYO', 'MAY'],
    ['6', 'JUN', 'JUNIO', 'JUN', 'JUNE'],
    ['7', 'JUL', 'JULIO', 'JUL', 'JULY'],
    ['8', 'AGO', 'AGOSTO', 'AUG', 'AUGUST'],
    ['9', 'SEP', 'SEPTIEMBRE', 'SEP', 'SEPTEMBER'],
    ['10', 'OCT', 'OCTUBRE', 'OCT', 'OCTOBER'],
    ['11', 'NOV', 'NOVIEMBRE', 'NOV', 'NOVEMBER'],
    ['12', 'DIC', 'DICIEMBRE', 'DEC', 'DECEMBER']
  ];

  for (let m = 0; m < 12; m++) {
    const kws = monthLabels[m];
    const startSearch = Math.max(0, idxAgrupacion + 1);
    for (let c = startSearch; c < headers.length; c++) {
      if (headers[c] === undefined || headers[c] === null) continue;
      const hStr = normalizeText(headers[c]);
      const prevRowCell = headerRowIdx > 0 && rows[headerRowIdx - 1] ? normalizeText(rows[headerRowIdx - 1][c]) : '';
      
      const matchHeader = kws.includes(hStr) || kws.some(kw => hStr === kw || hStr.startsWith(kw + '-') || hStr.startsWith(kw + ' '));
      const matchPrev = prevRowCell && (kws.includes(prevRowCell) || kws.some(kw => prevRowCell === kw || prevRowCell.startsWith(kw + '-') || prevRowCell.startsWith(kw + ' ')));
      
      if (matchHeader || matchPrev) {
        monthCols[m] = c;
        break;
      }
    }
  }

  // Fallback a columnas fijas si no se detectó alguna columna del mes
  for (let m = 0; m < 12; m++) {
    if (monthCols[m] === -1) {
      monthCols[m] = 11 + m;
    }
  }

  console.log('🔍 [comercialEngine] Columnas detectadas en PPTO:', {
    tipo: idxTipo,
    descProd: idxDescProd,
    agrupacion: idxAgrupacion,
    monthCols
  });

  const vol = [];
  const vta = [];

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;

    const tipoVal = (idxTipo >= 0 && idxTipo < r.length) ? String(r[idxTipo] || '').trim().toLowerCase() : '';
    const descProdVal = (idxDescProd >= 0 && idxDescProd < r.length) ? String(r[idxDescProd] || '').trim() : '';
    const agrupacionVal = (idxAgrupacion >= 0 && idxAgrupacion < r.length) ? String(r[idxAgrupacion] || '').trim() : '';

    if (!descProdVal && !agrupacionVal) continue;

    const entry = {
      descProd: descProdVal || '',
      agrupacion: agrupacionVal,
      monthValues: monthCols.map(colIdx => safeNum(r[colIdx]))
    };

    if (tipoVal.includes('volumen') || tipoVal.includes('cantidad') || tipoVal === 'volumen') {
      vol.push(entry);
    } else if (tipoVal.includes('venta') || tipoVal.includes('ingreso') || tipoVal.includes('ventas') || tipoVal.includes('monto')) {
      vta.push(entry);
    }
  }

  console.log(`📊 [comercialEngine] PPTO procesado exitosamente. Vol y Vta en cache: ${vol.length}, ${vta.length}`);
  if (vol.length > 0) {
    const uniqueAgrp = Array.from(new Set(vol.map(r => r.agrupacion))).slice(0, 30);
    console.log("🔍 [comercialEngine] Primeras 5 filas vol en PPTO:", vol.slice(0, 5));
    console.log("🔍 [comercialEngine] Valores de 'agrupacion' únicos (primeros 30) en PPTO vol:", uniqueAgrp);
  }
  return { vol, vta };
}

// ------------------------------------------------------------------
// AGREGACIONES DE NEGOCIO
// ------------------------------------------------------------------
function agregaData(rows, matcher, mesDesde, mesHasta) {
  let cantidad = 0, ingreso = 0;
  if (!matcher) return { cantidad, ingreso };
  rows.forEach(r => {
    if (r.mes >= mesDesde && r.mes <= mesHasta && matcher(r)) {
      cantidad += r.cantidad;
      ingreso  += r.ingreso;
    }
  });
  return { cantidad, ingreso };
}

function agregaDataHybrid(dataF, sixPlusSix, matcher, mesDesde, mesHasta, dataF_months) {
  let cantidad = 0, ingreso = 0;
  if (!matcher) return { cantidad, ingreso };

  for (let m = mesDesde; m <= mesHasta; m++) {
    if (dataF_months[m] || !sixPlusSix || (!sixPlusSix.vol.length && !sixPlusSix.vta.length)) {
      (dataF || []).forEach(r => {
        if (r.mes === m && matcher(r)) {
          cantidad += r.cantidad;
          ingreso  += r.ingreso;
        }
      });
    } else {
      if (sixPlusSix.vol) {
        sixPlusSix.vol.forEach(r => {
          if (matcher(r)) cantidad += sumPptoCols(r, m, m);
        });
      }
      if (sixPlusSix.vta) {
        sixPlusSix.vta.forEach(r => {
          if (matcher(r)) ingreso += sumPptoCols(r, m, m);
        });
      }
    }
  }
  return { cantidad, ingreso };
}

function agregaPPTOExact(pptoData, matcher, mesDesde, mesHasta) {
  let cantidad = 0, ingreso = 0;
  if (!matcher) return { cantidad, ingreso };
  pptoData.vol.forEach(r => {
    if (matcher(r)) cantidad += sumPptoCols(r, mesDesde, mesHasta);
  });
  pptoData.vta.forEach(r => {
    if (matcher(r)) ingreso += sumPptoCols(r, mesDesde, mesHasta);
  });
  return { cantidad, ingreso };
}

// ------------------------------------------------------------------
// CONSTRUIRE FILAS DE LA TABLA SEGÚN EL ÁRBOL
// ------------------------------------------------------------------
export function buildComercialTable(rawData, mesSeleccionado, isYTD) {
  const { dataF, data2025, ppto, sixPlusSix } = rawData;

  const mesDesde = isYTD ? 1 : mesSeleccionado;
  const mesHasta = mesSeleccionado;

  let controlVolumenTotal = 0;
  let controlVentasTotal = 0;

  const dataF_months = {};
  for (let m = 1; m <= 12; m++) {
    let vol = 0, vta = 0;
    (dataF || []).forEach(r => {
      if (r.mes === m) {
        vol += r.cantidad || 0;
        vta += r.ingreso || 0;
      }
    });
    dataF_months[m] = (Math.abs(vol) > 0.01 || Math.abs(vta) > 0.01);
  }

  let isSixPlusSixActive = false;
  for (let m = mesDesde; m <= mesHasta; m++) {
    if (!dataF_months[m] && sixPlusSix && (sixPlusSix.vol.length || sixPlusSix.vta.length)) {
      isSixPlusSixActive = true;
    }
  }

  (dataF || []).forEach(row => {
    if (row.mes >= mesDesde && row.mes <= mesHasta) {
      // Suma bruta antes de clasificar
      controlVolumenTotal += Number(row.vol_2026 || row.volumen || row.cantidad || 0);
      controlVentasTotal += Number(row.ventas_2026 || row.ventas || row.ingreso || 0);
      
      let categoria = classifyRow(row);
      if (!categoria || categoria === 'OTROS NO CLASIFICADOS') {
        console.warn("⚠️ Fuga de Volumen Detectada - Fila sin clasificar:", {
          agrupacion: row['Agrupacion CEO'] || row.agrupacion,
          descripcion: row.descProd || row.descripcion,
          volumen: row.vol_2026 || row.volumen || row.cantidad,
          ventas: row.ventas_2026 || row.ventas || row.ingreso
        });
      }
    }
  });

  // DIAGNÓSTICO: Registrar productos sin emparejar en la consola de desarrollo (F12)
  const reportOrphans = (sourceName, rows, isPpto = false) => {
    const unmapped = [];
    if (isPpto) {
      // Para PPTO (volumen)
      (rows.vol || []).forEach(r => {
        const cat = classifyRow(r);
        if (!cat) unmapped.push({ descProd: r.descProd, agrupacion: r.agrupacion, cantidad: sumPptoCols(r, mesDesde, mesHasta) });
      });
    } else {
      (rows || []).forEach(r => {
        if (r.mes >= mesDesde && r.mes <= mesHasta) {
          const cat = classifyRow(r);
          if (!cat) unmapped.push(r);
        }
      });
    }

    if (unmapped.length > 0) {
      console.groupCollapsed(`⚠️ [resumenComercial] Productos de (${sourceName}) sin clasificar en el Periodo:`, `Meses ${mesDesde}-${mesHasta}`);
      const countMap = {};
      unmapped.forEach(r => {
        const agrp = String(r.agrupacion || 'S/D').trim();
        const key = `${r.descProd} (Agrupación: ${agrp})`;
        const vol = isPpto ? r.cantidad : r.cantidad;
        countMap[key] = (countMap[key] || 0) + vol;
      });
      console.table(Object.entries(countMap)
        .map(([prod, vol]) => ({ "Producto/Agrupación": prod, "Volumen": Math.round(vol) }))
        .sort((a, b) => b.Volumen - a.Volumen)
      );
      console.groupEnd();
    }
  };

  try {
    reportOrphans("Real 2026", dataF);
    reportOrphans("Histórico 2025", data2025);
    reportOrphans("Presupuesto PPTO", ppto, true);
  } catch (diagError) {
    console.warn("[comercialEngine] Error en diagnóstico de productos:", diagError);
  }

  const results = {};

  ARBOL_COMERCIAL.forEach(node => {
    results[node.id] = {
      node,
      volumen: { a25: 0, a26: 0, ppto: 0 },
      ventas: { a25: 0, a26: 0, ppto: 0 },
      precio: { a25: 0, a26: 0, ppto: 0 }
    };
    if (node.matcher) {
        const compositeMatcher = (r) => {
          if (!r) return false;
          const cat = classifyRow(r);
          return cat === node.id;
        };

        const r26 = agregaDataHybrid(dataF, sixPlusSix, compositeMatcher, mesDesde, mesHasta, dataF_months);
        const r25 = agregaData(data2025, compositeMatcher, mesDesde, mesHasta);
        const rPpto = agregaPPTOExact(ppto, compositeMatcher, mesDesde, mesHasta);

        results[node.id].volumen = { a25: r25.cantidad, a26: r26.cantidad, ppto: rPpto.cantidad };
        results[node.id].ventas = { a25: r25.ingreso, a26: r26.ingreso, ppto: rPpto.ingreso };
    }
  });

  // Roll-up de árbol recursivo a padres
  const getChildrenOf = (pid) => ARBOL_COMERCIAL.filter(n => n.parent === pid);
  
  const rollUp = (pid) => {
    const children = getChildrenOf(pid);
    let vol25 = 0, vol26 = 0, volPpto = 0;
    let vta25 = 0, vta26 = 0, vtaPpto = 0;
    
    children.forEach(c => {
      rollUp(c.id);
      vol25 += results[c.id].volumen.a25;
      vol26 += results[c.id].volumen.a26;
      volPpto += results[c.id].volumen.ppto;
      vta25 += results[c.id].ventas.a25;
      vta26 += results[c.id].ventas.a26;
      vtaPpto += results[c.id].ventas.ppto;
    });

    if (children.length > 0) {
      if (!ARBOL_COMERCIAL.find(n => n.id === pid).matcher) { // Override sólo si no tiene matcher propio directo
          results[pid].volumen.a25 = vol25;
          results[pid].volumen.a26 = vol26;
          results[pid].volumen.ppto = volPpto;
          results[pid].ventas.a25 = vta25;
          results[pid].ventas.a26 = vta26;
          results[pid].ventas.ppto = vtaPpto;
      }
    }
  };

  ARBOL_COMERCIAL.filter(n => !n.parent).forEach(root => rollUp(root.id));

  const tableRows = [];
  
  let grandTotalVol = { a25: 0, a26: 0, ppto: 0 };
  let grandTotalVta = { a25: 0, a26: 0, ppto: 0 };

  ARBOL_COMERCIAL.filter(n => !n.parent).forEach(root => {
      grandTotalVol.a25 += results[root.id].volumen.a25;
      grandTotalVol.a26 += results[root.id].volumen.a26;
      grandTotalVol.ppto += results[root.id].volumen.ppto;

      grandTotalVta.a25 += results[root.id].ventas.a25;
      grandTotalVta.a26 += results[root.id].ventas.a26;
      grandTotalVta.ppto += results[root.id].ventas.ppto;
  });

  ARBOL_COMERCIAL.forEach(node => {
      const data = results[node.id];
      const p25 = data.volumen.a25 !== 0 ? data.ventas.a25 / data.volumen.a25 : null;
      const p26 = data.volumen.a26 !== 0 ? data.ventas.a26 / data.volumen.a26 : null;
      const pPpto = data.volumen.ppto !== 0 ? data.ventas.ppto / data.volumen.ppto : null;

      data.precio = { a25: p25, a26: p26, ppto: pPpto };
      tableRows.push(data);
  });

  console.log("📊 Control Bruto Excel:", controlVolumenTotal);
  console.log("📊 Total Motor (grandTotalVol):", grandTotalVol.a26);
  console.log("Diferencia (Fuga real):", controlVolumenTotal - grandTotalVol.a26);

  const resultadoFinal = { tableRows, grandTotalVol, grandTotalVta, mes: mesSeleccionado, isYTD, isSixPlusSixActive };
  console.log("🔥 RESULTADO DEL MOTOR:", JSON.parse(JSON.stringify(resultadoFinal)));
  return resultadoFinal;
}

// ------------------------------------------------------------------
// FORMATEADORES EN ESPAÑOL Y MONEDA DOP
// ------------------------------------------------------------------
const CURRENCY_FORMATTER = new Intl.NumberFormat('es-DO', {
  style: 'currency',
  currency: 'DOP',
  minimumFractionDigits: 1,
  maximumFractionDigits: 1
});

const NUMBER_FORMATTER = new Intl.NumberFormat('en-US');

function fmtVol(n) {
  if (n == null || n === 0) return '-';
  return NUMBER_FORMATTER.format(Math.round(n));
}

function fmtPrecio(n) {
  if (n == null || n === 0) return '-';
  return n.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function fmtMdop(n) {
  if (n == null || n === 0) return '-';
  const millions = n / 1_000_000;
  return millions.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function fmtPct(n, showPlus = false) {
  if (n == null || typeof n !== 'number' || !isFinite(n)) return '-';
  const pct = Math.round(n * 100);
  const sign = (showPlus && pct > 0) ? '+' : '';
  return `${sign}${pct}%`;
}

// ------------------------------------------------------------------
// RENDERER DE LA TABLA (ID: resumen-comercial-tbody)
// ------------------------------------------------------------------
function fmtVolDiff(n) {
  if (n == null || n === 0) return '-';
  const sign = n > 0 ? '+' : '';
  return `${sign}${NUMBER_FORMATTER.format(Math.round(n))}`;
}

function fmtPrecioDiff(n) {
  if (n == null || n === 0) return '-';
  const sign = n > 0 ? '+' : '';
  return `${sign}${n.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}`;
}

function fmtMdopDiff(n) {
  if (n == null || n === 0) return '-';
  const sign = n > 0 ? '+' : '';
  const millions = n / 1_000_000;
  return `${sign}${millions.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}`;
}

function getMonthMetrics(table, isPrevYear = false) {
  const map = {};
  table.tableRows.forEach(row => {
    map[row.node.id] = {
      vol: isPrevYear ? row.volumen.a25 : row.volumen.a26,
      vta: isPrevYear ? row.ventas.a25 : row.ventas.a26,
      px: isPrevYear ? row.precio.a25 : row.precio.a26
    };
  });
  return map;
}

// ------------------------------------------------------------------
// RENDERER DE LA TABLA (ID: resumen-comercial-tbody)
// ------------------------------------------------------------------
// Global toggle for groups
window.toggleComercialGroup = function(nodeId) {
  window.comercialCollapsedState = window.comercialCollapsedState || {};
  window.comercialCollapsedState[nodeId] = !window.comercialCollapsedState[nodeId];
  if (typeof window.renderResumenComercial === 'function') {
    window.renderResumenComercial();
  }
};

export function renderResumenComercial(mesSeleccionado, isYTD, viewType = 'resumen') {
  if (!comercialRawData) {
    console.warn('[comercialEngine] No hay datos cargados aún.');
    return;
  }

  // Inject custom button styles once
  if (!document.getElementById('comercial-toggle-styles')) {
    const styleEl = document.createElement('style');
    styleEl.id = 'comercial-toggle-styles';
    styleEl.innerHTML = `
      .comercial-toggle-btn {
        width: 18px;
        height: 18px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        border-radius: 4px !important;
        color: #475569 !important;
        font-size: 11px !important;
        font-weight: bold !important;
        line-height: normal !important;
        margin-right: 8px !important;
        cursor: pointer !important;
        padding: 0 !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
        transition: all 0.15s ease !important;
        vertical-align: middle !important;
      }
      .comercial-toggle-btn:hover {
        background-color: #f1f5f9 !important;
        border-color: #94a3b8 !important;
        color: #1e293b !important;
        transform: scale(1.05);
      }
      .comercial-toggle-btn:active {
        transform: scale(0.95);
      }
    `;
    document.head.appendChild(styleEl);
  }

  // Ensure collapsed state map is defined
  window.comercialCollapsedState = window.comercialCollapsedState || {};

  // Visibility checker helper
  const isNodeVisible = (id) => {
    let curr = ARBOL_COMERCIAL.find(n => n.id === id);
    while (curr && curr.parent) {
      if (window.comercialCollapsedState[curr.parent]) {
        return false;
      }
      curr = ARBOL_COMERCIAL.find(n => n.id === curr.parent);
    }
    return true;
  };

  // Toggle button builder helper
  const getToggleButtonHtml = (nodeId) => {
    const hasChildren = ARBOL_COMERCIAL.some(n => n.parent === nodeId);
    if (!hasChildren) {
      return `<span style="width: 26px; display: inline-block; shrink: 0; flex-shrink: 0;"></span>`;
    }
    const isCollapsed = window.comercialCollapsedState[nodeId] || false;
    const sign = isCollapsed ? '+' : '−';
    return `
      <button class="comercial-toggle-btn" data-id="${nodeId}" onclick="event.stopPropagation(); window.toggleComercialGroup('${nodeId}');">
        ${sign}
      </button>
    `;
  };

  // Buscar IDs de tabla en index.html
  const tbody = document.getElementById('resumen-comercial-tbody') || document.getElementById('tbody-resumen-comercial');
  const thead = document.getElementById('resumen-comercial-thead') || document.getElementById('thead-resumen-comercial');
  if (!tbody || !thead) {
    console.warn('[comercialEngine] No se encontraron elementos tbody o thead para la tabla comercial.');
    return;
  }

  const MESES = ['', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const mIndex = mesSeleccionado || 3;
  const mesName = MESES[mIndex] || 'Mes';

  const parentDepth = (id) => {
      let d = 0;
      let curr = ARBOL_COMERCIAL.find(n => n.id === id);
      while(curr && curr.parent) {
          d++;
          curr = ARBOL_COMERCIAL.find(n => n.id === curr.parent);
      }
      return d;
  };

  const formatPrcClr = (val) => {
      if (val === null || typeof val !== 'number' || !isFinite(val)) return { text: '-', color: 'inherit' };
      const text = fmtPct(val, true);
      return { text, color: val >= 0 ? '#16a34a' : '#dc2626' }; 
  };

  // 1. VISTA: MO MONTH OVER MONTH (MoM)
  if (viewType === 'mom') {
    let prevMonthName = '';
    let prevMetrics = {};
    let isPrev6x6 = false;
    if (mIndex === 1) {
      prevMonthName = 'Dic 25';
      const d25Table = buildComercialTable(comercialRawData, 12, false);
      prevMetrics = getMonthMetrics(d25Table, true);
    } else {
      prevMonthName = MESES[mIndex - 1] + ' 26';
      const prevTable = buildComercialTable(comercialRawData, mIndex - 1, false);
      prevMetrics = getMonthMetrics(prevTable, false);
      isPrev6x6 = prevTable.isSixPlusSixActive;
    }
    const currMonthName = MESES[mIndex] + ' 26';
    const currTable = buildComercialTable(comercialRawData, mIndex, false);
    const currMetrics = getMonthMetrics(currTable, false);

    const prevLbColor = isPrev6x6 ? '#f59e0b' : 'white';
    const currLbColor = currTable.isSixPlusSixActive ? '#f59e0b' : 'var(--primary)';
    const prevMonthLabel = isPrev6x6 ? `<span title="Incluye datos proyectados de 6+6">${prevMonthName.toUpperCase()} *</span>` : prevMonthName.toUpperCase();
    const currMonthLabel = currTable.isSixPlusSixActive ? `<span title="Incluye datos proyectados de 6+6">${currMonthName.toUpperCase()} *</span>` : currMonthName.toUpperCase();

    // RENDER THEAD MoM
    thead.innerHTML = `
      <colgroup>
        <col style="width: 22%;">
        <!-- Volumen = 24% -->
        <col style="width: 8%;">
        <col style="width: 8%;">
        <col style="width: 8%;">
        <!-- Precio = 24% -->
        <col style="width: 8%;">
        <col style="width: 8%;">
        <col style="width: 8%;">
        <!-- Ventas = 30% -->
        <col style="width: 10%;">
        <col style="width: 10%;">
        <col style="width: 10%;">
      </colgroup>
      <tr>
        <th rowspan="2" style="background:var(--sidebar); color:white; border-right: 2px solid rgba(255,255,255, 0.2); vertical-align:middle; text-transform:uppercase; font-size: 0.95rem; letter-spacing: 0.05em; padding: 12px 24px;">Categoría</th>
        <th colspan="3" style="text-align:center; background:#1e293b; color:white; border-bottom: 2px solid #38bdf8; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">Volumen (Unidades)</th>
        <th colspan="3" style="text-align:center; background:var(--sidebar); color:white; border-bottom: 2px solid #1e40af; border-left: 3px solid #475569; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">Precio (DOP)</th>
        <th colspan="3" style="text-align:center; background:#1e293b; color:white; border-bottom: 2px solid #38bdf8; border-left: 3px solid #475569; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">Ventas Netas (mDOP)</th>
      </tr>
      <tr>
        <th style="text-align:right; background:#334155; color:${prevLbColor}; font-size: 0.95rem;">${prevMonthLabel}</th>
        <th style="text-align:right; background:#334155; color:${currLbColor}; font-weight: 800; font-size: 0.95rem;">${currMonthLabel}</th>
        <th style="text-align:right; background:#1e293b; color:white; font-size: 0.95rem; border-right: 3px solid #475569;">MoM %</th>
        
        <th style="text-align:right; background:#1e293b; color:${prevLbColor}; font-size: 0.85rem;">${prevMonthLabel}</th>
        <th style="text-align:right; background:#1e293b; color:${currLbColor}; font-weight: 800; font-size: 0.85rem;">${currMonthLabel}</th>
        <th style="text-align:right; background:#334155; color:white; font-size: 0.85rem; border-right: 3px solid #475569;">MoM %</th>
        
        <th style="text-align:right; background:#334155; color:${prevLbColor}; font-size: 0.85rem;">${prevMonthLabel}</th>
        <th style="text-align:right; background:#334155; color:${currLbColor}; font-weight: 800; font-size: 0.85rem;">${currMonthLabel}</th>
        <th style="text-align:right; background:#1e293b; color:white; font-size: 0.85rem;">MoM %</th>
      </tr>
    `;

    // RENDER TBODY MoM
    let html = '';
    ARBOL_COMERCIAL.forEach(node => {
        if (!isNodeVisible(node.id)) return;

        const pData = prevMetrics[node.id] || { vol: 0, vta: 0, px: null };
        const cData = currMetrics[node.id] || { vol: 0, vta: 0, px: null };
        
        const varVol = pData.vol !== 0 ? (cData.vol - pData.vol) / Math.abs(pData.vol) : null;
        const varPx = (pData.px != null && pData.px !== 0) ? (cData.px - pData.px) / Math.abs(pData.px) : null;
        const varVta = pData.vta !== 0 ? (cData.vta - pData.vta) / Math.abs(pData.vta) : null;

        const vVol = formatPrcClr(varVol);
        const vPx = formatPrcClr(varPx);
        const vVta = formatPrcClr(varVta);

        let rowClass = '';
        let tdFirstStyle = '';
        const depth = parentDepth(node.id);

        const toggleBtn = getToggleButtonHtml(node.id);
        const labelContent = `<span style="display: inline-flex; align-items: center; vertical-align: middle;">${toggleBtn}<span>${node.label}</span></span>`;

        if (node.type === 'main') {
            rowClass = 'row-total';
            tdFirstStyle = `background: var(--sidebar) !important; color: white !important; font-weight: 800 !important; text-transform: uppercase; letter-spacing: 0.5px; border-right: 2px solid rgba(255,255,255,0.1); font-size: 1.15rem !important;`;
        } else if (node.type === 'sub') {
            rowClass = 'row-category';
            tdFirstStyle = `padding-left: 24px !important; font-weight: 600; border-right: 1px solid var(--border); text-transform: uppercase; font-size: 1.08rem !important;`;
        } else {
            let padding = 24 + (depth * 14); 
            tdFirstStyle = `padding-left: ${padding}px !important; color: var(--text-secondary) !important; border-right: 1px solid var(--border); font-size: 1.02rem !important;`;
        }

        html += `
        <tr class="${rowClass}">
            <td style="${tdFirstStyle}">${labelContent}</td>
            
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem;">${fmtVol(pData.vol)}</td>
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; font-weight: 700;">${fmtVol(cData.vol)}</td>
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; color:${vVol.color}; font-weight:bold; border-right: 3px solid #e2e8f0;">${vVol.text}</td>
            
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem;">${fmtPrecio(pData.px)}</td>
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; font-weight: 700;">${fmtPrecio(cData.px)}</td>
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; color:${vPx.color}; font-weight:bold; border-right: 3px solid #e2e8f0;">${vPx.text}</td>
            
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem;">${fmtMdop(pData.vta)}</td>
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; font-weight: 700; color:var(--primary); background:rgba(0,0,0,0.02);">${fmtMdop(cData.vta)}</td>
            <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; color:${vVta.color}; font-weight:bold;">${vVta.text}</td>
        </tr>
        `;
    });

    // TOTAL GENERAL MoM
    const totPrevVol = Object.values(prevMetrics).reduce((acc, x) => acc + (x.vol || 0), 0) / 2; // dividing by 2 due to category vs item roll-ups or root total. In our case, let's get standard total of Root total
    const rootNodes = ARBOL_COMERCIAL.filter(n => !n.parent);
    let totPVol = 0, totCVol = 0, totPVta = 0, totCVta = 0;
    rootNodes.forEach(node => {
       totPVol += prevMetrics[node.id]?.vol || 0;
       totCVol += currMetrics[node.id]?.vol || 0;
       totPVta += prevMetrics[node.id]?.vta || 0;
       totCVta += currMetrics[node.id]?.vta || 0;
    });

    const totVarVol = totPVol !== 0 ? (totCVol - totPVol) / Math.abs(totPVol) : null;
    const totVarVta = totPVta !== 0 ? (totCVta - totPVta) / Math.abs(totPVta) : null;
    const vTotVol = formatPrcClr(totVarVol);
    const vTotVta = formatPrcClr(totVarVta);

    html += `
        <tr style="background:var(--sidebar);">
            <td style="background:var(--sidebar) !important; color:white !important; font-weight:bold !important; text-transform:uppercase; border-right: 2px solid rgba(255,255,255, 0.2); font-size: 1.15rem !important;">TOTAL</td>
            <td style="text-align:right; background:var(--sidebar) !important; color:white !important; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtVol(totPVol)}</td>
            <td style="text-align:right; background:var(--sidebar) !important; color:white !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtVol(totCVol)}</td>
            <td style="text-align:right; background:var(--sidebar) !important; color:${vTotVol.color} !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); border-right: 3px solid #475569; font-family:var(--font-mono); font-size:1.15rem;">${vTotVol.text}</td>
            
            <td colspan="3" style="text-align:center; background:var(--sidebar) !important; border-top:2px solid var(--sidebar-accent); border-right: 3px solid #475569;"></td>
            
            <td style="text-align:right; background:var(--sidebar) !important; color:white !important; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtMdop(totPVta)}</td>
            <td style="text-align:right; background:var(--sidebar) !important; color:#38bdf8 !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtMdop(totCVta)}</td>
            <td style="text-align:right; background:var(--sidebar) !important; color:${vTotVta.color} !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${vTotVta.text}</td>
        </tr>
    `;
    tbody.innerHTML = html;
    return;
  }

  // 2. VISTA: ANÁLISIS DE VARIACIÓN (VARIACION)
  if (viewType === 'variacion') {
    // We only need the single table that matches the selected period state (isYTD)
    const compTable = buildComercialTable(comercialRawData, mesSeleccionado, isYTD);

    // Left Header Group & Right Header Group titles depend on whether we are in YTD mode or Monthly mode
    let leftHeaderLabel = isYTD ? 'YTD ACTUAL VS 2025 (YoY)' : 'MES ACTUAL VS 2025 (YoY)';
    let rightHeaderLabel = isYTD ? 'YTD ACTUAL VS PPTO' : 'MES ACTUAL VS PPTO';

    if (compTable.isSixPlusSixActive) {
      leftHeaderLabel += ' * (6+6)';
      rightHeaderLabel += ' * (6+6)';
    }

    // RENDER THEAD VARIACION
    thead.innerHTML = `
      <colgroup>
        <col style="width: 22%;">
        <!-- Left Group (YoY) = 39% -->
        <col style="width: 13%;">
        <col style="width: 13%;">
        <col style="width: 13%;">
        <!-- Right Group (PPTO) = 39% -->
        <col style="width: 13%;">
        <col style="width: 13%;">
        <col style="width: 13%;">
      </colgroup>
      <tr>
        <th rowspan="2" style="background:var(--sidebar); color:white; border-right: 2px solid rgba(255,255,255, 0.2); vertical-align:middle; text-transform:uppercase; font-size: 0.95rem; letter-spacing: 0.05em; padding: 12px 24px;">Categoría</th>
        <th colspan="3" style="text-align:center; background:#1e293b; color:white; border-bottom: 2px solid #38bdf8; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em; border-right: 3px solid #475569;">${leftHeaderLabel}</th>
        <th colspan="3" style="text-align:center; background:var(--sidebar); color:white; border-bottom: 2px solid #10b981; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">${rightHeaderLabel}</th>
      </tr>
      <tr>
        <th style="text-align:right; background:#334155; color:white; font-size: 0.85rem; padding: 6px 12px;">Volumen</th>
        <th style="text-align:right; background:#334155; color:white; font-size: 0.85rem; padding: 6px 12px;">Precio</th>
        <th style="text-align:right; background:#334155; color:white; font-size: 0.85rem; padding: 6px 12px; border-right: 3px solid #475569;">Total</th>
        
        <th style="text-align:right; background:#1e293b; color:white; font-size: 0.85rem; padding: 6px 12px;">Volumen</th>
        <th style="text-align:right; background:#1e293b; color:white; font-size: 0.85rem; padding: 6px 12px;">Precio</th>
        <th style="text-align:right; background:#1e293b; color:white; font-size: 0.85rem; padding: 6px 12px;">Total</th>
      </tr>
    `;

    // RENDER TBODY VARIACION
    let html = '';
    
    const getCellHtml = (diff, pct, isPrice = false, isVta = false, hasBorder = false) => {
        if (diff === null) {
            const brStyle = hasBorder ? 'border-right: 3px solid #cbd5e1;' : '';
            return `<td style="text-align:right; font-family:var(--font-mono); color:var(--text-secondary); font-size:1.05rem; border-bottom: 1px solid var(--border); ${brStyle}">-</td>`;
        }
        const clr = diff >= 0 ? '#16a34a' : '#dc2626';
        const fDiff = isPrice ? fmtPrecioDiff(diff) : (isVta ? fmtMdopDiff(diff) : fmtVolDiff(diff));
        const brStyle = hasBorder ? 'border-right: 3px solid #cbd5e1;' : '';
        return `
          <td style="text-align:right; padding: 8px 12px; border-bottom: 1px solid var(--border); ${brStyle}">
            <div style="font-family:var(--font-mono); font-size:1.05rem; font-weight:600; color:${clr}">${fDiff}</div>
            <div style="font-family:var(--font-mono); font-size:0.95rem; color:${clr}; font-weight:bold;">${fmtPct(pct, true)}</div>
          </td>
        `;
    };

    compTable.tableRows.forEach(row => {
        // Find corresponding node
        const node = ARBOL_COMERCIAL.find(n => n.id === row.node.id);
        if (!node) return;
        if (!isNodeVisible(node.id)) return;

        const vol26 = row.volumen.a26 || 0;
        const vol25 = row.volumen.a25 || 0;
        const volPpto = row.volumen.ppto || 0;
        const px26 = row.precio.a26 || 0;
        const px25 = row.precio.a25 || 0;
        const pxPpto = row.precio.ppto || 0;
        const vta26 = row.ventas.a26 || 0;
        const vta25 = row.ventas.a25 || 0;
        const vtaPpto = row.ventas.ppto || 0;

        // 1. Left Group: Actual vs 2025 (YoY) Calculations (Money Impacts)
        // Volumen impact: ((Volumen 2026 - Volumen 2025) * Precio 2026)
        const dVol25Money = (vol26 - vol25) * px26;
        // Precio impact: Volumen 2025 * (Precio 2026 - Precio 2025)
        const dPx25Money = vol25 * (px26 - px25);
        // Total impact: Ventas 2026 - Ventas 2025
        const dVta25 = vta26 - vta25;

        const pVol25 = vta25 !== 0 ? dVol25Money / Math.abs(vta25) : null;
        const pPx25 = vta25 !== 0 ? dPx25Money / Math.abs(vta25) : null;
        const pVta25 = vta25 !== 0 ? dVta25 / Math.abs(vta25) : null;

        // 2. Right Group: Actual vs PPTO Calculations (Money Impacts)
        // Volumen impact: ((Volumen 2026 - Volumen PPTO) * Precio 2026)
        const dVolPptoMoney = (vol26 - volPpto) * px26;
        // Precio impact: Volumen PPTO * (Precio 2026 - Precio PPTO)
        const dPxPptoMoney = volPpto * (px26 - pxPpto);
        // Total impact: Ventas 2026 - Ventas PPTO
        const dVtaPpto = vta26 - vtaPpto;

        const pVolPpto = vtaPpto !== 0 ? dVolPptoMoney / Math.abs(vtaPpto) : null;
        const pPxPpto = vtaPpto !== 0 ? dPxPptoMoney / Math.abs(vtaPpto) : null;
        const pVtaPpto = vtaPpto !== 0 ? dVtaPpto / Math.abs(vtaPpto) : null;

        let rowClass = '';
        let tdFirstStyle = '';
        const depth = parentDepth(node.id);

        const toggleBtn = getToggleButtonHtml(node.id);
        const labelContent = `<span style="display: inline-flex; align-items: center; vertical-align: middle;">${toggleBtn}<span>${node.label}</span></span>`;

        if (node.type === 'main') {
            rowClass = 'row-total';
            tdFirstStyle = `background: var(--sidebar) !important; color: white !important; font-weight: 800 !important; text-transform: uppercase; letter-spacing: 0.5px; border-right: 2px solid rgba(255,255,255,0.1); font-size: 1.15rem !important;`;
        } else if (node.type === 'sub') {
            rowClass = 'row-category';
            tdFirstStyle = `padding-left: 24px !important; font-weight: 600; border-right: 1px solid var(--border); text-transform: uppercase; font-size: 1.08rem !important;`;
        } else {
            let padding = 24 + (depth * 14); 
            tdFirstStyle = `padding-left: ${padding}px !important; color: var(--text-secondary) !important; border-right: 1px solid var(--border); font-size: 1.02rem !important;`;
        }

        html += `
        <tr class="${rowClass}">
            <td style="${tdFirstStyle}">${labelContent}</td>
            ${getCellHtml(dVol25Money, pVol25, false, true)}
            ${getCellHtml(dPx25Money, pPx25, false, true)}
            ${getCellHtml(dVta25, pVta25, false, true, true)}
            ${getCellHtml(dVolPptoMoney, pVolPpto, false, true)}
            ${getCellHtml(dPxPptoMoney, pPxPpto, false, true)}
            ${getCellHtml(dVtaPpto, pVtaPpto, false, true)}
        </tr>
        `;
    });

    // TOTALES GENERALES VARIACION (Money Impacts based on overall weighted average prices)
    const totVol26 = compTable.grandTotalVol.a26 || 0;
    const totVol25 = compTable.grandTotalVol.a25 || 0;
    const totVolPpto = compTable.grandTotalVol.ppto || 0;
    const totVta26 = compTable.grandTotalVta.a26 || 0;
    const totVta25 = compTable.grandTotalVta.a25 || 0;
    const totVtaPpto = compTable.grandTotalVta.ppto || 0;

    const totPx26 = totVol26 !== 0 ? totVta26 / totVol26 : 0;
    const totPx25 = totVol25 !== 0 ? totVta25 / totVol25 : 0;
    const totPxPpto = totVolPpto !== 0 ? totVtaPpto / totVolPpto : 0;

    // YTD or Month Actual vs 2025 (YoY) Totals
    const totDVol25Money = (totVol26 - totVol25) * totPx26;
    const totDPx25Money = totVol25 * (totPx26 - totPx25);
    const totDVta25 = totVta26 - totVta25;

    const totPVol25 = totVta25 !== 0 ? totDVol25Money / Math.abs(totVta25) : null;
    const totPPx25 = totVta25 !== 0 ? totDPx25Money / Math.abs(totVta25) : null;
    const totPVta25 = totVta25 !== 0 ? totDVta25 / Math.abs(totVta25) : null;

    // YTD or Month Actual vs PPTO Totals
    const totDVolPptoMoney = (totVol26 - totVolPpto) * totPx26;
    const totDPxPptoMoney = totVolPpto * (totPx26 - totPxPpto);
    const totDVtaPpto = totVta26 - totVtaPpto;

    const totPVolPpto = totVtaPpto !== 0 ? totDVolPptoMoney / Math.abs(totVtaPpto) : null;
    const totPPxPpto = totVtaPpto !== 0 ? totDPxPptoMoney / Math.abs(totVtaPpto) : null;
    const totPVtaPpto = totVtaPpto !== 0 ? totDVtaPpto / Math.abs(totVtaPpto) : null;

    const getTotalCellHtml = (diff, pct, isVta = false, hasBorder = false) => {
        const clr = diff >= 0 ? '#38bdf8' : '#ef4444';
        const fDiff = isVta ? fmtMdopDiff(diff) : fmtVolDiff(diff);
        const brStyle = hasBorder ? 'border-right: 3px solid #475569;' : '';
        return `
          <td style="text-align:right; background:var(--sidebar) !important; border-top:2px solid var(--sidebar-accent); ${brStyle}">
            <div style="font-family:var(--font-mono); font-size:1.15rem; font-weight:bold; color:${clr}">${fDiff}</div>
            <div style="font-family:var(--font-mono); font-size:1.0rem; color:${clr}; font-weight:bold;">${fmtPct(pct, true)}</div>
          </td>
        `;
    };

    html += `
        <tr style="background:var(--sidebar);">
            <td style="background:var(--sidebar) !important; color:white !important; font-weight:bold !important; text-transform:uppercase; border-right: 2px solid rgba(255,255,255, 0.2); font-size: 1.15rem !important;">TOTAL</td>
            ${getTotalCellHtml(totDVol25Money, totPVol25, true)}
            ${getTotalCellHtml(totDPx25Money, totPPx25, true)}
            ${getTotalCellHtml(totDVta25, totPVta25, true, true)}
            
            ${getTotalCellHtml(totDVolPptoMoney, totPVolPpto, true)}
            ${getTotalCellHtml(totDPxPptoMoney, totPPxPpto, true)}
            ${getTotalCellHtml(totDVtaPpto, totPVtaPpto, true)}
        </tr>
    `;

    tbody.innerHTML = html;
    return;
  }

  // 3. BASELINE DETAILED VIEW: RESUMEN DE VENTAS (DEFAULT)
  const table = buildComercialTable(comercialRawData, mesSeleccionado, isYTD);
  
  const m26Color = table.isSixPlusSixActive ? '#f59e0b' : 'var(--primary)'; // Amber color when using 6+6 projected data
  let m26LabelText = (isYTD ? `YTD-${mesName} 26` : `${mesName}-26`).toUpperCase();
  if (table.isSixPlusSixActive) m26LabelText += ' *';

  const mesLabel25 = (isYTD ? `YTD-${mesName} 25` : `${mesName}-25`).toUpperCase();
  const mesLabel26 = `<span title="Incluye datos proyectados de 6+6">${m26LabelText}</span>`;

  // RENDER THEAD
  thead.innerHTML = `
    <colgroup>
      <col style="width: 22%;">
      <!-- Volumen (Unidades) = 19.5% -->
      <col style="width: 6.5%;">
      <col style="width: 6.5%;">
      <col style="width: 6.5%;">
      <!-- Precio (DOP) = 19.5% -->
      <col style="width: 6.5%;">
      <col style="width: 6.5%;">
      <col style="width: 6.5%;">
      <!-- Ventas Netas (mDOP) = 19.5% -->
      <col style="width: 6.5%;">
      <col style="width: 6.5%;">
      <col style="width: 6.5%;">
      <!-- Variación % = 19.5% -->
      <col style="width: 9.75%;">
      <col style="width: 9.75%;">
    </colgroup>
    <tr>
      <th rowspan="2" style="background:var(--sidebar); color:white; border-right: 2px solid rgba(255,255,255, 0.2); vertical-align:middle; text-transform:uppercase; font-size: 0.95rem; letter-spacing: 0.05em; padding: 12px 24px;">Categoría</th>
      <th colspan="3" style="text-align:center; background:#1e293b; color:white; border-bottom: 2px solid #38bdf8; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">Volumen (Unidades)</th>
      <th colspan="3" style="text-align:center; background:var(--sidebar); color:white; border-bottom: 2px solid #1e40af; border-left: 3px solid #475569; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">Precio (DOP)</th>
      <th colspan="3" style="text-align:center; background:#1e293b; color:white; border-bottom: 2px solid #38bdf8; border-left: 3px solid #475569; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">Ventas Netas (mDOP)</th>
      <th colspan="2" style="text-align:center; background:var(--sidebar); color:white; border-bottom: 2px solid #10b981; border-left: 3px solid #475569; padding: 10px; font-size: 0.95rem; text-transform: uppercase; letter-spacing: 0.05em;">Variación %</th>
    </tr>
    <tr>
      <th style="text-align:right; background:#334155; color:white; font-size: 0.85rem;">${mesLabel25}</th>
      <th style="text-align:right; background:#334155; color:${m26Color}; font-weight: 800; font-size: 0.85rem;">${mesLabel26}</th>
      <th style="text-align:right; background:#1e3a8a; color:white; font-size: 0.85rem; border-right: 3px solid #475569;">PPTO</th>
      
      <th style="text-align:right; background:#1e293b; color:white; font-size: 0.85rem;">${mesLabel25}</th>
      <th style="text-align:right; background:#1e293b; color:${m26Color}; font-weight: 800; font-size: 0.85rem;">${mesLabel26}</th>
      <th style="text-align:right; background:#1e3a8a; color:white; font-size: 0.85rem; border-right: 3px solid #475569;">PPTO</th>
      
      <th style="text-align:right; background:#334155; color:white; font-size: 0.85rem;">${mesLabel25}</th>
      <th style="text-align:right; background:#334155; color:${m26Color}; font-weight: 800; font-size: 0.85rem;">${mesLabel26}</th>
      <th style="text-align:right; background:#1e3a8a; color:white; font-size: 0.85rem; border-right: 3px solid #475569;">PPTO</th>
      
      <th style="text-align:right; background:#1e293b; color:white; font-size: 0.85rem;">vs 2025</th>
      <th style="text-align:right; background:#1e293b; color:white; font-size: 0.85rem;">Vs PPTO</th>
    </tr>
  `;

  // RENDER TBODY
  let html = '';
  
  table.tableRows.forEach(row => {
      if (!isNodeVisible(row.node.id)) return;

      const var25 = row.ventas.a25 !== 0 ? (row.ventas.a26 - row.ventas.a25) / Math.abs(row.ventas.a25) : null;
      const varPpto = row.ventas.ppto !== 0 ? (row.ventas.a26 - row.ventas.ppto) / Math.abs(row.ventas.ppto) : null;

      let rowClass = '';
      let tdFirstStyle = '';
      
      const depth = parentDepth(row.node.id);

      const toggleBtn = getToggleButtonHtml(row.node.id);
      const labelContent = `<span style="display: inline-flex; align-items: center; vertical-align: middle;">${toggleBtn}<span>${row.node.label}</span></span>`;

      if (row.node.type === 'main') {
          rowClass = 'row-total';
          tdFirstStyle = `background: var(--sidebar) !important; color: white !important; font-weight: 800 !important; text-transform: uppercase; letter-spacing: 0.5px; border-right: 2px solid rgba(255,255,255,0.1); font-size: 1.15rem !important;`;
      } else if (row.node.type === 'sub') {
          rowClass = 'row-category';
          tdFirstStyle = `padding-left: 24px !important; font-weight: 600; border-right: 1px solid var(--border); text-transform: uppercase; font-size: 1.08rem !important;`;
      } else {
          let padding = 24 + (depth * 14); 
          tdFirstStyle = `padding-left: ${padding}px !important; color: var(--text-secondary) !important; border-right: 1px solid var(--border); font-size: 1.02rem !important;`;
      }

      const v25 = formatPrcClr(var25);
      const vPpto = formatPrcClr(varPpto);

      const pptoBg = 'rgba(30, 64, 175, 0.04)';

      html += `
      <tr class="${rowClass}">
          <td style="${tdFirstStyle}">${labelContent}</td>
          
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem;">${fmtVol(row.volumen.a25)}</td>
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; font-weight: 700;">${fmtVol(row.volumen.a26)}</td>
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; background:${pptoBg}; color:#1e40af; border-right: 3px solid #e2e8f0;">${fmtVol(row.volumen.ppto)}</td>
          
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem;">${fmtPrecio(row.precio.a25)}</td>
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; font-weight: 700;">${fmtPrecio(row.precio.a26)}</td>
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; background:${pptoBg}; color:#1e40af; border-right: 3px solid #e2e8f0;">${fmtPrecio(row.precio.ppto)}</td>
          
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem;">${fmtMdop(row.ventas.a25)}</td>
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; font-weight: 700; color:var(--primary); background:rgba(0,0,0,0.02);">${fmtMdop(row.ventas.a26)}</td>
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; background:${pptoBg}; color:#1e40af; border-right: 3px solid #e2e8f0;">${fmtMdop(row.ventas.ppto)}</td>
          
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; color:${v25.color}; font-weight:bold;">${v25.text}</td>
          <td style="text-align:right; font-family:var(--font-mono); font-size:1.05rem; color:${vPpto.color}; font-weight:bold;">${vPpto.text}</td>
      </tr>
      `;
  });

  const totVar25 = table.grandTotalVta.a25 !== 0 ? (table.grandTotalVta.a26 - table.grandTotalVta.a25) / Math.abs(table.grandTotalVta.a25) : null;
  const totVarPpto = table.grandTotalVta.ppto !== 0 ? (table.grandTotalVta.a26 - table.grandTotalVta.ppto) / Math.abs(table.grandTotalVta.ppto) : null;
  const vTot25 = formatPrcClr(totVar25);
  const vTotPpto = formatPrcClr(totVarPpto);

  // TOTALES GENERALES
  html += `
      <tr style="background:var(--sidebar);">
          <td style="background:var(--sidebar) !important; color:white !important; font-weight:bold !important; text-transform:uppercase; border-right: 2px solid rgba(255,255,255, 0.2); font-size: 1.15rem !important;">TOTAL</td>
          <td style="text-align:right; background:var(--sidebar) !important; color:white !important; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtVol(table.grandTotalVol.a25)}</td>
          <td style="text-align:right; background:var(--sidebar) !important; color:white !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtVol(table.grandTotalVol.a26)}</td>
          <td style="text-align:right; background:#1e3a8a !important; color:white !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); border-right: 3px solid #475569; font-family:var(--font-mono); font-size:1.15rem;">${fmtVol(table.grandTotalVol.ppto)}</td>
          
          <td colspan="3" style="text-align:center; background:var(--sidebar) !important; border-top:2px solid var(--sidebar-accent); border-right: 3px solid #475569;"></td>
          
          <td style="text-align:right; background:var(--sidebar) !important; color:white !important; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtMdop(table.grandTotalVta.a25)}</td>
          <td style="text-align:right; background:var(--sidebar) !important; color:#38bdf8 !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${fmtMdop(table.grandTotalVta.a26)}</td>
          <td style="text-align:right; background:#1e3a8a !important; color:white !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); border-right: 3px solid #475569; font-family:var(--font-mono); font-size:1.15rem;">${fmtMdop(table.grandTotalVta.ppto)}</td>
          
          <td style="text-align:right; background:var(--sidebar) !important; color:${vTot25.color} !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${vTot25.text}</td>
          <td style="text-align:right; background:var(--sidebar) !important; color:${vTotPpto.color} !important; font-weight:bold; border-top:2px solid var(--sidebar-accent); font-family:var(--font-mono); font-size:1.15rem;">${vTotPpto.text}</td>
      </tr>
  `;

  tbody.innerHTML = html;
}

// ------------------------------------------------------------------
// RENDERER DE LA TABLA P&G HORIZONTAL (ID: pg-horizontal-tbody)
// ------------------------------------------------------------------
export async function renderPgHorizontal() {
  // fallback if missing
  if (!comercialRawData || !comercialRawData.pgHorizontal || comercialRawData.pgHorizontal.length === 0) {
    try {
      const db = await openDB();
      const cached = await dbGet(db, 'COMERCIAL_KEY');
      if (cached && cached.data && cached.data.pgHorizontal && cached.data.pgHorizontal.length > 0) {
        if (!comercialRawData) comercialRawData = cached.data;
        else comercialRawData.pgHorizontal = cached.data.pgHorizontal;
        console.log('[comercialEngine] pgHorizontal restaurado de IndexedDB en render.');
      }
    } catch (e) {}
  }

  if (!comercialRawData || !comercialRawData.pgHorizontal || comercialRawData.pgHorizontal.length === 0) {
    console.warn('[comercialEngine] No hay datos de P&G Horizontal cargados aún.');
    const tbody = document.getElementById('pg-horizontal-tbody');
    const badge = document.getElementById('integrityBadge');
    if (badge) badge.style.display = 'none';
    if (tbody) tbody.innerHTML = `<tr><td colspan="19" style="text-align:center; padding:40px; color:var(--text-secondary); font-style:italic;">No se encontraron datos de P&G Horizontal en el archivo.</td></tr>`;
    return;
  }

  const tbody = document.getElementById('pg-horizontal-tbody');
  if (!tbody) return;

  const data = comercialRawData.pgHorizontal;
  if (data.length === 0) {
    const badge = document.getElementById('integrityBadge');
    if (badge) badge.style.display = 'none';
    tbody.innerHTML = `<tr><td colspan="19" style="text-align: center; padding: 40px; color: var(--text-secondary); font-style: italic;">No se encontraron datos de P&G Horizontal en el archivo.</td></tr>`;
    return;
  }

  const dropdown = document.getElementById('pg-dropdown-scenario');
  const scenarioLabel = dropdown ? dropdown.value : 'REAL AÑO ANTERIOR';
  
  let scenarioColIdx = -1;
  let occurrenceCount = 0;
  let targetOccurrence = scenarioLabel.endsWith(' 2') ? 2 : 1;
  let cleanLabel = scenarioLabel.replace(' 1', '').replace(' 2', '').replace(/\s+/g, ' ').trim().toUpperCase();
  
  // Find column index
  for (let i = 0; i < Math.min(20, data.length); i++) {
      const row = data[i];
      if (!row) continue;
      for (let j = 0; j < row.length; j++) {
          if (row[j] && typeof row[j] === 'string') {
              const cellVal = row[j].toUpperCase();
              if (cellVal.includes(cleanLabel) || cleanLabel.includes(cellVal)) {
                  occurrenceCount++;
                  if (occurrenceCount === targetOccurrence) {
                      scenarioColIdx = j;
                      break;
                  }
              }
          }
      }
      if (scenarioColIdx !== -1) break;
  }

  console.log("-> [PG Horizontal] cleanLabel:", cleanLabel, "Found at colIdx:", scenarioColIdx);

  const brandsData = {};
  const conceptsKeywords = {
     'UNIDADES TOTALES': 'unidades',
     'VENTAS NETAS': 'ventas',
     'COSTO DE VENTAS': 'costo',
     'UTILIDAD BRUTA': 'utilidad_bruta',
     'GASTOS LOGISTICOS EXTERNOS': 'logistica',
     'GASTOS LOGÍSTICOS EXTERNOS': 'logistica',
     'UTILIDAD POST LOGISTICOS': 'utilidad_post',
     'UTILIDAD POST LOGÍSTICOS': 'utilidad_post',
     'APOYO COMERCIAL': 'apoyo_comercial',
     'CONTRIBUCION DIRECTA': 'contribucion',
     'CONTRIBUCIÓN DIRECTA': 'contribucion'
  };

  const ignoreKeywords = [
     'HECTOLITROS', 'APOYO A MARCAS', 'REINTEGRO', 'INVESTIGACION',
     'APOYO VENTAS', 'GASTOS DE OPERACIÓN', 'GASTOS SALARIALES', 
     'GASTOS CENTRALES', 'UTILIDAD OPERATIVA', 'OTROS INGRESOS', 
     'DEPRECIACION', 'PERDIDA DE VALOR', 'DIFERENCIA EN CAMBIO', 
     'IMPUESTOS', 'UTILIDAD NETA', 'PRECIO NETO', 'COSTOS X UNIDAD', 
     'MARGEN X UNIDAD', 'MARGEN BRUTO', 'FC', 'MIX', 'RATIOS', 'GASTOS SIN MARCA', 'GASTO LOGISTICO SIN MARCA',
     'X UNIDAD', 'X HECTOLITRO', 'MARGEN', 'CANAL PREVENTA', 'NO ASIGNABLES', 'POR MARCA'
  ];

  let currentConceptKey = null;

  if (scenarioColIdx !== -1) {
      data.forEach(row => {
          if (!row) return;
          let conceptCellStr = '';
          for (let j = 0; j <= 3; j++) {
             if (row[j] && typeof row[j] === 'string' && isNaN(row[j])) {
                 conceptCellStr = row[j].trim().toUpperCase();
                 break;
             }
          }
          if (!conceptCellStr) return;

          let foundConcept = false;
          for (const [kw, key] of Object.entries(conceptsKeywords)) {
              if (conceptCellStr.includes(kw)) {
                 currentConceptKey = key;
                 foundConcept = true;
                 break;
              }
          }

          if (!foundConcept) {
             for (const kw of ignoreKeywords) {
                 if (conceptCellStr.includes(kw)) {
                    currentConceptKey = null;
                    foundConcept = true;
                    break;
                 }
             }
          }

          if (foundConcept) return;

          if (currentConceptKey) {
              let normalizedBrand = conceptCellStr.replace(/[^A-Z0-9]/g, '');
              if (!brandsData[normalizedBrand]) brandsData[normalizedBrand] = {};
              brandsData[normalizedBrand][currentConceptKey] = safeNum(row[scenarioColIdx]);
              // console.log("Brands Data update:", currentConceptKey, normalizedBrand, row[scenarioColIdx]);
          }
      });
  }

  const marcasPermitidas = [
    "APA BOTELLON 18.9 LTS (x1)",
    "APA BOTELLA 0.5 LTS (x20)",
    "APA BOTELLA 1.5 LTS (x12)",
    "APA OTRAS",
    "MAQUILA AGUA OTROS",
    "MAQUILA AGUA 1.5 LTS (x12)",
    "MAQUILA AGUA 0.5 LTS (x20)",
    "BON",
    "PA SABOR 0.5 LTS (x12)",
    "PA H+ 0.71 LTS (x12)",
    "Total general"
  ];
  
  const safeDiv = (num, den) => (den === 0 || isNaN(den) || !den) ? 0 : (num / den);

  let htmlTotales = '';
  let htmlUnitarios = '';

  const totals = {
      unidades: 0,
      ventas: 0,
      costo: 0,
      utilidad_bruta: 0,
      logistica: 0,
      utilidad_post: 0,
      apoyo_comercial: 0,
      contribucion: 0
  };

  marcasPermitidas.forEach(marca => {
      const isTotal = String(marca).toLowerCase().includes('total');
      const normMarca = marca.toUpperCase().replace(/[^A-Z0-9]/g, '');
      const rowData = brandsData[normMarca] || {};
      
      let unidades, ventas, costo, utilidad_bruta, logistica, utilidad_post, apoyo, contribucion;

      if (isTotal) {
          unidades = totals.unidades;
          ventas = totals.ventas;
          costo = totals.costo;
          utilidad_bruta = totals.utilidad_bruta;
          logistica = totals.logistica;
          utilidad_post = totals.utilidad_post;
          apoyo = totals.apoyo_comercial;
          contribucion = totals.contribucion;
      } else {
          unidades = rowData.unidades || 0;
          ventas = rowData.ventas || 0;
          costo = rowData.costo || 0;
          utilidad_bruta = rowData.utilidad_bruta || (ventas - costo);
          logistica = rowData.logistica || 0;
          
          // User specific rule:
          utilidad_post = utilidad_bruta - logistica;
          
          apoyo = rowData.apoyo_comercial || 0;
          contribucion = utilidad_post - apoyo;

          totals.unidades += unidades;
          totals.ventas += ventas;
          totals.costo += costo;
          totals.utilidad_bruta += utilidad_bruta;
          totals.logistica += logistica;
          totals.utilidad_post += utilidad_post;
          totals.apoyo_comercial += apoyo;
          totals.contribucion += contribucion;
      }

      const pct_mb = safeDiv(utilidad_bruta, ventas);
      const pct_log = safeDiv(logistica, ventas);
      const pct_postlog = safeDiv(utilidad_post, ventas);
      const pct_apoyo = safeDiv(apoyo, ventas);

      const divisor_un = unidades * 12;
      const px_un = safeDiv(ventas * 1000, divisor_un);
      const costo_un = safeDiv(costo * 1000, divisor_un);
      const mb_un = safeDiv(utilidad_bruta * 1000, divisor_un);
      const log_un = safeDiv(logistica * 1000, divisor_un);
      const upost_un = safeDiv(utilidad_post * 1000, divisor_un);

      const trStyle = isTotal 
        ? 'background: #bdd7ee; color: black; font-weight: bold; border-top: 2px solid #9cc2e5;'
        : 'background: white; border-bottom: 1px solid var(--border);';

      const formatCell = (val, isMoney, isPercent, isUnits) => {
         if (val === undefined || val === null || val === 0) return '-';
         let r = safeNum(val);
         if (isUnits) return fmtVol(r);
         if (isPercent) return fmtPct(r);
         return fmtPrecio(r);
      };

      const cUnidades = formatCell(unidades, false, false, true);
      const cVentas = formatCell(ventas, true, false, false);
      const cCosto = formatCell(costo, true, false, false);
      const cUtilidad = formatCell(utilidad_bruta, true, false, false);
      const cPctMB = formatCell(pct_mb, false, true, false);
      const cLogistica = formatCell(logistica, true, false, false);
      const cPctLog = formatCell(pct_log, false, true, false);
      const cUtilidadPost = formatCell(utilidad_post, true, false, false);
      const cPctPost = formatCell(pct_postlog, false, true, false);
      const cApoyo = formatCell(apoyo, true, false, false);
      const cPctApoyo = formatCell(pct_apoyo, false, true, false);
      const cContrib = formatCell(contribucion, true, false, false);

      const cPxUn = formatCell(px_un, true, false, false);
      const cCostoUn = formatCell(costo_un, true, false, false);
      const cMbUn = formatCell(mb_un, true, false, false);
      const cLogUn = formatCell(log_un, true, false, false);
      const cUpostUn = formatCell(upost_un, true, false, false);

      htmlTotales += `
        <tr style="${trStyle}">
          <td style="padding: 10px 12px; border-right: 1px solid rgba(0,0,0,0.1); font-weight: ${isTotal ? 'bold' : '600'}; font-size: 0.95rem; ${isTotal ? 'color: black;' : ''}">${marca}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem;">${cUnidades}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem;">${cVentas}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem;">${cCosto}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; font-weight: bold;">${cUtilidad}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 2px solid #cbd5e1; background: ${isTotal ? 'transparent' : '#f1f5f9'};">${cPctMB}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem;">${cLogistica}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 2px solid #cbd5e1; background: ${isTotal ? 'transparent' : '#f1f5f9'};">${cPctLog}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; font-weight: bold;">${cUtilidadPost}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 2px solid #cbd5e1; background: ${isTotal ? 'transparent' : '#f1f5f9'};">${cPctPost}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem;">${cApoyo}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 2px solid #cbd5e1; background: ${isTotal ? 'transparent' : '#f1f5f9'};">${cPctApoyo}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; font-weight: bold; border-right: 2px solid #cbd5e1;">${cContrib}</td>
        </tr>
      `;

      htmlUnitarios += `
        <tr style="${trStyle}">
          <td style="padding: 10px 12px; border-right: 1px solid rgba(0,0,0,0.1); font-weight: ${isTotal ? 'bold' : '600'}; font-size: 0.95rem; ${isTotal ? 'color: black;' : ''}">${marca}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 1px dashed #cbd5e1; border-left: 2px solid #cbd5e1; background: ${isTotal ? 'transparent' : '#f8fafc'};">${cPxUn}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 1px dashed #cbd5e1; background: ${isTotal ? 'transparent' : '#f8fafc'};">${cCostoUn}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 1px dashed #cbd5e1; background: ${isTotal ? 'transparent' : '#f8fafc'};">${cMbUn}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; border-right: 1px dashed #cbd5e1; background: ${isTotal ? 'transparent' : '#f8fafc'};">${cLogUn}</td>
          <td style="padding: 10px 12px; text-align: right; font-family: var(--font-mono); font-size: 0.95rem; background: ${isTotal ? 'transparent' : '#f8fafc'};">${cUpostUn}</td>
        </tr>
      `;
  });

  const tbodyTotales = document.getElementById('pg-horizontal-tbody');
  const tbodyUnitarios = document.getElementById('pg-horizontal-unitarios-tbody');
  if (tbodyTotales) tbodyTotales.innerHTML = htmlTotales;
  if (tbodyUnitarios) tbodyUnitarios.innerHTML = htmlUnitarios;
}

export async function processPgHorizontalWorkbook(workbook) {
  let pgSheetName = workbook.SheetNames.find(n => {
      const lower = n.toLowerCase();
      return lower.includes('analítico pyg') || lower.includes('analitico pyg') || 
             lower.includes('p&g') || lower.includes('horizontal') || lower.includes('pyg');
  }) || workbook.SheetNames[0];

  let pgData = [];
  if (pgSheetName) {
     const rows = XLSX.utils.sheet_to_json(workbook.Sheets[pgSheetName], {header: 1, defval: null});
     // Keep all rows so we can parse them based on headers
     pgData = rows;
  }

  if (!comercialRawData) {
      comercialRawData = {
          dataF: [],
          data2025: [],
          ppto: { vol: [], vta: [] },
          pgHorizontal: []
      };
  }

  comercialRawData.pgHorizontal = pgData;

  // Guardar en caché local
  try {
    const db = await openDB();
    await dbPut(db, 'COMERCIAL_KEY', { data: comercialRawData, timestamp: Date.now() });
    console.log("💾 [comercialEngine] Datos de P&G Horizontal persistidos con éxito en IndexedDB.");
  } catch (e) {
    console.warn('[comercialEngine] Fail to cache:', e);
  }

  return comercialRawData;
}

// ------------------------------------------------------------------
// PROCESADO DE LIBRO COMMERCIAL (EXCEL)
// ------------------------------------------------------------------
export async function processComercialWorkbook(workbook) {
  const sheetNames = workbook.SheetNames.map(n => n.toLowerCase());

  const findSheet = (keyword) => workbook.SheetNames.find(n => {
    const s = normalizeText(n);
    return s.includes(normalizeText(keyword));
  });

  const nameDataF    = findSheet('dataf') || sheetNames.find(n => n.includes('real 2026')) || workbook.SheetNames[0];
  const nameData2025 = findSheet('data 2025') || sheetNames.find(n => n.includes('historico')) || workbook.SheetNames[1];
  const namePPTO     = findSheet('ppto') || sheetNames.find(n => n.includes('presupuesto')) || workbook.SheetNames[2];
  const nameSixPlusSix = findSheet('6+6') || sheetNames.find(n => n.includes('6+6'));

  // Try to find the P&G Horizontal sheet based on known headers or keywords
  let pgSheetName = findSheet('analítico pyg') || findSheet('analitico pyg') || findSheet('margen bruto') || findSheet('p&g');
  if (!pgSheetName) {
    // Scan sheets to see if any has "UNIDADES TOTALES MES"
    for (const sheetName of workbook.SheetNames) {
       const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {header: 1});
       for (let i = 0; i < Math.min(rows.length, 10); i++) {
         if (rows[i] && rows[i].some(cell => String(cell).toUpperCase().includes('UNIDADES TOTALES'))) {
           pgSheetName = sheetName;
           break;
         }
       }
       if (pgSheetName) break;
    }
  }
  if (!pgSheetName) {
    // Fallback: Check if first sheet actually has the keyword
    const firstRows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
    const hasPgData = firstRows.slice(0, 10).some(r => r && r.some(c => String(c).toUpperCase().includes('UNIDADES TOTALES')));
    if (hasPgData) {
      pgSheetName = workbook.SheetNames[0];
    } else {
      pgSheetName = null;
    }
  }

  console.log('📌 Hoja detectada - DataF:', nameDataF);
  console.log('📌 Hoja detectada - Data 2025:', nameData2025);
  console.log('📌 Hoja detectada - PPTO:', namePPTO);
  console.log('📌 Hoja detectada - P&G Horizontal:', pgSheetName);

  let pgData = [];
  if (pgSheetName) {
     const rows = XLSX.utils.sheet_to_json(workbook.Sheets[pgSheetName], {header: 1, defval: null});
     pgData = rows;
  }

  // Preserve pgHorizontal if we already have it and current upload doesn't have it (or we fallback to existing)
  const safePgData = (() => {
    if (pgData && pgData.length > 0) return pgData;
    if (comercialRawData && comercialRawData.pgHorizontal && comercialRawData.pgHorizontal.length > 0) {
      return comercialRawData.pgHorizontal;
    }
    return [];
  })();

  let finalPgData = safePgData;
  if (finalPgData.length === 0) {
    try {
      const db = await openDB();
      const cached = await dbGet(db, 'COMERCIAL_KEY');
      if (cached && cached.data && cached.data.pgHorizontal && cached.data.pgHorizontal.length > 0) {
        finalPgData = cached.data.pgHorizontal;
        console.log('[comercialEngine] pgHorizontal recuperado de IndexedDB como fallback.');
      }
    } catch (e) {
      console.warn('[comercialEngine] No se pudo recuperar pgHorizontal desde IndexedDB:', e);
    }
  }

  comercialRawData = {
    dataF:   nameDataF ? parseDataF(workbook.Sheets[workbook.SheetNames[sheetNames.indexOf(nameDataF.toLowerCase())]]) : (comercialRawData?.dataF || []),
    data2025: nameData2025 ? parseData2025(workbook.Sheets[workbook.SheetNames[sheetNames.indexOf(nameData2025.toLowerCase())]]) : (comercialRawData?.data2025 || []),
    ppto:    namePPTO ? parsePPTO(workbook.Sheets[workbook.SheetNames[sheetNames.indexOf(namePPTO.toLowerCase())]]) : (comercialRawData?.ppto || { vol: [], vta: [] }),
    sixPlusSix: nameSixPlusSix ? parsePPTO(workbook.Sheets[workbook.SheetNames[sheetNames.indexOf(nameSixPlusSix.toLowerCase())]]) : (comercialRawData?.sixPlusSix || { vol: [], vta: [] }),
    pgHorizontal: finalPgData
  };

  // Guardar en caché local
  try {
    const db = await openDB();
    await dbPut(db, 'COMERCIAL_KEY', { data: comercialRawData, timestamp: Date.now() });
    console.log("💾 [comercialEngine] Datos del Resumen Comercial persistidos con éxito en IndexedDB.");
  } catch (e) {
    console.warn('[comercialEngine] Fail to cache:', e);
  }

  return comercialRawData;
}

// Cargar desde caché si existe
export async function loadComercialCache() {
  try {
    const db = await openDB();
    const record = await dbGet(db, 'COMERCIAL_KEY');
    if (record) {
      comercialRawData = record.data;
      console.log("📂 [comercialEngine] Caché de Resumen Comercial cargada de IndexedDB.");
      return true;
    }
  } catch (e) {
    console.warn('[comercialEngine] Sin caché disponible.');
  }
  return false;
}

export function hasComercialData() {
  return !!comercialRawData;
}

export function getComercialRawData() { return comercialRawData; }
export function getArbolComercial() { return ARBOL_COMERCIAL; }

export async function processManualFile(arrayBuffer) {
  try {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const result = await processComercialWorkbook(workbook);
    if (result) {
        const db = await openDB();
        await dbPut(db, 'COMERCIAL_KEY', { data: comercialRawData, timestamp: Date.now() });
        return result;
    }
  } catch (e) {
    console.error('[comercialEngine] Error en processManualFile:', e);
  }
  return null;
}

// ------------------------------------------------------------------
// INDEXEDDB HELPERS
// ------------------------------------------------------------------
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('PlanetaAzulDB', 7);
    req.onupgradeneeded = (e) => {
      if (!e.target.result.objectStoreNames.contains('finance_cache')) {
        e.target.result.createObjectStore('finance_cache');
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function dbPut(db, key, value) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction('finance_cache', 'readwrite');
    tx.objectStore('finance_cache').put(value, key);
    tx.oncomplete = resolve;
    tx.onerror = reject;
  });
}

function dbGet(db, key) {
  return new Promise((resolve) => {
    const req = db.transaction('finance_cache', 'readonly').objectStore('finance_cache').get(key);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => resolve(null);
  });
}

export function resetComercialEngine() {
  comercialRawData = null;
}
