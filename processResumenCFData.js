/**
 * Procesa y agrupa datos comerciales en las jerarquías financieras del "Resumen CF" (Rango I3:X46).
 * Diseñado minuciosamente para ser inmune a diferencias de formato, espacios en blanco y variaciones en nombres de columnas.
 */
export function processResumenCFData(data2025, data2026, dataPPTO, mesSeleccionado) {
  console.log("📊 [processResumenCFData] Iniciando proceso -> 2025:", data2025?.length, "| 2026:", data2026?.length, "| PPTO:", dataPPTO?.length, "| Mes seleccionado:", mesSeleccionado);

  const initCategory = () => ({
    volumen_2025: 0, volumen_2026: 0, volumen_ppto: 0,
    ventas_2025: 0, ventas_2026: 0, ventas_ppto: 0
  });

  const raw = {};
  const allKeys = [
    "BT5",
    "EVP",
    "Marcas Propias",
    "16 Oz_Propias",
    "1.5 LTS_Propias",
    "Otros EVP Agua_Propias",
    "100% Rpet_Propias",
    "5 LTS_Propias",
    "8 VASOS",
    "APA SPORT",
    "TETRA PACK",
    "16 OZ (FARDO 12)",
    "Marcas Privadas",
    "16 Oz_Privadas",
    "1.5 LTS_Privadas",
    "Otros EVP Agua_Privadas",
    "100% Rpet_Privadas",
    "5 LTS_Privadas",
    "Bebidas",
    "Agua Saborizada",
    "Hidractive+",
    "BON",
    "TOTAL"
  ];

  allKeys.forEach(k => {
    raw[k] = initCategory();
  });

  // HELPER 1: Limpia y parsea strings numéricos de cualquier formato
  const parseNum = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    if (typeof val === 'number') return val;
    const cleanStr = String(val).replace(/,/g, '').replace(/\$/g, '').trim();
    return parseFloat(cleanStr) || 0;
  };

  // HELPER 2: Obtiene dinámicamente columnas de un registro (safe-headers)
  const getCol = (row, keywords) => {
    if (!row) return undefined;

    const kwUpper = keywords.map(k => k.toUpperCase());
    if (kwUpper.some(k => k.includes('FAMILIA') || k.includes('FAMILY')) && row.familia !== undefined) return row.familia;
    if (kwUpper.some(k => k.includes('AGRUPACION') || k.includes('AGRUPACIÓN')) && row.agrupacion !== undefined) return row.agrupacion;
    if (kwUpper.some(k => k.includes('DESC') || k.includes('PROD') || k.includes('DESCRIPCION')) && row.descProd !== undefined) return row.descProd;
    if (kwUpper.some(k => k.includes('CANTIDAD') || k.includes('VOLUMEN') || k.includes('QTY')) && row.cantidad !== undefined) return row.cantidad;
    if (kwUpper.some(k => k.includes('INGRESO') || k.includes('VENTAS') || k.includes('MONTO') || k.includes('VALUE')) && row.ingreso !== undefined) return row.ingreso;

    if (Array.isArray(row)) {
      if (kwUpper.some(k => k.includes('FAMILIA') || k.includes('FAMILY'))) return row[12];
      if (kwUpper.some(k => k.includes('AGRUPACION') || k.includes('AGRUPACIÓN'))) return row[11];
      if (kwUpper.some(k => k.includes('DESC') || k.includes('PROD') || k.includes('DESCRIPCION'))) return row[11] || row[3];
      if (kwUpper.some(k => k.includes('CANTIDAD') || k.includes('VOLUMEN') || k.includes('QTY'))) return row[6];
      if (kwUpper.some(k => k.includes('INGRESO') || k.includes('VENTAS') || k.includes('MONTO') || k.includes('VALUE'))) return row[7];
      if (kwUpper.some(k => k.includes('PERIODO') || k.includes('PERIOD') || k.includes('MES'))) return row[21];
    }

    const keys = Object.keys(row);
    const upperKws = keywords.map(kw => kw.toUpperCase().trim());

    // Búsqueda exacta
    for (const key of keys) {
      const upperKey = key.toUpperCase().trim();
      if (upperKws.includes(upperKey)) return row[key];
    }

    // Búsqueda parcial
    for (const key of keys) {
      const upperKey = key.toUpperCase().trim();
      if (upperKws.some(kw => upperKey.includes(kw))) return row[key];
    }
    return undefined;
  };

  // HELPER 2.5: Obtiene el valor original del periodo
  const getPeriodoVal = (row) => {
    if (!row) return undefined;
    if (row.mes !== undefined) return row.mes;
    if (row.Periodo !== undefined) return row.Periodo;
    if (row.PERIODO !== undefined) return row.PERIODO;
    if (row.periodo !== undefined) return row.periodo;
    if (row.Period !== undefined) return row.Period;
    if (row.PERIOD !== undefined) return row.PERIOD;
    if (row.period !== undefined) return row.period;
    if (Array.isArray(row)) return row[21];
    return getCol(row, ['PERIODO', 'PERIOD', 'MES', 'PERIODOS']);
  };

  // HELPER 2.6: Parsea valores de fecha o códigos seriales de Excel a número de mes
  const excelDateToMonth = (val) => {
    if (val === undefined || val === null) return null;
    if (val instanceof Date) return val.getMonth() + 1;
    
    const trimmed = String(val).trim();
    const numVal = parseInt(trimmed, 10);
    if (!isNaN(numVal) && numVal >= 1 && numVal <= 12 && (trimmed.length === 1 || trimmed.length === 2)) {
      return numVal;
    }
    
    const num = Number(val);
    if (!isNaN(num) && num > 25569) { // serial Excel
      const d = new Date(Math.round((num - 25569) * 86400 * 1000));
      if (!isNaN(d.getTime())) return d.getMonth() + 1;
    }
    
    const dStr = new Date(trimmed);
    if (!isNaN(dStr.getTime())) return dStr.getMonth() + 1;
    
    return null;
  };

  // HELPER 2.7: Extrae Familia (Columna M)
  const getFamiliaColM = (row) => {
    if (!row) return '';
    if (Array.isArray(row)) return String(row[12] || '').trim().toUpperCase();
    if (row.familia !== undefined) return String(row.familia || '').trim().toUpperCase();
    if (row.FAMILIA !== undefined) return String(row.FAMILIA || '').trim().toUpperCase();
    if (row['Familia'] !== undefined) return String(row['Familia'] || '').trim().toUpperCase();
    return String(getCol(row, ['FAMILIA', 'FAMILY']) || '').trim().toUpperCase();
  };

  // HELPER 2.8: Extrae Desc Product (Columna L)
  const getDescProductColL = (row) => {
    if (!row) return '';
    if (Array.isArray(row)) return String(row[11] || '').trim().toUpperCase();
    const possibleKeys = [
      'Desc Product.', 'DESC PRODUCT.', 'descProduct', 'desc_product', 
      'Desc Product', 'DESC PRODUCT', 'desc_product.', 'agrupacion', 'AGRUPACION', 
      'agrupación', 'AGRUPACIÓN', 'descProd'
    ];
    for (const key of possibleKeys) {
      if (row[key] !== undefined) {
        return String(row[key] || '').trim().toUpperCase();
      }
    }
    return String(getCol(row, ['DESC PRODUCT.', 'DESC PRODUCT', 'DESC PROD', 'AGRUPACION', 'AGRUPACIÓN']) || '').trim().toUpperCase();
  };

  const isPrivRow = (fam, agrp, dsc) => {
    return fam.includes('PRIVADA') || fam.includes('PRIVADAS') || agrp.includes('MARCAS PRIVADAS') || agrp.includes('MARCA PRIVADA') || agrp.includes('PRIVADAS') || agrp.includes('PRIVADA') || dsc.includes('PRIVADA') || dsc.includes('PRIVADAS') || dsc.includes('PRICESMART') || dsc.includes('PRICE SMART') || dsc.includes('MEMBER');
  };

  // CLASIFICADOR DINÁMICO DE FILAS SEGÚN REGLAS DE NEGOCIO EXCEL RESUMEN CF
  const getCategory = (row) => {
    const familia = getFamiliaColM(row);
    const desc = getDescProductColL(row);

    const fUpper = String(getCol(row, ['FAMILIA', 'FAMILY']) || '').trim().toUpperCase();
    const aUpper = String(getCol(row, ['AGRUPACION', 'AGRUPACIÓN']) || '').trim().toUpperCase();
    const dUpper = String(getCol(row, ['DESC PRODUCT.', 'DESC PROD', 'DESCRIPCION MATERIAL', 'DESCRIPCION', 'PRODUCTO']) || '').trim().toUpperCase();

    // Normalizador determinístico: uppercase, trim, sin tildes, solo alfanuméricos y paréntesis, colapsar espacios.
    const normDp = dUpper.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\w\s\(\)]/gi, '').replace(/\s+/g, ' ');

    // REGLA PRIORIDAD #1 Y #2 (Aislado de inferencias semánticas)
    if (normDp === '16 OZ (FARDO 12)') return '16 OZ (FARDO 12)';
    if (normDp === '16 OZ') {
      if (fUpper === 'MARCAS PRIVADAS' || isPrivRow(fUpper, aUpper, dUpper)) return '16 Oz_Privadas';
      return '16 Oz_Propias';
    }

    // 1. BT5
    if (fUpper.includes('BT5') || dUpper.includes('BT5') || dUpper.includes('BOTELLON') || dUpper.includes('BOTELLÓN')) return 'BT5';

    // 2. BON
    if (fUpper.includes('BON') || aUpper.includes('BON') || dUpper.includes('BON') || fUpper.includes('CCN') || aUpper.includes('CCN') || dUpper.includes('CCN')) return 'BON';

    // 3. BEBIDAS
    if (fUpper.includes('BEBIDAS') || dUpper.includes('SABOR') || dUpper.includes('HIDRACTIVE') || dUpper.includes('ISO')) {
      if (dUpper.includes('SABOR')) return 'Agua Saborizada';
      return 'Hidractive+';
    }

    // 4. MARCAS PRIVADAS
    if (isPrivRow(fUpper, aUpper, dUpper)) {
      if (dUpper.includes('1.5')) return '1.5 LTS_Privadas';
      if (dUpper.includes('RPET') || dUpper.includes('R-PET')) return '100% Rpet_Privadas';
      if (/(?<![\d.,])5\s*(LTS|LT|L|LITER|LITROS|LITRO)\b/.test(dUpper)) return '5 LTS_Privadas';
      return 'Otros EVP Agua_Privadas';
    }

    // 5. MARCAS PROPIAS (Todo lo que sobró de EVP)
    if (dUpper.includes('1.5')) return '1.5 LTS_Propias';
    if (dUpper.includes('RPET') || dUpper.includes('R-PET')) return '100% Rpet_Propias';
    if (/(?<![\d.,])5\s*(LTS|LT|L|LITER|LITROS|LITRO)\b/.test(dUpper)) return '5 LTS_Propias';
    if (dUpper.includes('8 VASOS') || dUpper.includes('VASO')) return '8 VASOS';
    if (dUpper.includes('SPORT')) return 'APA SPORT';
    if (dUpper.includes('TETRA')) return 'TETRA PACK';
    
    // CORTAFUEGOS: Ya no devolvemos '16 Oz_Propias' aquí abajo para evitar basura.
    // Qualquer EVP no clasificado se va directo a Otros.
    return 'Otros EVP Agua_Propias';
  };

  // MAPAS DE CABECERAS DINÁMICOS
  let headers2025 = null;
  let headers2026 = null;

  // LECTOR INMUNE A DESPLAZAMIENTOS DE COLUMNAS EXCEL
  const extractRowData = (row, is2026) => {
      let obj = row;
      
      // 1. Si los datos vienen como arreglos (Excel crudo), mapeamos las posiciones dinámicamente
      if (Array.isArray(row)) {
          // Detectar si esta fila es la cabecera
          const hasHeaders = row.some(c => typeof c === 'string' && (c.toUpperCase().trim() === 'FAMILIA' || c.toUpperCase().trim().includes('DESC PRODUCT') || c.toUpperCase().trim() === 'PERIODO'));
          if (hasHeaders) {
              const map = {};
              row.forEach((col, i) => { 
                  if (col !== undefined && col !== null) {
                      map[String(col).trim().toUpperCase()] = i; 
                  }
              });
              if (is2026) headers2026 = map; else headers2025 = map;
              return null; // Es la fila de títulos, la saltamos
          }
          
          const activeMap = is2026 ? headers2026 : headers2025;
          if (activeMap) {
              const getIdx = (aliases) => {
                  for (const a of aliases) {
                      if (activeMap[a] !== undefined) return activeMap[a];
                  }
                  return -1;
              };
              const idxDesc = getIdx(['DESC PRODUCT.', 'DESC PRODUCT', 'DESC PRODUCTO', 'DESC PROD CF', 'DESC_PRODUCT', 'DESC PROD', 'DESCRIPCION']);
              const idxFam = getIdx(['FAMILIA', 'FAMILY', 'LINEA']);
              const idxPer = getIdx(['PERIODO', 'PERIOD', 'MES', 'PERIODOS']);
              const idxCant = getIdx(['CANTIDAD VENDIDA', 'CANTIDAD', 'VOLUMEN', 'CANT VEND', 'QTY']);
              const idxIng = getIdx(['INGRESO TOTAL', 'INGRESO', 'VENTAS', 'VENTA', 'MONTO', 'IMPORTE']);
              
              obj = {
                  'DESC PRODUCT.': idxDesc !== -1 ? row[idxDesc] : '',
                  'FAMILIA': idxFam !== -1 ? row[idxFam] : '',
                  'PERIODO': idxPer !== -1 ? row[idxPer] : '',
                  'CANTIDAD VENDIDA': idxCant !== -1 ? row[idxCant] : 0,
                  'INGRESO TOTAL': idxIng !== -1 ? row[idxIng] : 0
              };
          } else {
              // Fallback extremo por índices fijos si no detectó la cabecera aún
              obj = {
                  'DESC PRODUCT.': row[11] || '',
                  'FAMILIA': row[12] || '',
                  'PERIODO': row[21] || 0,
                  'CANTIDAD VENDIDA': row[6] || 0,
                  'INGRESO TOTAL': row[7] || 0
              };
          }
      }

      // 2. Extracción segura asegurando nombres de columnas
      let colL = '', colM = '', colV = 0, colG = 0, colH = 0;
      if (obj && typeof obj === 'object') {
          const keys = Object.keys(obj);
          for (const key of keys) {
              const kSafe = key.trim().toUpperCase();
              if (kSafe === 'DESC PRODUCT.' || kSafe === 'DESC PRODUCT' || kSafe === 'DESC PRODUCTO' || kSafe === 'DESC PROD CF' || kSafe === 'DESC_PRODUCT' || kSafe === 'DESC PROD') {
                  colL = String(obj[key] || '').trim().toUpperCase();
              }
              if (kSafe === 'FAMILIA') {
                  colM = String(obj[key] || '').trim().toUpperCase();
              }
              if (kSafe === 'PERIODO' || kSafe === 'PERIOD' || kSafe === 'MES') {
                  colV = parseInt(String(obj[key] || '').trim(), 10);
              }
              if (kSafe === 'CANTIDAD VENDIDA' || kSafe === 'CANTIDAD' || kSafe === 'VOLUMEN' || kSafe === 'CANT VEND') {
                  colG = parseNum(obj[key]);
              }
              if (kSafe === 'INGRESO TOTAL' || kSafe === 'INGRESO' || kSafe === 'VENTAS' || kSafe === 'VENTA') {
                  colH = parseNum(obj[key]);
              }
          }
      }
      return { colL, colM, colV, colG, colH };
  };

  // PROCESAMIENTO HISTÓRICO 2025
  (data2025 || []).forEach(row => {
    const dataExtracted = extractRowData(row, false);
    if (!dataExtracted) return; // Saltamos si era la cabecera
    const { colL, colM, colV, colG, colH } = dataExtracted;

    if (mesSeleccionado !== undefined && mesSeleccionado !== null && colV !== mesSeleccionado) return;

    let cat = getCategory(row);

    if (cat && raw[cat]) {
      raw[cat].volumen_2025 += colG;
      raw[cat].ventas_2025 += colH;
    }
  });

  // PROCESAMIENTO REAL 2026 (Estilo SUMIFS Excel)
  (data2026 || []).forEach(row => {
    const dataExtracted = extractRowData(row, true);
    if (!dataExtracted) return; // Saltamos si era la cabecera
    const { colL, colM, colV, colG, colH } = dataExtracted;
    
    // Filtro estricto del mes
    if (mesSeleccionado !== undefined && mesSeleccionado !== null && colV !== mesSeleccionado) return;

    let cat = null;
    const normDp = colL.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\w\s\(\).%-]/gi, '').replace(/\s+/g, ' ').trim();
    const fUpper = colM.trim().toUpperCase();

    // Regla Independiente Obligatoria
    if (normDp === '16 OZ (FARDO 12)') {
      cat = '16 OZ (FARDO 12)';
    }
    // BT5 Mantenemos como solicitado
    else if (fUpper.includes('BT5') || colL.includes('BT5') || colL.includes('BOTELLON') || colL.includes('BOTELLÓN')) {
      cat = 'BT5';
    }
    // SUMIFS exacto para MARCAS PROPIAS
    else if (fUpper === 'MARCAS PROPIAS') {
      if (normDp === '16 OZ' || normDp === '16 ONZ') cat = '16 Oz_Propias';
      else if (normDp === '1.5 LTS' || normDp.includes('1.5')) cat = '1.5 LTS_Propias';
      else if (normDp === '100% RPET' || normDp.includes('RPET') || normDp.includes('R-PET')) cat = '100% Rpet_Propias';
      else if (normDp === '5 LTS' || /(?<![\d.,])5\s*(LTS|LT|L)\b/.test(normDp)) cat = '5 LTS_Propias';
      else if (normDp === '8 VASOS' || normDp.includes('8 VASOS')) cat = '8 VASOS';
      else if (normDp === 'APA SPORT' || normDp.includes('SPORT')) cat = 'APA SPORT';
      else if (normDp === 'TETRA PACK' || normDp.includes('TETRA')) cat = 'TETRA PACK';
      else cat = 'Otros EVP Agua_Propias';
    } 
    // Fallback al getCategory anterior mientras avanzamos en la reconstrucción
    else {
      cat = getCategory(row);
    }

    if (cat && raw[cat]) {
      raw[cat].volumen_2026 += colG;
      raw[cat].ventas_2026 += colH;
    }
  });

  // PROCESAMIENTO PRESUPUESTO (PPTO)
  (dataPPTO || []).forEach(row => {
    const cat = getCategory(row);
    const tipo = String(getCol(row, ['TIPO']) || '').trim().toUpperCase();

    let rowSum = 0;
    if (row.monthValues && Array.isArray(row.monthValues)) {
      if (mesSeleccionado !== undefined && mesSeleccionado !== null) {
        rowSum = parseNum(row.monthValues[mesSeleccionado - 1]);
      } else {
        rowSum = row.monthValues.reduce((sum, val) => sum + parseNum(val), 0);
      }
    } else {
      const fyVal = getCol(row, ['FY', 'TOTAL', 'ANUAL', 'YTD']);
      if (fyVal !== undefined && fyVal !== '') {
        rowSum = parseNum(fyVal);
      } else {
        const skipCols = ['TIPO', 'AGRUPACION', 'AGRUPACIÓN', 'FAMILIA', 'DESC', 'LINEA', 'CODIGO', 'MATERIAL', 'PRODUCTO', 'DESCRIPCIÓN', 'DESCRIPCION'];
        let foundSpecificMonth = false;
        
        if (mesSeleccionado !== undefined && mesSeleccionado !== null) {
          const MONTH_KEYS_ES = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];
          const MONTH_KEYS_EN = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
          const targetEs = MONTH_KEYS_ES[mesSeleccionado - 1];
          const targetEn = MONTH_KEYS_EN[mesSeleccionado - 1];
          const targetNum = String(mesSeleccionado).padStart(2, '0');
          
          for (const key in row) {
            const safeKey = key.toUpperCase().trim();
            if (safeKey === targetEs || safeKey === targetEn || safeKey === targetNum || safeKey.startsWith(targetEs) || safeKey.startsWith(targetEn)) {
              rowSum += parseNum(row[key]);
              foundSpecificMonth = true;
            }
          }
        }
        
        if (!foundSpecificMonth) {
          for (const key in row) {
            const safeKey = key.toUpperCase().trim();
            if (!skipCols.some(skip => safeKey.includes(skip))) {
              rowSum += parseNum(row[key]);
            }
          }
        }
      }
    }

    if (raw[cat]) {
      if (tipo.includes('VOLUMEN') || tipo.includes('CANTIDAD') || tipo.includes('UNIT') || tipo.includes('QTY')) {
        raw[cat].volumen_ppto += rowSum;
      } else if (tipo.includes('INGRESO') || tipo.includes('VENTA') || tipo.includes('MONTO') || tipo.includes('VALUE')) {
        raw[cat].ventas_ppto += rowSum;
      }
    }
  });

  // CONSOLIDACIÓN CON REGLAS JERÁRQUICAS ESTRICTAS
  // 1. Marcas Propias -> Otros EVP Agua_Propias (Suma de sus sub-categorías de nivel inferior)
  raw["Otros EVP Agua_Propias"] = initCategory();
  const subcapsPropias = ["100% Rpet_Propias", "5 LTS_Propias", "8 VASOS", "APA SPORT", "TETRA PACK", "16 OZ (FARDO 12)"];
  subcapsPropias.forEach(sub => {
    raw["Otros EVP Agua_Propias"].volumen_2025 += raw[sub].volumen_2025;
    raw["Otros EVP Agua_Propias"].ventas_2025 += raw[sub].ventas_2025;
    raw["Otros EVP Agua_Propias"].volumen_2026 += raw[sub].volumen_2026;
    raw["Otros EVP Agua_Propias"].ventas_2026 += raw[sub].ventas_2026;
    raw["Otros EVP Agua_Propias"].volumen_ppto += raw[sub].volumen_ppto;
    raw["Otros EVP Agua_Propias"].ventas_ppto += raw[sub].ventas_ppto;
  });

  // 2. Marcas Propias = "16 Oz_Propias" + "1.5 LTS_Propias" + "Otros EVP Agua_Propias"
  raw["Marcas Propias"] = initCategory();
  const mpItems = ["16 Oz_Propias", "1.5 LTS_Propias", "Otros EVP Agua_Propias"];
  mpItems.forEach(item => {
    raw["Marcas Propias"].volumen_2025 += raw[item].volumen_2025;
    raw["Marcas Propias"].ventas_2025 += raw[item].ventas_2025;
    raw["Marcas Propias"].volumen_2026 += raw[item].volumen_2026;
    raw["Marcas Propias"].ventas_2026 += raw[item].ventas_2026;
    raw["Marcas Propias"].volumen_ppto += raw[item].volumen_ppto;
    raw["Marcas Propias"].ventas_ppto += raw[item].ventas_ppto;
  });

  // 3. Marcas Privadas -> Otros EVP Agua_Privadas (Suma de sus sub-categorías)
  raw["Otros EVP Agua_Privadas"] = initCategory();
  const subcapsPrivadas = ["100% Rpet_Privadas", "5 LTS_Privadas"];
  subcapsPrivadas.forEach(sub => {
    raw["Otros EVP Agua_Privadas"].volumen_2025 += raw[sub].volumen_2025;
    raw["Otros EVP Agua_Privadas"].ventas_2025 += raw[sub].ventas_2025;
    raw["Otros EVP Agua_Privadas"].volumen_2026 += raw[sub].volumen_2026;
    raw["Otros EVP Agua_Privadas"].ventas_2026 += raw[sub].ventas_2026;
    raw["Otros EVP Agua_Privadas"].volumen_ppto += raw[sub].volumen_ppto;
    raw["Otros EVP Agua_Privadas"].ventas_ppto += raw[sub].ventas_ppto;
  });

  // 4. Marcas Privadas = "16 Oz_Privadas" + "1.5 LTS_Privadas" + "Otros EVP Agua_Privadas"
  raw["Marcas Privadas"] = initCategory();
  const privItems = ["16 Oz_Privadas", "1.5 LTS_Privadas", "Otros EVP Agua_Privadas"];
  privItems.forEach(item => {
    raw["Marcas Privadas"].volumen_2025 += raw[item].volumen_2025;
    raw["Marcas Privadas"].ventas_2025 += raw[item].ventas_2025;
    raw["Marcas Privadas"].volumen_2026 += raw[item].volumen_2026;
    raw["Marcas Privadas"].ventas_2026 += raw[item].ventas_2026;
    raw["Marcas Privadas"].volumen_ppto += raw[item].volumen_ppto;
    raw["Marcas Privadas"].ventas_ppto += raw[item].ventas_ppto;
  });

  // 6. Bebidas = Agua Saborizada + Hidractive+
  raw["Bebidas"] = initCategory();
  const bebItems = ["Agua Saborizada", "Hidractive+"];
  bebItems.forEach(item => {
    raw["Bebidas"].volumen_2025 += raw[item].volumen_2025;
    raw["Bebidas"].ventas_2025 += raw[item].ventas_2025;
    raw["Bebidas"].volumen_2026 += raw[item].volumen_2026;
    raw["Bebidas"].ventas_2026 += raw[item].ventas_2026;
    raw["Bebidas"].volumen_ppto += raw[item].volumen_ppto;
    raw["Bebidas"].ventas_ppto += raw[item].ventas_ppto;
  });

  // 5. EVP = Marcas Propias + Marcas Privadas + Bebidas
  raw["EVP"] = initCategory();
  const evpItems = ["Marcas Propias", "Marcas Privadas", "Bebidas"];
  evpItems.forEach(item => {
    raw["EVP"].volumen_2025 += raw[item].volumen_2025;
    raw["EVP"].ventas_2025 += raw[item].ventas_2025;
    raw["EVP"].volumen_2026 += raw[item].volumen_2026;
    raw["EVP"].ventas_2026 += raw[item].ventas_2026;
    raw["EVP"].volumen_ppto += raw[item].volumen_ppto;
    raw["EVP"].ventas_ppto += raw[item].ventas_ppto;
  });

  // 7. TOTAL = BT5 + EVP + BON (Bebidas ya está incluido en EVP)
  raw["TOTAL"] = initCategory();
  const totalItems = ["BT5", "EVP", "BON"];
  totalItems.forEach(item => {
    raw["TOTAL"].volumen_2025 += raw[item].volumen_2025;
    raw["TOTAL"].ventas_2025 += raw[item].ventas_2025;
    raw["TOTAL"].volumen_2026 += raw[item].volumen_2026;
    raw["TOTAL"].ventas_2026 += raw[item].ventas_2026;
    raw["TOTAL"].volumen_ppto += raw[item].volumen_ppto;
    raw["TOTAL"].ventas_ppto += raw[item].ventas_ppto;
  });

  // CÁLCULO DE MÉTRICAS COMPLEMENTARIAS (MÉTRICAS EN MILLONES, PRECIOS UNITARIOS Y VARIACIONES)
  const finalResults = {};
  const safeDivVar = (num, den) => (den ? (num / den) - 1 : 0);

  allKeys.forEach(cat => {
    const data = raw[cat];

    // Expresamos obligatoriamente las Ventas Netas en Millones de DOP (mDOP)
    const v2025_mDOP = data.ventas_2025 / 1000000;
    const v2026_mDOP = data.ventas_2026 / 1000000;
    const vppto_mDOP = data.ventas_ppto / 1000000;

    // Fórmula del precio unitario solicitada: (Ventas Netas_mDOP * 1000000) / Volumen
    const precio_2025 = data.volumen_2025 ? (v2025_mDOP * 1000000) / data.volumen_2025 : 0;
    const precio_2026 = data.volumen_2026 ? (v2026_mDOP * 1000000) / data.volumen_2026 : 0;
    const precio_ppto = data.volumen_ppto ? (vppto_mDOP * 1000000) / data.volumen_ppto : 0;

    // Cálculo impecable de variaciones porcentuales (con manejo de división por cero)
    const var_yoy_vol = safeDivVar(data.volumen_2026, data.volumen_2025);
    const var_ppto_vol = safeDivVar(data.volumen_2026, data.volumen_ppto);

    const var_yoy_ventas = safeDivVar(v2026_mDOP, v2025_mDOP);
    const var_ppto_ventas = safeDivVar(v2026_mDOP, vppto_mDOP);

    finalResults[cat] = {
      Volumen: {
        Y2025: Math.round(data.volumen_2025),
        Y2026: Math.round(data.volumen_2026),
        PPTO: Math.round(data.volumen_ppto),
        Var_YoY: Number(var_yoy_vol.toFixed(4)),
        Var_PPTO: Number(var_ppto_vol.toFixed(4))
      },
      Ventas: {
        Y2025: Number(v2025_mDOP.toFixed(4)),
        Y2026: Number(v2026_mDOP.toFixed(4)),
        PPTO: Number(vppto_mDOP.toFixed(4)),
        Var_YoY: Number(var_yoy_ventas.toFixed(4)),
        Var_PPTO: Number(var_ppto_ventas.toFixed(4))
      },
      Precio: {
        Y2025: Number(precio_2025.toFixed(2)),
        Y2026: Number(precio_2026.toFixed(2)),
        PPTO: Number(precio_ppto.toFixed(2))
      }
    };
  });

  console.log('✅ [processResumenCFData] Jerarquía consolidada exitosamente:', finalResults);
  return finalResults;
}
