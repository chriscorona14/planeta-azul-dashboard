const fs = require('fs');
let code = fs.readFileSync('resumenComercialEngine.js', 'utf8');

const replacement = `const cleanA = (str) => String(str || '').toUpperCase().replace(/\\s+/g, '');

const ARBOL_COMERCIAL = [
  { 
    id: 'BT5_Total', 
    label: 'BOTELLÓN', 
    type: 'main', 
    matcher: (r) => { 
      const ag = cleanA(r.agrupacion); 
      return ag === 'BOTELLÓN' || ag === 'BOTELLON' || ag === 'BT5'; 
    } 
  },
  
  { id: 'BOTELLAS_Total', label: 'BOTELLAS', type: 'main' },
    { 
      id: 'APA_BOTELLA_0_5_LTS_x20', 
      parent: 'BOTELLAS_Total', 
      label: 'APA BOTELLA 0.5 LTS (x20)', 
      type: 'item',
      matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('APABOTELLA0.5LTS(X20)') || a === 'APABOTELLA0.5(X20)'; }
    },
    { 
      id: 'APA_BOTELLA_1_5_LTS_x12', 
      parent: 'BOTELLAS_Total', 
      label: 'APA BOTELLA 1.5 LTS (x12)', 
      type: 'item',
      matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('APABOTELLA1.5LTS(X12)') || a === 'APABOTELLA1.5(X12)'; }
    },
    { id: 'APA_OTROS', parent: 'BOTELLAS_Total', label: 'APA OTROS', type: 'sub' },
      { 
        id: 'APA_100_RPET_0_5LTS_x12', 
        parent: 'APA_OTROS', 
        label: 'APA 100% RPET 0.5LTS (x12)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('APA100%RPET0.5LTS(X12)') || a.includes('APA100%RPET0.5(X12)'); }
      },
      { 
        id: 'APA_BOTELLA_0_5LTS_x12', 
        parent: 'APA_OTROS', 
        label: 'APA BOTELLA 0.5LTS (x12)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('APABOTELLA0.5LTS(X12)') || a === 'APABOTELLA0.5(X12)'; }
      },
      { 
        id: 'APA_BOTELLA_5LTS_x4', 
        parent: 'APA_OTROS', 
        label: 'APA BOTELLA 5LTS (x4)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return a === 'APABOTELLA5LTS(X4)' || a === 'APABOTELLA5(X4)'; }
      },
      { 
        id: 'APA_BOTELLA_8_VASOS_1_89_LTS_x6', 
        parent: 'APA_OTROS', 
        label: 'APA BOTELLA 8 VASOS 1.89 LTS (x6)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('APABOTELLA8VASOS1.89LTS(X6)') || a.includes('APABOTELLA8VASOS1.89(X6)'); }
      },
      { 
        id: 'APA_SPORT_0_71_LTS_x12', 
        parent: 'APA_OTROS', 
        label: 'APA SPORT 0.71 LTS (x12)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('APASPORT0.71LTS(X12)') || a.includes('APASPORT0.71(X12)'); }
      },
      { 
        id: 'APA_TETRA_PACK_0_5_LTS_x18', 
        parent: 'APA_OTROS', 
        label: 'APA TETRA PACK 0.5 LTS (x18)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('APATETRAPACK0.5LTS(X18)') || a.includes('APATETRAPACK0.5(X18)') || a.includes('APATETRA0.5'); }
      },
      
    { 
      id: 'MAQUILA_AGUA_0_5_LTS_x20', 
      parent: 'BOTELLAS_Total', 
      label: 'MAQUILA AGUA 0.5 LTS (x20)', 
      type: 'sub',
      matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('MAQUILAAGUA0.5LTS(X20)') || a.includes('MAQUILA0.5LTS(X20)'); }
    },
    { 
      id: 'MAQUILA_AGUA_1_5_L_TS_x12', 
      parent: 'BOTELLAS_Total', 
      label: 'MAQUILA AGUA 1.5 L TS (x12)', 
      type: 'sub',
      matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('MAQUILAAGUA1.5L_TS(X12)') || a.includes('MAQUILAAGUA1.5LTS(X12)') || a.includes('MAQUILA1.5LTS(X12)') || a.includes('MAQUILAAGUA1.5LTS'); }
    },
    { id: 'MAQUILA_OTROS', parent: 'BOTELLAS_Total', label: 'MAQUILA OTROS', type: 'sub' },
      { 
        id: 'MAQUILA_100_RPET_0_5LTS_x12', 
        parent: 'MAQUILA_OTROS', 
        label: '100% RPET 0.5LTS (x12)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return (a.includes('100%RPET0.5LTS(X12)') || a.includes('RPET0.5LTS(X12)')) && !a.includes('APA100%RPET'); }
      },
      { 
        id: 'MAQUILA_BOTELLA_5LTS_x4', 
        parent: 'MAQUILA_OTROS', 
        label: 'BOTELLA 5LTS (x4)', 
        type: 'item',
        matcher: (r) => { const a = cleanA(r.agrupacion); return (a === 'BOTELLA5LTS(X4)' || a === 'BOTELLA5(X4)') && !a.includes('APA'); }
      },
      
  { id: 'BEBIDAS_Total', parent: 'BOTELLAS_Total', label: 'BEBIDAS', type: 'sub' },
    { 
      id: 'PA_SABOR_0_5_LTS_x12', 
      parent: 'BEBIDAS_Total', 
      label: 'PA SABOR 0.5 LTS (x12)', 
      type: 'item',
      matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('PASABOR0.5LTS(X12)') || a.includes('PASABOR0.5(X12)'); }
    },
    { 
      id: 'HIDRACTIVE_PLUS', 
      parent: 'BEBIDAS_Total', 
      label: 'HIDRACTIVE +', 
      type: 'item',
      matcher: (r) => { const a = cleanA(r.agrupacion); return a.includes('HIDRACTIVE+') || a === 'HIDRACTIVE'; }
    },
    
  { 
    id: 'BON_Total', 
    label: 'BON', 
    type: 'main',
    matcher: (r) => { const a = cleanA(r.agrupacion); return a === 'BON' || isBonRow(r); }
  }
];`

// Reemplazar la declaración vieja
let match = code.replace(/const ARBOL_COMERCIAL = \[\s*\{[\s\S]*?\}\s*\];/m, replacement);
fs.writeFileSync('resumenComercialEngine.js', match);
