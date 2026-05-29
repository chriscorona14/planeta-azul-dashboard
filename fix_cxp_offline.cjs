const fs = require('fs');
let mainJs = fs.readFileSync('main.js', 'utf8');

const sIdx = mainJs.indexOf("const req = db.transaction('finance_cache', 'readonly').objectStore('finance_cache').get(CACHE_KEY);");
const eIdx = mainJs.indexOf("const CACHE_CEO_KEY = 'CEO_VENTAS_KEY';");

if (sIdx === -1 || eIdx === -1) {
    console.log("Could not find insertion points");
    process.exit(1);
}

const insertionStr = `const cxpRecord = await new Promise((resolve) => {
            const req = db.transaction('finance_cache', 'readonly').objectStore('finance_cache').get('MASTER_STANDALONE_CXP_DATA');
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => resolve(null);
        });

        if (cxpRecord && cxpRecord.data) {
            window.cxpStandaloneData = cxpRecord.data;
            window.hasCxpAccess = true;
        }

        `;
        
mainJs = mainJs.slice(0, eIdx) + insertionStr + mainJs.slice(eIdx);
fs.writeFileSync('main.js', mainJs);
console.log("Injected offline loader for CXP");
