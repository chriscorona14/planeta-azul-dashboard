const fs = require('fs');

let content = fs.readFileSync('main.js', 'utf8');

// First inject the helper at the top, right after imports or strict mode if any, or just at the top
const helper = `\nasync function getFinanceDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open('PlanetaAzulDB', 4);
        req.onupgradeneeded = (e) => {
            if (!e.target.result.objectStoreNames.contains('finance_cache')) {
                e.target.result.createObjectStore('finance_cache');
            }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}\n`;

if (!content.includes('async function getFinanceDB()')) {
    content = helper + content;
}

// Replace the verbose blocks
// const db = await new Promise((resolve, reject) => { const req = indexedDB.open('PlanetaAzulDB', 3); ... });
// Since they vary slightly in formatting, regex replace is best

content = content.replace(/const db = await new Promise[^]*?indexedDB\.open\('PlanetaAzulDB', 3\)[^]*?\}\);/g, "const db = await getFinanceDB();");

content = content.replace(/const db = await new Promise\(r => \{ const req = indexedDB\.open\('PlanetaAzulDB', 3\); req\.onsuccess = \(\) => r\(req\.result\); \}\);/g, "const db = await getFinanceDB();");

fs.writeFileSync('main.js', content);
console.log("Updated main.js with IndexedDB fix.");
