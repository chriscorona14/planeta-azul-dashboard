import fs from 'fs';

let content = fs.readFileSync('main.js', 'utf8');

const regexToReplace = /if \(window\.globalFinancialData && window\.globalFinancialData\.length > 0\) \{[\s\S]*?console\.error\("Error saving updated CXP to indexedDB:", e\);\n\s*\}\n\s*\} else \{/g;

content = content.replace(regexToReplace, `if (window.globalFinancialData && window.globalFinancialData.length > 0) {
            window.cachedStandaloneCxp = { indices, cxpRows, rows };
            await window.applyCachedStandaloneCxp();
        } else {`);

fs.writeFileSync('main.js', content);
console.log('Fixed cxpWorkbook!');
