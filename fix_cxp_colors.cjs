const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');

// Replace isProjected criteria
content = content.replace(/const isProjected = m\.sortDate && m\.sortDate\.getFullYear\(\) >= 2026 && Math\.round\(cxpTotal\) === 0;/g, 
"const isProjected = m.cxpDetail && m.cxpDetail.isProjectedDetail;");

// Replace header color
content = content.replace(/const bg = isProjected \? '#e08924' : 'var\(--sidebar\)';/g,
"const bg = isProjected ? '#64748b' : 'var(--sidebar)';"); // Slate 500

// Replace table cell background
content = content.replace(/'rgba\(224, 137, 36, 0\.1\)'/g, "'rgba(100, 116, 139, 0.08)'"); // Slate 500 with 0.08 opacity

// Replace total cell background
content = content.replace(/'rgba\(224, 137, 36, 0\.2\)'/g, "'rgba(100, 116, 139, 0.15)'"); // Slate 500 with 0.15 opacity

fs.writeFileSync('main.js', content);
console.log("Colors and logic updated in main.js");
