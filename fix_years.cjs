const fs = require('fs');

let content = fs.readFileSync('main.js', 'utf8');

content = content.replace(/visibleMonths = visibleMonths\.filter\(m => \!isYear2025\(m\)\);/g, 'visibleMonths = visibleMonths.filter(m => isYear2026(m));');
content = content.replace(/const visibleMonths = data\.slice\(startIdx, endIdx \+ 1\)\.filter\(d => \!isYear2025\(d\)\);/g, 'const visibleMonths = data.slice(startIdx, endIdx + 1).filter(d => isYear2026(d));');

// And the continue conditions for YTD logic:
content = content.replace(/if \(isYear2025\(periodData\)\) continue;/g, 'if (!isYear2026(periodData)) continue;');
content = content.replace(/if \(isYear2025\(item\)\) continue;/g, 'if (!isYear2026(item)) continue;');
content = content.replace(/if \(isYear2025\(data\[k\]\)\) continue;/g, 'if (!isYear2026(data[k])) continue;');

fs.writeFileSync('main.js', content);
