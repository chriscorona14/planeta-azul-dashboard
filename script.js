const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');

// replace the !isYear2025 filter with isYear2026
content = content.replace(/!isYear2025\((\w+)\)/g, 'isYear2026($1)');
content = content.replace(/!isYear2025\(item\.d\)/g, 'isYear2026(item.d)');
content = content.replace(/!isYear2025\(data\[k\]\)/g, 'isYear2026(data[k])');

fs.writeFileSync('main.js', content);
