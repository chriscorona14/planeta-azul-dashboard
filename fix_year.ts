import fs from 'fs';

let content = fs.readFileSync('main.js', 'utf8');

// Replace standard identifier patterns
content = content.replace(/([a-zA-Z0-9_]+)\.sortDate\.getFullYear\(\)/g, 'getSortYear($1)');

// Replace array access patterns like data[endIdx].sortDate.getFullYear()
content = content.replace(/([a-zA-Z0-9_]+\[[a-zA-Z0-9_]+\])\.sortDate\.getFullYear\(\)/g, 'getSortYear($1)');

// Make sure sortDate check and getSortYear don't do weird logic.
// Actually, since getSortYear handles sortDate checking, things like `d.sortDate && getSortYear(d) === 2025` are perfectly fine.

fs.writeFileSync('main.js', content);
console.log('Fixed years!');
