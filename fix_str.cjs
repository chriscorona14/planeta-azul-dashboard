const fs = require('fs');
let s = fs.readFileSync('main.js', 'utf8');
s = s.replace("const isNegative = str.startsWith(''(') || str.startsWith('(');", "const isNegative = str.startsWith(\"\'(\") || str.startsWith('(');");
fs.writeFileSync('main.js', s);
