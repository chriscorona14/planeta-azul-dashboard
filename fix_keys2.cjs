const fs = require('fs');
let code = fs.readFileSync('main.js', 'utf8');
code = code.replace(/CEO_VENTAS_KEY_V3/g, 'CEO_VENTAS_KEY_V4');
fs.writeFileSync('main.js', code);
