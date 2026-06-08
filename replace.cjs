const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');
content = content.replace(/path: 'cover_bg.png'/g, "path: '/cover_bg.png'");
content = content.replace(/path: 'slide_bg.png'/g, "path: '/slide_bg.png'");
fs.writeFileSync('main.js', content);
