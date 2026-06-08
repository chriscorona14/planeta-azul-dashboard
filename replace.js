const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');
content = content.replace(/\.main-container, \.layout-wrapper \{/g, `html, body {
                        height: auto !important;
                        min-height: 100% !important;
                        overflow: visible !important;
                    }
                    .main-container, .layout-wrapper {`);
fs.writeFileSync('main.js', content);
