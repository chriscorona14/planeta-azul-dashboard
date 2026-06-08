const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');

// Looking for the master slide objects to add the logo text
let insertIdx = content.indexOf("{ circle: { x: 8.5, y: 4.5, w: 3, h: 3, fill: { color: '0096c7', transparency: 90 } } }");

if (insertIdx !== -1) {
    let before = content.substring(0, insertIdx);
    let after = content.substring(insertIdx);
    
    let replacement = `
                            { circle: { x: 8.5, y: 4.5, w: 3, h: 3, fill: { color: '0096c7', transparency: 90 } } },
                            { text: { text: "PLANETA AZUL\\nBEBIDAS", options: { x: 8.0, y: 4.8, w: 2, fill: 'none', color: '005b96', fontSize: 10, align: 'center', bold: true, fontFace: 'Segoe UI' } } }
    `;
    
    after = after.replace("{ circle: { x: 8.5, y: 4.5, w: 3, h: 3, fill: { color: '0096c7', transparency: 90 } } }", replacement.trim());
    
    fs.writeFileSync('main.js', before + after);
    console.log("Logo text added to PPTX template.");
} else {
    console.log("Could not find insertion point for PPTX logo");
}

// We should also make sure the PptxGenJS CDN works when imported as an ES module or via window.
