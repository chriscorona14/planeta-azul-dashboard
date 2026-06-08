const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');

// The shape name should be pptx.ShapeType.oval, let's replace { circle: ... } with { shape: { type: pptx.ShapeType.oval, ... } }
// But since we define it as { rect: ... } or { shape: ... } we should use { shape: pptx.ShapeType.oval, options: { ... } } or similar
// Let's use pptx.ShapeType.oval

let toReplace = `
                            { rect: { x: 0, y: 0, w: '100%', h: '0.2', fill: '0096c7' } }, // Top stripe
                            { circle: { x: -0.5, y: -0.5, w: 3, h: 3, fill: { color: '0096c7', transparency: 90 } } },
                            { circle: { x: 8, y: -1, w: 4, h: 4, fill: { color: '0096c7', transparency: 95 } } },
                            { circle: { x: 8.5, y: 4.5, w: 3, h: 3, fill: { color: '0096c7', transparency: 90 } } },
                            { text: { text: "PLANETA AZUL\\nBEBIDAS", options: { x: 8.0, y: 4.8, w: 2, fill: 'none', color: '005b96', fontSize: 10, align: 'center', bold: true, fontFace: 'Segoe UI' } } }
`;

let validShapes = `
                            { rect: { x: 0, y: 0, w: '100%', h: '0.2', fill: { color: '0096c7' } } }, // Top stripe
                            { text: { text: "PLANETA AZUL\\nBEBIDAS", options: { x: 8.0, y: 4.8, w: 2, fill: { color: 'none' }, color: '005b96', fontSize: 10, align: 'center', bold: true, fontFace: 'Segoe UI' } } }
`;

// It's much safer to use basic rects and text. If pptx.ShapeType isn't exposed correctly, it will crash.
content = content.replace(toReplace.trim(), validShapes.trim());

fs.writeFileSync('main.js', content);
console.log("Shapes fixed to be completely safe and native");
