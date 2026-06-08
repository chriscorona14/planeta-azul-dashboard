const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');

// Replace the pptx initialization and slide adding in bindPptxExport
const oldInit = `                if (!pptx) {
                    pptx = new pptxgen();
                    pptx.layout = 'LAYOUT_16x9';
                }
                let slide = pptx.addSlide();
                slide.addImage({ data: imgData, x: 0, y: 0, w: '100%', h: '100%' });`;

const newInit = `
                if (!pptx) {
                    pptx = new pptxgen();
                    pptx.layout = 'LAYOUT_16x9';
                    
                    // Define Master Slide
                    pptx.defineSlideMaster({
                        title: 'MASTER_SLIDE',
                        background: { fill: 'F0F8FF' }, // Light Alice Blue
                        objects: [
                            { rect: { x: 0, y: 0, w: '100%', h: '0.2', fill: '0096c7' } }, // Top stripe
                            { circle: { x: -0.5, y: -0.5, w: 3, h: 3, fill: { color: '0096c7', transparency: 90 } } },
                            { circle: { x: 8, y: -1, w: 4, h: 4, fill: { color: '0096c7', transparency: 95 } } },
                            { circle: { x: 8.5, y: 4.5, w: 3, h: 3, fill: { color: '0096c7', transparency: 90 } } }
                        ]
                    });

                    // Add Cover Slide
                    let cover = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
                    cover.addText('Planeta Azul', { x: 0.5, y: 2.0, w: 8, fontSize: 32, bold: true, color: '000000', fontFace: 'Segoe UI' });
                    
                    const sel = document.getElementById('monthSelector');
                    let mesStr = "Actuales";
                    if (sel && sel.options && sel.options[sel.selectedIndex]) {
                        mesStr = sel.options[sel.selectedIndex].text;
                    }
                    cover.addText('Resultados ' + mesStr, { x: 0.5, y: 2.6, w: 8, fontSize: 44, bold: true, color: '005b96', fontFace: 'Segoe UI' });
                    cover.addText('Comité Financiero', { x: 0.5, y: 3.5, w: 8, fontSize: 28, bold: true, color: '000000', fontFace: 'Segoe UI' });
                    
                    const currentDt = new Date();
                    const dtOpts = { month:'long', year:'numeric' };
                    let dtStr = currentDt.getDate() + ' de ' + currentDt.toLocaleDateString('es-ES', dtOpts);
                    dtStr = dtStr.charAt(0).toUpperCase() + dtStr.slice(1);
                    cover.addText(dtStr, { x: 0.5, y: 4.8, w: 8, fontSize: 14, color: '000000', fontFace: 'Segoe UI' });
                }

                let slide = pptx.addSlide({ masterName: 'MASTER_SLIDE' });
                // Add the slide title natively in PPTX to look professional
                slide.addText(forcedSubtitle, { 
                    x: 0.5, y: 0.3, w: 9, h: 0.5, 
                    fontSize: 24, bold: true, color: '005b96', fontFace: 'Segoe UI' 
                });

                // Add the captured image, but slightly smaller to let background and title show
                // We'll calculate a good fit, mostly centering it below the title
                slide.addImage({ data: imgData, x: 0.2, y: 0.9, w: 9.6, h: 4.5, sizing: { type: 'contain', w: 9.6, h: 4.5 } });
`;

// we also want to hide the HTML title and layout wrappers during PPTX capture so it's not rendered twice.
const pptxFuncStart = "const bindPptxExport = (btn) => {";
const pStartIndex = content.indexOf(pptxFuncStart);
if (pStartIndex === -1) {
    console.log("bindPptxExport not found");
    process.exit(1);
}

// Just substring it, replace oldInit, and replace "const bgColor = \"#ffffff\";" with "const bgColor = \"#F0F8FF\";"
let pptxBlock = content.substring(pStartIndex);
pptxBlock = pptxBlock.replace(oldInit, newInit);
pptxBlock = pptxBlock.replace('const bgColor = "#ffffff";', 'const bgColor = "#F0F8FF";');

// Let's also hide the titleLabel inside html2canvas
pptxBlock = pptxBlock.replace('let origTitleHTML = "";\n                if (titleLabel && forcedSubtitle) {\n                    origTitleHTML = titleLabel.innerHTML;\n                    titleLabel.innerHTML += ` <span style="font-size: 0.7em; color: white; background-color: var(--primary); padding: 4px 8px; border-radius: 8px; margin-left: 10px; vertical-align: middle;">${forcedSubtitle}</span>`;\n                }', 'let origTitleHTML = "";\n                let origTitleDisplay = "";\n                if (titleLabel) {\n                    origTitleHTML = titleLabel.innerHTML;\n                    origTitleDisplay = titleLabel.style.display;\n                    titleLabel.style.display = "none"; // Hide html title for pptx export\n                }');

pptxBlock = pptxBlock.replace('if (titleLabel && forcedSubtitle) titleLabel.innerHTML = origTitleHTML;', 'if (titleLabel) {\n                    titleLabel.innerHTML = origTitleHTML;\n                    titleLabel.style.display = origTitleDisplay;\n                }');

content = content.substring(0, pStartIndex) + pptxBlock;

fs.writeFileSync('main.js', content);
console.log("PPTX enhanced!");
