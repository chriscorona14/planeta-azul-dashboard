const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');

// We want to duplicate bindPdfExport into bindPptxExport.
const startStr = "const bindPdfExport = (btn) => {";
const endStr = "    bindPdfExport(btnExportPDF);";

const startIndex = content.indexOf(startStr);
const endIndex = content.indexOf(endStr) + endStr.length;

let pdfFunc = content.substring(startIndex, endIndex);

let pptxFunc = pdfFunc.replace("const bindPdfExport = (btn) => {", "const bindPptxExport = (btn) => {")
                      .replace("btn.innerHTML = '<i data-lucide=\"loader\" class=\"spin-icon\" style=\"width: 16px; height: 16px; display: inline-block; vertical-align: middle; margin-right: 4px;\"></i> Generando Master PDF...';",
                               "btn.innerHTML = '<i data-lucide=\"loader\" class=\"spin-icon\" style=\"width: 16px; height: 16px; display: inline-block; vertical-align: middle; margin-right: 4px;\"></i> Generando Master PPTX...';")
                      .replace("let pdf = null;", "let pptx = null;")
                      .replace("const addPageToPDF", "const addPageToPPTX")
                      .replace(/addPageToPDF/g, "addPageToPPTX")
                      .replace("if (!pdf) {\\n                    pdf = new jsPDF({ orientation: orientation, unit: 'mm', format: pageFormat });\\n                    pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, 'MEDIUM');\\n                } else {\\n                    pdf.addPage(pageFormat, orientation);\\n                    pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, 'MEDIUM');\\n                }",
                      `
                if (!pptx) {
                    pptx = new pptxgen();
                    // Optional: set layout to 16:9 for a professional look
                    pptx.layout = 'LAYOUT_16x9';
                }
                let slide = pptx.addSlide();
                // We add the captured image covering the slide
                slide.addImage({ data: imgData, x: 0, y: 0, w: '100%', h: '100%' });`)
                      .replace("pdf.save(`Reportes_Ejecutivos_Maestros_${dateStr}.pdf`);", "pptx.writeFile({ fileName: `Reportes_Ejecutivos_Maestros_${dateStr}.pptx` });")
                      .replace("Error generating Master PDF:", "Error generating Master PPTX:")
                      .replace("al generar el PDF:", "al generar el PPTX:")
                      .replace("bindPdfExport(btnExportPDF);", "const btnExportPPTX = document.getElementById('btn-export-pptx');\n    bindPptxExport(btnExportPPTX);");

// Make sure pdf stuff we replaced actually worked, we can use regexes
pptxFunc = pptxFunc.replace(/if \(!pdf\) \{[\s\S]*?\} else \{[\s\S]*?\}/, `                if (!pptx) {
                    pptx = new pptxgen();
                    pptx.layout = 'LAYOUT_16x9';
                }
                let slide = pptx.addSlide();
                slide.addImage({ data: imgData, x: 0, y: 0, w: '100%', h: '100%' });`);

// Insert pptxFunc right after pdfFunc
const newContent = content.substring(0, endIndex) + '\n\n    ' + pptxFunc + content.substring(endIndex);

fs.writeFileSync('main.js', newContent);
console.log("Appended pptx export code");
