const fs = require('fs');
let content = fs.readFileSync('main.js', 'utf8');

const oldStr = `                // Ensure the capture window is constrained to a standard desktop width
                const desktopWidth = 1440;
                
                const canvas = await html2canvas(contentToRender, {
                    scale: 2, 
                    useCORS: true,
                    logging: false,
                    backgroundColor: bgColor,
                    windowWidth: desktopWidth,
                    onclone: (clonedDoc) => {
                         // Fix the animation veil: animations mid-flight cause elements to have 50% opacity
                         const style = clonedDoc.createElement('style');
                         style.innerHTML = \`
                            * { 
                                animation: none !important; 
                                transition: none !important; 
                            }
                            .main-container, .layout-wrapper {
                                width: \${desktopWidth}px !important;
                                min-width: \${desktopWidth}px !important;
                                max-width: \${desktopWidth}px !important;
                                height: auto !important;
                                max-height: none !important;
                                overflow: visible !important;
                            }
                            .view-container.active {
                                width: 100% !important;
                            }
                            .card, .section-table, .pnl-detail-table, .table-container, .chart-box {
                                height: auto !important;
                                max-height: none !important;
                                overflow: visible !important;
                            }
                            table {
                                width: 100% !important; 
                                max-width: none !important;
                            }
                         \`;
                         clonedDoc.head.appendChild(style);

                         const mainCont = clonedDoc.querySelector('.main-container');
                         if (mainCont) {
                             mainCont.style.overflow = "visible";
                             mainCont.style.height = "max-content";
                         }
                    }
                });

                // RESTORE everything
                if (titleLabel && forcedSubtitle) {
                    titleLabel.innerHTML = origTitleHTML;
                }
                if (header) header.style.position = origHeaderPos;
                contentToRender.style.overflow = originalOverflow;
                contentToRender.style.height = originalHeight;
                if (layoutWrapper) {
                    layoutWrapper.style.height = layoutWrapperOrigHeight;
                    layoutWrapper.style.overflow = layoutWrapperOrigOverflow;
                }
                if (headerActions) headerActions.style.display = originalHeaderDisplay;
                if (pnlControls) pnlControls.style.display = pnlControlsDisplay;

                const styleEl = document.getElementById('pdf-expand-style');
                if (styleEl) styleEl.remove();

                const imgData = canvas.toDataURL('image/jpeg', 0.95);
                
                if (!pptx) {
                    const PptxGen = typeof pptxgen !== "undefined" ? pptxgen : (typeof PptxGenJS !== "undefined" ? PptxGenJS : window.PptxGenJS);
                    if (!PptxGen) {
                        throw new Error("PptxGenJS library is not loaded.");
                    }
                    pptx = new PptxGen();
                    pptx.layout = 'LAYOUT_16x9';
                    
                    // Define Master Slide
                    pptx.defineSlideMaster({
                        title: 'MASTER_SLIDE',
                        background: { fill: 'F0F8FF' }, // Light Alice Blue
                        objects: [
                            { rect: { x: 0, y: 0, w: '100%', h: '0.2', fill: { color: '0096c7' } } }, // Top stripe
                            // Decorative light circles for water ripple effect
                            { shape: { type: (typeof pptx !== 'undefined' && pptx && pptx.ShapeType ? pptx.ShapeType.oval : 'oval'), options: { x: -0.5, y: -0.5, w: 3, h: 3, fill: { color: '87CEFA', transparency: 85 } } } },
                            { shape: { type: (typeof pptx !== 'undefined' && pptx && pptx.ShapeType ? pptx.ShapeType.oval : 'oval'), options: { x: 8, y: -1, w: 4, h: 4, fill: { color: '0096c7', transparency: 90 } } } },
                            { shape: { type: (typeof pptx !== 'undefined' && pptx && pptx.ShapeType ? pptx.ShapeType.oval : 'oval'), options: { x: -1, y: 3, w: 5, h: 5, fill: { color: '87CEFA', transparency: 92 } } } },
                            { text: { text: "PLANETA AZUL\\nBEBIDAS", options: { x: 8.0, y: 4.8, w: 2, fill: { color: 'none' }, color: '005b96', fontSize: 10, align: 'center', bold: true, fontFace: 'Segoe UI' } } }
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
                slide.addImage({ data: imgData, x: 0.2, y: 0.2, w: 9.6, h: 5.2, sizing: { type: 'contain', w: 9.6, h: 5.2 } });`;

const newStr = `                // Ensure the capture window is constrained to a standard desktop width
                const desktopWidth = Math.max(1440, contentToRender.scrollWidth || 1440);
                
                const canvas = await html2canvas(contentToRender, {
                    scale: 2, 
                    useCORS: true,
                    logging: false,
                    backgroundColor: bgColor,
                    windowWidth: desktopWidth,
                    onclone: (clonedDoc) => {
                         // Fix the animation veil: animations mid-flight cause elements to have 50% opacity
                         const style = clonedDoc.createElement('style');
                         style.innerHTML = \`
                            * { 
                                animation: none !important; 
                                transition: none !important; 
                            }
                            .main-container, .layout-wrapper {
                                width: \${desktopWidth}px !important;
                                min-width: \${desktopWidth}px !important;
                                max-width: \${desktopWidth}px !important;
                                height: auto !important;
                                max-height: none !important;
                                overflow: visible !important;
                            }
                            .view-container.active {
                                width: 100% !important;
                            }
                            .card, .section-table, .pnl-detail-table, .table-container, .chart-box {
                                height: auto !important;
                                max-height: none !important;
                                overflow: visible !important;
                            }
                            table {
                                width: 100% !important; 
                                max-width: none !important;
                            }
                         \`;
                         clonedDoc.head.appendChild(style);

                         const mainCont = clonedDoc.querySelector('.main-container');
                         if (mainCont) {
                             mainCont.style.overflow = "visible";
                             mainCont.style.height = "max-content";
                         }
                    }
                });

                // RESTORE everything
                if (titleLabel && forcedSubtitle) {
                    titleLabel.innerHTML = origTitleHTML;
                }
                if (header) header.style.position = origHeaderPos;
                contentToRender.style.overflow = originalOverflow;
                contentToRender.style.height = originalHeight;
                if (layoutWrapper) {
                    layoutWrapper.style.height = layoutWrapperOrigHeight;
                    layoutWrapper.style.overflow = layoutWrapperOrigOverflow;
                }
                if (headerActions) headerActions.style.display = originalHeaderDisplay;
                if (pnlControls) pnlControls.style.display = pnlControlsDisplay;

                const styleEl = document.getElementById('pdf-expand-style');
                if (styleEl) styleEl.remove();

                const imgData = canvas.toDataURL('image/jpeg', 0.95);
                
                if (!pptx) {
                    const PptxGen = typeof pptxgen !== "undefined" ? pptxgen : (typeof PptxGenJS !== "undefined" ? PptxGenJS : window.PptxGenJS);
                    if (!PptxGen) {
                        throw new Error("PptxGenJS library is not loaded.");
                    }
                    pptx = new PptxGen();
                    pptx.layout = 'LAYOUT_16x9';
                    
                    const ovalType = (typeof pptx.ShapeType !== 'undefined' && pptx.ShapeType.oval) ? pptx.ShapeType.oval : 'oval';
                    
                    // Define Master Slide
                    pptx.defineSlideMaster({
                        title: 'MASTER_SLIDE',
                        background: { fill: 'FFFFFF' }, // Clean white background for the content to rest on
                        objects: [
                            // Soft watermark shapes to mimic Planeta Azul background attached layout.
                            { shape: { type: ovalType, options: { x: -2, y: -2, w: 5, h: 5, fill: { color: 'e0f2fe', transparency: 60 } } } },
                            { shape: { type: ovalType, options: { x: 8.5, y: -2, w: 6, h: 6, fill: { color: 'bae6fd', transparency: 70 } } } },
                            { shape: { type: ovalType, options: { x: -1, y: 3.5, w: 6, h: 6, fill: { color: 'e0f2fe', transparency: 60 } } } },
                            { shape: { type: ovalType, options: { x: 8.5, y: 4, w: 3, h: 3, fill: { color: 'bae6fd', transparency: 65 } } } },
                            // Logo Text Bottom Right
                            { text: { text: "PLANETA AZUL\\nBEBIDAS", options: { x: 8.0, y: 4.8, w: 2, fill: { color: 'none' }, color: '005b96', fontSize: 10, align: 'center', bold: true, fontFace: 'Segoe UI' } } }
                        ]
                    });

                    // Define Cover Slide
                    pptx.defineSlideMaster({
                        title: 'COVER_SLIDE',
                        background: { fill: 'F4F9FD' }, 
                        objects: [
                            { shape: { type: ovalType, options: { x: 4, y: -4, w: 10, h: 10, fill: { color: 'e0f2fe', transparency: 50 } } } },
                            { shape: { type: ovalType, options: { x: -2, y: 3, w: 6, h: 6, fill: { color: 'bae6fd', transparency: 60 } } } },
                            // Logo Text Bottom Right
                            { text: { text: "PLANETA AZUL\\nBEBIDAS", options: { x: 8.0, y: 4.8, w: 2, fill: { color: 'none' }, color: '005b96', fontSize: 10, align: 'center', bold: true, fontFace: 'Segoe UI' } } }
                        ]
                    });

                    // Add Cover Slide
                    let cover = pptx.addSlide({ masterName: 'COVER_SLIDE' });
                    cover.addText('Planeta Azul', { x: 0.5, y: 2.0, w: 8, fontSize: 36, bold: true, color: '000000', fontFace: 'Segoe UI' });
                    
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
                
                // Pure aspect ratio based scaling:
                const slideW = 10;
                const slideH = 5.625;
                const ratio = canvas.height / canvas.width;
                
                let imgW = slideW - 0.4;
                let imgH = imgW * ratio;
                if (imgH > (slideH - 0.4)) {
                    imgH = slideH - 0.4;
                    imgW = imgH / ratio;
                }
                
                const imgX = (slideW - imgW) / 2;
                const imgY = (slideH - imgH) / 2;
                
                slide.addImage({ data: imgData, x: imgX, y: imgY, w: imgW, h: imgH });`;

if (content.indexOf(oldStr) !== -1) {
    fs.writeFileSync('main.js', content.replace(oldStr, newStr));
    console.log("Replaced block perfectly");
} else {
    console.error("Block not found");
}
