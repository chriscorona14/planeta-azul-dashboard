import * as XLSX from 'xlsx';
import { financialEngine } from './financialEngine.js';

self.onmessage = function(e) {
    try {
        const buffer = e.data.buffer;
        
        // 1. Parsing the Excel file
        self.postMessage({ type: 'progress', progress: 50, message: "Decodificando archivo Excel en segundo plano..." });
        let workbook = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
        
        // 2. Processing data
        self.postMessage({ type: 'progress', progress: 75, message: "Ejecutando motor de datos financieros..." });
        let engineResult = financialEngine(workbook);
        
        workbook = null; // Free up memory

        if (engineResult.error || !engineResult.data || engineResult.data.length === 0) {
            self.postMessage({ 
                type: 'error', 
                error: engineResult.error || "No se pudieron extraer datos numéricos del archivo." 
            });
            return;
        }

        self.postMessage({ 
            type: 'done', 
            engineResult 
        });
    } catch (err) {
        self.postMessage({ type: 'error', error: err.message || "Ocurrió un error en el worker." });
    }
};
