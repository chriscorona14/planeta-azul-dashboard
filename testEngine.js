import * as XLSX from 'xlsx';
import { financialEngine } from './financialEngine.js';

let wb = XLSX.utils.book_new();
let ws = XLSX.utils.aoa_to_sheet([["Ingresos", 100], ["Ebitda", 20], ["Utilidad Neta", 10]]);
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

try {
    let result = financialEngine(wb);
    console.log("Success:", !!result);
} catch (e) {
    console.error("Worker error:", e);
}
