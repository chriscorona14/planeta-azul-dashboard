import * as XLSX from 'xlsx';
import { financialEngine } from './financialEngine.js';

let wb = XLSX.utils.book_new();

let balance = [
  ["Concept", "2025-12-01", "2026-04-01"],
  ["Deuda neta bancaria usd", 26.4, 26.9],
  ["4.0x", 2.5, 2.7],
  ["<= 2.0x", 1.7, 1.7],
  ["Capacidad de pago", 1.0, 0.8],
  ["Razon corriente", 0.5, 0.5],
  ["Caja y banco", 10, 10],
  ["Deuda financiera total", 100, 100],
  ["Ganancia acumulada", 50, 60]
];
let wsBalance = XLSX.utils.aoa_to_sheet(balance);
XLSX.utils.book_append_sheet(wb, wsBalance, "Balance Sheet mDOP");

let deuda = [
  ["Concept", "2025-12-01", "2026-04-01"],
  ["Promedio ponderado dop", 0.13, 0.13],
  ["Promedio ponderado usd", 0.05, 0.05],
  ["Deuda neta total usd", 35.4, 36.3]
];
let wsDeuda = XLSX.utils.aoa_to_sheet(deuda);
XLSX.utils.book_append_sheet(wb, wsDeuda, "Deuda");

let pnl = [
  ["Concept", "2025-12-01", "2026-04-01"],
  ["Ingresos", 1000, 1000],
  ["EBITDA", 500, 500]
];
let wsPnl = XLSX.utils.aoa_to_sheet(pnl);
XLSX.utils.book_append_sheet(wb, wsPnl, "P&L Mensual");

try {
    let result = financialEngine(wb);
    console.log(JSON.stringify(result.data.map(d => ({ date: d.date, d: d.deudaMetrics })), null, 2));
} catch (e) {
    console.error(e);
}
