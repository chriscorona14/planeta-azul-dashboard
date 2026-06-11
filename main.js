
async function getFinanceDB() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open('PlanetaAzulDB', 7);
        req.onupgradeneeded = (e) => {
            if (!e.target.result.objectStoreNames.contains('finance_cache')) {
                e.target.result.createObjectStore('finance_cache');
            }
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}
import * as XLSX from 'xlsx';
import { GoogleGenAI } from "@google/genai";
import * as d3 from 'd3';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { financialEngine, formatCurrency, formatRawCurrency, formatPercent, normalizeText, calculateYTD, formatSegmentName, formatDateKey } from "./financialEngine.js";
import { buildLLMInput } from "./buildLLMInput.js";
import { validateLLMInput } from "./validator.js";

// HELPER
function getSortYear(d) {
  if (!d || !d.sortDate) return 0;
  const dt = d.sortDate;
  if (typeof dt === 'string') {
    const parsed = new Date(dt);
    return isNaN(parsed.getTime()) ? 0 : parsed.getFullYear();
  }
  if (dt instanceof Date) return dt.getFullYear();
  return 0;
}

// --- PREVENCIÓN DE PANTALLA BLANCA (ERROR BOUNDARIES) ---
window.addEventListener('error', function(e) {
    console.error("Global error caught:", e);
    showGlobalError("Ocurrió un error inesperado en la aplicación. Por favor, recarga la página.");
});

window.addEventListener('unhandledrejection', function(e) {
    // Silencio en la UI para no interrumpir al usuario con errores de extensiones o red menores
});

function showGlobalError(msg) {
    if (!document.getElementById("global-error-banner")) {
        const banner = document.createElement('div');
        banner.id = "global-error-banner";
        banner.style = "position:fixed; bottom:20px; right:20px; max-width:400px; background:var(--danger, #e76f51); color:white; padding:16px; z-index:9999; font-weight:500; border-radius:8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); display:flex; flex-direction:column; gap:12px;";
        banner.innerHTML = `
            <div>
                <span style="vertical-align:middle;">${msg}</span>
            </div>
            <div style="display:flex; justify-content:flex-end; gap:8px;">
                <button onclick="this.parentElement.parentElement.remove()" style="padding:4px 12px; border:1px solid rgba(255,255,255,0.5); background:transparent; color:white; cursor:pointer; border-radius:4px; font-size:12px;">Ignorar</button>
                <button onclick="window.location.reload()" style="padding:4px 12px; border:none; background:white; color:var(--danger, #e76f51); cursor:pointer; border-radius:4px; font-size:12px; font-weight:bold;">Recargar</button>
            </div>
        `;
        document.body.appendChild(banner);
    }
}

// --- FUNCIONES DE APOYO PARA CACHÉ ---
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

function base64ToArrayBuffer(base64) {
    const binary_string = atob(base64);
    const bytes = new Uint8Array(binary_string.length);
    for (let i = 0; i < binary_string.length; i++) {
        bytes[i] = binary_string.charCodeAt(i);
    }
    return bytes.buffer;
}

let globalFinancialData = [];
Object.defineProperty(window, 'globalFinancialData', {
    get: function() { return globalFinancialData; },
    set: function(val) { globalFinancialData = val; },
    configurable: true,
    enumerable: true
});

function ensureMockFinancialData() {
    if (!globalFinancialData || globalFinancialData.length === 0) {
        if (window.hasVentasAccess || window.hasComercialAccess || (ceoData && ceoData.length > 0)) {
            console.log("⚡ [Fallback] Generando periodos base mockeados...");
            const fallback = [];
            const months = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
            for(let i=0; i<12; i++) {
                fallback.push({
                    date: `${months[i]} 2026`,
                    Periodo: `${String(i+1).padStart(2,'0')}-2026`,
                    sortDate: new Date(2026, i, 1),
                    kpis: { ingresos: 0, ebitda: 0, cashflow: 0, utilidadNeta: 0, margenBruto: 0 },
                    trend: { ingresos: 0, ebitda: 0, cashflow: 0, utilidadNeta: 0 },
                    balance: { deudaTotal: 0, ebitdaLTM: 0, efectivo: 0, roa: 0, roe: 0, ccc: 0 },
                    cashflowDetail: { ending: 0, ops: 0, inv: 0, fin: 0 },
                    pnl: {
                        categorias: { "Costo de Ventas": 0, "Ingresos": 0, "EBITDA": 0, "OPEX": 0 },
                        opexDetalle: { "Gastos Administrativos": 0, "Gastos de Mercadeo": 0 },
                        segments: {},
                        fullRows: []
                    },
                    ppto: { 
                        kpis: { ingresos: 0, ebitda: 0, cashflow: 0, utilidadNeta: 0, caja: 0 },
                        pnl: { categorias: {}, opexDetalle: {}, segments: {}, fullRows: [] }
                    },
                    ytd: { 
                        ingresos: 0, ebitda: 0, utilidadNeta: 0, 
                        ppto: { ingresos: 0, ebitda: 0, utilidadNeta: 0, caja: 0 }, 
                        trend: { ingresos: 0, ebitda: 0, utilidadNeta: 0 } 
                    },
                    estados: { fullRows: [] }
                });
            }
            fallback._isMock = true;
            globalFinancialData = fallback;
            if (typeof renderDashboard === 'function') {
                renderDashboard(globalFinancialData);
            }
        }
    }
}

let ceoData = null;
let ventasCeoData = null; // Will hold { columns: [], rows: [] } mapped
let ventasCeoCurrentMetric = 'Volumen';
let isYTDMode = false;
const loader = document.getElementById('loader');
const monthSelector = document.getElementById('monthSelector');

window.aiSummaryCache = {};
const AI_ADMIN_PASSWORD = import.meta.env.VITE_AI_ADMIN_PASSWORD || null;
window.aiEnabled = localStorage.getItem('aiEnabled') === 'true';

window.handleAiError = function(source, err) {
    const errorString = err && err.message ? err.message : String(err);
    if (errorString.includes("429") || errorString.includes("RESOURCE_EXHAUSTED") || errorString.includes("quota")) {
        if (window.aiEnabled) {
            console.warn(`[${source}] Cuota de API Gemini agotada. Cambiando a versión estática (Off).`);
            window.aiEnabled = false;
            localStorage.setItem('aiEnabled', 'false');
            applyAiUIState();
        }
    } else {
        console.warn(`[${source}] Detalle:`, errorString);
    }
};

function applyAiUIState() {
    const toggle = document.getElementById('toggleAiFeatures');
    if (toggle) toggle.checked = window.aiEnabled;

    const chatBtn = document.getElementById('openAiChatBtn');
    if (chatBtn) {
        chatBtn.style.display = window.aiEnabled ? 'flex' : 'none';
    }

    const summaryBox = document.getElementById('aiSummaryBox');
    if (summaryBox && !window.aiEnabled) {
        summaryBox.style.display = 'none';
    }
    
    const insightsSection = document.getElementById('ai-insights-section');
    if (insightsSection) {
        insightsSection.style.display = window.aiEnabled ? 'block' : 'none';
    }
    
    const btnRunSim = document.getElementById('btn-run-simulation');
    if (btnRunSim) {
        btnRunSim.style.display = window.aiEnabled ? 'flex' : 'none';
    }
    
    const simMenuItem = document.getElementById('sim-menu-item');
    if (simMenuItem) {
        simMenuItem.style.display = 'block';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    applyAiUIState();

    const toggle = document.getElementById('toggleAiFeatures');
    const modal = document.getElementById('aiPasswordModal');
    const pwInput = document.getElementById('aiPasswordInput');
    const cancelBtn = document.getElementById('aiPasswordCancel');
    const confirmBtn = document.getElementById('aiPasswordConfirm');

    if (toggle) {
        toggle.addEventListener('change', (e) => {
            if (e.target.checked) {
                // Trying to turn ON
                e.target.checked = false; // Prevent until authorized
                modal.style.display = 'flex';
                pwInput.value = '';
                pwInput.focus();
            } else {
                // Turning OFF is free
                window.aiEnabled = false;
                localStorage.setItem('aiEnabled', 'false');
                applyAiUIState();
            }
        });
    }

    const handleAuth = () => {
        if (pwInput.value === AI_ADMIN_PASSWORD) {
            window.aiEnabled = true;
            localStorage.setItem('aiEnabled', 'true');
            modal.style.display = 'none';
            applyAiUIState();
            // Try to generate summary for current view if present
            if (globalFinancialData && globalFinancialData.length > 0) {
                const idx = monthSelector ? parseInt(monthSelector.value, 10) : globalFinancialData.length - 1;
                generateExecutiveSummary(globalFinancialData, isNaN(idx) ? globalFinancialData.length - 1 : idx);
            }
        } else {
            alert('Acceso Denegado. Contraseña incorrecta.');
            pwInput.value = '';
        }
    };

    if (confirmBtn) confirmBtn.addEventListener('click', handleAuth);
    if (pwInput) pwInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') handleAuth();
    });
    if (cancelBtn) cancelBtn.addEventListener('click', () => {
        modal.style.display = 'none';
    });
});

async function generateExecutiveSummary(data, index) {
    if (!window.aiEnabled) return;
    
    const box = document.getElementById('aiSummaryBox');
    const curr = data[index];
    if (!box || !curr) return;
    
    const mesKey = curr.date || `m_${index}`;
    box.style.display = 'block';

    if (window.aiSummaryCache[mesKey]) {
        box.innerHTML = `<h3>Resumen Ejecutivo</h3>${window.aiSummaryCache[mesKey]}`;
        return;
    }
    
    box.innerHTML = '⏳ Analizando resultados financieros...';
    
    try {
        const contextData = {
            periodo: curr.date,
            kpis: curr.kpis,
            balance: curr.balance,
            cashflowDetail: curr.cashflowDetail,
            pnl_categorias: curr.pnl?.categorias
        };
        
        const promptInfo = JSON.stringify(contextData, null, 2);
        
        const promptText = `You are a senior financial analyst and strategic consultant specialized in the bottled water and beverage distribution industry.

Your task is to analyze the company’s live financial dashboard data and benchmark its performance against industry standards (both global and emerging markets).

IMPORTANT:
- If data is missing, estimate using reasonable financial assumptions based on industry behavior.
- Focus on operational reality, not accounting formality.

========================
CONTEXT: INDUSTRY BENCHMARKS
========================
Profitability:
- Gross Margin: 40% – 70%
- Operating Margin: 5% – 10%
- Net Margin: 2% – 15% (low performers) / up to 30% (optimized players)

Cost Structure:
- Distribution & logistics is typically the largest cost driver (can exceed 30% of revenue).
- Production cost is relatively low compared to logistics.

Financial Health:
- Debt Ratio: 0.4 – 0.5
- ROE: 15% – 25%
- EBITDA Multiple (valuation): 4x – 8x

Business Model Notes:
- This is a high-frequency, logistics-driven business.
- Profitability depends more on route efficiency and asset utilization than product cost.
- Customer density and delivery optimization are key performance drivers.

========================
LIVE DASHBOARD DATA
========================
${promptInfo}

========================
TASK
========================
Analyze the provided data and:
1. Identify key financial metrics (Revenue growth, margins, cost structure, cash flow).
2. Benchmark vs industry (Classify as Below/Within/Above industry and quantify deviation).
3. Diagnose the business (Identify if it's logistics efficient, margin constrained, etc.).
4. Identify root causes for any deviations.
5. Advanced insights (Detect structural risks like over-dependence on logistics or high working capital lock).
6. Competitive positioning.
7. Actionable recommendations (Provide 3–5 high-impact actions).

Additionally (Level God Insight):
- Detect if the company currently behaves more like a "distribution company" or a "manufacturing company" based on its cost structure.
- Estimate how much EBITDA improvement is possible from logistics optimization (in %).

========================
OUTPUT FORMAT
========================
Return structured output strictly in Markdown format:
1. Executive Summary (max 5 bullets)
2. Financial Benchmark Table (Metric | Company | Benchmark | Status | Variance)
3. Key Issues Identified
4. Root Cause Analysis
5. Strategic Positioning
6. Action Plan (prioritized by impact)

Be concise, analytical, and brutally honest. Focus on financial and operational reality.

========================
LANGUAGE
========================
IMPORTANT: Always return the full response in Spanish (Español).`;
        
        let timeoutId;
        const timeoutPromise = new Promise((_, reject) => {
            timeoutId = setTimeout(() => reject(new Error('AI Request Timeout (45s)')), 45000);
        });

        let apiCallPromise;
        try {
            apiCallPromise = getAI().models.generateContent({
                model: "gemini-2.5-flash",
                contents: promptText
            });
            apiCallPromise.catch(err => window.handleAiError("Summary", err));
        } catch (err) {
            apiCallPromise = Promise.reject(err);
            apiCallPromise.catch(()=> /* handled */ {});
        }

        let response;
        try {
            response = await Promise.race([
                apiCallPromise,
                timeoutPromise
            ]);
        } finally {
            clearTimeout(timeoutId);
        }
        
        const rawText = response.text || "No se pudo generar el resumen.";
        const textResponse = typeof marked !== 'undefined' ? marked.parse(rawText) : rawText.replace(/\n/g, '<br>');
        
        window.aiSummaryCache[mesKey] = textResponse;
        
        box.innerHTML = textResponse;
    } catch (err) {
        window.handleAiError("Summary", err);
        box.innerHTML = '⚠️ El análisis general de IA está temporalmente no disponible.';
    }
}

// Initialize Gemini (Lazy initialization)
let aiInstance = null;
function getAI() {
    if (!aiInstance) {
        if (!import.meta.env.VITE_GEMINI_API_KEY) {
            throw new Error("An API Key must be set (VITE_GEMINI_API_KEY is missing)");
        }
        aiInstance = new GoogleGenAI({ apiKey: import.meta.env.VITE_GEMINI_API_KEY });
    }
    return aiInstance;
}

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: import.meta.env.VITE_MSAL_CLIENT_ID || import.meta.env.VITE_MICROSOFT_CLIENT_ID,
        authority: import.meta.env.VITE_MSAL_TENANT_ID ? `https://login.microsoftonline.com/${import.meta.env.VITE_MSAL_TENANT_ID}` : "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin,
        navigateToLoginRequestUrl: false,
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
    }
};

let msalInstance;
if (window.msal) {
    msalInstance = new window.msal.PublicClientApplication(msalConfig);
}

let SHARPOINT_FILE_URL = localStorage.getItem('CUSTOM_ONEDRIVE_FILE_URL') || import.meta.env.VITE_ONEDRIVE_FILE_URL || import.meta.env.VITE_ONEDRIVE_ITEM_ID;
let SHARPOINT_VENTAS_FILE_URL = localStorage.getItem('CUSTOM_ONEDRIVE_VENTAS_URL') || import.meta.env.VITE_CEO_FILE_URL;
let RESUMEN_COMERCIAL_URL = localStorage.getItem('CUSTOM_RESUMEN_COMERCIAL_URL') || import.meta.env.VITE_RESUMEN_COMERCIAL_URL;
let PG_HORIZONTAL_URL = localStorage.getItem('CUSTOM_PG_HORIZONTAL_URL') || import.meta.env.VITE_PG_HORIZONTAL_URL;
let CXP_URL = localStorage.getItem('CUSTOM_CXP_URL') || import.meta.env.VITE_CXP_URL;
let COSTO_UNITARIO_URL = import.meta.env.VITE_COSTO_UNITARIO_URL;

const encodeUrlM365 = (url) => {
    if (!url || typeof url !== 'string' || url.trim().length === 0) return null;
    try {
        let cleanUrl = url.trim();
        // Remove known query params that break Graph API, but PRESERVE 'sourcedoc'
        cleanUrl = cleanUrl.replace(/[?&]download=1/gi, '');
        cleanUrl = cleanUrl.replace(/[?&]action=[^&]+/gi, '');
        cleanUrl = cleanUrl.replace(/[?&]e=[^&]+/gi, '');
        // Fix dangling ? if query string was emptied
        if (cleanUrl.endsWith('?')) {
            cleanUrl = cleanUrl.slice(0, -1);
        }
        return btoa(cleanUrl).replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
    } catch(e) {
        console.error("Error base64 encoding url:", url, e);
        return null;
    }
};

const resolveSharepointUrlClient = (inputUrl) => {
    if (!inputUrl) return inputUrl;
    
    let resolved = String(inputUrl).trim().replace(/&amp;/g, "&");
    
    // Clean braces of the guid if needed for check
    const cleanInput = resolved.replace(/^\{|\}$/g, "");
    if (/^[0-9a-fA-F\-]{36}$/.test(cleanInput)) {
        return `https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={${cleanInput}}&download=1`;
    }
    
    if (resolved.includes("sharepoint.com") || resolved.includes("onedrive.live.com")) {
        if (resolved.includes("action=embedview")) {
            resolved = resolved.replace(/action=embedview/g, "download=1");
        }
        
        resolved = resolved
            .replace(/&wdAllowInteractivity=[^&]*/g, "")
            .replace(/&wdHideGridlines=[^&]*/g, "")
            .replace(/&wdHideHeaders=[^&]*/g, "")
            .replace(/&wdDownloadButton=[^&]*/g, "")
            .replace(/&wdInConfigurator=[^&]*/g, "")
            .replace(/&edaebf=[^&]*/g, "");
            
        if (!resolved.includes("download=1")) {
            resolved += resolved.includes("?") ? "&download=1" : "?download=1";
        }
        
        return resolved;
    }
    
    return resolved;
};

window.updateM365UI = function(account) {
    const loginM365Btn = document.getElementById('loginM365Btn');
    const m365ActiveSession = document.getElementById('m365ActiveSession');
    const m365AccountEmail = document.getElementById('m365AccountEmail');
    
    const m365UrlMaster = document.getElementById('m365UrlMaster');
    const m365UrlVentas = document.getElementById('m365UrlVentas');
    const m365UrlComercial = document.getElementById('m365UrlComercial');
    const m365UrlPgHorizontal = document.getElementById('m365UrlPgHorizontal');
    const m365UrlCxp = document.getElementById('m365UrlCxp');
    const m365UrlCostoUnitario = document.getElementById('m365UrlCostoUnitario');

    if (account) {
        if (loginM365Btn) loginM365Btn.style.display = 'none';
        if (m365ActiveSession) m365ActiveSession.style.display = 'block';
        if (m365AccountEmail) m365AccountEmail.textContent = `Usuario: ${account.username || account.name || 'Conectado'}`;
        
        if (m365UrlMaster) m365UrlMaster.value = localStorage.getItem('CUSTOM_ONEDRIVE_FILE_URL') || '';
        if (m365UrlVentas) m365UrlVentas.value = localStorage.getItem('CUSTOM_ONEDRIVE_VENTAS_URL') || '';
        if (m365UrlComercial) m365UrlComercial.value = localStorage.getItem('CUSTOM_RESUMEN_COMERCIAL_URL') || '';
        if (m365UrlPgHorizontal) m365UrlPgHorizontal.value = localStorage.getItem('CUSTOM_PG_HORIZONTAL_URL') || '';
        if (m365UrlCxp) m365UrlCxp.value = localStorage.getItem('CUSTOM_CXP_URL') || '';
        if (m365UrlCostoUnitario) m365UrlCostoUnitario.value = localStorage.getItem('CUSTOM_COSTO_UNITARIO_URL') || '';
    } else {
        if (loginM365Btn) loginM365Btn.style.display = 'flex';
        if (m365ActiveSession) m365ActiveSession.style.display = 'none';
        
        if (m365UrlMaster) m365UrlMaster.value = '';
        if (m365UrlVentas) m365UrlVentas.value = '';
        if (m365UrlComercial) m365UrlComercial.value = '';
        if (m365UrlPgHorizontal) m365UrlPgHorizontal.value = '';
        if (m365UrlCxp) m365UrlCxp.value = '';
        if (m365UrlCostoUnitario) m365UrlCostoUnitario.value = '';
    }
};

async function connectM365() {
    if (!msalInstance) {
        alert("MSAL no inicializado.");
        return;
    }

    try {
        await msalInstance.initialize?.(); 
        await msalInstance.handleRedirectPromise?.();

        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.Read", "Files.Read", "Files.Read.All"],
            prompt: "select_account"
        });
        
        const token = loginResponse.accessToken;
        msalInstance.setActiveAccount(loginResponse.account);
        window.m365LoggedIn = true;
        window.updateM365UI(loginResponse.account);
        
        await fetchMasterData(token);
    } catch (error) {
        if (error.errorCode === "user_cancelled" || (error.message && error.message.includes("user_cancelled"))) {
            console.log("El usuario canceló el inicio de sesión.");
            return;
        }
        if (error.errorCode === "interaction_in_progress" || (error.message && error.message.includes("popup_window_error"))) {
            console.warn("Popup bloqueado o interacción en progreso.");
            alert("Bloqueador de ventanas emergentes detectado. Por favor, habilite las ventanas emergentes para Microsoft 365 o abra la aplicación en una nueva pestaña.");
            return;
        }
        console.error(error);
        alert("Error autenticando con Office 365: " + error.message);
    }
}

function showErrorUI(mensaje) {
    // 1. Detener procesos en segundo plano
    if (window.stop) window.stop();

    // 2. Limpiar la pantalla y aplicar el nuevo estilo corporativo
    document.body.innerHTML = `
        <div style="position:fixed; top:0; left:0; width:100%; height:100%; background:#f3f4f6; display:flex; align-items:center; justify-content:center; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; z-index:999999;">
            <div style="background:white; padding:40px; border-radius:12px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); text-align:center; max-width:450px; width:90%;">
                <div style="color:#004a99; font-size:50px; margin-bottom:20px;">🔒</div>
                <h2 style="color:#111827; margin:0 0 10px 0; font-size:1.5rem;">Acceso Restringido</h2>
                <p style="color:#6b7280; line-height:1.5; margin-bottom:30px;">${mensaje}<br><br>Por favor, contacta al administrador del sistema.</p>
                
                <button onclick="sessionStorage.clear(); localStorage.clear(); location.href='/';" 
                        style="background:#004a99; color:white; border:none; padding:12px 24px; border-radius:6px; font-weight:600; cursor:pointer; width:100%; transition: background 0.2s;">
                    Cerrar Sesión / Cambiar Cuenta
                </button>
                
                <p style="margin-top:20px; font-size:0.85rem; color:#9ca3af;">Finance Dashboard Pro | Planeta Azul</p>
            </div>
        </div>
    `;
}

window.hasMasterAccess = false;
window.hasVentasAccess = false;
window.hasComercialAccess = false;

window.applyRoleBasedUI = function(hasMaster, hasVentas, hasComercial = false) {
    const hasVentasFeature = 
        !!import.meta.env.VITE_CEO_FILE_URL || 
        (typeof window.runtimeConfig !== 'undefined' && !!window.runtimeConfig?.VITE_CEO_FILE_URL) || 
        !!document.getElementById("view-ventas-ceo");
    
    if (hasVentasFeature) {
        hasVentas = true;
        window.hasVentasAccess = true;
    }

    const mainContainer = document.querySelector('.main-container');
    const sidebarItems = document.querySelectorAll('.sidebar .menu-item');
    const monthSelector = document.getElementById('monthSelector');
    const viewModeToggle = document.querySelector('.view-mode-toggle');
    const dropZone = document.getElementById('dropZone');
    const sidebar = document.querySelector('.sidebar');
    const contentHeader = document.querySelector('.content-header');
    const headerActions = document.querySelector('.header-actions');
    const headerInfo = document.querySelector('.header-info');
    const viewContainers = document.querySelectorAll('.view-container');

    const groupSeguimiento = document.getElementById('grupo-seguimiento');
    const headerSeguimiento = groupSeguimiento ? groupSeguimiento.previousElementSibling : null;
    const groupVentas = document.getElementById('grupo-ventas');
    const headerVentas = groupVentas ? groupVentas.previousElementSibling : null;
    const groupCostos = document.getElementById('grupo-costos');
    const headerCostos = groupCostos ? groupCostos.previousElementSibling : null;

    const menuVentasCeo = document.getElementById('menu-ventas-ceo');
    const menuResumenComercial = document.getElementById('menu-resumen-comercial');
    const menuPgHorizontal = document.getElementById('menu-pg-horizontal');
    const menuCostoUnitario = document.getElementById('menu-costo-unitario');

    // Limpiar banner previo si existe
    let deniedBanner = document.getElementById('access-denied-banner');
    if (deniedBanner) deniedBanner.remove();

    const anyAccess = hasMaster || hasVentas || hasComercial;

    if (anyAccess) {
        if (mainContainer) mainContainer.style.display = '';
        if (sidebar) sidebar.style.display = '';
        if (contentHeader) contentHeader.style.display = '';
        if (headerActions) headerActions.style.display = 'flex';
        if (headerInfo) headerInfo.style.display = '';
        if (monthSelector) {
            monthSelector.style.display = (globalFinancialData && globalFinancialData.length > 0) ? 'block' : 'none';
        }
        if (viewModeToggle) {
            viewModeToggle.style.display = (globalFinancialData && globalFinancialData.length > 0) ? 'flex' : 'none';
        }
        if (dropZone) {
            const isConfigActive = document.getElementById('view-config')?.classList.contains('active');
            dropZone.style.display = isConfigActive ? 'block' : 'none';
        }

        // Grupo Seguimiento
        if (groupSeguimiento) groupSeguimiento.style.display = hasMaster ? '' : 'none';
        if (headerSeguimiento) headerSeguimiento.style.display = hasMaster ? 'flex' : 'none';

        // Grupo Ventas
        const showVentasGroup = hasVentas || hasComercial;
        if (groupVentas) groupVentas.style.display = showVentasGroup ? '' : 'none';
        if (headerVentas) headerVentas.style.display = showVentasGroup ? 'flex' : 'none';

        // Grupo Costos
        const showCostosGroup = hasComercial; // assuming 'hasComercial' allows viewing Costos for now
        if (groupCostos) groupCostos.style.display = showCostosGroup ? '' : 'none';
        if (headerCostos) headerCostos.style.display = showCostosGroup ? 'flex' : 'none';

        // Items individuales dentro de Ventas y Costos
        if (menuVentasCeo) menuVentasCeo.parentElement.style.display = hasVentas ? '' : 'none';
        if (menuResumenComercial) menuResumenComercial.parentElement.style.display = hasComercial ? '' : 'none';
        if (menuPgHorizontal) menuPgHorizontal.parentElement.style.display = hasComercial ? '' : 'none';
        if (menuCostoUnitario) menuCostoUnitario.parentElement.style.display = hasComercial ? '' : 'none';

        // Redirección si la vista activa ya no es accesible
        const activeView = Array.from(viewContainers).find(v => v.classList.contains('active'));
        if (activeView) {
            const isSeguimientoView = activeView.id.startsWith('view-') && activeView.id !== 'view-ventas-ceo' && activeView.id !== 'view-resumen-comercial';
            if (isSeguimientoView && !hasMaster) {
                 if (hasVentas) document.getElementById('menu-ventas-ceo')?.click();
                 else if (hasComercial) document.getElementById('menu-resumen-comercial')?.click();
            } else if (activeView.id === 'view-ventas-ceo' && !hasVentas) {
                 if (hasMaster) document.getElementById('menu-kpi')?.click();
                 else if (hasComercial) document.getElementById('menu-resumen-comercial')?.click();
            } else if (activeView.id === 'view-resumen-comercial' && !hasComercial) {
                 if (hasMaster) document.getElementById('menu-kpi')?.click();
                 else if (hasVentas) document.getElementById('menu-ventas-ceo')?.click();
            }
        }

    } else {
        // Escenario: Sin acceso - de todos modos dejamos al usuario operar e intentar cargar o configurar
        if (mainContainer) mainContainer.style.display = '';
        if (sidebar) sidebar.style.display = '';
        if (contentHeader) contentHeader.style.display = '';
        
        if (typeof window.handleZeroState === 'function') {
            window.handleZeroState();
        }
    }
};

async function fetchMasterData(token = null) {
    const statusEl = document.getElementById('engineStatus');
    const sidebarSyncDot = document.getElementById('sidebarSyncDot');
    const sidebarSyncText = document.getElementById('sidebarSyncText');

    // ==========================================
    // AUTO-RECUPERACIÓN SILENCIOSA DE SESIÓN MSAL
    // ==========================================
    if (!token && typeof msalInstance !== 'undefined' && msalInstance) {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            try {
                msalInstance.setActiveAccount(accounts[0]);
                const silentResponse = await msalInstance.acquireTokenSilent({
                    scopes: ["User.Read", "Files.Read", "Files.Read.All"],
                    account: accounts[0]
                });
                token = silentResponse.accessToken;
                window.m365LoggedIn = true;
                console.log("⚡ [M365] Token refrescado y re-vinculado de manera silenciosa.");
            } catch (err) {
                console.warn("⚠️ No se pudo reconectar de forma silenciosa al iniciar sync:", err);
            }
        }
    }

    if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--warning)';
    if (sidebarSyncText) {
        sidebarSyncText.innerText = 'Sincronizando...';
        sidebarSyncText.style.color = 'var(--warning)';
    }

    const viewContainers = document.querySelectorAll('.view-container');
    const dropZone = document.getElementById('dropZone');
    const loader = document.getElementById('loader');
    const loginBtn = document.getElementById('loginM365Btn');
    if (loginBtn) loginBtn.style.display = 'none';

    // ==========================================
    // 1. LA BARRERA SILENCIOSA (Stale-While-Revalidate)
    // ==========================================
    if (window.isMagicLoaded) {
        console.log("⚡ Modo Silencioso: Caché activa. Buscando actualizaciones en O365 sin bloquear la UI...");
        if (statusEl) {
            statusEl.style.background = '#e0f2fe';
            statusEl.style.color = '#0369a1';
            statusEl.style.borderColor = '#bae6fd';
            statusEl.innerHTML = "🔄 Buscando actualizaciones en Microsoft 365 en segundo plano...";
        }
    } else {
        console.log("No hay caché. Iniciando carga completa bloqueante...");
        viewContainers.forEach(v => v.style.display = 'none');
        if (dropZone) dropZone.style.display = 'none';
        
        if (loader) {
            loader.innerHTML = `
                <div style="background: white; padding: 40px; border-radius: 16px; box-shadow: var(--shadow-lg); width: 340px; text-align: center; border: 1px solid var(--border);">
                    <i data-lucide="loader" class="spin-icon" style="width: 48px; height: 48px; color: var(--primary); margin: 0 auto; margin-bottom: 20px; display: block;"></i>
                    <h4 style="font-size: 1.1rem; color: var(--text-primary); margin-bottom: 12px; font-weight: 600;">Sincronizando con M365...</h4>
                    <div style="width: 100%; height: 8px; background: #eef2f5; border-radius: 4px; overflow: hidden; margin-bottom: 12px;">
                        <div id="progressBar" style="width: 5%; height: 100%; background: var(--primary); transition: width 0.3s ease;"></div>
                    </div>
                    <p id="loadingText" style="font-size: 0.85rem; color: var(--text-secondary); margin: 0;">Conectando con Microsoft...</p>
                </div>
            `;
            loader.style.display = 'flex';
            if (window.lucide) window.lucide.createIcons();
        }
        if (statusEl) statusEl.innerHTML = "⏳ Estableciendo conexión con Microsoft Graph...";
    }

    // ==========================================
    // 2. DESCARGA DEL ARCHIVO (O365 Directo o Intercepción)
    // ==========================================
    try {
        const SYNC_TIMEOUT = 300000; // 5 minutos para tolerar descargas masivas (110k filas)
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), SYNC_TIMEOUT);
        let arrayBuffer = null;
        let arrayBufferCeo = null;
        
        const prevMaster = window.hasMasterAccess;
        const prevVentas = window.hasVentasAccess;
        const prevComercial = window.hasComercialAccess;

        window.hasMasterAccess = false;
        window.hasVentasAccess = false;
        window.hasComercialAccess = false;

        try {
            const pb = document.getElementById('progressBar');
            const lt = document.getElementById('loadingText');

            if (token) {
                // Get runtime config from the server to bypass build-time env var freezing
                let runtimeConfig = {};
                try {
                    const configRes = await fetch("/api/config");
                    const contentType = configRes.headers.get("content-type");
                    if (configRes.ok && contentType && contentType.includes("application/json")) {
                        runtimeConfig = await configRes.json();
                        window.runtimeConfig = runtimeConfig;
                    }
                } catch (e) {
                    console.warn("Could not fetch /api/config", e);
                }

                // Sincronizar variables activas en memoria con localStorage y config
                SHARPOINT_FILE_URL = localStorage.getItem('CUSTOM_ONEDRIVE_FILE_URL') || import.meta.env.VITE_ONEDRIVE_FILE_URL || import.meta.env.VITE_ONEDRIVE_ITEM_ID || runtimeConfig.VITE_ONEDRIVE_FILE_URL;
                SHARPOINT_VENTAS_FILE_URL = localStorage.getItem('CUSTOM_ONEDRIVE_VENTAS_URL') || import.meta.env.VITE_CEO_FILE_URL || runtimeConfig.VITE_CEO_FILE_URL || "";

                // Descarga Master Financiero (Directo)
                if (lt) lt.innerText = "Descargando Finanzas Master (5.0 MB)...";
                if (pb) pb.style.width = "20%";
                const resolvedUrl = resolveSharepointUrlClient(SHARPOINT_FILE_URL);
                const encodedUrl = encodeUrlM365(resolvedUrl);
                if (encodedUrl) {
                    const graphUrl = `https://graph.microsoft.com/v1.0/shares/u!${encodedUrl}/driveItem/content`;
                    const req = await fetch(graphUrl, { headers: { "Authorization": `Bearer ${token}` }, signal: controller.signal });
                    if (req.ok) {
                        arrayBuffer = await req.arrayBuffer();
                        window.hasMasterAccess = true;
                    } else {
                        console.warn("Graph API rejected Master sync. Status:", req.status, req.statusText, "- Intentando proxy con token...");
                        let paramsMaster = SHARPOINT_FILE_URL ? `?url=${encodeURIComponent(SHARPOINT_FILE_URL)}` : '';
                        const proxyReq = await fetch(`/api/downloadSync${paramsMaster}`, { 
                            headers: { "Authorization": `Bearer ${token}` },
                            signal: controller.signal 
                        });
                        if (proxyReq.ok) {
                            arrayBuffer = await proxyReq.arrayBuffer();
                            window.hasMasterAccess = true;
                        } else {
                            window.hasMasterAccess = false;
                            if (!window.isMagicLoaded && !(globalFinancialData && globalFinancialData.length > 0)) {
                                globalFinancialData = null;
                                try {
                                    const db = await getFinanceDB();
                                    const tx = db.transaction('finance_cache', 'readwrite');
                                    tx.objectStore('finance_cache').delete('MASTER_FINANCE_KEY');
                                } catch(e) {}
                            }
                        }
                    }
                }

                // Descarga Ventas CEO inmediata (Directo)
                if (lt) lt.innerText = "Descargando Ventas CEO (133 kB)...";
                if (pb) pb.style.width = "40%";
                const CEO_FILE_URL = SHARPOINT_VENTAS_FILE_URL || import.meta.env.VITE_CEO_FILE_URL || runtimeConfig.VITE_CEO_FILE_URL;
                const resolvedCeoUrl = resolveSharepointUrlClient(CEO_FILE_URL);
                const encodedCeoUrl = encodeUrlM365(resolvedCeoUrl);
                if (encodedCeoUrl) {
                    const graphUrlCeo = `https://graph.microsoft.com/v1.0/shares/u!${encodedCeoUrl}/driveItem/content`;
                    const reqCeo = await fetch(graphUrlCeo, { headers: { "Authorization": `Bearer ${token}` }, signal: controller.signal });
                    if (reqCeo.ok) {
                        arrayBufferCeo = await reqCeo.arrayBuffer();
                        window.hasVentasAccess = true;
                    } else {
                        console.warn("Graph API rejected Ventas sync. Status:", reqCeo.status, reqCeo.statusText, "- Intentando proxy...");
                        let paramsVentas = CEO_FILE_URL ? `?url=${encodeURIComponent(CEO_FILE_URL)}` : '';
                        const proxyReqVentas = await fetch(`/api/downloadSyncVentas${paramsVentas}`, { 
                            headers: { "Authorization": `Bearer ${token}` },
                            signal: controller.signal 
                        });
                        if (proxyReqVentas.ok) {
                            arrayBufferCeo = await proxyReqVentas.arrayBuffer();
                            window.hasVentasAccess = true;
                        } else {
                            window.hasVentasAccess = false;
                            ceoData = null;
                            try {
                                const db = await getFinanceDB();
                                const tx = db.transaction('finance_cache', 'readwrite');
                                tx.objectStore('finance_cache').delete('CEO_VENTAS_KEY_V3');
                            } catch(e) {}
                        }
                    }
                }

                // Descarga Resumen Comercial (Directo)
                if (lt) lt.innerText = "Descargando Resumen Comercial (5.0 MB)...";
                if (pb) pb.style.width = "60%";
                RESUMEN_COMERCIAL_URL = localStorage.getItem('CUSTOM_RESUMEN_COMERCIAL_URL') || import.meta.env.VITE_RESUMEN_COMERCIAL_URL || runtimeConfig.VITE_RESUMEN_COMERCIAL_URL;
                const resolvedComercialUrl = resolveSharepointUrlClient(RESUMEN_COMERCIAL_URL);
                const encodedComercialUrl = encodeUrlM365(resolvedComercialUrl);
                if (encodedComercialUrl) {
                    const graphUrlComercial = `https://graph.microsoft.com/v1.0/shares/u!${encodedComercialUrl}/driveItem/content`;
                    const reqComercial = await fetch(graphUrlComercial, { headers: { "Authorization": `Bearer ${token}` }, signal: controller.signal });
                    if (reqComercial.ok) {
                        const arrayBufferComercial = await reqComercial.arrayBuffer();
                        if (!window.resumenComercialEngine) {
                            try {
                                const engine = await import('./resumenComercialEngine.js');
                                window.resumenComercialEngine = engine;
                            } catch (e) {
                                console.error("Error importing resumenComercialEngine on demand:", e);
                            }
                        }
                        if (window.resumenComercialEngine) {
                            try {
                                const result = await window.resumenComercialEngine.processManualFile(arrayBufferComercial);
                                if (result) {
                                    window.hasComercialAccess = true;
                                }
                            } catch (e) {
                                console.error("Error processing comercial sync:", e);
                            }
                        }
                    } else {
                        console.warn("Graph API rejected Resumen Comercial sync. Status:", reqComercial.status, reqComercial.statusText, "- Intentando proxy...");
                        let paramsComercial = RESUMEN_COMERCIAL_URL ? `?url=${encodeURIComponent(RESUMEN_COMERCIAL_URL)}` : '';
                        const proxyReqComercial = await fetch(`/api/downloadSyncComercial${paramsComercial}`, {
                            headers: { "Authorization": `Bearer ${token}` },
                            signal: controller.signal
                        });
                        if (proxyReqComercial.ok) {
                            const arrayBufferComercial = await proxyReqComercial.arrayBuffer();
                            if (!window.resumenComercialEngine) {
                                try {
                                    const engine = await import('./resumenComercialEngine.js');
                                    window.resumenComercialEngine = engine;
                                } catch (e) {
                                    console.error("Error importing resumenComercialEngine on demand:", e);
                                }
                            }
                            if (window.resumenComercialEngine) {
                                try {
                                    const result = await window.resumenComercialEngine.processManualFile(arrayBufferComercial);
                                    if (result) {
                                        window.hasComercialAccess = true;
                                    }
                                } catch (e) {
                                    console.error("Error processing comercial sync:", e);
                                }
                            }
                        } else {
                            window.hasComercialAccess = false;
                            try {
                                const db = await getFinanceDB();
                                const tx = db.transaction('finance_cache', 'readwrite');
                                tx.objectStore('finance_cache').delete('COMERCIAL_KEY');
                            } catch(e) {}
                        }
                    }
                }

                // Descarga P&G Horizontal (Directo)
                if (lt) lt.innerText = "Descargando P&G Horizontal por Producto (29.8 MB)...";
                if (pb) pb.style.width = "83%";
                PG_HORIZONTAL_URL = localStorage.getItem('CUSTOM_PG_HORIZONTAL_URL') || import.meta.env.VITE_PG_HORIZONTAL_URL || runtimeConfig.VITE_PG_HORIZONTAL_URL;
                const resolvedPgUrl = resolveSharepointUrlClient(PG_HORIZONTAL_URL);
                const encodedPgUrl = encodeUrlM365(resolvedPgUrl);
                if (encodedPgUrl) {
                    const graphUrlPg = `https://graph.microsoft.com/v1.0/shares/u!${encodedPgUrl}/driveItem/content`;
                    const reqPg = await fetch(graphUrlPg, { headers: { "Authorization": `Bearer ${token}` }, signal: controller.signal });
                    if (reqPg.ok) {
                        const arrayBufferPg = await reqPg.arrayBuffer();
                        if (!window.resumenComercialEngine) {
                            try {
                                const engine = await import('./resumenComercialEngine.js');
                                window.resumenComercialEngine = engine;
                            } catch (e) {
                                console.error("Error importing resumenComercialEngine on demand:", e);
                            }
                        }
                        if (window.resumenComercialEngine) {
                            try {
                                const dataPg = new Uint8Array(arrayBufferPg);
                                const workbookPg = XLSX.read(dataPg, { type: 'array' });
                                await window.resumenComercialEngine.processPgHorizontalWorkbook(workbookPg);
                                window.hasComercialAccess = true;
                            } catch (e) {
                                console.error("Error processing pg horizontal sync:", e);
                            }
                        }
                    } else {
                        console.warn("Graph API rejected PG Horizontal sync. Status:", reqPg.status, reqPg.statusText, "- Intentando proxy...");
                        let paramsPg = PG_HORIZONTAL_URL ? `?url=${encodeURIComponent(PG_HORIZONTAL_URL)}` : '';
                        const proxyReqPg = await fetch(`/api/downloadSyncPgHorizontal${paramsPg}`, {
                            headers: { "Authorization": `Bearer ${token}` },
                            signal: controller.signal
                        });
                        if (proxyReqPg.ok) {
                            const arrayBufferPg = await proxyReqPg.arrayBuffer();
                            if (!window.resumenComercialEngine) {
                                try {
                                    const engine = await import('./resumenComercialEngine.js');
                                    window.resumenComercialEngine = engine;
                                } catch (e) {
                                    console.error("Error importing resumenComercialEngine on demand:", e);
                                }
                            }
                            if (window.resumenComercialEngine) {
                                try {
                                    const dataPg = new Uint8Array(arrayBufferPg);
                                    const workbookPg = XLSX.read(dataPg, { type: 'array' });
                                    await window.resumenComercialEngine.processPgHorizontalWorkbook(workbookPg);
                                    window.hasComercialAccess = true;
                                } catch (e) {
                                    console.error("Error processing pg horizontal proxy sync:", e);
                                }
                            }
                        } else {
                            console.warn("Proxy also rejected PG Horizontal sync. Status:", proxyReqPg.status);
                        }
                    }
                }

                // Descarga Detalle CxP (Directo)
                if (lt) lt.innerText = "Descargando Detalle CxP (1.2 MB)...";
                if (pb) pb.style.width = "90%";
                CXP_URL = localStorage.getItem('CUSTOM_CXP_URL') || import.meta.env.VITE_CXP_URL || runtimeConfig.VITE_CXP_URL;
                const resolvedCxpUrl = resolveSharepointUrlClient(CXP_URL);
                const encodedCxpUrl = encodeUrlM365(resolvedCxpUrl);
                if (encodedCxpUrl) {
                    const graphUrlCxp = `https://graph.microsoft.com/v1.0/shares/u!${encodedCxpUrl}/driveItem/content`;
                    const reqCxp = await fetch(graphUrlCxp, { headers: { "Authorization": `Bearer ${token}` }, signal: controller.signal });
                    if (reqCxp.ok) {
                        const arrayBufferCxp = await reqCxp.arrayBuffer();
                        try {
                            const workbookCxp = XLSX.read(new Uint8Array(arrayBufferCxp), { type: 'array' });
                            await window.processCxpWorkbook(workbookCxp);
                            window.hasCxpAccess = true;
                        } catch (e) {
                            console.error("Error processing m365 cxp sync:", e);
                        }
                    } else {
                        console.warn("Graph API rejected CxP sync. Status:", reqCxp.status, reqCxp.statusText, "- Intentando proxy...");
                        let paramsCxp = CXP_URL ? `?url=${encodeURIComponent(CXP_URL)}` : '';
                        const proxyReqCxp = await fetch(`/api/downloadSyncCxp${paramsCxp}`, {
                            headers: { "Authorization": `Bearer ${token}` },
                            signal: controller.signal
                        });
                        if (proxyReqCxp.ok) {
                            const arrayBufferCxp = await proxyReqCxp.arrayBuffer();
                            try {
                                const workbookCxp = XLSX.read(new Uint8Array(arrayBufferCxp), { type: 'array' });
                                await window.processCxpWorkbook(workbookCxp);
                                window.hasCxpAccess = true;
                            } catch (e) {
                                console.error("Error processing m365 cxp proxy sync:", e);
                            }
                        } else {
                            console.warn("Proxy also rejected CxP sync. Status:", proxyReqCxp.status);
                        }
                    }
                }

                // Descarga Costo Unitario (Directo)
                COSTO_UNITARIO_URL = import.meta.env.VITE_COSTO_UNITARIO_URL || runtimeConfig.VITE_COSTO_UNITARIO_URL;
                const resolvedCostoUrl = resolveSharepointUrlClient(COSTO_UNITARIO_URL);
                const encodedCostoUrl = encodeUrlM365(resolvedCostoUrl);
                if (encodedCostoUrl) {
                    const graphUrlCosto = `https://graph.microsoft.com/v1.0/shares/u!${encodedCostoUrl}/driveItem/content`;
                    const reqCosto = await fetch(graphUrlCosto, { headers: { "Authorization": `Bearer ${token}` }, signal: controller.signal });
                    if (reqCosto.ok) {
                        const arrayBufferCosto = await reqCosto.arrayBuffer();
                        try {
                            const engine = await import('./costoUnitarioEngine.js');
                            window.costoUnitarioEngine = engine;
                            const workbook = XLSX.read(new Uint8Array(arrayBufferCosto), { type: 'array' });
                            await window.costoUnitarioEngine.processCostoUnitarioWorkbook(workbook);
                            if (window.updateCostoUnitario) window.updateCostoUnitario();
                        } catch (e) { console.error("Error processing m365 costo sync:", e); }
                    } else {
                        let paramsCosto = COSTO_UNITARIO_URL ? `?url=${encodeURIComponent(COSTO_UNITARIO_URL)}` : '';
                        const proxyReqCosto = await fetch(`/api/downloadSyncCostoUnitario${paramsCosto}`, { headers: { "Authorization": `Bearer ${token}` }, signal: controller.signal });
                        if (proxyReqCosto.ok) {
                            const arrayBufferCosto = await proxyReqCosto.arrayBuffer();
                            try {
                                const engine = await import('./costoUnitarioEngine.js');
                                window.costoUnitarioEngine = engine;
                                const workbook = XLSX.read(new Uint8Array(arrayBufferCosto), { type: 'array' });
                                await window.costoUnitarioEngine.processCostoUnitarioWorkbook(workbook);
                                if (window.updateCostoUnitario) window.updateCostoUnitario();
                            } catch (e) { console.error("Error processing m365 costo proxy sync:", e); }
                        }
                    }
                }
            } else {
                // ==========================================
                // INTERCEPCIÓN EN VERCEL (Caché vs Bloqueo)
                // ==========================================
                const isVercel = window.location.hostname.includes("vercel.app");
                if (isVercel) {
                    console.warn("⚠️ Sincronización solicitada sin autenticación en Vercel. Interceptando para prevenir 504 Timeout.");

                    if (window.isMagicLoaded) {
                        // Si ya tenemos datos previos en caché, el usuario trabaja de forma segura sin caídas
                        console.log("Caché activa detectada. Ignorando peticiones lentas de backend.");
                        if (statusEl) {
                            statusEl.innerHTML = `ℹ️ <span style="font-weight:600; cursor:pointer;" onclick="connectM365()">Conectar cuenta de Microsoft</span> para actualizar datos en tiempo real. Trabajando con caché offline segura.`;
                            statusEl.style.background = '#eff6ff';
                            statusEl.style.color = '#1e40af';
                            statusEl.style.borderColor = '#bfdbfe';
                        }
                        if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = '#3b82f6'; // Azul Info
                        if (sidebarSyncText) {
                            sidebarSyncText.innerText = 'Caché Offline';
                            sidebarSyncText.style.color = '#3b82f6';
                            sidebarSyncText.style.cursor = 'pointer';
                            sidebarSyncText.title = "Haz clic para conectar con Microsoft 365 y actualizar en tiempo real";
                        }
                        
                        // Ocultamos loaders y restauramos vistas
                        if (loader) loader.style.display = 'none';
                        viewContainers.forEach(v => {
                            if (v.id === 'view-kpi-dashboard' || v.classList.contains('active')) {
                                v.style.display = 'block';
                            }
                        });
                        return; // Fin anticipado y seguro
                    } else {
                        // Si es el primer ingreso del usuario y no hay caché, no podemos seguir por backend por los 504 timeouts.
                        // Forzamos un lindo Overlay indicando la necesidad exclusiva de MSAL.
                        if (statusEl) {
                            statusEl.innerHTML = `⚠️ <span style="font-weight:600; cursor:pointer;" onclick="connectM365()">Inicio de sesión Microsoft Requerido</span> para descargar archivos pesados.`;
                            statusEl.style.background = '#fef2f2';
                            statusEl.style.color = '#991b1b';
                            statusEl.style.borderColor = '#fecaca';
                        }
                        
                        if (loader) {
                            loader.innerHTML = `
                                <div style="background: white; padding: 40px; border-radius: 16px; box-shadow: var(--shadow-lg); width: 440px; text-align: center; border: 1px solid var(--border); font-family: system-ui, sans-serif;">
                                    <div style="font-size: 48px; margin-bottom: 20px;">🔒</div>
                                    <h4 style="font-size: 1.25rem; color: #111827; margin-bottom: 12px; font-weight: 700; letter-spacing: -0.025em;">Autenticación OneDrive Requerida</h4>
                                    <p style="font-size: 0.9rem; color: #4b5563; margin-bottom: 28px; line-height: 1.6;">
                                        Los informes financieros para <b>Resumen Comercial (5.0 MB)</b> y <b>P&G Horizontal (29.8 MB)</b> exceden los límites máximos permitidos por el proxy interactivo del backend en Vercel.<br><br>
                                        Inicia sesión con tu cuenta organizativa para descargarlos directamente desde los servidores CDN de Microsoft a máxima velocidad de forma 100% segura.
                                    </p>
                                    <button id="loginM365BtnLoader" style="background: #004a99; color: white; border: none; padding: 12px 28px; border-radius: 8px; font-weight: 600; cursor: pointer; display: inline-flex; align-items: center; justify-content: center; gap: 8px; width: 100%; font-size: 0.95rem; box-shadow: 0 4px 6px -1px rgba(0, 74, 153, 0.2);">
                                        <svg style="width: 16px; height: 16px; fill: white;" viewBox="0 0 23 23"><path d="M0 0h11v11H0zM12 0h11v11H12zM0 12h11v11H0zM12 12h11v11H12z"/></svg>
                                        Conectar Microsoft 365
                                    </button>
                                </div>
                            `;
                            loader.style.display = 'flex';
                            const btnLoader = document.getElementById('loginM365BtnLoader');
                            if (btnLoader) {
                                btnLoader.addEventListener('click', connectM365);
                            }
                        }
                        
                        if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--danger)';
                        if (sidebarSyncText) {
                            sidebarSyncText.innerText = 'Iniciar Sesión M365';
                            sidebarSyncText.style.color = 'var(--danger)';
                            sidebarSyncText.style.cursor = 'pointer';
                        }
                        return; // Bloqueado tempranamente
                    }
                }

                // Fallback clásico local (solo si no es Vercel)
                let paramsMaster = SHARPOINT_FILE_URL ? `?url=${encodeURIComponent(SHARPOINT_FILE_URL)}` : '';
                const response = await fetch(`/api/downloadSync${paramsMaster}`, { signal: controller.signal });
                if (response.ok) {
                    arrayBuffer = await response.arrayBuffer();
                    window.hasMasterAccess = true;
                }

                let paramsVentas = SHARPOINT_VENTAS_FILE_URL ? `?url=${encodeURIComponent(SHARPOINT_VENTAS_FILE_URL)}` : '';
                const responseVentas = await fetch(`/api/downloadSyncVentas${paramsVentas}`, { signal: controller.signal });
                if (responseVentas.ok) {
                    arrayBufferCeo = await responseVentas.arrayBuffer();
                    window.hasVentasAccess = true;
                }

                let paramsComercial = RESUMEN_COMERCIAL_URL ? `?url=${encodeURIComponent(RESUMEN_COMERCIAL_URL)}` : '';
                const responseComercial = await fetch(`/api/downloadSyncComercial${paramsComercial}`, { signal: controller.signal });
                if (responseComercial.ok) {
                    const arrayBufferComercial = await responseComercial.arrayBuffer();
                    if (!window.resumenComercialEngine) {
                        try {
                            const engine = await import('./resumenComercialEngine.js');
                            window.resumenComercialEngine = engine;
                        } catch (e) {
                            console.error("Error importing resumenComercialEngine on demand:", e);
                        }
                    }
                    if (window.resumenComercialEngine) {
                        try {
                            const result = await window.resumenComercialEngine.processManualFile(arrayBufferComercial);
                            if (result) {
                                window.hasComercialAccess = true;
                            }
                        } catch (e) {
                            console.error("Error processing comercial sync:", e);
                        }
                    }
                }

                let paramsPg = PG_HORIZONTAL_URL ? `?url=${encodeURIComponent(PG_HORIZONTAL_URL)}` : '';
                const responsePg = await fetch(`/api/downloadSyncPgHorizontal${paramsPg}`, { signal: controller.signal });
                if (responsePg.ok) {
                    const arrayBufferPg = await responsePg.arrayBuffer();
                    if (!window.resumenComercialEngine) {
                        try {
                            const engine = await import('./resumenComercialEngine.js');
                            window.resumenComercialEngine = engine;
                        } catch (e) {
                            console.error("Error importing resumenComercialEngine on demand:", e);
                        }
                    }
                    if (window.resumenComercialEngine) {
                        try {
                            const dataPg = new Uint8Array(arrayBufferPg);
                            const workbookPg = XLSX.read(dataPg, { type: 'array' });
                            await window.resumenComercialEngine.processPgHorizontalWorkbook(workbookPg);
                            window.hasComercialAccess = true;
                        } catch (e) {
                            console.error("Error processing pg horizontal sync:", e);
                        }
                    }
                }

                let paramsCxp = CXP_URL ? `?url=${encodeURIComponent(CXP_URL)}` : '';
                const responseCxp = await fetch(`/api/downloadSyncCxp${paramsCxp}`, { signal: controller.signal });
                if (responseCxp.ok) {
                    const arrayBufferCxp = await responseCxp.arrayBuffer();
                    try {
                        const workbookCxp = XLSX.read(new Uint8Array(arrayBufferCxp), { type: 'array' });
                        await window.processCxpWorkbook(workbookCxp);
                        window.hasCxpAccess = true;
                    } catch (e) {
                        console.error("Error processing fallback cxp sync:", e);
                    }
                }

                let paramsCosto = COSTO_UNITARIO_URL ? `?url=${encodeURIComponent(COSTO_UNITARIO_URL)}` : '';
                const responseCosto = await fetch(`/api/downloadSyncCostoUnitario${paramsCosto}`, { signal: controller.signal });
                if (responseCosto.ok) {
                    const arrayBufferCosto = await responseCosto.arrayBuffer();
                    try {
                        const engine = await import('./costoUnitarioEngine.js');
                        window.costoUnitarioEngine = engine;
                        const workbook = XLSX.read(new Uint8Array(arrayBufferCosto), { type: 'array' });
                        window.costoUnitarioEngine.processCostoUnitarioWorkbook(workbook);
                        if (window.updateCostoUnitario) window.updateCostoUnitario();
                    } catch (e) { console.error("Error processing fallback costo sync:", e); }
                }
            }
            clearTimeout(timeoutId);
        } catch (err) {
            if (err.name === 'AbortError') console.warn("Tiempo de espera de red agotado.");
        }

        // Si operamos con caché (ya se cargaron los datos), restauramos los permisos que se perdieron con el reset
        if (window.isMagicLoaded) {
            // Restaurar hasMasterAccess ANTES de verificar globalFinancialData
            if (!window.hasMasterAccess) {
                const db = await getFinanceDB().catch(() => null);
                if (db) {
                    const cached = await new Promise(resolve => {
                        const tx = db.transaction('finance_cache', 'readonly');
                        const req = tx.objectStore('finance_cache').get('MASTER_FINANCE_KEY');
                        req.onsuccess = () => resolve(req.result);
                        req.onerror = () => resolve(null);
                    });
                    if (cached && cached.data && cached.data.length > 0) {
                        globalFinancialData = cached.data;
                        window.hasMasterAccess = true;
                    }
                }
            }
            if (!window.hasVentasAccess && (typeof ceoData !== 'undefined' && ceoData && ceoData.length > 0)) {
                window.hasVentasAccess = true;
            }
            if (!window.hasComercialAccess && (prevComercial || (window.resumenComercialEngine && typeof window.resumenComercialEngine.hasComercialData === 'function' && window.resumenComercialEngine.hasComercialData()))) {
                window.hasComercialAccess = true;
            }
        }

        // Aplicamos el RBAC después de las peticiones, pero si operamos con caché, mantenemos permisos
        if (typeof window.applyRoleBasedUI === 'function') {
            const hasCom = window.hasComercialAccess || false;
            window.applyRoleBasedUI(window.hasMasterAccess, window.hasVentasAccess, hasCom);
        }

        if (window._m365Interval) clearInterval(window._m365Interval);
        
        // Restore login button UI based on current auth state
        if (typeof window.updateM365UI === 'function') {
            const msalClient = window.msalInstance;
            const accounts = msalClient ? msalClient.getAllAccounts() : [];
            window.updateM365UI(accounts.length > 0 ? accounts[0] : null);
        }
        
        // Si no hay acceso a ninguno y no hay cache local, error general
        if (!arrayBuffer && !arrayBufferCeo) {
            if (window.isMagicLoaded) {
                if (statusEl) statusEl.innerHTML = "✅ Operando con Caché Local (Sin conexión nueva)";
                if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--success)';
                if (sidebarSyncText) sidebarSyncText.innerText = 'Caché Local';
                return;
            }
            throw new Error("No se pudo obtener ningún archivo fuente y no hay caché.");
        }

        // ==========================================
        // 3. PROCESAR CON WORKER
        // ==========================================
        if (arrayBuffer || window.isMagicLoaded) {
            if (arrayBuffer) {
                // Safeguard check - ensure SharePoint response isn't plain HTML / Login Redirect
                const previewText = new TextDecoder().decode(arrayBuffer.slice(0, 300));
                if (/^\s*<!doctype html/i.test(previewText) || /^\s*<html/i.test(previewText)) {
                    throw new Error("El archivo de Finanzas Master de SharePoint contiene HTML en vez de una hoja de cálculo (Revisar que esté compartido públicamente o iniciar sesión en M365).");
                }

                const engineResult = await new Promise((resolve, reject) => {
                    const worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
                    worker.onmessage = (e) => {
                        const data = e.data;
                        if (data.type === 'progress') {
                            if (loader && !window.isMagicLoaded) {
                                const lt = document.getElementById('loadingText');
                                if (lt) lt.innerText = data.message || "Procesando...";
                                const pb = document.getElementById('progressBar');
                                if (pb && data.progress) pb.style.width = `${data.progress}%`;
                            }
                        } else if (data.type === 'done') {
                            resolve(data.engineResult);
                            worker.terminate();
                        } else if (data.type === 'error') {
                            reject(new Error(data.error));
                            worker.terminate();
                        }
                    };
                    worker.onerror = (err) => {
                        reject(new Error(err.message || "Error fatal en el worker."));
                        worker.terminate();
                    };
                    worker.postMessage({ buffer: arrayBuffer }, [arrayBuffer]);
                });

                // 4. GUARDAR EN DISCO (INDEXEDDB) Y ACTUALIZAR UI SUAVEMENTE
                try {
                    const CACHE_KEY = 'MASTER_FINANCE_KEY';
                    const db = await getFinanceDB();
                    await new Promise((resolve, reject) => {
                        const tx = db.transaction('finance_cache', 'readwrite');
                        tx.objectStore('finance_cache').put({ data: engineResult.data, timestamp: Date.now() }, CACHE_KEY);
                        tx.oncomplete = resolve;
                        tx.onerror = reject;
                    });
                } catch (e) {
                    console.warn("⚠️ Error guardando caché en IndexedDB:", e);
                }

                globalFinancialData = engineResult.data;
                if (window.cachedStandaloneCxp && typeof window.applyCachedStandaloneCxp === 'function') {
                    await window.applyCachedStandaloneCxp();
                }
                window.isMagicLoaded = true;
                
                if (window.hasMasterAccess || window.isMagicLoaded) {
                    renderDashboard(globalFinancialData);
                }
            } else if (window.isMagicLoaded && globalFinancialData && globalFinancialData.length > 0) {
                // Fallback silencioso: usar caché existente aunque la API haya fallado
                console.log("⚡ [Fallback] Usando caché local tras fallo de API. Re-renderizando dashboard...");
                window.hasMasterAccess = !globalFinancialData._isMock;
                renderDashboard(globalFinancialData);
            }
        }
        
        if (loader) loader.style.display = 'none';
        if (arrayBuffer || arrayBufferCeo) {
            if (statusEl) statusEl.innerHTML = "✅ Sincronizado con O365";
            if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--success)';
            if (sidebarSyncText) {
                sidebarSyncText.innerText = 'Sincronizado';
                sidebarSyncText.style.color = 'var(--success)';
            }
        } else if (window.isMagicLoaded) {
            if (statusEl) {
                statusEl.innerHTML = "⚠️ Operando con Caché Local / Offline";
                statusEl.style.background = '#fef3c7';
                statusEl.style.color = '#92400e';
            }
            if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--warning, #f59e0b)';
            if (sidebarSyncText) {
                sidebarSyncText.innerText = 'Offline (Caché)';
                sidebarSyncText.style.color = 'var(--warning, #f59e0b)';
            }
        }
        if (window.updateLastUpdatedTime) {
            window.updateLastUpdatedTime();
        }

        // ==========================================
        // 5. PROCESAR VENTAS CEO
        // ==========================================
        if (arrayBufferCeo) {
            try {
                // Safeguard check - ensure SharePoint response isn't plain HTML / Login Redirect
                const previewText = new TextDecoder().decode(arrayBufferCeo.slice(0, 300));
                if (/^\s*<!doctype html/i.test(previewText) || /^\s*<html/i.test(previewText)) {
                    throw new Error("El archivo de Ventas CEO de SharePoint contiene HTML en vez de una hoja de cálculo (Revisar que esté compartido públicamente o iniciar sesión en M365).");
                }

                const resultCeo = await new Promise((resolve, reject) => {
                    const ceoWorker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
                    ceoWorker.onmessage = (e) => {
                        const data = e.data;
                        if (data.type === 'done_ventas') {
                            resolve(data.result);
                            ceoWorker.terminate();
                        } else if (data.type === 'error') {
                            reject(new Error(data.error));
                            ceoWorker.terminate();
                        }
                    };
                    ceoWorker.onerror = (err) => {
                        reject(new Error(err.message));
                        ceoWorker.terminate();
                    };
                    ceoWorker.postMessage({ buffer: arrayBufferCeo, fileType: 'ventas_ceo' }, [arrayBufferCeo]);
                });
                
                if (typeof window.processVentasCeoWorkbook === 'function') {
                    window.processVentasCeoWorkbook(null, null, resultCeo);
                }
            } catch (err) {
                console.warn("Fallo procesando Ventas CEO", err);
            }
        }
        
        ensureMockFinancialData();

    } catch (error) {
        console.error("Error en sincronización:", error);
        if (window._m365Interval) clearInterval(window._m365Interval);
        
        if (loader && !window.isMagicLoaded) loader.style.display = 'none';
        if (statusEl) {
            statusEl.style.background = '#fee2e2';
            statusEl.style.color = '#991b1b';
            statusEl.innerHTML = "⚠️ Sincronización fallida.";
        }
        
        // Si falló y tenemos caché, restauramos los accesos para que las vistas no desaparezcan
        if (window.isMagicLoaded) {
            if (!window.hasMasterAccess && (globalFinancialData && globalFinancialData.length > 0 && !globalFinancialData._isMock)) {
                window.hasMasterAccess = true;
            }
            if (!window.hasVentasAccess && (typeof ceoData !== 'undefined' && ceoData && ceoData.length > 0)) {
                window.hasVentasAccess = true;
            }
            if (!window.hasComercialAccess && ((typeof prevComercial !== 'undefined' && prevComercial) || (window.resumenComercialEngine && typeof window.resumenComercialEngine.hasComercialData === 'function' && window.resumenComercialEngine.hasComercialData()))) {
                window.hasComercialAccess = true;
            }
            if (typeof window.applyRoleBasedUI === 'function') {
                window.applyRoleBasedUI(window.hasMasterAccess, window.hasVentasAccess, window.hasComercialAccess);
            }
        } else {
            // Si falló y no tenemos caché, devolvemos a 0 al usuario para que no quede en pantalla blanca fantasma
            window.handleZeroState();
        }
    }
}
window.syncNavigationUI = function(menuId) {
    const titleLabel = document.getElementById('titleLabel');
    const titles = {
        'menu-kpi': "Torre de Control: Indicadores Clave",
        'menu-resumen': "Dashboard de Gestión Corporativa (RD$)",
        'menu-preliminar': "Estado de Resultado",
        'menu-pnl': "Estado de Resultados Detallado (RD$)",
        'menu-balance': "Balance General Consolidado (RD$)",
        'menu-cashflow': "Estado de Flujo de Efectivo (RD$)",
        'menu-deuda': "Zoom in Deuda (Millones DOP)",
        'menu-wc': "Capital de Trabajo (RD$)",
        'menu-cxp': "Detalle de Cuentas por Pagar (DOP)",
        'menu-estados': "Estados Financieros y KPIs (RD$)",
        'menu-simulador': "Simulador Estratégico (What-If)",
        'menu-pg-horizontal': "P&G Horizontal por Producto",
        'menu-ventas-ceo': "Ventas CEO",
        'menu-resumen-comercial': "Resumen Comercial",
        'menu-costo-unitario': "Costo Unitario",
        'menu-config': "Configuración y Auditoría",
        'menu-glosario': "Glosario de Términos y Metodologías Financieras",
        'menu-instructivo': "Manual del Usuario Corporativo"
    };
    if (titles[menuId] && titleLabel) titleLabel.textContent = titles[menuId];
};

window.handleZeroState = function() {
    const hasData = (globalFinancialData && globalFinancialData.length > 0) || 
                    (typeof ceoData !== 'undefined' && ceoData && ceoData.length > 0) || 
                    window.hasVentasAccess;
    
    const sidebar = document.querySelector('.sidebar');
    const headerActions = document.querySelector('.header-actions');
    const headerInfo = document.querySelector('.header-info');
    const dropZone = document.getElementById('dropZone');
    const viewContainers = document.querySelectorAll('.view-container');
    const mainContainer = document.querySelector('.main-container');
    const loginBtn = document.getElementById('loginM365Btn');

    // Siempre aplicar RBAC rules en vez de ocultar todo si hay acceso
    if (!hasData) {
        if(sidebar) sidebar.style.display = 'none';
        if(headerActions) headerActions.style.display = 'none';
        if(headerInfo) headerInfo.style.display = 'none';
        
        viewContainers.forEach(v => v.style.display = 'none');
        
        if (dropZone) {
            dropZone.style.display = 'block';
            dropZone.style.margin = '40px auto';
            if (mainContainer) mainContainer.appendChild(dropZone);
            if (loginBtn && !window.m365LoggedIn) loginBtn.style.display = 'flex';
        }
    } else {
        if(sidebar) sidebar.style.display = '';
        if(headerActions) headerActions.style.display = 'flex';
        if(headerInfo) headerInfo.style.display = 'block';
        
        viewContainers.forEach(v => v.style.display = '');

        const viewConfig = document.getElementById('view-config');
        const aiConfigPanel = document.getElementById('aiConfigPanel');
        if (dropZone && viewConfig) {
            dropZone.style.margin = '';
            if (aiConfigPanel && aiConfigPanel.nextSibling) {
                viewConfig.insertBefore(dropZone, aiConfigPanel.nextSibling);
            } else {
                viewConfig.appendChild(dropZone);
            }
        }
        
        if(typeof window.applyRoleBasedUI === 'function') {
           const inferredMaster = window.hasMasterAccess || (globalFinancialData && globalFinancialData.length > 0 && !globalFinancialData._isMock) ? true : false;
           const inferredVentas = window.hasVentasAccess || (typeof ceoData !== 'undefined' && ceoData && ceoData.length > 0) ? true : false;
           const inferredComercial = window.hasComercialAccess || false;
           window.applyRoleBasedUI(inferredMaster, inferredVentas, inferredComercial);
        }
    }
};

window.handleMSALLoginFailure = function() {
    const loginBtn = document.getElementById('loginM365Btn');
    if (loginBtn && !window.m365LoggedIn) loginBtn.style.display = 'flex';
    window.handleZeroState();
};

window.updateLastUpdatedTime = function(timestamp) {
    const lastUpdatedStatus = document.getElementById('lastUpdatedStatus');
    const lastUpdatedTime = document.getElementById('lastUpdatedTime');
    if (lastUpdatedStatus && lastUpdatedTime) {
        lastUpdatedStatus.style.display = 'block';
        const d = timestamp ? new Date(timestamp) : new Date();
        const dateStr = d.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: 'numeric' });
        const timeStr = d.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' });
        lastUpdatedTime.innerText = dateStr + ' ' + timeStr;
    }
};

async function loadCacheInstant() {
    try {
        const CACHE_VERSION = 'v5';
        if (localStorage.getItem('ventas_cache_version') !== CACHE_VERSION) {
            localStorage.setItem('ventas_cache_version', CACHE_VERSION);
            const db = await getFinanceDB();
            await new Promise((resolve, reject) => {
                const tx = db.transaction('finance_cache', 'readwrite');
                tx.objectStore('finance_cache').delete('MASTER_FINANCE_KEY');
                tx.objectStore('finance_cache').delete('CEO_VENTAS_KEY_V3');
                tx.objectStore('finance_cache').delete('CEO_VENTAS_KEY_V2');
                tx.objectStore('finance_cache').delete('COMERCIAL_KEY');
                tx.oncomplete = resolve;
                tx.onerror = reject;
            });
            console.log("Caché invalidado por nueva versión.");
            return false;
        }

        const CACHE_KEY = 'MASTER_FINANCE_KEY';
        const db = await getFinanceDB();

        const cachedRecord = await new Promise((resolve) => {
            const req = db.transaction('finance_cache', 'readonly').objectStore('finance_cache').get(CACHE_KEY);
            req.onsuccess = () => {
                const result = req.result;
                if (result) {
                    resolve(result);
                } else {
                    resolve(null);
                }
            };
            req.onerror = () => resolve(null);
        });
        
        // Also load Ventas CEO data
        const ceoCachedRecord = await new Promise((resolve) => {
            try {
                const req = db.transaction('finance_cache', 'readonly').objectStore('finance_cache').get('CEO_VENTAS_KEY_V3');
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => resolve(null);
            } catch (e) {
                resolve(null);
            }
        });

        const cxpCachedRecord = await new Promise((resolve) => {
            try {
                const req = db.transaction('finance_cache', 'readonly').objectStore('finance_cache').get('CXP_STANDALONE_KEY');
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => resolve(null);
            } catch (e) {
                resolve(null);
            }
        });
        if (cxpCachedRecord && cxpCachedRecord.data) {
            window.cxpStandaloneData = cxpCachedRecord.data; window.hasCxpAccess = true;
        }
        
        if (ceoCachedRecord && ceoCachedRecord.data && ceoCachedRecord.data.length > 0) {
            ceoData = ceoCachedRecord.data;
            if (typeof window.renderVentasCEO === 'function' && document.getElementById("view-ventas-ceo") && document.getElementById("view-ventas-ceo").classList.contains("active")) {
                window.renderVentasCEO();
            }
            window.isMagicLoaded = true;
        }

        // Try load Resumen Comercial as well
        try {
            const engine = await import('./resumenComercialEngine.js');
            window.resumenComercialEngine = engine;
            const hasComm = await engine.loadComercialCache();
            if (hasComm) {
                window.hasComercialAccess = true;
                window.isMagicLoaded = true;
                if (document.getElementById("view-resumen-comercial") && document.getElementById("view-resumen-comercial").classList.contains("active")) {
                    window.renderResumenComercial();
                }
                if (document.getElementById("view-pg-horizontal") && document.getElementById("view-pg-horizontal").classList.contains("active")) {
                    window.renderPgHorizontal();
                }
            }
        } catch(e) {}
        
        // Try load Costo Unitario as well
        try {
            const engineCU = await import('./costoUnitarioEngine.js');
            window.costoUnitarioEngine = engineCU;
            const hasCU = await engineCU.loadCostoUnitarioCache();
            if (hasCU) {
                window.hasCostoUnitarioData = () => true; 
                window.isMagicLoaded = true;
                if (document.getElementById("view-costo-unitario") && document.getElementById("view-costo-unitario").classList.contains("active")) {
                    if (typeof window.updateCostoUnitario === 'function') window.updateCostoUnitario();
                }
            }
        } catch(e) {}

        if (cachedRecord && cachedRecord.data && cachedRecord.data.length > 0) {
            console.log("🚀 Magic Load F5: Renderizando UI alzada instantáneamente.");
            window.isMagicLoaded = true; // 🔥 AÑADE ESTA LÍNEA AQUÍ
            globalFinancialData = cachedRecord.data;
            if (window.cachedStandaloneCxp && typeof window.applyCachedStandaloneCxp === 'function') {
                await window.applyCachedStandaloneCxp();
            }
            renderDashboard(globalFinancialData);
            if (window.updateLastUpdatedTime) {
                window.updateLastUpdatedTime(cachedRecord.timestamp);
            }
        }
        
        ensureMockFinancialData();

        if (window.isMagicLoaded) {
            if (typeof window.applyRoleBasedUI === 'function') {
                const inferredMaster = window.hasMasterAccess || (globalFinancialData && globalFinancialData.length > 0 && !globalFinancialData._isMock) ? true : false;
                const inferredVentas = window.hasVentasAccess || (typeof ceoData !== 'undefined' && ceoData && ceoData.length > 0) ? true : false;
                const inferredComercial = window.hasComercialAccess || false;
                window.applyRoleBasedUI(inferredMaster, inferredVentas, inferredComercial);
            }
            const loaderEl = document.getElementById('loader');
            if (loaderEl) loaderEl.style.display = 'none';
            return true;
        }
    } catch (e) {
        console.warn("⚠️ Magic Load omitido (caché no disponible):", e);
    }
    return false;
}

document.addEventListener('DOMContentLoaded', async () => {
    // 1. Escudo de Caché: Intenta cargar inmediatamente de SSD/IndexedDB
    const loadedFromCache = await loadCacheInstant();
    
    // 2. Si no hay caché, asegurar que se muestre Zero State
    if (!loadedFromCache) {
        window.handleZeroState();
    }

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', async () => {
            // Limpiar variables en memoria
            globalFinancialData = null;
            ceoData = null;
            window.isMagicLoaded = false;
            window.hasMasterAccess = false;
            window.hasVentasAccess = false;
            window.hasComercialAccess = false;
            window.m365LoggedIn = false;

            if (window.resumenComercialEngine && typeof window.resumenComercialEngine.resetComercialEngine === 'function') {
                window.resumenComercialEngine.resetComercialEngine();
            }
            
            // Limpiar caché
            try {
                const db = await getFinanceDB();
                await new Promise((resolve, reject) => {
                    const tx = db.transaction('finance_cache', 'readwrite');
                    tx.objectStore('finance_cache').delete('MASTER_FINANCE_KEY');
                    tx.objectStore('finance_cache').delete('CEO_VENTAS_KEY_V3');
                    tx.objectStore('finance_cache').delete('COMERCIAL_KEY');
                    tx.oncomplete = resolve;
                    tx.onerror = reject;
                });
            } catch (e) { console.error('Error clearing cache', e); }

            // Cerrar sesion MSAL
            if (msalInstance) {
                const activeAcc = msalInstance.getActiveAccount();
                if (activeAcc) {
                    await msalInstance.logoutPopup({ account: activeAcc }).catch(console.error);
                }
            }

            // Cambiar UI al estado inicial
            window.handleZeroState();
            
            // Update sidebar dot
            const sidebarSyncDot = document.getElementById('sidebarSyncDot');
            const sidebarSyncText = document.getElementById('sidebarSyncText');
            if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--warning)';
            if (sidebarSyncText) {
                sidebarSyncText.innerText = 'Desconectado';
                sidebarSyncText.style.color = 'rgba(255, 255, 255, 0.7)';
            }
        });
    }

    // MSAL Background Synchronization (No Bloqueante)
    if (msalInstance) {
        msalInstance.initialize?.().then(async () => {
            try {
                const redirectResponse = await msalInstance.handleRedirectPromise();
                if (redirectResponse) {
                    window.history.replaceState({}, document.title, window.location.pathname);
                    window.m365LoggedIn = true;
                    fetchMasterData(redirectResponse.accessToken);
                    return;
                }
            } catch (err) {
                console.error("MSAL Redirect Error:", err);
            }

            // Mandatory Silent Flow
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.setActiveAccount(accounts[0]);
                window.updateM365UI(accounts[0]);
                try {
                    const response = await msalInstance.acquireTokenSilent({
                        scopes: ["User.Read", "Files.Read", "Files.Read.All"],
                        account: accounts[0]
                    });
                    
                    // Background Update
                    window.m365LoggedIn = true;
                    fetchMasterData(response.accessToken);
                } catch (error) {
                    console.warn("Silent login failed (Token expire/cache missing):", error);
                    // Do not redirect to interactive login automatically to prevent interrupting the cached UI
                    // Only prompt if we have no cached data at all.
                    if (!loadedFromCache) {
                        console.log("Usuario conocido pero token expirado y sin caché: Se requiere inicio de sesión interactivo.");
                        // Evitar Redirect automático en iframes
                        if (window.parent === window) {
                            msalInstance.loginRedirect({
                                scopes: ["User.Read", "Files.Read", "Files.Read.All"]
                            });
                        } else {
                            console.warn("Redirección automática omitida debido a entorno de iFrame.");
                        }
                    } else {
                        const sidebarSyncText = document.getElementById('sidebarSyncText');
                        const sidebarSyncDot = document.getElementById('sidebarSyncDot');
                        if (sidebarSyncText) {
                            sidebarSyncText.innerText = 'Sesión expirada';
                            sidebarSyncText.style.color = '#ef4444';
                        }
                        if (sidebarSyncDot) {
                            sidebarSyncDot.style.backgroundColor = '#ef4444';
                        }
                    }
                }
            } else if (!loadedFromCache) {
                console.log("Usuario nuevo sin caché: Se requiere inicio de sesión interactivo.");
                // Evitar Redirect automático en iframes
                if (window.parent === window) {
                    msalInstance.loginRedirect({
                        scopes: ["User.Read", "Files.Read", "Files.Read.All"]
                    });
                } else {
                    console.warn("Redirección automática omitida debido a entorno de iFrame.");
                }
            }
        }).catch(err => {
             console.error("MSAL Initialization failed:", err);
             if (!loadedFromCache) {
                 window.handleMSALLoginFailure();
             }
        });
    } else if (!loadedFromCache) {
        window.handleMSALLoginFailure();
    }
    
    // Wire up interactive clickable sidebar sync status action
    const sidebarSyncStatus = document.getElementById('sidebarSyncStatus');
    if (sidebarSyncStatus) {
        sidebarSyncStatus.style.cursor = 'pointer';
        sidebarSyncStatus.title = "Haz clic para conectar con Microsoft 365 y sincronizar en tiempo real";
        sidebarSyncStatus.addEventListener('click', async () => {
            if (typeof msalInstance !== 'undefined' && msalInstance) {
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    try {
                        const silentResp = await msalInstance.acquireTokenSilent({
                            scopes: ["User.Read", "Files.Read", "Files.Read.All"],
                            account: accounts[0]
                        });
                        alert("Sincronizando de forma directa con Microsoft 365... (Garantiza velocidad ultra-rápida)");
                        await fetchMasterData(silentResp.accessToken);
                    } catch (e) {
                        console.error("Silent sync failed, prompt interactive logon:", e);
                        await connectM365();
                    }
                } else {
                    await connectM365();
                }
            } else {
                alert("Microsoft MSAL no está listo en esta sesión de navegador.");
            }
        });
    }

    // Wire up Office 365 Sync and Settings UI actions
    const loginM365Btn = document.getElementById('loginM365Btn');
    if (loginM365Btn) {
        loginM365Btn.addEventListener('click', connectM365);
    }

    const logoutM365Btn = document.getElementById('logoutM365Btn');
    if (logoutM365Btn) {
        logoutM365Btn.addEventListener('click', async () => {
            if (msalInstance) {
                const activeAcc = msalInstance.getActiveAccount();
                if (activeAcc) {
                    await msalInstance.logoutPopup({ account: activeAcc }).catch(console.error);
                }
            }
            window.m365LoggedIn = false;
            window.updateM365UI(null);
            window.handleZeroState();
        });
    }

    const saveM365UrlsBtn = document.getElementById('saveM365UrlsBtn');
    if (saveM365UrlsBtn) {
        saveM365UrlsBtn.addEventListener('click', () => {
            const masterUrl = document.getElementById('m365UrlMaster')?.value || '';
            const ventasUrl = document.getElementById('m365UrlVentas')?.value || '';
            const comercialUrl = document.getElementById('m365UrlComercial')?.value || '';
            const pgHorizontalUrl = document.getElementById('m365UrlPgHorizontal')?.value || '';
            const cxpUrl = document.getElementById('m365UrlCxp')?.value || '';
            const costoUnitarioUrl = document.getElementById('m365UrlCostoUnitario')?.value || '';
            
            localStorage.setItem('CUSTOM_ONEDRIVE_FILE_URL', masterUrl);
            localStorage.setItem('CUSTOM_ONEDRIVE_VENTAS_URL', ventasUrl);
            localStorage.setItem('CUSTOM_RESUMEN_COMERCIAL_URL', comercialUrl);
            localStorage.setItem('CUSTOM_PG_HORIZONTAL_URL', pgHorizontalUrl);
            localStorage.setItem('CUSTOM_CXP_URL', cxpUrl);
            localStorage.setItem('CUSTOM_COSTO_UNITARIO_URL', costoUnitarioUrl);
            
            SHARPOINT_FILE_URL = masterUrl || import.meta.env.VITE_ONEDRIVE_FILE_URL || import.meta.env.VITE_ONEDRIVE_ITEM_ID;
            SHARPOINT_VENTAS_FILE_URL = ventasUrl || import.meta.env.VITE_CEO_FILE_URL;
            RESUMEN_COMERCIAL_URL = comercialUrl || import.meta.env.VITE_RESUMEN_COMERCIAL_URL;
            PG_HORIZONTAL_URL = pgHorizontalUrl || import.meta.env.VITE_PG_HORIZONTAL_URL;
            CXP_URL = cxpUrl || import.meta.env.VITE_CXP_URL;
            COSTO_UNITARIO_URL = costoUnitarioUrl || import.meta.env.VITE_COSTO_UNITARIO_URL;
            
            alert("Enlaces de OneDrive guardados exitosamente.");
        });
    }

    const syncNowM365Btn = document.getElementById('syncNowM365Btn');
    if (syncNowM365Btn) {
        syncNowM365Btn.addEventListener('click', async () => {
            if (!msalInstance) return;
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                try {
                    const response = await msalInstance.acquireTokenSilent({
                        scopes: ["User.Read", "Files.Read", "Files.Read.All"],
                        account: accounts[0]
                    });
                    await fetchMasterData(response.accessToken);
                    alert("Datos sincronizados exitosamente desde OneDrive.");
                } catch (e) {
                    console.error("Manual sync failed, attempting interactive login:", e);
                    await connectM365();
                }
            } else {
                await connectM365();
            }
        });
    }
    const fileInput = document.getElementById('fileInput');
    const dropZone = document.getElementById('dropZone');

    // Setup Export and Mobile Menu
    const btnExportCSV = document.getElementById('btn-export-csv');
    const bindCsvExport = (btn) => {
        if (!btn) return;
        btn.addEventListener('click', () => {
            const mainContainer = document.querySelector('.main-container');
            const views = mainContainer.querySelectorAll('.view-container');
            let activeViewId = 'view-resumen';
            if (views && views.length > 0) {
                views.forEach(v => {
                    if (v.classList.contains('active')) activeViewId = v.id;
                });
            }

            let dataToExport = [];
            let prefix = 'reporte';
            const dateStr = new Date().toISOString().slice(0, 7);

            if (activeViewId === 'view-ventas-ceo') {
                if (!ceoData || ceoData.length === 0) {
                    alert('No hay datos de Ventas CEO para exportar.');
                    return;
                }
                dataToExport = ceoData.map(row => {
                    const flatObj = { Producto: row.Producto, Tipo: row.Tipo };
                    if (row.values) {
                        for (const key in row.values) {
                            flatObj[key] = row.values[key];
                        }
                    }
                    return flatObj;
                });
                prefix = 'reporte_ventas_ceo';
            } else {
                if (!globalFinancialData || globalFinancialData.length === 0) {
                    alert('No hay datos financieros para exportar.');
                    return;
                }

                function flattenObject(ob, prefix = '') {
                    const result = {};
                    for (const i in ob) {
                        if (Object.prototype.hasOwnProperty.call(ob, i)) {
                            if ((typeof ob[i]) === 'object' && ob[i] !== null && !Array.isArray(ob[i])) {
                                const flatObject = flattenObject(ob[i], prefix + i + '_');
                                for (const x in flatObject) {
                                    if (Object.prototype.hasOwnProperty.call(flatObject, x)) {
                                        result[x] = flatObject[x];
                                    }
                                }
                            } else {
                                result[prefix + i] = ob[i];
                            }
                        }
                    }
                    return result;
                }

                dataToExport = globalFinancialData.map(d => {
                    const flat = flattenObject(d);
                    const result = { Periodo: d.date };
                    for (const key in flat) {
                        if (key !== 'date') {
                            result[key] = flat[key];
                        }
                    }
                    return result;
                });
                prefix = 'reporte_planeta_azul';
            }

            const filename = `${prefix}_${dateStr}.csv`;
            downloadCSV(dataToExport, filename);
        });
    };
    
    bindCsvExport(btnExportCSV);

    const btnExportPDF = document.getElementById('btn-export-pdf');
    const bindPdfExport = (btn) => {
        if (!btn) return;
        btn.addEventListener('click', async () => {
            btn.disabled = true;
            const originalText = btn.innerHTML;
            btn.innerHTML = '<i data-lucide="loader" class="spin-icon" style="width: 16px; height: 16px; display: inline-block; vertical-align: middle; margin-right: 4px;"></i> Generando Master PDF...';
            if (typeof lucide !== 'undefined') lucide.createIcons();

            const previousViewId = document.querySelector('.view-container.active')?.id || 'view-config';
            const ytdCheckbox = document.getElementById('ytdToggle');
            const previousYTD = ytdCheckbox ? ytdCheckbox.checked : false;
            
            // Scroll to top before rendering to prevent html2canvas clipping sticky elements
            window.scrollTo(0, 0);

            const sleep = (ms) => new Promise(r => setTimeout(r, ms));
            let pdf = null;
            let isFirstPage = true;

            const addPageToPDF = async (forcedSubtitle = "") => {
                const layoutWrapper = document.querySelector('.layout-wrapper');
                const contentToRender = document.querySelector('.main-container');
                if (!contentToRender) return;

                // Expand parent fully to avoid cropping
                let layoutWrapperOrigHeight = "";
                let layoutWrapperOrigOverflow = "";
                if (layoutWrapper) {
                    layoutWrapperOrigHeight = layoutWrapper.style.height;
                    layoutWrapperOrigOverflow = layoutWrapper.style.overflow;
                    layoutWrapper.style.height = "auto";
                    layoutWrapper.style.overflow = "visible";
                }

                const headerActions = document.querySelector('.header-actions');
                let originalHeaderDisplay = '';
                if (headerActions) {
                    originalHeaderDisplay = headerActions.style.display;
                    headerActions.style.display = 'none';
                }
                
                // Keep perspective buttons visible so it's clear what's being viewed
                let pnlControls = contentToRender.querySelector('.pnl-controls');
                let pnlControlsDisplay = '';
                if (pnlControls) {
                    pnlControlsDisplay = pnlControls.style.display;
                    pnlControls.style.display = 'none';
                }

                const originalOverflow = contentToRender.style.overflow;
                const originalHeight = contentToRender.style.height;
                const originalPaddingBottom = contentToRender.style.paddingBottom;
                contentToRender.style.overflow = "visible";
                contentToRender.style.height = "max-content";
                contentToRender.style.paddingBottom = "100px";

                const expandStyles = document.createElement('style');
                expandStyles.id = 'pdf-expand-style';
                expandStyles.innerHTML = `
                    .main-container, .layout-wrapper {
                        width: max-content !important;
                        min-width: 100% !important;
                        height: auto !important;
                        max-height: none !important;
                        overflow: visible !important;
                    }
                    .view-container.active, .card, .pnl-detail-table, .table-container, .chart-box {
                        height: auto !important;
                        max-height: none !important;
                        overflow: visible !important;
                    }
                    table {
                        width: 100% !important; 
                        max-width: none !important;
                    }
                `;
                document.head.appendChild(expandStyles);
                
                await sleep(150); // allow layout to recalculate
                
                // Esperar a que el layout se estabilice (altura deja de cambiar)
                await new Promise(resolve => {
                    let lastHeight = 0;
                    let stable = 0;
                    const check = () => {
                        const h = contentToRender.scrollHeight;
                        if (h === lastHeight) { stable++; } else { stable = 0; lastHeight = h; }
                        if (stable >= 3) { resolve(); } else { requestAnimationFrame(check); }
                    };
                    requestAnimationFrame(check);
                });
                
                const header = document.getElementById('mainHeader');
                let origHeaderPos = "";
                if (header) {
                    origHeaderPos = header.style.position;
                    header.style.position = 'static'; 
                }

                // Add Subtitle if provided
                const titleLabel = document.getElementById('titleLabel');
                let origTitleHTML = "";
                if (titleLabel && forcedSubtitle) {
                    origTitleHTML = titleLabel.innerHTML;
                    titleLabel.innerHTML += ` <span style="font-size: 0.7em; color: white; background-color: var(--primary); padding: 4px 8px; border-radius: 8px; margin-left: 10px; vertical-align: middle;">${forcedSubtitle}</span>`;
                }

                // Use solid white background to avoid pale/veil effect from lighter grey
                const bgColor = "#ffffff";

                // Ensure the capture window captures the true scrollable width
                const fullWidth = Math.max(document.body.scrollWidth, contentToRender.scrollWidth);

                const canvas = await html2canvas(contentToRender, {
                    scale: 2, 
                    useCORS: true,
                    logging: false,
                    backgroundColor: bgColor,
                    windowWidth: fullWidth,
                    windowHeight: Math.max(document.body.scrollHeight, contentToRender.scrollHeight) + 200,
                    onclone: (clonedDoc) => {
                         // Fix the animation veil: animations mid-flight cause elements to have 50% opacity
                         const style = clonedDoc.createElement('style');
                         style.innerHTML = `
                            * { 
                                animation: none !important; 
                                transition: none !important; 
                            }
                            .main-container, .layout-wrapper {
                                width: max-content !important;
                                min-width: 100% !important;
                                height: auto !important;
                                max-height: none !important;
                                overflow: visible !important;
                            }
                            .view-container.active, .card, .section-table, .pnl-detail-table, .table-container, .chart-box {
                                height: auto !important;
                                max-height: none !important;
                                overflow: visible !important;
                            }
                            table {
                                width: 100% !important; 
                                max-width: none !important;
                            }
                         `;
                         clonedDoc.head.appendChild(style);

                         const mainCont = clonedDoc.querySelector('.main-container');
                         if (mainCont) {
                             mainCont.style.overflow = "visible";
                             mainCont.style.height = "max-content";
                             mainCont.style.maxHeight = "none";
                             mainCont.style.paddingTop = "40px";
                             mainCont.style.paddingBottom = "140px";
                         }
                         const activeTab = clonedDoc.querySelector('.view-container.active');
                         if (activeTab) {
                             activeTab.style.overflow = "visible"; 
                             activeTab.style.height = "max-content";
                             activeTab.style.maxHeight = "none";
                             activeTab.style.padding = "20px";
                             
                             // Darken secondary text slightly so it's not faded in PDF
                             const textSecs = activeTab.querySelectorAll('[style*="var(--text-secondary)"]');
                             textSecs.forEach(el => {
                                 if (el.style) el.style.color = "#475569";
                             });

                             // Prevent tables from cutting off
                             const tableCont = activeTab.querySelector('.pnl-detail-table');
                             if (tableCont) {
                                 tableCont.style.overflow = "visible";
                                 tableCont.style.maxHeight = "none";
                                 tableCont.style.height = "auto";
                             }
                             
                             const ventasList = activeTab.querySelectorAll('.pnl-detail-table');
                             ventasList.forEach(t => {
                                 t.style.overflow = "visible";
                                 t.style.maxHeight = "none";
                             });

                             // Forzar visibilidad de todos los elementos de la vista activa
                             const allChildren = activeTab ? activeTab.querySelectorAll('*') : [];
                             allChildren.forEach(el => {
                                 const cs = window.getComputedStyle(el);
                                 if (cs.display === 'none' && !el.closest('.header-actions') && !el.closest('.pnl-controls')) {
                                     // No forzar; solo loguear para debug
                                 }
                             });
                             // Asegurar que el contenedor raíz de la vista tenga dimensiones reales
                             if (activeTab) {
                                 activeTab.style.minHeight = activeTab.scrollHeight + 'px';
                                 activeTab.style.minWidth = activeTab.scrollWidth + 'px';
                                 activeTab.style.paddingBottom = '100px';
                             }
                             const clonedMain = clonedDoc.querySelector('.main-container');
                             if (clonedMain) {
                                 clonedMain.style.paddingBottom = '100px';
                             }
                         }
                    }
                });

                // RESTORE everything
                if (titleLabel && forcedSubtitle) titleLabel.innerHTML = origTitleHTML;
                if (header) header.style.position = origHeaderPos;
                contentToRender.style.overflow = originalOverflow;
                contentToRender.style.height = originalHeight;
                contentToRender.style.paddingBottom = originalPaddingBottom;
                if (layoutWrapper) {
                    layoutWrapper.style.height = layoutWrapperOrigHeight;
                    layoutWrapper.style.overflow = layoutWrapperOrigOverflow;
                }
                if (headerActions) headerActions.style.display = originalHeaderDisplay;
                if (pnlControls) pnlControls.style.display = pnlControlsDisplay;

                const styleEl = document.getElementById('pdf-expand-style');
                if (styleEl) styleEl.remove();

                const imgData = canvas.toDataURL('image/jpeg', 0.95);
                
                const pdfWidth = 420; // 420mm -> roughly A3 width
                const ratio = canvas.height / canvas.width;
                const pdfHeight = pdfWidth * ratio;
                const pageFormat = [pdfWidth, pdfHeight];
                const orientation = pdfWidth > pdfHeight ? 'landscape' : 'portrait';

                if (!pdf) {
                    pdf = new jsPDF({ orientation: orientation, unit: 'mm', format: pageFormat });
                    pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, 'MEDIUM');
                } else {
                    pdf.addPage(pageFormat, orientation);
                    pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight, undefined, 'MEDIUM');
                }
            };

            const toggleYTD = (enable) => {
                if (ytdCheckbox && ytdCheckbox.checked !== enable) {
                    ytdCheckbox.click();
                } else if (typeof updateUI === 'function') {
                    const sel = document.getElementById('monthSelector');
                    const idx = sel ? parseInt(sel.value) : NaN;
                    if (!isNaN(idx) && globalFinancialData) {
                        updateUI(globalFinancialData, idx);
                    }
                }
            };
            
            const clickElement = (id) => {
                const el = document.getElementById(id);
                if (el) el.click();
            };
            
            const showViewAndSync = (viewId, menuId) => {
                document.querySelectorAll('.view-container').forEach(v => v.classList.remove('active'));
                const v = document.getElementById(viewId);
                if (v) v.classList.add('active');
                if (window.syncNavigationUI) window.syncNavigationUI(menuId);
                
                if (typeof renderActiveViewLazy === 'function' && typeof globalFinancialData !== 'undefined' && globalFinancialData) {
                    const sel = document.getElementById('monthSelector');
                    if (sel) {
                        const idx = parseInt(sel.value);
                        if (!isNaN(idx)) renderActiveViewLazy(globalFinancialData, idx);
                    }
                }
            };

                try {
                // 1. KPIs
                if (document.getElementById('view-kpi')) {
                    showViewAndSync('view-kpi', 'menu-kpi'); 
                    await sleep(800); 
                    await addPageToPDF("Dashboard | " + (previousYTD ? "YTD" : "Mensual"));
                }

                // 2. Resumen Ejecutivo
                if (document.getElementById('view-resumen')) {
                    showViewAndSync('view-resumen', 'menu-resumen'); 
                    await sleep(800); 
                    await addPageToPDF("Gestión Corporativa | " + (previousYTD ? "YTD" : "Mensual"));
                }

                // 3. Resumen Comercial
                if (document.getElementById('view-resumen-comercial')) {
                    showViewAndSync('view-resumen-comercial', 'menu-resumen-comercial');
                    
                    toggleYTD(false); clickElement('btn-comercial-resumen'); await sleep(800); await addPageToPDF("Mensual | Resumen de Ventas");
                    toggleYTD(true); clickElement('btn-comercial-resumen'); await sleep(800); await addPageToPDF("YTD | Resumen de Ventas");
                    
                    toggleYTD(false); clickElement('btn-comercial-variacion'); await sleep(800); await addPageToPDF("Mensual | Análisis de Variación (vs PPTO / vs AA)");
                    toggleYTD(true); clickElement('btn-comercial-variacion'); await sleep(800); await addPageToPDF("YTD | Análisis de Variación (vs PPTO / vs AA)");
                    
                    toggleYTD(false); clickElement('btn-comercial-mom'); await sleep(800); await addPageToPDF("Mensual | Variación MoM");
                }

                // 4. Estado de Resultado (Mensual y YTD)
                if (document.getElementById('view-preliminar')) {
                    showViewAndSync('view-preliminar', 'menu-preliminar');
                    toggleYTD(false); await sleep(800); await addPageToPDF("Mensual");
                    toggleYTD(true); await sleep(800); await addPageToPDF("YTD");
                }

                // 5. Balance Sheet
                if (document.getElementById('view-balance')) {
                    showViewAndSync('view-balance', 'menu-balance');
                    toggleYTD(previousYTD);
                    clickElement('btn-balance-resumen');
                    await sleep(800);
                    await addPageToPDF("Balance Sheet | Resumen | " + (previousYTD ? "YTD" : "Mensual"));
                }

                // 6. Cash Flow
                if (document.getElementById('view-cashflow')) {
                    showViewAndSync('view-cashflow', 'menu-cashflow');
                    toggleYTD(previousYTD);
                    clickElement('btn-cashflow-resumen');
                    await sleep(800);
                    await addPageToPDF("Cash Flow | Resumen | " + (previousYTD ? "YTD" : "Mensual"));
                }

                // 6.5. Detalle CxP
                if (document.getElementById('view-cxp')) {
                    showViewAndSync('view-cxp', 'menu-cxp');

                    clickElement('btn-cxp-resumen');
                    await sleep(800);
                    await addPageToPDF("Resumen de Cuentas por Pagar (DOP)");
                }

                // 6.6. Zoom in Deuda
                if (document.getElementById('view-deuda')) {
                    showViewAndSync('view-deuda', 'menu-deuda');
                    await sleep(800);
                    await addPageToPDF("Zoom in Deuda");
                }

                // 7. Ventas CEO
                if (document.getElementById('view-ventas-ceo')) {
                    showViewAndSync('view-ventas-ceo', 'menu-ventas-ceo');
                    toggleYTD(previousYTD);
                    clickElement('btn-ventas-vol'); await sleep(800); await addPageToPDF("Métrica: Volumen (k de Unidades)");
                    clickElement('btn-ventas-monto'); await sleep(800); await addPageToPDF("Métrica: Monto (mDOP)");
                    clickElement('btn-ventas-precio'); await sleep(800); await addPageToPDF("Métrica: Precio Unitario");
                }

                // 8. Costo Unitario
                if (document.getElementById('view-costo-unitario') && typeof window.hasCostoUnitarioData === 'function' && window.hasCostoUnitarioData()) {
                    showViewAndSync('view-costo-unitario', 'menu-costo-unitario');
                    
                    clickElement('btn-costo-botellon');
                    clickElement('btn-costo-vista-resumen');
                    await sleep(800);
                    await addPageToPDF("Costo Unitario | Botellón | Resumen");
                    
                    clickElement('btn-costo-botella');
                    clickElement('btn-costo-vista-resumen');
                    await sleep(800);
                    await addPageToPDF("Costo Unitario | Botella 0.5 LTS | Resumen");
                }

                const dateStr = new Date().toISOString().slice(0, 10);
                pdf.save(`Reportes_Ejecutivos_Maestros_${dateStr}.pdf`);

            } catch(e) {
                console.error("Error generating Master PDF:", e);
                alert("Ocurrió un error al generar el PDF: " + String(e.message || e));
            } finally {
                toggleYTD(previousYTD);
                
                // Keep the menuid sync in check for restoration
                const navKeyMap = {
                    'view-kpi': 'menu-kpi', 'view-resumen': 'menu-resumen', 'view-preliminar': 'menu-preliminar',
                    'view-resumen-comercial': 'menu-resumen-comercial', 'view-ventas-ceo': 'menu-ventas-ceo',
                    'view-pg-horizontal': 'menu-pg-horizontal', 'view-costo-unitario': 'menu-costo-unitario', 'view-config': 'menu-config',
                    'view-deuda': 'menu-deuda', 'view-wc': 'menu-wc', 'view-cxp': 'menu-cxp',
                    'view-balance': 'menu-balance', 'view-cashflow': 'menu-cashflow', 'view-estados': 'menu-estados'
                };
                showViewAndSync(previousViewId, navKeyMap[previousViewId]);

                btn.disabled = false;
                btn.innerHTML = originalText;
                if (typeof lucide !== 'undefined') lucide.createIcons();
            }
        });
    };
    
    bindPdfExport(btnExportPDF);

    const menuToggleBtn = document.getElementById('menuToggleBtn');
    const sidebar = document.querySelector('.sidebar');
    if (menuToggleBtn && sidebar) {
        menuToggleBtn.addEventListener('click', () => {
            sidebar.classList.toggle('open');
        });
    }

    let pendingMainFile = null;
    let pendingVentasCeoFile = null;
    let pendingResumenComercialFile = null;
    let pendingPgHorizontalFile = null;
    let pendingCxpFile = null;
    let pendingCostoUnitarioFile = null;

    function updateProcessButton() {
        const btn = document.getElementById('processManualFilesBtn');
        if (btn) {
            if (pendingMainFile || pendingVentasCeoFile || pendingResumenComercialFile || pendingPgHorizontalFile || pendingCxpFile || pendingCostoUnitarioFile) {
                btn.disabled = false;
                btn.style.opacity = '1';
                btn.style.cursor = 'pointer';
            } else {
                btn.disabled = true;
                btn.style.opacity = '0.5';
                btn.style.cursor = 'not-allowed';
            }
        }
    }

    function makeSafeFile(file) {
        if (!file) return null;
        const safeBufferPromise = file.arrayBuffer().catch(err => {
            console.error("Error pre-reading file:", err);
            throw err;
        });
        return {
            name: file.name,
            size: file.size,
            type: file.type,
            arrayBuffer: () => safeBufferPromise
        };
    }

    if (fileInput) {
        fileInput.addEventListener('change', (e) => {
            const file = makeSafeFile(e.target.files[0]);
            if (file) {
                pendingMainFile = file;
                const nameEl = document.getElementById('fileInputName');
                if (nameEl) nameEl.textContent = file.name;
                updateProcessButton();
            }
        });
    }

    const uploadVentasCeoHomeInput = document.getElementById('upload-ventas-ceo-home');
    if (uploadVentasCeoHomeInput) {
        uploadVentasCeoHomeInput.addEventListener('change', (e) => {
            const file = makeSafeFile(e.target.files[0]);
            if (file) {
                pendingVentasCeoFile = file;
                const nameEl = document.getElementById('ventasCeoFileName');
                if (nameEl) nameEl.textContent = file.name;
                updateProcessButton();
            }
        });
    }

    const uploadResumenComercialHomeInput = document.getElementById('upload-resumen-comercial-home');
    if (uploadResumenComercialHomeInput) {
        uploadResumenComercialHomeInput.addEventListener('change', (e) => {
            const file = makeSafeFile(e.target.files[0]);
            if (file) {
                pendingResumenComercialFile = file;
                const nameEl = document.getElementById('resumenComercialFileName');
                if (nameEl) nameEl.textContent = file.name;
                updateProcessButton();
            }
        });
    }

    const uploadPgHorizontalHomeInput = document.getElementById('upload-pg-horizontal-home');
    if (uploadPgHorizontalHomeInput) {
        uploadPgHorizontalHomeInput.addEventListener('change', (e) => {
            const file = makeSafeFile(e.target.files[0]);
            if (file) {
                pendingPgHorizontalFile = file;
                const nameEl = document.getElementById('pgHorizontalFileName');
                if (nameEl) nameEl.textContent = file.name;
                updateProcessButton();
            }
        });
    }

    const uploadCxpHomeInput = document.getElementById('upload-cxp-home');
    if (uploadCxpHomeInput) {
        uploadCxpHomeInput.addEventListener('change', (e) => {
            const file = makeSafeFile(e.target.files[0]);
            if (file) {
                pendingCxpFile = file;
                const nameEl = document.getElementById('cxpFileName');
                if (nameEl) nameEl.textContent = file.name;
                updateProcessButton();
            }
        });
    }

    const uploadCostoUnitarioHomeInput = document.getElementById('upload-costo-unitario-home');
    if (uploadCostoUnitarioHomeInput) {
        uploadCostoUnitarioHomeInput.addEventListener('change', (e) => {
            const file = makeSafeFile(e.target.files[0]);
            if (file) {
                pendingCostoUnitarioFile = file;
                const nameEl = document.getElementById('costoUnitarioFileName');
                if (nameEl) nameEl.textContent = file.name;
                updateProcessButton();
            }
        });
    }

    const processManualFilesBtn = document.getElementById('processManualFilesBtn');
    if (processManualFilesBtn) {
        processManualFilesBtn.addEventListener('click', async () => {
            if (pendingVentasCeoFile) {
                await window.processVentasCeoFile(pendingVentasCeoFile);
            }
            if (pendingResumenComercialFile) {
                await window.processResumenComercialFile(pendingResumenComercialFile);
            }
            if (pendingPgHorizontalFile) {
                await window.processPgHorizontalFile(pendingPgHorizontalFile);
            }
            if (pendingCxpFile) {
                await window.processCxpFile(pendingCxpFile);
            }
            if (pendingCostoUnitarioFile) {
                await window.processCostoUnitarioFile(pendingCostoUnitarioFile);
            }
            if (pendingMainFile) {
                handleFileUpload({ target: { files: [pendingMainFile] } });
            } else if (pendingVentasCeoFile || pendingResumenComercialFile || pendingPgHorizontalFile || pendingCxpFile || pendingCostoUnitarioFile) {
                alert("Archivos secundarios procesados exitosamente.");
            }
        });
    }

    if (dropZone) {
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = 'var(--primary)';
            dropZone.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.style.borderColor = 'rgba(0, 150, 199, 0.4)';
            dropZone.style.background = 'transparent';
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = 'rgba(0, 150, 199, 0.4)';
            dropZone.style.background = 'transparent';
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                const originalFile = files[0];
                const safeBufferPromise = originalFile.arrayBuffer();
                const file = {
                    name: originalFile.name,
                    size: originalFile.size,
                    type: originalFile.type,
                    arrayBuffer: () => safeBufferPromise
                };
                
                const name = file.name.toLowerCase();
                
                // Intento heurístico:
                if (name.includes('ventas') && name.includes('ceo')) {
                    if (uploadVentasCeoHomeInput) uploadVentasCeoHomeInput.files = files;
                    pendingVentasCeoFile = file;
                    const nameEl = document.getElementById('ventasCeoFileName');
                    if (nameEl) nameEl.textContent = file.name;
                } else if (name.includes('resumen') || name.includes('comercial')) {
                    const uploadResumenBtn = document.getElementById('upload-resumen-comercial-home');
                    if (uploadResumenBtn) uploadResumenBtn.files = files;
                    pendingResumenComercialFile = file;
                    const nameEl = document.getElementById('resumenComercialFileName');
                    if (nameEl) nameEl.textContent = file.name;
                } else if (name.includes('p&g') || name.includes('horizontal') || name.includes('p_g') || (name.includes('p') && name.includes('g') && name.includes('horizontal'))) {
                    const uploadPgHorizontalBtn = document.getElementById('upload-pg-horizontal-home');
                    if (uploadPgHorizontalBtn) uploadPgHorizontalBtn.files = files;
                    pendingPgHorizontalFile = file;
                    const nameEl = document.getElementById('pgHorizontalFileName');
                    if (nameEl) nameEl.textContent = file.name;
                } else if (name.includes('cxp') || name.includes('pagar') || name.includes('aging') || name.includes('proveedores')) {
                    if (uploadCxpHomeInput) uploadCxpHomeInput.files = files;
                    pendingCxpFile = file;
                    const nameEl = document.getElementById('cxpFileName');
                    if (nameEl) nameEl.textContent = file.name;
                } else if (name.includes('costo') || name.includes('unitario')) {
                    if (uploadCostoUnitarioHomeInput) uploadCostoUnitarioHomeInput.files = files;
                    pendingCostoUnitarioFile = file;
                    const nameEl = document.getElementById('costoUnitarioFileName');
                    if (nameEl) nameEl.textContent = file.name;
                } else {
                    if (fileInput) fileInput.files = files;
                    pendingMainFile = file;
                    const nameEl = document.getElementById('fileInputName');
                    if (nameEl) nameEl.textContent = file.name;
                }
                updateProcessButton();
            }
        });
    }

    const resetUploadBtn = document.getElementById('resetUploadBtn');
    if (resetUploadBtn) {
        resetUploadBtn.addEventListener('click', () => {
            const dropZoneContent = document.getElementById('dropZoneContent');
            const uploadFeedback = document.getElementById('uploadFeedback');
            if (dropZoneContent) dropZoneContent.style.display = 'block';
            if (uploadFeedback) uploadFeedback.style.display = 'none';
        });
    }

    if (monthSelector) {
        monthSelector.addEventListener('change', (e) => {
            const index = parseInt(e.target.value);
            if (!isNaN(index)) updateUI(globalFinancialData, index);
            if (typeof window.updateCostoUnitario === 'function') {
                window.updateCostoUnitario();
            }
        });
    }

    const pgDropdownScenario = document.getElementById('pg-dropdown-scenario');
    if (pgDropdownScenario) {
        pgDropdownScenario.addEventListener('change', (e) => {
            if (window.renderPgHorizontal) {
                window.renderPgHorizontal();
            } else if (window.resumenComercialEngine && window.resumenComercialEngine.renderPgHorizontal) {
                window.resumenComercialEngine.renderPgHorizontal();
            }
        });
    }

    const ytdToggle = document.getElementById('ytdToggle');
    if (ytdToggle) {
        ytdToggle.addEventListener('change', (e) => {
            isYTDMode = e.target.checked;
            const labelMensual = document.getElementById('label-mensual');
            const labelYtd = document.getElementById('label-ytd');
            if(labelMensual && labelYtd) {
                if(isYTDMode) {
                    labelMensual.style.color = "var(--text-secondary)";
                    labelYtd.style.color = "var(--text-primary)";
                } else {
                    labelMensual.style.color = "var(--text-primary)";
                    labelYtd.style.color = "var(--text-secondary)";
                }
            }
            const index = parseInt(monthSelector.value);
            if (!isNaN(index)) updateUI(globalFinancialData, index);
        });
    }

    // Navigation Logic
    const menuLinks = document.querySelectorAll('.menu-item a');
    menuLinks.forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const id = link.getAttribute('id');
            if (!id) return;

            // Remove active from all menus and views
            menuLinks.forEach(m => m.classList.remove('active'));
            document.querySelectorAll('.view-container').forEach(v => v.classList.remove('active'));

            // Add active to clicked link and view
            link.classList.add('active');
            const viewId = id.replace('menu-', 'view-');
            const targetView = document.getElementById(viewId);
            if (targetView) targetView.classList.add('active');
            window.currentActiveView = viewId;

            // Trigger directo al navegar a CXP
            if (id === 'menu-cxp' && typeof window.renderCxpView === 'function') {
                const ms = document.getElementById('monthSelector');
                const pIdx = ms ? parseInt(ms.value, 10) : -1;
                window.renderCxpView(window.cxpStandaloneData || null, isNaN(pIdx) ? -1 : pIdx);
            }

            // Close mobile sidebar if open
            const sidebar = document.querySelector('.sidebar');
            if (sidebar && window.innerWidth <= 1024) {
                sidebar.classList.remove('open');
            }

            // Sync title
            if (window.syncNavigationUI) {
                window.syncNavigationUI(id);
            }

            const mainHeader = document.getElementById('mainHeader');
            if (mainHeader) {
                if (id === 'menu-kpi' || id === 'menu-resumen') {
                    mainHeader.classList.add('sticky-header');
                } else {
                    mainHeader.classList.remove('sticky-header');
                }
            }

            const periodContainer = document.getElementById('periodContainer');
            if (periodContainer) {
                if (id === 'menu-glosario' || id === 'menu-config') {
                    periodContainer.style.display = 'none';
                } else {
                    periodContainer.style.display = 'flex';
                }
            }

            const searchWrapper = document.getElementById('searchContainerWrapper');
            const viewModeToggle = document.querySelector('.view-mode-toggle');
            if (monthSelector) {
                if (id === 'menu-config' || id === 'menu-glosario' || id === 'menu-pg-horizontal') {
                    monthSelector.style.display = 'none';
                } else if (globalFinancialData && globalFinancialData.length > 0) {
                    monthSelector.style.display = 'block';
                }
            }
            
            if (viewModeToggle) {
                const viewsWithYTD = ['menu-kpi', 'menu-resumen', 'menu-preliminar', 'menu-resumen-comercial'];
                if (viewsWithYTD.includes(id)) {
                    viewModeToggle.style.display = 'flex';
                } else {
                    viewModeToggle.style.display = 'none';
                }
            }
            
            if (searchWrapper) {
                const viewsWithSearch = ['menu-resumen', 'menu-preliminar', 'menu-pnl', 'menu-balance', 'menu-cashflow', 'menu-deuda', 'menu-wc', 'menu-cxp', 'menu-estados'];
                if (viewsWithSearch.includes(id) && globalFinancialData && globalFinancialData.length > 0) {
                    searchWrapper.style.display = 'flex';
                } else {
                    searchWrapper.style.display = 'none';
                }
            }

            // CRÍTICO: Disparar resize para D3.js
            window.dispatchEvent(new Event('resize'));
            
            const dropZone = document.getElementById('dropZone');
            if (dropZone) {
                if (id === 'menu-config') {
                    dropZone.style.display = 'block';
                } else {
                    const hasActiveData = (globalFinancialData && globalFinancialData.length > 0) || (typeof ceoData !== 'undefined' && ceoData && ceoData.length > 0);
                    if (hasActiveData) {
                        dropZone.style.display = 'none';
                    }
                }
            }
            
            if (globalFinancialData && globalFinancialData.length > 0 && monthSelector) {
                const idx = parseInt(monthSelector.value);
                if (!isNaN(idx)) renderActiveViewLazy(globalFinancialData, idx);
            }
            if (id === 'menu-ventas-ceo' && typeof window.renderVentasCEO === 'function') {
                window.renderVentasCEO();
            }
            if (id === 'menu-resumen-comercial' && typeof window.renderResumenComercial === 'function') {
                window.renderResumenComercial();
            }
            if (id === 'menu-costo-unitario' && typeof window.updateCostoUnitario === 'function') {
                window.updateCostoUnitario();
            }
            if (id === 'menu-pg-horizontal' && typeof window.renderPgHorizontal === 'function') {
                window.renderPgHorizontal();
            }
        });
    });

    const accountSearch = document.getElementById('accountSearch');
    if (accountSearch) {
        accountSearch.addEventListener('focus', () => {
            const monthSelector = document.getElementById('monthSelector');
            const viewModeToggle = document.querySelector('.view-mode-toggle');
            const searchWrapper = document.getElementById('searchContainerWrapper');
            
            if (monthSelector) {
                monthSelector.setAttribute('data-prev-display', monthSelector.style.display || 'block');
                monthSelector.style.display = 'none';
            }
            if (viewModeToggle) {
                viewModeToggle.setAttribute('data-prev-display', viewModeToggle.style.display || 'flex');
                viewModeToggle.style.display = 'none';
            }
            if (searchWrapper) {
                searchWrapper.style.flex = '1';
                accountSearch.style.maxWidth = '100%';
            }
        });

        accountSearch.addEventListener('blur', () => {
            const monthSelector = document.getElementById('monthSelector');
            const viewModeToggle = document.querySelector('.view-mode-toggle');
            const searchWrapper = document.getElementById('searchContainerWrapper');

            if (monthSelector && monthSelector.hasAttribute('data-prev-display')) {
                monthSelector.style.display = monthSelector.getAttribute('data-prev-display');
            }
            if (viewModeToggle && viewModeToggle.hasAttribute('data-prev-display')) {
                viewModeToggle.style.display = viewModeToggle.getAttribute('data-prev-display');
            }
            if (searchWrapper) {
                searchWrapper.style.flex = 'initial';
                accountSearch.style.maxWidth = '300px';
            }
        });

        accountSearch.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase();
            
            // Filter desktop tables
            const tablesToFilter = ['pnlDetailedTable', 'balanceTable', 'covenantTable', 'cashflowTable', 'cfMetricsTable', 'tableResumenOperativo', 'tableVentasSegmento', 'tableCostosSegmento', 'tableMargenSegmento', 'tableOpex', 'table-estados', 'cxpTable'];
            tablesToFilter.forEach(tId => {
                const table = document.getElementById(tId);
                if (table) {
                    const rows = table.querySelectorAll('tbody tr');
                    rows.forEach(tr => {
                        const firstCell = tr.querySelector('td:first-child');
                        if (firstCell) {
                            const accountName = firstCell.textContent.toLowerCase();
                            if (accountName.includes(query)) {
                                tr.style.display = '';
                            } else {
                                tr.style.display = 'none';
                            }
                        }
                    });
                }
            });

            // Filter mobile cards
            const mobileContainersToFilter = [
                 'pnlMobileContainer', 'balanceMobileContainer', 'covenantMobileContainer', 
                 'cashflowMobileContainer', 'cfMetricsMobileContainer',
                 'resumenOperativoMobileContainer', 'ventasSegmentoMobileContainer', 'costosSegmentoMobileContainer', 'margenSegmentoMobileContainer', 'opexMobileContainer', 'estadosMobileContainer', 'cxpMobileContainer'
            ];
            mobileContainersToFilter.forEach(cId => {
                const container = document.getElementById(cId);
                if (container) {
                    const cards = container.querySelectorAll('.mobile-vertical-card');
                    cards.forEach(card => {
                        const titleEl = card.querySelector('.mobile-vertical-card-title span');
                        if (titleEl) {
                            const accountName = titleEl.textContent.toLowerCase();
                            if (accountName.includes(query)) {
                                card.style.display = '';
                            } else {
                                card.style.display = 'none';
                            }
                        }
                    });

                    // Hide empty accordion groups
                    const accordions = container.querySelectorAll('.mobile-accordion-group');
                    accordions.forEach(acc => {
                        const visibleCards = acc.querySelectorAll('.mobile-vertical-card[style=""]');
                        // if searching and no visible cards, hide the whole accordion
                        if (query !== '' && visibleCards.length === 0 && acc.querySelectorAll('.mobile-vertical-card').length > 0) {
                            acc.style.display = 'none';
                        } else {
                            acc.style.display = '';
                            if (query !== '') {
                                // Auto expand if searching
                                const content = acc.querySelector('.mobile-accordion-content');
                                if (content) content.classList.add('open');
                            }
                        }
                    });
                }
            });
        });
    }

    if (typeof lucide !== 'undefined') lucide.createIcons();
    
    // Global polished tooltip system for KPI Cards (matches chart style)
    let globalTooltip = d3.select("body").select(".d3-tooltip");
    if (globalTooltip.empty()) {
        globalTooltip = d3.select("body")
            .append("div")
            .attr("class", "d3-tooltip")
            .style("opacity", 0);
    }

    // Add event delegation for any element with data-tooltip
    document.addEventListener('mouseover', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger) {
            const text = trigger.getAttribute('data-tooltip');
            globalTooltip.style("opacity", 1)
                .html(text);
        }
    });

    document.addEventListener('mousemove', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger) {
            globalTooltip
                .style("left", (e.pageX + 15) + "px")
                .style("top", (e.pageY - 15) + "px");
        }
    });

    document.addEventListener('mouseout', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger) {
            globalTooltip.style("opacity", 0);
        }
    });

    // Support for touch devices (click to show/hide)
    document.addEventListener('click', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger && window.innerWidth < 1024) {
            const isVisible = globalTooltip.style("opacity") === "1";
            if (isVisible) {
                globalTooltip.style("opacity", 0);
            } else {
                const text = trigger.getAttribute('data-tooltip');
                globalTooltip.style("opacity", 1)
                    .html(text)
                    .style("left", (e.pageX + 15) + "px")
                    .style("top", (e.pageY - 15) + "px");
            }
        } else if (!trigger) {
            globalTooltip.style("opacity", 0);
        }
    });

    // Go to top button logic
    const mainContainer = document.querySelector('.main-container');
    const scrollTopBtn = document.getElementById('scrollTopBtn');
    if (mainContainer && scrollTopBtn) {
        mainContainer.addEventListener('scroll', () => {
            if (mainContainer.scrollTop > 300) {
                scrollTopBtn.classList.add('visible');
            } else {
                scrollTopBtn.classList.remove('visible');
            }
        });
        
        // Mobile fallback for body scroll
        window.addEventListener('scroll', () => {
             if (window.scrollY > 300) {
                 scrollTopBtn.classList.add('visible');
             } else {
                 scrollTopBtn.classList.remove('visible');
             }
        });

        scrollTopBtn.addEventListener('click', () => {
            mainContainer.scrollTo({ top: 0, behavior: 'smooth' });
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
    }

    // Handle window resize for D3 Charts redrawing and Mobile Accordions
    let resizeTimer;
    let lastWindowWidth = window.innerWidth;

    window.addEventListener('resize', () => {
        if (window.innerWidth === lastWindowWidth) {
            return;
        }
        lastWindowWidth = window.innerWidth;

        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => {
            if (globalFinancialData && globalFinancialData.length > 0 && monthSelector) {
                const idx = parseInt(monthSelector.value);
                if (!isNaN(idx)) {
                    renderActiveViewLazy(globalFinancialData, idx);
                }
            }
        }, 200);
    });
});

function downloadCSV(data, filename) {
    if (!data || !data.length) return;

    const columns = [];
    data.forEach(row => {
        Object.keys(row).forEach(key => {
            if (!columns.includes(key)) columns.push(key);
        });
    });

    let csvContent = '\uFEFF'; 
    csvContent += columns.map(c => `"${String(c).replace(/"/g, '""')}"`).join(';') + '\n';

    data.forEach(row => {
        const rowStr = columns.map(c => {
            const val = row[c] !== undefined && row[c] !== null ? String(row[c]) : '';
            return `"${val.replace(/"/g, '""')}"`;
        }).join(';');
        csvContent += rowStr + '\n';
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

// File Processing Logic Separated from Rendering
async function processFile(bufferPromise, progressCallback) {
    return new Promise(async (resolve, reject) => {
        // Simular progreso para dar feedback visual
        let simulatedProgress = 0;
        const progressInterval = setInterval(() => {
            if (simulatedProgress < 90) {
                simulatedProgress += 5;
                if (progressCallback) progressCallback(simulatedProgress, "Procesando archivo...");
            }
        }, 100);

        try {
            if (progressCallback) progressCallback(30, "Enviando al procesador en segundo plano...");
            
            const bufferData = await bufferPromise;

            const engineResult = await new Promise((workerResolve, workerReject) => {
                const worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
                worker.onmessage = (e) => {
                    const data = e.data;
                    if (data.type === 'progress') {
                        if (progressCallback) progressCallback(data.progress, data.message);
                    } else if (data.type === 'error') {
                        workerReject(new Error(data.error));
                        worker.terminate();
                    } else if (data.type === 'done') {
                        workerResolve(data.engineResult);
                        worker.terminate();
                    }
                };
                worker.onerror = (err) => {
                    console.error("Worker error details:", err);
                    workerReject(new Error("Worker error: " + (err.message || (err.error && err.error.message) || JSON.stringify(err) || "Unknown error")));
                    worker.terminate();
                };
                worker.postMessage({ buffer: bufferData }, [bufferData]);
            });
            
            if (progressCallback) progressCallback(80, "Validando estructura de datos...");
            
            const lastData = engineResult.data[engineResult.data.length - 1];
            if (!lastData || !lastData.balance) {
                clearInterval(progressInterval);
                return reject(new Error("Estructura de datos incompleta en el archivo."));
            }

            clearInterval(progressInterval);
            if (progressCallback) progressCallback(100, "Carga Completada");
            resolve(engineResult);

        } catch (err) {
            clearInterval(progressInterval);
            reject(new Error("Error procesando o leyendo el archivo: " + err.message));
        }
    });
}

async function handleFileUpload(e) {
    const file = e.target && e.target.files ? e.target.files[0] : (e.dataTransfer ? e.dataTransfer.files[0] : null);
    if (!file) return;
    
    // Start reading file immediately before any UI updates or event loop yields (fixes Drag & Drop permission loss)
    const bufferPromise = file.arrayBuffer().catch(err => {
        throw new Error("File read error: " + err.message);
    });

    // UI Elements
    const dropZoneContent = document.getElementById('dropZoneContent');
    const uploadFeedback = document.getElementById('uploadFeedback');
    const uploadProgressBar = document.getElementById('uploadProgressBar');
    const uploadMessage = document.getElementById('uploadMessage');
    const uploadTitle = document.getElementById('uploadTitle');
    const uploadIcon = document.getElementById('uploadIcon');
    const resetUploadBtn = document.getElementById('resetUploadBtn');

    // Reset and show feedback UI
    if (dropZoneContent) dropZoneContent.style.display = 'none';
    if (uploadFeedback) uploadFeedback.style.display = 'flex';
    if (resetUploadBtn) resetUploadBtn.style.display = 'none';
    if (uploadProgressBar) uploadProgressBar.style.width = '0%';
    if (uploadTitle) {
        uploadTitle.textContent = "Procesando archivo...";
        uploadTitle.style.color = "var(--text-primary)";
    }
    if (uploadIcon) {
        uploadIcon.setAttribute('data-lucide', 'loader');
        uploadIcon.classList.add('spin-icon');
        uploadIcon.style.color = "var(--primary)";
        if (window.lucide) window.lucide.createIcons();
    }
    
    try {
        const engineResult = await processFile(bufferPromise, (progress, message) => {
            if (uploadProgressBar) uploadProgressBar.style.width = `${progress}%`;
            if (uploadMessage) uploadMessage.textContent = message;
        });

        // AIAgent Analysis Logic
        if (uploadMessage) uploadMessage.textContent = "Validando datos...";
        
        const lastData = engineResult.data[engineResult.data.length - 1];
        const pnlResult = {
            ventas: engineResult.data.map(d => d.kpis.ingresos),
            ebitda: engineResult.data.map(d => d.kpis.ebitda)
        };
        const balanceResult = {
            activos: lastData.balance.activos || 0,
            pasivos: lastData.balance.pasivos || 0,
            patrimonio: lastData.balance.patrimonio || 0
        };

        const llmInput = buildLLMInput({
            pnlData: pnlResult,
            balanceData: balanceResult,
            source: "excel_upload"
        });

        const validation = validateLLMInput(llmInput);

        if (!validation.isValid) {
            console.warn("Validation Warnings:", validation.errors);
            if (uploadMessage) uploadMessage.textContent = `✅ Modelo Local: ${engineResult.modelType}`;
        } else {
            if (uploadMessage) uploadMessage.textContent = "🚀 Consultando Analista...";
            try {
                const aiResponse = await callAI(llmInput);
                if (uploadMessage) uploadMessage.textContent = `✅ Análisis Completado`;
                
                const lastIdx = engineResult.data.length - 1;
                if (aiResponse.alerts) {
                    engineResult.data[lastIdx].alerts = [...(engineResult.data[lastIdx].alerts || []), ...aiResponse.alerts];
                }
            } catch (aiErr) {
                window.handleAiError("Upload AI Check", aiErr);
                if (uploadMessage) uploadMessage.textContent = `⚠️ Usando motor local.`;
            }
        }

        // Set success state
        if (uploadProgressBar) uploadProgressBar.style.width = `100%`;
        if (uploadProgressBar) uploadProgressBar.style.background = `var(--success)`;
        if (uploadTitle) {
            uploadTitle.textContent = "¡Carga Exitosa!";
            uploadTitle.style.color = "var(--success)";
        }
        if (uploadIcon) {
            uploadIcon.setAttribute('data-lucide', 'check-circle');
            uploadIcon.classList.remove('spin-icon');
            uploadIcon.style.color = "var(--success)";
            if (window.lucide) window.lucide.createIcons();
        }
        if (resetUploadBtn) resetUploadBtn.style.display = 'inline-block';
        
        // Show success, then render
        setTimeout(async () => {
            // Clear caches to prevent memory leaks and stale data
            window.aiSummaryCache = {};
            window.aiAlertsCache = {};
            window.simSummaryCache = {};
            
            globalFinancialData = engineResult.data;
            if (window.cachedStandaloneCxp && typeof window.applyCachedStandaloneCxp === 'function') {
                await window.applyCachedStandaloneCxp();
            }
            
            // --- GUARDAR JSON PROCESADO EN INDEXEDDB ---
            try {
                const CACHE_KEY = 'MASTER_FINANCE_KEY';
                const db = await getFinanceDB();
                await new Promise((resolve, reject) => {
                    const tx = db.transaction('finance_cache', 'readwrite');
                    tx.objectStore('finance_cache').put({ data: engineResult.data, timestamp: Date.now() }, CACHE_KEY);
                    tx.oncomplete = resolve;
                    tx.onerror = reject;
                });
                console.log("✨ La Gran Victoria: JSON procesado guardado en IndexedDB con éxito.");
            } catch (e) {
                console.warn("⚠️ Error guardando caché manual en IndexedDB:", e);
            }

            renderDashboard(globalFinancialData);
        }, 500);

    } catch (err) {
        console.error("Upload error:", err);
        if (uploadProgressBar) uploadProgressBar.style.background = `var(--danger)`;
        if (uploadTitle) {
            uploadTitle.textContent = "Error al Cargar";
            uploadTitle.style.color = "var(--danger)";
        }
        if (uploadIcon) {
            uploadIcon.setAttribute('data-lucide', 'x-circle');
            uploadIcon.classList.remove('spin-icon');
            uploadIcon.style.color = "var(--danger)";
            if (window.lucide) window.lucide.createIcons();
        }
        if (uploadMessage) uploadMessage.textContent = err.message;
        if (resetUploadBtn) resetUploadBtn.style.display = 'inline-block';
        showError(err.message);
    }
}

async function callAI(payload) {
    let timeoutId;
    const timeoutPromise = new Promise((_, reject) => {
        timeoutId = setTimeout(() => reject(new Error('AI Request Timeout (45s)')), 45000);
    });

    let apiCallPromise;
    try {
        apiCallPromise = getAI().models.generateContent({
            model: "gemini-2.5-flash",
            contents: `Actúa como un Senior Financial Analyst y analiza estos datos de P&L y Balance.
        
        INSTRUCCIONES:
        1. Devuelve un JSON estrictamente válido.
        2. Proporciona insights sobre el EBITDA y la eficiencia operativa.
        3. Identifica variaciones atípicas.
        
        ESTRUCTURA REQUERIDA (NO OMITIR CAMPOS):
        {
          "date": "Periodo Actual",
          "kpis": { "ingresos": number, "ebitda": number, "margen_ebitda": number, "cashflow": number },
          "balance": { "activos": number, "pasivos": number, "patrimonio": number, "cuadra": boolean },
          "pnl": { "categorias": { "Categoria": valor, ... }, "segments": {} },
          "alerts": ["string"]
        }

        DATOS PARA ANALIZAR:
        ${JSON.stringify(payload, null, 2)}`,
            config: {
                responseMimeType: "application/json"
            }
        });
        apiCallPromise.catch(err => window.handleAiError("AI Engine", err));
    } catch (err) {
        apiCallPromise = Promise.reject(err);
        apiCallPromise.catch(()=> /* handled */ {});
    }

    let response;
    try {
        response = await Promise.race([apiCallPromise, timeoutPromise]);
    } finally {
        clearTimeout(timeoutId);
    }

    let text = response.text;
    text = text.replace(/```json/g, '').replace(/```/g, '').trim();
    
    try {
        return JSON.parse(text);
    } catch (e) {
        // Fallback: Try to extract just the first JSON object array or object
        const jsonMatch = text.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
        if (jsonMatch) {
            return JSON.parse(jsonMatch[0]);
        }
        throw e;
    }
}

function showError(msg) {
    const statusEl = document.getElementById('engineStatus');
    statusEl.style.background = '#fee2e2';
    statusEl.style.color = '#991b1b';
    statusEl.style.borderColor = '#fecaca';
    statusEl.innerHTML = `❌ ${msg}`;
}

/**
 * 🚀 MOBILE ACCORDION GENERATOR
 * Converts desktop tables into mobile-friendly vertical cards wrapped in accordions.
 */
function buildMobileAccordionsFromTable(tableId, containerId, customTitle = null, customSummary = null) {
    const table = document.getElementById(tableId);
    if (!table) return;
    const isMobile = window.innerWidth < 768;
    const container = document.getElementById(containerId);
    if(!container) return;

    if (!isMobile) {
        table.style.display = '';
        container.style.display = 'none';
        return;
    }

    // Determine if table is inside a section or just bare
    table.style.setProperty('display', 'none', 'important');
    container.style.display = 'block';

    const ths = Array.from(table.querySelectorAll('thead th'));
    const headers = ths.slice(1).map(th => th.innerText);

    const rows = Array.from(table.querySelectorAll('tbody tr'));
    
    let html = '';
    let currentGroupHtml = '';
    let currentGroupTitle = customTitle || 'Categoría / Cuentas';
    let currentGroupSummary = customSummary || '';
    
    let isSingleGroupTable = !rows.some(tr => tr.classList.contains('row-category'));

    const flushGroup = (newTitle, newSummary) => {
        if (currentGroupHtml !== '') {
             html += `<div class="mobile-accordion-group">
                <div class="mobile-accordion-header" onclick="this.nextElementSibling.classList.toggle('open')">
                    <div style="display:flex; flex-direction:column; gap:4px; max-width:85%;">
                        <span style="text-transform: uppercase;">${currentGroupTitle}</span>
                        ${currentGroupSummary ? `<span style="font-size:12px; color:var(--primary); font-weight: 800;">TOTAL: ${currentGroupSummary}</span>` : ''}
                    </div>
                    <i data-lucide="chevron-down" style="width:20px;height:20px;"></i>
                </div>
                <!-- Remove display none by default if it's a single group table so it opens by default or let user open it -->
                <div class="mobile-accordion-content ${isSingleGroupTable ? 'open' : ''}">
                    ${currentGroupHtml}
                </div>
             </div>`;
        }
        currentGroupHtml = '';
        currentGroupTitle = newTitle || customTitle || 'Categoría';
        currentGroupSummary = newSummary || '';
    };

    rows.forEach((tr, i) => {
        const tds = Array.from(tr.querySelectorAll('td'));
        if (tds.length < 2) return; // empty row or spacer

        const label = tds[0].innerText;
        const vals = tds.slice(1).map(td => td.innerText);
        
        const isTotal = tr.classList.contains('row-total');
        const isCategory = tr.classList.contains('row-category');
        
        if (isCategory) {
            flushGroup(label);
        } else if (isTotal) {
            // Find a valid numerical string to show as summary for the accordion
            let summaryVal = '';
            for(let j = vals.length - 1; j >= 0; j--) {
                if(vals[j] && vals[j] !== '-') { summaryVal = vals[j]; break; }
            }
            if(!currentGroupSummary) currentGroupSummary = summaryVal || vals[vals.length - 1];
            currentGroupHtml += createMobileCard(label, headers, vals);
            if (currentGroupTitle === 'Categoría') currentGroupTitle = label;
            
            // Only flush if we're dealing with a multi-category table like P&L
            if (!isSingleGroupTable && i < rows.length - 1) {
                flushGroup();
            }
        } else {
            currentGroupHtml += createMobileCard(label, headers, vals);
            // If it's the last row and a single group table, and we don't have a summary, we can try to guess it.
        }
    });

    flushGroup(); // flush remaining
    
    if (html === '') {
       container.innerHTML = '<div style="padding:20px; text-align:center; font-size:12px; color:var(--text-secondary);">No hay datos formatados para mostrar.</div>';
    } else {
       // Add Swipe Indicator (as requested by user)
       container.innerHTML = `<div class="swipe-indicator"> <i data-lucide="chevrons-down" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> Toca para interactuar</div>` + html;
       if (typeof lucide !== 'undefined') lucide.createIcons();
    }
}

function createMobileCard(label, headers, vals) {
    let cardHtml = `<div class="mobile-vertical-card">
        <div class="mobile-vertical-card-title">
            <span style="max-width:80%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${label}</span>
            <i data-lucide="trending-up" style="width:14px;height:14px;opacity:0.5;"></i>
        </div>`;
    vals.forEach((v, idx) => {
        if (headers[idx]) {
            cardHtml += `<div class="mobile-vertical-card-row">
                <span style="color:var(--text-secondary);">${headers[idx]}</span>
                <span style="font-weight:700;">${v}</span>
            </div>`;
        }
    });
    cardHtml += `</div>`;
    return cardHtml;
}

// Global UI Updater Function
function renderDashboard(data) {
    if (!data || data.length === 0) return;
    
    window.handleZeroState();
    
    // Filtro: No permitir seleccionar datos del 2025 en el dropdown global
    const filteredForSelector = data.map((d, i) => ({ d, i })).filter(item => isYear2026(item.d));
    
    monthSelector.innerHTML = filteredForSelector.map(item => `<option value="${item.i}">${item.d.date || 'Periodo'}</option>`).join('');
    monthSelector.style.display = 'block';
    
    // Show search input if one of the detailed views is active
    const searchWrapper = document.getElementById('searchContainerWrapper');
    if (searchWrapper) {
        const activeMenu = document.querySelector('.menu-item a.active');
        const viewsWithSearch = ['menu-resumen', 'menu-preliminar', 'menu-pnl', 'menu-balance', 'menu-cashflow', 'menu-deuda', 'menu-wc', 'menu-cxp', 'menu-estados'];
        if (activeMenu && viewsWithSearch.includes(activeMenu.id)) {
            searchWrapper.style.display = 'flex';
        }
    }
    
    const lastIdx = filteredForSelector.length > 0 ? filteredForSelector[filteredForSelector.length - 1].i : data.length - 1;
    monthSelector.value = lastIdx;
    
    // Yield rendering to prevent main thread blocking on mobile
    setTimeout(() => {
        updateUI(data, lastIdx);
    }, 10);
}

function getAggregatedData(dataArray, currentIndex, isYTD) {
    if (currentIndex < 0 || !dataArray[currentIndex]) return { currAgg: null, prevAgg: null };
    
    const curr = dataArray[currentIndex];
    
    if (!isYTD) {
        const operationalData = dataArray.filter(d => isYear2026(d));
        const currIdxInOp = operationalData.findIndex(d => d.date === curr.date);
        const operationalPrev = currIdxInOp > 0 ? operationalData[currIdxInOp - 1] : curr;
        return { 
            currAgg: curr, 
            prevAgg: operationalPrev 
        };
    }

    const targetYear = curr.sortDate ? getSortYear(curr) : 2026;
    const targetMonth = curr.sortDate ? curr.sortDate.getMonth() : 11;

    const createEmptyAgg = () => ({
        kpis: { ingresos: 0, ebitda: 0, cashflow: 0, utilidad: 0, margen_ebitda: 0 },
        pnl: { categorias: {}, segments: {}, opexDetalle: {} },
        ppto: {
            kpis: { ingresos: 0, ebitda: 0, cashflow: 0, utilidad: 0 },
            pnl: { categorias: {}, segments: {}, opexDetalle: {} }
        }
    });

    const sumAgg = (agg, source) => {
        if (!source) return;
        if (source.kpis) {
            agg.kpis.ingresos += source.kpis.ingresos || 0;
            agg.kpis.ebitda += source.kpis.ebitda || 0;
            agg.kpis.cashflow += source.kpis.cashflow || 0;
            agg.kpis.utilidad += source.kpis.utilidad || 0;
        }
        if (source.pnl && source.pnl.categorias) {
            for (const [key, val] of Object.entries(source.pnl.categorias)) {
                agg.pnl.categorias[key] = (agg.pnl.categorias[key] || 0) + val;
            }
        }
        if (source.pnl && source.pnl.segments) {
            for (const [seg, segData] of Object.entries(source.pnl.segments)) {
                if (!agg.pnl.segments[seg]) agg.pnl.segments[seg] = { ventas: 0, costos: 0 };
                agg.pnl.segments[seg].ventas += segData.ventas || 0;
                agg.pnl.segments[seg].costos += segData.costos || 0;
            }
        }
        if (source.pnl && source.pnl.opexDetalle) {
            for (const [key, val] of Object.entries(source.pnl.opexDetalle)) {
                agg.pnl.opexDetalle[key] = (agg.pnl.opexDetalle[key] || 0) + val;
            }
        }
        if (source.ppto) {
            if (source.ppto.kpis) {
                agg.ppto.kpis.ingresos += source.ppto.kpis.ingresos || 0;
                agg.ppto.kpis.ebitda += source.ppto.kpis.ebitda || 0;
                agg.ppto.kpis.cashflow += source.ppto.kpis.cashflow || 0;
                agg.ppto.kpis.utilidad += source.ppto.kpis.utilidad || 0;
            }
            if (source.ppto.pnl && source.ppto.pnl.categorias) {
                 for (const [key, val] of Object.entries(source.ppto.pnl.categorias)) {
                    agg.ppto.pnl.categorias[key] = (agg.ppto.pnl.categorias[key] || 0) + val;
                }
            }
            if (source.ppto.pnl && source.ppto.pnl.segments) {
                for (const [seg, segData] of Object.entries(source.ppto.pnl.segments)) {
                    if (!agg.ppto.pnl.segments[seg]) agg.ppto.pnl.segments[seg] = { ventas: 0, costos: 0 };
                    agg.ppto.pnl.segments[seg].ventas += segData.ventas || 0;
                    agg.ppto.pnl.segments[seg].costos += segData.costos || 0;
                }
            }
            if (source.ppto.pnl && source.ppto.pnl.opexDetalle) {
                for (const [key, val] of Object.entries(source.ppto.pnl.opexDetalle)) {
                    agg.ppto.pnl.opexDetalle[key] = (agg.ppto.pnl.opexDetalle[key] || 0) + val;
                }
            }
        }
    };

    const currAgg = createEmptyAgg();
    const prevAgg = createEmptyAgg();

    dataArray.forEach(d => {
        if (!d.sortDate) return;
        const dYear = getSortYear(d);
        const dMonth = d.sortDate.getMonth();

        if (dYear === targetYear && dMonth <= targetMonth) {
            sumAgg(currAgg, d);
        }
        
        if (dYear === targetYear - 1 && dMonth <= targetMonth) {
            sumAgg(prevAgg, d);
        }
    });

    currAgg.kpis.margen_ebitda = currAgg.kpis.ingresos !== 0 ? currAgg.kpis.ebitda / currAgg.kpis.ingresos : 0;
    prevAgg.kpis.margen_ebitda = prevAgg.kpis.ingresos !== 0 ? prevAgg.kpis.ebitda / prevAgg.kpis.ingresos : 0;

    return { currAgg, prevAgg };
}

function updateUI(data, index) {
    if (!data || !data[index]) return;
    const curr = data[index];
    
    // Usamos getAggregatedData para obtener currAgg y prevAgg basándonos en isYTDMode
    const { currAgg, prevAgg } = getAggregatedData(data, index, isYTDMode);
    
    // Variables for rendering explicitly (using the aggregated data for the specified cards)
    const aggKpis = currAgg.kpis || { ingresos: 0, ebitda: 0, cashflow: 0, margen_ebitda: 0 };
    const prevAggKpis = prevAgg.kpis || aggKpis;
    const aggPptoKpis = (currAgg.ppto && currAgg.ppto.kpis) ? currAgg.ppto.kpis : { ingresos: 0, ebitda: 0 };
    
    // Integrity Badge logic
    const integrityBadge = document.getElementById('integrityBadge');
    if (integrityBadge && curr.integrity) {
        integrityBadge.style.display = 'flex';
        if (curr.integrity.isBroken) {
            integrityBadge.className = 'integrity-fail';
            integrityBadge.innerHTML = `⚠️ Ajuste Detectado (Abs: ${formatCurrency(curr.integrity.gap)})`;
            integrityBadge.title = "La suma de Ingresos - Costos - Gastos no coincide con el EBITDA reportado";
        } else {
            integrityBadge.className = 'integrity-ok';
            integrityBadge.innerHTML = `✓ P&L Cuadrado`;
            integrityBadge.title = "Integridad de datos verificada operativamente";
        }
    }

    // Inyectar en Tarjetas de KPIs: Ingresos Netos, Costo de Ventas, EBITDA
    document.getElementById('kpi-ventas').textContent = formatCurrency(aggKpis.ingresos);
    document.getElementById('kpi-ebitda').textContent = formatCurrency(aggKpis.ebitda);

    const aggCategories = (currAgg.pnl && currAgg.pnl.categorias) ? currAgg.pnl.categorias : {};
    const aggTotalCost = aggCategories["Costo de Ventas"] || 0;
    const prevAggCategories = (prevAgg.pnl && prevAgg.pnl.categorias) ? prevAgg.pnl.categorias : aggCategories;
    const prevAggTotalCost = prevAggCategories["Costo de Ventas"] || 0;
    const aggPptoCategories = (currAgg.ppto && currAgg.ppto.pnl && currAgg.ppto.pnl.categorias) ? currAgg.ppto.pnl.categorias : {};

    document.getElementById('val-ratio').textContent = formatCurrency(aggTotalCost);

    const statusEl = document.getElementById('engineStatus');
    if (statusEl && curr.pnl && curr.pnl.detectedRows) {
        statusEl.innerHTML = `✅ Datos Detectados:<br>
            <b>Ingresos:</b> "${curr.pnl.detectedRows.ingresos || '?'}"<br>
            <b>EBITDA:</b> "${curr.pnl.detectedRows.ebitda || '?'}"<br>
            <b>OPEX:</b> "${curr.pnl.detectedRows.opex || '?'}"<br>
            <b>Balance:</b> "${curr.pnl.detectedRows.activos || 'No detectado'}"`;
    }
    
    document.getElementById('periodLabel').textContent = `Periodo de Análisis: ${curr.date || 'Actual'}`;
    updateTrend('sub-ventas', aggKpis.ingresos, prevAggKpis.ingresos, aggPptoKpis.ingresos || 0);
    const margin = ((aggKpis.margen_ebitda || 0) * 100).toFixed(1);
    updateTrend('sub-ebitda', aggKpis.ebitda, prevAggKpis.ebitda, aggPptoKpis.ebitda || 0, ` | Margen: ${margin}%`);
    updateTrend('sub-ratio', aggTotalCost, prevAggTotalCost, aggPptoCategories["Costo de Ventas"] || 0);

    // Renderizar resto condicionalmente
    renderActiveViewLazy(data, index);
}

function renderActiveViewLazy(data, index) {
    if (!data || !data[index]) return;
    const curr = data[index];
    const prevIdx = Math.max(0, index - 1);
    const prev = data[prevIdx];
    
    const operationalData = data.filter(d => isYear2026(d));
    const currIdxInOp = operationalData.findIndex(d => d.date === curr.date);
    const operationalPrev = currIdxInOp > 0 ? operationalData[currIdxInOp - 1] : curr;
    
    // We defer heavy operations via requestAnimationFrame and target only the view being displayed.
    requestAnimationFrame(() => {
        let viewPreliminar = document.getElementById("view-preliminar");
        if (viewPreliminar && viewPreliminar.classList.contains("active")) {
            renderPreliminaryView(data, index);
        }

        let viewPnl = document.getElementById("view-pnl");
        if (viewPnl && viewPnl.classList.contains("active")) {
            renderDetailedPnL(data, index);
            
            if (viewPnl) {
                let pnlDetailTable = viewPnl.querySelector(".pnl-detail-table");
                if (pnlDetailTable) {
                    if (!document.getElementById("marginTrendChart")) {
                        let marginContainer = document.createElement("div");
                        marginContainer.id = "marginTrendChart";
                        marginContainer.style.width = "100%";
                        marginContainer.style.height = "300px";
                        marginContainer.style.marginBottom = "30px";
                        pnlDetailTable.parentNode.insertBefore(marginContainer, pnlDetailTable);
                    }
                    if (!document.getElementById("waterfallChart")) {
                        let waterfallContainer = document.createElement("div");
                        waterfallContainer.id = "waterfallChart";
                        waterfallContainer.style.width = "100%";
                        waterfallContainer.style.height = "350px";
                        waterfallContainer.style.marginBottom = "30px";
                        pnlDetailTable.parentNode.insertBefore(waterfallContainer, pnlDetailTable);
                    }
                }
            }
            renderWaterfallChart(data, index);
            renderMarginTrendChart(data, index);
            
            // Hemos deshabilitado buildMobileAccordionsFromTable para pnlDetailedTable
            // para permitir el scroll horizontal manejado por CSS, pero el usuario pidió accordeon
            setTimeout(() => {
                buildMobileAccordionsFromTable('pnlDetailedTable', 'pnlMobileContainer');
            }, 10);
        }

        let viewBalance = document.getElementById("view-balance");
        if (viewBalance && viewBalance.classList.contains("active")) {
            renderBalanceSheet(data, index);
            setTimeout(() => {
                buildMobileAccordionsFromTable('balanceTable', 'balanceMobileContainer');
                buildMobileAccordionsFromTable('covenantTable', 'covenantMobileContainer');
            }, 10);
        }

        let viewCashflow = document.getElementById("view-cashflow");
        if (viewCashflow && viewCashflow.classList.contains("active")) {
            renderCashFlow(data, index);
            
            let cfDetailTable = viewCashflow.querySelector(".pnl-detail-table");
            if (cfDetailTable) {
                if (!document.getElementById("cashBridgeChart")) {
                    let cashBridgeContainer = document.createElement("div");
                    cashBridgeContainer.id = "cashBridgeChart";
                    cashBridgeContainer.style.width = "100%";
                    cashBridgeContainer.style.height = "350px";
                    cashBridgeContainer.style.marginBottom = "30px";
                    cfDetailTable.parentNode.insertBefore(cashBridgeContainer, cfDetailTable);
                }
            }
            renderCashBridgeChart(data, index);
            
            setTimeout(() => {
                buildMobileAccordionsFromTable('cashflowTable', 'cashflowMobileContainer');
                buildMobileAccordionsFromTable('cfMetricsTable', 'cfMetricsMobileContainer');
            }, 10);
        }
        
        let viewWc = document.getElementById("view-wc");
        if (viewWc && viewWc.classList.contains("active")) {
            renderWorkingCapital(data, index);
            setTimeout(() => {
                if (typeof buildMobileAccordionsFromTable === 'function') {
                    buildMobileAccordionsFromTable('wcTable', 'wcMobileContainer');
                }
            }, 10);
        }
        
        let viewCxp = document.getElementById("view-cxp");
        // Update cxp view always so it scales automatically for when the user navigates there
        if (viewCxp) {
            if (typeof window.renderCxpView === 'function') {
                window.renderCxpView(window.cxpStandaloneData || null, index);
            }
            setTimeout(() => {
                if (typeof buildMobileAccordionsFromTable === 'function') {
                    buildMobileAccordionsFromTable('cxpTable', 'cxpMobileContainer');
                }
            }, 10);
        }
        
        let viewKpi = document.getElementById("view-kpi");
        if (viewKpi && viewKpi.classList.contains("active")) {
            renderKPIDashboard(data, index);
        }

        let viewResumenComercial = document.getElementById("view-resumen-comercial");
        if (viewResumenComercial && viewResumenComercial.classList.contains("active")) {
            if (typeof window.renderResumenComercial === 'function') {
                window.renderResumenComercial();
            }
        }

        let viewPgHorizontal = document.getElementById("view-pg-horizontal");
        if (viewPgHorizontal && viewPgHorizontal.classList.contains("active")) {
            if (typeof window.renderPgHorizontal === 'function') {
                window.renderPgHorizontal();
            }
        }

        let viewDeuda = document.getElementById("view-deuda");
        if (viewDeuda && viewDeuda.classList.contains("active")) {
            renderDeudaView(data, index);
        }

        let viewEstados = document.getElementById("view-estados");
        if (viewEstados && viewEstados.classList.contains("active")) {
            renderEstadosFinancieros(data, index);
            setTimeout(() => {
                if (typeof buildMobileAccordionsFromTable === 'function') {
                    buildMobileAccordionsFromTable('table-estados', 'estadosMobileContainer');
                }
            }, 10);
        }

        let viewSimulador = document.getElementById("view-simulador");
        if (viewSimulador && viewSimulador.classList.contains("active")) {
            if (typeof window.runSimulationMath === 'function') {
                window.runSimulationMath();
            }
        }

        let viewResumen = document.getElementById("view-resumen");
        if (viewResumen && viewResumen.classList.contains("active")) {
            const { currAgg, prevAgg } = getAggregatedData(data, index, isYTDMode);
            
            // Dynamic headers logic
            let currentLabel = curr.date || 'Periodo Actual';
            let prevLabel = operationalPrev.date || 'Periodo Previo';

            if (isYTDMode && curr.sortDate) {
                const monthNames = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
                const m = curr.sortDate.getMonth();
                const y = getSortYear(curr);
                currentLabel = `Ene-${monthNames[m]} ${y}`;
                prevLabel = `Ene-${monthNames[m]} ${y - 1}`;
            }

            const headerIds = [
                'tableResumenOperativo',
                'tableVentasSegmento',
                'tableCostosSegmento',
                'tableMargenSegmento',
                'tableOpex'
            ];
            
            headerIds.forEach(id => {
                const table = document.getElementById(id);
                if (table) {
                    const ths = table.querySelectorAll('thead th');
                    if (ths.length > 2) {
                        ths[1].textContent = prevLabel; // 2nd col
                        ths[2].textContent = currentLabel; // 3rd col
                    }
                }
            });

            // Re-render resumen widgets fully
            const aggKpis = currAgg.kpis || { ingresos: 0, ebitda: 0, cashflow: 0, margen_ebitda: 0 };
            const aggCategories = (currAgg.pnl && currAgg.pnl.categorias) ? currAgg.pnl.categorias : {};
            const prevAggCategories = (prevAgg.pnl && prevAgg.pnl.categorias) ? prevAgg.pnl.categorias : aggCategories;
            const aggTotalCost = aggCategories["Costo de Ventas"] || 0;
            const aggPptoCategories = (currAgg.ppto && currAgg.ppto.pnl && currAgg.ppto.pnl.categorias) ? currAgg.ppto.pnl.categorias : {};

            const segments = (currAgg.pnl && currAgg.pnl.segments) ? currAgg.pnl.segments : {};
            const prevSegments = (prevAgg.pnl && prevAgg.pnl.segments) ? prevAgg.pnl.segments : {};
            const pptoSegments = (currAgg.ppto && currAgg.ppto.pnl && currAgg.ppto.pnl.segments) ? currAgg.ppto.pnl.segments : {};
            
            // Render Segmentos Ventas
            const segmentsSection = document.getElementById('segments-section');
            const segmentsBody = document.getElementById('segmentsBody');
            const segmentsList = Object.entries(segments).filter(([name]) => {
                const norm = name.toLowerCase();
                return !norm.includes('otras ventas') && !norm.includes('otros ingresos');
            });
            if (segmentsList.length > 0) {
                segmentsSection.style.display = 'block';
                segmentsBody.innerHTML = segmentsList.map(([name, dataSeg]) => {
                    const ventas = dataSeg.ventas || 0;
                    const prevVentas = prevSegments[name] ? prevSegments[name].ventas : 0;
                    const pptoVentas = pptoSegments[name] ? pptoSegments[name].ventas : 0;
                    const diff = ventas - prevVentas;
                    const diffPpto = ventas - pptoVentas;
                    const pctPart = aggKpis.ingresos !== 0 ? (ventas / aggKpis.ingresos) * 100 : 0;
                    const pctMoM = prevVentas !== 0 ? (diff / Math.abs(prevVentas)) * 100 : 0;
                    const pctPpto = pptoVentas !== 0 ? (diffPpto / Math.abs(pptoVentas)) * 100 : 0;
                    
                    const color = diff >= 0 ? 'var(--success)' : 'var(--danger)'; 
                    const colorPpto = diffPpto >= 0 ? 'var(--success)' : 'var(--danger)';
                    const valColor = ventas < 0 ? 'var(--danger)' : 'inherit';
                    const prevColor = prevVentas < 0 ? 'var(--danger)' : 'inherit';
                    const pptoColor = pptoVentas < 0 ? 'var(--danger)' : 'inherit';

                    return `<tr>
                        <td style="font-weight:600">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 2px;">
                                <span>${formatSegmentName(name)}</span>
                                <span style="font-size: 0.8rem; color: var(--text-secondary); font-weight: 500;">${pctPart.toFixed(1)}%</span>
                            </div>
                            <div class="bar-container"><div class="bar-fill" style="width: ${Math.min(100, Math.max(0, pctPart))}%;"></div></div>
                        </td>
                        <td style="color:${prevColor}">${formatCurrency(prevVentas)}</td>
                        <td style="color:${valColor}">${formatCurrency(ventas)}</td>
                        <td style="color:${pptoColor}">${formatCurrency(pptoVentas)}</td>
                        <td style="color:${color}">${diff >= 0 ? '+' : ''}${formatCurrency(diff)} (${pctMoM > 0 ? '+' : ''}${pctMoM.toFixed(1)}%)</td>
                        <td style="color:${colorPpto}">${diffPpto >= 0 ? '+' : ''}${formatCurrency(diffPpto)} (${pctPpto > 0 ? '+' : ''}${pctPpto.toFixed(1)}%)</td>
                    </tr>`;
                }).join('');
            } else {
                segmentsSection.style.display = 'none';
            }
            
            // Render Segmentos Costos
            const costSegmentsSection = document.getElementById('cost-segments-section');
            const costSegmentsBody = document.getElementById('costSegmentsBody');
            if (segmentsList.length > 0) {
                costSegmentsSection.style.display = 'block';
                costSegmentsBody.innerHTML = segmentsList.map(([name, dataSeg]) => {
                    const costos = dataSeg.costos || 0;
                    const prevCostos = prevSegments[name] ? prevSegments[name].costos : 0;
                    const pptoCostos = pptoSegments[name] ? pptoSegments[name].costos : 0;
                    
                    const diff = costos - prevCostos;
                    const diffPpto = costos - pptoCostos;
                    const pctVar = prevCostos !== 0 ? (diff / Math.abs(prevCostos)) * 100 : 0;
                    const pctVarPpto = pptoCostos !== 0 ? (diffPpto / Math.abs(pptoCostos)) * 100 : 0;
                    
                    const color = diff >= 0 ? 'var(--success)' : 'var(--danger)';
                    const colorPpto = diffPpto >= 0 ? 'var(--success)' : 'var(--danger)';
                    const valColor = costos < 0 ? 'var(--danger)' : 'inherit';
                    const prevColor = prevCostos < 0 ? 'var(--danger)' : 'inherit';
                    const pptoColor = pptoCostos < 0 ? 'var(--danger)' : 'inherit';
        
                    return `<tr>
                        <td style="font-weight:600">${formatSegmentName(name)}</td>
                        <td style="color:${prevColor}">${formatCurrency(prevCostos)}</td>
                        <td style="color:${valColor}">${formatCurrency(costos)}</td>
                        <td style="color:${pptoColor}">${formatCurrency(pptoCostos)}</td>
                        <td style="color:${color}">${diff >= 0 ? '+' : ''}${formatCurrency(diff)} (${pctVar > 0 ? '+' : ''}${pctVar.toFixed(1)}%)</td>
                        <td style="color:${colorPpto}">${diffPpto >= 0 ? '+' : ''}${formatCurrency(diffPpto)} (${pctVarPpto > 0 ? '+' : ''}${pctVarPpto.toFixed(1)}%)</td>
                    </tr>`;
                }).join('');
            } else {
                costSegmentsSection.style.display = 'none';
            }
            
            // Render Margen por segmento
            const margenSegmentsSection = document.getElementById('margen-segments-section');
            const margenSegmentsBody = document.getElementById('margenSegmentsBody');
            if (segmentsList.length > 0) {
                if(margenSegmentsSection) margenSegmentsSection.style.display = 'block';
                if(margenSegmentsBody) margenSegmentsBody.innerHTML = segmentsList.map(([name, dataSeg]) => {
                    const ventas = dataSeg.ventas || 0;
                    const costos = dataSeg.costos || 0;
                    const prevVentas = prevSegments[name] ? prevSegments[name].ventas : 0;
                    const prevCostos = prevSegments[name] ? prevSegments[name].costos : 0;
                    const pptoVentas = pptoSegments[name] ? pptoSegments[name].ventas : 0;
                    const pptoCostos = pptoSegments[name] ? pptoSegments[name].costos : 0;
                    
                    const margen = Math.abs(ventas) - Math.abs(costos);
                    const prevMargen = Math.abs(prevVentas) - Math.abs(prevCostos);
                    const pptoMargen = Math.abs(pptoVentas) - Math.abs(pptoCostos);
                    
                    const pctMargen = ventas !== 0 ? (margen / Math.abs(ventas)) * 100 : 0;
                    const pctPrevMargen = prevVentas !== 0 ? (prevMargen / Math.abs(prevVentas)) * 100 : 0;
                    const pctPptoMargen = pptoVentas !== 0 ? (pptoMargen / Math.abs(pptoVentas)) * 100 : 0;
                    
                    const diffPct = pctMargen - pctPrevMargen;
                    const diffPctPpto = pctMargen - pctPptoMargen;
                    
                    const color = diffPct >= 0 ? 'var(--success)' : 'var(--danger)';
                    const colorPpto = diffPctPpto >= 0 ? 'var(--success)' : 'var(--danger)';
                    const valColor = margen < 0 ? 'var(--danger)' : 'inherit';
                    const prevColor = prevMargen < 0 ? 'var(--danger)' : 'inherit';
        
                    return `<tr>
                        <td style="font-weight:600">${formatSegmentName(name)}</td>
                        <td>${pctPrevMargen.toFixed(1)}%</td>
                        <td style="font-weight:700">${pctMargen.toFixed(1)}%</td>
                        <td>${pctPptoMargen.toFixed(1)}%</td>
                        <td style="color:${color}; font-weight:700">${diffPct > 0 ? '+' : ''}${diffPct.toFixed(1)} pp</td>
                        <td style="color:${colorPpto}; font-weight:700">${diffPctPpto > 0 ? '+' : ''}${diffPctPpto.toFixed(1)} pp</td>
                    </tr>`;
                }).join('');
            } else {
                if(margenSegmentsSection) margenSegmentsSection.style.display = 'none';
            }
            
            // Render OPEX Detalle
            const opexSection = document.getElementById('opex-section');
            const opexBody = document.getElementById('opexBody');
            const opexDetalle = (currAgg.pnl && currAgg.pnl.opexDetalle) ? currAgg.pnl.opexDetalle : {};
            const prevOpexDetalle = (prevAgg.pnl && prevAgg.pnl.opexDetalle) ? prevAgg.pnl.opexDetalle : opexDetalle;
            const pptoOpexDetalle = (currAgg.ppto && currAgg.ppto.pnl && currAgg.ppto.pnl.opexDetalle) ? currAgg.ppto.pnl.opexDetalle : {};
        
            if (Object.keys(opexDetalle).length > 0) {
                opexSection.style.display = 'block';
                opexBody.innerHTML = Object.entries(opexDetalle).map(([cat, val]) => {
                    const prevVal = prevOpexDetalle[cat] || 0;
                    const pptoVal = pptoOpexDetalle[cat] || 0;
                    const diff = val - prevVal; 
                    const diffPpto = val - pptoVal;
                    const pct = prevVal !== 0 ? (diff / Math.abs(prevVal)) * 100 : 0;
                    const pctPpto = pptoVal !== 0 ? (diffPpto / Math.abs(pptoVal)) * 100 : 0;
                    const color = diff >= 0 ? 'var(--success)' : 'var(--danger)'; 
                    const colorPpto = diffPpto >= 0 ? 'var(--success)' : 'var(--danger)'; 
                    const valColor = val < 0 ? 'var(--danger)' : 'inherit';
                    const prevColor = prevVal < 0 ? 'var(--danger)' : 'inherit';
                    const pptoColor = pptoVal < 0 ? 'var(--danger)' : 'inherit';
        
                    return `<tr>
                        <td style="font-weight:600">${cat}</td>
                        <td style="color:${prevColor}">${formatCurrency(prevVal)}</td>
                        <td style="color:${valColor}">${formatCurrency(val)}</td>
                        <td style="color:${pptoColor}">${formatCurrency(pptoVal)}</td>
                        <td style="color:${color}">${diff >= 0 ? '+' : ''}${formatCurrency(diff)} (${pct > 0 ? '+' : ''}${pct.toFixed(1)}%)</td>
                        <td style="color:${colorPpto}">${diffPpto >= 0 ? '+' : ''}${formatCurrency(diffPpto)} (${pctPpto > 0 ? '+' : ''}${pctPpto.toFixed(1)}%)</td>
                    </tr>`;
                }).join('');
            } else {
                opexSection.style.display = 'none';
            }

            // Render Tabla Detallada General
            const tableBody = document.getElementById('tableBody');
            if (Object.keys(aggCategories).length > 0) {
                const filteredEntries = Object.entries(aggCategories).filter(([cat]) => 
                    !cat.toLowerCase().includes("opex") && 
                    !cat.toLowerCase().includes("extraordinarios") &&
                    !cat.toLowerCase().includes("utilidad")
                );
                
                tableBody.innerHTML = filteredEntries.map(([cat, val]) => {
                    const prevVal = prevAggCategories[cat] || 0;
                    const pptoVal = aggPptoCategories[cat] || 0;
                    const diff = val - prevVal;
                    const pct = prevVal !== 0 ? (diff / Math.abs(prevVal)) * 100 : 0;
                    const diffPpto = val - pptoVal;
                    const pctPpto = pptoVal !== 0 ? (diffPpto / Math.abs(pptoVal)) * 100 : 0;
                    
                    const color = diff >= 0 ? 'var(--success)' : 'var(--danger)';
                    const colorPpto = diffPpto >= 0 ? 'var(--success)' : 'var(--danger)';
                    const valColor = val < 0 ? 'var(--danger)' : 'inherit';
                    const prevColor = prevVal < 0 ? 'var(--danger)' : 'inherit';
                    const pptoColor = pptoVal < 0 ? 'var(--danger)' : 'inherit';
                    
                    return `<tr>
                        <td style="font-weight:600">${cat}</td>
                        <td style="color:${prevColor}">${formatCurrency(prevVal)}</td>
                        <td style="color:${valColor}">${formatCurrency(val)}</td>
                        <td style="color:${pptoColor}">${formatCurrency(pptoVal)}</td>
                        <td style="color:${color}">${diff >= 0 ? '+' : ''}${formatCurrency(diff)} (${pct > 0 ? '+' : ''}${pct.toFixed(1)}%)</td>
                        <td style="color:${colorPpto}">${diffPpto >= 0 ? '+' : ''}${formatCurrency(diffPpto)} (${pctPpto > 0 ? '+' : ''}${pctPpto.toFixed(1)}%)</td>
                    </tr>`;
                }).join('');
            }

            // Build Mobile views
            setTimeout(() => {
                buildMobileAccordionsFromTable('tableResumenOperativo', 'resumenOperativoMobileContainer', 'Resumen Operativo', '');
                buildMobileAccordionsFromTable('tableVentasSegmento', 'ventasSegmentoMobileContainer', 'Segmentos de Venta', formatCurrency(aggKpis.ingresos));
                buildMobileAccordionsFromTable('tableCostosSegmento', 'costosSegmentoMobileContainer', 'Desglose de Costos', formatCurrency(aggTotalCost));
                buildMobileAccordionsFromTable('tableMargenSegmento', 'margenSegmentoMobileContainer', 'Margen Bruto por Segmento', '');
                
                const currOpex = (currAgg.pnl && currAgg.pnl.opexDetalle) ? Object.values(currAgg.pnl.opexDetalle).reduce((acc, val) => acc + val, 0) : 0;
                buildMobileAccordionsFromTable('tableOpex', 'opexMobileContainer', 'Detalle de Gastos OPEX', formatCurrency(currOpex));
                
                // Trigger account search filter if active
                const searchInput = document.getElementById('accountSearch');
                if (searchInput && searchInput.value.trim() !== '') {
                    searchInput.dispatchEvent(new Event('input'));
                }
            }, 10);
        }
        
        let viewVentasCeo = document.getElementById("view-ventas-ceo");
        if (viewVentasCeo && viewVentasCeo.classList.contains("active")) {
            if (typeof window.renderVentasCEO === "function") {
                window.renderVentasCEO();
            }
        }
    });
}
/**
 * Helper to identify periods
 */
function isYear2025(d) {
    if (!d) return false;
    const dt = d.sortDate;
    const normDate = normalizeText(d.date || "");
    
    if (dt && typeof dt.getFullYear === 'function' && dt.getFullYear() === 2025) return true;
    if (dt && typeof dt === 'string') {
        const dObj = new Date(dt);
        if (!isNaN(dObj) && dObj.getFullYear() === 2025) return true;
    }
    if (normDate.includes("2025") || normDate.includes("-25") || normDate.includes("/25") || normDate.includes(" 25")) return true;
    
    return false;
}

function isYear2026(d) {
    if (!d) return false;
    const dt = d.sortDate;
    const normDate = normalizeText(d.date || "");
    
    if (dt && typeof dt.getFullYear === 'function' && dt.getFullYear() === 2026) return true;
    if (dt && typeof dt === 'string') {
        const dObj = new Date(dt);
        if (!isNaN(dObj) && dObj.getFullYear() === 2026) return true;
    }
    if (normDate.includes("2026") || normDate.includes("-26") || normDate.includes("/26") || normDate.includes(" 26")) return true;
    
    return false;
}

/**
 * Render the Balance Sheet Table
 */
function renderBalanceSheet(data, selectedIndex = -1) {
    const headerEl = document.getElementById('balanceHeader');
    const bodyEl = document.getElementById('balanceBody');
    const periodLabel = document.getElementById('balancePeriodLabel');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    const curr = data[endIdx];
    
    const startIdx = Math.max(0, endIdx - 5);
    let visibleMonths = data.slice(startIdx, endIdx + 1);

    // Fix Diciembre 2025 as the first column, filter out the rest of 2025
    visibleMonths = visibleMonths.filter(m => isYear2026(m));
    const dic2025Balance = data.find(d => isYear2025(d) && (d.date.toLowerCase().includes('dic') || d.date.toLowerCase().includes('dec')));
    if (dic2025Balance && !visibleMonths.includes(dic2025Balance)) {
        visibleMonths.unshift(dic2025Balance);
    }
    
    const periods = visibleMonths.map(d => d.date);

    periodLabel.textContent = `Situación Financiera al ${curr.date}`;

    headerEl.innerHTML = `
        <tr>
            <th>Concepto / Cuenta de Balance</th>
            ${periods.map(p => `<th>${p}</th>`).join('')}
        </tr>
    `;

    // Extract concepts
    let allConcepts = [];
    visibleMonths.forEach(d => {
        if (d.balance && d.balance.fullRows) {
            d.balance.fullRows.forEach(row => {
                if (!allConcepts.includes(row.concept)) allConcepts.push(row.concept);
            });
        }
    });

    // 1. Clasificación: Balance vs Covenants
    let balanceConcepts = [];
    let covenantConcepts = [];
    
    allConcepts.forEach(c => {
        const n = normalizeText(c);
        
        // Filtros solicitados por usuario (REA, redundantes y filas técnicas)
        if (n === "covenant deuda" || n === "rea" || n.trim() === "" || 
            n === "pasivos excluye deuda subordinada" || 
            n === "patrimonio incluye deuda subordinada" || 
            n === "pasivos - deuda subordinada") return;

        const isDebtRow = n.includes("deuda bruta") || n.includes("deuda total") || n.includes("deuda subordinada") || 
                          n.includes("deuda sin subordinada") || n.includes("deuda neta") || n.includes("ebitda");
        const isRatioRow = n.includes("apalancamiento") || n.includes("capacidad") || n.includes("razon corriente") || n.includes("covenant");
        
        // Efectivo es parte del bloque si está rodeado de deuda
        const isCovenant = isDebtRow || isRatioRow || n === "efectivo" || n.includes("ebitda r12");
        
        if (isCovenant) covenantConcepts.push(c);
        // Efectivo debe estar en AMBOS (Covenant y Balance)
        if (!isCovenant || n === "efectivo") balanceConcepts.push(c);
    });

    // Ordenamiento Estricto según imagen del Excel (Linear Extraction Mode)
    const getCovenantRank = (concept) => {
        const n = normalizeText(concept);
        if (n === "deuda bruta") return 1;
        if (n === "efectivo") return 2;
        if (n === "deuda neta" && !n.includes("ebitda") && !n.includes("subordinada")) return 3;
        if (n.includes("ltm ebitda") || (n.includes("ebitda") && !n.includes("ratio") && !n.includes("r12"))) return 4;
        
        if (n === "deuda total") return 5;
        if (n === "deuda subordinada") return 6;
        if (n === "deuda sin subordinada") return 7;
        if (n === "deuda neta sin subordinada") return 9;

        if (n.includes("deuda neta/ebitda") || n.includes("r12")) return 10;
        if (n.includes("apalancamiento")) return 11;
        if (n.includes("capacidad")) return 12;
        if (n.includes("razon corriente") || n.includes("corriente")) return 13;
        return 100;
    };

    covenantConcepts.sort((a, b) => getCovenantRank(a) - getCovenantRank(b));

    // 2. Filtro y Reordenamiento para el Balance
    let filteredBalance = [];
    let isSkipping = false;
    let foundGrandTotal = false;
    
    // Identificar posiciones especiales
    const utilidadesRetenidasIdx = balanceConcepts.findIndex(c => normalizeText(c).includes("utilidades retenidas"));
    const beneficioNetoConcept = balanceConcepts.find(c => {
        const n = normalizeText(c);
        return (n.includes("beneficio neto") && !n.includes("utilidades")) || 
               (n.includes("utilidad del ejercicio") && !n.includes("retenidas")) || 
               (n.includes("ganancia del periodo")) ||
               (n.includes("resultado del ejercicio") && !n.includes("retenidas"));
    });
    
    balanceConcepts.forEach(concept => {
        if (foundGrandTotal) return;
        const norm = normalizeText(concept);
        
        // Regla de Parada 1: Eliminar redundantes y cabeceras de Excel
        if (norm === "total pasivo y patrimonio" || norm === "total pasivo y capital") return;
        if (norm === "concepto" || norm === "cuentas" || norm === "descripcion" || norm === "balance sheet" || norm === "detalle") return;
        
        // Exclusión agresiva de Activos/Pasivos/Patrimonio como cabeceras puras
        const isHeaderOnly = norm === "activos" || norm === "pasivos" || norm === "patrimonio" || 
                             norm === "capital" || norm === "pasivo y capital" || 
                             norm === "activo" || norm === "pasivo" ||
                             norm.startsWith("activos:") || norm.startsWith("pasivos:") || 
                             norm.startsWith("patrimonio:");
        
        if (isHeaderOnly) return;
        if (norm.includes("estado de situacion") || norm.includes("reporte pa") || norm.includes("en mdop")) return;

        // Regla de Parada Final Fuerte
        if (norm.includes("total pasivo") && (norm.includes("capital") || norm.includes("accionista"))) {
            if (concept.trim().length > 10) { // Evitar falsos positivos cortos
                filteredBalance.push(concept);
                foundGrandTotal = true;
                return;
            }
        }

        // Evitar duplicar beneficio neto (se insertará debajo de utilidades retenidas)
        if (concept === beneficioNetoConcept && utilidadesRetenidasIdx !== -1) return;

        // Limpieza de firmas
        const isPatrimonioItem = norm.includes("utilidad") || norm.includes("beneficio") || norm.includes("ganancia") || 
                                 norm.includes("reserva") || norm.includes("capital") || norm.includes("patrimonio") || 
                                 norm.includes("rea") || norm.includes("resultados acumulados") || 
                                 norm.includes("ajuste") || norm.includes("manos de terceros");

        if (isSkipping && isPatrimonioItem) isSkipping = false;

        if (!isSkipping || isPatrimonioItem) {
            filteredBalance.push(concept);
            // Inserción de Beneficio Neto debajo de Utilidades Retenidas
            if (concept === balanceConcepts[utilidadesRetenidasIdx] && beneficioNetoConcept) {
                if (!filteredBalance.some(c => c === beneficioNetoConcept)) {
                    filteredBalance.push(beneficioNetoConcept);
                }
            }
        }
    });

    // Asegurar que "Efectivo" sea la primera fila del Balance General Consolidado
    const cashIndex = filteredBalance.findIndex(c => normalizeText(c) === "efectivo");
    if (cashIndex > 0) {
        const cashRow = filteredBalance.splice(cashIndex, 1)[0];
        filteredBalance.unshift(cashRow);
    }

    const renderRows = (concepts, targetBodyId) => {
        const bodyEl = document.getElementById(targetBodyId);
        if (!bodyEl) return;
        
        bodyEl.innerHTML = concepts.map(concept => {
            const norm = normalizeText(concept);
            const labelLower = norm;
            const isTotal = labelLower.includes("total") || labelLower.includes("sumas") || 
                            labelLower.includes("activo") || labelLower.includes("pasivo") || 
                            labelLower.includes("patrimonio") ||
                            labelLower.includes("ebitda") || labelLower.includes("apalancamiento") || 
                            labelLower.includes("capacidad de pago") || labelLower.includes("razon corriente");
            const isSubRow = concept.startsWith("  ") || concept.startsWith("\t") || 
                             norm.includes("acumulado") || norm.includes("depreciacion") || norm.includes("impuestos") ||
                             norm.includes("ganancia acumulada") || 
                             norm.includes("beneficio neto") || norm.includes("ganancia del periodo") ||
                             norm.includes("resultado del ejercicio") ||
                             norm.includes("activo en manos de terceros");
            
            const isMainCategory = (concept.trim() === concept.trim().toUpperCase() || 
                                    labelLower.includes("activos") || labelLower.includes("pasivos") || 
                                    labelLower.includes("patrimonio") || norm.includes("covenant") ||
                                    norm.includes("utilidades retenidas") || norm.includes("revaluacion de activos")) 
                                    && !isTotal && concept.trim().length > 3;

            const periodCells = visibleMonths.map(period => {
                const row = period.balance?.fullRows?.find(r => r.concept === concept);
                let val = row ? row.values[period.date] || 0 : 0;

                // Fallback para Beneficio Neto: si es 0 en el balance, tomarlo del P&L
                if (norm.includes("beneficio neto") || norm.includes("ganancia del periodo") || 
                    norm.includes("utilidad del ejercicio") || norm.includes("resultado del ejercicio") || 
                    norm.includes("utilidad neta") || norm.includes("ganancia neta")) {
                    if (val === 0 && period.pnl?.netIncome) {
                        val = period.pnl.netIncome;
                    }
                }
                
                const isRatio = (norm.includes("ratio") || norm.includes("indice") || norm.includes("razon") || 
                                 norm.includes("apalancamiento") || norm.includes("capacidad") || 
                                 norm.includes("ebitda r12") || norm.includes("ebitda ltm") ||
                                 norm.includes("deuda neta/ebitda") || concept.includes(" x ") || concept.endsWith(" x")) && 
                                 !norm.includes("cxp") && !norm.includes("otras cxp") && !norm.includes("cxc") && !norm.includes("pagar") && !norm.includes("cobrar");
                
                const color = val < 0 ? 'var(--danger)' : 'inherit';
                let displayVal;
                
                if (isRatio) {
                    displayVal = (val !== 0) ? (typeof val === 'number' ? val.toFixed(2) : val) + 'x' : "-";
                } else if ((norm.includes("covenant") || norm.includes("apalancamiento") || norm.includes("capacidad") || norm.includes("razon corriente") || norm.includes("ebitda r12")) && val !== 0 && !norm.includes("mdo") && !norm.includes("pagar") && !norm.includes("cobrar")) {
                    displayVal = (typeof val === 'number' ? val.toFixed(2) : val) + 'x';
                } else {
                    displayVal = formatCurrency(val);
                }

                // Si es una categoría principal y el valor es 0, ocultamos el valor para evitar confusión
                if (isMainCategory && val === 0) displayVal = "";

                return `<td style="color:${color}">${displayVal}</td>`;
            }).join('');

            let displayLabel = concept;
            if (norm === "ganancia del periodo") displayLabel = "Beneficio Neto del Periodo";
            
            let rowClass = isTotal ? 'row-total' : '';
            if (isMainCategory && !isSubRow) rowClass = 'row-category';
            
            const cellClass = isSubRow ? 'row-indent' : '';

            return `<tr class="${rowClass}">
                <td class="${cellClass}">${displayLabel}</td>
                ${periodCells}
            </tr>`;
        }).join('');
    };

    if (filteredBalance.length === 0) {
        bodyEl.innerHTML = `<tr><td colspan="${periods.length + 1}" style="text-align:center; padding:40px;">No se encontraron filas detalladas de Balance.</td></tr>`;
    } else {
        renderRows(filteredBalance, 'balanceBody');
    }
    
    // Render Covenant Section
    if (covenantConcepts.length > 0) {
        document.getElementById('covenant-section').style.display = 'block';
        document.getElementById('covenantHeader').innerHTML = `
            <tr>
                <th>Concepto / Ratio de Deuda</th>
                ${periods.map(p => `<th>${p}</th>`).join('')}
            </tr>
        `;
        renderRows(covenantConcepts, 'covenantBody');
    } else {
        document.getElementById('covenant-section').style.display = 'none';
    }

    renderBalanceResumen(dic2025Balance || visibleMonths[0], curr);
}

function renderBalanceResumen(dicData, currData) {
    const headerEl = document.getElementById('balanceResumenHeader');
    const bodyEl = document.getElementById('balanceResumenBody');
    if (!headerEl || !bodyEl) return;

    const dicLabel = dicData ? dicData.date : 'N/A';
    const currLabel = currData ? currData.date : 'N/A';
    
    headerEl.innerHTML = `
        <tr>
            <th style="background: var(--sidebar); color: white;">Concepto</th>
            <th style="text-align: right; background: var(--sidebar-dark); color: white;">${dicLabel}</th>
            <th style="text-align: right; background: var(--sidebar-dark); color: white;">${currLabel}</th>
            <th style="text-align: right; background: var(--sidebar-dark); color: white;">PPTO</th>
        </tr>
    `;

    if (!currData) return;

    const getVal = (source, labelMatch, sumAll = false, excludeMatch = null, exactMatch = false) => {
        if (!source || !source.balance || !source.balance.fullRows) return 0;
        const matches = Array.isArray(labelMatch) ? labelMatch : [labelMatch];
        const excludes = excludeMatch ? (Array.isArray(excludeMatch) ? excludeMatch : [excludeMatch]) : [];
        const dateKey = source.date || currData.date;

        const normalizeLocal = (str) => {
            if (!str) return "";
            return str.toString()
                .toLowerCase()
                .normalize("NFD")
                .replace(/[\u0300-\u036f]/g, "")
                .replace(/[.:;()]/g, " ")
                .replace(/\s+/g, " ")
                .trim();
        };

        const matchesNorm = matches.map(m => normalizeLocal(m));
        const excludesNorm = excludes.map(e => normalizeLocal(e));

        const rowsToProcess = (() => {
            let idx = source.balance.fullRows.findIndex(r => {
                const c = normalizeLocal(r.concept);
                return c.includes("total pasivo") && (c.includes("capital") || c.includes("accionista") || c.includes("patrimonio"));
            });
            if (idx === -1) idx = source.balance.fullRows.length;
            return source.balance.fullRows.slice(0, Math.max(0, idx + 1));
        })();

        if (sumAll) {
            let sum = 0;
            rowsToProcess.forEach(r => {
                const cNorm = normalizeLocal(r.concept);
                const isMatch = exactMatch ? matchesNorm.some(m => cNorm === m) : matchesNorm.some(m => cNorm.includes(m));
                if (isMatch && !excludesNorm.some(e => cNorm.includes(e))) {
                    if (r.values) {
                        sum += r.values[dateKey] || 0;
                    }
                }
            });
            return sum;
        } else {
            const row = rowsToProcess.find(r => {
                const cNorm = normalizeLocal(r.concept);
                const isMatch = exactMatch ? matchesNorm.some(m => cNorm === m) : matchesNorm.some(m => cNorm.includes(m));
                return isMatch && !excludesNorm.some(e => cNorm.includes(e));
            });
            if (row && row.values) return row.values[dateKey] || 0;
        }
        return 0;
    };

    const conceptRow20 = (currData.balance && currData.balance.conceptRow20) || (dicData && dicData.balance && dicData.balance.conceptRow20) || "";
    const intangibleMatch = conceptRow20 ? [conceptRow20, 'intangible'] : ['intangible'];

    const rowSpec = [
        { label: 'Efectivo', match: ['efectivo', 'inversion en cd'], sumAll: true },
        { label: 'CXC', match: 'cobrar', sumAll: true },
        { label: 'Inventarios', match: 'inventario' },
        { label: 'Gastos Pagados por Anticipado', match: 'pagados por anticipado' },
        { label: 'PPE', match: 'propiedad' },
        { label: 'Inversión en Acciones', match: 'acciones' },
        { label: 'Bienes Intangibles', match: intangibleMatch },
        { label: 'Impuestos', match: ['impuesto', 'taxes'] },
        { label: 'Otros Activos', match: 'otros activos' },
        { label: 'Total Activos', match: 'total activo', highlight: true, overlayBg: 'rgba(0,150,199,0.1)' },
        { type: 'separator' },
        { label: 'CXP', match: 'cuentas por pagar', exactMatch: true },
        { label: 'Deuda Financiera CP', match: 'corto plazo', exclude: 'relacionad' },
        { label: 'Deuda Financiera LP', match: 'largo plazo', exclude: 'relacionad' },
        { label: 'Deuda Accionista', match: ['deuda corto plazo relacionadas', 'deuda largo plazo relacionadas', 'deuda accionista'], sumAll: true },
        { label: 'Otros Pasivos', match: 'otros pasivos' },
        { label: 'Total Pasivos', match: 'total pasivos', exactMatch: true, highlight: true, overlayBg: 'rgba(0,150,199,0.1)' },
        { type: 'separator' },
        { label: 'Capital Suscrito y Pagado, Reserva Legal', match: ['suscrito y pagado', 'reserva legal', 'aporte para futuras'], sumAll: true },
        { label: 'Revaluaciones', match: 'revaluacion' },
        { label: 'Ganancias (Perdida) Acumuladas', match: 'acumulada' },
        { label: 'Ganancia del Periodo', match: ['beneficio neto del periodo', 'beneficio neto'] },
        { label: 'Total Patrimonio', match: 'total patrimonio', exactMatch: true, highlight: true, overlayBg: 'rgba(0,150,199,0.1)' },
        { label: 'Total Pasivo y Patrimonio', match: 'total pasivo y patrimonio', exactMatch: true, highlight: true, overlayBg: 'var(--sidebar)' }
    ];

    // Compute all values first
    const evaluatedRows = rowSpec.map(spec => {
        if (spec.type === 'separator') return { type: 'separator' };

        let vDic = getVal(dicData, spec.match, spec.sumAll, spec.exclude, spec.exactMatch);
        let vCurr = getVal(currData, spec.match, spec.sumAll, spec.exclude, spec.exactMatch);
        let vPpto = getVal(currData.ppto, spec.match, spec.sumAll, spec.exclude, spec.exactMatch);
        
        if (spec.label === 'Ganancia del Periodo') {
            const findBenNeto = (src, explicitDate) => {
                let res = 0;
                if (!src) return res;
                const dKey = explicitDate || src.date;
                const searchRow = (rows) => {
                    if (!rows) return null;
                    return [...rows].reverse().find(r => r.concept && (r.concept.toLowerCase().includes('beneficio neto') || r.concept.toLowerCase().includes('ganancia del periodo')));
                };
                let r = searchRow(src.balance?.fullRows) || searchRow(src.pnl?.fullRows);
                if (r && r.values) {
                    res = r.values[dKey] !== undefined ? r.values[dKey] : (Object.values(r.values)[0] || 0);
                }
                return res || src.pnl?.netIncome || 0;
            };
            
            if (vDic === 0) vDic = findBenNeto(dicData, dicData.date);
            if (vCurr === 0) vCurr = findBenNeto(currData, currData.date);
            if (vPpto === 0 && currData.ppto) vPpto = findBenNeto(currData.ppto, currData.date);
        }

        // Use PPTO object's period/date context for correct row evaluation fallback
        if (!vPpto && currData.ppto && currData.ppto.balance && currData.ppto.balance.fullRows) {
            const matches = Array.isArray(spec.match) ? spec.match : [spec.match];
            const excludes = spec.exclude ? (Array.isArray(spec.exclude) ? spec.exclude : [spec.exclude]) : [];
            
            const normalizeLocal = (str) => {
                if (!str) return "";
                return str.toString()
                    .toLowerCase()
                    .normalize("NFD")
                    .replace(/[\u0300-\u036f]/g, "")
                    .replace(/[.:;()]/g, " ")
                    .replace(/\s+/g, " ")
                    .trim();
            };

            const matchesNorm = matches.map(m => normalizeLocal(m));
            const excludesNorm = excludes.map(e => normalizeLocal(e));
            
            const rowsToProcess = (() => {
                let idx = currData.ppto.balance.fullRows.findIndex(r => {
                    const c = normalizeLocal(r.concept);
                    return c.includes("total pasivo") && (c.includes("capital") || c.includes("accionista") || c.includes("patrimonio"));
                });
                if (idx === -1) idx = currData.ppto.balance.fullRows.length;
                return currData.ppto.balance.fullRows.slice(0, Math.max(0, idx + 1));
            })();

            if (spec.sumAll) {
                let sum = 0;
                rowsToProcess.forEach(r => {
                    const cNorm = normalizeLocal(r.concept);
                    const isMatch = spec.exactMatch ? matchesNorm.some(m => cNorm === m) : matchesNorm.some(m => cNorm.includes(m));
                    if (isMatch && !excludesNorm.some(e => cNorm.includes(e))) {
                        if (r.values) {
                            const vals = Object.values(r.values);
                            if (vals.length > 0) sum += vals[0] || 0;
                        }
                    }
                });
                vPpto = sum;
            } else {
                const row = rowsToProcess.find(r => {
                    const cNorm = normalizeLocal(r.concept);
                    const isMatch = spec.exactMatch ? matchesNorm.some(m => cNorm === m) : matchesNorm.some(m => cNorm.includes(m));
                    return isMatch && !excludesNorm.some(e => cNorm.includes(e));
                });
                if(row && row.values) {
                    const vals = Object.values(row.values);
                    if(vals.length > 0) vPpto = vals[0] || 0;
                }
            }
        }

        return {
            label: spec.label,
            match: spec.match,
            highlight: spec.highlight,
            overlayBg: spec.overlayBg,
            vDic,
            vCurr,
            vPpto
        };
    });

    // Substraer "Impuestos" de "Otros Activos" porque en la fuente figuran como subcategoría
    const otrosActivosItem = evaluatedRows.find(r => r.label === 'Otros Activos');
    const impuestosItem = evaluatedRows.find(r => r.label === 'Impuestos');
    if (otrosActivosItem && impuestosItem) {
        otrosActivosItem.vDic = (otrosActivosItem.vDic || 0) - (impuestosItem.vDic || 0);
        otrosActivosItem.vCurr = (otrosActivosItem.vCurr || 0) - (impuestosItem.vCurr || 0);
        otrosActivosItem.vPpto = (otrosActivosItem.vPpto || 0) - (impuestosItem.vPpto || 0);
    }

    // Calcular "Otros Pasivos" dinámicamente = Total Pasivos de la fuente - CXP - Deuda Financiera CP - Deuda Financiera LP - Deuda Accionista
    const otrosPasivosItem = evaluatedRows.find(r => r.label === 'Otros Pasivos');
    const cxpItem = evaluatedRows.find(r => r.label === 'CXP');
    const dfcpItem = evaluatedRows.find(r => r.label === 'Deuda Financiera CP');
    const dflpItem = evaluatedRows.find(r => r.label === 'Deuda Financiera LP');
    const daItem = evaluatedRows.find(r => r.label === 'Deuda Accionista');
    const totalPasivosItem = evaluatedRows.find(r => r.label === 'Total Pasivos');

    if (totalPasivosItem && otrosPasivosItem) {
        const sumOthers = (item) => item ? (item.vDic || 0) : 0;
        const sumOthersCurr = (item) => item ? (item.vCurr || 0) : 0;
        const sumOthersPpto = (item) => item ? (item.vPpto || 0) : 0;

        const subtotalDic = sumOthers(cxpItem) + sumOthers(dfcpItem) + sumOthers(dflpItem) + sumOthers(daItem);
        const subtotalCurr = sumOthersCurr(cxpItem) + sumOthersCurr(dfcpItem) + sumOthersCurr(dflpItem) + sumOthersCurr(daItem);
        const subtotalPpto = sumOthersPpto(cxpItem) + sumOthersPpto(dfcpItem) + sumOthersPpto(dflpItem) + sumOthersPpto(daItem);

        otrosPasivosItem.vDic = totalPasivosItem.vDic - subtotalDic;
        otrosPasivosItem.vCurr = totalPasivosItem.vCurr - subtotalCurr;
        otrosPasivosItem.vPpto = totalPasivosItem.vPpto - subtotalPpto;
    }

    // Calcular "Ganancias (Perdida) Acumuladas" dinámicamente
    const gananciasAcumItem = evaluatedRows.find(r => r.label === 'Ganancias (Perdida) Acumuladas');
    const capSusItem = evaluatedRows.find(r => r.label === 'Capital Suscrito y Pagado, Reserva Legal');
    const revalItem = evaluatedRows.find(r => r.label === 'Revaluaciones');
    const gananciaPerItem = evaluatedRows.find(r => r.label === 'Ganancia del Periodo');
    const totPatriSource = evaluatedRows.find(r => r.label === 'Total Patrimonio');

    if (gananciasAcumItem && totPatriSource) {
        gananciasAcumItem.vDic = (totPatriSource.vDic || 0) - ((capSusItem?.vDic || 0) + (revalItem?.vDic || 0) + (gananciaPerItem?.vDic || 0));
        gananciasAcumItem.vCurr = (totPatriSource.vCurr || 0) - ((capSusItem?.vCurr || 0) + (revalItem?.vCurr || 0) + (gananciaPerItem?.vCurr || 0));
        gananciasAcumItem.vPpto = (totPatriSource.vPpto || 0) - ((capSusItem?.vPpto || 0) + (revalItem?.vPpto || 0) + (gananciaPerItem?.vPpto || 0));
    }

    const sumByLabels = (labels) => {
        let sd = 0, sc = 0, sp = 0;
        labels.forEach(lbl => {
            const item = evaluatedRows.find(r => r.label === lbl);
            if (item) {
                sd += item.vDic || 0;
                sc += item.vCurr || 0;
                sp += item.vPpto || 0;
            }
        });
        return { vDic: sd, vCurr: sc, vPpto: sp };
    };

    const activosLabels = ['Efectivo', 'CXC', 'Inventarios', 'Gastos Pagados por Anticipado', 'PPE', 'Inversión en Acciones', 'Bienes Intangibles', 'Impuestos', 'Otros Activos'];
    const pasivosLabels = ['CXP', 'Deuda Financiera CP', 'Deuda Financiera LP', 'Deuda Accionista', 'Otros Pasivos'];
    const patrimonioLabels = ['Capital Suscrito y Pagado, Reserva Legal', 'Revaluaciones', 'Ganancias (Perdida) Acumuladas', 'Ganancia del Periodo'];

    const assetsSum = sumByLabels(activosLabels);
    const liabsSum = sumByLabels(pasivosLabels);
    const patriSum = sumByLabels(patrimonioLabels);

    // Override the totals with their correct calculated sums
    const totAssets = evaluatedRows.find(r => r.label === 'Total Activos');
    if (totAssets) {
        totAssets.vDic = assetsSum.vDic;
        totAssets.vCurr = assetsSum.vCurr;
        totAssets.vPpto = assetsSum.vPpto;
    }

    const totLiabs = evaluatedRows.find(r => r.label === 'Total Pasivos');
    if (totLiabs) {
        totLiabs.vDic = liabsSum.vDic;
        totLiabs.vCurr = liabsSum.vCurr;
        totLiabs.vPpto = liabsSum.vPpto;
    }

    const totPatri = evaluatedRows.find(r => r.label === 'Total Patrimonio');
    // Removed override to keep original value

    const totLiabPatri = evaluatedRows.find(r => r.label === 'Total Pasivo y Patrimonio');
    if (totLiabPatri) {
        totLiabPatri.vDic = (totLiabs ? totLiabs.vDic : 0) + (totPatri ? totPatri.vDic : 0);
        totLiabPatri.vCurr = (totLiabs ? totLiabs.vCurr : 0) + (totPatri ? totPatri.vCurr : 0);
        totLiabPatri.vPpto = (totLiabs ? totLiabs.vPpto : 0) + (totPatri ? totPatri.vPpto : 0);
    }

    bodyEl.innerHTML = evaluatedRows.map(spec => {
        if (spec.type === 'separator') {
            return `<tr><td colspan="4" style="height: 16px; background: transparent; border: none;"></td></tr>`;
        }

        let styleStr = 'padding: 8px 12px; color: var(--text-primary);';
        if (spec.highlight) styleStr += ' font-weight: 700;';
        if (spec.overlayBg) styleStr += ` background-color: ${spec.overlayBg};`;
        if (spec.overlayBg === 'var(--sidebar)') styleStr += ' color: white;';
        
        return `<tr>
            <td style="${styleStr}">${spec.label}</td>
            <td style="text-align: right; ${styleStr}">${formatCurrency(spec.vDic)}</td>
            <td style="text-align: right; ${styleStr}">${formatCurrency(spec.vCurr)}</td>
            <td style="text-align: right; ${styleStr}">${formatCurrency(spec.vPpto)}</td>
        </tr>`;
    }).join('');
}

/**
 * Render the Working Capital Table
 */
function renderWorkingCapital(data, selectedIndex = -1) {
    const headerEl = document.getElementById('wcHeader');
    const bodyEl = document.getElementById('wcBody');
    const periodLabel = document.getElementById('wcPeriodLabel');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    const curr = data[endIdx];
    
    const startIdx = Math.max(0, endIdx - 5);
    // Filtrar para mostrar solo los 6 meses más recientes según lo requerido
    // Similar al P&L, pero asegurando mostrar +6 ultimos meses que es = los 6 ultimos. "deberia ver desde Febrero a Julio" (6 meses).
    // Si queremos que sea estricto para 2026, usamos `filter(d => isYear2026(d))` si el usuario lo requiere.
    let visibleMonths = data.slice(startIdx, endIdx + 1).filter(d => isYear2026(d));
    if (visibleMonths.length === 0) visibleMonths = data.slice(startIdx, endIdx + 1);

    if (periodLabel) {
        periodLabel.textContent = `${visibleMonths[0].date} - ${curr.date}`;
    }

    let headerHTML = '<tr><th>Concepto</th>';
    visibleMonths.forEach(m => {
        headerHTML += `<th>${m.date}</th>`;
    });
    headerHTML += '</tr>';
    headerEl.innerHTML = headerHTML;

    if (curr.wcFullRows && curr.wcFullRows.length > 0) {
        // Render dynamic rows from the Working Capital sheet
        let bodyHTML = '';
        curr.wcFullRows.forEach((row) => {
            if (row.isSpacer) {
                bodyHTML += `<tr style="border-bottom: none; background-color: transparent;"><td colspan="${visibleMonths.length + 1}" style="height: 1.5rem; padding: 0; border: none;"></td></tr>`;
                return;
            }
            
            const conceptName = row.concept;
            const normConcept = conceptName.toLowerCase();
            
            // Skip rows that look like full javascript date strings
            if (normConcept.includes('gmt') || normConcept.includes('00:00:00') || normConcept.includes('hora estándar')) {
                return;
            }

            const isDays = normConcept.includes('dso') || normConcept.includes('dpo') || normConcept.includes('dio') || normConcept.includes('days') || normConcept.includes('dias') || normConcept.includes('días');
            const isPercentage = normConcept.includes('%');
            const isRate = (normConcept.includes('tasa') && !normConcept.includes('impacto')) || normConcept.includes('var ') || normConcept === 'dop' || normConcept === 'eur' || normConcept === 'usd';
            
            bodyHTML += `<tr>
                <td style="font-weight:600; color:var(--text-main);">${conceptName}</td>`;
            
            visibleMonths.forEach(m => {
                const r = m.wcFullRows?.find(r => r.concept === conceptName);
                let val = r ? r.values[m.date] || 0 : 0;
                
                if (row.isHeader) val = "";
                
                if (val === "") {
                    bodyHTML += `<td></td>`;
                } else if (isDays) {
                    bodyHTML += `<td style="color: ${val < 0 ? 'var(--danger)' : 'inherit'};">${val.toFixed(1)}</td>`;
                } else if (isPercentage) {
                    bodyHTML += `<td style="color: ${val < 0 ? 'var(--danger)' : 'inherit'};">${(val * 100).toFixed(1)}%</td>`;
                } else if (isRate) {
                    bodyHTML += `<td style="color: ${val < 0 ? 'var(--danger)' : 'inherit'};">${val.toFixed(2)}</td>`;
                } else {
                    bodyHTML += `<td style="color: ${val < 0 ? 'var(--danger)' : 'inherit'};">${formatCurrency(val)}</td>`;
                }
            });
            bodyHTML += `</tr>`;
        });
        bodyEl.innerHTML = bodyHTML;
        return;
    }

    const wcConcepts = [
        { key: 'wc', label: 'Working Capital', isBold: true },
        // ROW REMOVED as requested by user
        { key: 'dpo', label: 'DPO (Días de Pago)', isBold: false },
        { key: 'dio', label: 'DIO (Días de Inventario)', isBold: false },
        { key: 'cxc', label: 'Cuentas por Cobrar', isBold: false },
        { key: 'inv', label: 'Inventario', isBold: false },
        { key: 'cxp', label: 'Cuentas por Pagar', isBold: false }
    ];

    let bodyHTML = '';
    wcConcepts.forEach(c => {
        bodyHTML += `<tr>
            <td style="${c.isBold ? 'font-weight:700; color:var(--text-main);' : 'padding-left:1.5rem; color:var(--text-secondary);'}">${c.label}</td>`;
        visibleMonths.forEach(m => {
            let val = m.wcDetail?.[c.key];
            if (val === undefined) { val = m.cashflowDetail?.[c.key]; }
            // If the value doesn't strictly exist mapped in cashflowDetail, try to search it
            if (val === undefined && m.pnl && m.pnl.fullRows) {
                 const localRow = m.pnl.fullRows.find(r => r.concept.toLowerCase().includes(c.label.toLowerCase()) || r.concept.toLowerCase() === c.key.toLowerCase());
                 if (localRow && localRow.values && localRow.values[m.date]) {
                     val = localRow.values[m.date];
                 }
            }
            if (val === undefined && m.balance && m.balance.fullRows && c.key === 'cxc') {
                 let sum = 0;
                 m.balance.fullRows.forEach(r => {
                     if (r.concept.toLowerCase().includes("cobrar")) {
                         if (r.values && r.values[m.date] !== undefined) {
                             sum += r.values[m.date] || 0;
                         }
                     }
                 });
                 val = sum;
            }
            if (val === undefined && m.balance && m.balance.fullRows && c.key === 'cxp') {
                 let sum = 0;
                 m.balance.fullRows.forEach(r => {
                     if (r.concept.toLowerCase().includes("pagar")) {
                         if (r.values && r.values[m.date] !== undefined) {
                             sum += r.values[m.date] || 0;
                         }
                     }
                 });
                 val = sum;
            }
            if (val === undefined && m.balance && m.balance.fullRows && c.key === 'inv') {
                 let sum = 0;
                 m.balance.fullRows.forEach(r => {
                     if (r.concept.toLowerCase().includes("inventario")) {
                         if (r.values && r.values[m.date] !== undefined) {
                             sum += r.values[m.date] || 0;
                         }
                     }
                 });
                 val = sum;
            }
            if (val === undefined) { val = 0; }
            
            if (['dso', 'dpo', 'dio'].includes(c.key)) {
                bodyHTML += `<td>${val.toFixed(1)}</td>`;
            } else {
                bodyHTML += `<td class="${val < 0 ? 'negative-val' : ''}">${formatCurrency(val)}</td>`;
            }
        });
        bodyHTML += `</tr>`;
    });

    bodyEl.innerHTML = bodyHTML;
}

window.renderCxpView = function(overrideData, globalIdx = -1) {
    let baseData = overrideData;
    if (Array.isArray(baseData) || !baseData || !baseData.labels) {
        baseData = window.cxpStandaloneData;
    }
    const headerEl = document.getElementById('cxpHeader');
    const bodyEl = document.getElementById('cxpBody');
    const periodLabel = document.getElementById('cxpPeriodLabel');
    if (!headerEl || !bodyEl) return;
    if (!baseData || !baseData.labels) {
        bodyEl.innerHTML = '<tr><td colspan="6" style="text-align:center;' +
            'color:var(--text-secondary);padding:24px;">' +
            'Carga el archivo CXP_Historico.xlsx para visualizar este módulo.' +
            '</td></tr>';
        return;
    }

    // Determine target month from globalFinancialData
    const ms = document.getElementById('monthSelector');
    let useIdx = globalIdx !== -1 ? globalIdx : (ms ? parseInt(ms.value, 10) : -1);
    if (isNaN(useIdx)) useIdx = -1;
    if (useIdx === -1 && globalFinancialData) {
        useIdx = globalFinancialData.length - 1;
    }

    let endSliceIdx = baseData.labels.length - 1;
    let foundIdx = -1;
    if (useIdx !== -1 && globalFinancialData && globalFinancialData[useIdx]) {
        if (globalFinancialData[useIdx].sortDate) {
            const gDate = new Date(globalFinancialData[useIdx].sortDate);
            const m = gDate.getMonth() + 1;
            const y = gDate.getFullYear();
            
            if (baseData.periods) {
                for (let i = baseData.periods.length - 1; i >= 0; i--) {
                    const parts = String(baseData.periods[i]).split('/');
                    const pM = parseInt(parts[0], 10);
                    const pY = parseInt(parts[1], 10);
                    if (pY < y || (pY === y && pM <= m)) {
                        foundIdx = i;
                        break;
                    }
                }
            }
        }
        
        if (foundIdx !== -1) {
            endSliceIdx = foundIdx;
        }
    }

    if (globalFinancialData && baseData.periods && baseData.periods.length > 0) {
        baseData.CostosYTD = baseData.periods.map((p, i) => {
            const parts = String(p).split('/');
            const m = parseInt(parts[0], 10);
            const y = parseInt(parts[1], 10);
            
            const gItem = globalFinancialData.find(d => {
                if (d.sortDate) {
                    const dt = new Date(d.sortDate);
                    return dt.getMonth() + 1 === m && dt.getFullYear() === y;
                }
                return false;
            });
            
            if (gItem && gItem.wcFullRows) {
                const r = gItem.wcFullRows.find(rw => {
                    const c = String(rw.concept).toLowerCase().trim();
                    return c.includes("total costos") || c.includes("costos ytd") || c.includes("opex + capex") || (c.includes("capex") && c.includes("opex") && c.includes("costo"));
                });
                if (r && r.values) {
                    // Try exact match
                    if (r.values[gItem.date] !== undefined && r.values[gItem.date] !== null && r.values[gItem.date] !== "") return r.values[gItem.date];
                    
                    // Fallback to fuzzy match keys in r.values
                    for (let key in r.values) {
                        if (String(key).toLowerCase().trim() === String(gItem.date).toLowerCase().trim()) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                    
                    // If all fails, fall back to matching month and year string in key
                    const shortMonths = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
                    const tgtStr = shortMonths[m-1] + "-" + String(y).slice(-2);
                    for (let key in r.values) {
                        const kl = String(key).toLowerCase();
                        if (kl.includes(shortMonths[m-1]) && kl.includes(String(y).slice(-2))) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                }
            }
            return baseData.CostosYTD[i]; // fallback to existing
        });
        
        baseData.DPO = baseData.periods.map((p, i) => {
            const parts = String(p).split('/');
            const m = parseInt(parts[0], 10);
            const y = parseInt(parts[1], 10);
            
            const gItem = globalFinancialData.find(d => {
                if (d.sortDate) {
                    const dt = new Date(d.sortDate);
                    return dt.getMonth() + 1 === m && dt.getFullYear() === y;
                }
                return false;
            });
            
            if (gItem && gItem.wcFullRows) {
                const r = gItem.wcFullRows.find(rw => {
                    const c = String(rw.concept).toLowerCase().trim();
                    return c === "dpo" || c.includes("dpo") || c.includes("días de cuentas") || c.includes("dias de cuentas");
                });
                if (r && r.values) {
                    if (r.values[gItem.date] !== undefined && r.values[gItem.date] !== null && r.values[gItem.date] !== "") return r.values[gItem.date];
                    
                    for (let key in r.values) {
                         if (String(key).toLowerCase().trim() === String(gItem.date).toLowerCase().trim()) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                    
                    const shortMonths = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
                    for (let key in r.values) {
                        const kl = String(key).toLowerCase();
                        if (kl.includes(shortMonths[m-1]) && kl.includes(String(y).slice(-2))) {
                            if (Math.abs(Number(r.values[key])) > 0) return r.values[key];
                        }
                    }
                }
            }
            return baseData.DPO[i]; // fallback to existing
        });
    }

    function matchFinAndCxpPeriod(gItem, pStr) {
        if (!gItem || !pStr) return false;
        const parts = String(pStr).split('/');
        if (parts.length < 2) return false;
        const pM = parseInt(parts[0], 10);
        let pY = parseInt(parts[1], 10);
        if (pY < 100) pY += 2000;
        
        const gSortDate = gItem.sortDate;
        if (gSortDate) {
            const d = new Date(gSortDate);
            if (!isNaN(d)) {
                return d.getMonth() + 1 === pM && d.getFullYear() === pY;
            }
        }
        
        // Fallback: compare month and year using isYear2025/isYear2026 and month name
        const normDate = (gItem.date || "").toLowerCase();
        const shortMonths = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
        const pMonthShort = shortMonths[pM - 1];
        const isTargetYear = pY === 2025 ? isYear2025(gItem) : (pY === 2026 ? isYear2026(gItem) : false);
        
        return isTargetYear && (normDate.includes(pMonthShort) || normDate.includes(String(pM)));
    }

    let labels = [];
    let periods = [];
    let corriente = [];
    let balanceGeneral = [];
    let cxp = [];
    let provisionales = [];
    let otrosProveedores = [];
    let total = [];
    let costosYTD = [];
    let dpo = [];
    let aging = {
        "0_30": [],
        "31_60": [],
        "61_90": [],
        "91_120": [],
        "121_150": [],
        "151_180": [],
        "180Mas": []
    };
    let top14Saldos = {};
    if (baseData.Top14Names) {
        baseData.Top14Names.forEach(k => {
            top14Saldos[k] = [];
        });
    }

    if (globalFinancialData && globalFinancialData.length > 0) {
        const endIdx = useIdx >= 0 && useIdx < globalFinancialData.length ? useIdx : globalFinancialData.length - 1;
        const startIdx = Math.max(0, endIdx - 5);
        let visibleFinancialMonths = globalFinancialData.slice(startIdx, endIdx + 1);
        visibleFinancialMonths = visibleFinancialMonths.filter(m => isYear2026(m));
        const dic2025Fin = globalFinancialData.find(d => isYear2025(d) && (d.date.toLowerCase().includes('dic') || d.date.toLowerCase().includes('dec')));
        if (dic2025Fin && !visibleFinancialMonths.includes(dic2025Fin)) {
            visibleFinancialMonths.unshift(dic2025Fin);
        }

        visibleFinancialMonths.forEach(gItem => {
            labels.push(gItem.date ? gItem.date.toUpperCase() : "");
            
            // Find corresponding index in baseData.periods
            const cxpIdx = baseData.periods ? baseData.periods.findIndex(p => matchFinAndCxpPeriod(gItem, p)) : -1;
            
            if (cxpIdx !== -1) {
                periods.push(baseData.periods[cxpIdx]);
                corriente.push(baseData.Corriente[cxpIdx] || 0);
                balanceGeneral.push(baseData.BalanceGeneral[cxpIdx] || 0);
                cxp.push(baseData.CXP[cxpIdx] || 0);
                provisionales.push(baseData.Provisionales[cxpIdx] || 0);
                otrosProveedores.push((baseData.OtrosProveedores || [])[cxpIdx] || 0);
                total.push((baseData.Total || [])[cxpIdx] || 0);
                costosYTD.push((baseData.CostosYTD || [])[cxpIdx] || 0);
                dpo.push((baseData.DPO || [])[cxpIdx] || 0);
                
                aging["0_30"].push(baseData.Aging["0_30"][cxpIdx] || 0);
                aging["31_60"].push(baseData.Aging["31_60"][cxpIdx] || 0);
                aging["61_90"].push(baseData.Aging["61_90"][cxpIdx] || 0);
                aging["91_120"].push(baseData.Aging["91_120"][cxpIdx] || 0);
                aging["121_150"].push(baseData.Aging["121_150"][cxpIdx] || 0);
                aging["151_180"].push(baseData.Aging["151_180"][cxpIdx] || 0);
                aging["180Mas"].push(baseData.Aging["180Mas"][cxpIdx] || 0);
                
                if (baseData.Top14Names) {
                    baseData.Top14Names.forEach(k => {
                        const arr = baseData.Top14Saldos ? baseData.Top14Saldos[k] : null;
                        top14Saldos[k].push(arr ? (arr[cxpIdx] || 0) : 0);
                    });
                }
            } else {
                // Future month or non-existent in CXP files -> fill with 0
                let m = 1;
                let y = 2026;
                if (gItem.sortDate) {
                    const dt = new Date(gItem.sortDate);
                    m = dt.getMonth() + 1;
                    y = dt.getFullYear();
                } else {
                    const norm = (gItem.date || "").toLowerCase();
                    const shortMonths = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
                    const foundMonthIdx = shortMonths.findIndex(sm => norm.includes(sm));
                    if (foundMonthIdx !== -1) {
                         m = foundMonthIdx + 1;
                    }
                    if (isYear2025(gItem)) y = 2025;
                }
                const pStr = `${m}/${y}`;
                periods.push(pStr);
                corriente.push(0);
                balanceGeneral.push(0);
                cxp.push(0);
                provisionales.push(0);
                otrosProveedores.push(0);
                total.push(0);
                costosYTD.push(0);
                dpo.push(0);
                
                aging["0_30"].push(0);
                aging["31_60"].push(0);
                aging["61_90"].push(0);
                aging["91_120"].push(0);
                aging["121_150"].push(0);
                aging["151_180"].push(0);
                aging["180Mas"].push(0);
                
                if (baseData.Top14Names) {
                    baseData.Top14Names.forEach(k => {
                        top14Saldos[k].push(0);
                    });
                }
            }
        });
    } else {
        // Fallback to original slicing if globalFinancialData is empty
        const startSliceIdx = Math.max(0, endSliceIdx - 5);
        let visibleIndices = [];
        for (let i = startSliceIdx; i <= endSliceIdx; i++) {
            visibleIndices.push(i);
        }
        visibleIndices = visibleIndices.filter(idx => {
            const pStr = baseData.periods[idx] || "";
            return pStr.endsWith('/2026') || pStr.endsWith('/26');
        });
        const dic2025Idx = baseData.periods.findIndex(p => {
            const pStr = String(p);
            return pStr === "12/2025" || pStr === "12/25";
        });
        if (dic2025Idx !== -1 && !visibleIndices.includes(dic2025Idx)) {
            visibleIndices.unshift(dic2025Idx);
        }
        
        visibleIndices.forEach(idx => {
            labels.push((baseData.labels[idx] || "").toUpperCase());
            periods.push(baseData.periods[idx]);
            corriente.push(baseData.Corriente[idx] || 0);
            balanceGeneral.push(baseData.BalanceGeneral[idx] || 0);
            cxp.push(baseData.CXP[idx] || 0);
            provisionales.push(baseData.Provisionales[idx] || 0);
            otrosProveedores.push((baseData.OtrosProveedores || [])[idx] || 0);
            total.push((baseData.Total || [])[idx] || 0);
            costosYTD.push((baseData.CostosYTD || [])[idx] || 0);
            dpo.push((baseData.DPO || [])[idx] || 0);
            
            aging["0_30"].push(baseData.Aging["0_30"][idx] || 0);
            aging["31_60"].push(baseData.Aging["31_60"][idx] || 0);
            aging["61_90"].push(baseData.Aging["61_90"][idx] || 0);
            aging["91_120"].push(baseData.Aging["91_120"][idx] || 0);
            aging["121_150"].push(baseData.Aging["121_150"][idx] || 0);
            aging["151_180"].push(baseData.Aging["151_180"][idx] || 0);
            aging["180Mas"].push(baseData.Aging["180Mas"][idx] || 0);
            
            if (baseData.Top14Names) {
                baseData.Top14Names.forEach(k => {
                    const arr = baseData.Top14Saldos ? baseData.Top14Saldos[k] : null;
                    top14Saldos[k].push(arr ? (arr[idx] || 0) : 0);
                });
            }
        });
    }

    const data = {
        labels,
        periods,
        Corriente: corriente,
        Aging: aging,
        BalanceGeneral: balanceGeneral,
        CXP: cxp,
        Provisionales: provisionales,
        OtrosProveedores: otrosProveedores,
        Top14Names: baseData.Top14Names || [],
        Top14Saldos: top14Saldos,
        Total: total,
        CostosYTD: costosYTD,
        DPO: dpo
    };

    if (periodLabel) {
        let yr = "20XX";
        if (data.periods && data.periods.length > 0) {
            let lastP = data.periods[data.periods.length - 1];
            yr = String(lastP).split('/')[1] || "20XX";
        }
        periodLabel.innerText = `Análisis de Cuentas Por Pagar ${yr}`;
    }

    let headerHTML = '<tr><th style="white-space: normal; word-wrap: break-word; width: 250px;">Concepto</th>';
    data.labels.forEach((lbl) => {
        headerHTML += `<th style="text-align:right;">${lbl}</th>`;
    });
    headerHTML += '</tr>';
    headerEl.innerHTML = headerHTML;
    
    const formatCurrencyStr = (v, minDec = 2, maxDec = 2) => {
        if (v === undefined || v === null) return '-';
        if (typeof v !== 'number') return v;
        return v.toLocaleString('en-US', { minimumFractionDigits: minDec, maximumFractionDigits: maxDec });
    };

    const formatInt = (v) => v !== undefined && v !== null ? Math.round(v).toLocaleString('en-US') : '-';

    let bodyHTML = '';

    const addRow = (label, values, isTotal = false) => {
        const rowStyle = isTotal ? 'font-weight:700; background:rgba(0,0,0,0.02);' : '';
        let h = `<tr style="${rowStyle}">` 
            + `<td style="white-space: normal !important; word-break: break-word;">${label}</td>`;
        
        values.forEach((v) => {
            const valNum = parseFloat(v);
            const isNegative = !isNaN(valNum) && valNum < -0.009;
            const textCls = isNegative ? 'negative-val' : '';
            h += `<td style="text-align:right;" class="${textCls}">${formatCurrencyStr(v)}</td>`;
        });
        h += '</tr>';
        return h;
    };

    // Resumen General
    bodyHTML += addRow('Balance General', data.BalanceGeneral, true);
    
    // Aging
    bodyHTML += `<tr><td colspan="${data.labels.length + 1}" style="font-weight:700; background:rgba(0,0,0,0.04); font-size: 0.85rem; text-transform: uppercase;">Aging</td></tr>`;
    bodyHTML += addRow('CXP', data.CXP);
    bodyHTML += addRow('Proveedores provisión sin fact', data.Provisionales);
    bodyHTML += addRow('Corriente (Saldo No Vencido)', data.Corriente);
    bodyHTML += addRow('0 a 30', data.Aging['0_30']);
    bodyHTML += addRow('31 a 60', data.Aging['31_60']);
    bodyHTML += addRow('61 a 90', data.Aging['61_90']);
    bodyHTML += addRow('91 a 120', data.Aging['91_120']);
    bodyHTML += addRow('121 a 150', data.Aging['121_150']);
    bodyHTML += addRow('151 a 180', data.Aging['151_180']);
    bodyHTML += addRow('> 180', data.Aging['180Mas']);

    // Proveedores Top 14
    bodyHTML += `<tr style="height:20px"><td colspan="${data.labels.length + 1}"></td></tr>`;
    
    let provLine = '<tr><td style="font-weight:700; background:rgba(0,0,0,0.04); font-size: 0.85rem; text-transform: uppercase;">Saldos de Top Proveedores</td>';
    data.labels.forEach((lbl) => provLine += `<td style="text-align:right; font-weight:700; background:rgba(0,0,0,0.04);">${lbl}</td>`);
    provLine += '</tr>';
    bodyHTML += provLine;

    data.Top14Names.forEach((name) => {
        bodyHTML += addRow(name, data.Top14Saldos[name] || [0,0,0,0,0]);
    });

    bodyHTML += addRow('Otros Proveedores', data.OtrosProveedores);
    bodyHTML += addRow('Total', data.Total, true);
    
    bodyHTML += `<tr style="height:20px"><td colspan="${data.labels.length + 1}"></td></tr>`;
    
    // Costos YTD
    let costosRow = '<tr><td>Costos + Gasto (Opex+Capex) YTD</td>';
    data.CostosYTD.forEach((v) => {
        // formatCostos can be up to thousands, we can use 2 decimals
        costosRow += `<td style="text-align:right;">${formatCurrencyStr(v)}</td>`
    });
    costosRow += '</tr>';
    bodyHTML += costosRow;
    
    // DPO
    let dpoRow = '<tr><td>DPO</td>';
    data.DPO.forEach((v) => {
        dpoRow += `<td style="text-align:right;">${formatInt(v)}</td>`
    });
    dpoRow += '</tr>';
    bodyHTML += dpoRow;

    bodyEl.innerHTML = bodyHTML;

    // Render Summary view if active or updated
    if (typeof window.renderCxpResumen === 'function') {
        window.renderCxpResumen(data, data.labels.length - 1);
    }
}

window.renderCxpResumen = function(data, selectedIdx = -1) {
    const resumenBody = document.getElementById('cxpResumenBody');
    const chartContainer = document.getElementById('cxpResumenChart');
    if (!resumenBody || !chartContainer) return;

    if (!data || !data.labels || data.labels.length === 0) {
        resumenBody.innerHTML = '<tr><td colspan="3" style="text-align: center; padding: 24px; color: var(--text-secondary); font-style: italic;">Carga el archivo CXP_Historico.xlsx para visualizar el resumen.</td></tr>';
        chartContainer.innerHTML = '';
        return;
    }

    let idxLast = selectedIdx !== -1 ? selectedIdx : data.labels.length - 1;
    if (idxLast < 0) idxLast = 0;
    if (idxLast >= data.labels.length) idxLast = data.labels.length - 1;

    let idxDec = -1;
    for (let i = idxLast; i >= 0; i--) {
        const lbl = String(data.labels[i]).toLowerCase();
        if ((lbl.includes('dic') || lbl.includes('dec')) && i !== idxLast) {
            idxDec = i;
            break;
        }
    }
    if (idxDec === -1) idxDec = 0;

    const lblDec = data.labels[idxDec] || 'dic 2025';
    const lblLast = data.labels[idxLast] || 'abr 2026';

    // Set headers
    const headerCol1Select = document.getElementById('cxpResumenCol1');
    const headerCol2Select = document.getElementById('cxpResumenCol2');
    if (headerCol1Select) headerCol1Select.textContent = lblDec.toLowerCase();
    if (headerCol2Select) headerCol2Select.textContent = lblLast.toLowerCase();

    // Helper to format values as integers, empty if undefined or null
    const formatValInt = (v) => {
        if (v === undefined || v === null) return '-';
        return Math.round(v).toLocaleString('en-US');
    };

    let tableHTML = '';
    
    // Top 5 suppliers
    const top5Names = (data.Top14Names || []).slice(0, 5);
    let top5SumDec = 0;
    let top5SumLast = 0;

    top5Names.forEach(name => {
        const rowSaldos = data.Top14Saldos[name] || [];
        const valDec = rowSaldos[idxDec] || 0;
        const valLast = rowSaldos[idxLast] || 0;

        top5SumDec += valDec;
        top5SumLast += valLast;

        tableHTML += `
            <tr>
              <td style="text-align: left; white-space: normal !important; word-break: break-word;">${name}</td>
              <td style="text-align: right;">${formatValInt(valDec)}</td>
              <td style="text-align: right;">${formatValInt(valLast)}</td>
            </tr>
        `;
    });

    // Otros Proveedores = Total - top5Sum
    const totalDec = data.Total[idxDec] || 0;
    const totalLast = data.Total[idxLast] || 0;

    const otrosDec = Math.max(0, totalDec - top5SumDec);
    const otrosLast = Math.max(0, totalLast - top5SumLast);

    tableHTML += `
        <tr>
          <td style="text-align: left;">Otros Proveedores</td>
          <td style="text-align: right;">${formatValInt(otrosDec)}</td>
          <td style="text-align: right;">${formatValInt(otrosLast)}</td>
        </tr>
    `;

    // Total Row
    tableHTML += `
        <tr class="total-row" style="background-color: var(--row-hover);">
          <td style="text-align: left;"><strong>Total</strong></td>
          <td style="text-align: right;"><strong>${formatValInt(totalDec)}</strong></td>
          <td style="text-align: right;"><strong>${formatValInt(totalLast)}</strong></td>
        </tr>
    `;

    // DPO Row
    const dpoDec = (data.DPO || [])[idxDec] || 0;
    const dpoLast = (data.DPO || [])[idxLast] || 0;

    tableHTML += `
        <tr class="total-row" style="background-color: var(--row-hover);">
          <td style="text-align: left;"><strong>DPO</strong></td>
          <td style="text-align: right;"><strong>${formatValInt(dpoDec)}</strong></td>
          <td style="text-align: right;"><strong>${formatValInt(dpoLast)}</strong></td>
        </tr>
    `;

    resumenBody.innerHTML = tableHTML;

    // --- CHART GENERATION ---
    const corrienteDec = data.Corriente[idxDec] || 0;
    const agingFields = ['0_30', '31_60', '61_90', '91_120', '121_150', '151_180', '180Mas'];
    const vencidoDec = agingFields.reduce((sum, f) => sum + (data.Aging[f][idxDec] || 0), 0);

    const corrienteLast = data.Corriente[idxLast] || 0;
    const vencidoLast = agingFields.reduce((sum, f) => sum + (data.Aging[f][idxLast] || 0), 0);

    const baseDec = (corrienteDec + vencidoDec) || 1;
    const baseLast = (corrienteLast + vencidoLast) || 1;

    const pctCorrienteDec = Math.round((corrienteDec / baseDec) * 100);
    const pctVencidoDec = Math.round((vencidoDec / baseDec) * 100);

    const pctCorrienteLast = Math.round((corrienteLast / baseLast) * 100);
    const pctVencidoLast = Math.round((vencidoLast / baseLast) * 100);

    const maxVal = Math.max(baseDec, baseLast, 100) * 1.12; 
    const svgHeight = 460;
    const drawHeight = 350; 
    const baseY = 410; 
    const pixelsPerUnit = drawHeight / maxVal;

    const hCorrienteDec = corrienteDec * pixelsPerUnit;
    const hVencidoDec = vencidoDec * pixelsPerUnit;

    const hCorrienteLast = corrienteLast * pixelsPerUnit;
    const hVencidoLast = vencidoLast * pixelsPerUnit;

    const midYCorrienteDec = baseY - (hCorrienteDec / 2);
    const midYVencidoDec = baseY - hCorrienteDec - (hVencidoDec / 2);

    const midYCorrienteLast = baseY - (hCorrienteLast / 2);
    const midYVencidoLast = baseY - hCorrienteLast - (hVencidoLast / 2);

    const barW = 120;
    const xDec = 140;
    const xLast = 420;

    const codeSVG = `
      <svg width="100%" height="100%" viewBox="0 0 580 460" style="font-family: inherit;">
        <!-- Base line -->
        <line x1="30" y1="${baseY}" x2="550" y2="${baseY}" stroke="#cbd5e1" stroke-width="2" />

        <!-- ================= DIC BAR ================= -->
        <!-- Corriente (Blue) -->
        <rect x="${xDec - barW/2}" y="${baseY - hCorrienteDec}" width="${barW}" height="${hCorrienteDec}" fill="#0070c0" rx="2" />
        <!-- Vencido (Red) -->
        <rect x="${xDec - barW/2}" y="${baseY - hCorrienteDec - hVencidoDec}" width="${barW}" height="${hVencidoDec}" fill="#b91c1c" rx="2" />

        <!-- Labels inside Dec Bar -->
        ${hCorrienteDec > 16 ? `
        <text x="${xDec}" y="${midYCorrienteDec + 5}" fill="white" font-size="12" font-weight="bold" text-anchor="middle">${Math.round(corrienteDec)}</text>
        ` : ''}
        ${hVencidoDec > 16 ? `
        <text x="${xDec}" y="${midYVencidoDec + 5}" fill="white" font-size="12" font-weight="bold" text-anchor="middle">${Math.round(vencidoDec)}</text>
        ` : ''}

        <!-- Total on top of Dec Bar -->
        <text x="${xDec}" y="${baseY - hCorrienteDec - hVencidoDec - 10}" fill="#1e293b" font-size="13" font-weight="bold" text-anchor="middle">${Math.round(totalDec)}</text>

        <!-- Label below Dec Bar -->
        <text x="${xDec}" y="${baseY + 22}" fill="#475569" font-size="11" font-weight="bold" text-anchor="middle">${lblDec.toUpperCase()}</text>

        <!-- Annotations Dec Bar -->
        <text x="${xDec + barW/2 + 10}" y="${midYVencidoDec}" fill="#b91c1c" font-size="12" font-weight="bold" text-anchor="start">${pctVencidoDec}% Vencido</text>
        <text x="${xDec + barW/2 + 10}" y="${midYCorrienteDec}" fill="#0070c0" font-size="12" font-weight="bold" text-anchor="start">${pctCorrienteDec}% Corriente</text>


        <!-- ================= LAST BAR ================= -->
        <!-- Corriente (Blue) -->
        <rect x="${xLast - barW/2}" y="${baseY - hCorrienteLast}" width="${barW}" height="${hCorrienteLast}" fill="#0070c0" rx="2" />
        <!-- Vencido (Red) -->
        <rect x="${xLast - barW/2}" y="${baseY - hCorrienteLast - hVencidoLast}" width="${barW}" height="${hVencidoLast}" fill="#b91c1c" rx="2" />

        <!-- Labels inside Last Bar -->
        ${hCorrienteLast > 16 ? `
        <text x="${xLast}" y="${midYCorrienteLast + 5}" fill="white" font-size="12" font-weight="bold" text-anchor="middle">${Math.round(corrienteLast)}</text>
        ` : ''}
        ${hVencidoLast > 16 ? `
        <text x="${xLast}" y="${midYVencidoLast + 5}" fill="white" font-size="12" font-weight="bold" text-anchor="middle">${Math.round(vencidoLast)}</text>
        ` : ''}

        <!-- Total on top of Last Bar -->
        <text x="${xLast}" y="${baseY - hCorrienteLast - hVencidoLast - 10}" fill="#1e293b" font-size="13" font-weight="bold" text-anchor="middle">${Math.round(totalLast)}</text>

        <!-- Label below Last Bar -->
        <text x="${xLast}" y="${baseY + 22}" fill="#475569" font-size="11" font-weight="bold" text-anchor="middle">${lblLast.toUpperCase()}</text>

        <!-- Annotations Last Bar -->
        <text x="${xLast + barW/2 + 10}" y="${midYVencidoLast}" fill="#b91c1c" font-size="12" font-weight="bold" text-anchor="start">${pctVencidoLast}% Vencido</text>
        <text x="${xLast + barW/2 + 10}" y="${midYCorrienteLast}" fill="#0070c0" font-size="12" font-weight="bold" text-anchor="start">${pctCorrienteLast}% Corriente</text>
      </svg>
    `;

    chartContainer.innerHTML = codeSVG;
};

// Ensure the old function name works due to previous references if any
function renderCxpView(overrideData, globalIdx = -1) {
    if (window.renderCxpView) {
        window.renderCxpView(overrideData, globalIdx);
    }
}



function renderDeudaView(data, selectedIndex = -1) {
    if (!data || data.length === 0) return;
    
    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    const curr = data[endIdx];
    
    // Buscar Dic-25 para los saldos iniciales (si existe en los datos)
    const dic2025 = data.find(d => d.sortDate && getSortYear(d) === 2025 && d.sortDate.getMonth() === 11) || curr;
    
    const pLabel = document.getElementById("deudaPeriodLabel");
    if (pLabel) pLabel.textContent = `Millones DOP | ${curr.date}`;

    const ext = (obj) => obj && obj.deudaMetrics ? obj.deudaMetrics : {};
    
    const currD = ext(curr);
    const dicD = ext(dic2025);
    
    const currDetail = currD.debtDetail || { bancos: {}, tasasPorBanco: {} };
    const dicDetail = dicD.debtDetail || { bancos: {}, tasasPorBanco: {} };

    const formatPercent = (val) => val === null || val === undefined ? "N/A" : (val * 100).toFixed(2) + "%";
    const formatRatio = (val) => val === null || val === undefined ? "N/A" : val.toFixed(2) + "x";
    const formatNum = (val, decimals=0) => val === null || val === undefined ? "N/A" : (decimals > 0 ? formatCurrency(val).replace('$', '') : formatCurrency(val).replace('$', '').split('.')[0]);
    const formatFx = (val) => val === null || val === undefined || val === 0 ? "N/A" : val.toFixed(2);

    // Tabla Bancos
    const banksHeader = document.getElementById('deudaBancosHeader');
    const banksBody = document.getElementById('deudaBancosBody');
    if (banksHeader && banksBody) {
        banksHeader.innerHTML = `
            <tr>
                <th style="background: var(--sidebar); color: white; text-align: left; padding: 12px 16px;">Banco</th>
                <th style="background: var(--sidebar); color: white; text-align: right; padding: 12px 16px;">${dic2025.date}</th>
                <th style="background: var(--sidebar); color: white; text-align: right; padding: 12px 16px;">${curr.date}</th>
                <th style="background: var(--sidebar); color: white; text-align: right; padding: 12px 16px;">Tasa Promedio</th>
            </tr>
        `;
        
        const bancoList = [
            { id: 'Banco Popular', label: 'Banco Popular' },
            { id: 'Banco Santa Cruz', label: 'Banco Santa Cruz' },
            { id: 'Scotiabank', label: 'Scotiabank' },
            { id: 'Loganville', label: 'Loganville' }
        ];
        
        let banksHTML = '';
        let totDic = 0, totCurr = 0;
        
        bancoList.forEach(b => {
            const vDic = dicDetail.bancos[b.id] || 0;
            const vCurr = currDetail.bancos[b.id] || 0;
            const tCurr = currDetail.tasasPorBanco[b.id];
            
            totDic += vDic;
            totCurr += vCurr;
            
            banksHTML += `
                <tr>
                    <td style="font-weight: 600; text-align: left; padding: 12px 16px;">${b.label}</td>
                    <td class="num" style="text-align: right; padding: 12px 16px;">${vDic === 0 ? "" : formatNum(vDic)}</td>
                    <td class="num" style="text-align: right; padding: 12px 16px;">${vCurr === 0 ? "" : formatNum(vCurr)}</td>
                    <td class="num" style="text-align: right; padding: 12px 16px;">${tCurr === undefined || tCurr === null || tCurr === 0 ? "" : formatPercent(tCurr)}</td>
                </tr>
            `;
        });
        
        banksHTML += `
            <tr style="background: var(--sidebar); color: white; font-weight: bold;">
                <td style="font-weight: bold; text-align: left; padding: 12px 16px;">Total</td>
                <td class="num" style="font-weight: bold; text-align: right; padding: 12px 16px;">${totDic === 0 ? "" : formatNum(totDic)}</td>
                <td class="num" style="font-weight: bold; text-align: right; padding: 12px 16px;">${totCurr === 0 ? "" : formatNum(totCurr)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;"></td>
            </tr>
        `;
        banksBody.innerHTML = banksHTML;
    }

    // Tabla Indicadores
    const indHeader = document.getElementById('deudaIndicadoresHeader');
    const indBody = document.getElementById('deudaIndicadoresBody');
    if (indHeader && indBody) {
        indHeader.innerHTML = `
            <tr>
                <th style="background: var(--sidebar); color: white; text-align: left; padding: 12px 16px;">Indicador</th>
                <th style="background: var(--sidebar); color: white; text-align: right; padding: 12px 16px;">${dic2025.date}</th>
                <th style="background: var(--sidebar); color: white; text-align: right; padding: 12px 16px;">${curr.date}</th>
            </tr>
        `;
        
        const highlightStyle = (valStr, isBad) => {
            if (valStr === "N/A" || valStr === "") return valStr;
            if (isBad) return `<span style="color: #ef4444; font-weight: bold;">${valStr}</span>`;
            return valStr;
        };
        
        const iDic = dicD;
        const iCurr = currD;
        
        indBody.innerHTML = `
            <tr style="background: #e2e8f0; border-bottom: 2px solid #cbd5e1;"><td colspan="3" style="font-weight: 800; color: var(--sidebar); font-size: 0.9rem; padding-top: 12px; padding-bottom: 12px; text-align: left; padding-left: 16px;">Tasa de Interés y Tipo de Cambio</td></tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Tasa DOP</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatPercent(iDic.tasaDop)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatPercent(iCurr.tasaDop)}</td>
            </tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Tasa USD</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${iDic.tasaUsd !== null && iDic.tasaUsd !== undefined && iDic.tasaUsd !== 0 ? formatPercent(iDic.tasaUsd) : formatFx(iDic.tasaCambio)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${iCurr.tasaUsd !== null && iCurr.tasaUsd !== undefined && iCurr.tasaUsd !== 0 ? formatPercent(iCurr.tasaUsd) : formatFx(iCurr.tasaCambio)}</td>
            </tr>
            <tr style="background: #e2e8f0; border-bottom: 2px solid #cbd5e1;"><td colspan="3" style="font-weight: 800; color: var(--sidebar); font-size: 0.9rem; padding-top: 12px; padding-bottom: 12px; text-align: left; padding-left: 16px;">Indicadores</td></tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Deuda Neta USD <span style="font-size:0.75rem; color:var(--text-secondary);">(M USD)</span></td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatNum(iDic.deudaNetaUsd, 1)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatNum(iCurr.deudaNetaUsd, 1)}</td>
            </tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Deuda Neta Bancaria USD <span style="font-size:0.75rem; color:var(--text-secondary);">(M USD)</span></td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatNum(iDic.deudaNetaBancUSD, 1)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatNum(iCurr.deudaNetaBancUSD, 1)}</td>
            </tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Deuda Neta Bancaria / EBITDA R12 (&lt;=4.0x)</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${highlightStyle(formatRatio(iDic.covenantLean), iDic.covenantLean > 4)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${highlightStyle(formatRatio(iCurr.covenantLean), iCurr.covenantLean > 4)}</td>
            </tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Apalancamiento (&lt;=2.0x)</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${highlightStyle(formatRatio(iDic.apalancamiento), iDic.apalancamiento > 2)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${highlightStyle(formatRatio(iCurr.apalancamiento), iCurr.apalancamiento > 2)}</td>
            </tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Capacidad de Pago</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatRatio(iDic.capacidadPago)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${formatRatio(iCurr.capacidadPago)}</td>
            </tr>
            <tr>
                <td style="font-weight: 600; text-align: left; padding: 12px 16px;">Razón Corriente (&gt;=1.5x)</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${highlightStyle(formatRatio(iDic.razonCorriente), iDic.razonCorriente !== null && iDic.razonCorriente < 1.5)}</td>
                <td class="num" style="text-align: right; padding: 12px 16px;">${highlightStyle(formatRatio(iCurr.razonCorriente), iCurr.razonCorriente !== null && iCurr.razonCorriente < 1.5)}</td>
            </tr>
        `;
    }

    renderDeudaChart(data, endIdx, dic2025);
}

function renderDeudaChart(data, selectedIndex, dic2025) {
    const container = document.getElementById('deuda-chart-container');
    if (!container) return;
    d3.select(container).selectAll("*").remove();

    const curr = data[selectedIndex];
    if (!curr || !dic2025) {
        container.innerHTML = '<p style="color: var(--text-secondary);">[Insuficientes Datos para Gráfico]</p>';
        return;
    }

    const ext = (obj) => obj && obj.deudaMetrics && obj.deudaMetrics.debtDetail ? obj.deudaMetrics.debtDetail.bancos : {};

    const bDic = ext(dic2025);
    const bCurr = ext(curr);
    
    // Dataset para apilados
    const labels = [dic2025.date, curr.date];
    const bancos = [
        { key: 'Banco Popular', color: '#0040C1', label: 'Banco Popular', logo: 'popular.png' },
        { key: 'Banco Santa Cruz', color: '#00B050', label: 'Banco Santa Cruz', logo: 'santacruz.png' },
        { key: 'Scotiabank', color: '#FF0000', label: 'Scotiabank', logo: 'scotiabank.png' },
        { key: 'Loganville', color: '#ED7D31', label: 'Loganville', logo: null }
    ];

    const chartData = labels.map((l, i) => {
        let obj = { label: l, total: 0 };
        const dataSrc = i === 0 ? bDic : bCurr;
        let y0 = 0;
        bancos.forEach(b => {
            const val = dataSrc[b.key] || 0;
            obj[b.key] = {
                val: val,
                y0: y0,
                y1: y0 + val
            };
            y0 += val;
            obj.total += val;
        });
        return obj;
    });

    const width = container.clientWidth || 400;
    const height = 350;
    const margin = { top: 40, right: 30, bottom: 40, left: 195 };

    const svgWrapper = d3.select(container).append("div").style("width", "100%").style("display", "flex").style("justify-content", "center");
    
    const svg = svgWrapper.append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand()
        .domain(labels)
        .range([0, width - margin.left - margin.right])
        .padding(0.55);

    const maxY = d3.max(chartData, d => d.total) * 1.1;
    const y = d3.scaleLinear()
        .domain([0, maxY || 1])
        .range([height - margin.top - margin.bottom, 0]);

    // Grid
    svg.append("g")
        .attr("class", "grid")
        .call(d3.axisLeft(y).ticks(6).tickSize(-width + margin.left + margin.right).tickFormat(""))
        .style("stroke-dasharray", "3,3")
        .style("stroke-opacity", 0.1);

    svg.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).tickSize(0).tickPadding(10))
        .selectAll("text")
        .style("font-size", "12px")
        .style("font-weight", "600")
        .style("color", "var(--text-secondary)");

    svg.append("g")
        .call(d3.axisLeft(y).ticks(5).tickFormat(d => (d/1000).toFixed(1) + 'k'))
        .selectAll("text")
        .style("font-size", "11px")
        .style("color", "var(--text-secondary)");

    const yearGroups = svg.selectAll(".yearGroup")
        .data(chartData)
        .enter().append("g")
        .attr("class", "yearGroup")
        .attr("transform", d => `translate(${x(d.label)},0)`);

    bancos.forEach(b => {
        yearGroups.append("rect")
            .attr("x", 0)
            .attr("width", x.bandwidth())
            .attr("y", d => y(d[b.key].y1))
            .attr("height", d => y(d[b.key].y0) - y(d[b.key].y1))
            .style("fill", b.color)
            .style("stroke", "#fff")
            .style("stroke-width", 1);
            
        // Number inside bar
        yearGroups.append("text")
            .filter(d => d[b.key].val > d.total * 0.05) // only if > 5% of total
            .attr("x", x.bandwidth() / 2)
            .attr("y", d => y(d[b.key].y0 + (d[b.key].val / 2)))
            .attr("dy", "0.3em")
            .attr("text-anchor", "middle")
            .style("fill", "#fff")
            .style("font-size", "11px")
            .style("font-weight", "bold")
            .text(d => Math.round(d[b.key].val));
            
        // Logos on the left side (only for the first bar)
        if (chartData[0] && chartData[0][b.key].val > 0) {
            const yCenter = y(chartData[0][b.key].y0 + (chartData[0][b.key].val / 2));
            if (b.logo) {
                svg.append("image")
                    .attr("href", b.logo)
                    .attr("x", -175) // to the left of the y-axis
                    .attr("y", yCenter - 20)
                    .attr("width", 135)
                    .attr("height", 40)
                    .attr("preserveAspectRatio", "xMaxYMid meet");
            } else {
                svg.append("text")
                    .attr("x", -75)
                    .attr("y", yCenter)
                    .attr("dy", "0.35em")
                    .attr("text-anchor", "end")
                    .style("font-weight", "bold")
                    .style("font-size", "14px")
                    .style("fill", "var(--sidebar)")
                    .text(b.label);
            }
        }
    });

    // Totales top
    yearGroups.append("text")
        .attr("x", x.bandwidth() / 2)
        .attr("y", d => y(d.total) - 8)
        .attr("text-anchor", "middle")
        .style("fill", "var(--sidebar)")
        .style("font-size", "12px")
        .style("font-weight", "bold")
        .text(d => formatCurrency(d.total).replace('$', '').split('.')[0]);

    // Custom HTML Legend with Logos
    const legendDiv = d3.select(container).append("div")
        .style("display", "flex")
        .style("justify-content", "center")
        .style("gap", "1.5rem")
        .style("margin-top", "1rem")
        .style("flex-wrap", "wrap")
        .style("align-items", "center");

    bancos.forEach(b => {
        const item = legendDiv.append("div")
            .style("display", "flex")
            .style("align-items", "center")
            .style("gap", "0.4rem");
            
        // Pequeño indicador de color
        item.append("div")
            .style("width", "10px")
            .style("height", "10px")
            .style("border-radius", "2px")
            .style("background-color", b.color);
            
        if (b.logo) {
            item.append("img")
                .attr("src", b.logo)
                .attr("alt", b.label)
                .style("height", "18px")
                .style("object-fit", "contain");
        } else {
            item.append("span")
                .style("font-size", "13px")
                .style("font-weight", "700")
                .style("color", "var(--sidebar)")
                .text(b.label);
        }
    });
}

/**
 * Render the Cash Flow Table
 */
function renderCashFlow(data, selectedIndex = -1) {
    const headerEl = document.getElementById('cashflowHeader');
    const bodyEl = document.getElementById('cashflowBody');
    const periodLabel = document.getElementById('cashflowPeriodLabel');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    const curr = data[endIdx];
    
    const startIdx = Math.max(0, endIdx - 5);
    let visibleMonths = data.slice(startIdx, endIdx + 1);

    // Fix Diciembre 2025 as the first column, filter out the rest of 2025
    visibleMonths = visibleMonths.filter(m => isYear2026(m));
    const dic2025Cash = data.find(d => isYear2025(d) && (d.date.toLowerCase().includes('dic') || d.date.toLowerCase().includes('dec')));
    if (dic2025Cash && !visibleMonths.includes(dic2025Cash)) {
        visibleMonths.unshift(dic2025Cash);
    }
    
    const periods = visibleMonths.map(d => d.date);
    periodLabel.textContent = `Análisis de Ciclo de Caja al ${curr.date}`;

    headerEl.innerHTML = `
        <tr>
            <th>Concepto / Flujo de Efectivo</th>
            ${periods.map(p => `<th>${p}</th>`).join('')}
        </tr>
    `;

    // Definition of rows in order
    const rowSpec = [
        { key: 'beginning', label: 'Efectivo Inicial', isHeader: true },
        { key: 'ebitda', label: 'EBITDA (Origen P&L)', isSource: true },
        { type: 'separator', label: 'Cambios en Capital de Trabajo' },
        { key: 'cxc', label: ' (Aumento)/Disminución CxC', indent: true },
        { key: 'inv', label: ' (Aumento)/Disminución Inventario', indent: true },
        { key: 'cxp', label: ' Aumento/(Disminución) CxP', indent: true },
        { key: 'wc', label: 'Total Cambios Capital Trabajo', isTotal: true },
        { type: 'separator', label: 'Otros Ajustes Operativos' },
        { key: 'taxes', label: 'Taxes', indent: true },
        { key: 'extraordinary', label: 'Gastos/Ingresos Extraordinarios', indent: true },
        { key: 'operating', label: 'FLUJO DE CAJA OPERATIVO', isTotal: true },
        { type: 'separator', label: 'Actividades de Inversión' },
        { key: 'capex', label: 'Inversiones de Capital (CAPEX)', indent: true },
        { type: 'separator', label: 'Actividades de Financiamiento' },
        { key: 'netDebt', label: 'Variación Deuda Neta', indent: true },
        { key: 'interest', label: 'Gastos Financieros / Intereses', indent: true },
        { key: 'dividends', label: 'Actividades con Accionistas / Otros', indent: true },
        { key: 'financing', label: 'Total Flujo Financiamiento', isTotal: true },
        { key: 'change', label: 'VARIACIÓN NETA DE EFECTIVO', isHeader: true },
        { key: 'ending', label: 'Efectivo Final', isHeader: true }
    ];

    bodyEl.innerHTML = rowSpec.map(spec => {
        if (spec.type === 'separator') {
            return `<tr class="row-category"><td colspan="${periods.length + 1}" style="background:rgba(0,0,0,0.02); font-weight:700; font-size:0.75rem; color:var(--text-secondary); text-transform:uppercase; letter-spacing:0.5px;">${spec.label}</td></tr>`;
        }

        const cells = visibleMonths.map(period => {
            let val = 0;
            if (spec.key === 'ebitda') {
                val = period.kpis?.ebitda || 0;
            } else {
                val = period.cashflowDetail?.[spec.key] || 0;
            }

            const color = val < 0 ? 'var(--danger)' : 'inherit';
            return `<td style="color:${color}">${formatCurrency(val)}</td>`;
        }).join('');

        let className = spec.isTotal ? 'row-total' : (spec.isHeader ? 'row-category' : '');
        let cellClass = spec.indent ? 'row-indent' : '';
        if (spec.isHeader) className += ' row-highlight';

        return `<tr class="${className}">
            <td class="${cellClass}">${spec.label}</td>
            ${cells}
        </tr>`;
    }).join('');

    // Metrics Section (DSO, DPO, DIO)
    const metricsBody = document.getElementById('cfMetricsBody');
    const metricsHeader = document.getElementById('cfMetricsHeader');
    const metricsSection = document.getElementById('cf-metrics-section');

    const hasMetrics = visibleMonths.some(m => m.cashflowDetail?.dso || m.cashflowDetail?.dpo || m.cashflowDetail?.dio);
    
    if (hasMetrics && metricsBody && metricsHeader) {
        metricsSection.style.display = 'block';
        metricsHeader.innerHTML = `<tr><th>Indice de Eficiencia (Días)</th>${periods.map(p => `<th>${p}</th>`).join('')}</tr>`;
        
        const metricRows = [
            { key: 'dso', label: 'DSO (Días Cuentas por Cobrar)' },
            { key: 'dpo', label: 'DPO (Días Cuentas por Pagar)' },
            { key: 'dio', label: 'DIO (Días Rotación Inventario)' },
            { key: 'ccc', label: 'CCC (Ciclo de Conversión de Efectivo)' }
        ];

        metricsBody.innerHTML = metricRows.map(m => {
            const cells = visibleMonths.map(p => {
                let val = 0;
                if (m.key === 'ccc') {
                    const dso = p.cashflowDetail?.dso || 0;
                    const dio = p.cashflowDetail?.dio || 0;
                    const dpo = p.cashflowDetail?.dpo || 0;
                    val = dso + dio - dpo;
                } else {
                    val = p.cashflowDetail?.[m.key] || 0;
                }
                return `<td>${Math.round(val)} días</td>`;
            }).join('');
            return `<tr><td>${m.label}</td>${cells}</tr>`;
        }).join('');
    } else {
        metricsSection.style.display = 'none';
    }

    // Render Cash Flow Resumen
    renderCashFlowResumen(dic2025Cash || visibleMonths[0], curr, data, endIdx);
}

function renderCashFlowResumen(dicData, currData, fullData, selectedIndex) {
    const headerEl = document.getElementById('cashflowResumenHeader');
    const bodyEl = document.getElementById('cashflowResumenBody');
    if (!headerEl || !bodyEl) return;

    const dicLabel = dicData ? dicData.date : 'N/A';
    const currLabel = currData ? currData.date : 'N/A';
    
    headerEl.innerHTML = `
        <tr>
            <th style="background: var(--sidebar); color: white;">Concepto</th>
            <th style="text-align: right; background: var(--sidebar-dark); color: white;">${dicLabel}</th>
            <th style="text-align: right; background: var(--sidebar-dark); color: white;">YTD ${currLabel}</th>
            <th style="text-align: right; background: var(--sidebar-dark); color: white;">PPTO YTD</th>
        </tr>
    `;

    if (!currData) return;

    // Function to sum year-to-date or full year data based on an up-to index
    const calcYTD = (targetItem, isPpto = false) => {
        let result = {};
        if (!targetItem) return result;
        const targetYear = targetItem.sortDate ? getSortYear(targetItem) : 2026;
        const targetIdx = fullData.indexOf(targetItem);
        
        let firstMonth;
        for (let k = 0; k <= targetIdx; k++) {
            const item = fullData[k];
            if (item.sortDate && getSortYear(item) === targetYear) {
                if (!firstMonth) firstMonth = item;
                
                const detail = isPpto ? (item.ppto?.cashflowDetail || {}) : (item.cashflowDetail || {});
                const kpis = isPpto ? (item.ppto?.kpis || {}) : (item.kpis || {});

                Object.keys(detail).forEach(key => {
                    if (key !== 'beginning' && key !== 'ending' && key !== 'ccc' && key !== 'dso' && key !== 'dpo' && key !== 'dio') {
                        result[key] = (result[key] || 0) + (detail[key] || 0);
                    }
                });
                result.ebitda = (result.ebitda || 0) + (kpis.ebitda || 0);
            }
        }
        
        // Point in time values
        const endDetail = isPpto ? (targetItem.ppto?.cashflowDetail || {}) : (targetItem.cashflowDetail || {});
        const startDetail = firstMonth ? (isPpto ? (firstMonth.ppto?.cashflowDetail || {}) : (firstMonth.cashflowDetail || {})) : {};
        
        result.beginning = startDetail.beginning || 0;
        result.ending = endDetail.ending || 0;
        result.dso = endDetail.dso || 0;
        result.dpo = endDetail.dpo || 0;
        result.dio = endDetail.dio || 0;
        result.ccc = endDetail.ccc || 0;
        
        result.cxp = (result.cxp || 0);
        
        result.change = result.ending - result.beginning;

        // Internal verification for debugging purposes
        const flowSum = (result.operating||0) + (result.capex||0) + (result.netDebt||0) + 
                        (result.interest||0) + (result.extraordinary||0) + (result.diferencialCambiario||0);
        const expectedChange = result.ending - result.beginning;
        if (Math.abs(flowSum - expectedChange) > 1) {
            console.warn(`[CF Resumen] Descuadre ${isPpto?'PPTO':'Real'}: flujos=${flowSum.toFixed(2)}, cambio esperado=${expectedChange.toFixed(2)}, diferencia=${(flowSum-expectedChange).toFixed(2)}`);
        }

        return result;
    };

    const targetYearDic = dicData && dicData.sortDate ? getSortYear(dicData) : 2025;
    const ytdDic = calcYTD(dicData, false);
    
    // Override starting cash if necessary based on Y-1 ending (if available in fullData)
    const prevYearItem = fullData.find(d => d.sortDate && getSortYear(d) === targetYearDic - 1 && d.sortDate.getMonth() === 11);
    if(prevYearItem && prevYearItem.cashflowDetail) {
        ytdDic.beginning = prevYearItem.cashflowDetail.ending || ytdDic.beginning;
    }

    const ytdCurr = calcYTD(currData, false);
    const ytdPpto = calcYTD(currData, true);

    // Override starting cash for Curr using dicData ending
    const isPrevYearDec = dicData && dicData.date && (dicData.date.toLowerCase().includes('dic') || dicData.date.toLowerCase().includes('dec'));
    if (dicData && dicData.cashflowDetail && isPrevYearDec) {
        ytdCurr.beginning = dicData.cashflowDetail.ending || ytdCurr.beginning;
    }
    // For PPTO YTD, we do NOT override beginning cash with December 2025's Actual ending cash,
    // as PPTO has its own budgeted starting cash balance which aligns with all its budgeted cash flows.

    const rowSpec = [
        { key: 'beginning', label: 'Efectivo Inicial', highlight: true },
        { key: 'ebitda', label: 'EBITDA', highlight: true },
        { key: 'cxc', label: 'Cuentas por Cobrar' },
        { key: 'inv', label: 'Inventario' },
        { key: 'otrosActivos', label: 'Otros Activos' },
        { key: 'cxp', label: 'Cuentas por Pagar' },
        { key: 'pasivoLaboral', label: 'Pasivo Laboral' },
        { key: 'otrosPasivos', label: 'Otros Pasivos' },
        { key: 'taxes', label: 'Taxes' },
        { key: 'operating', label: 'Flujo de Caja Operativo', highlight: true, overlayBg: 'rgba(0,150,199,0.1)' },
        { key: 'capex', label: 'CAPEX', highlight: true },
        { key: 'netDebt', label: 'Aumento Deuda Neta' },
        { key: 'interest', label: 'Gastos de Interes' },
        { key: 'extraordinary', label: 'Ingresos (Gastos) Extraordinarios' },
        { key: 'diferencialCambiario', label: 'Diferencial Cambiario' },
        { key: 'change', label: 'Cambio en Efectivo', highlight: true, overlayBg: 'rgba(0,150,199,0.1)' },
        { key: 'ending', label: 'Efectivo Final', highlight: true, overlayBg: 'rgba(0,150,199,0.1)' },
        { type: 'separator' },
        { key: 'dso', label: 'DSO', isMetric: true },
        { key: 'dpo', label: 'DPO', isMetric: true },
        { key: 'dio', label: 'DIO', isMetric: true }
    ];

    const evaluateCFColumn = (ytdSource, isDic = false) => {
        const valMap = {};
        rowSpec.forEach(spec => {
            if (spec.type !== 'separator') {
                valMap[spec.key] = ytdSource ? ytdSource[spec.key] || 0 : 0;
            }
        });

        // 1. Calculate Flujo de Caja Operativo = EBITDA + CxC + Inv + OtrosActivos + CxP + PasivoLaboral + OtrosPasivos + Taxes
        valMap['operating'] = (valMap['ebitda'] || 0) + 
                              (valMap['cxc'] || 0) + 
                              (valMap['inv'] || 0) + 
                              (valMap['otrosActivos'] || 0) + 
                              (valMap['cxp'] || 0) + 
                              (valMap['pasivoLaboral'] || 0) + 
                              (valMap['otrosPasivos'] || 0) + 
                              (valMap['taxes'] || 0);

        // 2. Calculate Cambio en Efectivo = Ending - Beginning (guarantees perfect consistency)
        valMap['change'] = (valMap['ending'] || 0) - (valMap['beginning'] || 0);

        // 3. Calculate Efectivo Final = Beginning + Change
        valMap['ending'] = (valMap['beginning'] || 0) + (valMap['change'] || 0);

        // 4. Calculate Net Debt as the balancing plug of the cash flow statement:
        // change = operating + capex + netDebt + interest + extraordinary + diferencialCambiario
        // therefore netDebt = change - (operating + capex + interest + extraordinary + diferencialCambiario)
        if (isDic) {
            valMap['netDebt'] = ytdSource ? ytdSource['netDebt'] || 0 : 0;
        } else {
            valMap['netDebt'] = valMap['change'] - (
                (valMap['operating'] || 0) + 
                (valMap['capex'] || 0) + 
                (valMap['interest'] || 0) + 
                (valMap['extraordinary'] || 0) + 
                (valMap['diferencialCambiario'] || 0)
            );
        }

        return valMap;
    };

    const finalYtdDic = evaluateCFColumn(ytdDic, true);
    const finalYtdCurr = evaluateCFColumn(ytdCurr, false);
    const finalYtdPpto = evaluateCFColumn(ytdPpto, false);

    console.log("CASH FLOW RESUMEN TAXES DEBUG:");
    console.log("DIC Data target element:", dicData ? dicData.date : null);
    console.log("CURR Data target element:", currData ? currData.date : null);
    console.log("ytdDic.taxes:", ytdDic.taxes, "finalYtdDic.taxes:", finalYtdDic.taxes);
    console.log("ytdCurr.taxes:", ytdCurr.taxes, "finalYtdCurr.taxes:", finalYtdCurr.taxes);
    console.log("ytdPpto.taxes:", ytdPpto.taxes, "finalYtdPpto.taxes:", finalYtdPpto.taxes);

    bodyEl.innerHTML = rowSpec.map(spec => {
        if (spec.type === 'separator') {
            return `<tr><td colspan="4" style="height: 16px; background: transparent; border: none;"></td></tr>`;
        }

        const vDic = finalYtdDic[spec.key] !== undefined ? finalYtdDic[spec.key] : (ytdDic ? ytdDic[spec.key] || 0 : 0);
        const vCurr = finalYtdCurr[spec.key] !== undefined ? finalYtdCurr[spec.key] : (ytdCurr ? ytdCurr[spec.key] || 0 : 0);
        const vPpto = finalYtdPpto[spec.key] !== undefined ? finalYtdPpto[spec.key] : (ytdPpto ? ytdPpto[spec.key] || 0 : 0);
        
        let styleStr = 'padding: 8px 12px;';
        if (spec.highlight) styleStr += ' font-weight: 700;';
        if (spec.overlayBg) styleStr += ` background-color: ${spec.overlayBg};`;
        if (spec.isMetric) styleStr += ' background-color: #e2e8f0; font-weight: 600; font-size: 0.85rem;';
        
        const formatFn = spec.isMetric ? (v) => Math.round(v) : formatCurrency;

        return `<tr>
            <td style="${styleStr}">${spec.label}</td>
            <td style="text-align: right; ${styleStr}">${formatFn(vDic)}</td>
            <td style="text-align: right; ${styleStr}">${formatFn(vCurr)}</td>
            <td style="text-align: right; ${styleStr}">${formatFn(vPpto)}</td>
        </tr>`;
    }).join('');
}

function renderPreliminaryView(data, selectedIndex = -1) {
    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    const curr = data[endIdx];
    if (!curr) return;
    
    const labelEl = document.getElementById('preliminarDateLabel');
    if(labelEl) {
        labelEl.innerText = isYTDMode ? `YTD ${curr.date}` : curr.date;
    }

    const prevYearItem = data.find(d => {
        const dDate = d.sortDate;
        const cDate = curr.sortDate;
        return dDate.getMonth() === cDate.getMonth() && dDate.getFullYear() === cDate.getFullYear() - 1;
    }) || null;

    const tableBody = document.getElementById('preliminarBody');
    if (!tableBody) return;

    const getVal = (item, keywords, context) => {
        if (!item) return 0;

        const findValueInRows = (rows, kw) => {
            // Filter by context if provided
            let searchRows = rows;
            if (context) {
                const idxMargen = rows.findIndex(r => r.concept && r.concept.toLowerCase().includes("margen bruto"));
                let idxCosto = -1;
                for (let i = 0; i < rows.length; i++) {
                     let cLabel = rows[i].concept ? rows[i].concept.trim().toLowerCase() : "";
                     if (cLabel === "costo de ventas" || cLabel === "costos" || cLabel === "costos de operacion") {
                          idxCosto = i;
                          break;
                     }
                }
                
                if (context === "ventas" && idxCosto !== -1) {
                     searchRows = rows.slice(0, idxCosto);
                } else if (context === "costos" && idxCosto !== -1) {
                     const end = idxMargen !== -1 ? idxMargen : rows.length;
                     searchRows = rows.slice(idxCosto, end);
                } else if (context === "gastos" && idxMargen !== -1) {
                     searchRows = rows.slice(idxMargen);
                }
            }

            const findValueRobust = (rowObj, searchDate) => {
                if (!rowObj || !rowObj.values) return null;
                if (Object.prototype.hasOwnProperty.call(rowObj.values, searchDate)) return rowObj.values[searchDate];
                // Try to find matching month-year without caring about format
                const sDate = searchDate.toLowerCase().replace(/[^a-z0-9]/g, '');
                for (let k in rowObj.values) {
                    const kStr = k.toLowerCase().replace(/[^a-z0-9]/g, '');
                    if (kStr.includes(sDate) || sDate.includes(kStr) || kStr === sDate) {
                        return rowObj.values[k];
                    }
                }
                return null;
            };

            const exact = searchRows.find(r => r.concept && r.concept.trim().toLowerCase() === kw && !r.concept.toLowerCase().includes('acumulada'));
            if (exact) {
                const bestVal = findValueRobust(exact, item.date);
                if (bestVal !== null) return bestVal;
            }

            const partials = searchRows.filter(r => r.concept && r.concept.toLowerCase().includes(kw) && !r.concept.toLowerCase().includes('acumulada'));
            partials.sort((a, b) => a.concept.length - b.concept.length);
            
            for (const p of partials) {
                const bestVal = findValueRobust(p, item.date);
                if (bestVal !== null) return bestVal;
            }
            return null;
        };

        for (const kw of keywords) {
            let pnlRows = item.pnl?.fullRows || [];
            let sourceRows = (kw.includes("descuento") || kw.includes("devolucion") || kw.includes("devolución")) && item.estados && item.estados.fullRows && item.estados.fullRows.length > 0 ? item.estados.fullRows : pnlRows;
            let val = findValueInRows(sourceRows, kw.trim());
            if (val !== null && val !== 0) return val;

            if (!context) {
                val = findValueInRows(item.balance?.fullRows || [], kw.trim());
                if (val !== null && val !== 0) return val;

                val = findValueInRows(item.wcFullRows || [], kw.trim());
                if (val !== null && val !== 0) return val;
            }
        }
        
        for (const kw of keywords) {
             let pnlRows = item.pnl?.fullRows || [];
             let sourceRows = (kw.includes("descuento") || kw.includes("devolucion") || kw.includes("devolución")) && item.estados && item.estados.fullRows && item.estados.fullRows.length > 0 ? item.estados.fullRows : pnlRows;
             let val = findValueInRows(sourceRows, kw.trim());
             if (val !== null) return val;
        }

        return 0;
    };

    const structure = [
        { label: "Ventas Brutas", match: ["ventas brutas", "ventas bon", "ventas p6"], isTotal: false },
        { label: "Descuentos y Devoluciones", match: ["descuentos y devoluciones"], isTotal: false, id: "desc" },
        { label: "Descuentos", match: ["descuentos", "descuento", "descuento sobre ventas", "descuento en ventas", "menos descuentos"], isTotal: false, isSubItem: true, parentId: "desc" },
        { label: "Devoluciones", match: ["devoluciones", "devolucion", "devoluciones sobre ventas", "devolución", "devolución en ventas", "menos devoluciones"], isTotal: false, isSubItem: true, parentId: "desc" },
        { label: "Ventas netas", match: ["ventas netas", "ingresos"], isTotal: true, id: "vnetas" },
        { label: "EVP", match: ["evp"], isTotal: false, isSubItem: true, restrictTo: "ventas", parentId: "vnetas" },
        { label: "BT5", match: ["bt5"], isTotal: false, isSubItem: true, restrictTo: "ventas", parentId: "vnetas" },
        { label: "BON", match: ["bon", "p6"], isTotal: false, isSubItem: true, restrictTo: "ventas", parentId: "vnetas" },
        { label: "Otros Ingresos", match: ["otras ventas", "otros ingresos"], isTotal: false, isSubItem: true, parentId: "vnetas" },
        { label: "Costo de ventas", match: ["costo de ventas", "costos", "costo", "costos de operacion"], isTotal: false, id: "costos" },
        { label: "EVP ", match: ["evp", "costo evp"], isTotal: false, isSubItem: true, restrictTo: "costos", parentId: "costos" },
        { label: "BT5 ", match: ["bt5", "costo bt5"], isTotal: false, isSubItem: true, restrictTo: "costos", parentId: "costos" },
        { label: "BON ", match: ["bon", "p6", "costo bon"], isTotal: false, isSubItem: true, restrictTo: "costos", parentId: "costos" },
        { label: "Margen Bruto", match: ["margen bruto", "gross margin", "utilidad bruta", "gross profit"], isTotal: true },
        { label: "Gastos Administrativos", match: ["gastos administrativos", "administrativo", "admin "], isTotal: false, id: "gadmin" },
        { label: "ITBIS", match: ["itbis", "impuesto itbis"], isTotal: false, isSubItem: true, restrictTo: "gastos", parentId: "gadmin" },
        { label: "Otros Gastos Operativos", match: ["otros gastos operativos", "otros gastos"], isTotal: false, isSubItem: true, restrictTo: "gastos", parentId: "gadmin" },
        { label: "Gastos Administrativos ", match: ["gastos administrativos", "administrativo", "admin "], isTotal: false, isSubItem: true, restrictTo: "gastos", parentId: "gadmin" },
        { label: "Gastos de Mercadeo", match: ["gastos de mercadeo", "mercadeo"], isTotal: false },
        { label: "Gastos de Ventas", match: ["gastos de ventas", "comercial"], isTotal: false },
        { label: "Gastos de Logística", match: ["gastos de logística", "logistica", "logística"], isTotal: false },
        { label: "EBITDA", match: ["ebitda"], isTotal: true },
        { label: "EBITDA (mUSD)", match: ["ebitda usd", "ebitda (usd)", "ebitda us$", "ebitda us", "ebitda en dolares"], isTotal: false },
        { label: "D & A", match: ["depreciación y amortización", "depreciacion y amortizacion", "d&a", "d & a", "depreciaciones", "depreciacion", "amortizacion"], isTotal: false },
        { label: "Intereses Netos", match: ["intereses", "financiero"], isTotal: false, id: "intnetos" },
        { label: "Ingresos Financieros", match: ["ingresos financieros"], isTotal: false, isSubItem: true, parentId: "intnetos" },
        { label: "Gastos Financieros", match: ["gastos financieros"], isTotal: false, isSubItem: true, parentId: "intnetos" },
        { label: "Diferencial cambiario", match: ["diferencial cambiario", "cambiario"], isTotal: false },
        { label: "Ingresos (Gastos) Extraordinarios", match: ["extraordinario", "extraordinarios", "ingresos (gastos) extraordinarios"], isTotal: false },
        { label: "Utilidad antes de impuesto", match: ["utilidad antes de impuesto", "antes de impuesto", "net income before"], isTotal: true },
        { label: "Tasa de cierre USD", match: ["fx", "tasa cambio", "tasa de cambio", "dop", "tipo de cambio", "t.c", "tc", "tasa cambio cierre", "fx eop"], isTotal: false }
    ];

    window.expandedGroups = window.expandedGroups || new Set();

    window.toggleGroup = function(groupId, btn) {
        const rows = document.querySelectorAll('.subitem-' + groupId);
        let isHidden = false;
        rows.forEach(r => {
            if (r.style.display === 'none') {
                r.style.display = '';
                isHidden = false;
            } else {
                r.style.display = 'none';
                isHidden = true;
            }
        });
        if (btn) btn.innerHTML = isHidden ? '+' : '-';
        
        if (isHidden) {
            window.expandedGroups.delete(groupId);
        } else {
            window.expandedGroups.add(groupId);
        }
    }

    let html = '';
    
    const resolveRowStandardVal = (targetItem, rowLabel, rowMatch, context) => {
        if (!targetItem) return { actual: 0, ppto: 0 };
        let actual = getVal(targetItem, rowMatch, context);
        let ppto = 0;

        if (targetItem.ppto && targetItem.ppto.pnl) {
            if (targetItem.ppto.pnl.categorias) {
                if (rowLabel === "Ventas netas") ppto = targetItem.ppto.kpis.ingresos || 0;
                if (rowLabel === "Costo de ventas") ppto = targetItem.ppto.pnl.categorias["Costo de Ventas"] || 0;
                if (rowLabel === "EBITDA") ppto = targetItem.ppto.kpis.ebitda || targetItem.ppto.pnl.categorias["EBITDA"] || 0;
                if (rowLabel === "Utilidad antes de impuesto") ppto = targetItem.ppto.kpis.utilidad || 0;
                if (rowLabel === "Margen Bruto") ppto = (targetItem.ppto.kpis.ingresos || 0) + (targetItem.ppto.pnl.categorias["Costo de Ventas"] || 0);
                if (rowLabel === "EBITDA (mUSD)") {
                    const ebitdaPpto = targetItem.ppto.kpis.ebitda || targetItem.ppto.pnl.categorias["EBITDA"] || 0;
                    const tasaPpto = targetItem.ppto.tasaCambio || 1;
                    ppto = tasaPpto !== 0 ? (ebitdaPpto / tasaPpto) : 0;
                }

                if (rowLabel === "Gastos Administrativos" || rowLabel === "Gastos Administrativos ") ppto = targetItem.ppto.pnl.opexDetalle?.["Gastos Administrativos"] || 0;
                if (rowLabel === "Gastos de Mercadeo") ppto = targetItem.ppto.pnl.opexDetalle?.["Gastos de Mercadeo"] || 0;
                if (rowLabel === "Gastos de Ventas") ppto = targetItem.ppto.pnl.opexDetalle?.["Gastos de Ventas (Comercial)"] || 0;
                if (rowLabel === "Gastos de Logística") ppto = targetItem.ppto.pnl.opexDetalle?.["Gastos de Logística"] || 0;
            }
            if (targetItem.ppto.pnl.segments) {
                if (rowLabel.trim() === "EVP") ppto = context === "costos" ? targetItem.ppto.pnl.segments["EVP"]?.costos || 0 : targetItem.ppto.pnl.segments["EVP"]?.ventas || 0;
                if (rowLabel.trim() === "BT5") ppto = context === "costos" ? targetItem.ppto.pnl.segments["BT5"]?.costos || 0 : targetItem.ppto.pnl.segments["BT5"]?.ventas || 0;
                if (rowLabel.trim() === "BON" || rowLabel.trim() === "P6") ppto = context === "costos" ? (targetItem.ppto.pnl.segments["BON"]?.costos || targetItem.ppto.pnl.segments["P6"]?.costos || 0) : (targetItem.ppto.pnl.segments["BON"]?.ventas || targetItem.ppto.pnl.segments["P6"]?.ventas || 0);
                if (rowLabel.trim() === "Otros Ingresos") ppto = (targetItem.ppto.pnl.segments["Otras Ventas"]?.ventas || 0) + (targetItem.ppto.pnl.segments["Otros Ingresos"]?.ventas || 0);
            }
        }

        if (rowLabel === "Tasa de cierre USD") {
            if (targetItem.ppto && targetItem.ppto.tasaCambio !== undefined) {
                ppto = targetItem.ppto.tasaCambio;
            }
        }

        if (rowLabel === "Descuentos") {
            let actualDesc = getVal(targetItem, ["descuentos", "descuento", "descuento sobre ventas", "descuento en ventas", "menos descuentos"]);
            let actualDev = getVal(targetItem, ["devoluciones", "devolucion", "devoluciones sobre ventas", "devolución", "devolución en ventas", "menos devoluciones"]);
            actualDesc = actualDesc !== 0 ? -Math.abs(actualDesc) : 0;
            actualDev = actualDev !== 0 ? -Math.abs(actualDev) : 0;
            
            if (targetItem.ppto) {
                 const pptoVentasNetas = targetItem.ppto?.kpis?.ingresos || 0;
                 ppto = pptoVentasNetas !== 0 ? pptoVentasNetas * -0.0387749832441863 : 0;
            }
            if ((actualDesc + actualDev) === 0) {
                 const actualVentasNetas = targetItem?.kpis?.ingresos || 0;
                 actual = actualVentasNetas !== 0 ? actualVentasNetas * -0.0387749832441863 : 0;
            } else {
                 actual = actualDesc;
            }
        }

        if (rowLabel === "Devoluciones") {
            let actualDesc = getVal(targetItem, ["descuentos", "descuento", "descuento sobre ventas", "descuento en ventas", "menos descuentos"]);
            let actualDev = getVal(targetItem, ["devoluciones", "devolucion", "devoluciones sobre ventas", "devolución", "devolución en ventas", "menos devoluciones"]);
            actualDesc = actualDesc !== 0 ? -Math.abs(actualDesc) : 0;
            actualDev = actualDev !== 0 ? -Math.abs(actualDev) : 0;
            
            if ((actualDesc + actualDev) === 0) {
                actual = 0;
            } else {
                actual = actualDev;
            }
        }

        if (rowLabel === "Descuentos y Devoluciones") {
            let actualDesc = getVal(targetItem, ["descuentos", "descuento", "descuento sobre ventas", "descuento en ventas", "menos descuentos"]);
            let actualDev = getVal(targetItem, ["devoluciones", "devolucion", "devoluciones sobre ventas", "devolución", "devolución en ventas", "menos devoluciones"]);
            actualDesc = actualDesc !== 0 ? -Math.abs(actualDesc) : 0;
            actualDev = actualDev !== 0 ? -Math.abs(actualDev) : 0;
            let sumD = actualDesc + actualDev;
            
            if (targetItem.ppto) {
                 const pptoVentasNetas = targetItem.ppto?.kpis?.ingresos || 0;
                 ppto = pptoVentasNetas !== 0 ? pptoVentasNetas * -0.0387749832441863 : 0;
            }
            if (sumD === 0) {
                 const actualVentasNetas = targetItem?.kpis?.ingresos || 0;
                 actual = actualVentasNetas !== 0 ? actualVentasNetas * -0.0387749832441863 : 0;
            } else {
                 actual = sumD;
            }
        }

        if (rowLabel === "Ventas Brutas") {
            const vNet = getVal(targetItem, ["ventas netas", "ingresos"]) || targetItem?.kpis?.ingresos || 0;
            let actualDesc = getVal(targetItem, ["descuentos", "descuento", "descuento sobre ventas", "descuento en ventas", "menos descuentos"]);
            let actualDev = getVal(targetItem, ["devoluciones", "devolucion", "devoluciones sobre ventas", "devolución", "devolución en ventas", "menos devoluciones"]);
            actualDesc = actualDesc !== 0 ? -Math.abs(actualDesc) : 0;
            actualDev = actualDev !== 0 ? -Math.abs(actualDev) : 0;
            let sumD = actualDesc + actualDev;
            
            if (sumD === 0) {
                 const actualVentasNetas = targetItem?.kpis?.ingresos || 0;
                 sumD = actualVentasNetas !== 0 ? actualVentasNetas * -0.0387749832441863 : 0;
            }
            actual = vNet + Math.abs(sumD);

            if (targetItem.ppto) {
                 const pptoVentasNetas = targetItem.ppto?.kpis?.ingresos || 0;
                 const pptoDescuentos = pptoVentasNetas * -0.0387749832441863;
                 ppto = pptoVentasNetas + Math.abs(pptoDescuentos);
            }
        }

        if (rowLabel === "D & A") {
            const depActual = getVal(targetItem, ["depreciacion", "depreciación"]) || 0;
            const amoActual = getVal(targetItem, ["amortizacion", "amortización"]) || 0;
            const combined = getVal(targetItem, ["depreciacion y amortizacion", "depreciación y amortización", "d&a", "d & a"]);
            // If we found a combined field, use it, else sum.
            actual = combined ? combined : (depActual + amoActual);
        }

        if (rowLabel === "Margen Bruto") {
            const vNet = getVal(targetItem, ["ventas netas", "ingresos"]) || targetItem?.kpis?.ingresos || 0;
            const cVentas = getVal(targetItem, ["costo de ventas", "costos", "costo", "costos de operacion"]) || targetItem?.pnl?.categorias?.["Costo de Ventas"] || 0;
            actual = vNet + cVentas;
        }

        if (rowLabel === "Ventas netas") {
            const evp = getVal(targetItem, ["evp"], "ventas") || 0;
            const bt5 = getVal(targetItem, ["bt5"], "ventas") || 0;
            const bon = getVal(targetItem, ["bon", "p6"], "ventas") || 0;
            const otros = getVal(targetItem, ["otras ventas", "otros ingresos"]) || 0;
            actual = evp + bt5 + bon + otros;
        }
        if (rowLabel === "EBITDA" && !actual) actual = targetItem?.kpis?.ebitda || 0;
        if (rowLabel === "Costo de ventas" && !actual) actual = targetItem?.pnl?.categorias?.["Costo de Ventas"] || 0;
        if (rowLabel === "Utilidad antes de impuesto" && !actual) actual = targetItem?.kpis?.utilidad || 0;

        // Generic fallback for any unmapped PPTO value
        if (targetItem.ppto && targetItem.ppto.pnl && targetItem.ppto.pnl.fullRows && ppto === 0 && rowLabel !== "Margen Bruto" && rowLabel !== "Ventas Brutas" && rowLabel !== "Descuentos y Devoluciones" && rowLabel !== "Ventas netas" && rowLabel !== "Utilidad antes de impuesto" && rowLabel !== "Costo de ventas" && rowLabel !== "EBITDA") {
            // Re-use `getVal` but against the ppto object!
            const tempVal = getVal({ pnl: { fullRows: targetItem.ppto.pnl.fullRows }, date: targetItem.date }, rowMatch, context);
            if (tempVal !== null && tempVal !== 0) {
                 ppto = tempVal;
            } else if (rowLabel === "D & A") {
                 const d1 = getVal({ pnl: { fullRows: targetItem.ppto.pnl.fullRows }, date: targetItem.date }, ["depreciacion", "depreciación"]);
                 const d2 = getVal({ pnl: { fullRows: targetItem.ppto.pnl.fullRows }, date: targetItem.date }, ["amortizacion", "amortización"]);
                 const dc = getVal({ pnl: { fullRows: targetItem.ppto.pnl.fullRows }, date: targetItem.date }, ["depreciacion y amortizacion", "d&a"]);
                 ppto = dc ? dc : (d1 + d2);
            }
        }

        return { actual, ppto };
    };

    const getYtdAggregated = (targetIdx, row) => {
        if (targetIdx < 0 || targetIdx >= data.length) return { actual: 0, ppto: 0 };
        const item = data[targetIdx];
        if (!item) return { actual: 0, ppto: 0 };

        if (isYTDMode && row.label === "Tasa de cierre USD") {
            const targetYear = getSortYear(item);
            let sumEbitdaLocal = 0;
            let sumEbitdaUsd = 0;
            
            const parseDirtyNumber = (val) => {
                if (!val) return 0;
                if (typeof val === 'number') return val;
                let cleaned = val.toString().replace(/[^0-9.-]+/g, "");
                return Number(cleaned) || 0;
            };

            for (let k = targetIdx; k >= 0; k--) {
                const iterItem = data[k];
                if (getSortYear(iterItem) !== targetYear) break;
                
                let localVal = 0;
                let usdVal = 0;

                if (iterItem.pnl && iterItem.pnl.fullRows) {
                    const localRow = iterItem.pnl.fullRows.find(r => r.concept === "EBITDA");
                    if (localRow) localVal = localRow.values[iterItem.date] || 0;
                    if (localVal === 0 && iterItem.kpis?.ebitda) localVal = iterItem.kpis.ebitda;

                    const usdRow = iterItem.pnl.fullRows.find(r => r.concept === "EBITDA US$");
                    if (usdRow) usdVal = usdRow.values[iterItem.date] || 0;
                }

                sumEbitdaLocal += parseDirtyNumber(localVal);
                sumEbitdaUsd += parseDirtyNumber(usdVal);
            }

            let actualFxAcumulado = 0;
            if (sumEbitdaUsd !== 0 && sumEbitdaLocal !== 0) {
                actualFxAcumulado = sumEbitdaLocal / sumEbitdaUsd;
            } else {
                actualFxAcumulado = resolveRowStandardVal(item, row.label, row.match, row.restrictTo).actual;
            }
            
            // Calculate YTD PPTO FX 
            let pptoEbitdaLocal = 0;
            let pptoEbitdaUsd = 0;
            
            for (let k = targetIdx; k >= 0; k--) {
                const iterItem = data[k];
                if (getSortYear(iterItem) !== targetYear) break;
                
                let pptoLocalVal = 0;
                let pptoRate = iterItem.ppto?.tasaCambio || resolveRowStandardVal(iterItem, row.label, row.match, row.restrictTo).ppto || 1;
                
                if (iterItem.ppto?.kpis?.ebitda) {
                    pptoLocalVal = iterItem.ppto.kpis.ebitda;
                } else if (iterItem.ppto?.pnl?.fullRows) {
                    const localRowPpto = iterItem.ppto.pnl.fullRows.find(r => r.concept === "EBITDA");
                    if (localRowPpto) pptoLocalVal = localRowPpto.values[iterItem.date] || 0;
                }
                
                let valNum = parseDirtyNumber(pptoLocalVal);
                pptoEbitdaLocal += valNum;
                if (pptoRate !== 0) {
                    pptoEbitdaUsd += valNum / pptoRate;
                }
            }
            
            let pptoFxAcumulado = 0;
            if (pptoEbitdaUsd !== 0 && pptoEbitdaLocal !== 0) {
                pptoFxAcumulado = pptoEbitdaLocal / pptoEbitdaUsd;
            } else {
                pptoFxAcumulado = resolveRowStandardVal(item, row.label, row.match, row.restrictTo).ppto;
            }
            
            return { actual: actualFxAcumulado, ppto: pptoFxAcumulado };
        }

        const isTasa = row.label.includes("Tasa") || row.label.includes("Diferencial");
        if (!isYTDMode || isTasa) return resolveRowStandardVal(item, row.label, row.match, row.restrictTo);

        const targetYear = getSortYear(item);
        let sumActual = 0;
        let sumPpto = 0;
        for (let i = targetIdx; i >= 0; i--) {
            if (getSortYear(data[i]) !== targetYear) break;
            const vals = resolveRowStandardVal(data[i], row.label, row.match, row.restrictTo);
            sumActual += vals.actual;
            sumPpto += vals.ppto;
        }
        return { actual: sumActual, ppto: sumPpto };
    };

    const endIdxCurr = endIdx;
    const endIdxPrev = prevYearItem ? data.indexOf(prevYearItem) : -1;

    const vNetPrevVals = getYtdAggregated(endIdxPrev, {label: "Ventas netas", match: ["ventas netas", "ingresos"]});
    const vNetCurrVals = getYtdAggregated(endIdxCurr, {label: "Ventas netas", match: ["ventas netas", "ingresos"]});
    const vNetPrev = vNetPrevVals.actual;
    const vNetCurr = vNetCurrVals.actual;
    const vNetPpto = vNetCurrVals.ppto;

    const varianceThreshold = 1.0; 
    let costVar=[], adminVar=[], logisticaVar=[], diffCambiarioVar=[], extraordinariosVar=[];

    structure.forEach(row => {
        const prevVals = getYtdAggregated(endIdxPrev, row);
        const currVals = getYtdAggregated(endIdxCurr, row);
        
        let prevVal = prevVals.actual;
        let currVal = currVals.actual;
        let pptoVal = currVals.ppto;

        let pPrev = vNetPrev ? prevVal / vNetPrev : 0;
        let pCurr = vNetCurr ? currVal / vNetCurr : 0;
        let pPpto = vNetPpto ? pptoVal / vNetPpto : 0;

        const isTasa = row.label.includes("Tasa") || row.label.includes("mUSD");
        const hidePct = row.label === "Ventas Brutas" || row.label === "Descuentos y Devoluciones";
        let styleText = row.isTotal ? "font-weight:700; background: rgba(0,0,0,0.03);" : "";
        let bgRow = row.isTotal ? "background: rgba(0,0,0,0.03);" : "";
        
        let paddingLeft = "16px";
        if (row.isSubItem) {
             paddingLeft = "40px";
             styleText += " color: var(--text-secondary);";
        }

        const formatRowVal = (v) => {
            if (v === 0 || !v) return '-';
            if (row.label.includes("Tasa")) return '$ ' + v.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
            if (row.label.includes("mUSD")) return v.toLocaleString('en-US', {minimumFractionDigits: 1, maximumFractionDigits: 1});
            return formatCurrency(v);
        };

        const variancePpto = currVal - pptoVal;
        
        if (Math.abs(variancePpto) >= varianceThreshold && !row.isTotal && pptoVal !== 0 && currVal !== 0) {
            const isNegativeVal = row.label.includes("Costo") || row.label.includes("Gasto");
            const varFormatted = (variancePpto > 0 ? "+" : "") + variancePpto.toFixed(1) + "mDOP";
            
            if (row.label.includes("Costo") || row.label.includes("Mercadeo") || row.label.includes("Ventas") || row.label.includes("Descuentos")) {
                costVar.push(`${varFormatted} en ${row.label}`);
            }
            else if (row.label.includes("Administrativo")) adminVar.push(`${varFormatted} en ${row.label}`);
            else if (row.label.includes("Logística")) logisticaVar.push(`${varFormatted} en sub-cuentas (aprox.)`);
            else if (row.label.includes("cambiario")) diffCambiarioVar.push(`${varFormatted}`);
            else if (row.label.includes("Extraordinarios")) extraordinariosVar.push(`${varFormatted}`);
        }

        let isExpanded = window.expandedGroups.has(row.id);
        let collapseBtn = "";
        if (row.id) {
            collapseBtn = `<button class="collapse-btn" onclick="toggleGroup('${row.id}', this)">${isExpanded ? '-' : '+'}</button>`;
        }
        
        let idAttr = "";
        let displayStyle = "";
        if (row.isSubItem) {
             idAttr = `class="subitem-row subitem-${row.parentId}"`;
             let isParentExpanded = window.expandedGroups.has(row.parentId);
             if (!isParentExpanded) {
                 displayStyle = "display: none;";
             }
        }

        let prelimPulseClass = "";
        let prelimPulseTitle = "";
        if (pptoVal && pptoVal !== 0 && !row.isTotal) {
            const devPct = (currVal - pptoVal) / Math.abs(pptoVal);
            if (Math.abs(devPct) > 0.15) {
                const labelLower = row.label.toLowerCase();
                const isExpense = labelLower.includes("costo") || labelLower.includes("gasto") || labelLower.includes("itbis") || labelLower.includes("descuento") || labelLower.includes("devolucion") || labelLower.includes("d & a");
                const isPositiveBetter = !isExpense;
                const isBetter = isPositiveBetter ? (devPct > 0) : (devPct < 0);
                prelimPulseClass = isBetter ? "pulse-pos" : "pulse-neg";
                prelimPulseTitle = `Desviación de ${(devPct * 100).toFixed(1)}% respecto al presupuesto (${formatRowVal(pptoVal)}) para ${row.label}`;
            }
        }
        const innerClassStr = prelimPulseClass ? `class="${prelimPulseClass}" title="${prelimPulseTitle}" style="display:inline-block; padding: 2px 6px;"` : '';

        html += `<tr ${idAttr} style="${bgRow} ${displayStyle}">
            <td style="border-right: 1px solid rgba(0,0,0,0.05); padding: 5px; text-align: center; vertical-align: middle;">${collapseBtn}</td>
            <td style="border-right: 1px solid rgba(0,0,0,0.05); padding: 14px 16px; padding-left: ${paddingLeft}; font-size: 1.05em; ${styleText}">${formatSegmentName(row.label)}</td>
            <td style="text-align:right; padding: 14px 16px; font-size: 1.25em; ${styleText}">${formatRowVal(prevVal)}</td>
            <td style="text-align:right; color:var(--text-secondary); font-size:1.05em; padding: 14px 16px;">${(isTasa || hidePct) ? '' : (Math.abs(pPrev)*100).toFixed(0)+'%'}</td>
            <td style="text-align:right; border-left: 2px solid var(--sidebar-accent); background:rgba(0, 150, 199, 0.05); padding: 14px 16px; font-size: 1.25em; ${styleText}"><div ${innerClassStr}>${formatRowVal(currVal)}</div></td>
            <td style="text-align:right; background:rgba(0, 150, 199, 0.05); font-weight:600; font-size:1.05em; padding: 14px 16px;">${(isTasa || hidePct) ? '' : (Math.abs(pCurr)*100).toFixed(0)+'%'}</td>
            <td style="text-align:right; border-left: 1px solid rgba(0,0,0,0.05); padding: 14px 16px; font-size: 1.25em; ${styleText}">${formatRowVal(pptoVal)}</td>
            <td style="text-align:right; color:var(--text-secondary); font-size:1.05em; padding: 14px 16px;">${(isTasa || hidePct) ? '' : (Math.abs(pPpto)*100).toFixed(0)+'%'}</td>
        </tr>`;
    });

    tableBody.innerHTML = html;

    setTimeout(() => {
        if (typeof buildMobileAccordionsFromTable === 'function') {
            buildMobileAccordionsFromTable('preliminarTable', 'preliminarMobileContainer', 'Resumen Ejecutivo');
        }
    }, 10);
}

/**
 * Render the full P&L Detail Table with a 6-month rolling window
 */
function renderDetailedPnL(data, selectedIndex = -1) {
    const headerEl = document.getElementById('pnlDetailedHeader');
    const bodyEl = document.getElementById('pnlDetailedBody');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    // Use current selection as anchor, or last month if not specified
    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    // Show up to 6 months including the selected one (min 1)
    const startIdx = Math.max(0, endIdx - 5);
    
    // Filtro: No mostrar 2025 en el P&L
    const visibleMonths = data.slice(startIdx, endIdx + 1).filter(d => isYear2026(d));
    const periods = visibleMonths.map(d => d.date);
    
    // Header
    headerEl.innerHTML = `
        <tr>
            <th>Concepto / Cuenta</th>
            ${periods.map(p => `<th>${p}</th>`).join('')}
            <th>Acum. Periodo</th>
        </tr>
    `;

    let allConcepts = [];
    visibleMonths.forEach(d => {
        if (d.pnl && d.pnl.fullRows) {
            d.pnl.fullRows.forEach(row => {
                if (!allConcepts.includes(row.concept)) {
                    allConcepts.push(row.concept);
                }
            });
        }
    });

    // Filtros solicitados:
    allConcepts = allConcepts.filter(c => {
        const nc = normalizeText(c);
        if (nc === "concepto" || nc === "cuentas" || nc === "descripcion" || nc === "p&l" || nc === "resultado" || nc === "detalle") return false;
        if (nc.includes("en mdop") || nc.includes("reporte pa") || nc.includes("seguimiento gerencial") || nc.includes("margen operacional") || nc === "margen neto" || nc === "margen bruto ordinario") return false;
        return true;
    });

    // 2. Eliminar desde "PPE acumulado" hacia abajo
    const ppeIndex = allConcepts.findIndex(c => normalizeText(c).includes("ppe acumulado"));
    if (ppeIndex !== -1) {
        allConcepts = allConcepts.slice(0, ppeIndex);
    }

    if (allConcepts.length === 0) {
        bodyEl.innerHTML = `<tr><td colspan="${periods.length + 2}" style="text-align:center; padding:40px;">No se encontraron filas detalladas en el reporte para este periodo.</td></tr>`;
        return;
    }

    bodyEl.innerHTML = allConcepts.map(concept => {
        let periodAccumTotal = 0;
        const normConcept = normalizeText(concept);
        const isPercentage = normConcept.includes("margen") || normConcept.includes("margin") || normConcept.includes("%");
        const isFX = normConcept === "fx" || normConcept.includes("tasa de cambio") || normConcept === "tasa cambio" || normConcept === "tasa de cambio cierre" || normConcept === "tipo de cambio" || normConcept.includes("tasa proyectada");

        const isEbitdaMargin = normConcept.includes("ebitda");
        const isGrossMargin = normConcept.includes("bruto");
        const isNetMargin = normConcept.includes("neto") || normConcept.includes("utilidad neta") || normConcept.includes("resultado neto");
        const isGgadm = normConcept.includes("ggadm");

        const targetYearForAccum = getSortYear(data[endIdx]);
        // Calculate YTD (Year to Date) total from the first data point up to endIdx
        for (let k = 0; k <= endIdx; k++) {
            const periodData = data[k];
            if (getSortYear(periodData) !== targetYearForAccum) continue; 
            
            let matchingRows = periodData.pnl?.fullRows?.filter(r => r.concept === concept) || [];
            const val = matchingRows.reduce((sum, r) => sum + (r.values[periodData.date] || 0), 0);
            periodAccumTotal += val;
        }
        
        const parseDirtyNumberForMargin = (val) => {
            if (!val) return 0;
            if (typeof val === 'number') return val;
            let cleaned = val.toString().replace(/[^0-9.-]+/g, "");
            return Number(cleaned) || 0;
        };

        const periodCells = visibleMonths.map(period => {
            let matchingRows = period.pnl?.fullRows?.filter(r => r.concept === concept) || [];
            let val = matchingRows.reduce((sum, r) => sum + (r.values[period.date] || 0), 0);
            
            if (isPercentage) {
                const denRows = period.pnl?.fullRows?.filter(r => {
                    const nc = normalizeText(r.concept);
                    return nc === "ventas netas" || nc === "total ingresos" || nc === "ingresos" || nc.includes("ventas netas");
                }) || [];
                let denVal = denRows.length > 0 ? denRows.reduce((sum, r) => sum + (r.values[period.date] || 0), 0) : (period.kpis?.ingresos || 0);
                
                let numVal = 0;
                if (isEbitdaMargin) {
                    const numRow = period.pnl?.fullRows?.find(r => normalizeText(r.concept) === "ebitda" || normalizeText(r.concept).includes("ebitda ") || normalizeText(r.concept).includes(" ebitda"));
                    numVal = numRow ? numRow.values[period.date] || 0 : (period.kpis?.ebitda || 0);
                } else if (isGrossMargin) {
                    const numRow = period.pnl?.fullRows?.find(r => normalizeText(r.concept) === "margen bruto" || normalizeText(r.concept) === "utilidad bruta");
                    numVal = numRow ? numRow.values[period.date] || 0 : (period.kpis?.margen_bruto * period.kpis?.ingresos || 0);
                } else if (isNetMargin) {
                    const numRow = period.pnl?.fullRows?.find(r => normalizeText(r.concept) === "utilidad neta" || normalizeText(r.concept) === "ganancia del periodo" || normalizeText(r.concept) === "resultado neto");
                    numVal = numRow ? numRow.values[period.date] || 0 : (period.kpis?.utilidad || 0);
                } else if (isGgadm) {
                    const numRow = period.pnl?.fullRows?.find(r => normalizeText(r.concept) === "total ggadm" || normalizeText(r.concept).includes("gastos administrativos"));
                    numVal = numRow ? numRow.values[period.date] || 0 : 0;
                }

                numVal = parseDirtyNumberForMargin(numVal);
                denVal = parseDirtyNumberForMargin(denVal);

                if ((isEbitdaMargin || isGrossMargin || isNetMargin || isGgadm) && denVal !== 0) {
                    val = numVal / denVal;
                }
            }
            
            const color = val < 0 ? 'var(--danger)' : 'inherit';
            
            let displayVal;
            if (isPercentage) {
                displayVal = formatPercent(val);
            } else if (isFX) {
                displayVal = val.toFixed(2);
            } else {
                displayVal = formatCurrency(val);
            }

            // PPTO budget value for the same concept
            let pptoRows = period.ppto?.pnl?.fullRows || [];
            let pptoRow = pptoRows.find(r => r.concept === concept);
            let pptoVal = 0;
            if (pptoRow && pptoRow.values) {
                pptoVal = pptoRow.values[period.date] || 0;
            }

            let pulseClass = "";
            let pulseTitle = "";
            if (pptoVal && pptoVal !== 0) {
                const devPct = (val - pptoVal) / Math.abs(pptoVal);
                if (Math.abs(devPct) > 0.15) {
                    const conceptLower = concept.toLowerCase();
                    const isExpense = conceptLower.includes("costo") || conceptLower.includes("gasto") || conceptLower.includes("itbis") || conceptLower.includes("descuento") || conceptLower.includes("devolucion") || conceptLower.includes("devolución") || conceptLower.includes("d & a") || conceptLower.includes("depreciacion") || conceptLower.includes("amortizacion");
                    const isPositiveBetter = !isExpense;
                    const isBetter = isPositiveBetter ? (devPct > 0) : (devPct < 0);
                    pulseClass = isBetter ? "pulse-pos" : "pulse-neg";
                    
                    let formattedPptoVal;
                    if (isPercentage) {
                        formattedPptoVal = formatPercent(pptoVal);
                    } else if (isFX) {
                        formattedPptoVal = pptoVal.toFixed(2);
                    } else {
                        formattedPptoVal = formatCurrency(pptoVal);
                    }
                    pulseTitle = `Desviación de ${(devPct * 100).toFixed(1)}% respecto al presupuesto (${formattedPptoVal}) para ${concept}`;
                }
            }

            const innerAttributes = pulseClass ? `class="${pulseClass}" title="${pulseTitle}" style="display:inline-block; padding: 2px 6px;"` : "";

            return `<td style="text-align: right; color:${color};"><div ${innerAttributes}>${displayVal}</div></td>`;
        }).join('');

        const labelLower = concept.toLowerCase();
        const isTotal = labelLower.includes("total") || 
                        labelLower.includes("ebitda") || 
                        labelLower.includes("utilidad") ||
                        labelLower.includes("resultado") ||
                        labelLower.includes("ggadm") ||
                        labelLower.includes("ventas netas") ||
                        labelLower.includes("costo de venta") ||
                        labelLower.includes("ebit");
        
        const isSubRow = concept.startsWith("  ") || concept.startsWith("\t") || 
                         concept.toLowerCase().includes("costos ") || 
                         concept.toLowerCase().includes("gastos ") || 
                         concept.toLowerCase().includes("impuestos") || 
                         normConcept.includes("diferencial cambiario") || 
                         normConcept.includes("ingresos financieros") || 
                         normConcept.includes("gastos extraordinarios");
        const rowClass = isTotal ? 'row-total' : '';
        const cellClass = isSubRow ? 'row-indent' : '';

        // Acumulado del periodo
        let displayAccum;
        if (isPercentage) {
            const targetYear = getSortYear(data[endIdx]);
            let sumNum = 0;
            let sumDen = 0;
            
            const isEbitdaMargin = normConcept.includes("ebitda");
            const isGrossMargin = normConcept.includes("bruto");
            const isNetMargin = normConcept.includes("neto") || normConcept.includes("utilidad neta") || normConcept.includes("resultado neto");
            const isGgadm = normConcept.includes("ggadm");
            
            for (let k = endIdx; k >= 0; k--) {
                const item = data[k];
                if (getSortYear(item) !== targetYear) break;
                
                if (item.pnl && item.pnl.fullRows) {
                    const parseDirtyNumber = (val) => {
                        if (!val) return 0;
                        if (typeof val === 'number') return val;
                        let cleaned = val.toString().replace(/[^0-9.-]+/g, "");
                        return Number(cleaned) || 0;
                    };

                    const denRows = item.pnl.fullRows.filter(r => {
                        const nc = normalizeText(r.concept);
                        return nc === "ventas netas" || nc === "total ingresos" || nc === "ingresos" || nc.includes("ventas netas");
                    });
                    let denVal = denRows.length > 0 ? denRows.reduce((sum, r) => sum + (r.values[item.date] || 0), 0) : (item.kpis?.ingresos || 0);
                    
                    let numVal = 0;
                    if (isEbitdaMargin) {
                        const numRows = item.pnl.fullRows.filter(r => normalizeText(r.concept) === "ebitda" || normalizeText(r.concept).includes("ebitda ") || normalizeText(r.concept).includes(" ebitda"));
                        numVal = numRows.length > 0 ? numRows.reduce((sum, r) => sum + (r.values[item.date] || 0), 0) : (item.kpis?.ebitda || 0);
                    } else if (isGrossMargin) {
                        const numRows = item.pnl.fullRows.filter(r => normalizeText(r.concept) === "margen bruto" || normalizeText(r.concept) === "utilidad bruta");
                        numVal = numRows.length > 0 ? numRows.reduce((sum, r) => sum + (r.values[item.date] || 0), 0) : (item.kpis?.margen_bruto * item.kpis?.ingresos || 0);
                    } else if (isNetMargin) {
                        const numRows = item.pnl.fullRows.filter(r => normalizeText(r.concept) === "utilidad neta" || normalizeText(r.concept) === "ganancia del periodo" || normalizeText(r.concept) === "resultado neto");
                        numVal = numRows.length > 0 ? numRows.reduce((sum, r) => sum + (r.values[item.date] || 0), 0) : (item.kpis?.utilidad || 0);
                    } else if (isGgadm) {
                        const numRows = item.pnl.fullRows.filter(r => normalizeText(r.concept) === "total ggadm" || normalizeText(r.concept).includes("gastos administrativos"));
                        numVal = numRows.length > 0 ? numRows.reduce((sum, r) => sum + (r.values[item.date] || 0), 0) : 0;
                    }

                    numVal = parseDirtyNumber(numVal);
                    denVal = parseDirtyNumber(denVal);

                    if (isEbitdaMargin || isGrossMargin || isNetMargin || isGgadm) {
                        sumNum += numVal;
                        sumDen += denVal;
                    } else {
                        // For generic percentages not strictly defined, just fallback
                        const genRows = item.pnl.fullRows.filter(r => r.concept === concept);
                        let genVal = genRows.length > 0 ? genRows.reduce((sum, r) => sum + (r.values[item.date] || 0), 0) : 0;
                        sumNum += parseDirtyNumber(genVal); // Just sum the % if nothing else matches (or use last val as before)
                        sumDen = 0;
                    }
                }
            }
            
            if (sumDen !== 0 && sumNum !== 0) {
                displayAccum = formatPercent(sumNum / sumDen);
            } else {
                const lastItem = visibleMonths[visibleMonths.length - 1];
                const lastMatchingRows = lastItem.pnl?.fullRows?.filter(r => r.concept === concept) || [];
                const lastVal = lastMatchingRows.reduce((sum, r) => sum + (r.values[lastItem.date] || 0), 0);
                displayAccum = formatPercent(lastVal);
            }
        } else if (isFX) {
            let sumEbitdaLocal = 0;
            let sumEbitdaUsd = 0;

            const targetYear = getSortYear(data[endIdx]);
            
            // Función auxiliar para limpiar números sucios del Excel
            const parseDirtyNumber = (val) => {
                if (!val) return 0;
                if (typeof val === 'number') return val;
                let cleaned = val.toString().replace(/[^0-9.-]+/g, "");
                return Number(cleaned) || 0;
            };

            // Suma manual iterando los meses YTD
            for (let k = endIdx; k >= 0; k--) {
                const item = data[k];
                if (getSortYear(item) !== targetYear) break;
                
                let localVal = 0;
                let usdVal = 0;

                if (item.pnl && item.pnl.fullRows) {
                    // Acceso directo con las llaves exactas
                    const localRow = item.pnl.fullRows.find(r => r.concept === "EBITDA");
                    if (localRow) localVal = localRow.values[item.date] || 0;
                    if (localVal === 0 && item.kpis?.ebitda) localVal = item.kpis.ebitda;

                    const currentRateRow = item.pnl.fullRows.find(r => r.concept === concept);
                    const currentRate = currentRateRow ? parseDirtyNumber(currentRateRow.values[item.date]) : 0;
                    
                    if (currentRate !== 0) {
                        usdVal = localVal / currentRate;
                    } else {
                        const usdRow = item.pnl.fullRows.find(r => r.concept === "EBITDA US$");
                        if (usdRow) usdVal = usdRow.values[item.date] || 0;
                    }
                }

                sumEbitdaLocal += parseDirtyNumber(localVal);
                sumEbitdaUsd += parseDirtyNumber(usdVal);
            }

            // Cálculo final de la Tasa Implícita YTD
            let fxAcumulado = 0;
            if (sumEbitdaUsd !== 0 && sumEbitdaLocal !== 0) {
                fxAcumulado = sumEbitdaLocal / sumEbitdaUsd;
            }
            displayAccum = fxAcumulado > 0 ? fxAcumulado.toFixed(2) : '-';
        } else {
            displayAccum = formatCurrency(periodAccumTotal);
        }

        let displayConcept = concept;
        if (displayConcept.trim() === "Ventas P6") {
            displayConcept = "Ventas BON";
        }
        displayConcept = formatSegmentName(displayConcept);

        return `<tr class="${rowClass}">
            <td class="${cellClass}">${displayConcept}</td>
            ${periodCells}
            <td style="font-weight:700; background:rgba(0,0,0,0.02);">${displayAccum}</td>
        </tr>`;
    }).join('');
}

/**
 * 🚀 KPI DASHBOARD: Torre de Control
 */
function renderKPIDashboard(data, selectedIndex) {
    const curr = data[selectedIndex];
    if (!curr) return;

    const kpis = curr.kpis || { ingresos: 0, ebitda: 0, cashflow: 0 };
    const prev = selectedIndex > 0 ? data[selectedIndex - 1] : curr;
    const prevKpis = prev.kpis || kpis;

    const getPeriodStr = (idx) => {
        if (idx < 0 || !data[idx]) return "N/A";
        const item = data[idx];
        if (!isYTDMode || !item.sortDate) return item.date;
        try {
            const targetYear = getSortYear(item);
            let startItem = item;
            for (let i = idx; i >= 0; i--) {
                if (data[i].sortDate && getSortYear(data[i]) === targetYear) {
                    startItem = data[i];
                } else {
                    break;
                }
            }
            if (!startItem.date || typeof startItem.date !== 'string' || !item.date || typeof item.date !== 'string') return item.date;
            const startMonth = startItem.date.split(' ')[0];
            const endMonth = item.date.split(' ')[0];
            if (startMonth === endMonth) return item.date;
            return `${startMonth} a ${endMonth} ${targetYear}`;
        } catch(e) {
            return item.date;
        }
    };

    // 1. Update Score Cards
    const levValue = curr.balance.ebitdaLTM > 0 ? (curr.balance.deudaTotal / curr.balance.ebitdaLTM).toFixed(2) : "0.00";
    document.getElementById('dash-lev').textContent = levValue + 'x';

    const statusLev = document.getElementById('status-lev');
    if (statusLev) {
        const prevLevValue = prev.balance.ebitdaLTM > 0 ? (prev.balance.deudaTotal / prev.balance.ebitdaLTM).toFixed(2) : "0.00";
        const diffLev = parseFloat(levValue) - parseFloat(prevLevValue);
        if (Math.abs(diffLev) < 0.01) {
            statusLev.textContent = "Estable";
            statusLev.style.color = "var(--text-secondary)";
        } else if (diffLev < 0) {
            statusLev.textContent = "▲ Mejorando";
            statusLev.style.color = "var(--success)";
        } else {
            statusLev.textContent = "▼ Cayendo";
            statusLev.style.color = "var(--danger)";
        }
        const prevDateStr = getPeriodStr(selectedIndex - 1);
        statusLev.setAttribute('data-tooltip', `<strong>Previo (${prevDateStr}):</strong> <br/> ${prevLevValue}x`);
    }

    // Secondary CEO KPIs
    let utilidad = kpis.utilidad || 0;
    let margenNeto = kpis.margen_neto || 0;
    let margenBruto = kpis.margen_bruto || 0;
    
    let prevMargenNeto = prevKpis.margen_neto || 0;
    let prevMargenBruto = prevKpis.margen_bruto || 0;

    // Helper to calculate accumulated metrics for a given index up to the beginning of its year
    const getAccumulatedRatios = (idx) => {
        if (idx < 0) return null;
        const targetItem = data[idx];
        const targetYear = getSortYear(targetItem);
        let sumIngresos = 0, sumUtilidad = 0, sumGrossMargin = 0;
        
        for (let i = idx; i >= 0; i--) {
            const item = data[i];
            if (getSortYear(item) !== targetYear) break;
            sumIngresos += (item.kpis?.ingresos || 0);
            sumUtilidad += (item.kpis?.utilidad || 0);
            sumGrossMargin += (item.kpis?.margen_bruto || 0) * (item.kpis?.ingresos || 0);
        }
        
        const monthNumber = targetItem.sortDate.getMonth() + 1;
        return {
            utilidad: sumUtilidad,
            margen_bruto: sumIngresos !== 0 ? sumGrossMargin / sumIngresos : 0,
            margen_neto: sumIngresos !== 0 ? sumUtilidad / sumIngresos : 0,
            monthNumber: monthNumber
        };
    };

    let dispMonthNumber = 1;
    if (isYTDMode) {
        const currYTD = getAccumulatedRatios(selectedIndex);
        if (currYTD) {
            utilidad = currYTD.utilidad;
            margenNeto = currYTD.margen_neto;
            margenBruto = currYTD.margen_bruto;
            dispMonthNumber = currYTD.monthNumber;
        }
        
        const prevYTD = getAccumulatedRatios(selectedIndex - 1);
        if (prevYTD) {
            prevMargenNeto = prevYTD.margen_neto;
            prevMargenBruto = prevYTD.margen_bruto;
        } else {
            prevMargenNeto = margenNeto;
            prevMargenBruto = margenBruto;
        }
    }

    document.getElementById('dash-margen-neto').textContent = formatPercent(margenNeto);
    const margenBrutoEl = document.getElementById('dash-margen-bruto');
    if (margenBrutoEl) margenBrutoEl.textContent = formatPercent(margenBruto);

    // ROE (Utilidad LTM / Patrimonio) - Estimado
    const patrimonio = curr.balance.patrimonio > 0 ? curr.balance.patrimonio : 1; 
    const activos = curr.balance.activos > 0 ? curr.balance.activos : 1;
    
    // Si la utilidad y activos son > 0 lo mostramos. Si patrimonio = 0 (anomalía), no mostrar div by zero
    // En YTD mode, anualizamos la utilidad acumulada dividiéndola por el número de meses y multiplicando por 12
    const annualizedUtility = isYTDMode ? (utilidad / dispMonthNumber) * 12 : utilidad * 12;
    const roe = curr.balance.patrimonio !== 0 ? (annualizedUtility / curr.balance.patrimonio) : 0;
    const roa = curr.balance.activos !== 0 ? (annualizedUtility / curr.balance.activos) : 0;

    let prevUtilidad = prevKpis.utilidad || 0;
    let prevDispMonthNumber = 1;
    if (isYTDMode) {
        const prevYTD = getAccumulatedRatios(selectedIndex - 1);
        if (prevYTD) {
            prevUtilidad = prevYTD.utilidad;
            prevDispMonthNumber = prevYTD.monthNumber;
        }
    }
    const prevAnnualizedUtility = isYTDMode ? (prevUtilidad / prevDispMonthNumber) * 12 : prevUtilidad * 12;
    const prevPatrimonio = prev.balance.patrimonio > 0 ? prev.balance.patrimonio : 1;
    const prevActivos = prev.balance.activos > 0 ? prev.balance.activos : 1;
    const prevRoe = prev.balance.patrimonio !== 0 ? (prevAnnualizedUtility / prevPatrimonio) : 0;
    const prevRoa = prev.balance.activos !== 0 ? (prevAnnualizedUtility / prevActivos) : 0;
    
    document.getElementById('dash-roe').textContent = formatPercent(roe);
    document.getElementById('dash-roa').textContent = formatPercent(roa);

    // CCC = DSO + DIO - DPO
    const dso = curr.cashflowDetail?.dso || 0;
    const dio = curr.cashflowDetail?.dio || 0;
    const dpo = curr.cashflowDetail?.dpo || 0;
    const ccc = dso + dio - dpo;
    
    document.getElementById('dash-ccc').textContent = `${ccc.toFixed(0)} días`;
    
    const prevDso = prev.cashflowDetail?.dso || 0;
    const prevDio = prev.cashflowDetail?.dio || 0;
    const prevDpo = prev.cashflowDetail?.dpo || 0;
    const prevCcc = prevDso + prevDio - prevDpo;
    const cccDiff = prevCcc !== 0 ? ((ccc - prevCcc) / Math.abs(prevCcc)) * 100 : 0;
    const statusCcc = document.getElementById('status-ccc');
    if (statusCcc) {
        if (ccc === 0 && prevCcc === 0) {
            statusCcc.textContent = "Estable";
            statusCcc.style.color = "var(--text-secondary)";
        } else if (ccc < prevCcc) {
            statusCcc.textContent = "▲ Mejorando";
            statusCcc.style.color = "var(--success)";
        } else if (ccc > prevCcc) {
            statusCcc.textContent = "▼ Cayendo";
            statusCcc.style.color = "var(--danger)";
        } else {
            statusCcc.textContent = "Estable";
            statusCcc.style.color = "var(--text-secondary)";
        }
        const prevDateStr = getPeriodStr(selectedIndex - 1);
        statusCcc.setAttribute('data-tooltip', `<strong>Previo (${prevDateStr}):</strong> <br/> ${prevCcc.toFixed(0)} días`);
    }

    const updateBulletChart = (idPrefix, realValMonthly, pptoValMonthly, realYtd, pptoYtd) => {
        const valueId = `dash-${idPrefix}`;
        const barId = `bullet-val-${idPrefix}`;
        const targetId = `bullet-target-${idPrefix}`;
        const labelId = `bullet-label-${idPrefix}`;
        const targetTextId = `target-${idPrefix}`;

        const valueEl = document.getElementById(valueId);
        const barEl = document.getElementById(barId);
        const targetEl = document.getElementById(targetId);
        const labelEl = document.getElementById(labelId);
        const targetTextEl = document.getElementById(targetTextId);

        if (!valueEl || !barEl || !targetEl || !labelEl) return;

        const dispReal = isYTDMode ? realYtd : realValMonthly;
        const dispPpto = isYTDMode ? pptoYtd : pptoValMonthly;

        valueEl.textContent = formatCurrency(dispReal);
        if (targetTextEl) {
            targetTextEl.textContent = `Meta PPTO: ${formatCurrency(dispPpto)}`;
        }
        
        let pct = dispPpto !== 0 ? (dispReal / Math.abs(dispPpto)) * 100 : 0;
        if (dispPpto === 0 && dispReal > 0) pct = 100;
        else if (dispPpto === 0 && dispReal < 0) pct = 0;

        // Visual logic based on idea that Target is the 80% line.
        // Scale max to 125% of Target, or Actual.
        const maxVisualPct = Math.max(125, pct + 5); 
        const targetVisualPos = dispPpto !== 0 ? (100 / maxVisualPct) * 100 : 0;
        const barVisualPos = (Math.max(0, pct) / maxVisualPct) * 100;

        barEl.style.width = `${Math.min(100, barVisualPos)}%`;
        targetEl.style.left = `${Math.min(98, targetVisualPos)}%`;
        
        // Semantic color
        let color = '#2a9d8f';
        if (pct >= 100) color = '#2a9d8f';
        else color = '#e76f51';

        barEl.style.backgroundColor = color;
        labelEl.textContent = `${pct.toFixed(1)}% vs PPTO ${isYTDMode ? 'YTD' : 'Mes'}`;
        labelEl.style.color = (pct >= 100) ? 'var(--success)' : 'var(--danger)';
    };

    const ytdData = calculateYTD(data, selectedIndex);

    updateBulletChart('ingresos', kpis.ingresos, curr.ppto?.kpis?.ingresos || 0, ytdData.real.ingresos, ytdData.ppto.ingresos);
    updateBulletChart('ebitda', kpis.ebitda, curr.ppto?.kpis?.ebitda || 0, ytdData.real.ebitda, ytdData.ppto.ebitda);
    updateBulletChart('cash', kpis.cashflow, curr.ppto?.kpis?.cashflow || 0, ytdData.real.cashflow, ytdData.ppto.cashflow);

    const updateScoreTrend = (id, currVal, prevVal) => {
        const el = document.getElementById(id);
        if (!el) return;
        const diff = currVal - prevVal;
        const pct = prevVal !== 0 ? (diff / Math.abs(prevVal)) * 100 : 0;
        el.style.color = (diff >= 0 ? 'var(--success)' : 'var(--danger)');
        const timeLabel = isYTDMode ? 'año ant.' : 'mes ant.';
        el.textContent = `${diff >= 0 ? '▲' : '▼'} ${Math.abs(pct).toFixed(1)}% vs ${timeLabel}`;
    };

    updateBulletChart('utilidad', utilidad, curr.ppto?.kpis?.utilidad || 0, ytdData.real.utilidad, ytdData.ppto.utilidad);
    
    // update simple status for ratios
    const updateRatioStatus = (elId, diff, prevVal) => {
        const el = document.getElementById(elId);
        if(!el) return;
        el.textContent = diff >= 0 ? "▲ Mejorando" : "▼ Cayendo";
        el.style.color = diff >= 0 ? "var(--success)" : "var(--danger)";
        if (Math.abs(diff) < 0.001) {
            el.textContent = "Estable";
            el.style.color = "var(--text-secondary)";
        }
        if (prevVal !== undefined) {
            const prevDateStr = getPeriodStr(selectedIndex - 1);
            el.setAttribute('data-tooltip', `<strong>Previo (${prevDateStr}):</strong> <br/> ${formatPercent(prevVal)}`);
        }
    };
    
    updateRatioStatus('status-margen-neto', margenNeto - prevMargenNeto, prevMargenNeto);
    updateRatioStatus('status-margen-bruto', margenBruto - prevMargenBruto, prevMargenBruto);

    // ROE, ROA status
    const evaluateStatus = (elId, val, prevVal) => {
        const el = document.getElementById(elId);
        if(!el) return;
        if(val > 0.15) { el.textContent = "Óptimo"; el.style.color = "var(--success)"; }
        else if(val > 0) { el.textContent = "Adecuado"; el.style.color = "var(--info)"; }
        else if(val === 0) { el.textContent = "Insuficiente Data"; el.style.color = "var(--text-secondary)"; }
        else { el.textContent = "Bajo Cero (Atención)"; el.style.color = "var(--danger)"; }
        
        if (prevVal !== undefined) {
             const prevDateStr = getPeriodStr(selectedIndex - 1);
             el.setAttribute('data-tooltip', `<strong>Previo (${prevDateStr}):</strong> <br/> ${formatPercent(prevVal)}`);
        }
    };
    evaluateStatus('status-roe', roe, prevRoe);
    evaluateStatus('status-roa', roa, prevRoa);

    // -- Variación Interanual (YoY) --
    let yoyData = null;
    try {
        const prevYearValue = curr.sortDate && !isNaN(new Date(curr.sortDate)) ? new Date(curr.sortDate).getFullYear() - 1 : 2025;
        
        // 1. Try exact month match first
        yoyData = data.find(d => {
            if (d.sortDate && curr.sortDate) {
                const dDate = new Date(d.sortDate);
                const cDate = new Date(curr.sortDate);
                if (!isNaN(dDate) && !isNaN(cDate)) {
                    return dDate.getMonth() === cDate.getMonth() && 
                           dDate.getFullYear() === prevYearValue;
                }
            }
            return false;
        });

        // 2. Fallback: Any data from previous year
        if (!yoyData) {
            yoyData = data.find(d => {
                if (d.sortDate) {
                    const dDate = new Date(d.sortDate);
                    return !isNaN(dDate) && dDate.getFullYear() === prevYearValue;
                }
                const dNorm = normalizeText(d.date || "");
                return dNorm.includes(prevYearValue.toString()) || dNorm.includes(prevYearValue.toString().slice(-2));
            });
        }
    } catch (err) {
        console.warn("Could not find yoyData by date matching.", err);
    }
    
    if (!yoyData) {
        console.warn("Could not find yoyData fallback. Current date:", curr.date, "Available data:", data.map(d=>d.date));
        yoyData = selectedIndex >= 12 ? data[selectedIndex - 12] : null;
    }

    const calcYoY = (currValue, yoyItem, elPrefix) => {
        const valueEl = document.getElementById(`yoy-${elPrefix}`);
        const statusEl = document.getElementById(`yoy-status-${elPrefix}`);
        if (!yoyItem || (!isYTDMode && !yoyItem)) {
            if (valueEl) valueEl.textContent = "N/A";
            if (statusEl) {
                statusEl.textContent = "Sin datos año ant.";
                statusEl.style.color = "var(--text-secondary)";
            }
            return;
        }

        let finalCurrValue = currValue;
        let finalPrevValue = elPrefix === 'caja' 
            ? (yoyItem.kpis?.cashEnding || yoyItem.kpis?.cashflow || 0)
            : (elPrefix === 'utilidad' ? (yoyItem.kpis?.utilidad || 0) 
            : (yoyItem.kpis?.[elPrefix] || 0));

        if (isYTDMode) {
            const ytdKey = elPrefix === 'caja' ? 'cashflow' : elPrefix;
            finalCurrValue = ytdData.real[ytdKey] || 0;
            
            const yoyIndex = data.indexOf(yoyItem);
            const prevYtdData = yoyIndex >= 0 ? calculateYTD(data, yoyIndex) : null;
            finalPrevValue = prevYtdData ? (prevYtdData.real[ytdKey] || 0) : finalPrevValue;
        }
        
        const diff = finalCurrValue - finalPrevValue;
        const pct = finalPrevValue !== 0 ? (diff / Math.abs(finalPrevValue)) * 100 : (finalCurrValue !== 0 ? 100 : 0);
        
        if (valueEl) {
            valueEl.textContent = `${pct > 0 ? '+' : ''}${pct.toFixed(1)}%`;
            const yoyDateStr = yoyItem ? getPeriodStr(data.indexOf(yoyItem)) : 'año ant.';
            valueEl.setAttribute('data-tooltip', `<strong>Vs. ${yoyDateStr}:</strong> <br/> ${formatCurrency(finalPrevValue)}`);
            valueEl.removeAttribute('title');
            if (valueEl.parentElement) {
                valueEl.parentElement.setAttribute('data-tooltip', `<strong>Vs. ${yoyDateStr}:</strong> <br/> ${formatCurrency(finalPrevValue)}`);
                valueEl.parentElement.removeAttribute('title');
            }
        }
        if (statusEl) {
            if (pct >= 0.01) {
                statusEl.textContent = "▲ Mejorando";
                statusEl.style.color = "var(--success)";
            } else if (pct <= -0.01) {
                statusEl.textContent = "▼ Cayendo";
                statusEl.style.color = "var(--danger)";
            } else {
                statusEl.textContent = "Estable";
                statusEl.style.color = "var(--text-secondary)";
            }
        }
    };
    
    calcYoY(kpis.ingresos, yoyData, 'ingresos');
    calcYoY(kpis.ebitda, yoyData, 'ebitda');
    calcYoY(utilidad, yoyData, 'utilidad');
    const currentDisplayCash = kpis.cashEnding || kpis.cashflow;
    calcYoY(currentDisplayCash, yoyData, 'caja');
    // --------------------------------

    // 2. Render Sparklines using D3
    const renderSparkline = (containerId, values, color) => {
        const container = document.getElementById(containerId);
        if (!container) return;
        container.innerHTML = '';
        
        const width = container.clientWidth;
        const height = 40;
        const margin = { top: 2, right: 2, bottom: 2, left: 2 };

        const svg = d3.select(`#${containerId}`)
            .append("svg")
            .attr("width", width)
            .attr("height", height);

        const x = d3.scaleLinear()
            .domain([0, Math.max(1, values.length - 1)])
            .range([margin.left, width - margin.right]);

        const y = d3.scaleLinear()
            .domain([d3.min(values) || 0, d3.max(values) || 0])
            .range([height - margin.bottom, margin.top]);

        const line = d3.line()
            .x((d, i) => x(i))
            .y(d => y(d))
            .curve(d3.curveBasis);

        svg.append("path")
            .datum(values)
            .attr("fill", "none")
            .attr("stroke", color)
            .attr("stroke-width", 2)
            .attr("d", line);
    };

    // Filtro: No mostrar 2025 en el Dashboard (Gráficos)
    const rollingData = data.slice(Math.max(0, selectedIndex - 11), selectedIndex + 1).filter(d => isYear2026(d));
    renderSparkline('spark-ingresos', rollingData.map(d => d.kpis.ingresos), 'var(--success)');
    renderSparkline('spark-ebitda', rollingData.map(d => d.kpis.ebitda), 'var(--primary)');
    renderSparkline('spark-cash', rollingData.map(d => d.kpis.cashflow), 'var(--info)');

    // 3. Main Trend Charts
    requestAnimationFrame(() => {
        renderMarginChart(rollingData);
        
        requestAnimationFrame(() => {
            renderCashFlowChart(rollingData);
            
            requestAnimationFrame(() => {
                // 4. Alerts
                renderDashboardAlerts(curr, data, selectedIndex);

                // 5. Covenants Container & Gauges
                let covenantsContainer = document.getElementById('covenantsContainer');
                if (!covenantsContainer) {
                    covenantsContainer = document.createElement('div');
                    covenantsContainer.id = 'covenantsContainer';
                    covenantsContainer.style.display = 'flex';
                    covenantsContainer.style.flexDirection = 'row';
                    covenantsContainer.style.flexWrap = 'wrap';
                    covenantsContainer.style.marginTop = '10px';
                    covenantsContainer.style.marginBottom = '20px';
                    covenantsContainer.style.gap = '15px';
                    
                    const alertsSection = document.getElementById('dashboard-alerts-section');
                    if (alertsSection) {
                        alertsSection.parentNode.insertBefore(covenantsContainer, alertsSection);
                    }
                }
                renderCovenantGauges(data, selectedIndex);
            });
        });
    });

    // -- AI Executive Summary Injection --
    let aiContainer = document.getElementById('aiSummaryContainer');
    if (!aiContainer) {
        if(!document.getElementById('ai-summary-styles')) {
            const style = document.createElement('style');
            style.id = 'ai-summary-styles';
            style.innerHTML = `
                .ai-summary-container {
                    margin-top: 24px;
                    margin-bottom: 24px;
                    display: flex;
                    flex-direction: column;
                    gap: 12px;
                }
                .ai-button {
                    background-color: #0f172a;
                    color: #ffffff;
                    border: none;
                    padding: 12px 24px;
                    border-radius: 12px;
                    font-size: 0.95rem;
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.2s ease;
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    width: fit-content;
                    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                }
                .ai-button:hover {
                    background-color: #1e293b;
                    transform: translateY(-1px);
                }
                .ai-button:active {
                    transform: translateY(0);
                }
                .ai-summary-box {
                    background-color: #f8fafc;
                    border-left: 4px solid #6366f1;
                    padding: 20px;
                    border-radius: 8px;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                    font-size: 0.9rem;
                    line-height: 1.6;
                    color: #334155;
                    animation: fadeIn 0.3s ease;
                }
                .ai-summary-box h3 {
                    margin-top: 0;
                    color: #1e293b;
                    font-size: 1rem;
                    margin-bottom: 12px;
                }
                .ai-summary-box ul {
                    margin-left: 20px;
                    margin-bottom: 0;
                }
                .ai-summary-box li {
                    margin-bottom: 8px;
                }
                .ai-summary-box li:last-child {
                    margin-bottom: 0;
                }
                .ai-summary-box table {
                    width: 100%;
                    border-collapse: collapse;
                    margin: 16px 0;
                }
                .ai-summary-box th, .ai-summary-box td {
                    border: 1px solid #cbd5e1;
                    padding: 8px 12px;
                    text-align: left;
                }
                .ai-summary-box th {
                    background-color: #f1f5f9;
                    font-weight: 700;
                    color: #1e293b;
                }
            `;
            document.head.appendChild(style);
        }

        aiContainer = document.createElement('div');
        aiContainer.id = 'aiSummaryContainer';
        aiContainer.className = 'ai-summary-container';
        
        const btn = document.createElement('button');
        btn.id = 'btnGenerateAI';
        btn.className = 'ai-button';
        btn.innerHTML = '✨ Generar Resumen Ejecutivo del Mes';
        
        const box = document.createElement('div');
        box.id = 'aiSummaryBox';
        box.className = 'ai-summary-box';
        box.style.display = 'none';
        
        aiContainer.appendChild(btn);
        aiContainer.appendChild(box);
        
        const alertsSection = document.getElementById('dashboard-alerts-section');
        if (alertsSection && alertsSection.parentNode) {
            alertsSection.parentNode.insertBefore(aiContainer, alertsSection.nextSibling);
        }
    }
    
    const aiSummaryBox = document.getElementById('aiSummaryBox');
    if (aiSummaryBox) {
        aiSummaryBox.style.display = 'none';
        aiSummaryBox.innerHTML = '';
    }
    
    let btnGenerateAI = document.getElementById('btnGenerateAI');
    if (btnGenerateAI) {
        const newBtn = btnGenerateAI.cloneNode(true);
        btnGenerateAI.parentNode.replaceChild(newBtn, btnGenerateAI);
        
        newBtn.addEventListener('click', () => {
            generateExecutiveSummary(data, selectedIndex);
        });
    }
}

function renderMarginChart(originalRollingData) {
    const marginContainer = document.getElementById('marginChart');
    if (!marginContainer) return;
    
    const parentView = marginContainer.closest('.view-container');
    if (parentView && window.getComputedStyle(parentView).display === 'none') {
        return;
    }

    d3.select('#marginChart').selectAll('*').remove();

    const isMobile = window.innerWidth < 768;
    const rollingData = isMobile ? originalRollingData.slice(-3) : originalRollingData;

    const width = marginContainer.clientWidth;
    const height = 250;
    const margin = isMobile 
        ? { top: 20, right: 15, bottom: 35, left: 35 } 
        : { top: 20, right: 30, bottom: 40, left: 50 };

    // Tooltip
    let tooltip = d3.select("body").select(".d3-tooltip");
    if (tooltip.empty()) {
        tooltip = d3.select("body").append("div").attr("class", "d3-tooltip");
    }

    const svg = d3.select("#marginChart")
        .append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const data = rollingData.map(d => ({
        date: d.date,
        margin: (d.kpis.margen_ebitda || 0) * 100,
        ebitda: d.kpis.ebitda || 0
    }));

    const x = d3.scalePoint()
        .domain(data.map(d => d.date))
        .range([0, width - margin.left - margin.right]);

    const y = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.margin) * 1.2])
        .range([height - margin.top - margin.bottom, 0]);

    svg.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).tickSize(0).tickPadding(10))
        .selectAll("text")
        .style("font-size", isMobile ? "8px" : "10px")
        .style("color", "var(--text-secondary)");

    svg.append("g")
        .call(d3.axisLeft(y).ticks(5).tickFormat(d => d + "%"))
        .style("font-size", isMobile ? "8px" : "10px");

    const line = d3.line()
        .x(d => x(d.date))
        .y(d => y(d.margin))
        .curve(d3.curveMonotoneX);

    svg.append("path")
        .datum(data)
        .attr("fill", "none")
        .attr("stroke", "var(--primary)")
        .attr("stroke-width", 3)
        .attr("d", line);

    svg.selectAll(".dot")
        .data(data)
        .enter().append("circle")
        .attr("class", "dot")
        .attr("cx", d => x(d.date))
        .attr("cy", d => y(d.margin))
        .attr("r", 5)
        .attr("fill", "white")
        .attr("stroke", "var(--primary)")
        .attr("stroke-width", 2)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
            d3.select(this).attr("r", 8).attr("fill", "var(--primary)");
            tooltip.style("opacity", 1)
                   .html(`<strong>${d.date}</strong><br/>Margen EBITDA: ${d.margin.toFixed(1)}%<br/>Monto EBITDA: ${formatCurrency(d.ebitda)}`);
        })
        .on("mousemove", function(event) {
            tooltip.style("left", (event.pageX + 10) + "px")
                   .style("top", (event.pageY - 28) + "px");
        })
        .on("mouseout", function() {
            d3.select(this).attr("r", 5).attr("fill", "white");
            tooltip.style("opacity", 0);
        });
}

function renderCashFlowChart(originalRollingData) {
    const cashContainer = document.getElementById('cashFlowChart');
    if (!cashContainer) return;

    const parentView = cashContainer.closest('.view-container');
    if (parentView && window.getComputedStyle(parentView).display === 'none') {
        return;
    }

    d3.select('#cashFlowChart').selectAll('*').remove();

    const isMobile = window.innerWidth < 768;
    const rollingData = isMobile ? originalRollingData.slice(-3) : originalRollingData;

    const width = cashContainer.clientWidth;
    const height = 250;
    const margin = isMobile 
        ? { top: 20, right: 15, bottom: 35, left: 35 } 
        : { top: 20, right: 30, bottom: 40, left: 60 };

    // Tooltip
    let tooltip = d3.select("body").select(".d3-tooltip");
    if (tooltip.empty()) {
        tooltip = d3.select("body").append("div").attr("class", "d3-tooltip");
    }

    const svg = d3.select("#cashFlowChart")
        .append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand()
        .domain(rollingData.map(d => d.date))
        .range([0, width - margin.left - margin.right])
        .padding(0.3);

    const y = d3.scaleLinear()
        .domain([d3.min(rollingData, d => d.kpis.cashflow), d3.max(rollingData, d => d.kpis.cashflow) * 1.1])
        .range([height - margin.top - margin.bottom, 0]);

    svg.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).tickSize(0).tickPadding(10))
        .selectAll("text")
        .style("font-size", isMobile ? "8px" : "10px");

    svg.append("g")
        .call(d3.axisLeft(y).ticks(5).tickFormat(d => {
            if (d === 0) return "0";
            if (Math.abs(d) >= 1000000) return (d / 1000000).toFixed(1).replace(/\.0$/, '') + "M";
            if (Math.abs(d) >= 1000) return (d / 1000).toFixed(0) + "K";
            return d.toString();
        }))
        .style("font-size", isMobile ? "8px" : "10px");

    svg.selectAll(".bar")
        .data(rollingData)
        .enter().append("rect")
        .attr("class", "bar")
        .attr("x", d => x(d.date))
        .attr("width", x.bandwidth())
        .attr("y", d => y(Math.max(0, d.kpis.cashflow)))
        .attr("height", d => Math.abs(y(d.kpis.cashflow) - y(0)))
        .attr("fill", d => d.kpis.cashflow >= 0 ? "var(--success)" : "var(--danger)")
        .attr("rx", 4)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
            d3.select(this).style("opacity", 0.8);
            tooltip.style("opacity", 1)
                   .html(`<strong>${d.date}</strong><br/>Flujo de Caja: ${formatCurrency(d.kpis.cashflow)}`);
        })
        .on("mousemove", function(event) {
            tooltip.style("left", (event.pageX + 10) + "px")
                   .style("top", (event.pageY - 28) + "px");
        })
        .on("mouseout", function() {
            d3.select(this).style("opacity", 1);
            tooltip.style("opacity", 0);
        });
}

window.aiAlertsCache = window.aiAlertsCache || {};

async function renderDashboardAlerts(curr, globalData, selectedIndex) {
    const container = document.getElementById('alertsContainer');
    if (!container) return;
    
    // Función de renderizado fallback o estático
    const renderStaticAlerts = () => {
        const alerts = [];
        const kpis = curr.kpis;
        const margin = kpis.margen_ebitda * 100;

        if (margin < 15) {
            alerts.push({ type: 'warning', text: `Margen EBITDA bajo (${margin.toFixed(1)}%). Se recomienda revisar eficiencia operativa.` });
        }
        
        if (curr.integrity && curr.integrity.isBroken) {
            alerts.push({ type: 'danger', text: "Descuadre detectable en la integridad del P&L. Verifique los costos directos." });
        }

        if (curr.balance.activos !== 0 && (curr.balance.activos < curr.balance.pasivos)) {
            alerts.push({ type: 'danger', text: "Patrimonio Negativo detectado. Riesgo de insolvencia técnica." });
        }

        if (alerts.length === 0) {
            container.innerHTML = '<div class="alert-card alert-success">No se detectan anomalías financieras críticas en este periodo.</div>';
        } else {
            container.innerHTML = alerts.map(a => `
                <div class="alert-card alert-${a.type}">
                    <i data-lucide="${a.type === 'danger' ? 'alert-octagon' : 'alert-triangle'}"></i>
                    <span>${a.text}</span>
                </div>
            `).join('');
        }
        if (typeof lucide !== 'undefined') lucide.createIcons();
    };

    if (!window.aiEnabled || !globalData || selectedIndex === undefined) {
        renderStaticAlerts();
        return;
    }

    const mesKey = curr.mes || `mes-${selectedIndex}`;
    
    if (window.aiAlertsCache[mesKey]) {
        container.innerHTML = window.aiAlertsCache[mesKey];
        if (typeof lucide !== 'undefined') lucide.createIcons();
        return;
    }

    container.innerHTML = `
        <div class="alert-card alert-warning" style="justify-content: center; background: transparent; border: none; box-shadow: none;">
            <i data-lucide="loader-2" class="spin-icon"></i>
            <span>Analizando historial de datos y anomalías con IA...</span>
        </div>
    `;
    if (typeof lucide !== 'undefined') lucide.createIcons();

    // Prepare historical contextual data for the AI
    let historicalData = [];
    const startIndex = Math.max(0, selectedIndex - 3);
    for(let i = startIndex; i <= selectedIndex; i++) {
        historicalData.push({
            mes: globalData[i].mes,
            ingresos: globalData[i].kpis.ingresos,
            ebitda: globalData[i].kpis.ebitda,
            margen_ebitda: globalData[i].kpis.margen_ebitda,
            activos: globalData[i].balance.activos,
            pasivos: globalData[i].balance.pasivos
        });
    }

    try {
        let timeoutId;
        const timeoutPromise = new Promise((_, reject) => {
            timeoutId = setTimeout(() => reject(new Error('AI Request Timeout (45s)')), 45000);
        });

        const promptText = `Actúa como un Auditor y Analista Financiero Senior. Revisa los siguientes datos financieros históricos (del mes actual: ${curr.mes} y últimos meses) y detecta anomalías, riesgos, o desviaciones significativas en las tendencias.
        
        INSTRUCCIONES:
        1. Devuelve un JSON estrictamente válido con un arreglo de objetos. Cada objeto debe tener:
           - "type": "danger" (problema crítico), "warning" (advertencia), o "success" (mejora notable).
           - "text": Descripción concisa e incisiva de la anomalía enfocada en las tendencias (máximo 2 oraciones).
        2. El análisis debe estar en español y mostrar tu razonamiento de impacto financiero.
        3. Fíjate en caídas abruptas de ingresos, márgenes, o incrementos inusuales en pasivos a lo largo del tiempo que podrían requerir atención inmediata de la gerencia. No repitas descripciones, consolida la información para que sea legible y de alto nivel ejecutivo.
        4. No inventes alertas si los datos no cambian de manera significativa. Si no hay anomalías detectables o los datos parecen muy planos, devuelve un array vacío [].
        
        DATOS HISTÓRICOS (Últimos meses):
        ${JSON.stringify(historicalData, null, 2)}`;

        let apiCallPromise;
        try {
            apiCallPromise = getAI().models.generateContent({
                model: "gemini-2.5-flash",
                contents: promptText,
                config: {
                    responseMimeType: "application/json"
                }
            });
            apiCallPromise.catch(err => window.handleAiError("Alerts", err));
        } catch (err) {
            apiCallPromise = Promise.reject(err);
            apiCallPromise.catch(()=> /* handled */ {});
        }

        let response;
        try {
            response = await Promise.race([apiCallPromise, timeoutPromise]);
        } finally {
            clearTimeout(timeoutId);
        }
        
        let text = response.text;
        text = text.replace(/```json/g, '').replace(/```/g, '').trim();
        
        let alertsData = [];
        try {
            alertsData = JSON.parse(text);
        } catch (e) {
            window.handleAiError("Alerts Parse", e);
            alertsData = [];
        }

        if (!Array.isArray(alertsData) || alertsData.length === 0) {
            window.aiAlertsCache[mesKey] = '<div class="alert-card alert-success" style="background: rgba(42, 157, 143, 0.1); color: var(--success);"><i data-lucide="check-circle" style="color: var(--success);"></i><span>No se detectan anomalías financieras críticas o desviaciones inusuales para este periodo frente al historial reciente.</span></div>';
            renderStaticAlerts(); // Fallback to basic static alerts if AI finds nothing
            if (!container.innerHTML.includes('alert-success')) {
                window.aiAlertsCache[mesKey] = container.innerHTML;
            } else {
                container.innerHTML = window.aiAlertsCache[mesKey];
            }
        } else {
            const html = alertsData.map(a => {
                let icon = 'info';
                let iconColor = 'var(--text-primary)';
                if (a.type === 'danger') { icon = 'alert-octagon'; iconColor = 'var(--danger)'; }
                else if (a.type === 'warning') { icon = 'alert-triangle'; iconColor = 'var(--warning)'; }
                else if (a.type === 'success') { icon = 'trending-up'; iconColor = 'var(--success)'; }
                
                return `
                    <div class="alert-card alert-${a.type || 'warning'}" style="border-left: 4px solid ${iconColor};">
                        <i data-lucide="${icon}" style="color: ${iconColor};"></i>
                        <span><strong>IA AI-Detect:</strong> ${a.text}</span>
                    </div>
                `;
            }).join('');
            window.aiAlertsCache[mesKey] = html;
            container.innerHTML = window.aiAlertsCache[mesKey];
        }
        
        if (typeof lucide !== 'undefined') lucide.createIcons();

    } catch (err) {
        window.handleAiError("Alerts", err);
        renderStaticAlerts();
    }
}

function updateTrend(id, curr, prev, ppto = null, suffix = "") {
    const el = document.getElementById(id);
    if (!el) return;
    const diff = curr - prev;
    const pct = prev !== 0 ? (diff / Math.abs(prev)) * 100 : 0;
    
    const timeLabel = isYTDMode ? "año ant." : "mes ant.";
    
    let html = '';
    if (diff >= 0.01) {
        html = `<span style="color:var(--success)">▲ ${pct.toFixed(1)}%</span> vs ${timeLabel}`;
    } else if (diff <= -0.01) {
        html = `<span style="color:var(--danger)">▼ ${Math.abs(pct).toFixed(1)}%</span> vs ${timeLabel}`;
    } else {
        html = `Sin cambios vs ${timeLabel}`;
    }

    if (ppto !== null && ppto !== 0) {
        const diffPpto = curr - ppto;
        const pctPpto = (diffPpto / Math.abs(ppto)) * 100;
        if (diffPpto >= 0.01) {
            html += ` | <span style="color:var(--success)">▲ ${pctPpto.toFixed(1)}%</span> vs PPTO`;
        } else if (diffPpto <= -0.01) {
            html += ` | <span style="color:var(--danger)">▼ ${Math.abs(pctPpto).toFixed(1)}%</span> vs PPTO`;
        } else {
            html += ` | En PPTO`;
        }
    }
    
    el.innerHTML = html + suffix;
}

/**
 * Render Estados Financieros based on wide format
 */
function renderEstadosFinancieros(data, selectedIndex = -1) {
    console.log("-> renderEstadosFinancieros executing");
    const headerEl = document.getElementById('header-estados');
    const bodyEl = document.getElementById('body-estados');
    if (!headerEl || !bodyEl) return;
    
    // Nueva validación de fuente: 
    if (!data || data.length === 0 || !data[0].kpis) {
        bodyEl.innerHTML = '<tr><td colspan="100%" style="text-align:center; padding: 20px; color: var(--text-secondary);">Por favor, sincronice el Master Financiero para ver esta sección</td></tr>';
        return;
    }

    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    // Show up to 12 months including the selected one
    const startIdx = Math.max(0, endIdx - 11);
    
    // We do NOT want to show full 12 always if data doesn't have it, but slice will handle that
    const visibleMonths = data.slice(startIdx, endIdx + 1).filter(d => isYear2026(d));
    const periods = visibleMonths.map(d => d.date);

    // The financial engine correctly scales values, so we don't need a multiplier hack here
    const applyMultiplier = 1;

    const formatLocalMillions = (v) => {
        if (v === 0 || !v) return "-";
        const formatted = Math.abs(v).toLocaleString('en-US', {
            minimumFractionDigits: 1,
            maximumFractionDigits: 1
        });
        return `${v < 0 ? '-' : ''}${formatted}`;
    };

    // Header
    headerEl.innerHTML = `
        <tr>
            <th style="width: 250px;">Concepto</th>
            ${periods.map(p => `<th style="text-align: right;">${p}</th>`).join('')}
            <th style="text-align: right; background: #f0f9ff; color: #0369a1;">Acumulado YTD</th>
        </tr>
    `;

    // 1. Gather all concepts present in the data to ensure we don't miss any
    let allDataConcepts = [];
    visibleMonths.forEach(d => {
        const sourceRows = (d.estados && d.estados.fullRows && d.estados.fullRows.length > 0) ? d.estados.fullRows : (d.pnl && d.pnl.fullRows ? d.pnl.fullRows : []);
        if (sourceRows) {
            sourceRows.forEach(row => {
                const norm = normalizeText(row.concept);
                const stringsToHide = ['otras ventas', 'otros ingresos'];
                if (norm !== "x" && norm !== "año" && norm !== "mes" && norm !== "columna" && norm !== "(dop)" && norm !== "diferencial cambiario por operaciones" && norm !== "diferencial cambiario por deuda" && !stringsToHide.includes(norm)) {
                    if (!allDataConcepts.includes(row.concept)) allDataConcepts.push(row.concept);
                }
            });
        }
    });

    // 2. Define the explicit requested structure
    const EXPLICIT_STRUCTURE = [
        { label: "Ventas Netas", type: "bold", dataKey: ["Ventas Netas", "Ventas totales", "Ingresos"] },
        { label: "Ventas BT5", type: "indent" },
        { label: "Ventas EVP", type: "indent" },
        { label: "Ventas P6", type: "indent" },
        { label: "Descuentos", type: "indent", dataKey: ["Descuentos", "Descuento", "Descuento sobre ventas", "Descuento en ventas", "Menos descuentos", "Menos descuentos y devoluciones"] },
        { label: "Devoluciones", type: "indent", dataKey: ["Devoluciones", "Devolucion", "Devoluciones sobre ventas", "Devolución", "Devolución en ventas", "Menos devoluciones", "Descuentos y devoluciones"] },
        { label: "Costo de Ventas", type: "bold" },
        { label: "Costo de Ventas", type: "indent" },
        { label: "Costo de Ventas Otros Ingresos", type: "indent" },
        { label: "Utilidad Bruta", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "GGADM", type: "normal" },
        { label: "Gastos de Personal", type: "indent" },
        { label: "Seguros", type: "indent" },
        { label: "Servicios Básicos", type: "indent", dataKey: ["Servicios Básicos", "Servicios Basicos"] },
        { label: "Combustibles", type: "indent" },
        { label: "Otros Gastos", type: "indent" },
        { label: "ITBIS", type: "indent" },
        { label: "Mercadeo y Ventas", type: "indent" },
        { label: "Honorarios Profesionales", type: "indent" },
        { label: "Alquiler", type: "indent" },
        { label: "Mantenimiento y Reparación", type: "indent", dataKey: ["Mantenimiento y Reparación", "Mantenimiento y Reparacion"] },
        { label: "Otros Gastos Operativos", type: "indent" },
        { label: "EBITDA", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "Depreciación y Amortización", type: "normal", dataKey: ["Depreciación y Amortización", "Depreciacion y Amortizacion"] },
        { label: "Depreciación y Amortización Gasto", type: "indent", dataKey: ["Depreciación y Amortización Gasto", "Depreciacion y Amortizacion Gasto"] },
        { label: "Depreciación y Amortización Costo", type: "indent", dataKey: ["Depreciación y Amortización Costo", "Depreciacion y Amortizacion Costo"] },
        { label: "EBIT", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "Ingreso(gasto) de Interés", type: "italic", dataKey: ["Ingreso(gasto) de Interés", "Ingreso(gasto) de interes", "Ingreso (gasto) de Interés", "Ingreso (gasto) de Interes"] },
        { label: "Ingresos Financieros", type: "indent" },
        { label: "Gastos Financieros", type: "indent" },
        { label: "Ingreso (gasto) de Interés", type: "normal", borderT: "1px dashed var(--text-secondary)", dataKey: ["Ingreso (gasto) de Interés", "Ingreso (gasto) de Interes", "Ingreso(gasto) de Interés", "Ingreso(gasto) de Interes"] },
        { label: "Diferencial Cambiario", type: "normal" },
        { label: "Gastos Extraordinarios", type: "normal" },
        { label: "Ingreso Antes de Impuestos", type: "normal", borderT: "1px dashed var(--text-secondary)" },
        { label: "Impuestos", type: "normal", dataKey: ["Impuesto Sobre la Renta", "Taxes", "Impuestos"] },
        { label: "Beneficio Neto", type: "bold", borderT: "2px solid #000", dataKey: ["Beneficio Neto", "Beneficio Neto del Periodo"] },
        
        { label: "Empty_1", type: "empty" },
        { label: "EBITDA USD", type: "bold" },
        { label: "Empty_2", type: "empty" },
        
        { label: "KPIs y Drivers", type: "category", subtitle: "(DOP)" },
        { label: "Análisis Horizontal", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "Crecimiento Ventas", type: "ratio" },
        { label: "Crecimiento EBITDA DOP", type: "ratio" },
        { label: "Crecimiento EBITDA USD", type: "ratio" },
        { label: "Crecimiento Beneficio Neto", type: "ratio" },

        { label: "Análisis Vertical", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "Costo de Ventas / Ventas", type: "ratio" },
        { label: "GGADM / Ventas", type: "ratio" },
        { label: "D&A / Ventas", type: "ratio" },
        { label: "CAPEX / Ventas", type: "ratio" },

        { label: "Análisis Margen", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "Gross margin", type: "ratio", dataKey: ["Gross margin", "Gross Margin"] },
        { label: "EBITDA margin", type: "ratio", dataKey: ["EBITDA margin", "EBITDA Margin"] },
        { label: "EBIT margin", type: "ratio", dataKey: ["EBIT margin", "EBIT Margin"] },
        { label: "Margen Neto", type: "ratio", dataKey: ["Margen Neto", "Margen neto"] },

        { label: "Rentabilidad", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "ROIC", type: "ratio" },
        { label: "ROE", type: "ratio" },
        { label: "ROA", type: "ratio" },
        { label: "Ingreso Interes / (Efectivo + CDs)", type: "ratio", dataKey: ["Ingreso Interes / (Efectivo + CDs)", "Ingreso Interés / (Efectivo + CDs)"] },

        { label: "Variables Macro", type: "bold", borderT: "1px dashed var(--text-secondary)" },
        { label: "Tasa de cierre USD", type: "decimal", dataKey: ["Tasa de cierre USD", "FX EOP", "Tasa Cambio Cierre", "FX"] }
    ];

    // Compute occurrences to handle duplicate labels correctly (e.g., "Costo de Ventas" twice)
    const labelOccurrences = {};
    const structuredItems = EXPLICIT_STRUCTURE.map(item => {
        const keys = item.dataKey ? (Array.isArray(item.dataKey) ? item.dataKey.map(k => normalizeText(k)) : [normalizeText(item.dataKey)]) : [normalizeText(item.label)];
        const primaryKey = keys[0];
        if (!labelOccurrences[primaryKey]) labelOccurrences[primaryKey] = 0;
        const occIndex = labelOccurrences[primaryKey];
        if (item.type !== "empty" && item.type !== "category_main" && item.type !== "category") {
            labelOccurrences[primaryKey]++;
        }
        return { ...item, matchKeys: keys, occIndex };
    });

    const isRatio = (type) => type === "ratio";
    const isDecimal = (type) => type === "decimal";

    let tbBody = '';
    const processedConcepts = new Set();

    structuredItems.forEach(item => {
        if (item.type === "empty") {
            tbBody += `<tr><td colspan="${periods.length + 2}" style="height: 24px;"></td></tr>`;
            return;
        }

        let rowBgColor = '';
        let cellStyle = 'color: var(--text-primary);';
        let commonTdStyle = '';
        
        if (item.borderT) {
            commonTdStyle += `border-top: ${item.borderT}; `;
        }
        
        if (item.type === "category_main") {
            rowBgColor = 'background: rgb(132,159,186);';
            cellStyle = 'color: white; font-weight: 700; font-size: 1.1em;';
        } else if (item.type === "category") {
            rowBgColor = 'background: #e0f2fe;';
            cellStyle = 'color: #0369a1; font-weight: 800; font-size: 1.1em;';
        } else if (item.type === "bold") {
            cellStyle += 'font-weight: 700;';
        } else if (item.type === "italic") {
            cellStyle += 'font-style: italic;';
        } else if (item.type === "indent") {
            cellStyle += 'padding-left: 24px; font-weight: 500;';
        } else {
            cellStyle += 'font-weight: 500;';
        }

        let labelHtml = formatSegmentName(item.label);
        if (item.subtitle) {
            labelHtml += `<div style="font-size: 0.75rem; font-weight: 600; color:var(--text-secondary); margin-top:2px;">${item.subtitle}</div>`;
        }
        
        let firstStyle = `${cellStyle} ${commonTdStyle}`;
        if (item.type === "category_main") {
            firstStyle += ` background: rgb(132,159,186) !important; color: white !important;`;
        } else if (item.type === "category") {
            firstStyle += ` background: #e0f2fe !important; color: #0369a1 !important;`;
        } else {
            firstStyle += ` background: var(--card, #ffffff) !important;`;
        }
        let rowHtml = `<td style="${firstStyle}">${labelHtml}</td>`;
        
        let total = 0;
        let isTotalizable = !isRatio(item.type) && !isDecimal(item.type) && item.type !== "category_main" && item.type !== "category";
        let anyVal = false;

        periods.forEach(p => {
            let val = 0;
            const periodData = visibleMonths.find(d => d.date === p);
            const sourceRows = (periodData && periodData.estados && periodData.estados.fullRows && periodData.estados.fullRows.length > 0) 
                 ? periodData.estados.fullRows 
                 : (periodData && periodData.pnl && periodData.pnl.fullRows ? periodData.pnl.fullRows : []);
            
            if (sourceRows.length > 0 && item.type !== "category_main" && item.type !== "category") {
                // Find all matches for this key
                const matches = sourceRows.filter(r => item.matchKeys.includes(normalizeText(r.concept)));
                if (matches[item.occIndex]) {
                    const matchedRow = matches[item.occIndex];
                    processedConcepts.add(matchedRow.concept);
                    if (matchedRow.values[p] !== undefined) {
                        val = matchedRow.values[p];
                        if (!isRatio(item.type) && !isDecimal(item.type)) {
                            val = val * applyMultiplier;
                        }
                        if (val !== 0) anyVal = true;
                    }
                }
            }

            // Fallback robusto para campos calculados si no se extrajeron bien
            if (!anyVal && val === 0 && item.type !== "category_main" && item.type !== "category" && item.type !== "empty") {
                const normLabel = normalizeText(item.label);
                if (normLabel === "utilidad bruta" || normLabel === "margen bruto") {
                    const vNet = periodData?.kpis?.ingresos || 0;
                    const cVentas = periodData?.pnl?.categorias?.["Costo de Ventas"] || 0;
                    val = vNet + cVentas;
                    if (val !== 0) anyVal = true;
                } else if (normLabel === "ebitda" || normLabel === "ebitda ajustado") {
                    val = periodData?.kpis?.ebitda || 0;
                    if (val !== 0) anyVal = true;
                } else if (normLabel === "utilidad antes de impuesto" || normLabel === "utilidad neta" || normLabel === "beneficio neto del periodo") {
                    val = periodData?.kpis?.utilidad || 0;
                    if (val !== 0) anyVal = true;
                } else if (normLabel === "ventas netas" || normLabel === "ingresos") {
                    val = periodData?.kpis?.ingresos || periodData?.pnl?.categorias?.["Ingresos"] || 0;
                    if (val !== 0) anyVal = true;
                } else if (normLabel === "costo de ventas") {
                    val = periodData?.pnl?.categorias?.["Costo de Ventas"] || 0;
                    if (item.occIndex === 0) {
                        if (val !== 0) anyVal = true;
                    } else {
                        val = 0; // Evitar duplicado en el indent
                    }
                } else if (normLabel === "ggadm" || normLabel === "total ggadm") {
                    // Tomar OPEX general aproximado si falla la extracción exacta
                    const opDet = periodData?.pnl?.opexDetalle || {};
                    val = -Math.abs((opDet["Gastos Administrativos"] || 0) + (opDet["Gastos de Mercadeo"] || 0) + (opDet["Gastos de Ventas (Comercial)"] || 0) + (opDet["Gastos de Logística"] || 0));
                    if (val === 0 && periodData?.kpis?.ingresos) {
                        // Opex fallback
                        const opex = periodData.kpis.ingresos + (periodData.pnl?.categorias?.["Costo de Ventas"] || 0) - periodData.kpis.ebitda;
                        val = -Math.abs(opex);
                    }
                    if (val !== 0) anyVal = true;
                }
            }

            if (item.type === "category_main" || item.type === "category") {
                rowHtml += `<td style="${commonTdStyle}"></td>`;
            } else if (isRatio(item.type)) {
                rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-size:0.9rem; ${commonTdStyle}">${val === 0 ? '-' : formatPercent(val)}</td>`;
            } else if (isDecimal(item.type)) {
                rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-size:0.9rem; ${commonTdStyle}">${val === 0 ? '-' : val.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>`;
            } else {
                rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-size:0.9rem; ${commonTdStyle}">${val === 0 ? '-' : formatLocalMillions(val)}</td>`;
                total += val;
            }
        });

        // Add accumulated column
        if (item.type === "category_main" || item.type === "category") {
            rowHtml += `<td style="${commonTdStyle}"></td>`;
        } else if (isTotalizable) {
            rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-weight: 700; background: #f0f9ff; color: #0369a1; font-size:0.9rem; ${commonTdStyle}">${total === 0 ? '-' : formatLocalMillions(total)}</td>`;
        } else {
            rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-weight: 700; background: #f0f9ff; color: #0369a1; ${commonTdStyle}">-</td>`;
        }

        // ALWAYS render row since the user wants to see the explicit structure
        let rowClasses = [];
        if (item.type === "category_main" || item.type === "category") rowClasses.push('row-category');
        if (item.type === "bold") rowClasses.push('row-total');
        
        let rowClassStr = rowClasses.length > 0 ? ` class="${rowClasses.join(' ')}"` : '';
        tbBody += `<tr style="${rowBgColor}"${rowClassStr}>${rowHtml}</tr>`;
    });

    // Unmapped items (Otras Cuentas)
    const allConceptsInData = new Set();
    visibleMonths.forEach(periodData => {
        const sourceRows = (periodData.estados && periodData.estados.fullRows) 
                 ? periodData.estados.fullRows 
                 : (periodData.pnl && periodData.pnl.fullRows ? periodData.pnl.fullRows : []);
        if (sourceRows) {
            sourceRows.forEach(r => allConceptsInData.add(r.concept));
        }
    });

    const stringsToHide = ['otras ventas', 'otros ingresos'];
    const unmappedConcepts = Array.from(allConceptsInData).filter(c => {
        const norm = normalizeText(c);
        return !processedConcepts.has(c) && !stringsToHide.includes(norm);
    });

    if (unmappedConcepts.length > 0) {
        tbBody += `<tr class="row-category"><td colspan="${periods.length + 2}" style="background:rgba(0,0,0,0.02); font-weight:700; font-size:0.75rem; color:var(--text-secondary); text-transform:uppercase; letter-spacing:0.5px;">Otras Cuentas (No Mapeadas)</td></tr>`;
        unmappedConcepts.forEach(concept => {
            let rowHtml = `<td style="font-size:0.9rem; padding-left: 24px; font-weight: 500; color:var(--text-primary); border-bottom: 1px solid var(--border-color);">${concept}</td>`;
            let total = 0;
            let anyVal = false;
            periods.forEach(p => {
                let val = 0;
                const periodData = visibleMonths.find(d => d.date === p);
                const sourceRows = (periodData && periodData.estados && periodData.estados.fullRows) 
                     ? periodData.estados.fullRows 
                     : (periodData && periodData.pnl && periodData.pnl.fullRows ? periodData.pnl.fullRows : []);
                
                if (sourceRows && sourceRows.length > 0) {
                    const rowData = sourceRows.find(r => r.concept === concept);
                    if (rowData && rowData.values[p] !== undefined) {
                        val = rowData.values[p];
                        if (val !== 0) anyVal = true;
                    }
                }
                rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-size:0.9rem; color:var(--text-primary); border-bottom: 1px solid var(--border-color);">${val === 0 ? '-' : formatLocalMillions(val)}</td>`;
                total += val;
            });
            rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-weight: 700; background: #f0f9ff; color: #0369a1; font-size:0.9rem; border-bottom: 1px solid var(--border-color);">${total === 0 ? '-' : formatLocalMillions(total)}</td>`;
            if (anyVal) {
                tbBody += `<tr>${rowHtml}</tr>`;
            }
        });
    }

    bodyEl.innerHTML = tbBody;
}

// --------------------------------------------------------------------------------------
// PILAR B: Módulo de Desempeño Operativo (Gráficas Avanzadas D3)
// --------------------------------------------------------------------------------------

function renderWaterfallChart(data, index) {
    if (!data || !data[index]) return;
    const domElement = document.getElementById('waterfallChart');
    if (!domElement) return;

    if (domElement.__resizeObserver) {
        domElement.__resizeObserver.disconnect();
    }

    const ro = new ResizeObserver(entries => {
        const cw = domElement.clientWidth;
        if (cw > 0 && domElement.__lastWidth !== cw) {
            domElement.__lastWidth = cw;
            requestAnimationFrame(() => {
                renderWaterfallChartInternal(data, index, cw);
            });
        }
    });
    domElement.__resizeObserver = ro;
    ro.observe(domElement);

    const initialCw = domElement.clientWidth;
    if (initialCw > 0) {
        domElement.__lastWidth = initialCw;
        renderWaterfallChartInternal(data, index, initialCw);
    }
}

function renderWaterfallChartInternal(data, index, cw) {
    if (!data || !data[index]) return;
    const curr = data[index];
    const containerId = "#waterfallChart";
    
    const container = d3.select(containerId);
    if (container.empty()) return;
    
    const node = container.node();
    const parentView = node.closest ? node.closest('.view-container') : null;
    if (parentView && window.getComputedStyle(parentView).display === 'none') {
        return;
    }

    container.selectAll('*').remove();

    // Extraer datos reales del mes o YTD
    let ventasNetas = 0;
    let costoVentas = 0;
    let ebitdaReal = 0;
    
    let gAdm = 0;
    let gMerc = 0;
    let gCom = 0;
    let gLog = 0;
    let totalOpex = 0;

    if (isYTDMode) {
        const targetYear = getSortYear(curr);
        for (let k = 0; k <= index; k++) {
            const periodData = data[k];
            if (getSortYear(periodData) !== targetYear) continue;
            const pCats = periodData.pnl?.categorias || {};
            const oDet = periodData.pnl?.opexDetalle || {};
            
            ventasNetas += Math.abs(pCats["Ingresos"] || 0);
            costoVentas += Math.abs(pCats["Costo de Ventas"] || 0);
            ebitdaReal += Math.abs(pCats["EBITDA"] || 0);
            
            gAdm += Math.abs(oDet["Gastos Administrativos"] || 0);
            gMerc += Math.abs(oDet["Gastos de Mercadeo"] || 0);
            gCom += Math.abs(oDet["Gastos de Ventas (Comercial)"] || 0);
            gLog += Math.abs(oDet["Gastos de Logística"] || 0);
            
            totalOpex += Math.abs(pCats["OPEX"] || 0);
        }
    } else {
        const pnlCats = curr.pnl?.categorias || {};
        const opexDet = curr.pnl?.opexDetalle || {};
        
        ventasNetas = Math.abs(pnlCats["Ingresos"] || 0);
        costoVentas = Math.abs(pnlCats["Costo de Ventas"] || 0);
        ebitdaReal = Math.abs(pnlCats["EBITDA"] || 0);
        
        gAdm = Math.abs(opexDet["Gastos Administrativos"] || 0);
        gMerc = Math.abs(opexDet["Gastos de Mercadeo"] || 0);
        gCom = Math.abs(opexDet["Gastos de Ventas (Comercial)"] || 0);
        gLog = Math.abs(opexDet["Gastos de Logística"] || 0);
        totalOpex = Math.abs(pnlCats["OPEX"] || 0);
    }
    
    let otrosGastos = totalOpex - gAdm - gMerc - gCom - gLog;
    if (otrosGastos < 0) otrosGastos = 0;

    let current = Math.abs(ventasNetas);
    const chartData = [];
    
    // 1. Inicio: Ventas Netas
    chartData.push({
        name: "Ventas Netas",
        isTotal: true,
        start: 0,
        end: current,
        value: current,
        color: "var(--sidebar-accent)"
    });

    const addDeduction = (name, amount, isPositive = false) => {
        if (amount !== 0) {
            chartData.push({
                name: name,
                isTotal: false,
                start: current,
                end: current + (isPositive ? Math.abs(amount) : -Math.abs(amount)),
                value: isPositive ? Math.abs(amount) : -Math.abs(amount),
                color: isPositive ? "var(--success)" : "var(--danger)"
            });
            current += (isPositive ? Math.abs(amount) : -Math.abs(amount));
        }
    };

    addDeduction("Costo de Ventas", costoVentas);
    addDeduction("Gastos Adm", gAdm);
    addDeduction("Gastos Merc/Com", gMerc + gCom);
    
    const restos = gLog + otrosGastos;
    addDeduction("Log/Otros Gastos", restos);

    // Si llegados a este punto 'current' no es igual al ebitdaReal, agregamos un ajuste (D&A u otras partidas)
    const ebitdaGap = ebitdaReal - current;
    if (Math.abs(ebitdaGap) > 0.1) {
        addDeduction("Otros", Math.abs(ebitdaGap), ebitdaGap > 0);
    }

    chartData.push({
        name: "EBITDA",
        isTotal: true,
        start: 0,
        end: current,
        value: current,
        color: "var(--sidebar-accent)"
    });

    // Drawing the Waterfall
    const width = cw;
    const height = container.node().clientHeight || 350;
    const margin = { top: 30, right: 20, bottom: 40, left: 60 };

    const svg = container.append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand()
        .domain(chartData.map(d => d.name))
        .range([0, width - margin.left - margin.right])
        .padding(0.3);

    const maxVal = d3.max(chartData, d => Math.max(d.start, d.end));
    const minVal = d3.min(chartData, d => Math.min(d.start, d.end));

    const y = d3.scaleLinear()
        .domain([Math.min(0, minVal * 1.1), maxVal * 1.1])
        .range([height - margin.top - margin.bottom, 0]);

    // X Axis
    svg.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).tickSizeOuter(0))
        .selectAll("text")
        .style("font-weight", "600")
        .style("color", "var(--text-primary)")
        .style("font-size", window.innerWidth < 768 ? "9px" : "11px");

    // Y Axis
    svg.append("g")
        .call(d3.axisLeft(y).ticks(5).tickFormat(d => d + "M"))
        .style("color", "var(--text-secondary)")
        .style("font-size", "10px");

    let tooltip = d3.select("body").select(".d3-tooltip");
    if (tooltip.empty()) {
        tooltip = d3.select("body").append("div").attr("class", "d3-tooltip").style("opacity", 0);
    }

    svg.selectAll(".bar")
        .data(chartData)
        .enter().append("rect")
        .attr("class", "bar")
        .attr("x", d => x(d.name))
        .attr("width", x.bandwidth())
        .attr("y", d => y(Math.max(d.start, d.end)))
        .attr("height", d => Math.max(1, Math.abs(y(d.start) - y(d.end))))
        .attr("fill", d => d.color)
        .attr("rx", 4)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
            d3.select(this).style("filter", "brightness(1.1)");
            const valFormated = formatCurrency(Math.abs(d.value));
            const weightStr = ventasNetas ? (Math.abs(d.value) / ventasNetas * 100).toFixed(1) + "%" : "0%";
            const label = d.isTotal ? "Total" : (d.value > 0 ? "Adición" : "Deducción");
            const sign = (!d.isTotal && d.value < 0) ? "-" : "";
            tooltip.style("opacity", 1)
                   .html(`<strong>${d.name}</strong><br/>${label}: ${valFormated}<br/>Peso vs Ventas: ${sign}${weightStr}`);
        })
        .on("mousemove", function(event) {
            tooltip.style("left", (event.pageX + 15) + "px")
                   .style("top", (event.pageY - 15) + "px");
        })
        .on("mouseout", function() {
            d3.select(this).style("filter", "none");
            tooltip.style("opacity", 0);
        });

    // Connecting lines
    svg.selectAll(".connector")
        .data(chartData.slice(0, -1))
        .enter().append("line")
        .attr("class", "connector")
        .attr("x1", d => x(d.name) + x.bandwidth())
        .attr("y1", d => y(d.end))
        .attr("x2", (d, i) => x(chartData[i + 1].name))
        .attr("y2", d => y(d.end))
        .attr("stroke", "var(--text-secondary)")
        .attr("stroke-dasharray", "3,3")
        .attr("stroke-width", 1);

    // Labels
    svg.selectAll(".label")
        .data(chartData)
        .enter().append("text")
        .attr("class", "label")
        .attr("x", d => x(d.name) + x.bandwidth() / 2)
        .attr("y", d => y(Math.max(d.start, d.end)) - 5)
        .attr("text-anchor", "middle")
        .style("font-size", "10px")
        .style("font-weight", "bold")
        .style("fill", "var(--text-primary)")
        .text(d => (d.value > 0 ? '' : '') + d.value.toFixed(1) + 'M');

    // Title
    svg.append("text")
        .attr("x", 0)
        .attr("y", -10)
        .style("font-size", "14px")
        .style("font-weight", "800")
        .style("fill", "var(--sidebar-dark)")
        .text(`Puente de Rentabilidad: Ventas a EBITDA (${isYTDMode ? 'YTD ' : ''}${curr.date})`);
}

function renderMarginTrendChart(globalData, index) {
    if (!globalData || globalData.length === 0) return;
    const domElement = document.getElementById('marginTrendChart');
    if (!domElement) return;

    if (domElement.__resizeObserver) {
        domElement.__resizeObserver.disconnect();
    }

    const ro = new ResizeObserver(entries => {
        const cw = domElement.clientWidth;
        if (cw > 0 && domElement.__lastWidth !== cw) {
            domElement.__lastWidth = cw;
            requestAnimationFrame(() => {
                renderMarginTrendChartInternal(globalData, index, cw);
            });
        }
    });
    domElement.__resizeObserver = ro;
    ro.observe(domElement);

    const initialCw = domElement.clientWidth;
    if (initialCw > 0) {
        domElement.__lastWidth = initialCw;
        renderMarginTrendChartInternal(globalData, index, initialCw);
    }
}

function renderMarginTrendChartInternal(globalData, index, cw) {
    if (!globalData || globalData.length === 0) return;
    const containerId = "#marginTrendChart";
    
    const container = d3.select(containerId);
    if (container.empty()) return;
    
    const node = container.node();
    const parentView = node.closest ? node.closest('.view-container') : null;
    if (parentView && window.getComputedStyle(parentView).display === 'none') {
        return;
    }

    container.selectAll('*').remove();

    // Filtramos para ignorar 2025 base y sacar datos de PPTO vs Real 
    // y limitamos los datos hasta el mes seleccionado (index)
    const isMobile = window.innerWidth < 768;
    const slicedData = globalData.slice(0, index !== undefined ? index + 1 : globalData.length);
    const validData = slicedData.filter(d => isYear2026(d));
    if (validData.length === 0) return;

    // Tomamos al menos los últimos 12 meses (o 6 en mobile)
    const elementsToSlice = isMobile ? -6 : -12;
    const chartData = validData.slice(elementsToSlice).map(d => ({
        date: d.date,
        realIngresos: d.kpis.ingresos || 0,
        pptoIngresos: (d.ppto && d.ppto.kpis && d.ppto.kpis.ingresos) ? d.ppto.kpis.ingresos : 0,
        realMargen: (d.kpis.margen_ebitda || 0) * 100,
        pptoMargen: (d.ppto && d.ppto.kpis && d.ppto.kpis.ingresos !== 0) ? ((d.ppto.kpis.ebitda || 0) / (d.ppto.kpis.ingresos || 1)) * 100 : ((d.ppto && d.ppto.pnl && d.ppto.pnl.categorias && d.ppto.pnl.categorias.EBITDA) ? (d.ppto.pnl.categorias.EBITDA / (d.ppto.pnl.categorias.Ingresos || 1) * 100) : 0)
    }));

    const width = cw;
    const height = container.node().clientHeight || 300;
    const margin = { top: 40, right: 50, bottom: 40, left: 50 };

    const svg = container.append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand()
        .domain(chartData.map(d => d.date))
        .range([0, width - margin.left - margin.right])
        .padding(0.4);

    const maxIngresos = d3.max(chartData, d => Math.max(d.realIngresos, d.pptoIngresos));
    const yLeft = d3.scaleLinear()
        .domain([0, maxIngresos * 1.15])
        .range([height - margin.top - margin.bottom, 0]);

    const maxMargen = d3.max(chartData, d => Math.max(d.realMargen, d.pptoMargen));
    const yRight = d3.scaleLinear()
        .domain([0, maxMargen * 1.2])
        .range([height - margin.top - margin.bottom, 0]);

    // Ejes
    svg.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).tickSizeOuter(0))
        .selectAll("text")
        .style("font-weight", "600")
        .style("color", "var(--text-primary)")
        .style("font-size", isMobile ? "9px" : "11px");

    svg.append("g")
        .call(d3.axisLeft(yLeft).ticks(5).tickFormat(d => d + "M"))
        .style("color", "var(--text-secondary)")
        .style("font-size", "10px");

    svg.append("g")
        .attr("transform", `translate(${width - margin.left - margin.right}, 0)`)
        .call(d3.axisRight(yRight).ticks(5).tickFormat(d => d + "%"))
        .style("color", "var(--text-secondary)")
        .style("font-size", "10px");

    let tooltip = d3.select("body").select(".d3-tooltip");

    // Barras (Fondo PPTO)
    svg.selectAll(".bar-ppto")
        .data(chartData)
        .enter().append("rect")
        .attr("class", "bar-ppto")
        .attr("x", d => x(d.date))
        .attr("width", x.bandwidth())
        .attr("y", d => yLeft(d.pptoIngresos))
        .attr("height", d => Math.max(0, height - margin.top - margin.bottom - yLeft(d.pptoIngresos)))
        .attr("fill", "rgba(148, 163, 184, 0.6)")
        .attr("stroke", "rgba(71, 85, 105, 1)")
        .attr("stroke-width", 2)
        .attr("stroke-dasharray", "4,4")
        .attr("rx", 4)
        .style("opacity", d => d.pptoIngresos > 0 ? 1 : 0);

    // Barras (Frente Real)
    svg.selectAll(".bar-real")
        .data(chartData)
        .enter().append("rect")
        .attr("class", "bar-real")
        .attr("x", d => x(d.date) + x.bandwidth() * 0.15)
        .attr("width", x.bandwidth() * 0.7)
        .attr("y", d => yLeft(Math.max(0, d.realIngresos)))
        .attr("height", d => Math.max(0, height - margin.top - margin.bottom - yLeft(Math.max(0, d.realIngresos))))
        .attr("fill", "var(--sidebar-accent)")
        .attr("rx", 3)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
            d3.select(this).style("filter", "brightness(1.1)");
            tooltip.style("opacity", 1)
                   .html(`<strong>${d.date}</strong><br/>Real: ${formatCurrency(d.realIngresos)}<br/>PPTO: ${formatCurrency(d.pptoIngresos)}`);
        })
        .on("mousemove", function(event) {
            tooltip.style("left", (event.pageX + 15) + "px")
                   .style("top", (event.pageY - 15) + "px");
        })
        .on("mouseout", function() {
            d3.select(this).style("filter", "none");
            tooltip.style("opacity", 0);
        });

    // Línea PPTO
    const linePpto = d3.line()
        .x(d => x(d.date) + x.bandwidth() / 2)
        .y(d => yRight(d.pptoMargen))
        .curve(d3.curveMonotoneX);

    svg.append("path")
        .datum(chartData.filter(d => d.pptoMargen > 0))
        .attr("fill", "none")
        .attr("stroke", "#94a3b8")
        .attr("stroke-width", 2)
        .attr("stroke-dasharray", "5,5")
        .attr("d", linePpto);

    // Línea Real
    const lineReal = d3.line()
        .x(d => x(d.date) + x.bandwidth() / 2)
        .y(d => yRight(d.realMargen))
        .curve(d3.curveMonotoneX);

    svg.append("path")
        .datum(chartData)
        .attr("fill", "none")
        .attr("stroke", "var(--warning)")
        .attr("stroke-width", 3)
        .attr("d", lineReal);

    // Puntos Línea Real
    svg.selectAll(".dot-real")
        .data(chartData)
        .enter().append("circle")
        .attr("class", "dot-real")
        .attr("cx", d => x(d.date) + x.bandwidth() / 2)
        .attr("cy", d => yRight(d.realMargen))
        .attr("r", 4)
        .attr("fill", "white")
        .attr("stroke", "var(--warning)")
        .attr("stroke-width", 2)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
             d3.select(this).attr("r", 6);
             tooltip.style("opacity", 1)
                    .html(`<strong>${d.date}</strong><br/>Margen Real: ${d.realMargen.toFixed(1)}%<br/>Margen PPTO: ${(d.pptoMargen || 0).toFixed(1)}%`);
        })
        .on("mousemove", function(event) {
             tooltip.style("left", (event.pageX + 15) + "px")
                    .style("top", (event.pageY - 15) + "px");
        })
        .on("mouseout", function() {
             d3.select(this).attr("r", 4);
             tooltip.style("opacity", 0);
        });

    // Legends y Title
    svg.append("text")
        .attr("x", 0)
        .attr("y", -20)
        .style("font-size", "14px")
        .style("font-weight", "800")
        .style("fill", "var(--sidebar-dark)")
        .text("Ingresos vs PPTO y Evolución de Margen EBITDA (%)");

    // Leyenda
    const legendX = isMobile ? 0 : width - margin.left - margin.right - 250;
    const legendY = isMobile ? -5 : -25;
    const legend = svg.append("g").attr("transform", `translate(${legendX}, ${legendY})`);
    
    legend.append("rect").attr("x", 0).attr("y", 0).attr("width", 10).attr("height", 10).attr("fill", "var(--sidebar-accent)");
    legend.append("text").attr("x", 15).attr("y", 9).style("font-size", "10px").text("Real");
    
    legend.append("rect").attr("x", 50).attr("y", 0).attr("width", 10).attr("height", 10).attr("fill", "rgba(148, 163, 184, 0.6)").attr("stroke", "rgba(71, 85, 105, 1)").attr("stroke-width", 1).attr("stroke-dasharray", "2,2");
    legend.append("text").attr("x", 65).attr("y", 9).style("font-size", "10px").text("PPTO");
    
    legend.append("line").attr("x1", 105).attr("y1", 5).attr("x2", 125).attr("y2", 5).attr("stroke", "var(--warning)").attr("stroke-width", 2);
    legend.append("text").attr("x", 130).attr("y", 9).style("font-size", "10px").text("Mg Real");
    
    legend.append("line").attr("x1", 180).attr("y1", 5).attr("x2", 200).attr("y2", 5).attr("stroke", "#94a3b8").attr("stroke-dasharray", "3,3").attr("stroke-width", 2);
    legend.append("text").attr("x", 205).attr("y", 9).style("font-size", "10px").text("Mg PPTO");
}

// --------------------------------------------------------------------------------------
// PILAR C: Módulo de Liquidez (Gráficas Avanzadas D3)
// --------------------------------------------------------------------------------------

function renderCashBridgeChart(data, index) {
    if (!data || !data[index]) return;
    const domElement = document.getElementById('cashBridgeChart');
    if (!domElement) return;

    if (domElement.__resizeObserver) {
        domElement.__resizeObserver.disconnect();
    }

    const ro = new ResizeObserver(entries => {
        const cw = domElement.clientWidth;
        if (cw > 0 && domElement.__lastWidth !== cw) {
            domElement.__lastWidth = cw;
            requestAnimationFrame(() => {
                renderCashBridgeChartInternal(data, index, cw);
            });
        }
    });
    domElement.__resizeObserver = ro;
    ro.observe(domElement);

    const initialCw = domElement.clientWidth;
    if (initialCw > 0) {
        domElement.__lastWidth = initialCw;
        renderCashBridgeChartInternal(data, index, initialCw);
    }
}

function renderCashBridgeChartInternal(data, index, cw) {
    if (!data || !data[index]) return;
    const curr = data[index];
    const containerId = "#cashBridgeChart";
    
    const container = d3.select(containerId);
    if (container.empty()) return;
    
    const node = container.node();
    const parentView = node.closest ? node.closest('.view-container') : null;
    if (parentView && window.getComputedStyle(parentView).display === 'none') {
        return;
    }
    
    container.selectAll('*').remove();

    let beginning = 0;
    let operating = 0;
    let capex = 0;
    let netDebt = 0;
    let interest = 0;
    let dividends = 0;
    let ending = 0;

    let change = 0;
    
    if (isYTDMode) {
        let firstIdx = 0;
        const targetYear = getSortYear(curr);
        for (let k = 0; k <= index; k++) {
            if (getSortYear(data[k]) === targetYear) {
                firstIdx = k;
                break;
            }
        }
        beginning = data[firstIdx]?.cashflowDetail?.beginning || 0;
        ending = data[index]?.cashflowDetail?.ending || 0; 
        
        for (let k = firstIdx; k <= index; k++) {
            if (getSortYear(data[k]) !== targetYear) continue;
            const det = data[k]?.cashflowDetail || {};
            operating += (det.operating || 0);
            capex += (det.capex || 0);
            netDebt += (det.netDebt || 0);
            interest += (det.interest || 0);
            dividends += (det.dividends || 0);
        }
        change = operating + capex + netDebt + interest + dividends;
    } else {
        const det = curr.cashflowDetail || {};
        beginning = det.beginning || 0;
        operating = det.operating || 0;
        capex = det.capex || 0;
        netDebt = det.netDebt || 0;
        interest = det.interest || 0;
        dividends = det.dividends || 0;
        ending = det.ending || 0;
        change = operating + capex + netDebt + interest + dividends;
    }

    let current = beginning;
    const chartData = [];
    
    // 1. Inicio: Efectivo Inicial
    chartData.push({
        name: "Efectivo Inicial",
        isTotal: true,
        start: 0,
        end: current,
        value: current,
        color: "var(--sidebar-accent)"
    });

    const addVariation = (name, amount) => {
        if (Math.abs(amount) > 0.001) {
            let isPositive = amount >= 0;
            // Para tooltips y demás, el 'value' numérico se mantiene, pero si es una salida el start>end.
            chartData.push({
                name: name,
                isTotal: false,
                start: current,
                end: current + amount,
                value: amount,
                color: isPositive ? "var(--success)" : "var(--danger)"
            });
            current += amount;
        }
    };

    addVariation("Flujo de Caja Operativo", operating);
    addVariation("CAPEX", capex);
    addVariation("Deuda Bancaria", netDebt);
    addVariation("Gastos de Interés", interest);
    addVariation("Otros Flujos", dividends);

    const gap = ending - current;
    if (Math.abs(gap) > 0.1) {
        addVariation("Ajustes", gap);
    }

    chartData.push({
        name: "Efectivo Final",
        isTotal: true,
        start: 0,
        end: ending,
        value: ending,
        color: "var(--sidebar-accent)"
    });

    // D3 Setup
    const isMobile = window.innerWidth < 1024;
    const width = cw;
    const height = 350;
    
    // Márgenes más generosos abajo para que quepan las etiquetas en móvil
    const margin = isMobile 
        ? { top: 40, right: 20, bottom: 90, left: 50 } 
        : { top: 40, right: 30, bottom: 80, left: 80 };

    const svg = container.append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand()
        .domain(chartData.map(d => d.name))
        .range([0, width - margin.left - margin.right])
        .padding(0.3);

    const allValues = chartData.map(d => d.start).concat(chartData.map(d => d.end));
    const yMin = Math.min(0, d3.min(allValues)) * 1.25;
    const yMax = Math.max(0, d3.max(allValues)) * 1.25;

    const y = d3.scaleLinear()
        .domain([yMin, yMax])
        .range([height - margin.top - margin.bottom, 0]);

    // Gridlines
    svg.append("g")
        .attr("class", "grid")
        .call(d3.axisLeft(y).ticks(5).tickSize(-(width - margin.left - margin.right)).tickFormat(""))
        .selectAll("line")
        .style("stroke", "#e2e8f0")
        .style("stroke-dasharray", "3,3");
    svg.selectAll(".domain").remove();

    // Axes
    const xAxisY = height - margin.top - margin.bottom;
    svg.append("g")
        .attr("transform", `translate(0,${xAxisY})`) 
        .call(d3.axisBottom(x).tickSize(0))
        .selectAll("text")
        .style("text-anchor", "end")
        .attr("dx", "-.8em")
        .attr("dy", ".15em")
        .attr("transform", "rotate(-25)")
        .style("font-size", isMobile ? "9px" : "11px")
        .style("font-weight", "600")
        .style("fill", "var(--text-secondary)");
    
    svg.select(".domain").remove();

    svg.append("g")
        .call(d3.axisLeft(y).ticks(5).tickFormat(d => d.toFixed(0) + "M"))
        .selectAll("text")
        .style("font-size", "11px")
        .style("fill", "var(--text-secondary)")
        .style("font-weight", "600");
    svg.select(".domain").remove();

    // Tooltip
    let tooltip = d3.select("body").select(".d3-tooltip");

    // Bars
    svg.selectAll(".waterfall-bar")
        .data(chartData)
        .enter().append("rect")
        .attr("class", "waterfall-bar")
        .attr("x", d => x(d.name))
        .attr("y", d => y(Math.max(d.start, d.end)))
        .attr("height", d => Math.abs(y(d.start) - y(d.end)) || 1) // prevent 0 height
        .attr("width", x.bandwidth())
        .attr("fill", d => d.color)
        .attr("rx", 4)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
             d3.select(this).style("filter", "brightness(1.1)");
             const valText = (d.value > 0 && !d.isTotal ? "+" : "") + (d.value).toFixed(1) + "M";
             tooltip.style("opacity", 1)
                    .html(`<strong>${d.name}</strong><br/>RD$ ${valText}`);
        })
        .on("mousemove", function(event) {
             tooltip.style("left", (event.pageX + 15) + "px")
                    .style("top", (event.pageY - 15) + "px");
        })
        .on("mouseout", function() {
             d3.select(this).style("filter", "none");
             tooltip.style("opacity", 0);
        });

    // Conector lines
    svg.selectAll(".connector")
        .data(chartData.slice(0, -1))
        .enter().append("line")
        .attr("class", "connector")
        .attr("x1", d => x(d.name) + x.bandwidth())
        .attr("y1", d => y(d.end))
        .attr("x2", (d, i) => x(chartData[i+1].name))
        .attr("y2", d => y(d.end))
        .style("stroke", "var(--text-secondary)")
        .style("stroke-dasharray", "4,4")
        .style("stroke-width", 1);

    // Etiqueta de valores
    svg.selectAll(".bar-label")
        .data(chartData)
        .enter().append("text")
        .attr("class", "bar-label")
        .attr("x", d => x(d.name) + x.bandwidth() / 2)
        .attr("y", d => {
            if (d.end >= d.start) {
                return y(d.end) - 5;
            } else {
                return y(d.end) + 15;
            }
        })
        .style("text-anchor", "middle")
        .style("font-size", isMobile ? "9px" : "11px")
        .style("font-weight", "700")
        .style("fill", "var(--sidebar-dark)")
        .text(d => {
            const val = d.value;
            return (val > 0 && !d.isTotal ? "+" : "") + val.toFixed(1) + "M";
        });

    // Title
    svg.append("text")
        .attr("x", 0)
        .attr("y", -15)
        .style("font-size", "14px")
        .style("font-weight", "800")
        .style("fill", "var(--sidebar-dark)")
        .text(`Puente de Efectivo (Cash Bridge) - ${isYTDMode ? 'YTD ' : ''}${curr.date}`);
}

// --------------------------------------------------------------------------------------
// PILAR D: Módulo de Riesgo y Covenants
// --------------------------------------------------------------------------------------

function renderCovenantGauges(data, index) {
    if (!data || !data[index]) return;
    const curr = data[index];
    
    const d3Container = d3.select('#covenantsContainer');
    if (d3Container.empty()) return;
    d3Container.selectAll('*').remove();
    
    let container = document.getElementById('covenantsContainer');
    if (!container) return;
    
    const parentView = container.closest('.view-container');
    if (parentView && window.getComputedStyle(parentView).display === 'none') {
        return;
    }
    
    // Create card wrappers
    const isMobile = window.innerWidth < 768;
    const createCard = (id, title) => {
        const div = document.createElement('div');
        div.className = 'chart-card';
        div.style.flex = '1';
        div.style.minWidth = isMobile ? '100%' : 'calc(50% - 10px)';
        div.style.backgroundColor = 'white';
        div.style.padding = '15px';
        div.style.borderRadius = '12px';
        div.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';
        div.innerHTML = `<h4 style="margin-bottom: 10px; font-weight: 600; font-size: 13px; color: var(--text); text-align: center;">${title}</h4><div id="${id}" style="display: flex; justify-content: center; position: relative;"></div>`;
        container.appendChild(div);
        return id;
    };
    
    const levId = createCard('gaugeLeverage', 'Apalancamiento (Deuda / EBITDA)');
    const covId = createCard('gaugeCoverage', 'Endeudamiento (Pasivo / Patrimonio)');
    
    // Cálculos
    const deudaTotal = curr.balance ? (curr.balance.deudaTotal || 0) : 0;
    const pasivos = curr.balance ? (curr.balance.pasivos || 0) : 0;
    const patrimonio = curr.balance ? (curr.balance.patrimonio || 0) : 0;
    
    let ebitdaYTD = 0;
    
    const targetYear = curr.sortDate ? new Date(curr.sortDate).getFullYear() : 2026;
    let currentMonthNum = curr.sortDate ? new Date(curr.sortDate).getMonth() + 1 : 1;
    if (isNaN(currentMonthNum) || currentMonthNum < 1) currentMonthNum = 1;

    for (let k = 0; k <= index; k++) {
        const d = data[k];
        const dYear = d.sortDate ? new Date(d.sortDate).getFullYear() : targetYear;
        if (dYear === targetYear && isYear2026(d)) {
            ebitdaYTD += (d.kpis.ebitda || 0);
        }
    }
    
    const ebitdaAnualizado = (ebitdaYTD / currentMonthNum) * 12;
    
    let leverageValue = (ebitdaAnualizado > 0) ? (deudaTotal / ebitdaAnualizado) : 0;
    if (leverageValue < 0) leverageValue = 0;
    if (ebitdaAnualizado <= 0) leverageValue = 0; // fallback if negative ebitda
    
    let debtEquityValue = (patrimonio > 0) ? (pasivos / patrimonio) : 0;
    if (patrimonio <= 0) debtEquityValue = 99.9; // Negative equity or 0
    if (debtEquityValue < 0 && debtEquityValue !== 99.9) debtEquityValue = 0;
    
    // Helper para Semicírculos (Half-Donut)
    const drawHalfDonut = (selectorId, value, threshold, limitMax, colorLogic) => {
        const wrapper = d3.select(`#${selectorId}`);
        wrapper.selectAll('*').remove();
        
        const width = 200;
        const height = 100; 
        const margin = 10;
        const radius = Math.min(width, height * 2) / 2 - margin;
        
        const svg = wrapper.append("svg")
            .attr("width", width)
            .attr("height", height)
            .append("g")
            .attr("transform", `translate(${width / 2},${height - 10})`);
            
        const arc = d3.arc()
            .innerRadius(radius * 0.6)
            .outerRadius(radius)
            .startAngle(-Math.PI / 2);
            
        // Fondo Gris
        svg.append("path")
            .datum({ endAngle: Math.PI / 2 })
            .style("fill", "#e2e8f0")
            .attr("d", arc);
            
        let cappedVal = Math.min(Math.max(value, 0), limitMax);
        const angle = -Math.PI / 2 + (cappedVal / limitMax) * Math.PI;
        const color = colorLogic(value);

        const foreground = svg.append("path")
            .datum({ endAngle: -Math.PI / 2 })
            .style("fill", color)
            .attr("d", arc);

        foreground.transition()
            .duration(1000)
            .attrTween("d", function(d) {
                const i = d3.interpolate(d.endAngle, angle);
                return function(t) {
                    d.endAngle = i(t);
                    return arc(d);
                }
            });

        // Threshold Marker
        const thresholdAngle = -Math.PI / 2 + (threshold / limitMax) * Math.PI;
        const lineLen = radius + 5;
        const innerLen = radius * 0.6 - 5;
        svg.append("line")
            .attr("x1", innerLen * Math.sin(thresholdAngle))
            .attr("y1", -innerLen * Math.cos(thresholdAngle))
            .attr("x2", lineLen * Math.sin(thresholdAngle))
            .attr("y2", -lineLen * Math.cos(thresholdAngle))
            .attr("stroke", "#0f172a")
            .attr("stroke-width", 2)
            .attr("stroke-dasharray", "2,2");

        // Valor Numérico Central
        let displayValText = value === 99.9 ? "N/A" : value.toFixed(1) + "x";
        if (value > 50 && value !== 99.9) displayValText = ">50.0x";
        
        svg.append("text")
            .attr("text-anchor", "middle")
            .attr("y", -10) // Centered baseline relative to origin
            .style("font-size", "20px")
            .style("font-weight", "800")
            .style("fill", color)
            .text(displayValText);
            
        // Etiqueta del Threshold
        svg.append("text")
            .attr("text-anchor", "middle")
            .attr("y", 5) // Below centerline inside the hole
            .style("font-size", "10px")
            .style("fill", "var(--text-secondary)")
            .text("Límite: " + threshold.toFixed(1) + "x");
    };

    // Covenant 1: Apalancamiento (< 3.0 Verde, >3 Rojo)
    const colorLev = (val) => {
        if (val === 0) return "var(--text-secondary)";
        if (val <= 2.5) return "var(--success)";
        if (val <= 3.0) return "var(--warning)";
        return "var(--danger)";
    };
    drawHalfDonut(levId, leverageValue, 3.0, 5.0, colorLev);

    // Covenant 2: Endeudamiento (< 2.0 Verde, >2.5 Rojo)
    const colorDebtEq = (val) => {
        if (val === 0) return "var(--text-secondary)";
        if (val === 99.9) return "var(--danger)"; // Negative equity
        if (val <= 1.5) return "var(--success)";
        if (val <= 2.5) return "var(--warning)";
        return "var(--danger)";
    };
    drawHalfDonut(covId, debtEquityValue, 2.5, 4.0, colorDebtEq);
}

// --------------------------------------------------------------------------------------
// CFO CO-PILOT E INTERACCIÓN CON IA (CHAT Y WHAT-IF)
// --------------------------------------------------------------------------------------

// 1. Lógica del Panel Lateral (Chat)
document.addEventListener('DOMContentLoaded', () => {
    const aiChatSidebar = document.getElementById('aiChatSidebar');
    const openAiChatBtn = document.getElementById('openAiChatBtn');
    const closeAiChat = document.getElementById('closeAiChat');
    const aiChatInput = document.getElementById('aiChatInput');
    const sendAiChatBtn = document.getElementById('sendAiChatBtn');
    const chatMessages = document.getElementById('chatMessages');

    if (openAiChatBtn) {
        openAiChatBtn.addEventListener('click', () => {
            aiChatSidebar.classList.add('open');
            aiChatInput.focus();
        });
    }

    if (closeAiChat) {
        closeAiChat.addEventListener('click', () => {
            aiChatSidebar.classList.remove('open');
        });
    }

    const appendMessage = (text, isUser) => {
        const msgDiv = document.createElement('div');
        msgDiv.className = 'chat-msg ' + (isUser ? 'user-msg' : 'ai-msg');
        msgDiv.innerHTML = text;
        chatMessages.appendChild(msgDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    };

    function getDashboardContext() {
        let masterContext = "";
        if (globalFinancialData && globalFinancialData.length > 0) {
            const monthSelector = document.getElementById('monthSelector');
            const idx = monthSelector ? parseInt(monthSelector.value, 10) : globalFinancialData.length - 1;
            const curr = globalFinancialData[idx || globalFinancialData.length - 1];

            masterContext = `
            Datos actuales del dashboard al ${curr.date}:
            - Ingresos (Kpis): RD$ ${(curr.kpis?.ingresos || 0).toFixed(2)}M
            - EBITDA (Kpis): RD$ ${(curr.kpis?.ebitda || 0).toFixed(2)}M
            - Utilidad Neta (Kpis): RD$ ${(curr.kpis?.utilidad || 0).toFixed(2)}M
            - Flujo de Caja (Generación): RD$ ${(curr.kpis?.cashflow || 0).toFixed(2)}M
            - Efectivo Final: RD$ ${(curr.cashflowDetail?.ending || 0).toFixed(2)}M
            - Margen Bruto: ${((curr.kpis?.margen_bruto || 0) * 100).toFixed(1)}%
            - Margen Neto: ${((curr.kpis?.margen_neto || 0) * 100).toFixed(1)}%
            - Deuda Total: RD$ ${(curr.balance?.deudaTotal || 0).toFixed(2)}M
            - Apalancamiento (Deuda/EBITDA): ${((curr.balance?.deudaTotal || 0) / (curr.balance?.ebitdaLTM || 1)).toFixed(2)}x
            `;
        }

        let ventasCeoContext = "";
        if (ceoData && ceoData.length > 0) {
            ventasCeoContext = `
            =================================
            System Prompt — Procesamiento de Ventas CEO (Planeta Azul)
            Eres el analista financiero de datos de Planeta Azul S.A. con acceso al archivo Excel de Ventas CEO.
            Tu rol: Procesas datos de Ventas CEO.
            Cuando se te pida analizar el archivo:
            Extrae la jerarquía de productos de "Tablas Consejo".
            La estructura actual en ceoData ya tiene:
            values: reales mensuales.
            pptoValues: presupuestos.
            FY2024, PO26, etc.
            Si ves 0 en los meses reales de 2026 hacia adelante, usa el valor PPTO. (El sistema pre-procesó y aplicó fallback a los objetos).
            Si el usuario te pide extraer datos, responde SOLO con JSON válido, sin markdown, sin texto adicional (como en el requerimiento "Extrae los datos de Volumen...").
            
            Jerarquía de productos actual extraída y disponible en data:
            AGUA PLANETA AZUL (padre de: APA BOTELLON 18.9, APA BOTELLA 0.5, etc.)
            MAQUILAS (padre de: MAQUILA AGUA OTROS...)
            BEBIDAS (padre de: BON, PA SABOR 0.5...)
            
            Dato en crudo de Ventas CEO (JSON):
            ${JSON.stringify(ceoData.map(d => ({ Producto: d.Producto, Tipo: d.Tipo, parentId: d.parentId, hasChildren: d.hasChildren, values: d.values, pptoValues: d.pptoValues, FY2024: d.FY2024, PO26: d.PO26 }))) }
            =================================
            `;
        }

        if (!masterContext && !ventasCeoContext) {
             return "No hay datos cargados en el sistema.";
        }

        return masterContext + ventasCeoContext + "\nEste es el contexto para tus respuestas.";
    }

    const handleChatSubmit = async () => {
        if (!window.aiEnabled) return;
        const question = aiChatInput.value.trim();
        if (!question) return;

        appendMessage(question, true);
        aiChatInput.value = '';

        const context = getDashboardContext();
        appendMessage('<i data-lucide="loader" class="spin-icon"></i> Analizando...', false);
        lucide.createIcons();

        try {
            const prompt = `Eres el CFO Co-Pilot de Planeta Azul. Eres analítico y directo.
Responde a esta pregunta basándote únicamente en el siguiente contexto financiero. Sé breve (máximo 3-4 oraciones) y usa métricas. Da la respuesta en formato HTML si necesitas negritas.
Contexto:
${context}
Pregunta: ${question}`;

            let timeoutId;
            const timeoutPromise = new Promise((_, reject) => {
                timeoutId = setTimeout(() => reject(new Error('AI Request Timeout (45s)')), 45000);
            });
            let apiCallPromise;
            try {
                apiCallPromise = getAI().models.generateContent({
                    model: "gemini-2.5-flash",
                    contents: prompt,
                });
                apiCallPromise.catch(err => window.handleAiError("Chat", err));
            } catch (err) {
                apiCallPromise = Promise.reject(err);
                apiCallPromise.catch(()=> /* handled */ {});
            }
            
            let response;
            try {
                response = await Promise.race([apiCallPromise, timeoutPromise]);
            } finally {
                clearTimeout(timeoutId);
            }

            // Reemplazar spinner con la respuesta
            chatMessages.lastChild.remove(); 
            appendMessage(response.text, false);
            lucide.createIcons();
        } catch (err) {
            chatMessages.lastChild.remove();
            appendMessage('Lo siento, hubo un problema al procesar tu solicitud: ' + err.message, false);
        }
    };

    if (sendAiChatBtn) {
        sendAiChatBtn.addEventListener('click', handleChatSubmit);
    }
    if (aiChatInput) {
        aiChatInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') handleChatSubmit();
        });
    }

    // 2. Lógica del Simulador What-If
    const simVentasVolInp = document.getElementById('sim-ventas-vol');
    const simPreciosInp = document.getElementById('sim-precios');
    const simCogsInp = document.getElementById('sim-cogs');
    const simOpexInp = document.getElementById('sim-opex');
    const simDsoInp = document.getElementById('sim-dso');
    
    const labelVentasVol = document.getElementById('label-sim-ventas-vol');
    const labelPrecios = document.getElementById('label-sim-precios');
    const labelCogs = document.getElementById('label-sim-cogs');
    const labelOpex = document.getElementById('label-sim-opex');
    const labelDso = document.getElementById('label-sim-dso');

    window.runSimulationMath = () => {
        if (!globalFinancialData || globalFinancialData.length === 0) return null;
        
        const monthSelector = document.getElementById('monthSelector');
        const idx = monthSelector ? parseInt(monthSelector.value, 10) : globalFinancialData.length - 1;
        const curr = globalFinancialData[idx || globalFinancialData.length - 1];

        // Setup Real values (Base Actual)
        const realIngresos = curr.kpis?.ingresos || 0;
        const realEbitda = curr.kpis?.ebitda || 0;
        const realCaja = curr.cashflowDetail?.ending || 0;

        let cogs = curr.pnl?.categorias?.["Costo de Ventas"] || 0;
        let opex = 0;
        if (curr.pnl?.categorias?.OPEX) {
            opex = curr.pnl.categorias.OPEX;
        } else if (curr.pnl?.opexDetalle) {
            opex = -Math.abs(Object.values(curr.pnl.opexDetalle).reduce((acc, val) => acc + val, 0));
        } else {
            opex = realEbitda - realIngresos - cogs;
        }

        // Fallbacks in case everything is zero
        if (cogs === 0 && opex === 0 && realEbitda > 0) {
             cogs = -Math.abs(realIngresos * 0.4);
             opex = (realEbitda - realIngresos - cogs);
        }

        // Obtener porcentajes seleccionados por el usuario
        const pctVentasVol = parseInt(simVentasVolInp.value, 10) / 100;
        const pctPrecios = parseInt(simPreciosInp.value, 10) / 100;
        const pctCogs = parseInt(simCogsInp.value, 10) / 100;
        const pctOpex = parseInt(simOpexInp.value, 10) / 100;
        const extraDso = parseInt(simDsoInp.value, 10);

        // -------------- MOTOR MATEMÁTICO --------------
        // 1. Efecto en Ingresos: (Volumen * Precio)
        const simIngresosVolumen = realIngresos * (1 + pctVentasVol);
        const simIngresos = simIngresosVolumen * (1 + pctPrecios);
        
        // Lógica Mejorada: COGS es 100% variable con el VOLUMEN de ventas, no con el precio.
        // Para el OPEX, asumimos que un 40% es variable (logística, comisiones, etc.) y 60% es fijo.
        const variableOpexRatio = 0.4;
        
        // 1. Efecto Volumen (Crecen por las ventas - volumen)
        const cogsPorVolumen = cogs * (1 + pctVentasVol);
        const opexFijo = opex * (1 - variableOpexRatio);
        const opexVariablePorVolumen = opex * variableOpexRatio * (1 + pctVentasVol);
        
        // 2. Efecto Inflación/Eficiencia independiente para COGS y OPEX
        const simCogs = cogsPorVolumen * (1 + pctCogs);
        const simOpex = (opexFijo + opexVariablePorVolumen) * (1 + pctOpex);
        
        // varCostos será negativo si los costos suben (los costos ya son valores negativos)
        const varCostos = (simCogs - cogs) + (simOpex - opex); 
        
        // Nuevo EBITDA = Real Ebitda + Delta Ingresos + Delta Costos
        const simEbitda = realEbitda + (simIngresos - realIngresos) + varCostos;

        // 2. Simulación Caja (Impacto de Cuentas por Cobrar + delta EBITDA)
        // Cada día de DSO atrapa: (Ingresos Anualizados / 365) en capital de trabajo. (Aprox mensual: Ingresos Mensuales / 30)
        const dailySales = simIngresos / 30;
        const cashTrappedByDso = extraDso * dailySales;

        const deltaEbitda = simEbitda - realEbitda;
        
        // Nuevo Saldo de Caja = Caja Actual + (Aumento Ebitda) - (Efectivo retenido por más días de Cuentas por Cobrar)
        const simCaja = realCaja + deltaEbitda - cashTrappedByDso;
        // ----------------------------------------------

        // Renderizar Resultados
        document.getElementById('sim-base-ebitda').textContent = `Base Actual: RD$ ${realEbitda.toFixed(1)}M`;
        document.getElementById('sim-base-caja').textContent = `Base Actual: RD$ ${realCaja.toFixed(1)}M`;

        const resEbitdaEl = document.getElementById('sim-result-ebitda');
        const resCajaEl = document.getElementById('sim-result-caja');

        resEbitdaEl.textContent = `RD$ ${simEbitda.toFixed(1)}M`;
        resCajaEl.textContent = `RD$ ${simCaja.toFixed(1)}M`;

        resEbitdaEl.style.color = simEbitda >= realEbitda ? 'var(--success)' : 'var(--danger)';
        resCajaEl.style.color = simCaja >= realCaja ? 'var(--success)' : 'var(--danger)';

        // Render Comparative Table
        const tbody = document.getElementById('sim-comparison-tbody');
        if (tbody) {
            const getRowHTML = (label, isBold, baseVal, simVal, prefix = '') => {
                const diff = simVal - baseVal;
                // Format diff properly
                const diffColor = diff >= 0 ? 'var(--success)' : 'var(--danger)';
                  
                const diffText = diff === 0 ? '-' : `${diff > 0 ? '+' : ''}${formatCurrency(diff)}`;
                
                return `
                <tr style="${isBold ? 'font-weight: 700; background-color: #f8fafc;' : ''}">
                  <td style="padding:10px 12px; border-bottom: 1px solid #e2e8f0; color: #334155;">${prefix}${label}</td>
                  <td style="padding:10px 12px; text-align:right; border-bottom: 1px solid #e2e8f0; color: #475569;">${formatCurrency(baseVal)}</td>
                  <td style="padding:10px 12px; text-align:right; border-bottom: 1px solid #e2e8f0; color: #0f172a;">${formatCurrency(simVal)}</td>
                  <td style="padding:10px 12px; text-align:right; border-bottom: 1px solid #e2e8f0; color: ${diffColor}; font-weight: 600;">${diffText}</td>
                </tr>
                `;
            };

            const realUtilidadBruta = realIngresos + cogs;
            const simUtilidadBruta = simIngresos + simCogs;

            // Margins
            const mbReal = realIngresos > 0 ? (realUtilidadBruta/realIngresos)*100 : 0;
            const mbSim = simIngresos > 0 ? (simUtilidadBruta/simIngresos)*100 : 0;

            const mebReal = realIngresos > 0 ? (realEbitda/realIngresos)*100 : 0;
            const mebSim = simIngresos > 0 ? (simEbitda/simIngresos)*100 : 0;

            let html = `
                ${getRowHTML('Ingresos Brutos', true, realIngresos, simIngresos)}
                ${getRowHTML('Costos Directos (COGS)', false, cogs, simCogs, '&nbsp;&nbsp;')}
                ${getRowHTML(`Margen Bruto`, true, realUtilidadBruta, simUtilidadBruta)}
                <tr style="font-size: 0.75rem; background-color: #f1f5f9;">
                  <td colspan="4" style="padding:6px 12px; text-align:right; color: #64748b;">
                    % Margen Base: <strong>${mbReal.toFixed(1)}%</strong> &nbsp;&nbsp;|&nbsp;&nbsp; % Simulado: <strong>${mbSim.toFixed(1)}%</strong>
                  </td>
                </tr>
                ${getRowHTML('OPEX (Fijo + Variable)', false, opex, simOpex, '&nbsp;&nbsp;')}
                ${getRowHTML('EBITDA', true, realEbitda, simEbitda)}
                <tr style="font-size: 0.75rem; background-color: #f1f5f9;">
                  <td colspan="4" style="padding:6px 12px; text-align:right; color: #64748b;">
                    % Margen Base: <strong>${mebReal.toFixed(1)}%</strong> &nbsp;&nbsp;|&nbsp;&nbsp; % Simulado: <strong>${mebSim.toFixed(1)}%</strong>
                  </td>
                </tr>
                ${getRowHTML('Efectivo Total (Caja)', true, realCaja, simCaja)}
            `;
            tbody.innerHTML = html;
        }
        
        return { pctVentasVol, pctPrecios, pctCogs, pctOpex, extraDso, realEbitda, simEbitda, realCaja, simCaja };
    };

    const updateLabels = () => {
        if(labelVentasVol) labelVentasVol.textContent = simVentasVolInp.value + '%';
        if(labelPrecios) labelPrecios.textContent = simPreciosInp.value + '%';
        if(labelCogs) labelCogs.textContent = simCogsInp.value + '%';
        if(labelOpex) labelOpex.textContent = simOpexInp.value + '%';
        if(labelDso) labelDso.textContent = simDsoInp.value + ' Días';
        window.runSimulationMath();
    };

    if (simVentasVolInp) simVentasVolInp.addEventListener('input', updateLabels);
    if (simPreciosInp) simPreciosInp.addEventListener('input', updateLabels);
    if (simCogsInp) simCogsInp.addEventListener('input', updateLabels);
    if (simOpexInp) simOpexInp.addEventListener('input', updateLabels);
    if (simDsoInp) simDsoInp.addEventListener('input', updateLabels);

    const btnResetSim = document.getElementById('btn-reset-simulation');
    if (btnResetSim) {
        btnResetSim.addEventListener('click', () => {
            if (simVentasVolInp) simVentasVolInp.value = 0;
            if (simPreciosInp) simPreciosInp.value = 0;
            if (simCogsInp) simCogsInp.value = 0;
            if (simOpexInp) simOpexInp.value = 0;
            if (simDsoInp) simDsoInp.value = 0;
            
            const aiInsightEl = document.getElementById('sim-ai-insight');
            if (aiInsightEl) {
                if (!window.aiEnabled) {
                     aiInsightEl.innerHTML = '<em>Funciones avanzadas deshabilitadas. Habilítelas en Configuración para ver insights estratégicos.</em>';
                } else {
                     aiInsightEl.innerHTML = '<em>Genera un insight de IA ejecutando una simulación.</em>';
                }
            }
            updateLabels();
        });
    }

    const btnRunSim = document.getElementById('btn-run-simulation');
    window.simSummaryCache = {};

    if (btnRunSim) {
        btnRunSim.addEventListener('click', async () => {
            if (!globalFinancialData || globalFinancialData.length === 0) {
                alert("Por favor, sube los datos financieros primero.");
                return;
            }

            const monthSelector = document.getElementById('monthSelector');
            const idx = monthSelector ? parseInt(monthSelector.value, 10) : globalFinancialData.length - 1;
            const curr = globalFinancialData[idx || globalFinancialData.length - 1];

            const simData = window.runSimulationMath();
            if (!simData) return;
            const { pctVentasVol, pctPrecios, pctCogs, pctOpex, extraDso, realEbitda, simEbitda, realCaja, simCaja } = simData;

            // Generar Insight IA
            const simInsightEl = document.getElementById('sim-ai-insight');
            
            if (!window.aiEnabled) {
                 simInsightEl.innerHTML = '<em>Funciones avanzadas deshabilitadas. Habilítelas en Configuración para ver insights estratégicos.</em>';
                 return;
            }

            simInsightEl.innerHTML = '<em><i data-lucide="loader" class="spin-icon"></i> Generando Insight Estratégico...</em>';
            lucide.createIcons();

            // Cache check
            const cacheKey = `vv${pctVentasVol}_vp${pctPrecios}_cc${pctCogs}_co${pctOpex}_d${extraDso}_m${curr.date || 'base'}`;
            if (window.simSummaryCache[cacheKey]) {
                simInsightEl.innerHTML = window.simSummaryCache[cacheKey];
                lucide.createIcons();
                return;
            }

            try {
                const simContext = `
El usuario simuló las siguientes variaciones en el mes actual (${curr.date}):
- Crecimiento de Ventas (Volumen): ${(pctVentasVol * 100).toFixed(0)}%
- Incremento de Precios: ${(pctPrecios * 100).toFixed(0)}%
- Eficiencia/Inflación COGS: ${(pctCogs * 100).toFixed(0)}%
- Eficiencia/Inflación OPEX: ${(pctOpex * 100).toFixed(0)}%
- Aumento de Días de Cobro (DSO): ${extraDso} días

Resultados calculados matemáticamente:
- EBITDA Base: RD$ ${realEbitda.toFixed(2)}M -> Simulado: RD$ ${simEbitda.toFixed(2)}M
- Caja Base: RD$ ${realCaja.toFixed(2)}M -> Simulada: RD$ ${simCaja.toFixed(2)}M

Redacta UNA SOLA ORACIÓN para el CFO de advertencia o recomendación estratégica. Ejemplo: "Este aumento en ventas drenará tu liquidez en RD$ 15M debido al relajamiento de los cobros comerciales."
                `;

                let timeoutId;
                const timeoutPromise = new Promise((_, reject) => {
                    timeoutId = setTimeout(() => reject(new Error('AI Request Timeout (45s)')), 45000);
                });
                let apiCallPromise;
                try {
                    apiCallPromise = getAI().models.generateContent({
                        model: "gemini-2.5-flash",
                        contents: simContext,
                    });
                    apiCallPromise.catch(err => window.handleAiError("Sim", err));
                } catch (err) {
                    apiCallPromise = Promise.reject(err);
                    apiCallPromise.catch(()=> /* handled */ {});
                }
                
                let response;
                try {
                    response = await Promise.race([apiCallPromise, timeoutPromise]);
                } finally {
                    clearTimeout(timeoutId);
                }

                const finalHtml = `<strong><i data-lucide="sparkles" style="display: inline; width: 16px; height: 16px; vertical-align: text-bottom; margin-right: 4px;"></i> Insight Bot:</strong> ${response.text}`;
                simInsightEl.innerHTML = finalHtml;
                window.simSummaryCache[cacheKey] = finalHtml;
                lucide.createIcons();

            } catch (err) {
                window.handleAiError("Sim", err);
                simInsightEl.innerHTML = `<em>IA no disponible por el momento.</em>`;
            }
        });
    }
    
    async function loadVentasCeoData(token) {
        if (!SHARPOINT_VENTAS_FILE_URL || !token) {
            console.warn('Ventas CEO: SHARPOINT_VENTAS_FILE_URL o token no configurado. Intentando carga local...');
            try {
                const res = await fetch('/public/ventasCEO.csv');
                if(!res.ok) {
                    console.warn('Ventas CEO CSV not found');
                    return;
                }
                const csvText = await res.text();
                const lines = csvText.split(/\r?\n/).map(l => l.split(','));
                window.processVentasCeoWorkbook(null, lines);
            } catch(e) {
                console.error("Error loading ventas CEO local", e);
            }
            return;
        }

        try {
            const encodedUrl = btoa(SHARPOINT_VENTAS_FILE_URL).replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
            const graphUrl = `https://graph.microsoft.com/v1.0/shares/u!${encodedUrl}/driveItem/content`;
            
            const req = await fetch(graphUrl, { headers: { "Authorization": `Bearer ${token}` }});
            if (req.ok) {
                const arrayBuffer = await req.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                let bestSheetName = workbook.SheetNames[0];
                const consejoSheet = workbook.SheetNames.find(n => n.toLowerCase().includes('consejo'));
                if (consejoSheet) {
                    bestSheetName = consejoSheet;
                } else {
                    let maxScore = -1;
                    for (let name of workbook.SheetNames) {
                        const sheetTmp = workbook.Sheets[name];
                        const rowsTmp = XLSX.utils.sheet_to_json(sheetTmp, { header: 1 });
                        let score = 0;
                        for (let r of rowsTmp) {
                            if (!r) continue;
                            for (let c of r) {
                                if (c === undefined || c === null) continue;
                                const term = String(c).toLowerCase().trim();
                                if (term === 'producto' || term === 'descripción' || term === 'descripcion') score += 10;
                                if (term === 'tipo') score += 5;
                                if (term.includes('ventas netas dop') || term.includes('ventas netas')) score += 8;
                                if (term.includes('volumen unidades') || term.includes('volumen')) score += 8;
                                if (term === '2026' || term.includes('ppto')) score += 5;
                            }
                        }
                        if (score > maxScore && score > 0) {
                            maxScore = score;
                            bestSheetName = name;
                        }
                    }
                }
                window.processVentasCeoWorkbook(workbook);
            } else {
                 console.error("Error fetching Ventas CEO from OneDrive", req.statusText);
            }
        } catch(e) {
            console.error("Error loading Ventas CEO from OneDrive", e);
        }
    }
    
    window.processVentasCeoWorkbook = async function(workbook, fallbackLines, workerResult) {
        ceoData = []; // Vaciado de estado para evitar duplicados
        if (fallbackLines) {
            ceoData = parseConsejoSheet(fallbackLines);
            // Save to IndexedDB
            try {
                const db = await getFinanceDB();
                await new Promise((resolve, reject) => {
                    const tx = db.transaction('finance_cache', 'readwrite');
                    tx.objectStore('finance_cache').put({ data: ceoData, timestamp: Date.now() }, 'CEO_VENTAS_KEY_V3');
                    tx.oncomplete = resolve;
                    tx.onerror = reject;
                });
            } catch (err) {
                console.warn("⚠️ Error saving Ventas CEO to IndexedDB:", err);
            }
            window.hasVentasAccess = true;
            if (typeof window.applyRoleBasedUI === 'function') {
                window.applyRoleBasedUI(window.hasMasterAccess, true, window.hasComercialAccess);
            }
            let viewVentasCeo = document.getElementById("view-ventas-ceo");
            if (viewVentasCeo && viewVentasCeo.classList.contains("active")) {
                window.renderVentasCEO();
            }
            return;
        }
        
        try {
            // Find sheets
            let consejoSheetName, dataSheetName, bestSheetName;
            
            if (workerResult) {
                consejoSheetName = workerResult.consejoSheetName;
                dataSheetName = workerResult.dataSheetName;
                bestSheetName = workerResult.bestSheetName;
            } else {
                consejoSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('consejo'));
                dataSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('data por mes'));
            }
            
            if (!consejoSheetName) {
                let bestRows;
                if (workerResult) {
                     bestRows = workerResult.bestRows;
                } else {
                    bestSheetName = workbook.SheetNames[0];
                    let maxScore = -1;
                    for (let name of workbook.SheetNames) {
                        const sheetTmp = workbook.Sheets[name];
                        const rowsTmp = XLSX.utils.sheet_to_json(sheetTmp, { header: 1 });
                        let score = 0;
                        for (let r of rowsTmp) {
                            if (!r) continue;
                            for (let c of r) {
                                if (c === undefined || c === null) continue;
                                const term = String(c).toLowerCase().trim();
                                if (term === 'producto' || term === 'descripción' || term === 'descripcion') score += 10;
                                if (term === 'tipo') score += 5;
                                if (term.includes('ventas netas dop') || term.includes('ventas netas')) score += 8;
                                if (term.includes('volumen unidades') || term.includes('volumen')) score += 8;
                                if (term === '2026' || term.includes('ppto')) score += 5;
                            }
                        }
                        if (score > maxScore && score > 0) {
                            maxScore = score;
                            bestSheetName = name;
                        }
                    }
                    bestRows = XLSX.utils.sheet_to_json(workbook.Sheets[bestSheetName], { header: 1 });
                }
                ceoData = parseConsejoSheet(bestRows);
                if (!ceoData || ceoData.length === 0) {
                    const selectedSheetName = bestSheetName || (workbook?.SheetNames ? workbook.SheetNames[0] : null);
                    console.warn("Ventas CEO produjo 0 filas. Diagnóstico (no consejoSheetName):", {
                        sheetNames: workbook?.SheetNames,
                        selectedSheetName: selectedSheetName,
                        firstRows: bestRows ? bestRows.slice(0, 10) : null,
                        detectedHeaders: bestRows ? (bestRows[0] || []) : [],
                        rawRowsLength: bestRows ? bestRows.length : 0,
                        parsedLength: 0
                    });
                    ceoData = null;
                    let viewVentasCeo = document.getElementById("view-ventas-ceo");
                    if (viewVentasCeo && viewVentasCeo.classList.contains("active")) {
                        window.renderVentasCEO();
                    }
                    return;
                }
                // Save to IndexedDB
                try {
                    const db = await getFinanceDB();
                    await new Promise((resolve, reject) => {
                        const tx = db.transaction('finance_cache', 'readwrite');
                        tx.objectStore('finance_cache').put({ data: ceoData, timestamp: Date.now() }, 'CEO_VENTAS_KEY_V3');
                        tx.oncomplete = resolve;
                        tx.onerror = reject;
                    });
                } catch (err) {
                    console.warn("⚠️ Error saving Ventas CEO to IndexedDB:", err);
                }
                let viewVentasCeo = document.getElementById("view-ventas-ceo");
                if (viewVentasCeo && viewVentasCeo.classList.contains("active")) {
                    window.renderVentasCEO();
                }
                return;
            }

            // Both sheets are available or at least Consejo is
            // Start by extracting "Tablas Consejo" for Volumen
            let rawConsejoObjects;
            if (workerResult) {
                rawConsejoObjects = workerResult.consejoRows;
            } else {
                const consejoSheet = workbook.Sheets[consejoSheetName];
                rawConsejoObjects = XLSX.utils.sheet_to_json(consejoSheet, { range: 2, defval: 0 });
            }
            
            const tempParsedRows = parseConsejoFromObjects(rawConsejoObjects);
            let baseHierarchy = tempParsedRows.filter(d => d.Tipo === 'Volumen');
            let finalData = [];

            if (dataSheetName) {
                // If "data por mes" exists, compute Volumen, Monto and Precio Unitario
                let dataRows;
                if (workerResult) {
                    dataRows = workerResult.dataRows;
                } else {
                    const dataSheet = workbook.Sheets[dataSheetName];
                    dataRows = XLSX.utils.sheet_to_json(dataSheet, { header: 1 });
                }
                const detailedData = extractDetailedData(dataRows);
                
                const categories = [...new Set(baseHierarchy.map(r => r.Producto))];
                
                categories.forEach(cat => {
                    const hierarchyRow = baseHierarchy.find(d => d.Producto === cat);
                    if(hierarchyRow) {
                        // Compute Volumen 
                        const newVolRow = computeCategoryFromDetailed(cat, detailedData, hierarchyRow, 'Volumen');
                        finalData.push(newVolRow);
                        
                        // Compute Monto 
                        const montoRow = computeCategoryFromDetailed(cat, detailedData, hierarchyRow, 'Monto (MM DOP)');
                        finalData.push(montoRow);
                        
                        // Compute Precio Unitario
                        let precioRow = { 
                            Producto: cat, 
                            Tipo: "Precio Unitario", 
                            hasChildren: hierarchyRow.hasChildren,
                            parentId: hierarchyRow.parentId,
                            id: hierarchyRow.id,
                            values: {},
                            pptoValues: {}
                        };
                        
                        Object.keys(newVolRow.values || {}).forEach(k => {
                            let volVal = newVolRow.values[k] || 0;
                            let montoVal = montoRow.values[k] || 0;
                            precioRow.values[k] = volVal ? ((montoVal * 1000000) / volVal) : 0;
                        });
                        
                        Object.keys(newVolRow.pptoValues || {}).forEach(k => {
                            let volVal = newVolRow.pptoValues[k] || 0;
                            let montoVal = montoRow.pptoValues[k] || 0;
                            precioRow.pptoValues[k] = volVal ? ((montoVal * 1000000) / volVal) : 0;
                        });
                        
                        ['FY2024', 'PO25', 'PO26'].forEach(y => {
                            let volVal = newVolRow[y] || 0;
                            let montoVal = montoRow[y] || 0;
                            precioRow[y] = volVal ? ((montoVal * 1000000) / volVal) : 0;
                        });
                        
                        finalData.push(precioRow);
                    }
                });

            } else {
                // fallback 
                finalData = tempParsedRows; 
            }
            
            // 1. Recalcular Padres (las filas que funcionan como 'padres' de los grupos deben contemplarse como la suma del valor de los 'hijos')
            const parentIds = [...new Set(finalData.filter(d => d.hasChildren).map(d => d.id))];
            parentIds.forEach(pId => {
                ['Volumen', 'Monto (MM DOP)'].forEach(tipo => {
                    let parentRow = finalData.find(d => d.id === pId && d.Tipo === tipo);
                    if (parentRow) {
                        let children = finalData.filter(d => d.parentId === pId && d.Tipo === tipo && d.Producto !== "PA H+ 0.68 LTS (X12)");
                        // Limpiar valores del padre para sumar
                        Object.keys(parentRow.values).forEach(k => parentRow.values[k] = 0);
                        if(parentRow.pptoValues) Object.keys(parentRow.pptoValues).forEach(k => parentRow.pptoValues[k] = 0);
                        else parentRow.pptoValues = {};
                        
                        ['FY2024', 'PO25', 'PO26'].forEach(y => { if(parentRow[y] !== undefined) parentRow[y] = 0; });
                        
                        children.forEach(c => {
                            Object.keys(c.values).forEach(k => {
                                parentRow.values[k] = (parentRow.values[k] || 0) + (c.values[k] || 0);
                            });
                            if(c.pptoValues) {
                                Object.keys(c.pptoValues).forEach(k => {
                                    parentRow.pptoValues[k] = (parentRow.pptoValues[k] || 0) + (c.pptoValues[k] || 0);
                                });
                            }
                            ['FY2024', 'PO25', 'PO26'].forEach(y => {
                                parentRow[y] = (parentRow[y] || 0) + (c[y] || 0);
                            });
                        });
                    }
                });
                
                // Recalcular Precio Unitario para el padre = Monto / Volumen
                let parentPrecio = finalData.find(d => d.id === pId && d.Tipo === 'Precio Unitario');
                let parentVol = finalData.find(d => d.id === pId && d.Tipo === 'Volumen');
                let parentMonto = finalData.find(d => d.id === pId && d.Tipo === 'Monto (MM DOP)');
                
                if (!parentPrecio && parentVol && parentMonto) {
                    parentPrecio = { Producto: parentVol.Producto, Tipo: 'Precio Unitario', values: {}, pptoValues: {} };
                    parentPrecio.id = parentVol.id;
                    parentPrecio.hasChildren = parentVol.hasChildren;
                    parentPrecio.parentId = parentVol.parentId;
                    finalData.push(parentPrecio);
                }
                
                if (parentPrecio && parentVol && parentMonto) {
                    parentPrecio.pptoValues = {};
                    let allKeys = new Set([...Object.keys(parentVol.values || {}), ...Object.keys(parentVol.pptoValues || {})]);
                    allKeys.forEach(k => {
                        let volVal = parentVol.values[k] || 0;
                        let montoVal = parentMonto.values[k] || 0;
                        parentPrecio.values[k] = volVal ? ((montoVal * 1000000) / volVal) : 0;
                        
                        let volPpto = parentVol.pptoValues?.[k] || 0;
                        let montoPpto = parentMonto.pptoValues?.[k] || 0;
                        parentPrecio.pptoValues[k] = volPpto ? ((montoPpto * 1000000) / volPpto) : 0;
                    });
                    ['FY2024', 'PO25', 'PO26'].forEach(y => {
                        let volVal = parentVol[y] || 0;
                        let montoVal = parentMonto[y] || 0;
                        parentPrecio[y] = volVal ? ((montoVal * 1000000) / volVal) : 0;
                    });
                }
            });

            // 2. Recalcular TOTAL y TOTAL SIN BON si están en el hierarchy
            ['TOTAL', 'TOTAL SIN BON'].forEach(tot => {
               ['Volumen', 'Monto (MM DOP)'].forEach(tipo => {
                   let totRow = finalData.find(d => d.Producto === tot && d.Tipo === tipo);
                   if(totRow) {
                       totRow.pptoValues = totRow.pptoValues || {};
                       // Las filas raíz
                       const mainItems = finalData.filter(d => d.Tipo === tipo && ['AGUA PLANETA AZUL', 'MAQUILAS', 'BEBIDAS'].includes(d.Producto.toUpperCase().trim()));
                       // Las filas BONIF o BON (son hijos)
                       const bonifItems = finalData.filter(d => d.parentId && d.Tipo === tipo && d.Producto.includes('BON'));
                       
                       // Collect all possible keys
                       let allKeys = new Set();
                       mainItems.forEach(d => {
                           if(d.values) Object.keys(d.values).forEach(k => allKeys.add(k));
                           if(d.pptoValues) Object.keys(d.pptoValues).forEach(k => allKeys.add(k));
                       });
                       bonifItems.forEach(d => {
                           if(d.values) Object.keys(d.values).forEach(k => allKeys.add(k));
                           if(d.pptoValues) Object.keys(d.pptoValues).forEach(k => allKeys.add(k));
                       });

                       allKeys.forEach(k => {
                           let sum = 0;
                           let sumPpto = 0;
                           mainItems.forEach(d => {
                               sum += (d.values[k] || 0);
                               sumPpto += (d.pptoValues?.[k] || 0);
                           });
                           
                           if (tot === 'TOTAL SIN BON') {
                               // Restar bonificaciones
                               bonifItems.forEach(d => {
                                   sum -= (d.values[k] || 0);
                                   sumPpto -= (d.pptoValues?.[k] || 0);
                               });
                           }
                           
                           totRow.values[k] = sum;
                           totRow.pptoValues[k] = sumPpto;
                       });
                       
                       ['FY2024', 'PO25', 'PO26'].forEach(y => {
                           let sum = 0;
                           let sumPpto = 0;
                           mainItems.forEach(d => {
                               sum += (d[y] || 0);
                               sumPpto += (d.pptoValues?.[y] || 0); // Note: FY2024 etc. are direct properties, not in pptoValues, wait!
                           });
                           
                           if (tot === 'TOTAL SIN BON') {
                               bonifItems.forEach(d => {
                                   sum -= (d[y] || 0);
                                   sumPpto -= (d.pptoValues?.[y] || 0);
                               });
                           }
                           
                           totRow[y] = sum;
                           // we don't have totRow.pptoValues[y]
                       });
                   }
               });
               
               // Precio Unitario para totales
               let totPrecio = finalData.find(d => d.Producto === tot && d.Tipo === 'Precio Unitario');
               let totVol = finalData.find(d => d.Producto === tot && d.Tipo === 'Volumen');
               let totMonto = finalData.find(d => d.Producto === tot && d.Tipo === 'Monto (MM DOP)');
               
               if (!totPrecio && totVol && totMonto) {
                   totPrecio = { Producto: tot, Tipo: 'Precio Unitario', values: {}, pptoValues: {} };
                   if (totVol.id) totPrecio.id = totVol.id;
                   if (totVol.hasChildren !== undefined) totPrecio.hasChildren = totVol.hasChildren;
                   if (totVol.parentId) totPrecio.parentId = totVol.parentId;
                   finalData.push(totPrecio);
               }
               
               if (totPrecio && totVol && totMonto) {
                    totPrecio.pptoValues = totPrecio.pptoValues || {};
                    let allKeys = new Set([...Object.keys(totVol.values || {}), ...Object.keys(totVol.pptoValues || {})]);
                    allKeys.forEach(k => {
                        let volVal = totVol.values[k] || 0;
                        let montoVal = totMonto.values[k] || 0;
                        totPrecio.values[k] = volVal ? ((montoVal * 1000000) / volVal) : 0;
                        
                        let volPpto = totVol.pptoValues?.[k] || 0;
                        let montoPpto = totMonto.pptoValues?.[k] || 0;
                        totPrecio.pptoValues[k] = volPpto ? ((montoPpto * 1000000) / volPpto) : 0;
                    });
                    ['FY2024', 'PO25', 'PO26'].forEach(y => {
                        let volVal = totVol[y] || 0;
                        let montoVal = totMonto[y] || 0;
                        totPrecio[y] = volVal ? ((montoVal * 1000000) / volVal) : 0;
                    });
               }
            });
            
        // Dividir el Volumen entre 1000 para llevar a miles (k)
            finalData.forEach(row => {
                if (row.Tipo === 'Volumen') {
                    if (row.values) {
                        Object.keys(row.values).forEach(k => row.values[k] /= 1000);
                    }
                    if (row.pptoValues) {
                        Object.keys(row.pptoValues).forEach(k => row.pptoValues[k] /= 1000);
                    }
                    ['PO25', 'PO26'].forEach(y => {
                        if(row[y] !== undefined) row[y] /= 1000;
                    });
                     // FY2024 will be restored below, so no need to divide it here
                }
            });

        // Restaurar FY2024 desde Tablas Consejo para TODOS los productos y tipos
        // ya que el valor allí es un promedio anual y debe mantenerse fijo, 
        // y ya viene formateado en la escala correcta (k, o MM DOP)
        finalData.forEach(row => {
            let consejoRow = tempParsedRows.find(d => {
                if (d.Producto !== row.Producto) return false;
                if (row.Tipo === 'Monto (MM DOP)') {
                    let cType = String(d.Tipo).toUpperCase();
                    return cType.includes('MONTO') || cType.includes('VALOR') || cType.includes('VENTAS') || d.Tipo === row.Tipo;
                }
                return d.Tipo === row.Tipo;
            });
            if (consejoRow && consejoRow['FY2024'] !== undefined) {
                row['FY2024'] = consejoRow['FY2024'];
                row.__real24 = consejoRow['FY2024'];
            }
        });

            if (!finalData || finalData.length === 0) {
                const selectedSheetName = consejoSheetName || bestSheetName || dataSheetName || (workbook?.SheetNames ? workbook.SheetNames[0] : null);
                let rawRows = null;
                if (selectedSheetName && workbook?.Sheets?.[selectedSheetName]) {
                    rawRows = XLSX.utils.sheet_to_json(workbook.Sheets[selectedSheetName], { header: 1 });
                }
                
                console.warn("Ventas CEO produjo 0 filas. Diagnóstico:", {
                    sheetNames: workbook?.SheetNames,
                    selectedSheetName: selectedSheetName,
                    firstRows: rawRows ? rawRows.slice(0, 10) : null,
                    detectedHeaders: rawRows ? (rawRows[0] || []) : [],
                    rawRowsLength: rawRows ? rawRows.length : 0,
                    parsedLength: 0
                });
                ceoData = null;
                let viewVentasCeo = document.getElementById("view-ventas-ceo");
                if (viewVentasCeo && viewVentasCeo.classList.contains("active")) {
                    window.renderVentasCEO();
                }
                return;
            }

            ceoData = finalData;
            console.log("Ventas CEO data loaded.", ceoData.length);
            
            // Save to IndexedDB
            try {
                const db = await getFinanceDB();
                await new Promise((resolve, reject) => {
                    const tx = db.transaction('finance_cache', 'readwrite');
                    tx.objectStore('finance_cache').put({ data: ceoData, timestamp: Date.now() }, 'CEO_VENTAS_KEY_V3');
                    tx.oncomplete = resolve;
                    tx.onerror = reject;
                });
            } catch (err) {
                console.warn("⚠️ Error saving Ventas CEO to IndexedDB:", err);
            }
            
            window.hasVentasAccess = true;
            if (typeof window.applyRoleBasedUI === 'function') {
                window.applyRoleBasedUI(window.hasMasterAccess, true, window.hasComercialAccess);
            }

            let viewVentasCeo = document.getElementById("view-ventas-ceo");
            if (viewVentasCeo && viewVentasCeo.classList.contains("active")) {
                window.renderVentasCEO();
            }
        } catch(e) {
            console.error("Error formatting Ventas CEO workbook", e);
        }
    }

    function computeCategoryFromDetailed(category, dataRows, hierarchyRow, targetMetric) {
        let resultRow = { 
            Producto: category, 
            Tipo: targetMetric, 
            hasChildren: hierarchyRow.hasChildren, 
            parentId: hierarchyRow.parentId, 
            id: hierarchyRow.id, 
            values: {} 
        };
        const catUpper = category.toUpperCase().trim();
        
        // Accumulate data
        let accumulated = {};
        let accumulated_ppto = {};
        let fyAccum = { 'FY2024': 0, 'PO25': 0, 'PO26': 0 };

        let foundAny = false;
        
        dataRows.forEach(row => {
            if(!row.Concepto && !row.Producto && !row['Descripción']) return;
            const desc = String(row.Concepto || row.Producto || row['Descripción']).toUpperCase().trim();
            const isMatch = desc === catUpper || (catUpper.startsWith('APA ') && desc.includes(catUpper)) || desc.includes(catUpper);
            
            let isMetricMatch = false;
            const metricStr = String(row.Tipo || row.Métrica || '').toUpperCase();
            
            if (targetMetric === 'Volumen') {
                isMetricMatch = metricStr.includes('VOLUMEN') || metricStr.includes('CAJA') || metricStr.includes('UNIDAD');
            } else {
                isMetricMatch = metricStr.includes('MONTO') || metricStr.includes('VENTAS') || metricStr.includes('VALOR');
            }
            
            if (isMatch && isMetricMatch) {
                foundAny = true;
                
                Object.keys(row.values || {}).forEach(k => {
                    accumulated[k] = (accumulated[k] || 0) + (row.values[k] || 0);
                    if (row.isPpto && row.isPpto[k]) {
                        if (!resultRow.isPpto) resultRow.isPpto = {};
                        resultRow.isPpto[k] = true;
                    }
                });
                Object.keys(row.pptoValues || {}).forEach(k => {
                    accumulated_ppto[k] = (accumulated_ppto[k] || 0) + (row.pptoValues[k] || 0);
                });
                
                ['FY2024', 'PO25', 'PO26'].forEach(y => {
                    fyAccum[y] += (row[y] || 0);
                });
            }
        });
        
        resultRow.pptoValues = {};
        if (foundAny) {
            // Apply divided by 1M ONLY for Monto
            const divisor = targetMetric === 'Monto (MM DOP)' ? 1000000 : 1;
            Object.keys(accumulated).forEach(k => {
                resultRow.values[k] = (accumulated[k] || 0) / divisor;
            });
            Object.keys(accumulated_ppto).forEach(k => {
                resultRow.pptoValues[k] = (accumulated_ppto[k] || 0) / divisor;
            });
            ['FY2024', 'PO25', 'PO26'].forEach(y => {
                resultRow[y] = fyAccum[y] / divisor;
            });
        } else if (targetMetric === 'Volumen') {
            // Fallback for Volumen: just use the data from hierarchyRow (Tablas Consejo)
            if (hierarchyRow && hierarchyRow.values) {
                Object.keys(hierarchyRow.values).forEach(k => {
                    resultRow.values[k] = hierarchyRow.values[k];
                });
                ['FY2024', 'PO25', 'PO26'].forEach(y => {
                    resultRow[y] = hierarchyRow[y] || 0;
                });
            }
        }
        
        // Initialize zeros for keys present in hierarchyRow but missing
        if(hierarchyRow && hierarchyRow.values) {
           Object.keys(hierarchyRow.values).forEach(k => {
               if(resultRow.values[k] === undefined) resultRow.values[k] = 0;
           });
           ['FY2024', 'PO25', 'PO26'].forEach(y => {
                if(resultRow[y] === undefined) resultRow[y] = 0;
           });
        }
        
        return resultRow;
    }
    
    function extractDetailedData(rows) {
        if(rows.length < 2) return [];
        let headerRowIdx = rows.findIndex(row => row && row.some(cell => {
             const text = String(cell).toLowerCase().trim();
             return text === 'concepto' || text === 'producto' || text === 'descripción';
        }));
        if(headerRowIdx === -1) headerRowIdx = 0;
        const headers = rows[headerRowIdx] || [];
        // Ensure standard names
        headers.forEach((h, idx) => {
             if(!h) return;
             const text = String(h).toLowerCase().trim();
             if(text === 'concepto' || text === 'producto' || text === 'descripción') headers[idx] = 'Concepto';
             if(text === 'tipo' || text === 'métrica') headers[idx] = 'Tipo';
        });
        
        let maxRowLength = Math.max(...rows.map(r => r ? r.length : 0));
        for(let i=0; i<maxRowLength; i++) {
             if(!headers[i]) headers[i] = `Col_${i}`;
        }
        
        let pptoStartIdx = 9999;
        const scanStart = Math.max(0, headerRowIdx - 2);
        for (let i = scanStart; i <= headerRowIdx; i++) {
            const r = rows[i];
            if (!r) continue;
            r.forEach((cell, idx) => {
                if (!cell) return;
                const s = String(cell).toUpperCase().trim();
                if (s === 'PPTO' || s === 'PRESUPUESTO' || s.includes('PPTO 2026') || s.includes('PRESUP')) {
                    if (idx < pptoStartIdx) pptoStartIdx = idx;
                }
            });
        }
        if (pptoStartIdx === 9999) {
            headers.forEach((cell, idx) => {
                if (cell) {
                    const s = String(cell).toUpperCase().trim();
                    if (s.includes('PPTO') || s.includes('PRESUPUESTO')) {
                        if (idx < pptoStartIdx) pptoStartIdx = idx;
                    }
                }
            });
        }
        if (pptoStartIdx === 9999) pptoStartIdx = 39;

        const monthsMap = {
            'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
            'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12'
        };

        let data = [];
        for(let i = headerRowIdx + 1; i < rows.length; i++) {
             const row = rows[i];
             if(!row) continue;
             
             let obj = { values: {}, pptoValues: {} };
             headers.forEach((h, idx) => {
                 if(row[idx] === undefined) return;
                 let val = parseFloat(String(row[idx]).replace(/,/g, '')) || 0;
                 
                 let isDate = false;
                 let dateStr = "";
                 const headerText = String(h).trim();
                 const textK = headerText.toLowerCase();

                 if (headerText.includes('-01 00:00:00') || (headerText.startsWith('202') && headerText.length >= 7)) {
                     isDate = true;
                     dateStr = headerText.slice(0, 7);
                 } else if (!isNaN(headerText) && Number(headerText) > 40000 && Number(headerText) < 50000) {
                     isDate = true;
                     let dObj = new Date((Number(headerText) - 25569) * 86400 * 1000);
                     dateStr = dObj.toISOString().slice(0, 7);
                 } else {
                     const match = textK.match(/([a-z]{3})[-/ ]?(\d{2,4})/);
                     if(match && monthsMap[match[1]]) {
                         let y = match[2];
                         if(y.length === 2) y = "20" + y;
                         dateStr = `${y}-${monthsMap[match[1]]}`;
                         isDate = true;
                     }
                 }

                 if (isDate) {
                     const isPptoCol = idx >= pptoStartIdx || textK.includes('ppto') || textK.includes('presupuesto');
                     if (isPptoCol) {
                         if (!obj.pptoValues[dateStr]) obj.pptoValues[dateStr] = val;
                     } else {
                         if (!obj.values[dateStr]) obj.values[dateStr] = val;
                     }
                 } else {
                     if (headerText === 'Concepto' || headerText === 'Tipo') {
                         obj[headerText] = row[idx];
                     } else if(textK.includes('fy') || textK.includes('real 2024') || headerText === 'FY2024' || headerText === 'PO26' || headerText === 'PO25') {
                         let yKey = 'FY2024';
                         if(textK.includes('po26')) yKey = 'PO26';
                         else if(textK.includes('po25')) yKey = 'PO25';
                         obj[yKey] = val;
                     }
                 }
             });
             
             // Fallback Real 2026 to PPTO if Real is 0
             for (let m = 1; m <= 12; m++) {
                 let dateStr = `2026-${String(m).padStart(2, '0')}`;
                 let realVal = obj.values[dateStr] || 0;
                 let pptoVal = obj.pptoValues[dateStr] || 0;
                 if (realVal === 0 && pptoVal !== 0) {
                     obj.values[dateStr] = pptoVal;
                     if (!obj.isPpto) obj.isPpto = {};
                     obj.isPpto[dateStr] = true;
                 }
             }
             
             let po26Sum = 0;
             for (let m = 1; m <= 12; m++) {
                 let dateStr = `2026-${String(m).padStart(2, '0')}`;
                 po26Sum += obj.pptoValues[dateStr] || 0;
             }
             obj['PO26'] = po26Sum;
             
             if(obj.Concepto) data.push(obj);
        }
        return data;
    }

    function parseConsejoFromObjects(objectsInput) {
        if (!objectsInput || objectsInput.length === 0) return [];

        let objects = objectsInput;
        
        // 1. AUTO-DETECTOR Y LIMPIEZA DE MATRICES
        if (Array.isArray(objectsInput[0])) {
            let totalRowIndex = -1;
            
            for (let i = 0; i < objectsInput.length; i++) {
                if (!objectsInput[i]) continue;
                let hasTotal = objectsInput[i].some(c => c && String(c).toUpperCase().trim() === 'TOTAL');
                if (hasTotal) {
                    totalRowIndex = i;
                    break;
                }
            }
            
            if (totalRowIndex !== -1) {
                let headerRowIndex = totalRowIndex > 0 ? totalRowIndex - 1 : 0;
                let headers = objectsInput[headerRowIndex];
                let constructedObjects = [];
                
                for (let i = totalRowIndex; i < objectsInput.length; i++) {
                    if (!objectsInput[i] || objectsInput[i].length === 0) continue;
                    let obj = {};
                    for (let j = 0; j < objectsInput[i].length; j++) {
                        let key = (headers[j] !== undefined && headers[j] !== null && String(headers[j]).trim() !== '') 
                                    ? String(headers[j]).trim() 
                                    : `__GHOST_${j}`; // Renombrado para no confundir con __EMPTY
                        obj[key] = objectsInput[i][j];
                    }
                    constructedObjects.push(obj);
                }
                objects = constructedObjects;
            }
        }

        let parsedRows = [];
        let currentType = "Volumen";
        let currentParentId = null;
        let totalCount = 0;
        
        objects.forEach(r => {
            const keys = Object.keys(r);
            if(keys.length === 0) return;
            
            // 2. CORRECCIÓN CRÍTICA: Ignorar fantasmas y forzar columna inicial (keys[0])
            let prodKey = keys.find(k => k.toLowerCase().includes('producto') || k.toLowerCase().includes('descrip'));
            if (!prodKey) prodKey = keys[0]; 
            
            let tipoKey = keys.find(k => k.toLowerCase().includes('tipo'));

            let prodVal = String(r[prodKey] || '').trim();
            if(!prodVal || prodVal === '0') return;
            
            let firstCell = prodVal.toUpperCase();

            // CORTAFUEGOS
            if(firstCell.includes("PLANETA AZUL") && firstCell.includes("VENTAS")) return;
            if(firstCell.includes("PRECIO UNITARIO") && firstCell.length < 20) return;
            
            if(firstCell === 'TOTAL') {
                totalCount++;
                if (totalCount === 1) currentType = "Volumen";
                else if (totalCount === 2) currentType = "Monto (MM DOP)";
                else if (totalCount === 3) currentType = "Precio Unitario";
            }
            
            prodVal = prodVal.replace(/\s*\(\s*ZUMOS\s*\)\s*/i, ' ').replace(/\s{2,}/g, ' ').trim();
            firstCell = prodVal.toUpperCase();
            
            if(firstCell.includes("PRECIO UNITARIO")) return;
            
            let isParent = false;
            let parentId = null;
            let objId = firstCell.replace(/[^a-zA-Z0-9]/g, '_');
            
            if (firstCell === 'BEBIDAS' || firstCell === 'AGUA PLANETA AZUL' || firstCell === 'MAQUILAS') {
                isParent = true;
                currentParentId = objId;
                parentId = null;
            } else if (firstCell === 'TOTAL' || firstCell.startsWith('TOTAL SIN BON') || firstCell === 'TOTAL AÑO') {
                currentParentId = null;
            } else {
                parentId = currentParentId;
            }

            let obj = {
                Producto: prodVal,
                Tipo: currentType,
                hasChildren: isParent,
                parentId: parentId,
                id: objId,
                values: {},
                pptoValues: {}
            };
            
            if(tipoKey && r[tipoKey] && String(r[tipoKey]).trim() !== '' && String(r[tipoKey]).trim() !== '0') {
                obj.Tipo = String(r[tipoKey]);
                currentType = obj.Tipo;
            }
            
            keys.forEach(k => {
                if(k === prodKey || k === tipoKey || k.includes('__GHOST_')) return;
                
                let textK = String(k).trim().toLowerCase();
                let val = parseFloat(String(r[k]).replace(/,/g, ''));
                if (isNaN(val)) val = 0; // Prevenir arrastre de NaN
                
                let isDate = false;
                let dateStr = "";
                
                if (k.includes('-01 00:00:00') || (k.startsWith('202') && k.length >= 7)) {
                    isDate = true;
                    dateStr = k.slice(0, 7);
                } 
                else if (!isNaN(k) && Number(k) > 40000 && Number(k) < 60000) {
                    isDate = true;
                    let dObj = new Date(Math.round((Number(k) - 25569) * 86400 * 1000));
                    let y = dObj.getUTCFullYear();
                    let m = String(dObj.getUTCMonth() + 1).padStart(2, '0');
                    dateStr = `${y}-${m}`;
                } else {
                    const monthsMap = {
                        'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
                        'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12',
                        'jan': '01', 'apr': '04', 'aug': '08', 'dec': '12'
                    };
                    const match = textK.match(/^([a-z]{3})[-\/ ]?(\d{2,4})$/);
                    if(match && monthsMap[match[1].toLowerCase()]) {
                        let y = match[2];
                        if(y.length === 2) y = "20" + y;
                        dateStr = `${y}-${monthsMap[match[1].toLowerCase()]}`;
                        isDate = true;
                    } else if (textK.match(/^\d{4}-\d{2}$/) || textK.match(/^\d{4}-\d{2}-\d{2}/)) {
                        dateStr = textK.substring(0, 7);
                        isDate = true;
                    } else if (textK.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/)) {
                        const mMatch = textK.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                        let m = mMatch[1].padStart(2, '0');
                        let y = mMatch[3];
                        if (y.length === 2) y = "20" + y;
                        dateStr = `${y}-${m}`;
                        isDate = true;
                    }
                }

                if (isDate) {
                    obj.values[dateStr] = val;
                } else if (textK.includes('fy') || textK.includes('real 2024') || k === 'FY2024') {
                    obj['FY2024'] = val;
                } else if (textK.includes('po26') || textK.includes('ppto') || textK.includes('presupuesto')) {
                    obj['PO26'] = val;
                } else if (textK.includes('po25')) {
                    obj['PO25'] = val;
                }
            });
            parsedRows.push(obj);
        });
        
        return parsedRows;
    }

    function parseConsejoSheet(rows) {
        try {
            if (rows.length === 0) return [];

            let startIndex = -1;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                if (row && row.length > 0) {
                    const firstCell = String(row[0] || '').toUpperCase().trim();
                    if (firstCell === 'TOTAL' || firstCell === 'BEBIDAS') {
                        startIndex = i - 1;
                        break;
                    }
                }
            }

            if (startIndex === -1) {
                startIndex = rows.findIndex(row => row && row.some(cell => {
                    const text = String(cell).toUpperCase().trim();
                    return text === 'TOTAL' || text === 'BEBIDAS';
                })) - 1;
            }

            if (startIndex < 0) startIndex = 0;

            const headers = rows[startIndex] || [];
            
            let prodColIdx = -1;
            let tipoColIdx = -1;
            headers.forEach((h, idx) => {
                if(!h) return;
                const text = String(h).toLowerCase().trim();
                let isProdMatch = text === 'producto' || text === 'descripción' || text === 'descripcion' || text === 'artículo' || text === 'item' || text === 'concepto' || text === 'conceptos';
                if(isProdMatch) headers[idx] = 'Producto';
                if(text === 'tipo') headers[idx] = 'Tipo';
            });
            
            const totalRow = rows.find(row => row && row.some(cell => String(cell).toUpperCase().trim().startsWith('TOTAL')));
            if(totalRow && !headers.includes('Producto')) {
                const tIdx = totalRow.findIndex(cell => String(cell).toUpperCase().trim().startsWith('TOTAL'));
                if(tIdx !== -1) {
                    headers[tIdx] = 'Producto';
                }
            }
            if(totalRow && !headers.includes('Tipo')) {
                const pIdx = headers.indexOf('Producto');
                if(pIdx !== -1 && pIdx + 1 < headers.length) {
                    headers[pIdx + 1] = 'Tipo';
                }
            }
            
            if(!headers.includes('Producto')) {
                if(String(headers[0]||'').toLowerCase().replace(/\./g, '').trim() === 'no') {
                    headers[1] = 'Producto';
                } else {
                    headers[0] = 'Producto';
                    if(!headers[0]) {
                        const firstValIdx = headers.findIndex(h => h && String(h).trim() !== '');
                        if(firstValIdx !== -1) headers[firstValIdx] = 'Producto';
                        else headers[0] = 'Producto';
                    }
                }
            }
            
            let maxRowLength = Math.max(...rows.map(r => r ? r.length : 0));
            for(let i=0; i<maxRowLength; i++) {
                if(!headers[i]) headers[i] = `Col_${i}`;
            }

            let pptoStartIdx = 9999;
            const scanStart = Math.max(0, startIndex - 2);
            for (let i = scanStart; i <= startIndex; i++) {
                const r = rows[i];
                if (!r) continue;
                r.forEach((cell, idx) => {
                    if (!cell) return;
                    const s = String(cell).toUpperCase().trim();
                    if (s === 'PPTO' || s === 'PRESUPUESTO' || s.includes('PPTO 2026') || s.includes('PRESUP')) {
                        if (idx < pptoStartIdx) pptoStartIdx = idx;
                    }
                });
            }
            if (pptoStartIdx === 9999) {
                headers.forEach((cell, idx) => {
                    if (cell) {
                        const s = String(cell).toUpperCase().trim();
                        if (s.includes('PPTO') || s.includes('PRESUPUESTO')) {
                            if (idx < pptoStartIdx) pptoStartIdx = idx;
                        }
                    }
                });
            }
            if (pptoStartIdx === 9999) pptoStartIdx = 39;

            const data = [];
            for(let i = startIndex + 1; i < rows.length; i++) {
                const row = rows[i];
                if(!row) continue;
                
                let isRowEmpty = true;
                for(let c=0; c<row.length; c++) { if(row[c] !== undefined && row[c] !== null && String(row[c]).trim() !== '') { isRowEmpty = false; break; } }
                if(isRowEmpty) continue;
                
                const firstFewStrs = row.slice(0, 5).map(c => String(c||'').toUpperCase());
                if(firstFewStrs.some(s => s.includes("PRECIO UNITARIO"))) continue;
                
                let obj = {};
                headers.forEach((h, idx) => {
                    if(h !== undefined && h !== null && row[idx] !== undefined) {
                        let headerStr = String(h).trim();
                        obj[headerStr] = row[idx];
                    }
                });
                data.push(obj);
            }
            
            let parsedRows = [];
            let currentType = "Volumen"; 
            let tableCount = 0;
            let currentParentId = null;
            
            data.forEach(d => {
                const dProdStr = String(d.Producto || '').toUpperCase().trim();
                if (dProdStr === 'TOTAL') {
                    tableCount++;
                    if (tableCount === 1) currentType = "Volumen";
                    else if (tableCount === 2) currentType = "Monto (MM DOP)";
                    else if (tableCount === 3) currentType = "Precio Unitario";
                }
                
                let cellValues = Object.values(d).map(v => String(v).toUpperCase().trim());
                if (cellValues.some(v => v.includes('VOLUMEN') && !v.includes('PRECIO'))) {
                    currentType = 'Volumen';
                } else if (cellValues.some(v => v.includes('VENTAS') || v === 'MONTO' || v.includes('NETAS DOP'))) {
                    currentType = 'Ventas Netas DOP';
                }

                if(d.Tipo && String(d.Tipo).trim() !== '') {
                    currentType = String(d.Tipo).trim();
                }
                
                if (dProdStr.includes('VOLUMEN') && !dProdStr.includes('PRECIO')) currentType = 'Volumen';
                if (dProdStr.includes('VENTAS') || dProdStr === 'MONTO' || dProdStr.includes('NETAS DOP')) currentType = 'Ventas Netas DOP';

                if (!d.Producto || 
                    dProdStr === 'VOLUMEN' || 
                    dProdStr === 'VENTAS NETAS DOP' || 
                    dProdStr === 'VOLUMEN UNIDADES' ||
                    dProdStr === 'MONTO' ||
                    dProdStr.includes('PRECIO UNITARIO') ||
                    dProdStr === 'TABLAS CONSEJO') {
                    return;
                }
                
                let finalType = "Monto (MM DOP)";
                let cTypeUpper = currentType.toUpperCase();
                if(cTypeUpper.includes("VOLUMEN")) {
                    finalType = "Volumen";
                }
                
                let rawProducto = String(d.Producto).replace(/\s*\(\s*ZUMOS\s*\)\s*/i, ' ').replace(/\s{2,}/g, ' ').trim();
                let firstCell = rawProducto.toUpperCase();
                
                let isParent = false;
                let parentId = null;
                let objId = firstCell.replace(/[^a-zA-Z0-9]/g, '_');
                
                if (firstCell === 'BEBIDAS' || firstCell === 'AGUA PLANETA AZUL' || firstCell === 'MAQUILAS') {
                    isParent = true;
                    currentParentId = objId;
                    parentId = null;
                } else if (firstCell === 'TOTAL' || firstCell.startsWith('TOTAL SIN BON') || firstCell === 'TOTAL AÑO') {
                    currentParentId = null;
                } else {
                    parentId = currentParentId;
                }

                let p = {
                    Producto: rawProducto,
                    Tipo: finalType,
                    hasChildren: isParent,
                    parentId: parentId,
                    id: objId,
                    values: {},
                    pptoValues: {}
                };
                
                Object.keys(d).forEach(k => {
                    let isDate = false;
                    let dateStr = k;
                    let textK = String(k).trim().toLowerCase();
                    
                    if (k.includes('-01 00:00:00') || (k.startsWith('202') && k.length >= 7)) {
                        isDate = true;
                        dateStr = k.slice(0, 7);
                    } else if (!isNaN(k) && Number(k) > 40000 && Number(k) < 50000) {
                        isDate = true;
                        let dObj = new Date((Number(k) - 25569) * 86400 * 1000);
                        dateStr = dObj.toISOString().slice(0, 7);
                    } else {
                        const monthsMap = {
                            'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
                            'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12'
                        };
                        const regex = /([a-z]{3})[-/ ]?(\d{2,4})/;
                        const match = textK.match(regex);
                        if(match && monthsMap[match[1].toLowerCase()]) {
                            let y = match[2];
                            if(y.length === 2) y = "20" + y;
                            dateStr = `${y}-${monthsMap[match[1].toLowerCase()]}`;
                            isDate = true;
                        }
                    }
                    
                    if(isDate) {
                        let val = parseFloat(String(d[k]).replace(/,/g, '')) || 0;
                        if(finalType === "Monto (MM DOP)") val = val / 1000000;
                        
                        let colIndex = headers.indexOf(k);
                        const isPptoCol = colIndex >= pptoStartIdx || textK.includes('ppto') || textK.includes('presupuesto');
                        if (isPptoCol) {
                            p.pptoValues[dateStr] = val;
                        } else {
                            p.values[dateStr] = val;
                        }
                    } else if(textK.includes('fy') || textK.includes('real 2024') || k === 'FY2024' || k === 'PO26' || k === 'PO25') {
                        let yKey = 'FY2024';
                        if(textK.includes('po26')) yKey = 'PO26';
                        else if(textK.includes('po25')) yKey = 'PO25';
                        
                        let val = parseFloat(String(d[k]).replace(/,/g, '')) || 0;
                        if(finalType === "Monto (MM DOP)") val = val / 1000000;
                        p[yKey] = val;
                    }
                });
                
                // Fallback for 2026 missing actuals
                let lastNonZeroDate = "";
                const sortedDates = Object.keys(p.values).sort();
                for (let d of sortedDates) {
                    if (p.values[d] !== 0) {
                        lastNonZeroDate = d;
                    }
                }
                
                Object.keys(p.pptoValues).sort().forEach(d => {
                    if (d.startsWith('2026-') && d > lastNonZeroDate) {
                        if (p.values[d] === 0 || p.values[d] === undefined) {
                            p.values[d] = p.pptoValues[d] || 0;
                        }
                    }
                });

                parsedRows.push(p);
            });
            
            const productos = [...new Set(parsedRows.map(r => r.Producto))];
            productos.forEach(prod => {
                const montoRow = parsedRows.find(r => r.Producto === prod && r.Tipo === "Monto (MM DOP)");
                const volRow = parsedRows.find(r => r.Producto === prod && r.Tipo === "Volumen");
                if(montoRow && volRow) {
                    let p = { 
                        Producto: prod, 
                        Tipo: "Precio Unitario", 
                        hasChildren: montoRow.hasChildren,
                        parentId: montoRow.parentId,
                        id: montoRow.id,
                        values: {}, 
                        pptoValues: {} 
                    };
                    Object.keys(montoRow.values || {}).forEach(k => {
                        let volVal = volRow.values[k] || 0;
                        let div = volVal ? ((montoRow.values[k] * 1000000) / volVal) : 0;
                        p.values[k] = div;
                    });
                    
                    Object.keys(montoRow.pptoValues || {}).forEach(k => {
                        let volVal = volRow.pptoValues[k] || 0;
                        let div = volVal ? ((montoRow.pptoValues[k] * 1000000) / volVal) : 0;
                        p.pptoValues[k] = div;
                    });
                    
                    ['FY2024', 'PO25', 'PO26'].forEach(y => {
                        if (volRow[y] && montoRow[y]) {
                            p[y] = ((montoRow[y] * 1000000) / volRow[y]);
                        } else {
                            p[y] = 0;
                        }
                    });
                    
                    parsedRows.push(p);
                }
            });
            
            return parsedRows;
        } catch(e) {
            console.error("Error parsing Ventas CEO", e);
            return [];
        }
    }
    
    window.processVentasCeoFile = async function(file) {
        return new Promise(async (resolve) => {
            try {
                const buffer = await file.arrayBuffer();
                const data = new Uint8Array(buffer);
                const workbook = XLSX.read(data, { type: 'array' });
                let bestSheetName = workbook.SheetNames[0];
                const consejoSheet = workbook.SheetNames.find(n => n.toLowerCase().includes('consejo'));
                if (consejoSheet) {
                    bestSheetName = consejoSheet;
                } else {
                    for (let name of workbook.SheetNames) {
                        const sheetTmp = workbook.Sheets[name];
                        const rowsTmp = XLSX.utils.sheet_to_json(sheetTmp, { header: 1 });
                        const hasProducto = rowsTmp.some(r => r && r.some(c => String(c).toLowerCase().trim() === 'producto' || String(c).toLowerCase().trim() === 'descripción'));
                        if (hasProducto) {
                            bestSheetName = name;
                            break;
                        }
                    }
                }
                window.processVentasCeoWorkbook(workbook);
                resolve();
            } catch(e) {
                console.warn("No es un Excel válido, intentando como texto (CSV)...", e);
                try {
                    const text = await file.text();
                    const lines = text.split(/\r?\n/).map(l => l.split(','));
                    window.processVentasCeoWorkbook(null, lines);
                } catch (textError) {
                    console.error("No se pudo leer como CSV", textError);
                }
                resolve();
            }
        });
    };
    
    window.processResumenComercialFile = async function(file) {
        return new Promise(async (resolve) => {
            try {
                const buffer = await file.arrayBuffer();
                const data = new Uint8Array(buffer);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Dynamically import to avoid top-level issues if any
                const engine = await import('./resumenComercialEngine.js');
                window.resumenComercialEngine = engine;
                
                const monthSelector = document.getElementById('monthSelector');
                const m = monthSelector ? parseInt(monthSelector.value) : 3;

                // Process data using the new motor financiero
                await engine.processComercialWorkbook(workbook);
                
                window.hasComercialAccess = true;
                if (typeof window.applyRoleBasedUI === 'function') {
                    window.applyRoleBasedUI(window.hasMasterAccess, window.hasVentasAccess, true);
                }

                window.renderResumenComercial();
                resolve();
            } catch(e) {
                console.error("Error procesando Excel de Resumen Comercial", e);
                resolve();
            }
        });
    };

    window.buildVentasCeoFromComercial = async function() {
        if (!window.resumenComercialEngine) return false;
        const rawData = window.resumenComercialEngine.getComercialRawData();
        if (!rawData) return false;
        
        let csvRows = [];
        try {
            const res = await fetch('./ventasCEO_summary.csv');
            const text = await res.text();
            csvRows = text.split('\n').filter(l => l.trim().length > 0);
        } catch(e) {
            console.warn("Could not fetch ventasCEO_summary.csv", e);
        }
        
        const csvMapper = {
            'Volumen': {},
            'Monto (MM DOP)': {},
            'Precio Unitario': {}
        };
        
        let currentTypeSection = 'Volumen';
        let sectionCount = 0;
        csvRows.forEach(row => {
            if(row.startsWith('---')) {
                sectionCount++;
                if(sectionCount === 1) currentTypeSection = 'Monto (MM DOP)';
                else if(sectionCount === 2) currentTypeSection = 'Precio Unitario';
                return;
            }
            const cols = row.split(',');
            if (cols[0] && cols[0] !== 'Producto') {
                const prod = cols[0].toUpperCase().trim();
                const real2024 = parseFloat(cols[1]) || 0;
                let po26 = parseFloat(cols[4]) || 0; 
                csvMapper[currentTypeSection][prod] = {
                     FY2024: real2024,
                     PO26: po26
                };
            }
        });

        const monthTables = {};
        for (let m = 1; m <= 12; m++) {
            monthTables[m] = window.resumenComercialEngine.buildComercialTable(rawData, m, false);
        }
        
        const CEO_MAPPINGS = [
          { Producto: "TOTAL", match: null, isParent: false },
          { Producto: "AGUA PLANETA AZUL", match: null, isParent: true, parentId: null },
          { Producto: "APA BOTELLON 18.9 LTS (x1)", match: ["BT5_Total"], parentId: "AGUA PLANETA AZUL" },
          { Producto: "APA BOTELLA 0.5 LTS ( x20)", match: ["APA_BOTELLA_0_5_LTS_x20"], parentId: "AGUA PLANETA AZUL" },
          { Producto: "APA BOTELLA 1.5 LTS (x12)", match: ["APA_BOTELLA_1_5_LTS_x12"], parentId: "AGUA PLANETA AZUL" },
          { Producto: "APA OTRAS", match: ["APA_OTROS"], parentId: "AGUA PLANETA AZUL" },
          { Producto: "MAQUILAS", match: null, isParent: true, parentId: null },
          { Producto: "MAQUILA AGUA OTROS", match: ["MAQUILA_OTROS"], parentId: "MAQUILAS" }, 
          { Producto: "MAQUILA AGUA 1.5 L TS (x12)", match: ["MAQUILA_AGUA_1_5_L_TS_x12"], parentId: "MAQUILAS" },
          { Producto: "MAQUILA AGUA 0.5 LTS (x20)", match: ["MAQUILA_AGUA_0_5_LTS_x20"], parentId: "MAQUILAS" },
          { Producto: "BEBIDAS", match: null, isParent: true, parentId: null },
          { Producto: "BON", match: ["BON_Total"], parentId: "BEBIDAS" }, 
          { Producto: "PA SABOR 0.5 LTS (x12)", match: ["PA_SABOR_0_5_LTS_x12"], parentId: "BEBIDAS" },
          { Producto: "PA H+ 0.68 LTS (X12)", match: ["HIDRACTIVE_PLUS"], parentId: "BEBIDAS" },
          { Producto: "TOTAL SIN BON", match: null, isParent: false }
        ];

        const TYPE_MAPPINGS = [
           { key: 'Volumen', tableKey: 'volumen', rawDivisor: 1000 },
           { key: 'Monto (MM DOP)', tableKey: 'ventas', rawDivisor: 1000000 }, 
           { key: 'Precio Unitario', tableKey: 'precio', rawDivisor: 1 }
        ];

        const finalData = [];
        
        TYPE_MAPPINGS.forEach(tm => {
            CEO_MAPPINGS.forEach(cMap => {
                let p = {
                    Producto: cMap.Producto,
                    Tipo: tm.key,
                    hasChildren: cMap.isParent,
                    parentId: cMap.parentId,
                    id: cMap.isParent ? cMap.Producto.replace(/[^a-zA-Z0-9]/g, '_') : cMap.Producto.replace(/[^a-zA-Z0-9]/g, '_'),
                    values: {},
                    pptoValues: {}
                };
                
                const upProd = cMap.Producto.toUpperCase().trim();
                let matchingProdName = upProd;
                if (csvMapper[tm.key] && !csvMapper[tm.key][upProd]) {
                     matchingProdName = Object.keys(csvMapper[tm.key]).find(k => k.includes('H+') || k.includes(upProd));
                }

                if (csvMapper[tm.key] && csvMapper[tm.key][matchingProdName || upProd]) {
                    p.FY2024 = csvMapper[tm.key][matchingProdName || upProd].FY2024;
                    p.PO26 = csvMapper[tm.key][matchingProdName || upProd].PO26;
                } else if (!cMap.isParent) {
                    p.FY2024 = 0; p.PO26 = 0;
                }
                
                if (cMap.match) {
                    for (let m = 1; m <= 12; m++) {
                        let mStr = m.toString().padStart(2, '0');
                        let sum25 = 0, sum26 = 0, sumPpto = 0;
                        cMap.match.forEach(nodeId => {
                             let row = monthTables[m].tableRows.find(r => r.node.id === nodeId);
                             if (row) {
                                 if (tm.tableKey === 'precio') {
                                     sum25 = row[tm.tableKey].a25;
                                     sum26 = row[tm.tableKey].a26;
                                     sumPpto = row[tm.tableKey].ppto;
                                 } else {
                                     sum25 += row[tm.tableKey].a25 || 0;
                                     sum26 += row[tm.tableKey].a26 || 0;
                                     sumPpto += row[tm.tableKey].ppto || 0;
                                 }
                             }
                        });
                        
                        p.values[`2025-${mStr}`] = sum25 ? sum25 / tm.rawDivisor : 0;
                        p.values[`2026-${mStr}`] = sum26 ? sum26 / tm.rawDivisor : 0;
                        p.pptoValues[`2026-${mStr}`] = sumPpto ? sumPpto / tm.rawDivisor : 0;
                    }
                }
                finalData.push(p);
            });
        });
        
        const getParentGroups = () => {
            const groups = {};
            finalData.forEach(d => {
                if (d.parentId) {
                    if(!groups[d.parentId]) groups[d.parentId] = [];
                    groups[d.parentId].push(d);
                }
            });
            return groups;
        };
        const groups = getParentGroups();
        
        TYPE_MAPPINGS.forEach(tm => {
            const typesRows = finalData.filter(d => d.Tipo === tm.key);
            
            Object.keys(groups).forEach(pId => {
                const parentRow = typesRows.find(r => r.Producto === pId); 
                if (parentRow) {
                    const children = typesRows.filter(r => r.parentId === pId && r.Producto !== "PA H+ 0.68 LTS (X12)");
                    if (tm.key !== 'Precio Unitario') {
                        parentRow.FY2024 = children.reduce((acc, c) => acc + (c.FY2024||0), 0);
                        parentRow.PO26 = children.reduce((acc, c) => acc + (c.PO26||0), 0);
                        ['2025', '2026'].forEach(y => {
                            for(let m=1;m<=12;m++) {
                                let mStr = `${y}-${m.toString().padStart(2, '0')}`;
                                parentRow.values[mStr] = children.reduce((acc, c) => acc + (c.values[mStr]||0), 0);
                            }
                        });
                        for(let m=1;m<=12;m++) {
                            let mStr = `2026-${m.toString().padStart(2, '0')}`;
                            parentRow.pptoValues[mStr] = children.reduce((acc, c) => acc + (c.pptoValues[mStr]||0), 0);
                        }
                    }
                }
            });
            
            const totalRow = typesRows.find(r => r.Producto === 'TOTAL');
            const totalSinBonRow = typesRows.find(r => r.Producto === 'TOTAL SIN BON');
            const parents = typesRows.filter(r => r.hasChildren && ['AGUA PLANETA AZUL', 'MAQUILAS', 'BEBIDAS'].includes(r.Producto.toUpperCase().trim()));
            const parentBon = typesRows.find(r => r.Producto === 'BON');
            
            if (totalRow && tm.key !== 'Precio Unitario') {
                totalRow.FY2024 = parents.reduce((acc, c) => acc + (c.FY2024||0), 0);
                totalRow.PO26 = parents.reduce((acc, c) => acc + (c.PO26||0), 0);
                ['2025', '2026'].forEach(y => {
                    for(let m=1;m<=12;m++) {
                        let mStr = `${y}-${m.toString().padStart(2, '0')}`;
                        totalRow.values[mStr] = parents.reduce((acc, c) => acc + (c.values[mStr]||0), 0);
                    }
                });
                for(let m=1;m<=12;m++) {
                    let mStr = `2026-${m.toString().padStart(2, '0')}`;
                    totalRow.pptoValues[mStr] = parents.reduce((acc, c) => acc + (c.pptoValues[mStr]||0), 0);
                }
            }
            if (totalSinBonRow && tm.key !== 'Precio Unitario') {
                totalSinBonRow.FY2024 = (totalRow.FY2024 || 0) - (parentBon?.FY2024 || 0);
                totalSinBonRow.PO26 = (totalRow.PO26 || 0) - (parentBon?.PO26 || 0);
                ['2025', '2026'].forEach(y => {
                    for(let m=1;m<=12;m++) {
                        let mStr = `${y}-${m.toString().padStart(2, '0')}`;
                        totalSinBonRow.values[mStr] = (totalRow.values[mStr]||0) - (parentBon?.values[mStr]||0);
                    }
                });
                for(let m=1;m<=12;m++) {
                    let mStr = `2026-${m.toString().padStart(2, '0')}`;
                    totalSinBonRow.pptoValues[mStr] = (totalRow.pptoValues[mStr]||0) - (parentBon?.pptoValues[mStr]||0);
                }
            }
        });
        
        const pxRows = finalData.filter(d => d.Tipo === 'Precio Unitario');
        pxRows.forEach(pxRow => {
            if (pxRow.hasChildren || pxRow.Producto === 'TOTAL' || pxRow.Producto === 'TOTAL SIN BON') {
                const volRow = finalData.find(r => r.Tipo === 'Volumen' && r.Producto === pxRow.Producto);
                const montoRow = finalData.find(r => r.Tipo === 'Monto (MM DOP)' && r.Producto === pxRow.Producto);
                if (volRow && montoRow) {
                    pxRow.FY2024 = volRow.FY2024 ? (montoRow.FY2024 * 1000000) / (volRow.FY2024 * 1000) : 0;
                    pxRow.PO26 = volRow.PO26 ? (montoRow.PO26 * 1000000) / (volRow.PO26 * 1000) : 0;
                    ['2025', '2026'].forEach(y => {
                        for(let m=1;m<=12;m++) {
                            let mStr = `${y}-${m.toString().padStart(2, '0')}`;
                            let v = volRow.values[mStr]||0;
                            let mo = montoRow.values[mStr]||0;
                            pxRow.values[mStr] = v ? (mo * 1000000) / (v * 1000) : 0;
                        }
                    });
                    for(let m=1;m<=12;m++) {
                        let mStr = `2026-${m.toString().padStart(2, '0')}`;
                        let v = volRow.pptoValues[mStr]||0;
                        let mo = montoRow.pptoValues[mStr]||0;
                        pxRow.pptoValues[mStr] = v ? (mo * 1000000) / (v * 1000) : 0;
                    }
                }
            }
        });

        finalData.forEach(p => {
             let lastNonZeroDate = "2026-00"; 
             Object.keys(p.values).sort().forEach(d => {
                 if (d.startsWith('2026-') && p.values[d] !== 0) {
                     lastNonZeroDate = d;
                 }
             });
             Object.keys(p.pptoValues).sort().forEach(d => {
                 if (d.startsWith('2026-') && d > lastNonZeroDate) {
                     if (p.values[d] === 0 || p.values[d] === undefined) {
                         p.values[d] = p.pptoValues[d] || 0;
                     }
                 }
             });
        });

        ceoData = finalData;
        
        // Also update IndexedDB with the generated data
        try {
            const db = await getFinanceDB();
            await new Promise((resolve, reject) => {
                const tx = db.transaction('finance_cache', 'readwrite');
                const store = tx.objectStore('finance_cache');
                store.put({ id: 'ceo_ventas', data: ceoData, timestamp: Date.now() });
                tx.oncomplete = () => resolve();
                tx.onerror = () => reject(tx.error);
            });
        } catch(err) {
            console.warn("Could not cache dynamically generated ceoData", err);
        }

        window.hasVentasAccess = true;
        console.log("🔥 Ventas CEO dynamically built from Comercial");
        let viewVentasCeo = document.getElementById("view-ventas-ceo");
        if (viewVentasCeo && viewVentasCeo.classList.contains("active")) {
            window.renderVentasCEO();
        }
        
        // Also remove lock on the sidebar
        const svNode = document.querySelector('[data-view="view-ventas-ceo"]');
        if (svNode) {
            const lNode = svNode.querySelector('.lucide-lock');
            if (lNode) lNode.remove();
            svNode.style.pointerEvents = 'auto';
            svNode.style.opacity = '1';
        }
        
        return true;
    };

    window.comercialCurrentView = 'resumen';
    window.currentCostoProd = 'botellon';

    window.costoUnitarioVista = 'tendencia';
    window.setCostoVista = function(vista) {
        window.costoUnitarioVista = vista;
        document.querySelectorAll('.costo-vista-btn').forEach(b => {
             b.classList.remove('active');
             b.style.background = 'transparent';
             b.style.color = 'var(--text-secondary)';
             b.style.boxShadow = 'none';
        });
        const btn = document.getElementById('btn-costo-vista-' + vista);
        if (btn) {
             btn.classList.add('active');
             btn.style.background = 'white';
             btn.style.color = 'var(--primary)';
             btn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
        }
        if (typeof window.updateCostoUnitario === 'function') {
             window.updateCostoUnitario();
        }
    };

    window.updateCostoUnitario = function() {
        const selector = document.getElementById('monthSelector');
        let m = selector && !isNaN(parseInt(selector.value)) ? parseInt(selector.value) : 3;
        let m_idx = 3;
        if (selector && Array.isArray(globalFinancialData) && globalFinancialData.length > m && globalFinancialData[m]) {
             let item = globalFinancialData[m];
             const match = item.date ? item.date.match(/ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC/i) : null;
             const monthsArr = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
             if (match) {
                 m_idx = monthsArr.indexOf(match[0].toUpperCase());
             }
        }
        if (window.costoUnitarioEngine && window.costoUnitarioEngine.hasCostoUnitarioData()) {
            window.costoUnitarioEngine.renderCostoUnitario(m_idx, window.currentCostoProd, window.costoUnitarioVista);
        }
    };

    function updateComercialButtonsVisuals() {
        const resetBtn = (id) => {
            const btn = document.getElementById(id);
            if (btn) {
                btn.style.background = 'transparent';
                btn.style.color = 'var(--text-secondary)';
                btn.style.boxShadow = 'none';
            }
        };
        resetBtn('btn-comercial-resumen');
        resetBtn('btn-comercial-mom');
        resetBtn('btn-comercial-variacion');

        let activeId = 'btn-comercial-resumen';
        if (window.comercialCurrentView === 'mom') {
            activeId = 'btn-comercial-mom';
        } else if (window.comercialCurrentView === 'variacion') {
            activeId = 'btn-comercial-variacion';
        }

        const activeBtn = document.getElementById(activeId);
        if (activeBtn) {
            activeBtn.style.background = 'white';
            activeBtn.style.color = 'var(--primary)';
            activeBtn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
        }
    }

    window.processPgHorizontalFile = async function(file) {
        return new Promise(async (resolve) => {
            try {
                const buffer = await file.arrayBuffer();
                const data = new Uint8Array(buffer);
                const workbook = XLSX.read(data, { type: 'array' });
                if (!window.resumenComercialEngine) {
                     const engine = await import('./resumenComercialEngine.js');
                     window.resumenComercialEngine = engine;
                }
                const engine = window.resumenComercialEngine;
                await engine.processPgHorizontalWorkbook(workbook);

                window.hasComercialAccess = true;
                if (typeof window.applyRoleBasedUI === 'function') {
                    window.applyRoleBasedUI(window.hasMasterAccess, window.hasVentasAccess, true);
                }

                window.renderPgHorizontal();
                resolve();
            } catch(e) {
                console.error("Error procesando Excel de P&G Horizontal", e);
                resolve();
            }
        });
    };

    
    window.processCxpWorkbook = async function(workbook) {
    // 1. Identify sheets
    const names = workbook.SheetNames;
    const historicoName = names.find(n => n.includes("Historico") && n.includes("CXP"));
    const balanzaName = names.find(n => n.includes("Balanza"));
    const analisisName = names.find(n => n.includes("Analisis") || n.includes("Análisis"));
    
    if (!historicoName || !balanzaName || !analisisName) {
        console.warn("Could not find one of the required sheets (Historico CXP, Balanza, Analisis).");
        return;
    }

    const { sheet_to_json } = XLSX.utils;
    const cleanVal = (v) => {
        if (v === undefined || v === null || v === '') return 0;
        if (typeof v === 'number') return v;
        let str = String(v).trim();
        const isNegative = str.startsWith("'(") || str.startsWith('(');
        str = str.replace(/[^0-9.-]/g, '');
        let n = parseFloat(str);
        if (isNaN(n)) return 0;
        return isNegative ? -Math.abs(n) : n;
    };
    
    // --- PASO 1 - IDENTIFICAR PERIODO ---
    const histRows = sheet_to_json(workbook.Sheets[historicoName], { header: 1 });
    if (histRows.length === 0) return;
    
    const histHeaders = histRows[0];
    let periodColIdx = -1;
    let nombreSocioColIdx = -1;
    let totalCxpColIdx = -1;
    let colNoVencido = -1, col0_30 = -1, col31_60 = -1, col61_90 = -1, col91_120 = -1, col121_150 = -1, col151_180 = -1, col180Mas = -1;

    for (let i = 0; i < histHeaders.length; i++) {
        let val = String(histHeaders[i]).toLowerCase().trim();
        if (val === 'period') periodColIdx = i;
        if (val === 'nombresocio') nombreSocioColIdx = i;
        if (val === 'total saldo cxp') totalCxpColIdx = i;
        if (val === 'saldo no vencido') colNoVencido = i;
        if (val === '0 a 30') col0_30 = i;
        if (val === '31 a 60') col31_60 = i;
        if (val === '61 a 90') col61_90 = i;
        if (val === '91 a 120') col91_120 = i;
        if (val === '121 a 150') col121_150 = i;
        if (val === '151 a 180') col151_180 = i;
        if (val === '> 180') col180Mas = i;
    }

    let maxDate = new Date(0);
    let maxPeriod = "";
    
    for (let i = 1; i < histRows.length; i++) {
        let p = histRows[i][periodColIdx];
        if (p) {
            let parts = String(p).split('/');
            if (parts.length === 2) {
                let m = parseInt(parts[0], 10);
                let y = parseInt(parts[1], 10);
                let d = new Date(y, m - 1, 1);
                if (d > maxDate) {
                    maxDate = d;
                    maxPeriod = String(p);
                }
            }
        }
    }
    
    if (maxPeriod === "") {
        console.warn("No valid periods found in Historico");
        return;
    }
    
    let dates = [];
    let curD = new Date(maxDate.getFullYear(), maxDate.getMonth(), 1);
    for (let i = 0; i < 24; i++) {
        dates.unshift(new Date(curD.getFullYear(), curD.getMonth(), 1));
        curD.setMonth(curD.getMonth() - 1);
    }
    
    const periods = dates.map(d => d.getMonth() + 1 + "/" + d.getFullYear());
    const shortMonthsList = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    const labels = dates.map(d => shortMonthsList[d.getMonth()] + " " + d.getFullYear());

    // --- PASO 2 - HOJA BALANZA ---
    const balRows = sheet_to_json(workbook.Sheets[balanzaName], { header: 1 });
    const balYears = balRows[0];
    const balMonthsName = balRows[1];
    
    let balIndices = [];
    for (let i = 0; i < periods.length; i++) {
        let y = dates[i].getFullYear();
        let m = dates[i].getMonth() + 1;
        let foundIdx = -1;
        for (let j = 2; j < balYears.length; j++) {
            if (balYears[j] == y && balMonthsName[j] == m) {
                foundIdx = j;
                break;
            }
        }
        balIndices.push(foundIdx);
    }
    
    const rowCXP = balRows[12];
    const rowOtrasCXP = balRows[21];
    const rowBal = balRows[24];
    
    let arrBalGen = [], arrCXP = [], arrProveedoresProv = [];
    for (let i = 0; i < periods.length; i++) {
        let idx = balIndices[i];
        if (idx !== -1) {
            arrBalGen.push(-cleanVal(rowBal[idx]) / 1000000);
            arrCXP.push(-cleanVal(rowCXP[idx]) / 1000000);
            arrProveedoresProv.push(-cleanVal(rowOtrasCXP[idx]) / 1000000);
        } else {
            arrBalGen.push(0); arrCXP.push(0); arrProveedoresProv.push(0);
        }
    }

    // --- PASO 3 - AGING ---
    let arrCorriente = Array(periods.length).fill(0);
    let arr0_30 = Array(periods.length).fill(0);
    let arr31_60 = Array(periods.length).fill(0);
    let arr61_90 = Array(periods.length).fill(0);
    let arr91_120 = Array(periods.length).fill(0);
    let arr121_150 = Array(periods.length).fill(0);
    let arr151_180 = Array(periods.length).fill(0);
    let arr180mas = Array(periods.length).fill(0);

    for (let i = 1; i < histRows.length; i++) {
        let p = String(histRows[i][periodColIdx]);
        let pIdx = periods.indexOf(p);
        if (pIdx !== -1) {
            arrCorriente[pIdx] += -cleanVal(histRows[i][colNoVencido]);
            arr0_30[pIdx] += -cleanVal(histRows[i][col0_30]);
            arr31_60[pIdx] += -cleanVal(histRows[i][col31_60]);
            arr61_90[pIdx] += -cleanVal(histRows[i][col61_90]);
            arr91_120[pIdx] += -cleanVal(histRows[i][col91_120]);
            arr121_150[pIdx] += -cleanVal(histRows[i][col121_150]);
            arr151_180[pIdx] += -cleanVal(histRows[i][col151_180]);
            arr180mas[pIdx] += -cleanVal(histRows[i][col180Mas]);
        }
    }
    
    const toMM = (arr) => arr.map(v => v / 1000000);
    arrCorriente = toMM(arrCorriente);
    arr0_30 = toMM(arr0_30);
    arr31_60 = toMM(arr31_60);
    arr61_90 = toMM(arr61_90);
    arr91_120 = toMM(arr91_120);
    arr121_150 = toMM(arr121_150);
    arr151_180 = toMM(arr151_180);
    arr180mas = toMM(arr180mas);
    
    // --- PASO 4 - TOP 14 PROVEEDORES ---
    let lastMonthProvTotals = {};
    for (let i = 1; i < histRows.length; i++) {
        let p = String(histRows[i][periodColIdx]);
        if (p === maxPeriod) {
            let prov = String(histRows[i][nombreSocioColIdx] || '');
            let v = -cleanVal(histRows[i][totalCxpColIdx]) / 1000000;
            if (!lastMonthProvTotals[prov]) lastMonthProvTotals[prov] = 0;
            lastMonthProvTotals[prov] += v;
        }
    }
    
    let provList = Object.keys(lastMonthProvTotals).map(k => ({name: k, val: lastMonthProvTotals[k]}));
    provList.sort((a,b) => Math.abs(b.val) - Math.abs(a.val));
    let top14Names = provList.slice(0, 14).map(p => p.name);
    
    let top14Saldos = {};
    for (let n of top14Names) {
        top14Saldos[n] = Array(periods.length).fill(0);
    }
    
    for (let i = 1; i < histRows.length; i++) {
        let p = String(histRows[i][periodColIdx]);
        let pIdx = periods.indexOf(p);
        if (pIdx !== -1) {
            let prov = String(histRows[i][nombreSocioColIdx] || '');
            if (top14Saldos[prov]) {
                top14Saldos[prov][pIdx] += -cleanVal(histRows[i][totalCxpColIdx]) / 1000000;
            }
        }
    }
    
    let arrOtros = Array(periods.length).fill(0);
    let arrTotal = Array(periods.length).fill(0);
    for (let i = 0; i < periods.length; i++) {
        let sumTop14 = 0;
        for (let n of top14Names) {
            sumTop14 += top14Saldos[n][i];
        }
        arrOtros[i] = arrBalGen[i] - sumTop14;
        arrTotal[i] = sumTop14 + arrOtros[i];
    }

    // --- PASO 5 - COSTOS YTD ---
    const anaRows = sheet_to_json(workbook.Sheets[analisisName], { header: 1 });
    const anaDates = anaRows[2]; // Fila 3, índice 2
    
    let anaCostos = null;
    let foundCostosRowIdx = -1;
    for (let r = 0; r < anaRows.length; r++) {
        if (anaRows[r] && anaRows[r][0]) {
            let rowStr = String(anaRows[r][0]).toLowerCase().trim();
            if (rowStr.includes("total costos") || rowStr.includes("costos ytd") || rowStr.includes("opex + capex")) {
                anaCostos = anaRows[r];
                foundCostosRowIdx = r;
                break;
            }
        }
    }
    
    // If we didn't find it dynamically by string matching common names, try row 38 as fallback:
    if (!anaCostos && anaRows[38]) {
        anaCostos = anaRows[38];
    }
    
    let arrCostosYTD = Array(periods.length).fill(0);
    
    const EXCEL_EPOCH = new Date(1899, 11, 30); // in excel 1= 1900-01-01 but there's leaps
    const getMonthAndYearFromExcel = (cell) => {
        if (!cell) return null;
        if (typeof cell === 'number') {
            let d = new Date(EXCEL_EPOCH.getTime() + cell * 86400000);
            return { m: d.getMonth() + 1, y: d.getFullYear() };
        }
        if (cell instanceof Date) {
            return { m: cell.getMonth() + 1, y: cell.getFullYear() };
        }
        let tk = String(cell).substring(0,10);
        let d = new Date(tk + 'T12:00:00Z');
        if (!isNaN(d.getTime())) return { m: d.getMonth() + 1, y: d.getFullYear() };
        
        let d2 = new Date(cell);
        if (!isNaN(d2.getTime())) return { m: d2.getMonth() + 1, y: d2.getFullYear() };
        return null;
    };
    
    let anaIndices = [];
    for (let i = 0; i < periods.length; i++) {
        let y = dates[i].getFullYear();
        let m = dates[i].getMonth() + 1;
        let foundIdx = -1;
        
        if (anaDates) {
            for (let j = 1; j < anaDates.length; j++) {
                let dt = getMonthAndYearFromExcel(anaDates[j]);
                if (dt && dt.m === m && dt.y === y) {
                    foundIdx = j; break;
                }
            }
        }
        anaIndices.push(foundIdx);
    }
    
    for (let i = 0; i < periods.length; i++) {
        let idx = anaIndices[i];
        if (idx !== -1 && anaCostos) {
            arrCostosYTD[i] = cleanVal(anaCostos[idx]); // Ya en MM
        }
    }
    
    // --- PASO 6 - DPO ---
    let arrDPO = Array(periods.length).fill(0);
    for (let i = 0; i < periods.length; i++) {
        let cytd = arrCostosYTD[i];
        if (cytd > 0) {
            arrDPO[i] = Math.round(arrBalGen[i] / (cytd / 30));
        }
    }

    window.cxpStandaloneData = {
        labels,
        periods,
        BalanceGeneral: arrBalGen,
        CXP: arrCXP,
        Provisionales: arrProveedoresProv,
        Corriente: arrCorriente,
        Aging: {
            "0_30": arr0_30,
            "31_60": arr31_60,
            "61_90": arr61_90,
            "91_120": arr91_120,
            "121_150": arr121_150,
            "151_180": arr151_180,
            "180Mas": arr180mas
        },
        Top14Names: top14Names,
        Top14Saldos: top14Saldos,
        OtrosProveedores: arrOtros,
        Total: arrTotal,
        CostosYTD: arrCostosYTD,
        DPO: arrDPO
    };
    
    try {
        const db = await getFinanceDB();
        const tx = db.transaction('finance_cache', 'readwrite');
        tx.objectStore('finance_cache').put({ data: window.cxpStandaloneData, timestamp: Date.now() }, 'CXP_STANDALONE_KEY');
    } catch(e) {
        console.error("Error saving to standalone indexeddb cxp", e);
    }
    
    if (window.currentActiveView === 'view-cxp') { 
        if (typeof window.renderCxpView === 'function') {
            window.renderCxpView(window.cxpStandaloneData);
        }
    }
}

window.processCxpFile = async function(file) {
        return new Promise(async (resolve) => {
            try {
                const buffer = await file.arrayBuffer();
                const data = new Uint8Array(buffer);
                const workbook = XLSX.read(data, { type: 'array' });
                await window.processCxpWorkbook(workbook);
                
                window.hasCxpAccess = true;
                if (typeof window.applyRoleBasedUI === 'function') {
                    window.applyRoleBasedUI(window.hasMasterAccess, window.hasVentasAccess, window.hasComercialAccess);
                }
                
                if (window.currentActiveView === 'view-cxp-detail') {
                    if (typeof renderCxpView === 'function') {
                        const idx = monthSelector ? parseInt(monthSelector.value, 10) : (globalFinancialData ? globalFinancialData.length - 1 : -1);
                        renderCxpView(globalFinancialData, idx);
                    }
                }
                resolve(true);
            } catch (e) {
                console.error("Error processing manual CxP file:", e);
                resolve(false);
            }
        });
    };

    window.processCostoUnitarioFile = async function(file) {
        return new Promise(async (resolve) => {
            try {
                const buffer = await file.arrayBuffer();
                const engine = await import('./costoUnitarioEngine.js');
                window.costoUnitarioEngine = engine;
                await engine.processManualFile(buffer);
                
                if (typeof window.applyRoleBasedUI === 'function') {
                    window.applyRoleBasedUI(window.hasMasterAccess, window.hasVentasAccess, window.hasComercialAccess);
                }

                if (typeof window.updateCostoUnitario === 'function') {
                    window.updateCostoUnitario();
                }
                resolve(true);
            } catch(e) {
                console.error("Error processing manual Costo Unitario file:", e);
                resolve(false);
            }
        });
    };

    // Attach click events on load/execution
    setTimeout(() => {
        document.getElementById('btn-comercial-resumen')?.addEventListener('click', () => {
            window.comercialCurrentView = 'resumen';
            updateComercialButtonsVisuals();
            window.renderResumenComercial();
        });
        document.getElementById('btn-comercial-mom')?.addEventListener('click', () => {
            window.comercialCurrentView = 'mom';
            updateComercialButtonsVisuals();
            window.renderResumenComercial();
        });
        document.getElementById('btn-comercial-variacion')?.addEventListener('click', () => {
            window.comercialCurrentView = 'variacion';
            updateComercialButtonsVisuals();
            window.renderResumenComercial();
        });

        document.getElementById('btn-costo-botellon')?.addEventListener('click', () => {
            window.currentCostoProd = 'botellon';
            document.getElementById('btn-costo-botellon').className = "costo-prod-btn active";
            document.getElementById('btn-costo-botellon').style.background = "white";
            document.getElementById('btn-costo-botellon').style.color = "var(--primary)";
            document.getElementById('btn-costo-botellon').style.boxShadow = "0 1px 2px rgba(0,0,0,0.05)";
            
            document.getElementById('btn-costo-botella').className = "costo-prod-btn";
            document.getElementById('btn-costo-botella').style.background = "transparent";
            document.getElementById('btn-costo-botella').style.color = "var(--text-secondary)";
            document.getElementById('btn-costo-botella').style.boxShadow = "none";
            
            if (typeof window.updateCostoUnitario === 'function') {
                window.updateCostoUnitario();
            }
        });

        document.getElementById('btn-costo-botella')?.addEventListener('click', () => {
            window.currentCostoProd = 'botella';
            document.getElementById('btn-costo-botella').className = "costo-prod-btn active";
            document.getElementById('btn-costo-botella').style.background = "white";
            document.getElementById('btn-costo-botella').style.color = "var(--primary)";
            document.getElementById('btn-costo-botella').style.boxShadow = "0 1px 2px rgba(0,0,0,0.05)";
            
            document.getElementById('btn-costo-botellon').className = "costo-prod-btn";
            document.getElementById('btn-costo-botellon').style.background = "transparent";
            document.getElementById('btn-costo-botellon').style.color = "var(--text-secondary)";
            document.getElementById('btn-costo-botellon').style.boxShadow = "none";

            if (typeof window.updateCostoUnitario === 'function') {
                window.updateCostoUnitario();
            }
        });

        document.getElementById('btn-cashflow-detalle')?.addEventListener('click', () => {
            const btnDetalle = document.getElementById('btn-cashflow-detalle');
            const btnResumen = document.getElementById('btn-cashflow-resumen');
            if (btnDetalle) { btnDetalle.style.background = 'white'; btnDetalle.style.color = 'var(--primary)'; btnDetalle.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)'; }
            if (btnResumen) { btnResumen.style.background = 'transparent'; btnResumen.style.color = 'var(--text-secondary)'; btnResumen.style.boxShadow = 'none'; }
            
            const detailContainer = document.getElementById('cashflow-detalle-container');
            const resumenContainer = document.getElementById('cashflow-resumen-container');
            if(detailContainer) detailContainer.style.display = 'block';
            if(resumenContainer) resumenContainer.style.display = 'none';
        });

        document.getElementById('btn-cashflow-resumen')?.addEventListener('click', () => {
            const btnDetalle = document.getElementById('btn-cashflow-detalle');
            const btnResumen = document.getElementById('btn-cashflow-resumen');
            if (btnResumen) { btnResumen.style.background = 'white'; btnResumen.style.color = 'var(--primary)'; btnResumen.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)'; }
            if (btnDetalle) { btnDetalle.style.background = 'transparent'; btnDetalle.style.color = 'var(--text-secondary)'; btnDetalle.style.boxShadow = 'none'; }
            
            const detailContainer = document.getElementById('cashflow-detalle-container');
            const resumenContainer = document.getElementById('cashflow-resumen-container');
            if(detailContainer) detailContainer.style.display = 'none';
            if(resumenContainer) resumenContainer.style.display = 'block';
        });

        document.getElementById('btn-balance-detalle')?.addEventListener('click', () => {
            const btnDetalle = document.getElementById('btn-balance-detalle');
            const btnResumen = document.getElementById('btn-balance-resumen');
            if (btnDetalle) { btnDetalle.style.background = 'white'; btnDetalle.style.color = 'var(--primary)'; btnDetalle.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)'; }
            if (btnResumen) { btnResumen.style.background = 'transparent'; btnResumen.style.color = 'var(--text-secondary)'; btnResumen.style.boxShadow = 'none'; }
            
            const detailContainer = document.getElementById('balance-detalle-container');
            const resumenContainer = document.getElementById('balance-resumen-container');
            if(detailContainer) detailContainer.style.display = 'block';
            if(resumenContainer) resumenContainer.style.display = 'none';
        });

        document.getElementById('btn-balance-resumen')?.addEventListener('click', () => {
            const btnDetalle = document.getElementById('btn-balance-detalle');
            const btnResumen = document.getElementById('btn-balance-resumen');
            if (btnResumen) { btnResumen.style.background = 'white'; btnResumen.style.color = 'var(--primary)'; btnResumen.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)'; }
            if (btnDetalle) { btnDetalle.style.background = 'transparent'; btnDetalle.style.color = 'var(--text-secondary)'; btnDetalle.style.boxShadow = 'none'; }
            
            const detailContainer = document.getElementById('balance-detalle-container');
            const resumenContainer = document.getElementById('balance-resumen-container');
            if(detailContainer) detailContainer.style.display = 'none';
            if(resumenContainer) resumenContainer.style.display = 'block';
        });

        document.getElementById('btn-cxp-detalle')?.addEventListener('click', () => {
            const btnDetalle = document.getElementById('btn-cxp-detalle');
            const btnResumen = document.getElementById('btn-cxp-resumen');
            if (btnDetalle) { btnDetalle.style.background = 'white'; btnDetalle.style.color = 'var(--primary)'; btnDetalle.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)'; }
            if (btnResumen) { btnResumen.style.background = 'transparent'; btnResumen.style.color = 'var(--text-secondary)'; btnResumen.style.boxShadow = 'none'; }
            
            const detailContainer = document.getElementById('cxp-detalle-container');
            const resumenContainer = document.getElementById('cxp-resumen-container');
            if(detailContainer) detailContainer.style.display = 'block';
            if(resumenContainer) resumenContainer.style.display = 'none';
        });

        document.getElementById('btn-cxp-resumen')?.addEventListener('click', () => {
            const btnDetalle = document.getElementById('btn-cxp-detalle');
            const btnResumen = document.getElementById('btn-cxp-resumen');
            if (btnResumen) { btnResumen.style.background = 'white'; btnResumen.style.color = 'var(--primary)'; btnResumen.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)'; }
            if (btnDetalle) { btnDetalle.style.background = 'transparent'; btnDetalle.style.color = 'var(--text-secondary)'; btnDetalle.style.boxShadow = 'none'; }
            
            const detailContainer = document.getElementById('cxp-detalle-container');
            const resumenContainer = document.getElementById('cxp-resumen-container');
            if(detailContainer) detailContainer.style.display = 'none';
            if(resumenContainer) resumenContainer.style.display = 'block';
        });

        const btnToggleComercial = document.getElementById('btn-toggle-comercial-view');
        if (btnToggleComercial) {
            btnToggleComercial.addEventListener('click', () => {
                const table = document.getElementById('resumen-comercial-table');
                if (!table) return;
                table.classList.toggle('card-view-tbl');
                const isCard = table.classList.contains('card-view-tbl');
                const lbl = document.getElementById('text-toggle-comercial-view');
                if (lbl) lbl.textContent = isCard ? 'Table View' : 'Card View';
                const icon = btnToggleComercial.querySelector('i');
                if (icon) {
                    icon.setAttribute('data-lucide', isCard ? 'table' : 'layout-grid');
                    if (window.lucide) window.lucide.createIcons();
                }
            });
        }

        const btnTogglePg = document.getElementById('btn-toggle-pg-view');
        if (btnTogglePg) {
            btnTogglePg.addEventListener('click', () => {
                const table1 = document.getElementById('pg-horizontal-table');
                const table2 = document.getElementById('pg-horizontal-unitarios-table');
                if (table1) table1.classList.toggle('card-view-tbl');
                if (table2) table2.classList.toggle('card-view-tbl');
                const isCard = (table1 && table1.classList.contains('card-view-tbl')) || (table2 && table2.classList.contains('card-view-tbl'));
                const lbl = document.getElementById('text-toggle-pg-view');
                if (lbl) lbl.textContent = isCard ? 'Table View' : 'Card View';
                const icon = btnTogglePg.querySelector('i');
                if (icon) {
                    icon.setAttribute('data-lucide', isCard ? 'table' : 'layout-grid');
                    if (window.lucide) window.lucide.createIcons();
                }
            });
        }
    }, 500);

    function applyMobileDataLabels(tableId, theadId) {
        const thead = document.getElementById(theadId);
        const table = document.getElementById(tableId);
        if (!thead || !table) return;
        
        const tbody = table.querySelector('tbody');
        if (!tbody) return;

        const trs = thead.querySelectorAll('tr');
        if (!trs.length) return;
        
        let matrix = [];
        for(let i=0; i<trs.length; i++) matrix.push([]);
        
        trs.forEach((tr, rowIndex) => {
            const cells = tr.querySelectorAll('th, td');
            let colIndex = 0;
            
            cells.forEach(cell => {
                while(matrix[rowIndex][colIndex] !== undefined) colIndex++;
                
                const rowSpan = parseInt(cell.getAttribute('rowspan') || 1);
                const colSpan = parseInt(cell.getAttribute('colspan') || 1);
                const txt = cell.innerText.split('\n')[0].replace(' (DOP)','').replace(' (UNIDADES)','').replace(' (mDOP)','').replace(' (MDOP)','').trim();
                
                for(let r=0; r<rowSpan; r++) {
                   for(let c=0; c<colSpan; c++) {
                       if(matrix[rowIndex + r]) {
                           matrix[rowIndex + r][colIndex + c] = txt;
                       }
                   }
                }
            });
        });
        
        let labels = [];
        let cols = matrix[0] ? matrix[0].length : 0;
        for(let c=0; c<cols; c++) {
           let parts = [];
           for(let r=0; r<matrix.length; r++) {
               let p = matrix[r][c];
               if (p && !parts.includes(p)) parts.push(p);
           }
           labels.push(parts.join(" - "));
        }
        
        const bodyTrs = tbody.querySelectorAll('tr');
        bodyTrs.forEach(tr => {
            const tds = tr.querySelectorAll('td');
            tds.forEach((td, idx) => {
                if (labels[idx]) {
                    td.setAttribute('data-label', labels[idx]);
                }
            });
        });
    }

    window.renderPgHorizontal = function() {
        if (!window.resumenComercialEngine) return;
        const selector = document.getElementById('monthSelector');
        let periodText = 'Periodo Actual';
        if (selector && globalFinancialData) {
            const idx = parseInt(selector.value);
            if (!isNaN(idx) && globalFinancialData[idx]) {
                const periodoInfo = globalFinancialData[idx].Periodo;
                if (periodoInfo && typeof periodoInfo === 'string') {
                    const parts = periodoInfo.split('-');
                    if (parts.length === 2) {
                        periodText = `${parts[0]} ${parts[1]}`;
                    } else {
                        periodText = periodoInfo;
                    }
                }
            }
        }
        
        window.resumenComercialEngine.renderPgHorizontal();
        setTimeout(() => {
            applyMobileDataLabels('pg-horizontal-table', 'pg-horizontal-thead');
            applyMobileDataLabels('pg-horizontal-unitarios-table', 'pg-horizontal-unitarios-thead');
        }, 10);
    };

    window.currentCostoProd = 'botellon';
    
    window.updateCostoUnitario = function() {
        const selector = document.getElementById('monthSelector');
        let m = selector && !isNaN(parseInt(selector.value)) ? parseInt(selector.value) : 3;
        let m_idx = 3;
        if (selector && Array.isArray(globalFinancialData) && globalFinancialData.length > m && globalFinancialData[m]) {
             let item = globalFinancialData[m];
             const match = item.date ? item.date.match(/ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC/i) : null;
             const monthsArr = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
             if (match) {
                 m_idx = monthsArr.indexOf(match[0].toUpperCase());
             }
        }
        if (window.costoUnitarioEngine && window.costoUnitarioEngine.hasCostoUnitarioData()) {
            window.costoUnitarioEngine.renderCostoUnitario(m_idx, window.currentCostoProd, window.costoUnitarioVista);
        }
    };

    window.renderResumenComercial = function() {
        if (!window.resumenComercialEngine) return;
        const selector = document.getElementById('monthSelector');
        let m = 3;
        let periodText = 'Periodo Actual';
        if (selector && globalFinancialData) {
            const idx = parseInt(selector.value);
            const item = globalFinancialData[idx];
            if (item) {
                periodText = item.date || 'Periodo';
                let dateObj = item.sortDate;
                if (dateObj) {
                    const d = new Date(dateObj);
                    if (!isNaN(d.getTime())) {
                        m = d.getMonth() + 1;
                    }
                } else if (item.date) {
                    const MESES_SEARCH = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
                    const lower = item.date.toLowerCase();
                    const found = MESES_SEARCH.findIndex(mes => lower.includes(mes));
                    if (found !== -1) m = found + 1;
                }
            }
        }
        
        const isYTD = (typeof isYTDMode !== 'undefined') ? isYTDMode : false;
        if (isYTD) periodText = "YTD " + periodText;
        
        const label = document.getElementById('resumenComercialPeriodLabel');
        if (label) {
            let viewLabel = "Resumen de Ventas";
            if (window.comercialCurrentView === 'mom') viewLabel = "MoM";
            if (window.comercialCurrentView === 'variacion') viewLabel = "Análisis de Variación";
            label.textContent = `| ${viewLabel} | ${periodText}`;
        }

        window.resumenComercialEngine.renderResumenComercial(m, isYTD, window.comercialCurrentView);
        setTimeout(() => {
            applyMobileDataLabels('resumen-comercial-table', 'resumen-comercial-thead');
            
            // Dynamic sticky header adjustment to prevent overlaps
            const thead = document.getElementById('resumen-comercial-thead');
            if (thead) {
                const firstRowTh = thead.querySelector('tr:first-child th:nth-child(2)');
                if (firstRowTh) {
                    const h = firstRowTh.getBoundingClientRect().height;
                    const secondRowThs = thead.querySelectorAll('tr:nth-child(2) th');
                    secondRowThs.forEach(th => th.style.top = `${h - 0.5}px`);
                }
            }
        }, 10);

        // Sync Mobile Accordions
        // Hemos deshabilitado esto a peticion del usuario dado a que los
        // grandes volumenes de datos crashean movil por problemas de memoria.
        // setTimeout(() => {
        //     if (typeof buildMobileAccordionsFromTable === 'function') {
        //         buildMobileAccordionsFromTable('resumen-comercial-table', 'resumenComercialMobileContainer');
        //     }
        // }, 100);
    };

    // Setup file upload listener for Ventas CEO (from detailed view if any)
    const handleVentasCeoUpload = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        await window.processVentasCeoFile(file);
        alert("Datos de Ventas CEO cargados exitosamente.");
    };
    
    const uploadVentasCeoInput = document.getElementById('upload-ventas-ceo');
    if (uploadVentasCeoInput) {
        uploadVentasCeoInput.addEventListener('change', handleVentasCeoUpload);
    }
    
    // El input del home (upload-ventas-ceo-home) ahora se maneja con el botón Procesar Archivos

    window.togglePrevVentasCEO = function() {
        const newVal = !window.ventasCeoColPrev;
        window.ventasCeoColPrev = newVal;
        window.ventasCeoColCurr = newVal;
        window.renderVentasCEO();
    };

    window.toggleCurrVentasCEO = function() {
        const newVal = !window.ventasCeoColCurr;
        window.ventasCeoColPrev = newVal;
        window.ventasCeoColCurr = newVal;
        window.renderVentasCEO();
    };

    window.renderVentasCEO = function(skipChart = false) {
        if (!ceoData || ceoData.length === 0) {
            // Attempt to build dynamically if Comercial is loaded
            if (window.hasComercialAccess && typeof window.buildVentasCeoFromComercial === 'function') {
                if (!window.buildingCeoDataFlag) {
                    window.buildingCeoDataFlag = true;
                    window.buildVentasCeoFromComercial().then(() => {
                        window.buildingCeoDataFlag = false;
                    }).catch(e => {
                        console.error('Failed to dynamically build CEO data', e);
                        window.buildingCeoDataFlag = false;
                    });
                    return; // Wait for the build to finish, it calls renderVentasCEO() itself
                }
            }

            const container = document.getElementById("ventas-ceo-table-container");
            const chartBox = document.querySelector("#view-ventas-ceo .chart-box");
            if (chartBox) chartBox.style.display = 'none';
            const cardsContainer = document.getElementById('ventas-ceo-cards-container');
            if (cardsContainer) cardsContainer.style.display = 'none';
            
            if (container) {
                container.style.display = 'block';
                container.innerHTML = `
                  <div style="padding:24px; background:white; border:1px solid var(--border); border-radius:8px; margin: 20px 0;">
                    <h3 style="margin:0 0 8px 0; color: var(--text);">Ventas CEO sin datos</h3>
                    <p style="margin:0; color:var(--text-secondary);">
                      El archivo de Ventas CEO se encontró, pero el parser produjo 0 filas.
                      Revisa la hoja, encabezados o estructura del Excel.
                    </p>
                  </div>
                `;
            }
            return;
        }

        const chartBox = document.querySelector("#view-ventas-ceo .chart-box");
        if (chartBox) chartBox.style.display = 'block';

        const isMobile = window.innerWidth <= 768;

        const thead = document.getElementById('ventas-ceo-thead');
        const tbody = document.getElementById('ventas-ceo-tbody');
        if(!thead || !tbody) return;
        
        tbody.innerHTML = '';
        thead.innerHTML = '';
        
        const tableContainer = document.getElementById('ventas-ceo-table-container');
        let cardsContainer = document.getElementById('ventas-ceo-cards-container');
        
        if (isMobile) {
            if (tableContainer) tableContainer.style.display = 'none';
            if (!cardsContainer) {
                cardsContainer = document.createElement('div');
                cardsContainer.id = 'ventas-ceo-cards-container';
                cardsContainer.className = 'ceo-cards-container';
                if (tableContainer && tableContainer.parentElement) {
                    tableContainer.parentElement.insertBefore(cardsContainer, tableContainer);
                }
            }
            cardsContainer.style.display = 'flex';
            cardsContainer.innerHTML = '';
        } else {
            if (tableContainer) tableContainer.style.display = '';
            if (cardsContainer) cardsContainer.style.display = 'none';
        }
        
        if(!window.expandedVentasCeoGroups) window.expandedVentasCeoGroups = new Set();
        
        const displayData = ceoData.filter(d => {
            if (d.Tipo !== ventasCeoCurrentMetric) return false;
            const p = d.Producto ? d.Producto.trim().toUpperCase() : '';
            return p !== 'TOTAL' && p !== 'TOTAL SIN BON' && p !== 'TOTAL SIN BON.' && p !== 'TOTAL AÑO' && p !== 'PA H+ 0.68 LTS (X12)' && p !== 'VENTAS NETAS DOP';
        });
        const isPrecio = ventasCeoCurrentMetric === 'Precio Unitario';
        const decimals = isPrecio ? 1 : 0;
        
        displayData.forEach(d => {
            // id, parentId, and hasChildren are already correctly set in parseConsejoFromObjects.
            // Just double check that we don't need CEO_HIERARCHY anymore.
        });
        
        // Dynamically get active month
        let selectedDate = new Date();
        let foundDate = false;
        const monthSelector = document.getElementById('monthSelector');
        if (monthSelector && globalFinancialData && globalFinancialData.length > 0) {
           const idx = parseInt(monthSelector.value, 10);
           if (!isNaN(idx) && globalFinancialData[idx]) {
               const item = globalFinancialData[idx];
               if (item.sortDate) {
                   selectedDate = new Date(item.sortDate);
               } else if (item.date) {
                   const mMatch = String(item.date).toLowerCase().match(/([a-z]+) (\d{4})/);
                   if (mMatch) {
                       const mIndex = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'].findIndex(m => m === mMatch[1].slice(0,3));
                       if (mIndex >= 0) {
                           selectedDate = new Date(parseInt(mMatch[2]), mIndex, 1);
                       } else {
                           selectedDate = new Date(item.date);
                       }
                   } else {
                       selectedDate = new Date(item.date);
                   }
               }
               
               if (!isNaN(selectedDate.getTime())) {
                   foundDate = true;
               }
           }
        } 
        
        if (!foundDate && ceoData && ceoData.length > 0) {
            // Find the most recent date available in ceoData values
            let maxKey = null;
            ceoData.forEach(d => {
                if (d.values) {
                    Object.keys(d.values).forEach(k => {
                        if (!maxKey || k > maxKey) maxKey = k;
                    });
                }
            });
            if (maxKey && typeof maxKey === 'string') {
                const [y, m] = maxKey.split('-');
                if (y && m) {
                    const tempDate = new Date(parseInt(y), parseInt(m) - 1, 1);
                    if (!isNaN(tempDate.getTime())) {
                        selectedDate = tempDate;
                        foundDate = true;
                    }
                }
            }
        }
        
        if (!foundDate || isNaN(selectedDate.getTime())) {
            selectedDate = new Date(); // Ultimate fallback
        }
        
        const currYear = selectedDate.getFullYear();
        const currMonth = selectedDate.getMonth();
        
        const formatM = (y, m) => {
           const d = new Date(y, m, 1);
           const str = d.toLocaleDateString('es-ES', { month: 'short' }).replace('.','');
           return `${str}-${String(y).slice(2)}`;
        };
        const formatKey = (y, m) => {
           let sm = String(m+1).padStart(2,'0');
           return `${y}-${sm}`;
        };
        
        const getMonthsArr = (endY, endM, count) => {
            let res = [];
            for(let i=count-1; i>=0; i--) {
                let m = endM - i;
                let y = endY;
                while(m < 0) { m += 12; y -= 1; }
                while(m > 11) { m -= 12; y += 1; }
                res.push({ key: formatKey(y, m), label: formatM(y, m) });
            }
            return res;
        };
        
        const currMonths = getMonthsArr(currYear, currMonth, 4);
        const prevMonths = getMonthsArr(currYear-1, currMonth, 4);
        
        const prevBtnStr = `<button onclick="togglePrevVentasCEO()" style="border: 1px solid rgba(255,255,255,0.4); background:transparent; color:white; border-radius:4px; width:20px; height:20px; display:inline-flex; align-items:center; justify-content:center; cursor:pointer; margin-left:8px; font-weight:bold; padding:0; line-height:1;" title="Mostrar/Ocultar meses">${window.ventasCeoColPrev ? '+' : '-'}</button>`;
        const currBtnStr = `<button onclick="toggleCurrVentasCEO()" style="border: 1px solid rgba(255,255,255,0.4); background:transparent; color:white; border-radius:4px; width:20px; height:20px; display:inline-flex; align-items:center; justify-content:center; cursor:pointer; margin-left:8px; font-weight:bold; padding:0; line-height:1;" title="Mostrar/Ocultar meses">${window.ventasCeoColCurr ? '+' : '-'}</button>`;
        
        const prevAvgLabel = `Promedio ${prevMonths[0].label} - ${prevMonths[3].label}${prevBtnStr}`;
        const currAvgLabel = `Promedio ${currMonths[0].label} - ${currMonths[3].label}${currBtnStr}`;
        
        let thHtml = `<th style="width: 24px; min-width: 24px; max-width: 24px; text-align: center; border:none; background: var(--sidebar); color: white; padding: 0;"></th>
                      <th style="text-align:left; background: var(--sidebar); color: white; border:none; border-right:1px solid rgba(255,255,255,0.2); padding: 12px 16px;">Producto</th>`;
        
        const addTh = (label, bg, color) => {
            thHtml += `<th style="background:${bg}; color:${color}; border:none; text-align:right; padding: 12px 8px; font-size: 0.7rem;">${label}</th>`;
        };
        
        // Static columns
        addTh(`Real 2024`, 'var(--sidebar)', 'white');
        addTh('<span title="Prom. Mensual Año Ant." style="cursor:help;">REAL AÑO ANT.</span>', 'var(--sidebar)', 'white');
        addTh('Var %', 'var(--sidebar)', 'white');
        addTh('PPTO', 'var(--sidebar)', 'white');
        
        const checkIsPpto = (monthKey) => !!(ceoData && ceoData.some(d => d.isPpto && d.isPpto[monthKey]));

        // Prev Months
        const displayPrevMonths = window.ventasCeoColPrev ? prevMonths.slice(2) : prevMonths;
        displayPrevMonths.forEach(m => {
            const isPpto = checkIsPpto(m.key);
            const label = isPpto ? `${m.label} (PPTO)` : m.label;
            const bg = isPpto ? '#e08924' : 'var(--sidebar)';
            addTh(label, bg, 'white');
        });
        addTh(prevAvgLabel, '#73A5C6', 'white'); // distinctive color
        
        // Curr Months
        const displayCurrMonths = window.ventasCeoColCurr ? currMonths.slice(2) : currMonths;
        displayCurrMonths.forEach(m => {
            const isPpto = checkIsPpto(m.key);
            const label = isPpto ? `${m.label} (PPTO)` : m.label;
            const bg = isPpto ? '#e08924' : 'var(--sidebar)';
            addTh(label, bg, 'white');
        });
        addTh(currAvgLabel, '#73A5C6', 'white');
        addTh('Var %', 'var(--sidebar)', 'white');
        
        if (!isMobile) {
            thead.innerHTML = `<tr>${thHtml}</tr>`;
        }
        
        const formatVal = (val) => {
            return parseFloat(val || 0).toLocaleString('es-DO', {minimumFractionDigits: decimals, maximumFractionDigits: decimals});
        };
        const formatPct = (val) => {
            return (parseFloat(val || 0) * 100).toFixed(1) + '%';
        };
        
        const renderRowContent = (row, isTotal) => {
            let html = '';
            
            const reqY = currYear;
            const reqM = String(currMonth + 1).padStart(2, '0');
            const prevY = currYear - 1;
            
            const currKey = `${reqY}-${reqM}`;
            const prevKey = `${prevY}-${reqM}`;
            
            // Assume upcoming year is reqY + 2
            const poStrKey = `${reqY + 2}-${reqM}`;
            
            let real24 = row.FY2024 || 0;
            
            let prevYearSum = 0;
            let prevYearCount = 12;
            for(let m=1; m<=12; m++) {
                let pKey = `${prevY}-${String(m).padStart(2, '0')}`;
                prevYearSum += (row.values[pKey] || 0);
            }
            let realAnoAnt = prevYearSum / prevYearCount;
            
            let po26 = (row.pptoValues && row.pptoValues[poStrKey]) ? row.pptoValues[poStrKey] : 0;
            
            if (po26 === 0 && row.pptoValues) {
                let altKey = `${reqY}-${reqM}`;
                if (row.pptoValues[altKey]) po26 = row.pptoValues[altKey];
            }
            
            row.__real24 = real24;
            row.__realAnoAnt = realAnoAnt;
            row.__po26 = po26;
            
            const varPct = realAnoAnt ? ((realAnoAnt - real24)/real24) : 0;
            
            const cellStyle = isTotal ? "font-weight:800; font-size:0.85rem;" : "font-size:0.95rem;";
            
            html += `<td style="text-align:right; ${cellStyle}">${formatVal(real24)}</td>`;
            html += `<td style="text-align:right; ${cellStyle}">${formatVal(realAnoAnt)}</td>`;
            html += `<td style="text-align:right; ${cellStyle}">${formatPct(varPct)}</td>`;
            html += `<td style="text-align:right; ${cellStyle}">${formatVal(po26)}</td>`;
            
            let prevSum = 0, prevCount = 0;
            prevMonths.forEach(m => {
                let v = row.values[m.key] || 0;
                prevSum += v; prevCount++;
                if (displayPrevMonths.find(dm => dm.key === m.key)) {
                    html += `<td style="text-align:right; ${cellStyle}">${formatVal(v)}</td>`;
                }
            });
            const prevAvg = prevCount ? (prevSum/prevCount) : 0;
            row.__prevAvg = prevAvg;
            html += `<td style="text-align:right; font-weight:600; font-size:0.95rem; background:rgba(115,165,198,0.1);">${formatVal(prevAvg)}</td>`;
            
            let currSum = 0, currCount = 0;
            currMonths.forEach(m => {
                let v = row.values[m.key] || 0;
                currSum += v; currCount++;
                if (displayCurrMonths.find(dm => dm.key === m.key)) {
                    html += `<td style="text-align:right; ${cellStyle}">${formatVal(v)}</td>`;
                }
            });
            const currAvg = currCount ? (currSum/currCount) : 0;
            row.__currAvg = currAvg;
            html += `<td style="text-align:right; font-weight:600; font-size:0.95rem; background:rgba(115,165,198,0.1);">${formatVal(currAvg)}</td>`;
            
            const varAvg = prevAvg ? ((currAvg - prevAvg)/prevAvg) : 0;
            html += `<td style="text-align:right; font-weight:600; font-size:0.95rem;">${formatPct(varAvg)}</td>`;
            
            return html;
        };
        
        let tbHtml = '';
        let cardsHtml = '';
        displayData.forEach(row => {
            let isVisible = true;
            if (row.parentId) {
                isVisible = window.expandedVentasCeoGroups.has(row.parentId);
            }
            if (!isVisible) return;
            
            let isExpanded = window.expandedVentasCeoGroups.has(row.id);
            let collapseBtn = "";
            let rowStyle = "";
            
            let rowOnclick = "";
            let rowHover = "";
            if (row.hasChildren) {
                collapseBtn = `<button class="collapse-btn" style="width: 16px; height: 16px; display: inline-flex; align-items: center; justify-content: center; padding: 0; line-height: 1; pointer-events: none; border: 1px solid var(--border); border-radius: 4px; background: white; color: var(--text-primary);">${isExpanded ? '-' : '+'}</button>`;
                rowStyle += "font-weight: 600; cursor: pointer; ";
                rowOnclick = `onclick="toggleVentasCeoGroup('${row.id}')"`;
                rowHover = `onmouseover="this.style.background='rgba(0,0,0,0.02)'" onmouseout="this.style.background=''"`;
            } else if (row.parentId) {
                rowStyle += "padding-left: 24px; color: var(--text-secondary); ";
            } else {
                rowStyle += "font-weight: 500; ";
            }

            if (isMobile) {
                renderRowContent(row, false); // Calculate data properties
                let currMonthVal = row.__real24;
                let varPct = row.__realAnoAnt ? ((row.__real24 - row.__realAnoAnt)/row.__realAnoAnt) : 0;
                let avgVal = row.__currAvg;
                
                let pctColor = varPct > 0 ? '#10b981' : (varPct < 0 ? '#ef4444' : 'var(--text-secondary)');
                let clickAttr = row.hasChildren ? `onclick="toggleVentasCeoGroup('${row.id}')" style="cursor:pointer;"` : '';
                let titleMargin = row.parentId ? "margin-left: 16px; border-left: 2px solid var(--border); padding-left: 8px;" : "";
                
                let selectedMonthLabel = formatM(currYear, currMonth);

                cardsHtml += `
                <div class="ceo-card" ${clickAttr}>
                    <div class="ceo-card-title" style="${titleMargin}">
                        <span>${formatSegmentName(row.Producto)}</span>
                        ${row.hasChildren ? `<span>${isExpanded ? '▼' : '►'}</span>` : ''}
                    </div>
                    <div class="ceo-card-metrics-grid">
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Real 2024</span>
                            <span class="ceo-card-metric-value">${formatVal(currMonthVal)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Real Año Ant.</span>
                            <span class="ceo-card-metric-value">${formatVal(row.__realAnoAnt)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Var %</span>
                            <span class="ceo-card-metric-value" style="color: ${pctColor}">${formatPct(varPct)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">PPTO</span>
                            <span class="ceo-card-metric-value">${formatVal(row.__po26)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">${selectedMonthLabel}</span>
                            <span class="ceo-card-metric-value">${formatVal(currMonthVal)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Prom. 4M actual</span>
                            <span class="ceo-card-metric-value">${formatVal(avgVal)}</span>
                        </div>
                    </div>
                </div>`;
            } else {
                tbHtml += `<tr data-group="${row.parentId || ''}" id="ventasceo-row-${row.id}" ${rowOnclick} ${rowHover}>
                              <td style="width: 24px; min-width: 24px; max-width: 24px; text-align: center; vertical-align: middle; border-right: 1px solid rgba(0,0,0,0.05); padding: 0; cursor: ${row.hasChildren ? 'pointer':'default'};">${collapseBtn}</td>
                              <td style="text-align:left; border-right: 1px solid rgba(0,0,0,0.05); padding: 12px 16px; ${rowStyle}">${formatSegmentName(row.Producto)}</td>`;
                tbHtml += renderRowContent(row, false);
                tbHtml += '</tr>';
            }
        });
        
        const totalRow = ceoData.find(d => {
            const p = d.Producto ? d.Producto.trim().toUpperCase() : '';
            return d.Tipo === ventasCeoCurrentMetric && p === 'TOTAL';
        });
        
        if(totalRow) {
             if (isMobile) {
                renderRowContent(totalRow, true);
                let varPct = totalRow.__realAnoAnt ? ((totalRow.__real24 - totalRow.__realAnoAnt)/totalRow.__realAnoAnt) : 0;
                let pctColor = varPct > 0 ? '#10b981' : (varPct < 0 ? '#ef4444' : 'var(--text-secondary)');
                let selectedMonthLabel = formatM(currYear, currMonth);
                
                cardsHtml = `
                <div class="ceo-card" style="background: #eef2f5; border: 2px solid var(--border);">
                    <div class="ceo-card-title">
                        <span>TOTAL</span>
                    </div>
                    <div class="ceo-card-metrics-grid">
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Real 2024</span>
                            <span class="ceo-card-metric-value">${formatVal(totalRow.__real24)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Real Año Ant.</span>
                            <span class="ceo-card-metric-value">${formatVal(totalRow.__realAnoAnt)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Var %</span>
                            <span class="ceo-card-metric-value" style="color: ${pctColor}">${formatPct(varPct)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">PPTO</span>
                            <span class="ceo-card-metric-value">${formatVal(totalRow.__po26)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">${selectedMonthLabel}</span>
                            <span class="ceo-card-metric-value">${formatVal(totalRow.__real24)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Prom. 4M actual</span>
                            <span class="ceo-card-metric-value">${formatVal(totalRow.__currAvg)}</span>
                        </div>
                    </div>
                </div>` + cardsHtml;
             } else {
                 const tRowHtml = renderRowContent(totalRow, true);
                 tbHtml = `<tr style="background:#eef2f5;">
                              <td style="width: 24px; min-width: 24px; max-width: 24px; text-align: center; border-right: 1px solid rgba(0,0,0,0.05); padding: 0;"></td>
                              <td style="text-align:left; font-weight:800; border-right: 1px solid rgba(0,0,0,0.05); padding: 12px 16px;">TOTAL</td>` 
                               + tRowHtml + 
                           '</tr>' + tbHtml;
             }
        }
        
        const tsbRow = ceoData.find(d => {
            const p = d.Producto ? d.Producto.trim().toUpperCase() : '';
            return d.Tipo === ventasCeoCurrentMetric && (p === 'TOTAL SIN BON' || p === 'TOTAL SIN BON.');
        });
        if(tsbRow) {
             if (isMobile) {
                renderRowContent(tsbRow, true);
                let varPct = tsbRow.__realAnoAnt ? ((tsbRow.__real24 - tsbRow.__realAnoAnt)/tsbRow.__realAnoAnt) : 0;
                let pctColor = varPct > 0 ? '#10b981' : (varPct < 0 ? '#ef4444' : 'var(--text-secondary)');
                let selectedMonthLabel = formatM(currYear, currMonth);
                
                cardsHtml += `
                <div class="ceo-card" style="background: #eef2f5; border: 2px solid #10b981;">
                    <div class="ceo-card-title" style="color: #10b981;">
                        <span>TOTAL SIN BON</span>
                    </div>
                    <div class="ceo-card-metrics-grid">
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Real 2024</span>
                            <span class="ceo-card-metric-value" style="color: #10b981;">${formatVal(tsbRow.__real24)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Real Año Ant.</span>
                            <span class="ceo-card-metric-value" style="color: #10b981;">${formatVal(tsbRow.__realAnoAnt)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Var %</span>
                            <span class="ceo-card-metric-value" style="color: ${pctColor}">${formatPct(varPct)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">PPTO</span>
                            <span class="ceo-card-metric-value" style="color: #10b981;">${formatVal(tsbRow.__po26)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">${selectedMonthLabel}</span>
                            <span class="ceo-card-metric-value" style="color: #10b981;">${formatVal(tsbRow.__real24)}</span>
                        </div>
                        <div class="ceo-card-metric">
                            <span class="ceo-card-metric-label">Prom. 4M actual</span>
                            <span class="ceo-card-metric-value" style="color: #10b981;">${formatVal(tsbRow.__currAvg)}</span>
                        </div>
                    </div>
                </div>`;
             } else {
                 tbHtml += `<tr style="background:#eef2f5;">
                               <td style="width: 24px; min-width: 24px; max-width: 24px; text-align: center; border-right: 1px solid rgba(0,0,0,0.05); padding: 0;"></td>
                               <td style="text-align:left; font-weight:800; color: #10b981; border-right: 1px solid rgba(0,0,0,0.05); padding: 12px 16px;">${formatSegmentName('TOTAL SIN BON')}</td>` 
                               + renderRowContent(tsbRow, true).replace(/<td/g, '<td style="color: #10b981;"') + 
                           '</tr>';
             }
        }

        if (!isMobile) {
            tbody.innerHTML = tbHtml;
        }
        if (isMobile) {
            const container = document.getElementById('ventas-ceo-cards-container');
            if (container) container.innerHTML = cardsHtml;
        }
        
        let chartMonths = ['__real24', '__realAnoAnt', '__po26'];
        let chartLabels = ['Real 2024', 'REAL AÑO ANTERIOR', 'PPTO'];
        
        const dividers = [];

        dividers.push({ left: 'PPTO' }); // Divider after PPTO

        displayPrevMonths.forEach((m, idx) => { 
            const isPpto = checkIsPpto(m.key);
            const label = isPpto ? `${m.label} (PPTO)` : m.label;
            chartMonths.push(m.key); chartLabels.push(label); 
            if (idx === displayPrevMonths.length - 1) {
                dividers.push({ left: label });
            }
        });
        
        const prevAvgChartLabel = `Promedio\n${prevMonths[0].label} - ${prevMonths[3].label}`;
        chartMonths.push('__prevAvg'); chartLabels.push(prevAvgChartLabel);
        dividers.push({ left: prevAvgChartLabel });
        
        displayCurrMonths.forEach((m, idx) => { 
            const isPpto = checkIsPpto(m.key);
            const label = isPpto ? `${m.label} (PPTO)` : m.label;
            chartMonths.push(m.key); chartLabels.push(label);
            if (idx === displayCurrMonths.length - 1) {
                dividers.push({ left: label });
            }
        });
        
        const currAvgChartLabel = `Promedio\n${currMonths[0].label} - ${currMonths[3].label}`;
        chartMonths.push('__currAvg'); chartLabels.push(currAvgChartLabel);

        const chartDataRows = displayData.filter(d => {
            const isVisible = !d.parentId || window.expandedVentasCeoGroups.has(d.parentId);
            if (!isVisible) return false;
            
            const isExpanded = d.hasChildren && window.expandedVentasCeoGroups.has(d.id);
            if (isExpanded) return false;
            
            return d.values !== undefined;
        });
        
        if (skipChart) return;

        renderVentasCeoChart(chartDataRows, chartMonths, chartLabels, dividers);
        updateVentasButtons();
    }
    
    if(window.ventasCeoColPrev === undefined) window.ventasCeoColPrev = true;
    if(window.ventasCeoColCurr === undefined) window.ventasCeoColCurr = true;

    window.toggleVentasCeoGroup = function(groupId, btn) {
        if(window.expandedVentasCeoGroups.has(groupId)) {
            window.expandedVentasCeoGroups.delete(groupId);
        } else {
            window.expandedVentasCeoGroups.add(groupId);
        }
        window.renderVentasCEO(true);
    };
    
    document.getElementById('btn-ventas-expandir')?.addEventListener('click', () => {
        if(!window.expandedVentasCeoGroups) window.expandedVentasCeoGroups = new Set();
        (ceoData || []).forEach(d => {
            if(d.hasChildren && d.id) window.expandedVentasCeoGroups.add(d.id);
        });
        window.renderVentasCEO(true);
    });
    
    document.getElementById('btn-ventas-colapsar')?.addEventListener('click', () => {
        if(window.expandedVentasCeoGroups) window.expandedVentasCeoGroups.clear();
        window.renderVentasCEO(true);
    });
    
    function renderVentasCeoChart(displayData, dateCols, dateLabels, dividers) {
        const container = document.getElementById('ventas-ceo-chart');
        if (!container) return;
        container.innerHTML = '';
        
        if (displayData.length === 0) return;
        
        const chartData = [];
        dateCols.forEach((c, idx) => {
            let label = dateLabels ? dateLabels[idx] : c;
            const isPpto = label.includes('(PPTO)');
            let item = { label: label, date: c, isPpto: isPpto };
            let total = 0;
            displayData.forEach(row => {
                let val = 0;
                if(row.values && row.values[c] !== undefined) {
                    val = parseFloat(row.values[c]);
                } else {
                    val = parseFloat(row[c]) || 0;
                }
                if (isNaN(val)) val = 0;
                item[row.Producto] = val;
                total += val;
            });
            item.total = total;
            chartData.push(item);
        });
        
        const isMobile = window.innerWidth <= 768;
        const margin = isMobile ? { top: 24, right: 12, bottom: 64, left: 42 } : { top: 40, right: 300, bottom: 140, left: 70 };
        const width = container.clientWidth;
        const height = container.clientHeight;
        const boundedWidth = width - margin.left - margin.right;
        const boundedHeight = height - margin.top - margin.bottom;

        if (boundedWidth <= 0 || boundedHeight <= 0) return;

        const svg = d3.select(container)
            .append("svg")
            .attr("width", width)
            .attr("height", height);

        const g = svg.append("g")
            .attr("transform", `translate(${margin.left},${margin.top})`);
            
        const seriesKeys = displayData.map(d => d.Producto);
        const stack = d3.stack().keys(seriesKeys)(chartData);
        
        const x = d3.scaleBand()
            .domain(chartData.map(d => d.label))
            .range([0, boundedWidth])
            .padding(0.2);

        const yMax = d3.max(chartData, d => d.total) * 1.1;
        const y = d3.scaleLinear()
            .domain([0, yMax])
            .range([boundedHeight, 0])
            .nice();

        const top3PerDate = {};
        chartData.forEach(d => {
            const vals = seriesKeys.map(k => ({ key: k, val: d[k] || 0 }));
            vals.sort((a,b) => b.val - a.val);
            top3PerDate[d.label] = vals.slice(0, 3).map(v => v.key);
        });

        const colorScale = d3.scaleOrdinal(d3.schemeTableau10).domain(seriesKeys);

        const defs = svg.append("defs");
        defs.append("pattern")
            .attr("id", "pattern-stripe")
            .attr("width", 8)
            .attr("height", 8)
            .attr("patternUnits", "userSpaceOnUse")
            .attr("patternTransform", "rotate(45)")
            .append("rect")
            .attr("width", 4)
            .attr("height", 8)
            .attr("transform", "translate(0,0)")
            .attr("fill", "rgba(255, 255, 255, 0.3)");

        const layer = g.selectAll("g.layer")
            .data(stack)
            .enter().append("g")
            .attr("class", "layer")
            .attr("fill", d => colorScale(d.key));
            
        const rects = layer.selectAll("g.bar-group")
            .data(d => d)
            .enter().append("g")
            .attr("class", "bar-group");

        // Base colored rectangle
        rects.append("rect")
            .attr("x", d => x(d.data.label))
            .attr("y", d => y(d[1]))
            .attr("height", d => Math.max(0, y(d[0]) - y(d[1])))
            .attr("width", x.bandwidth())
            .attr("opacity", d => d.data.isPpto ? 0.6 : 1)
            .on("mouseover", function(event, d) {
                const subName = formatSegmentName(d3.select(this.parentNode.parentNode).datum().key);
                d3.select(this).attr("opacity", d.data.isPpto ? 0.8 : 0.8);
                const tip = d3.select("body").append("div")
                    .attr("class", "d3-tooltip")
                    .style("opacity", 1)
                    .html(`<strong>${subName}${d.data.isPpto ? ' (PPTO)' : ''}</strong><br/>${d.data.label}<br/>Valor: ${(d[1]-d[0]).toLocaleString('es-DO', {maximumFractionDigits:1})}`);
                const rect = this.getBoundingClientRect();
                tip.style("left", (rect.left + window.pageXOffset) + "px")
                   .style("top", (rect.top + window.pageYOffset - 40) + "px");
            })
            .on("mouseout", function(event, d) {
                d3.select(this).attr("opacity", d.data.isPpto ? 0.6 : 1);
                d3.selectAll(".d3-tooltip").remove();
            });

        // Overlay pattern for PPTO
        rects.append("rect")
            .filter(d => d.data.isPpto)
            .attr("x", d => x(d.data.label))
            .attr("y", d => y(d[1]))
            .attr("height", d => Math.max(0, y(d[0]) - y(d[1])))
            .attr("width", x.bandwidth())
            .attr("fill", "url(#pattern-stripe)")
            .style("pointer-events", "none");

        const isPrecio = ventasCeoCurrentMetric === 'Precio Unitario';
        const formatter = new Intl.NumberFormat('es-DO', { 
            minimumFractionDigits: isPrecio ? 1 : 0, 
            maximumFractionDigits: isPrecio ? 1 : 0 
        });

        layer.selectAll("text.segment-label")
            .data(d => d)
            .enter().append("text")
            .attr("class", "segment-label")
            .attr("x", d => x(d.data.label) + x.bandwidth() / 2)
            .attr("y", d => y(d[1]) + (y(d[0]) - y(d[1])) / 2 + 3)
            .attr("text-anchor", "middle")
            .attr("fill", "white")
            .style("font-size", "12px")
            .style("font-weight", "600")
            .style("pointer-events", "none")
            .text(function(d) {
                if (isMobile) return ""; // Hide internal stack labels on mobile
                const subName = d3.select(this.parentNode).datum().key;
                const top3 = top3PerDate[d.data.label] || [];
                const heightPx = y(d[0]) - y(d[1]);
                if (top3.includes(subName) && heightPx > 20) {
                    return formatter.format(d[1] - d[0]);
                }
                return "";
            });

        if (!isPrecio) {
            g.selectAll("text.total-label")
                .data(chartData)
                .enter().append("text")
                .attr("class", "total-label")
                .attr("x", d => x(d.label) + x.bandwidth() / 2)
                .attr("y", d => y(d.total) - 8)
                .attr("text-anchor", "middle")
                .style("font-size", "13px")
                .style("font-weight", "bold")
                .style("fill", "var(--text-primary)")
                .text(d => formatter.format(d.total));
        }

        g.append("g")
            .attr("transform", `translate(0,${boundedHeight})`)
            .call(d3.axisBottom(x).tickSize(0).tickPadding(10))
            .selectAll("text")
            .style("text-anchor", "middle")
            .style("font-size", "11px")
            .style("font-weight", "600")
            .style("fill", "var(--sidebar-dark)")
            .each(function(d) {
                 const textGroup = d3.select(this);
                 textGroup.text(""); // Clear original text
                 if(d.includes("\n")) {
                     const lines = d.split("\n");
                     lines.forEach((line, index) => {
                         textGroup.append("tspan")
                             .attr("x", 0)
                             .attr("dy", index === 0 ? "0" : "1.2em")
                             .text(line);
                     });
                 } else {
                     textGroup.append("tspan")
                         .attr("x", 0)
                         .attr("dy", "0")
                         .text(d);
                 }
            });

        if (dividers && dividers.length > 0) {
            dividers.forEach(div => {
                let leftX = x(div.left);
                if (leftX !== undefined) {
                    let dividerX = leftX + x.bandwidth() + (x.step() - x.bandwidth()) / 2;
                    g.append("line")
                        .attr("x1", dividerX)
                        .attr("y1", -20)
                        .attr("x2", dividerX)
                        .attr("y2", boundedHeight + 40)
                        .attr("stroke", "#0ea5e9") // Light blue
                        .attr("stroke-width", 1.5);
                }
            });
        }

        g.append("g")
            .call(d3.axisLeft(y).ticks(5))
            .selectAll("text")
            .style("font-size", "12px");
        
        // Add additional annotations for __currAvg vs __prevAvg
        const currIdx = chartData.findIndex(d => d.date === '__currAvg');
        const prevIdx = chartData.findIndex(d => d.date === '__prevAvg');
        
        // Helper to format short names in the legend if needed so it doesn't overflow
        const getShortName = (name) => {
            if (!name) return "";
            name = formatSegmentName(name);
            if (name.length <= 25) return name;
            let s = name.toUpperCase();
            if (s.includes("BOTELLON") && s.includes("18.9")) return s.includes("MAQUILA") ? "MAQ. BOTELLON 18.9L" : "APA BOTELLON 18.9L";
            if (s.includes("1.5 LTS")) return s.includes("MAQUILA") ? "MAQ. 1.5L" : "APA 1.5L";
            if (s.includes("0.5 LTS")) return s.includes("MAQUILA") ? "MAQ. 0.5L" : (s.includes("SABOR") ? "SABOR 0.5L" : "APA 0.5L");
            if (s.includes("0.71 LTS")) return s.includes("H+") ? "PA H+ 0.71L" : "PA 0.71L";
            if (s.includes("OTRAS") || s.includes("OTROS")) return s.includes("MAQUILA") ? "MAQ. OTROS" : "APA OTRAS";
            return name.slice(0, 22) + '...';
        };

        if (currIdx !== -1 && prevIdx !== -1) {
            const currItem = chartData[currIdx];
            const prevItem = chartData[prevIdx];
            
            const formatterPct = new Intl.NumberFormat('es-DO', { style: 'percent', minimumFractionDigits: 0, maximumFractionDigits: 1, signDisplay: 'always' });
            
            const lastBarX = x(currItem.label) + x.bandwidth();
            const lineX = lastBarX + 15;
            
            g.append("line")
                .attr("x1", lineX)
                .attr("y1", 0)
                .attr("x2", lineX)
                .attr("y2", boundedHeight)
                .attr("stroke", "var(--sidebar)")
                .attr("stroke-width", 2)
                .attr("stroke-dasharray", "4,4")
                .attr("opacity", 0.5);
                
            if (!isMobile) {
                const annotG = g.append("g")
                    .attr("transform", `translate(${lineX + 10}, 0)`);
                    
                annotG.append("text")
                    .attr("x", 0)
                    .attr("y", 10)
                    .attr("fill", "var(--sidebar)")
                    .style("font-size", "13px")
                    .style("font-weight", "bold")
                    .html(`<tspan x="0" dy="0">Ult. 4 meses</tspan><tspan x="0" dy="16">% vs. AA</tspan>`);
                    
                let totalPct = prevItem.total !== 0 ? (currItem.total - prevItem.total) / prevItem.total : (currItem.total > 0 ? 1 : 0);
                
                annotG.append("text")
                    .attr("x", 0)
                    .attr("y", 55)
                    .attr("fill", "var(--primary)")
                    .style("font-size", "14px")
                    .style("font-weight", "bold")
                    .text(`Total: ${formatterPct.format(totalPct)}`);
                
                // Build pct map
                const pctMap = {};
                seriesKeys.forEach(k => {
                    let prev = prevItem[k] || 0;
                    let curr = currItem[k] || 0;
                    pctMap[k] = prev !== 0 ? (curr - prev) / prev : (curr > 0 ? 1 : 0);
                });
    
                // Put legend exactly here
                const legend = annotG.append("g")
                    .attr("font-family", "sans-serif")
                    .attr("font-size", 12)
                    .attr("text-anchor", "start")
                    .selectAll("g")
                    .data(seriesKeys)
                    .enter().append("g")
                    .attr("transform", (d, i) => `translate(0,${i * 20 + 80})`);
    
                legend.append("rect")
                    .attr("x", 0)
                    .attr("width", 15)
                    .attr("height", 15)
                    .attr("fill", colorScale);
    
                legend.append("text")
                    .attr("x", 20)
                    .attr("y", 7.5)
                    .attr("dy", "0.32em")
                    .style("font-size", "12px")
                    .attr("fill", "var(--text-primary)")
                    .html(d => {
                        let shortName = getShortName(d);
                        let pctValue = pctMap[d];
                        // Make the percentage bold and colored
                        let color = pctValue >= 0 ? "var(--success)" : "var(--destructive)";
                        let pctStr = formatterPct.format(pctValue);
                        let displayStr = `${shortName}: `;
                        return `<tspan>${displayStr}</tspan><tspan fill="${color}" font-weight="bold">${pctStr}</tspan>`;
                    });
            }
        } else if (!isMobile) {
            // standard legend
            const legend = svg.append("g")
                .attr("font-family", "sans-serif")
                .attr("font-size", 12)
                .attr("text-anchor", "start")
                .selectAll("g")
                .data(seriesKeys)
                .enter().append("g")
                .attr("transform", (d, i) => `translate(${width - 200},${i * 20 + margin.top})`); // pushed legend a bit left

            legend.append("rect")
                .attr("x", 0)
                .attr("width", 15)
                .attr("height", 15)
                .attr("fill", colorScale);

            legend.append("text")
                .attr("x", 20)
                .attr("y", 7.5)
                .attr("dy", "0.32em")
                .style("font-size", "12px")
                .attr("fill", "var(--text-primary)")
                .text(d => getShortName(d));
        }
    }
    
    // PG Horizontal Perspective Toggles
    const btnPgTotales = document.getElementById('btn-pg-totales');
    const btnPgUnitarios = document.getElementById('btn-pg-unitarios');
    const pgTableTotales = document.getElementById('pg-horizontal-table');
    const pgTableUnitarios = document.getElementById('pg-horizontal-unitarios-table');

    if (btnPgTotales && btnPgUnitarios && pgTableTotales && pgTableUnitarios) {
        btnPgTotales.addEventListener('click', () => {
            btnPgTotales.style.background = 'white';
            btnPgTotales.style.color = 'var(--primary)';
            btnPgTotales.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
            
            btnPgUnitarios.style.background = 'transparent';
            btnPgUnitarios.style.color = 'var(--text-secondary)';
            btnPgUnitarios.style.boxShadow = 'none';

            pgTableTotales.style.display = 'table';
            pgTableUnitarios.style.display = 'none';
        });

        btnPgUnitarios.addEventListener('click', () => {
            btnPgUnitarios.style.background = 'white';
            btnPgUnitarios.style.color = 'var(--primary)';
            btnPgUnitarios.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
            
            btnPgTotales.style.background = 'transparent';
            btnPgTotales.style.color = 'var(--text-secondary)';
            btnPgTotales.style.boxShadow = 'none';

            pgTableTotales.style.display = 'none';
            pgTableUnitarios.style.display = 'table';
        });
    }

    document.getElementById('btn-ventas-vol')?.addEventListener('click', () => {
        ventasCeoCurrentMetric = 'Volumen';
        updateVentasButtons();
        window.renderVentasCEO();
    });
    document.getElementById('btn-ventas-monto')?.addEventListener('click', () => {
        ventasCeoCurrentMetric = 'Monto (MM DOP)';
        updateVentasButtons();
        window.renderVentasCEO();
    });
    document.getElementById('btn-ventas-precio')?.addEventListener('click', () => {
        ventasCeoCurrentMetric = 'Precio Unitario';
        updateVentasButtons();
        window.renderVentasCEO();
    });
    
    function updateVentasButtons() {
        const resetBtn = (id) => {
            const btn = document.getElementById(id);
            if(btn) {
                btn.style.background = 'transparent';
                btn.style.color = 'var(--text-secondary)';
                btn.style.boxShadow = 'none';
            }
        };
        resetBtn('btn-ventas-vol');
        resetBtn('btn-ventas-monto');
        resetBtn('btn-ventas-precio');
        
        let activeId = 'btn-ventas-vol';
        let chartTitle = 'Volumen (k) de Unidades';
        
        if(ventasCeoCurrentMetric === 'Monto (MM DOP)') { 
            activeId = 'btn-ventas-monto';
            chartTitle = 'Monto (mDOP)';
        }
        if(ventasCeoCurrentMetric === 'Precio Unitario') { 
            activeId = 'btn-ventas-precio';
            chartTitle = 'Precio Unitario';
        }
        
        const activeBtn = document.getElementById(activeId);
        if(activeBtn) {
            activeBtn.style.background = 'white';
            activeBtn.style.color = 'var(--primary)';
            activeBtn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
        }
        
        const titleEl = document.getElementById('ventas-ceo-chart-title');
        if(titleEl) {
            titleEl.textContent = chartTitle;
        }
    }

    const btnGenerateInsights = document.getElementById('btn-generate-insights');
    if (btnGenerateInsights) {
        btnGenerateInsights.addEventListener('click', async () => {
            const btn = btnGenerateInsights;
            const content = document.getElementById('ai-insights-content');
            
            const monthSelector = document.getElementById('monthSelector');
            const idx = monthSelector ? parseInt(monthSelector.value, 10) : (globalFinancialData ? globalFinancialData.length - 1 : 0);
            
            const dataToAnalyze = (globalFinancialData && globalFinancialData.length > 0) ? globalFinancialData[idx] : null;

            if (!dataToAnalyze) {
                content.innerHTML = '<span style="color: var(--destructive)">No hay datos disponibles para analizar. Cargue un archivo financiero o espere a que se procese.</span>';
                return;
            }

            btn.disabled = true;
            btn.innerHTML = '<i data-lucide="loader" class="spin"></i> Analizando...';
            lucide.createIcons();

            try {
                const summaryInfo = {
                    date: dataToAnalyze.date,
                    kpis: dataToAnalyze.kpis,
                    pnl_summary: dataToAnalyze.pnl,
                    balance: dataToAnalyze.balance
                };

                const response = await fetch('/api/gemini/insights', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ financialData: summaryInfo })
                });

                if (!response.ok) {
                    throw new Error(`Error: ${response.status} ${response.statusText}`);
                }

                const result = await response.json();
                
                if (result.error) {
                    throw new Error(result.error);
                }

                content.innerHTML = result.insight;
            } catch (err) {
                console.error(err);
                content.innerHTML = `<span style="color: var(--destructive)">Falló la generación de insights: ${err.message}</span>`;
            } finally {
                btn.disabled = false;
                btn.innerHTML = 'Actualizar Resumen';
                lucide.createIcons();
            }
        });
    }

    // =======================================================
    // MANUAL DEL USUARIO, IMPRESIÓN PDF & TOUR INTERACTIVO
    // =======================================================

    // Función para navegación y sombreado
    window.highlightInstNav = function(element) {
        document.querySelectorAll('.inst-nav-link').forEach(link => {
            link.classList.remove('active');
            link.style.color = 'var(--text-secondary)';
            link.style.background = 'transparent';
        });
        element.classList.add('active');
        element.style.color = 'white';
        element.style.background = '#38bdf8';
    };

    // Estilos personalizados para el Tour Virtual Spotlight
    const tourStyle = document.createElement('style');
    tourStyle.innerHTML = `
        #tour-modal-backdrop {
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(15, 23, 42, 0.4);
            z-index: 10000;
            pointer-events: none;
            display: none;
            transition: opacity 0.3s;
        }
        #tour-spotlight-box {
            position: absolute;
            border-radius: 8px;
            box-shadow: 0 0 0 9999px rgba(15, 23, 42, 0.7);
            z-index: 10001;
            pointer-events: auto;
            display: none;
            border: 3px solid #38bdf8;
            box-sizing: border-box;
            transition: all 0.3s ease;
        }
        #tour-tooltip-box {
            position: fixed;
            width: 380px;
            background: #ffffff;
            color: #1e293b;
            border-radius: 12px;
            padding: 24px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.3);
            z-index: 10002;
            display: none;
            box-sizing: border-box;
            border-top: 5px solid #0284c7;
            transition: all 0.3s ease;
        }
        .tour-btn {
            padding: 8px 16px;
            font-size: 0.85rem;
            font-weight: 700;
            border-radius: 6px;
            cursor: pointer;
            border: none;
            transition: 0.2s;
        }
        .tour-btn-primary {
            background: #0284c7;
            color: white;
        }
        .tour-btn-secondary {
            background: #f1f5f9;
            color: #475569;
        }
        .tour-btn-primary:hover {
            background: #0369a1;
        }
        .tour-btn-secondary:hover {
            background: #e2e8f0;
        }
    `;
    document.head.appendChild(tourStyle);

    // Crear elementos del Tour en el DOM si no existen
    let backdrop = document.getElementById('tour-modal-backdrop');
    if (!backdrop) {
        backdrop = document.createElement('div');
        backdrop.id = 'tour-modal-backdrop';
        document.body.appendChild(backdrop);
    }
    let spotlight = document.getElementById('tour-spotlight-box');
    if (!spotlight) {
        spotlight = document.createElement('div');
        spotlight.id = 'tour-spotlight-box';
        document.body.appendChild(spotlight);
    }
    let tooltip = document.getElementById('tour-tooltip-box');
    if (!tooltip) {
        tooltip = document.createElement('div');
        tooltip.id = 'tour-tooltip-box';
        document.body.appendChild(tooltip);
    }

    let currentTourStep = 0;
    const tourSteps = [
        {
            elementId: "monthSelector",
            title: "🗓️ 1. Panel Global de Periodos",
            text: "Selector global del periodo de análisis. A diferencia de un archivo estático, modificar el mes aquí actualiza en tiempo real todas las vistas, gráficas y tablas de la plataforma.",
            action: () => { document.getElementById("menu-kpi")?.click(); }
        },
        {
            elementId: "ytdToggleContainer",
            title: "🔄 2. Selector Temporal (Mensual vs Acumulado)",
            text: "Alterna el foco de análisis con un clic. Visualiza el desempeño aislado del mes seleccionado (Mensual) o el acumulado contable desde enero hasta dicho periodo (Acumulado YTD).",
            action: () => { document.getElementById("menu-kpi")?.click(); }
        },
        {
            elementId: "icon-seguimiento",
            title: "📂 3. Navegación Agrupada",
            text: "Los módulos corporativos están agrupados estratégicamente. Haz clic en encabezados como 'Modelo de Seguimiento' o 'Ventas' para desplegar u ocultar tableros sin saturar el entorno visual.",
            action: () => { 
               // Forzar visualmente que al menos la sección este visible
               const grupo = document.getElementById('grupo-seguimiento');
               if (grupo && grupo.style.maxHeight === '0px') {
                   document.getElementById('icon-seguimiento').click();
               }
            }
        },
        {
            elementId: "cxp-view-toggles",
            title: "🗂️ 4. Sub-vistas Internas (Detalles y Resúmenes)",
            text: "Maximiza el análisis en pantallas complejas (ej. Detalle CxP, P&G Horizontal) utilizando las pestañas superiores derechas, permitiendo transiciones fluidas entre resúmenes gráficos y data analítica dura.",
            action: () => { 
               document.getElementById("menu-cxp")?.click(); 
            }
        },
        {
            elementId: "btn-comercial-resumen",
            title: "📊 5. Perspectiva: Resumen de Ventas",
            text: "Despliega el consolidado base. Identifica inmediatamente desviaciones en volumen, precio promedio unitario neto y recaudación de ventas por categoría, contrastado frente al presupuesto (PPTO) o año anterior.",
            action: () => { 
               document.getElementById("menu-resumen-comercial")?.click(); 
               setTimeout(() => {
                   document.getElementById("btn-comercial-resumen")?.click();
               }, 50);
            }
        },
        {
            elementId: "btn-comercial-mom",
            title: "📈 6. Perspectiva: Tendencia Secuencial (MoM)",
            text: "Habilita la evaluación secuencial mes a mes (Month-over-Month). Fundamental para detectar cambios dinámicos de mercado, crecimientos sostenidos o caídas inusuales en ciclos cortos.",
            action: () => { 
               document.getElementById("menu-resumen-comercial")?.click(); 
               setTimeout(() => {
                   document.getElementById("btn-comercial-mom")?.click();
               }, 50);
            }
        },
        {
            elementId: "btn-comercial-variacion",
            title: "🔍 7. Perspectiva: Análisis de Variación",
            text: "Aísla matemáticamente la varianza comercial. El sistema disgrega si una mejora en los ingresos netos fue el resultado directo de una agresiva estrategia de Precios o de una mayor penetración física (Volumen).",
            action: () => { 
               document.getElementById("menu-resumen-comercial")?.click(); 
               setTimeout(() => {
                   document.getElementById("btn-comercial-variacion")?.click();
               }, 50);
            }
        },
        {
            elementId: "tour-alert-target",
            title: "🚨 8. Alertas de Desviación",
            text: "Lógica: Los indicadores visuales intermitentes en el Estado de Resultados son alertas de desviación. Rojo intermitente señala déficit agudo y Verde intermitente señala superávit extraordinario en la ejecución (>15% vs meta).",
            action: () => { 
               document.getElementById("menu-pnl")?.click(); 
               setTimeout(() => {
                   let pulseTarget = document.querySelector('#pnlDetailedTable .pulse-neg') || document.querySelector('#pnlDetailedTable .pulse-pos');
                   if (pulseTarget) {
                       pulseTarget.id = 'tour-alert-target';
                   } else {
                       let tbody = document.getElementById('pnlDetailedBody');
                       if (tbody && tbody.rows.length > 0) {
                           tbody.rows[0].cells[0].id = 'tour-alert-target';
                       }
                   }
               }, 100);
            }
        },
        {
            elementId: "menu-simulador",
            title: "🚀 9. Operando el Simulador (What-If)",
            text: "Entorno de proyección prospectivo. Ajusta los parámetros interactivos (Sliders de Precio, Costos, Volumen) para modelar su efecto instantáneo en la rentabilidad u operativa (EBITDA) y previsiones de liquidity.",
            action: () => { document.getElementById("menu-simulador")?.click(); }
        },
        {
            elementId: "tour-export-center",
            title: "📥 10. Centro de Descarga",
            text: "Exporte el Dashboard Ejecutivo en formatos profesionales. Puede descargar reportes consolidados en CSV o generar un documento PDF de alta calidad con el layout del tablero actual.",
            action: () => { 
               document.getElementById("menu-config")?.click(); 
               setTimeout(() => {
                   document.getElementById('tour-export-center')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
               }, 100);
            }
        },
        {
            elementId: "btn-download-manual-pdf",
            title: "🎓 11. Guía de Interpretación",
            text: "¡Ha completado el preámbulo funcional! Utilice este botón para exportar la Guía Corporativa PDF, la cual asiste en cómo interpretar los balances financieros clínicamente a profundidad.",
            action: () => { document.getElementById("menu-instructivo")?.click(); }
        }
    ];

    function renderTourStep() {
        const step = tourSteps[currentTourStep];
        if (!step) {
            endTour();
            return;
        }

        // Ejecutar acción del paso (cambio de tab, scroll, etc.)
        if (typeof step.action === 'function') {
            step.action();
        }

        // Dar un breve delay estructurado para permitir renderizados de pestañas
        setTimeout(() => {
            const target = document.getElementById(step.elementId);
            if (!target) {
                // Si el elemento no existe o está oculto, pasar al siguiente
                nextStep();
                return;
            }

            backdrop.style.display = 'block';
            tooltip.style.display = 'block';
            spotlight.style.display = 'block';

            // Configurar contenido del diálogo flotante
            tooltip.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                    <span style="font-size:0.75rem; text-transform:uppercase; font-weight:800; color:#0284c7; background:rgba(2,132,199,0.1); padding:2px 8px; border-radius:12px;">Paso ${currentTourStep + 1} de ${tourSteps.length}</span>
                    <button onclick="window.endTour()" style="background:none; border:none; font-size:1.1rem; color:#94a3b8; cursor:pointer;">&times;</button>
                </div>
                <h4 style="margin:0 0 8px 0; font-size:1.1rem; font-weight:800; color:#1e293b;">${step.title}</h4>
                <p style="margin:0 0 20px 0; font-size:0.88rem; line-height:1.5; color:#475569;">${step.text}</p>
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <button onclick="window.prevStep()" class="tour-btn tour-btn-secondary" ${currentTourStep === 0 ? 'disabled style="opacity:0.5; cursor:not-allowed;"' : ''}>Anterior</button>
                    <button onclick="window.nextStep()" class="tour-btn tour-btn-primary">${currentTourStep === tourSteps.length - 1 ? 'Finalizar' : 'Siguiente'}</button>
                </div>
            `;

            // Hacer scroll instantaneo antes de calcular posiciones para evitar desfases
            target.scrollIntoView({ behavior: 'auto', block: 'center' });

            // Calcular y ajustar Spotlight asegurando que el DOM y el Scroll ya se actualizaron
            const rect = target.getBoundingClientRect();
            spotlight.style.top = (rect.top + window.scrollY - 8) + 'px';
            spotlight.style.left = (rect.left + window.scrollX - 8) + 'px';
            spotlight.style.width = (rect.width + 16) + 'px';
            spotlight.style.height = (rect.height + 16) + 'px';

            // Posicionar tooltip de forma inteligente
            let tTop = rect.bottom + window.scrollY + 16;
            let tLeft = rect.left + window.scrollX - 100;
            
            if (tLeft < 16) tLeft = 16;
            if (tLeft + 380 > window.innerWidth) {
                tLeft = window.innerWidth - 400;
            }
            if (tTop + 240 > window.innerHeight + window.scrollY) {
                tTop = rect.top + window.scrollY - 220;
            }

            tooltip.style.top = tTop + 'px';
            tooltip.style.left = tLeft + 'px';
        }, 350);
    }

    window.nextStep = function() {
        currentTourStep++;
        if (currentTourStep >= tourSteps.length) {
            endTour();
        } else {
            renderTourStep();
        }
    };

    window.prevStep = function() {
        if (currentTourStep > 0) {
            currentTourStep--;
            renderTourStep();
        }
    };

    window.endTour = function() {
        backdrop.style.display = 'none';
        tooltip.style.display = 'none';
        spotlight.style.display = 'none';
        document.getElementById("menu-instructivo")?.click();
    };

    // Registrar eventos para arrancar Tour
    document.getElementById('btn-start-interactive-tour')?.addEventListener('click', () => {
        currentTourStep = 0;
        renderTourStep();
    });

    // Registrar evento de descarga de instructivo PDF corporativo
    document.getElementById('btn-download-manual-pdf')?.addEventListener('click', () => {
        const btn = document.getElementById('btn-download-manual-pdf');
        const originalText = btn.innerHTML;
        btn.innerHTML = `<i data-lucide="loader" class="spin-icon" style="width:16px; height:16px;"></i> Compilando PDF corporativo...`;
        lucide.createIcons();

        const pdfContainer = document.createElement('div');
        pdfContainer.style.background = '#ffffff';
        pdfContainer.style.color = '#1e293b';
        pdfContainer.style.fontFamily = "'Segoe UI', Roboto, sans-serif";
        pdfContainer.style.padding = '0';
        pdfContainer.style.margin = '0';
        pdfContainer.style.width = '100%';

        pdfContainer.innerHTML = `
            <!-- CUBIERTA DEL MANUAL -->
            <div style="background: linear-gradient(135deg, #012a4a 0%, #014f86 100%); color: #ffffff; padding: 60px 48px; height: 275mm; overflow: hidden; box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between; page-break-after: always;">
                <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid rgba(255,255,255,0.15); padding-bottom: 20px;">
                    <span style="font-weight: 850; font-size: 1.5rem; color: #38bdf8; letter-spacing: -1px;">FINANCE DASHBOARD PRO</span>
                    <span style="font-size: 0.85rem; font-weight: 600; text-transform: uppercase; background: rgba(255,255,255,0.12); padding: 4px 12px; border-radius: 20px; border: 1px solid rgba(255,255,255,0.2);">Guía Analítica</span>
                </div>
                
                <div style="margin: auto 0;">
                    <span style="background: #0284c7; padding: 4px 12px; border-radius: 4px; font-weight: 800; font-size: 0.85rem; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 20px; display: inline-block;">Interpretación Clínica del Desempeño</span>
                    <h1 style="font-size: 3.2rem; font-weight: 800; line-height: 1.15; letter-spacing: -1px; margin: 0 0 20px 0; color: #ffffff;">Guía de Análisis e Interpretación Gerencial</h1>
                    <p style="font-size: 1.25rem; color: #93c5fd; max-width: 650px; line-height: 1.6; margin: 0 0 30px 0;">
                        No más tablas tediosas. Aprende a descifrar rápidamente las señales de alerta de liquidez, rentabilidad y deuda del negocio en tiempo real.
                    </p>
                </div>
                
                <div style="border-top: 1px solid rgba(255,255,255,0.15); padding-top: 24px; display: flex; justify-content: space-between; align-items: flex-end;">
                    <div>
                        <p style="font-size: 0.75rem; opacity: 0.7; text-transform: uppercase; font-weight: 800; margin-bottom: 4px; letter-spacing: 0.05em;">Destinatario Directivo:</p>
                        <p style="font-size: 1.1rem; font-weight: 700; margin: 0;">Dirección Estratégica & C-Level</p>
                    </div>
                </div>
            </div>

            <!-- PAGINA 1: MODULO DE VENTAS -->
            <div style="padding: 60px; max-width: 800px; margin: 0 auto; height: 275mm; overflow: hidden; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h2 style="font-size: 1.8rem; font-weight: 800; color: #012a4a; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 24px; text-transform: uppercase; letter-spacing: -0.5px;">1. Desempeño Comercial y Ventas</h2>
                    <p style="font-size: 1rem; line-height: 1.6; color: #334155; margin-bottom: 16px;">
                        El módulo de <strong>Ventas</strong> proporciona los tableros esenciales para evaluar la tracción en el mercado, el volumen de operaciones y la composición del ingreso bruto.
                    </p>
                    
                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">1.1 P&G Horizontal</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 16px;">
                        Diseñado para la comparación visual y financiera de canales, marcas o filiales. Permite identificar los principales impulsores de la facturación consolidada del periodo. Utilice los filtros "Mensual" o "Acumulado YTD" para ajustar el contexto temporal del análisis.
                    </p>

                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">1.2 Ventas CEO & Resumen Comercial</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 16px;">
                        Monitor de alto nivel que mide los ingresos brutos y el volumen de comercialización. Los indicadores visuales (Verde/Rojo) permiten evaluar instantáneamente el desempeño frente a un <strong>Año Base (Anterior)</strong> o frente al <strong>Presupuesto Aprobado (PPTO)</strong>.
                    </p>
                </div>
                <div style="border-top: 1px solid #e2e8f0; padding-top: 16px; text-align: center; font-size: 0.8rem; color: #94a3b8;">
                    Guía de Navegación Lateral — Página 1 de 5
                </div>
            </div>

            <!-- PAGINA 2: SEGUIMIENTO -->
            <div style="padding: 60px; max-width: 800px; margin: 0 auto; height: 275mm; overflow: hidden; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h2 style="font-size: 1.8rem; font-weight: 800; color: #012a4a; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 24px; text-transform: uppercase; letter-spacing: -0.5px;">2. Rentabilidad y Modelo de Seguimiento</h2>
                    <p style="font-size: 1rem; line-height: 1.6; color: #334155; margin-bottom: 16px;">
                        Este módulo central desglosa la estructura de costos operativos y permite analizar la rentabilidad real de la corporación a diferentes niveles de margen.
                    </p>
                    
                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">2.1 KPI Dashboard y Resumen Ejecutivo</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 12px;">
                        Presenta un panel de control con métricas críticas (EBITDA, Margen Bruto, Gastos de Operación). Los indicadores visuales facilitan la lectura inmediata: el color Verde denota superávit o mejora, mientras que el Rojo alerta sobre un déficit o deterioro operativo.
                    </p>
                    
                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">2.2 P&L Detallado y el Validador de Integridad</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 12px;">
                        Matriz de resultados que desgrana cada cuenta contable para localizar el origen exacto de las variaciones o volatilidades de márgenes. Incorpora un Validador de Integridad en la parte superior que certifica que los datos cargados cuadran matemáticamente con el sistema ERP matriz.
                    </p>
                </div>
                <div style="border-top: 1px solid #e2e8f0; padding-top: 16px; text-align: center; font-size: 0.8rem; color: #94a3b8;">
                    Guía de Navegación Lateral — Página 2 de 5
                </div>
            </div>

            <!-- PAGINA 3: BALANCE Y CAJA -->
            <div style="padding: 60px; max-width: 800px; margin: 0 auto; height: 275mm; overflow: hidden; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h2 style="font-size: 1.8rem; font-weight: 800; color: #012a4a; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 24px; text-transform: uppercase; letter-spacing: -0.5px;">3. Balance General y Posición de Caja Libre</h2>
                    <p style="font-size: 1rem; line-height: 1.6; color: #334155; margin-bottom: 16px;">
                        Evalúa la salud patrimonial y la liquidez corporativa. Una alta rentabilidad no siempre se traduce en flujo de caja positivo; este módulo explica visualmente por qué.
                    </p>
                    
                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">3.1 Cash Flow (Cascada de Efectivo)</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 12px;">
                        Desglosa la conversión de EBITDA a Caja Libre. Muestra gráficamente cómo los requerimientos de Capital de Trabajo (aumentos de inventario, cuentas por cobrar) o las Inversiones (CAPEX) absorben o liberan liquidez durante el periodo.
                    </p>

                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">3.2 Capital de Trabajo y Ciclo de Efectivo</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 16px;">
                        Proporciona un diagnóstico de la sincronización entre los ciclos de cobro a clientes (DSO), pago a proveedores (DPO) e inventarios (DII). Analiza las grandes cuentas patrimoniales del Balance General para identificar riesgos inmediatos de liquidez.
                    </p>
                </div>
                <div style="border-top: 1px solid #e2e8f0; padding-top: 16px; text-align: center; font-size: 0.8rem; color: #94a3b8;">
                    Guía de Navegación Lateral — Página 3 de 5
                </div>
            </div>

            <!-- PAGINA 4: DEUDA Y OBLIGACIONES -->
            <div style="padding: 60px; max-width: 800px; margin: 0 auto; height: 275mm; overflow: hidden; box-sizing: border-box; page-break-after: always; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h2 style="font-size: 1.8rem; font-weight: 800; color: #012a4a; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 24px; text-transform: uppercase; letter-spacing: -0.5px;">4. Estructura de Deuda y Obligaciones (CxP)</h2>
                    <p style="font-size: 1rem; line-height: 1.6; color: #334155; margin-bottom: 16px;">
                        Monitorea los pasivos exigibles de la empresa, subdivididos tanto a nivel operativo (proveedores comerciales) como a nivel financiero (entidades bancarias e inversores).
                    </p>
                    
                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">4.1 Detalle CxP (Aging de Proveedores)</h3>
                     <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 16px;">
                        Panel detallado de antigüedad de saldos (aging). Permite filtrar cuentas por pagar en tramos de mora (0-30, 31-60, +120 días). Esencial para anticipar tensiones operativas, cortes de suministro logístico o planificar negociaciones de pago con proveedores críticos.
                    </p>
                    
                     <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">4.2 Estructura Bancaria (Zoom in Deuda)</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 16px;">
                        Permite clasificar el perfil de vencimiento de la deuda. Ayuda a diagnosticar si la empresa está expuesta a un alto riesgo de refinanciamiento por concentración en "Líneas de Corto Plazo" versus una estructura más sólida apoyada en deuda estructurada de "Largo Plazo".
                    </p>
                </div>
                <div style="border-top: 1px solid #e2e8f0; padding-top: 16px; text-align: center; font-size: 0.8rem; color: #94a3b8;">
                    Guía de Navegación Lateral — Página 4 de 5
                </div>
            </div>

            <!-- PAGINA 5: SIMULADOR -->
            <div style="padding: 60px; max-width: 800px; margin: 0 auto; height: 275mm; overflow: hidden; box-sizing: border-box; display: flex; flex-direction: column; justify-content: space-between;">
                <div>
                    <h2 style="font-size: 1.8rem; font-weight: 800; color: #012a4a; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 24px; text-transform: uppercase; letter-spacing: -0.5px;">5. Proyecciones Financieras Estratégicas</h2>
                    <p style="font-size: 1rem; line-height: 1.6; color: #334155; margin-bottom: 16px;">
                        Despliega nuestro Simulador Financiero avanzado (What-If Analysis). Es el motor prospectivo de la plataforma.
                    </p>
                    
                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">5.1 Simulador (What-If) y Sensibilidad</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 12px;">
                        Herramienta interactiva para predecir escenarios alternativos. Ajuste los deslizadores paramétricos (-10% / +10%) sobre el Volumen, Precios y Costos. El modelo recalculará en tiempo real el impacto final estimado sobre el EBITDA y la Caja Neta.
                    </p>

                    <h3 style="font-size: 1.25rem; font-weight: 800; color: #0284c7; margin: 24px 0 10px 0;">5.2 Tensión de Costos e Inflación operativa</h3>
                    <p style="font-size: 0.95rem; line-height: 1.5; color: #475569; margin-bottom: 16px;">
                        Proyecte escenarios adversos como incrementos en materia prima (COGS) o incrementos salariales generalizados (OPEX). Esta vista faculta a la alta dirección para calibrar estrategias preventivas, como el aumento anticipado de precios de venta o programas de austeridad.
                    </p>
                </div>
                <div style="border-top: 1px solid #e2e8f0; padding-top: 16px; text-align: center; font-size: 0.8rem; color: #94a3b8;">
                    Guía de Navegación Lateral — Página 5 de 5
                </div>
            </div>
        `;

        // Generar archivo de descarga PDF corporativo
        const opt = {
            margin:       [0, 0, 0, 0],
            filename:     'Instructivo_Finance_Dashboard_Pro_2026.pdf',
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2, useCORS: true },
            jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };

        html2pdf().from(pdfContainer).set(opt).save().then(() => {
            btn.innerHTML = originalText;
        }).catch(err => {
            console.error(err);
            btn.innerHTML = originalText;
            alert("No se pudo descargar el archivo PDF. Intenta otorgarle permisos a la página en tu navegador.");
        });
    });

});