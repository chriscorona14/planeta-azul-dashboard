import "dotenv/config";
import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";

async function startServer() {
  const app = express();
  const PORT = 3000;

  // API routes FIRST
  app.use(express.json({ limit: '10mb' }));

  app.post("/api/gemini/insights", async (req, res) => {
    try {
      const { financialData } = req.body;
      if (!financialData) {
         return res.status(400).json({ error: "Missing financialData" });
      }

      const { GoogleGenAI } = require("@google/genai");
      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey) {
         return res.status(500).json({ error: "GEMINI_API_KEY environment variable is missing" });
      }

      const ai = new GoogleGenAI({
         apiKey: apiKey,
         httpOptions: { headers: { 'User-Agent': 'aistudio-build' } }
      });

      const prompt = `You are a Senior CFO analyzing financial data for a company. 
I will provide you with the most recent financial data summary.
Please provide a concise Executive Summary highlighting:
1. Key performance indicators and their significance.
2. Positive and negative trends.
3. Potential risks or opportunities shown in the data.

Please format your response in Spanish, using simple HTML (e.g. <b>, <ul>, <li>, <br>) without a markdown wrapper.
Keep the overall summary under 300 words and be highly analytical and direct.

Data: ${JSON.stringify(financialData)}`;

      const response = await ai.models.generateContent({
        model: "gemini-3.1-pro-preview",
        contents: prompt,
      });

      res.json({ insight: response.text });
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  function resolveSharepointUrl(inputUrl: string | undefined, defaultUrl: string): string {
    if (!inputUrl) return defaultUrl;
    
    let resolved = String(inputUrl).trim().replace(/&amp;/g, "&");
    
    // Clean braces of the guid if needed for check
    const cleanInput = resolved.replace(/^\{|\}$/g, "");
    if (/^[0-9a-fA-F\-]{36}$/.test(cleanInput)) {
      const personalMatch = defaultUrl.match(/personal\/([^\/]+)/i);
      const personalPath = personalMatch ? personalMatch[1] : "marcos_ojeda_planetaazulrd_com";
      return `https://aguaplanetaazul2-my.sharepoint.com/personal/${personalPath}/_layouts/15/Doc.aspx?sourcedoc={${cleanInput}}&download=1`;
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
  }

  app.get("/api/downloadSync", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_ONEDRIVE_ITEM_ID || process.env.VITE_ONEDRIVE_FILE_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={cfe13828-c964-447a-8147-feb8de79816c}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const fetchHeaders: any = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      };
      if (req.headers.authorization) {
          fetchHeaders["Authorization"] = req.headers.authorization;
      }
      const response = await fetch(url, { headers: fetchHeaders });
      
      if (!response.ok) {
        return res.status(response.status).json({ error: `SharePoint rejected the request: ${response.status} ${response.statusText}. Ensure the file is shared publicly.` });
      }
      
      const contentType = response.headers.get("content-type") || "";
      const buffer = await response.arrayBuffer();
      const preview = new TextDecoder().decode(buffer.slice(0, 300));

      if (
        contentType.includes("text/html") ||
        /^\s*<!doctype html/i.test(preview) ||
        /^\s*<html/i.test(preview)
      ) {
        return res.status(403).json({ error: "Finanzas Master: SharePoint devolvió HTML/login en vez del archivo Excel. Revisar permisos o URL pública." });
      }
      
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.send(Buffer.from(buffer));
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  app.get("/api/downloadSyncVentas", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_CEO_FILE_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/christopher_corona_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={0dded43b-deb4-4017-b8e7-849aa0ca29ac}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const fetchHeaders: any = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      };
      if (req.headers.authorization) {
          fetchHeaders["Authorization"] = req.headers.authorization;
      }
      const response = await fetch(url, { headers: fetchHeaders });
      
      if (!response.ok) {
        return res.status(response.status).json({ error: `SharePoint rejected the request: ${response.status} ${response.statusText}. Ensure the file is shared publicly.` });
      }
      
      const contentType = response.headers.get("content-type") || "";
      const buffer = await response.arrayBuffer();
      const preview = new TextDecoder().decode(buffer.slice(0, 300));

      if (
        contentType.includes("text/html") ||
        /^\s*<!doctype html/i.test(preview) ||
        /^\s*<html/i.test(preview)
      ) {
        return res.status(403).json({ error: "Ventas CEO: SharePoint devolvió HTML/login en vez del archivo Excel. Revisar permisos o URL pública." });
      }
      
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.send(Buffer.from(buffer));
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  app.get("/api/downloadSyncComercial", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_RESUMEN_COMERCIAL_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={PLACEHOLDER-COMERCIAL}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const fetchHeaders: any = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      };
      if (req.headers.authorization) {
          fetchHeaders["Authorization"] = req.headers.authorization;
      }
      const response = await fetch(url, { headers: fetchHeaders });
      
      if (!response.ok) {
        return res.status(response.status).json({ error: `SharePoint rejected the request: ${response.status} ${response.statusText}. Ensure the file is shared publicly.` });
      }
      
      const buffer = await response.arrayBuffer();
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.send(Buffer.from(buffer));
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  app.get("/api/downloadSyncPgHorizontal", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_PG_HORIZONTAL_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={PLACEHOLDER-PG}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const fetchHeaders: any = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      };
      if (req.headers.authorization) {
          fetchHeaders["Authorization"] = req.headers.authorization;
      }
      const response = await fetch(url, { headers: fetchHeaders });
      
      if (!response.ok) {
        return res.status(response.status).json({ error: `SharePoint rejected the request: ${response.status} ${response.statusText}. Ensure the file is shared publicly.` });
      }
      
      const buffer = await response.arrayBuffer();
      
      // TEST: Logging sheet names
      try {
          const XLSX = require('xlsx');
          const workbook = XLSX.read(buffer, { type: 'buffer' });
          console.log("PG Horizontal Sheet Names:", workbook.SheetNames);
          let pgSheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('analítico pyg') || n.toLowerCase().includes('analitico pyg')) || workbook.SheetNames[0];
          console.log("Found PG Sheet:", pgSheetName);
          const data = XLSX.utils.sheet_to_json(workbook.Sheets[pgSheetName], {header: 1, defval: null});
          console.log("First 10 rows:");
          console.dir(data.slice(0, 10), { depth: null });
      } catch (e) {
          console.error("XLSX test failed", e);
      }

      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.send(Buffer.from(buffer));
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  app.get("/api/downloadSyncCxp", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_CXP_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={da78e2c9-ceb1-4f4a-8752-9b1927700779}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const fetchHeaders: any = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      };
      if (req.headers.authorization) {
          fetchHeaders["Authorization"] = req.headers.authorization;
      }
      const response = await fetch(url, { headers: fetchHeaders });
      
      if (!response.ok) {
        return res.status(response.status).json({ error: `SharePoint rejected the request: ${response.status} ${response.statusText}. Ensure the file is shared publicly.` });
      }
      
      const buffer = await response.arrayBuffer();
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.send(Buffer.from(buffer));
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  // Vite middleware for development
  app.get("/api/config", (req, res) => {
    res.json({
      VITE_ONEDRIVE_FILE_URL: process.env.VITE_ONEDRIVE_ITEM_ID || process.env.VITE_ONEDRIVE_FILE_URL,
      VITE_CEO_FILE_URL: process.env.VITE_CEO_FILE_URL,
      VITE_RESUMEN_COMERCIAL_URL: process.env.VITE_RESUMEN_COMERCIAL_URL,
      VITE_PG_HORIZONTAL_URL: process.env.VITE_PG_HORIZONTAL_URL,
      VITE_CXP_URL: process.env.VITE_CXP_URL
    });
  });

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
        if (req.path.startsWith('/api/')) {
            return res.status(404).json({ error: "Not Found" });
        }
        res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.get("/api/downloadSyncCostoUnitario", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_COSTO_UNITARIO_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/christopher_corona_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={738547b0-8a34-4527-bbf0-1c4e9e12075a}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True&edaebf=rslc0&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const fetchHeaders: any = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      };
      if (req.headers.authorization) {
          fetchHeaders["Authorization"] = req.headers.authorization;
      }
      const response = await fetch(url, { headers: fetchHeaders });
      
      if (!response.ok) {
        return res.status(response.status).json({ error: `SharePoint rejected the request: ${response.status} ${response.statusText}. Ensure the file is shared publicly.` });
      }
      
      const buffer = await response.arrayBuffer();
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.send(Buffer.from(buffer));
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
    }
  });

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
