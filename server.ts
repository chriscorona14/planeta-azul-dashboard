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
    // If it's a GUID, format it into the full SharePoint URL
    if (/^[0-9a-fA-F\-]{36}$/.test(inputUrl) || /^\{[0-9a-fA-F\-]{36}\}$/.test(inputUrl)) {
      return `https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc=${inputUrl.startsWith('{') ? inputUrl : '{' + inputUrl + '}'}&download=1`;
    }
    return inputUrl;
  }

  app.get("/api/downloadSync", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_ONEDRIVE_ITEM_ID || process.env.VITE_ONEDRIVE_FILE_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={cfe13828-c964-447a-8147-feb8de79816c}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const response = await fetch(url, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
      });
      
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

  app.get("/api/downloadSyncVentas", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_CEO_FILE_URL || process.env.VITE_ONEDRIVE_VENTAS_ITEM_ID, "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={654321-URL-PLACEHOLDER}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const response = await fetch(url, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
      });
      
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

  app.get("/api/downloadSyncComercial", async (req, res) => {
    try {
      const customUrl = typeof req.query.url === "string" ? req.query.url : undefined;
      const url = resolveSharepointUrl(customUrl || process.env.VITE_RESUMEN_COMERCIAL_URL, "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={PLACEHOLDER-COMERCIAL}&download=1");
      if (!url.includes("sharepoint.com") && !url.includes("onedrive.live.com")) {
        return res.status(400).json({ error: "Invalid Microsoft 365 file URL." });
      }
      const response = await fetch(url, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
      });
      
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
      const response = await fetch(url, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
      });
      
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

  // Vite middleware for development
  app.get("/api/config", (req, res) => {
    res.json({
      VITE_ONEDRIVE_FILE_URL: process.env.VITE_ONEDRIVE_ITEM_ID || process.env.VITE_ONEDRIVE_FILE_URL,
      VITE_CEO_FILE_URL: process.env.VITE_CEO_FILE_URL || process.env.VITE_ONEDRIVE_VENTAS_ITEM_ID,
      VITE_RESUMEN_COMERCIAL_URL: process.env.VITE_RESUMEN_COMERCIAL_URL,
      VITE_PG_HORIZONTAL_URL: process.env.VITE_PG_HORIZONTAL_URL
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

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
