import * as XLSX from "xlsx";
import * as dotenv from "dotenv";

dotenv.config();

// Helper to resolve SharePoint direct download URLs
function resolveSharepointUrl(url: string | undefined, fallback: string): string {
  const current = url || fallback;
  if (current.includes("sharepoint.com") && !current.includes("download=1") && !current.includes("download=true")) {
    if (current.includes("Doc.aspx") || current.includes("onedoc.aspx")) {
      return current + "&download=1";
    }
  }
  return current;
}

async function run() {
  const fallbackUrl = "https://aguaplanetaazul2-my.sharepoint.com/personal/christopher_corona_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={0dded43b-deb4-4017-b8e7-849aa0ca29ac}&download=1";
  const rawUrl = process.env.VITE_CEO_FILE_URL || fallbackUrl;
  const url = resolveSharepointUrl(rawUrl, fallbackUrl);

  console.log("Fetching from URL:", url);

  try {
    const res = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
      }
    });

    if (!res.ok) {
      console.error(`HTTP error: ${res.status} ${res.statusText}`);
      const text = await res.text();
      console.error("Preview of response body:", text.slice(0, 500));
      return;
    }

    const contentType = res.headers.get("content-type") || "";
    console.log("Content-Type:", contentType);

    const arrayBuffer = await res.arrayBuffer();
    const preview = new TextDecoder().decode(arrayBuffer.slice(0, 300));

    if (
      contentType.includes("text/html") ||
      /^\s*<!doctype html/i.test(preview) ||
      /^\s*<html/i.test(preview)
    ) {
      console.error("Error: SharePoint returned HTML/login page instead of Excel. Please confirm the file shared status or URL.");
      console.log("HTML Preview:\n", preview);
      return;
    }

    const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
    console.log("Workbook Sheets:", workbook.SheetNames);

    for (const sheetName of workbook.SheetNames) {
      console.log(`\n================ SHEET: ${sheetName} ================`);
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1 });
      console.log(`Total raw rows: ${rows.length}`);
      
      const nonEmptyRows = rows.filter(r => r && r.length > 0);
      console.log(`First 10 non-empty rows:`);
      nonEmptyRows.slice(0, 10).forEach((r, idx) => {
        console.log(`Row ${idx + 1}:`, r.slice(0, 12));
      });

      // Let's do some metric column scanning
      let matchScore = 0;
      for (const r of rows) {
        if (!r) continue;
        for (const c of r) {
          if (c === undefined || c === null) continue;
          const term = String(c).toLowerCase().trim();
          if (term === 'producto' || term === 'descripción' || term === 'descripcion') matchScore += 10;
          if (term === 'tipo') matchScore += 5;
          if (term.includes('ventas netas') || term.includes('ventas')) matchScore += 8;
          if (term.includes('volumen') || term.includes('unidades')) matchScore += 8;
          if (term === '2026' || term.includes('ppto')) matchScore += 5;
        }
      }
      console.log(`Match score for sheet "${sheetName}": ${matchScore}`);
    }

  } catch (error) {
    console.error("Error running diagnostics:", error);
  }
}

run();
