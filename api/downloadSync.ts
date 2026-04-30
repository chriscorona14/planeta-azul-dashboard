export default async function handler(req: any, res: any) {
  try {
    const url = process.env.VITE_ONEDRIVE_FILE_URL || "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={cfe13828-c964-447a-8147-feb8de79816c}&download=1";
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
  } catch (e: any) {
    console.error(e);
    res.status(500).json({ error: e instanceof Error ? e.message : String(e) });
  }
}
