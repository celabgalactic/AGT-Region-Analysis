import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Support direct local serving of NMSFuturaProBook_Kerned.ttf, geonms-font.ttf (fallback), and AGTicon.png
  app.get("/NMSFuturaProBook_Kerned.ttf", (req, res) => {
    res.sendFile(path.join(process.cwd(), "src", "NMSFuturaProBook_Kerned.ttf"));
  });
  app.get("/geonms-font.ttf", (req, res) => {
    res.sendFile(path.join(process.cwd(), "src", "NMSFuturaProBook_Kerned.ttf"));
  });
  app.get("/AGTicon.png", (req, res) => {
    res.sendFile(path.join(process.cwd(), "AGTicon.png"));
  });
  app.get("/AGTIcon.png", (req, res) => {
    res.sendFile(path.join(process.cwd(), "AGTicon.png"));
  });

  // Proxy route for Google Drive assets to bypass CORS
  app.get("/api/asset-proxy", async (req, res) => {
    const fileId = req.query.id as string;
    if (!fileId) return res.status(400).send("Missing ID");

    try {
      // Try Direct Download link first
      const driveUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
      const response = await fetch(driveUrl, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': '*/*',
          'Accept-Language': 'en-US,en;q=0.9',
        },
        redirect: 'follow'
      });
      
      if (!response.ok) {
        // Log details if possible (visible in container logs)
        console.error(`Fetch failed for ${fileId}: ${response.status} ${response.statusText}`);
        throw new Error(`Failed to fetch from Drive: ${response.status}`);
      }

      const contentType = response.headers.get("content-type");
      
      // If we got an HTML page, it might be a "too large to scan" or login page
      if (contentType && contentType.includes("text/html")) {
        console.warn(`Proxy received HTML for ${fileId}, likely Google Drive warning page.`);
        // Note: For truly large files, we'd need to extract a confirm token from the HTML.
        // For fonts/icons, this shouldn't happen unless permissions are wrong.
      }

      if (contentType) res.setHeader("Content-Type", contentType);
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.setHeader("Cache-Control", "public, max-age=31536000");

      const arrayBuffer = await response.arrayBuffer();
      res.send(Buffer.from(arrayBuffer));
    } catch (error) {
      console.error("Proxy error:", error);
      res.status(500).send("Error proxying asset");
    }
  });

  // Vite middleware for development
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
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
