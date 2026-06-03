const express = require("express");
const https = require("https");
const path = require("path");
const devCerts = require("office-addin-dev-certs");

async function startServer() {
  const app = express();
  const port = 3000;

  // Serve static files from root directory
  app.use(express.static(path.join(__dirname)));

  // Redirect root to taskpane.html
  app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "src", "taskpane", "taskpane.html"));
  });

  try {
    // Generate/retrieve localhost SSL certificates
    const sslOptions = await devCerts.getHttpsServerOptions();
    const server = https.createServer(sslOptions, app);

    server.listen(port, () => {
      console.log(`\n======================================================`);
      console.log(`[Server] Dev server running at https://localhost:${port}`);
      console.log(`[Server] Taskpane URL: https://localhost:${port}/src/taskpane/taskpane.html`);
      console.log(`======================================================\n`);
    });
  } catch (error) {
    console.error("[Server] Failed to start HTTPS dev server:", error);
    process.exit(1);
  }
}

startServer();
