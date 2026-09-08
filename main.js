const { app, BrowserWindow, dialog, powerSaveBlocker } = require("electron");
const fs = require("fs");
const http = require("http");
const path = require("path");

let mainWindow = null;
let serverPort = null;
let serverUrl = "";
let powerSaveBlockerId = null;

app.commandLine.appendSwitch("disable-background-timer-throttling");
app.commandLine.appendSwitch("disable-renderer-backgrounding");
app.commandLine.appendSwitch("disable-backgrounding-occluded-windows");

function startPowerSaveBlocker() {
  if (powerSaveBlockerId !== null && powerSaveBlocker.isStarted(powerSaveBlockerId)) {
    return;
  }
  powerSaveBlockerId = powerSaveBlocker.start("prevent-app-suspension");
}

function waitForServer(url, timeoutMs = 15000) {
  const startedAt = Date.now();
  return new Promise((resolve, reject) => {
    const attempt = () => {
      const req = http.get(`${url.replace(/\/$/, "")}/api/health`, (res) => {
        res.resume();
        if (res.statusCode && res.statusCode >= 200 && res.statusCode < 500) {
          resolve();
        } else {
          retry();
        }
      });

      req.on("error", retry);
      req.setTimeout(1200, () => {
        req.destroy();
        retry();
      });
    };

    const retry = () => {
      if (Date.now() - startedAt > timeoutMs) {
        reject(new Error("Server start timeout."));
        return;
      }
      setTimeout(attempt, 300);
    };

    attempt();
  });
}

function ensureServerStarted() {
  if (serverUrl) {
    return serverUrl;
  }

  const configuredUrl = String(process.env.POS_SERVER_URL || "").trim();
  if (configuredUrl) {
    serverUrl = configuredUrl.replace(/\/$/, "");
    return serverUrl;
  }

  if (serverPort) {
    return `http://127.0.0.1:${serverPort}`;
  }

  const userDataDir = app.getPath("userData");
  const dataDir = path.join(userDataDir, "data");
  const dataFile = path.join(dataDir, "pos-data.xlsx");
  const legacyFile = path.join(userDataDir, "db", "pos-data.xlsx");
  fs.mkdirSync(dataDir, { recursive: true });
  if (!fs.existsSync(dataFile) && fs.existsSync(legacyFile)) {
    fs.copyFileSync(legacyFile, dataFile);
  }
  process.env.POS_DATA_FILE = dataFile;
  process.env.PORT = process.env.PORT || "3210";
  process.env.HOST = process.env.HOST || "0.0.0.0";
  serverPort = Number(process.env.PORT);

  require(path.join(__dirname, "..", "server.js"));
  serverUrl = `http://127.0.0.1:${serverPort}`;
  return serverUrl;
}

async function createWindow() {
  const targetUrl = ensureServerStarted();
  await waitForServer(targetUrl);

  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1100,
    minHeight: 720,
    autoHideMenuBar: true,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      backgroundThrottling: false
    }
  });

  await mainWindow.loadURL(`${targetUrl}/login`);
  mainWindow.on("show", startPowerSaveBlocker);
  mainWindow.on("focus", startPowerSaveBlocker);
  mainWindow.on("closed", () => {
    mainWindow = null;
  });
}

app.whenReady().then(async () => {
  try {
    startPowerSaveBlocker();
    await createWindow();
  } catch (err) {
    dialog.showErrorBox("Startup Error", err.message || "Failed to launch application.");
    app.quit();
  }
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow().catch((err) => {
      dialog.showErrorBox("Startup Error", err.message || "Failed to reopen application.");
      app.quit();
    });
  }
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("before-quit", () => {
  if (powerSaveBlockerId !== null && powerSaveBlocker.isStarted(powerSaveBlockerId)) {
    powerSaveBlocker.stop(powerSaveBlockerId);
  }
  powerSaveBlockerId = null;
});
