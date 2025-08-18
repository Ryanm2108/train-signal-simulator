const { app, BrowserWindow } = require("electron");
const path = require("path");
const isDev = process.env.ELECTRON_START_URL;

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
    },
  });

  if (isDev) {
    // Development: load React dev server
    win.loadURL("http://localhost:3000");
  } else {
    // Production: load the built React files
    win.loadFile(path.join(__dirname, "build", "index.html"));
  }
}

app.on("ready", createWindow);
