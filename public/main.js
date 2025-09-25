const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const crypto = require('crypto');
const { exec } = require('child_process');

// ----------------- Config -----------------
const BASE_DIR = "C:\\DSQ Enterprise";
// Use global Python
const pythonExe = "python"; // assumes Python is in PATH
const mainUrl = 'https://dsq-beta.vercel.app/index.html';

// ----------------- Helpers -----------------
function hashFile(filePath) {
    if (!fs.existsSync(filePath)) return null;
    return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

// ----------------- Browser Window -----------------
async function createWindow() {
    const ses = session.defaultSession;
    await ses.clearCache();
    await ses.clearStorageData();
    console.log('🗑️ Cache and storage cleared.');

    const win = new BrowserWindow({
        width: 1200,
        height: 900,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            webSecurity: false
        }
    });

    await win.loadURL(mainUrl);
    console.log('✅ index.html loaded from Vercel.');
}

// ----------------- DSQ.py Download & Run -----------------
ipcMain.on('run-python', (event, url) => {
    const filePath = path.join(BASE_DIR, 'DSQ.py');
    const tempPath = filePath + '.new';
    console.log("📥 DSQ.py download request for:", url);

    const file = fs.createWriteStream(tempPath);
    const cacheBustedUrl = url.includes('?') ? `${url}&_=${Date.now()}` : `${url}?_=${Date.now()}`;

    https.get(cacheBustedUrl, (response) => {
        if (response.statusCode !== 200) {
            console.error(`❌ Failed to download DSQ.py: ${response.statusCode}`);
            file.close(); if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
            return;
        }

        response.pipe(file);
        file.on('finish', () => {
            file.close(() => {
                const newHash = crypto.createHash('sha256').update(fs.readFileSync(tempPath)).digest('hex');
                const oldHash = hashFile(filePath);

                if (newHash !== oldHash) {
                    fs.renameSync(tempPath, filePath);
                    console.log("⬆️ DSQ.py updated.");
                } else {
                    fs.unlinkSync(tempPath);
                    console.log("⏩ Using cached DSQ.py");
                }

                // Run global Python
                exec(`"${pythonExe}" "${filePath}"`, { windowsHide: true }, (err, stdout, stderr) => {
                    if (err) console.error("❌ Python execution error:", err);
                    if (stdout) console.log("🐍 Python stdout:\n", stdout);
                    if (stderr) console.error("🐍 Python stderr:\n", stderr);
                });
            });
        });
    }).on('error', (err) => {
        console.error('❌ DSQ.py download error:', err.message);
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
    });
});

// ----------------- App Lifecycle -----------------
app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
