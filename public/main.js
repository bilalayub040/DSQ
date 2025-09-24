const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const crypto = require('crypto');
const unzipper = require('unzipper');
const { exec } = require('child_process');

// ----------------- Config -----------------
const BASE_DIR = "C:\\DSQ Enterprise";
const pythonDir = path.join(BASE_DIR, 'python');
const pythonExe = path.join(pythonDir, 'python.exe');
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

// ----------------- Portable Python -----------------
function ensurePythonInstalled(callback) {
    if (fs.existsSync(pythonExe)) {
        console.log('🐍 Portable Python already exists.');
        callback();
        return;
    }

    console.log('⚠️ Python not found, downloading portable version...');
    const zipUrl = 'https://www.python.org/ftp/python/3.12.2/python-3.12.2-embed-amd64.zip';
    const zipPath = path.join(BASE_DIR, 'python_embed.zip');
    const file = fs.createWriteStream(zipPath);

    https.get(zipUrl, (res) => {
        if (res.statusCode !== 200) {
            console.error('❌ Failed to download Python zip:', res.statusCode);
            return;
        }

        res.pipe(file);
        file.on('finish', () => {
            file.close(() => {
                console.log('✅ Python zip downloaded. Extracting...');
                fs.createReadStream(zipPath)
                    .pipe(unzipper.Extract({ path: pythonDir }))
                    .on('close', () => {
                        console.log('✅ Python portable ready.');
                        fs.unlinkSync(zipPath);
                        callback();
                    });
            });
        });
    }).on('error', (err) => {
        console.error('❌ Python download error:', err.message);
        if (fs.existsSync(zipPath)) fs.unlinkSync(zipPath);
    });
}

// ----------------- DSQ.py Download & Run -----------------
ipcMain.on('run-python', (event, url) => {
    ensurePythonInstalled(() => {
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

                    if (!fs.existsSync(pythonExe)) {
                        console.error("❌ Portable Python not found at", pythonExe);
                        return;
                    }

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
});

// ----------------- App Lifecycle -----------------
app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
