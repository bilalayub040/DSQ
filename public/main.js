const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const crypto = require('crypto');
const { exec } = require('child_process');

const BASE_DIR = "C:\\DSQ Enterprise";
const mainUrl = 'https://dsq-beta.vercel.app/index.html';

function hashFile(filePath) {
    if (!fs.existsSync(filePath)) return null;
    return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

async function createWindow() {
    const ses = session.defaultSession;
    await ses.clearCache();
    await ses.clearStorageData();

    const win = new BrowserWindow({
        width: 1200,
        height: 900,
        webPreferences: { nodeIntegration: true, contextIsolation: false, webSecurity: false }
    });

    await win.loadURL(mainUrl);
}

// ----------------- Run Python/EXE -----------------
ipcMain.on('run-python', (event, url) => {
    const filePath = path.join(BASE_DIR, 'subs.exe');
    const tempPath = filePath + '.new';

    const file = fs.createWriteStream(tempPath);
    const cacheBustedUrl = url.includes('?') ? `${url}&_=${Date.now()}` : `${url}?_=${Date.now()}`;

    https.get(cacheBustedUrl, (response) => {
        if (response.statusCode !== 200) {
            console.error(`❌ Failed to download: ${response.statusCode}`);
            file.close(); if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
            event.sender.send('process-finished'); // notify renderer even on error
            return;
        }

        response.pipe(file);
        file.on('finish', () => {
            file.close(() => {
                const newHash = crypto.createHash('sha256').update(fs.readFileSync(tempPath)).digest('hex');
                const oldHash = hashFile(filePath);

                if (newHash !== oldHash) {
                    fs.renameSync(tempPath, filePath);
                } else { fs.unlinkSync(tempPath); }

                // Run EXE
                const child = exec(`"${filePath}"`, { windowsHide: true });

                child.stdout.on('data', (data) => {
                    // Optionally print logs
                    console.log(data.toString());
                    if (data.toString().includes('APP_READY')) {
                        event.sender.send('process-finished'); // hide spinner when app is ready
                    }
                });

                child.stderr.on('data', (data) => console.error(data.toString()));
                child.on('exit', () => event.sender.send('process-finished'));
            });
        });
    }).on('error', (err) => {
        console.error('❌ Download error:', err.message);
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
        event.sender.send('process-finished'); // notify renderer
    });
});

// ----------------- App Lifecycle -----------------
app.whenReady().then(createWindow);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
