const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');
const https = require('https');
const crypto = require('crypto');

function hashFile(filePath) {
    if (!fs.existsSync(filePath)) return null;
    return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

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

    const url = 'https://dsq-beta.vercel.app/index.html';
    await win.loadURL(url);
    console.log('✅ index.html loaded from Vercel.');
}

ipcMain.on('run-python', (event, url) => {
    const filePath = path.join(app.getPath('userData'), 'temp_script.py');
    const tempPath = filePath + '.new';

    console.log("📥 Download request for:", url);

    const file = fs.createWriteStream(tempPath);
    const cacheBustedUrl = url.includes('?') ? `${url}&_=${Date.now()}` : `${url}?_=${Date.now()}`;

    https.get(cacheBustedUrl, (response) => {
        if (response.statusCode !== 200) {
            console.error(`❌ Failed to download: ${response.statusCode}`);
            file.close(); fs.unlinkSync(tempPath);
            return;
        }

        response.pipe(file);
        file.on('finish', () => {
            file.close(() => {
                const newHash = crypto.createHash('sha256')
                                      .update(fs.readFileSync(tempPath)).digest('hex');
                const oldHash = hashFile(filePath);

                if (newHash !== oldHash) {
                    // overwrite with new version
                    fs.renameSync(tempPath, filePath);
                    console.log("⬆️ Python script updated.");
                } else {
                    // same as before → discard temp
                    fs.unlinkSync(tempPath);
                    console.log("⏩ No update, using cached Python script.");
                }

                // always run the current version
                exec(`python "${filePath}"`, { windowsHide: true }, (err, stdout, stderr) => {
                    if (err) console.error("❌ Python execution error:", err);
                    if (stdout) console.log("🐍 Python stdout:\n", stdout);
                    if (stderr) console.error("🐍 Python stderr:\n", stderr);
                });
            });
        });
    }).on('error', (err) => {
        console.error('❌ Download error:', err.message);
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
    });
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
