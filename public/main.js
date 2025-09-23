const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');
const https = require('https');

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

    // Intercept requests for cache busting
    ses.webRequest.onBeforeRequest((details, callback) => {
        let url = details.url;
        if (url.startsWith('http') && !url.includes('_=')) {
            if (url.indexOf('?') > -1) {
                url += `&_=${Date.now()}`;
            } else {
                url += `?_=${Date.now()}`;
            }
            callback({ redirectURL: url });
        } else {
            callback({});
        }
    });

    const url = 'https://dsq-beta.vercel.app/index.html';
    await win.loadURL(url);
    console.log('✅ index.html loaded from Vercel.');

    win.webContents.on('did-finish-load', () => {
        console.log('✅ Renderer finished loading.');
    });
}

ipcMain.on('run-python', (event, url) => {
    const filePath = path.join(app.getPath('userData'), 'temp_script.py');

    console.log("📥 Download request for:", url);
    console.log("📂 Saving to:", filePath);

    // 🗑️ Delete old file if it exists
    if (fs.existsSync(filePath)) {
        try {
            fs.unlinkSync(filePath);
            console.log('🗑️ Old Python file removed.');
        } catch (err) {
            console.error('❌ Error removing old Python file:', err);
        }
    }

    // Add cache-buster to ensure fresh download from Vercel
    const cacheBustedUrl = url.includes('?') ? `${url}&_=${Date.now()}` : `${url}?_=${Date.now()}`;

    const file = fs.createWriteStream(filePath);
    https.get(cacheBustedUrl, (response) => {
        console.log("🌐 HTTP Status:", response.statusCode, response.statusMessage);
        console.log("📏 Content-Type:", response.headers['content-type']);

        response.pipe(file);
        file.on('finish', () => {
            file.close(() => {
                console.log('✅ Python script downloaded.');
                exec(`python "${filePath}"`, (err, stdout, stderr) => {
                    if (err) console.error("❌ Python execution error:", err);
                    if (stdout) console.log("🐍 Python stdout:\n", stdout);
                    if (stderr) console.error("🐍 Python stderr:\n", stderr);
                });
            });
        });
    }).on('error', (err) => {
        fs.unlink(filePath, () => {});
        console.error('❌ Download error:', err.message);
    });
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
