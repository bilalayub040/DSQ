//thrth
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

    const url = 'https://dsq-beta.vercel.app/index.html';
    await win.loadURL(url);
    console.log('✅ index.html loaded from Vercel.');
}

ipcMain.on('run-python', (event, url) => {
    const filePath = path.join(app.getPath('userData'), 'DSQ.py');

    if (fs.existsSync(filePath)) {
        console.log("⏩ File already exists, running cached version.");
        runPython(filePath);
        return;
    }

    console.log("📥 File not found, downloading:", url);

    const file = fs.createWriteStream(filePath);
    https.get(url, (response) => {
        if (response.statusCode !== 200) {
            console.error(`❌ Failed to download: ${response.statusCode}`);
            file.close();
            fs.unlinkSync(filePath);
            return;
        }

        response.pipe(file);
        file.on('finish', () => {
            file.close(() => {
                console.log("⬇️ DSQ.py downloaded and saved.");
                runPython(filePath);
            });
        });
    }).on('error', (err) => {
        console.error('❌ Download error:', err.message);
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    });
});

function runPython(filePath) {
    exec(`python "${filePath}"`, { windowsHide: true }, (err, stdout, stderr) => {
        if (err) console.error("❌ Python execution error:", err);
        if (stdout) console.log("🐍 Python stdout:\n", stdout);
        if (stderr) console.error("🐍 Python stderr:\n", stderr);
    });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
