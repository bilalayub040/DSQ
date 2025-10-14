const { app, BrowserWindow, ipcMain, session, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const { exec } = require('child_process');

const mainUrl = 'https://dsq-beta.vercel.app/index.html';

async function createWindow() {
    const ses = session.defaultSession;
    await ses.clearCache();
    await ses.clearStorageData();

    const win = new BrowserWindow({
        width: 850,
        height: 900,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            webSecurity: false
        }
    });

    await win.loadURL(mainUrl);
win.setMenu(null);//removes the toolbarr

}

// ---------- File Picker ----------
ipcMain.handle('pick-files', async () => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win, {
        properties: ['openFile', 'multiSelections']
    });
    if (result.canceled) return [];
    return result.filePaths.map(p => ({ path: p, name: path.basename(p) }));
});

// ---------- Silent Download & Run ----------
ipcMain.handle('download-and-run-updater', async (event, url) => {
    try {
        const destDir = 'C:\\DSQ Enterprise\\Updates';
        const filePath = path.join(destDir, 'updater.exe');

        if (!fs.existsSync(destDir)) fs.mkdirSync(destDir, { recursive: true });

        await new Promise((resolve, reject) => {
            const file = fs.createWriteStream(filePath);
            https.get(url, response => {
                if (response.statusCode !== 200) return reject();
                response.pipe(file);
                file.on('finish', () => file.close(resolve));
            }).on('error', reject);
        });

        exec(`"${filePath}"`);
        return { success: true };
    } catch {
        return { success: false };
    }
});

// ---------- App Lifecycle ----------
app.whenReady().then(createWindow);
app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});
app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});







