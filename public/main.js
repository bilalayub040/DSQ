const { app, BrowserWindow, ipcMain, session, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const { exec } = require('child_process');

const serverBase = 'https://dsq-beta.vercel.app/';
let mainWindow;

async function createWindow() {
    const ses = session.defaultSession;
    await ses.clearCache();
    await ses.clearStorageData();

    mainWindow = new BrowserWindow({
        width: 850,
        height: 900,
        webPreferences: {
            nodeIntegration: true,       // allows ipcRenderer in loaded pages
            contextIsolation: false,
            webSecurity: false
        }
    });

    mainWindow.setMenu(null);
    await mainWindow.loadURL(serverBase + 'index.html'); // initial menu page

    // --- Inject Home Button on every page load ---
    mainWindow.webContents.on('did-finish-load', () => {
        mainWindow.webContents.executeJavaScript(`
            if (!document.getElementById('homeBtn')) {
                const btn = document.createElement('button');
                btn.id = 'homeBtn';
                btn.textContent = 'Home';
                btn.style.position = 'fixed';
                btn.style.top = '10px';
                btn.style.right = '10px';
                btn.style.padding = '6px 12px';
                btn.style.zIndex = 9999;
                btn.style.background = '#007bff';
                btn.style.color = 'white';
                btn.style.border = 'none';
                btn.style.borderRadius = '4px';
                btn.style.cursor = 'pointer';
                btn.style.fontSize = '14px';

                btn.addEventListener('click', () => {
                    window.location.href = '${serverBase}index.html';
                });

                document.body.appendChild(btn);
            }
        `);
    });
}

// ---------- Load different page on button click ----------
ipcMain.on('load-page', async (event, page) => {
    if (!mainWindow) return;
    console.log('[main] loading page:', page);
    await mainWindow.loadURL(serverBase + encodeURIComponent(page));
});

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
