const { app, BrowserWindow, ipcMain, session, dialog } = require('electron');
const path = require('path');

// ----------------- Config -----------------
const mainUrl = 'https://dsq-beta.vercel.app/index.html';

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

// ----------------- IPC: Spinner Coordination -----------------
ipcMain.on('task-started', (event) => {
    // notify renderer that main is busy
    event.sender.send('main-busy');
});

ipcMain.on('task-finished', (event) => {
    // notify renderer that main is done
    event.sender.send('main-done');
});

// ----------------- IPC: File Picker -----------------
ipcMain.handle('pick-files', async (event) => {
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win, {
        title: 'Select attachments',
        properties: ['openFile', 'multiSelections']
    });

    if (result.canceled) return [];
    return result.filePaths.map(p => ({
        path: p,
        name: path.basename(p)
    }));
});

// ----------------- App Lifecycle -----------------
app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
