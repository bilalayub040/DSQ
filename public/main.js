const { app, BrowserWindow, ipcMain, session } = require('electron');
const fs = require('fs');
const path = require('path');
const https = require('https');
const AdmZip = require('adm-zip');

// ----------------- Config -----------------
const mainUrl = 'https://dsq-beta.vercel.app/index.html';
const depsUrl = 'https://yourserver.com/deps.zip'; // replace with your deps.zip URL
const depsPath = path.join(__dirname, 'deps.zip');
const extractFolder = path.join(__dirname, 'deps'); // folder where deps.zip will be extracted

// ----------------- Helper: Download file -----------------
function downloadFile(url, dest) {
    return new Promise((resolve, reject) => {
        const file = fs.createWriteStream(dest);
        https.get(url, (res) => {
            if (res.statusCode !== 200) {
                return reject(new Error(`Failed to get '${url}' (${res.statusCode})`));
            }
            res.pipe(file);
        });

        file.on('finish', () => file.close(resolve));
        file.on('error', (err) => reject(err));
    });
}

// ----------------- Helper: Extract zip with merge -----------------
function unzipMerge(zipFilePath, targetFolder) {
    const zip = new AdmZip(zipFilePath);
    zip.getEntries().forEach((entry) => {
        const entryPath = path.join(targetFolder, entry.entryName);

        if (entry.isDirectory) {
            if (!fs.existsSync(entryPath)) {
                fs.mkdirSync(entryPath, { recursive: true });
            }
        } else {
            const dirName = path.dirname(entryPath);
            if (!fs.existsSync(dirName)) {
                fs.mkdirSync(dirName, { recursive: true });
            }
            fs.writeFileSync(entryPath, entry.getData());
        }
    });

    console.log(`📦 Extracted ${zipFilePath} to ${targetFolder} (merged)`);
}

// ----------------- Browser Window -----------------
async function createWindow() {
    const ses = session.defaultSession;
    await ses.clearCache();
    await ses.clearStorageData();
    console.log('🗑️ Cache and storage cleared.');

    // Download and merge deps.zip
    try {
        console.log('⬇️ Downloading deps.zip...');
        await downloadFile(depsUrl, depsPath);
        console.log('📦 Download complete.');

        console.log('🔀 Extracting and merging deps.zip...');
        unzipMerge(depsPath, extractFolder);
    } catch (err) {
        console.error('❌ Failed to download or extract deps.zip:', err);
    }

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
ipcMain.on('task-started', (event) => event.sender.send('main-busy'));
ipcMain.on('task-finished', (event) => event.sender.send('main-done'));

// ----------------- App Lifecycle -----------------
app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
