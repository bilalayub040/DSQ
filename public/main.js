const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const { exec, execSync } = require('child_process');
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

function ensurePythonInstalled(callback) {
    try {
        // check if python is in PATH
        execSync('python --version');
        console.log('🐍 Python is already installed.');
        callback();
    } catch {
        console.log('⚠️ Python not found, downloading installer...');
        const installerUrl = 'https://www.python.org/ftp/python/3.12.2/python-3.12.2-amd64.exe';
        const installerPath = path.join(app.getPath('userData'), 'python_installer.exe');
        const file = fs.createWriteStream(installerPath);

        https.get(installerUrl, (response) => {
            if (response.statusCode !== 200) {
                console.error(`❌ Failed to download Python: ${response.statusCode}`);
                file.close(); fs.unlinkSync(installerPath);
                return;
            }

            response.pipe(file);
            file.on('finish', () => {
                file.close(() => {
                    console.log('✅ Python installer downloaded. Installing in background...');
                    // silent install
                    exec(`"${installerPath}" /quiet InstallAllUsers=1 PrependPath=1`, (err) => {
                        if (err) {
                            console.error('❌ Python installation failed:', err);
                            return;
                        }
                        console.log('✅ Python installed successfully.');
                        callback();
                    });
                });
            });
        }).on('error', (err) => {
            console.error('❌ Download error:', err.message);
            if (fs.existsSync(installerPath)) fs.unlinkSync(installerPath);
        });
    }
}

ipcMain.on('run-python', (event, url) => {
    ensurePythonInstalled(() => {
        const filePath = path.join(app.getPath('userData'), 'DSQ.py');
        const tempPath = filePath + '.new';
        console.log("📥 DSQ.py download request for:", url);

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
                        fs.renameSync(tempPath, filePath);
                        console.log("⬆️ DSQ.py updated.");
                    } else {
                        fs.unlinkSync(tempPath);
                        console.log("⏩ Using cached DSQ.py");
                    }

                    // run Python script
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
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
