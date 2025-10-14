const { app, BrowserWindow, ipcMain, session, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const { spawn, exec } = require('child_process');

const serverBase = 'https://dsq-beta.vercel.app/';
const userFile = 'C:\\DSQ Enterprise\\assets\\USER.txt';
const loadExe = 'C:\\DSQ Enterprise\\assets\\load.exe';

let mainWindow;

// --------- Utility Functions ----------
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function waitForUserFile(timeoutMs = 15000) {
    let waited = 0;
    while (!fs.existsSync(userFile) && waited < timeoutMs) {
        await sleep(500);
        waited += 500;
    }
    if (!fs.existsSync(userFile)) throw new Error('USER.txt not generated');
}

function getValidEmail() {
    if (!fs.existsSync(userFile)) return null;
    const lines = fs.readFileSync(userFile, 'utf-8').split(/\r?\n/).filter(Boolean);
    for (const line of lines) {
        if (line.includes('@dsq.qa') || line.includes('@vodafone.qa')) return line;
    }
    return null;
}

async function runLoadExe() {
    return new Promise((resolve, reject) => {
        if (!fs.existsSync(loadExe)) return reject(new Error('load.exe not found'));
        const child = spawn(loadExe, { windowsHide: true });
        child.on('error', reject);
        child.on('exit', (code) => {
            if (code === 0) resolve();
            else reject(new Error(`load.exe exited with code ${code}`));
        });
    });
}

// --------- Create Main Window ----------
async function createWindow() {
    const ses = session.defaultSession;
    await ses.clearCache();
    await ses.clearStorageData();

    mainWindow = new BrowserWindow({
        width: 850,
        height: 900,
		icon: 'C:\\DSQ Enterprise\\icon.ico',  // <-- Add this line

        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            webSecurity: false
        }
    });

    mainWindow.setMenu(null);

    // Animated validating overlay
    const overlayHtml = `
    <html>
    <head>
        <style>
            body { margin:0; display:flex; justify-content:center; align-items:center; height:100vh; background:#f9f9f9; font-family:Arial,sans-serif; }
            .overlay-container { text-align:center; }
            .spinner {
                margin: 0 auto 20px;
                width: 50px; height: 50px;
                border: 5px solid #ccc;
                border-top-color: #007bff;
                border-radius: 50%;
                animation: spin 1s linear infinite;
            }
            @keyframes spin { 100% { transform: rotate(360deg); } }
            .text { font-size:18px; color:#333; font-weight:bold; }
        </style>
    </head>
    <body>
        <div class="overlay-container">
            <div class="spinner"></div>
            <div class="text">Validating User...</div>
        </div>
    </body>
    </html>
    `;
    mainWindow.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(overlayHtml));

    // ---------- Validation Process ----------
    try {
        // Delete USER.txt
        if (fs.existsSync(userFile)) fs.unlinkSync(userFile);

        // Run load.exe
        await runLoadExe();

        // Wait for USER.txt
        await waitForUserFile();

        // Check valid email
        const email = getValidEmail();
        if (!email) throw new Error('Not Valid User!');

        // Load actual UI after validation
        await mainWindow.loadURL(serverBase + 'index.html');

        // Inject Home button & user email
        mainWindow.webContents.on('did-finish-load', () => {
            mainWindow.webContents.executeJavaScript(`
                // Home button
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

                // Email display at top center
                if (!document.getElementById('userEmail')) {
                    const span = document.createElement('span');
                    span.id = 'userEmail';
                    span.textContent = 'User: ${email}';
                    span.style.position = 'fixed';
                    span.style.top = '10px';
                    span.style.left = '50%';
                    span.style.transform = 'translateX(-50%)';
                    span.style.fontSize = '14px';
                    span.style.fontWeight = 'bold';
                    span.style.color = '#333';
                    span.style.zIndex = 9999;
                    document.body.appendChild(span);
                }
            `);
        });

    } catch (err) {
        const msg = err.message || 'Validation Failed';
        await mainWindow.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(`
            <html>
                <body style="display:flex;justify-content:center;align-items:center;height:100vh;background:#f9f9f9;">
                    <h2 style="font-family:Arial,sans-serif;color:red;">${msg}</h2>
                </body>
            </html>
        `));
        setTimeout(() => app.quit(), 2500);
    }
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
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
