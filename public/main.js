const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');
const https = require('https');

const etagFile = path.join(app.getPath('userData'), 'temp_script.etag');
const scriptFile = path.join(app.getPath('userData'), 'temp_script.py');

function getSavedETag() {
    return fs.existsSync(etagFile) ? fs.readFileSync(etagFile, 'utf8') : null;
}
function saveETag(etag) {
    fs.writeFileSync(etagFile, etag, 'utf8');
}

function runPython(filePath) {
    BrowserWindow.getAllWindows()[0].webContents.send('loading', true);

    exec(`python "${filePath}"`, { windowsHide: true }, (err, stdout, stderr) => {
        if (err) console.error("❌ Python execution error:", err);
        if (stdout) console.log("🐍 Python stdout:\n", stdout);
        if (stderr) console.error("🐍 Python stderr:\n", stderr);

        BrowserWindow.getAllWindows()[0].webContents.send('loading', false);
    });
}

ipcMain.on('run-python', (event, url) => {
    const cacheBustedUrl = url.includes('?') ? `${url}&_chk=${Date.now()}` : `${url}?_chk=${Date.now()}`;
    const options = new URL(cacheBustedUrl);
    options.method = 'HEAD'; // just metadata, no download

    console.log("🔎 Checking for updates...");

    const req = https.request(options, (res) => {
        const remoteETag = res.headers['etag'] || res.headers['last-modified'];
        const savedETag = getSavedETag();

        if (remoteETag && savedETag && remoteETag === savedETag && fs.existsSync(scriptFile)) {
            console.log("✅ No update found, running cached file.");
            runPython(scriptFile);
        } else {
            console.log("⬆️ Update detected, downloading new script...");
            const file = fs.createWriteStream(scriptFile);
            https.get(url, (response) => {
                response.pipe(file);
                file.on('finish', () => {
                    file.close(() => {
                        if (remoteETag) saveETag(remoteETag);
                        console.log("✅ Script updated & saved.");
                        runPython(scriptFile);
                    });
                });
            }).on('error', (err) => {
                console.error("❌ Download error:", err.message);
                if (fs.existsSync(scriptFile)) {
                    console.log("⚠️ Falling back to cached script.");
                    runPython(scriptFile);
                }
            });
        }
    });

    req.on('error', (err) => {
        console.error("❌ HEAD request error:", err.message);
        if (fs.existsSync(scriptFile)) {
            console.log("⚠️ Falling back to cached script.");
            runPython(scriptFile);
        }
    });

    req.end();
});
