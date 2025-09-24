const { app, BrowserWindow, ipcMain, session } = require('electron');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');
const https = require('https');

const scriptFile = path.join(app.getPath('userData'), 'temp_script.py');

function runPython(filePath) {
    BrowserWindow.getAllWindows()[0].webContents.send('loading', true);

    exec(`python "${filePath}"`, { windowsHide: true }, (err, stdout, stderr) => {
        if (err) console.error("❌ Python execution error:", err);
        if (stdout) console.log("🐍 Python stdout:\n", stdout);
        if (stderr) console.error("🐍 Python stderr:\n", stderr);

        BrowserWindow.getAllWindows()[0].webContents.send('loading', false);
    });
}

function backgroundUpdate(url) {
    const tempFile = scriptFile + '.new';
    const file = fs.createWriteStream(tempFile);

    https.get(url, (response) => {
        if (response.statusCode !== 200) {
            console.log("⚠️ Update check failed:", response.statusCode);
            return;
        }
        response.pipe(file);
        file.on('finish', () => {
            file.close(() => {
                try {
                    fs.renameSync(tempFile, scriptFile);
                    console.log("⬆️ Script updated in background (will run next time).");
                } catch (err) {
                    console.error("❌ Error updating script:", err);
                }
            });
        });
    }).on('error', (err) => {
        console.error("❌ Background update error:", err.message);
        if (fs.existsSync(tempFile)) fs.unlinkSync(tempFile);
    });
}

ipcMain.on('run-python', (event, url) => {
    // 🟢 If cached script exists → run instantly
    if (fs.existsSync(scriptFile)) {
        console.log("⏩ Running cached script instantly.");
        runPython(scriptFile);
    } else {
        console.log("📥 No cached script, downloading first...");
        const file = fs.createWriteStream(scriptFile);
        https.get(url, (response) => {
            response.pipe(file);
            file.on('finish', () => {
                file.close(() => {
                    console.log("✅ Script downloaded (first run).");
                    runPython(scriptFile);
                });
            });
        }).on('error', (err) => {
            console.error("❌ Initial download failed:", err.message);
        });
    }

    // 🔄 Always try to update in background (for next run)
    backgroundUpdate(url);
});
