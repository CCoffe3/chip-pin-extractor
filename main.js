const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1400,
        height: 900,
        minWidth: 1024,
        minHeight: 768,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
            enableRemoteModule: false,
        },
        icon: path.join(__dirname, 'icon.ico'),
    });

    mainWindow.loadFile('index.html');

    // 开发模式下打开开发者工具
    if (process.argv.includes('--dev')) {
        mainWindow.webContents.openDevTools();
    }

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// 处理文件选择对话框
ipcMain.handle('select-files', async (event, options) => {
    const { dialog } = require('electron');
    const result = await dialog.showOpenDialog(mainWindow, options);
    return result;
});

// 处理保存文件对话框
ipcMain.handle('save-file', async (event, options) => {
    const { dialog } = require('electron');
    const result = await dialog.showSaveDialog(mainWindow, options);
    return result;
});
