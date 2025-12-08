const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs-extra');

// 设置自定义缓存目录，解决权限问题
const userDataPath = app.getPath('userData');
const customCachePath = path.join(userDataPath, 'customCache');

// 确保自定义缓存目录存在
fs.ensureDirSync(customCachePath);

// 设置应用程序使用自定义缓存目录
app.setPath('cache', customCachePath);
app.setPath('sessionData', path.join(userDataPath, 'sessionData'));
app.setPath('temp', path.join(userDataPath, 'temp'));

// 禁用SSL证书验证（解决网络环境导致的证书错误）
app.commandLine.appendSwitch('ignore-certificate-errors');
app.commandLine.appendSwitch('ignore-certificate-errors-spki-list');
app.commandLine.appendSwitch('allow-insecure-localhost');
app.commandLine.appendSwitch('disable-features', 'OutOfBlinkCors');
app.commandLine.appendSwitch('disable-web-security');

// 禁用GPU加速，解决部分GPU相关的缓存错误
app.disableHardwareAcceleration();

// 禁用沙盒模式，解决部分权限问题
app.enableSandbox = false;
app.commandLine.appendSwitch('no-sandbox');
app.commandLine.appendSwitch('disable-gpu-sandbox');

// 禁用GPU缓存，直接解决GPU cache创建失败问题
app.commandLine.appendSwitch('disable-gpu-cache');
app.commandLine.appendSwitch('disable-shader-cache');
app.commandLine.appendSwitch('disable-offscreen-rendering');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
  });

  mainWindow.loadFile('index.html');

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === null) {
    createWindow();
  }
});

ipcMain.handle('select-directory', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: '选择项目文件夹',
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('save-file', async (event, data, filename) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: filename,
    filters: [{ name: 'Excel Files', extensions: ['xlsx'] }],
  });
  if (!result.canceled && result.filePath) {
    await fs.writeFile(result.filePath, data);
    return result.filePath;
  }
  return null;
});

ipcMain.handle('select-file', async (event, filters, title) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: filters || [{ name: 'All Files', extensions: ['*'] }],
    title: title || '选择文件'
  });
  return result.canceled ? null : result.filePaths[0];
});