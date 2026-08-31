'use strict';
const { app, BrowserWindow, Menu } = require('electron');

Menu.setApplicationMenu(null);   // File/Edit/View 메뉴바는 이 앱엔 쓸모가 없다

function createWindow() {
  const win = new BrowserWindow({
    width: 1100,
    height: 800,
    webPreferences: { contextIsolation: true, nodeIntegration: false },
  });
  win.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});
