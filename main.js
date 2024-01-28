const { app, BrowserWindow, ipcMain } = require('electron');
const { spawn, exec } = require('child_process');
const path = require('path');

let mainWindow;
let sidecarProcess = null;
let funtoolProcess = null;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        },
        autoHideMenuBar: true  // 自动隐藏菜单栏
    });
    // 当 Electron 应用准备好后，检查 WeChat 程序
    checkWeChat();

    mainWindow.loadFile('index.html');

    ipcMain.on('start-funtool', () => {
        const timeutc = new Date().toLocaleString('zh-CN', { hour12: false }).replace(/\//g, '-').replace(/:/g, '-').replace(/ /g, '-');
        if (!funtoolProcess) {
            const execPath = path.join(__dirname, 'assets', 'funtool_wx=3.9.2.23.exe');
            funtoolProcess = spawn(execPath);
            funtoolProcess.on('spawn', () => {
                mainWindow.webContents.send('action-result', timeutc + ':funtool已启动！');
            });
            funtoolProcess.on('error', (err) => {
                mainWindow.webContents.send('action-result', timeutc + ':启动funtool时发生错误: ' + err);
            });
        } else {
            mainWindow.webContents.send('action-result', timeutc + ':funtool已在运行中...');
        }
    });

    ipcMain.on('stop-funtool', () => {
        const timeutc = new Date().toLocaleString('zh-CN', { hour12: false }).replace(/\//g, '-').replace(/:/g, '-').replace(/ /g, '-');

        if (funtoolProcess) {
            funtoolProcess.kill();
            funtoolProcess = null;
            mainWindow.webContents.send('action-result', timeutc + ':funtool已停止。');
        } else {
            mainWindow.webContents.send('action-result', timeutc + ':funtool未在运行...');
        }
    });

    ipcMain.on('start-sidecar', () => {
        const timeutc = new Date().toLocaleString('zh-CN', { hour12: false }).replace(/\//g, '-').replace(/:/g, '-').replace(/ /g, '-');

        if (!sidecarProcess) {
            const execPath = path.join(__dirname, 'assets','wxbot-sidecar.exe');
            sidecarProcess = spawn(execPath);
            sidecarProcess.on('spawn', () => {
                mainWindow.webContents.send('action-result', timeutc + ':Sidecar已启动！');
            });
            sidecarProcess.on('error', (err) => {
                mainWindow.webContents.send('action-result', timeutc + ':启动Sidecar时发生错误: ' + err);
            });
        } else {
            mainWindow.webContents.send('action-result', timeutc + 'Sidecar已在运行中...');
        }
    });

    ipcMain.on('stop-sidecar', () => {
        const timeutc = new Date().toLocaleString('zh-CN', { hour12: false }).replace(/\//g, '-').replace(/:/g, '-').replace(/ /g, '-');

        if (sidecarProcess) {
            sidecarProcess.kill();
            sidecarProcess = null;
            mainWindow.webContents.send('action-result', timeutc + ':Sidecar已停止!');
        } else {
            mainWindow.webContents.send('action-result', timeutc + ':Sidecar未在运行...');
        }
    });

}

function checkWeChat() {
    // 检查 WeChat 是否启动的命令，根据你的系统情况适当修改
    const checkWeChatCommand = 'tasklist | findstr WeChat.exe';

    exec(checkWeChatCommand, (err, stdout, stderr) => {
        let result;
        if (err || stderr) {
            result = '无法检查 WeChat 状态。';
        } else if (stdout) {
            result = 'WeChat 正在运行。';
            // 这里可以添加获取 WeChat 版本的代码
            const filePath = 'C:\\Program Files (x86)\\Tencent\\WeChat\\WeChat.exe';
            const command = `(Get-Item '${filePath}').VersionInfo | Select-Object -ExpandProperty ProductVersion`;

            exec(`powershell -command "${command}"`, (error, stdout, stderr) => {
                if (error) {
                    console.error(`执行的错误: ${error}`);
                    return;
                }
                if (stderr) {
                    console.error(`执行的错误: ${stderr}`);
                    return;
                }

                result = `本机安装的版本号: ${stdout.trim()}`;
                console.log(`WeChat Version: ${stdout.trim()}`);
                mainWindow.webContents.send('wechat-check-result', result);
            });
        } else {
            result = 'WeChat 未运行。';
        }
        // 将检查结果发送到渲染进程
        mainWindow.webContents.send('wechat-check-result', result);
    });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});
