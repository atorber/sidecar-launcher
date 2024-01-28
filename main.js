const { app, BrowserWindow, ipcMain } = require('electron');
const { spawn, exec } = require('child_process');
const path = require('path');
const {
    WechatyBuilder,
    log,
} = require('wechaty')
const { FileBox } = require('file-box')
const { PuppetBridge } = require('wechaty-puppet-bridge')

let mainWindow;
let sidecarProcess = null;
let funtoolProcess = null;
let result = '程序启动，载入中...';

async function onLogin(user) {
    log.info('onLogin', '%s login', user)
    result = `${user} login<br>` + result;
    mainWindow.webContents.send('action-result', result);

    const roomList = await bot.Room.findAll()
    console.info('room count:', roomList.length)
    const contactList = await bot.Contact.findAll()
    console.info('contact count:', contactList.length)
    result = `contact count: ${contactList.length},room count: ${roomList.length}<br>` + result;
    mainWindow.webContents.send('action-result', result);

}

async function onMessage(message) {
    log.info('onMessage', JSON.stringify(message))
    const text = message.text()
    const room = message.room()
    const talker = message.talker()
    // 生成2024-1-28 21:51:06格式的时间
    const timeutc = new Date().toLocaleString()
    result = `${timeutc} ${room ? '[' + await room.topic() + ']' : ''} ${talker.name()}: ${text}<br>` + result
    mainWindow.webContents.send('action-result', result); 
    // 1. send Image
    if (/^ding$/i.test(message.text())) {
        const fileBox = FileBox.fromUrl('https://wechaty.github.io/wechaty/images/bot-qr-code.png')
        await message.say(fileBox)
    }

    // 2. send Text

    if (/^dong$/i.test(message.text())) {
        await message.say('dingdingding')
    }

}

const puppet = new PuppetBridge({
    nickName: '大师'  // 登录微信的昵称
})
const bot = WechatyBuilder.build({
    name: 'ding-dong-bot',
    puppet,
})

bot.on('login', onLogin)
bot.on('message', onMessage)

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
        const timeutc = new Date().toLocaleString();
        if (!funtoolProcess) {
            const execPath = path.join(__dirname, 'assets', 'funtool_wx=3.9.2.23.exe');
            funtoolProcess = spawn(execPath);
            funtoolProcess.on('spawn', () => {
                result = `${timeutc}:funtool已启动！<br>` + result;
                mainWindow.webContents.send('action-result', result);
            });
            funtoolProcess.on('error', (err) => {
                result = `${timeutc}:启动funtool时发生错误: ${err}<br>` + result;
                mainWindow.webContents.send('action-result', result);
            });
        } else {
            result = `${timeutc}:funtool已在运行中...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('stop-funtool', () => {
        const timeutc = new Date().toLocaleString();

        if (funtoolProcess) {
            funtoolProcess.kill();
            funtoolProcess = null;
            result = `${timeutc}:funtool已停止...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        } else {
            result = `${timeutc}:funtool未在运行...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('start-sidecar', () => {
        const timeutc = new Date().toLocaleString();

        if (!sidecarProcess) {
            const execPath = path.join(__dirname, 'assets', 'wxbot-sidecar.exe');
            sidecarProcess = spawn(execPath);
            sidecarProcess.on('spawn', () => {
                result = `${timeutc}:Sidecar已启动！<br>` + result;
                mainWindow.webContents.send('action-result', result);
            });
            sidecarProcess.on('error', (err) => {
                result = `${timeutc}:启动Sidecar时发生错误: ${err}<br>` + result;
                mainWindow.webContents.send('action-result', result);
            });
        } else {
            result = `${timeutc}:Sidecar已在运行中...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('stop-sidecar', () => {
        const timeutc = new Date().toLocaleString();
        if (sidecarProcess) {
            sidecarProcess.kill();
            sidecarProcess = null;
            result = `${timeutc}:Sidecar已停止...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        } else {
            result = `${timeutc}:Sidecar未在运行...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    bot.start()
    .then(() => {
        return log.info('StarterBot', 'Starter Bot Started.')
    })
    .catch(console.error)
}

function checkWeChat() {
    // 检查 WeChat 是否启动的命令，根据你的系统情况适当修改
    const checkWeChatCommand = 'tasklist | findstr WeChat.exe';

    exec(checkWeChatCommand, (err, stdout, stderr) => {
        let resultVersion = '';
        if (err || stderr) {
            resultVersion = '无法检查 WeChat 状态...';
        } else if (stdout) {
            resultVersion = 'WeChat 正在运行...';
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

                resultVersion = `本机安装的版本号: ${stdout.trim()}`;
                console.log(`WeChat Version: ${stdout.trim()}`);
                mainWindow.webContents.send('wechat-check-result', resultVersion);
            });
        } else {
            resultVersion = 'WeChat 未运行...';
        }
        // 将检查结果发送到渲染进程
        mainWindow.webContents.send('wechat-check-result', resultVersion);
    });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});
