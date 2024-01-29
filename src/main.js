const {
    app,
    BrowserWindow,
    ipcMain,
    shell,
    dialog
} = require('electron');
const { spawn, exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const {
    WechatyBuilder,
    log,
} = require('wechaty')

const { FileBox } = require('file-box')
const { PuppetBridge } = require('wechaty-puppet-bridge')
const xlsx = require('xlsx');

let mainWindow;
let sidecarProcess = null;
let funtoolProcess = null;
let botProcess = null;

let jobs = [];

const getTime = () => {
    const timeutc = new Date().toLocaleString();
    return timeutc;
}

let result = `${getTime()}:程序启动，等待操作...<br>`;

async function onLogin(user) {
    log.info('onLogin', '%s login', user)
    result = `${getTime()}:${user} login<br>` + result;
    mainWindow.webContents.send('action-result', result);

    const roomList = await bot.Room.findAll()
    console.info('room count:', roomList.length)
    const contactList = await bot.Contact.findAll()
    console.info('contact count:', contactList.length)
    result = `${getTime()}:contact count: ${contactList.length},room count: ${roomList.length}<br>` + result;
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
        width: 900,
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

        if (!funtoolProcess) {
            const execPath = path.join(__dirname, 'assets', 'funtool_wx=3.9.2.23.exe');
            funtoolProcess = spawn(execPath);
            funtoolProcess.on('spawn', () => {
                result = `${getTime()}:funtool已启动！<br>` + result;
                mainWindow.webContents.send('action-result', result);
            });
            funtoolProcess.on('error', (err) => {
                result = `${getTime()}:启动funtool时发生错误: ${err}<br>` + result;
                mainWindow.webContents.send('action-result', result);
            });
        } else {
            result = `${getTime()}:funtool已在运行中...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('stop-funtool', () => {

        if (funtoolProcess) {
            funtoolProcess.kill();
            funtoolProcess = null;
            result = `${getTime()}:funtool已停止...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        } else {
            result = `${getTime()}:funtool未在运行...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('start-sidecar', () => {
        const timeutc = getTime();

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
        const timeutc = getTime();
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

    ipcMain.on('start-bot', () => {
        const timeutc = getTime();

        if (!botProcess) {
            bot.start()
                .then(() => {
                    botProcess = true;
                    result = `${timeutc}:Bot已启动！<br>` + result;
                    mainWindow.webContents.send('action-result', result);
                    return log.info('StarterBot', 'Starter Bot Started.')
                })
                .catch((err) => {
                    console.error('bot start error', err)
                    result = `${timeutc}:启动Bot时发生错误...${err}<br>` + result;
                    mainWindow.webContents.send('action-result', result);
                })
        } else {
            result = `${timeutc}:Bot已在运行中...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('stop-bot', () => {
        const timeutc = new Date().toLocaleString();
        if (botProcess) {
            bot.stop().then(() => {
                botProcess = null;
                result = `${timeutc}:bot已停止...<br>` + result;
                mainWindow.webContents.send('action-result', result);
            }).catch((err) => {
                console.error('bot stop error', err)
                result = `${timeutc}:bot停止时发生错误...${err}<br>` + result;
                mainWindow.webContents.send('action-result', result);
            })
        } else {
            result = `${timeutc}:Bot未在运行...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('restart-app', () => {
        app.relaunch();
        app.exit();
    });

    ipcMain.on('download-template', async (event) => {
        const templatePath = path.join(__dirname, 'assets', 'contacts.xlsx');

        // 弹出保存文件对话框
        const { filePath } = await dialog.showSaveDialog(mainWindow, {
            title: '保存模板',
            defaultPath: 'template.xlsx',
            filters: [
                { name: 'Excel Files', extensions: ['xlsx'] }
            ]
        });

        if (filePath) {
            // 将模板文件复制到用户选择的位置
            fs.copyFile(templatePath, filePath, (err) => {
                if (err) {
                    console.error('Error copying the template file', err);
                    result = `${getTime()}:下载模板失败: ${err.message}<br>` + result;
                    mainWindow.webContents.send('action-result', result);
                } else {
                    // 打开文件夹并选中文件
                    shell.showItemInFolder(filePath);
                    result = `${getTime()}:模板已下载到: ${filePath}<br>` + result;
                    mainWindow.webContents.send('action-result', result);
                }
            });
        }
    });

    ipcMain.on('upload-excel', (event, filePath) => {
        try {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet);
            jobs = data.filter(item => item['name'] && item['id'] && item['text']);

            result = `${getTime()}:读取Excel文件成功,待发送消息：${jobs.length}<br>` + result;
            mainWindow.webContents.send('action-result', result);

            jobs.forEach((item, index) => {
                const { id, name, text } = item;
                result = `${index + 1} ${name} ${text}<br>` + result;
            });

            result = '任务列表：<br>' + result;
            mainWindow.webContents.send('action-result', result);
        } catch (error) {
            console.error('Error reading the Excel file', error);
            result = `${getTime()}:读取Excel文件出错: ${error.message}<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('export-contacts', async () => {
        try {
            const filePath = await exportContactsToExcel();
            result = `${getTime()}:联系人已导出到: ${filePath}<br>` + result;
            mainWindow.webContents.send('action-result', result);
        } catch (error) {
            result = `${getTime()}:导出联系人时出错: ${error.message}<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    });

    ipcMain.on('send-message', async (event, message) => {
        const timeutc = getTime();
        if (botProcess) {
            if (jobs.length > 0) {
                result = `${timeutc}:开始发送消息...<br>` + result;
                mainWindow.webContents.send('action-result', result);

                for (let i = 0; i < jobs.length; i++) {
                    const item = jobs[i];
                    const { id, name, text } = item;
                    const contact = await bot.Contact.find({ id });
                    if (contact) {
                        await contact.say(text);
                        result = `${timeutc}:${name} ${text}<br>` + result;
                        mainWindow.webContents.send('action-result', result);
                    } else {
                        result = `${timeutc}:未找到联系人 ${name}<br>` + result;
                        mainWindow.webContents.send('action-result', result);
                    }
                }
                result = `${timeutc}:发送消息完成！<br>` + result;
                mainWindow.webContents.send('action-result', result);
            } else {
                result = `${timeutc}:没有待发送的消息...<br>` + result;
                mainWindow.webContents.send('action-result', result);
            }
        } else {
            result = `${timeutc}:Bot未启动...<br>` + result;
            mainWindow.webContents.send('action-result', result);
        }
    })

    bot.start()
        .then(() => {
            console.log('bot start success')
            const timeutc = new Date().toLocaleString();

            botProcess = true;
            result = `${timeutc}:Bot已启动！<br>` + result;
            mainWindow.webContents.send('action-result', result);
            return log.info('StarterBot', 'Starter Bot Started.')
        })
        .catch((err) => {
            const timeutc = new Date().toLocaleString();

            console.error('bot start error', err)
            result = `${timeutc}:启动Bot时发生错误...${err}<br>` + result;
            mainWindow.webContents.send('action-result', result);
        })
}

async function exportContactsToExcel() {
    try {
        const contactList = await bot.Contact.findAll();
        const data = contactList.filter(contact => contact.friend())
            .map(contact => ({
                id: contact.id,
                name: contact.name(),
                alias: contact.alias() || 'N/A', // 如果没有别名，显示 'N/A'
                text: '',
                state: '',
            }));

        // 转换数据为工作表
        const worksheet = xlsx.utils.json_to_sheet(data);
        const workbook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Contacts');

        // 生成Excel文件并保存
        const filePath = path.join(__dirname, 'assets', 'contacts.xlsx');
        xlsx.writeFile(workbook, filePath);
        result = `${getTime()}:联系人已导出到: ${filePath}<br>` + result;
        mainWindow.webContents.send('action-result', result);
        console.log('Contacts have been exported to Excel.');
        return filePath;
    } catch (error) {
        result = `${getTime()}:导出联系人时出错: ${error.message}<br>` + result;
        mainWindow.webContents.send('action-result', result);
        console.error('Failed to export contacts:', error);
        throw error; // 抛出错误供上层处理
    }
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

                resultVersion = `本机安装的客户端版本: ${stdout.trim()}`;
                console.log(resultVersion);
                try {
                    mainWindow.webContents.send('wechat-check-result', resultVersion);
                    console.log('send wechat-version');
                } catch (e) {
                    console.error(e);
                }
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
