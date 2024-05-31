const { app, BrowserWindow, ipcMain, shell, dialog } = require("electron");
const { spawn, exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const { WechatyBuilder, ScanStatus, log } = require("wechaty");
const qrcodeTerminal = require("qrcode-terminal");
const { FileBox } = require("file-box");
const axios = require("axios");
const {
  PuppetBridgeJwpingWxbot,
  PuppetBridgeTtttupupWxhelperV3091019,
  PuppetBridgeTtttupupWxhelperV3090223,
  PuppetBridgeAtorberFusedV3090825,
  PuppetBridgeTtttupupWxhelperV3090581,
} = require("wechaty-puppet-bridge");
const xlsx = require("xlsx");
const startMqtt = require("./utils/mqtt-broker.js");

const { MqttGateway } = require("wechat-remote");

const v1 = process.versions;
console.log("v1:" + v1);

// 启动MQTT
startMqtt();

let bot;
let mainWindow;
let funtoolProcess = null;
let botProcess = null;
let wechatVersion = "";
let bridgeIsOn = false;
let mqttIsOn = false;
let wechatyIsOn = false;
let room;
let contact;
let replyJson = {};
let chatMessage = "";
// 读取配置文件
const config = require("./assets/config.json");

let jobs = [];

const getTime = () => {
  const timeutc = new Date().toLocaleString();
  return timeutc;
};

let result = `${getTime()}:程序启动，等待操作...<br>`;

async function onScan(qrcode, status) {
  log.info("StarterBot", "onScan: %s(%s)", ScanStatus[status], status);
  if (status === ScanStatus.Waiting || status === ScanStatus.Timeout) {
    const qrcodeImageUrl = [
      "https://wechaty.js.org/qrcode/",
      encodeURIComponent(qrcode),
    ].join("");
    log.info(
      "StarterBot",
      "onScan: %s(%s) - %s",
      ScanStatus[status],
      status,
      qrcodeImageUrl
    );

    qrcodeTerminal.generate(qrcode, { small: true }); // show qrcode on console
    result = `${getTime()}:登录二维码地址 ${qrcodeImageUrl}<br>` + result;
    mainWindow.webContents.send("action-result", result);
    let file;
    let filePath = "";
    try {
      file = FileBox.fromQRCode(qrcode);
      filePath = "assets/" + file.name;
      log.info("StarterBot", "onScan: %s", filePath);
      await file.toFile(filePath, true);
      mainWindow.webContents.send("qrcode-result", filePath);
    } catch (e) {
      result = `${getTime()}:二维码保存失败：${filePath}<br>` + result;
      log.error("二维码保存失败：", e);
      mainWindow.webContents.send("action-result", result);
    }
  } else {
    log.info("StarterBot", "onScan: %s(%s)", ScanStatus[status], status);
  }
}

async function onLogin(user) {
  log.info("onLogin", "%s login", user);
  result = `${getTime()}:${user} login<br>` + result;
  mainWindow.webContents.send("action-result", result);

  const roomList = await bot.Room.findAll();
  console.info("room count:", roomList.length);
  const contactList = await bot.Contact.findAll();
  console.info("contact count:", contactList.length);
  result =
    `${getTime()}:contact count: ${contactList.length},room count: ${
      roomList.length
    }<br>` + result;
  mainWindow.webContents.send("action-result", result);
}

async function onMessage(message) {
  log.info("onMessage", JSON.stringify(message));

  const text = message.text();
  const room = message.room();
  const talker = message.talker();
  // 生成2024-1-28 21:51:06格式的时间
  const timeutc = new Date().toLocaleString();
  result =
    `${timeutc} ${
      room ? "[" + (await room.topic()) + "]" : ""
    } ${talker.name()}: ${text}<br>` + result;
  mainWindow.webContents.send("action-result", result);
  // 1. send Image
  if (/^ding$/i.test(message.text())) {
    const fileBox = FileBox.fromUrl(
      "https://wechaty.github.io/wechaty/images/bot-qr-code.png"
    );
    await message.say(fileBox);
  }

  // 2. send Text

  if (/^dong$/i.test(message.text())) {
    await message.say("dingdingding");
  }
}

const startBot = () => {
  if (!wechatyIsOn) {
    let puppet = "";
    if (wechatVersion === "3.9.2.1000") {
      result = `${getTime()}:当前版本支持Bridge,使用wxhelper！<br>` + result;
      mainWindow.webContents.send("action-result", result);
      puppet = new PuppetBridgeTtttupupWxhelperV3090223();
      // puppet = new PuppetXp();
    } else if (wechatVersion === "3.9.5.1000") {
      result = `${getTime()}:当前版本支持Bridge,使用wxhelper！<br>` + result;
      mainWindow.webContents.send("action-result", result);
      puppet = new PuppetBridgeTtttupupWxhelperV3090581();
    } else if (wechatVersion === "3.9.8.1000") {
      result = `${getTime()}:当前版本支持Bridge,使用wxhelper！<br>` + result;
      mainWindow.webContents.send("action-result", result);
      puppet = new PuppetBridgeAtorberFusedV3090825();
    } else if (wechatVersion === "3.9.10.1000") {
      result = `${getTime()}:当前版本支持Bridge,使用wxhelper！<br>` + result;
      mainWindow.webContents.send("action-result", result);
      puppet = new PuppetBridgeTtttupupWxhelperV3091019();
    } else {
      result = `${getTime()}:不支持当前版本,使用web版微信登录！<br>` + result;
      mainWindow.webContents.send("action-result", result);
      puppet = "wechaty-puppet-wechat4u";
      bot.on("scan", onScan);
    }
    bot = WechatyBuilder.build({
      name: "ding-dong-bot",
      puppet,
    });

    const config = {
      events: [
        "login",
        "logout",
        "reset",
        "ready",
        "dirty",
        "dong",
        "error",
        // 'heartbeat',
        "friendship",
        "message",
        "post",
        "room-invite",
        "room-join",
        "room-leave",
        "room-topic",
        "scan",
      ],
      mqtt: {
        clientId: "ding-dong-test01", // 替换成自己的clientId，建议不少于16个字符串
        host: "127.0.0.1",
        password: "",
        port: 1883,
        username: "",
      },
      options: {
        secrectKey: "",
        simple: false,
      },
      token: "",
    };

    if (mqttIsOn) {
      bot.use(MqttGateway(config));
    }

    bot.on("login", onLogin);
    bot.on("message", onMessage);

    bot
      .start()
      .then(() => {
        wechatyIsOn = true;
        mainWindow.webContents.send("wechatyIsOn-result", wechatyIsOn);
        result = `${getTime()}:Bot已启动！<br>` + result;
        mainWindow.webContents.send("action-result", result);
        return log.info("StarterBot", "Starter Bot Started.");
      })
      .catch((err) => {
        console.error("bot start error", err);
        wechatyIsOn = false;
        mainWindow.webContents.send("wechatyIsOn-result", wechatyIsOn);
        result = `${getTime()}:启动Bot时发生错误...${err}<br>` + result;
        mainWindow.webContents.send("action-result", result);
      });
  } else {
    result = `${getTime()}:Bot已在运行中...<br>` + result;
    mainWindow.webContents.send("action-result", result);
  }
};

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1080,
    height: 720,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
    },
    autoHideMenuBar: true, // 自动隐藏菜单栏
  });
  // 当 Electron 应用准备好后，检查 WeChat 程序
  checkWeChat();

  mainWindow.loadFile("index.html");

  // 将config传递给渲染进程
  mainWindow.webContents.send("config", config);

  ipcMain.on("start-funtool", () => {
    let sidecarText = "";
    if (!funtoolProcess) {
      let execPath;
      if (wechatVersion === "3.9.2.1000") {
        execPath = path.join(__dirname, "assets", "funtool_wx=3.9.2.23.exe");
        funtoolProcess = spawn(execPath);
        sidecarText = "http://127.0.0.1:5555";
      } else if (wechatVersion === "3.9.8.1000") {
        execPath = path.join(__dirname, "assets", "wxbot-sidecar.exe");
        sidecarText = "http://127.0.0.1:8080";
        // 运行spawn(execPath)指定-b --auto-click-login参数，自动点击登录
        // funtoolProcess = spawn(execPath, ['-b', '--auto-click-login']);
        funtoolProcess = spawn(execPath);
      } else {
        result = `${getTime()}:不支持当前版本！<br>` + result;
        mainWindow.webContents.send("action-result", result);
      }

      funtoolProcess.on("spawn", () => {
        result =
          `${getTime()}:Bridge已启动,请求地址：${sidecarText}<br>` + result;
        mainWindow.webContents.send("action-result", result);
        bridgeIsOn = true;
        mainWindow.webContents.send("bridgeIsOn-result", bridgeIsOn);
        // 倒计时5秒后启动bot
        result = `${getTime()}:倒计时5秒后启动bot...<br>` + result;
        mainWindow.webContents.send("action-result", result);
        setTimeout(() => {
          startBot();
        }, 3000);
      });
      funtoolProcess.on("error", (err) => {
        result = `${getTime()}:启动Bridge时发生错误: ${err}<br>` + result;
        mainWindow.webContents.send("action-result", result);
        bridgeIsOn = false;
        mainWindow.webContents.send("bridgeIsOn-result", bridgeIsOn);
      });
    } else {
      result =
        `${getTime()}:Bridge已在运行中,请求地址：${sidecarText}<br>` + result;
      mainWindow.webContents.send("action-result", result);
    }
  });

  ipcMain.on("stop-funtool", () => {
    if (funtoolProcess) {
      funtoolProcess.kill();
      funtoolProcess = null;
      bridgeIsOn = false;
      mainWindow.webContents.send("bridgeIsOn-result", bridgeIsOn);
      result = `${getTime()}:funtool已停止...<br>` + result;
      mainWindow.webContents.send("action-result", result);
    } else {
      result = `${getTime()}:funtool未在运行...<br>` + result;
      mainWindow.webContents.send("action-result", result);
    }
  });

  ipcMain.on("start-bot", () => {
    startBot();
  });

  ipcMain.on("stop-bot", () => {
    const timeutc = getTime();
    if (wechatyIsOn && bot) {
      bot
        .stop()
        .then(() => {
          wechatyIsOn = false;
          mainWindow.webContents.send("wechatyIsOn-result", wechatyIsOn);
          result = `${timeutc}:bot已停止...<br>` + result;
          mainWindow.webContents.send("action-result", result);
        })
        .catch((err) => {
          console.error("bot stop error", err);
          result = `${timeutc}:bot停止时发生错误...${err}<br>` + result;
          mainWindow.webContents.send("action-result", result);
        });
      bot = null;
      wechatyIsOn = false;
    } else {
      result = `${timeutc}:Bot未在运行...<br>` + result;
      mainWindow.webContents.send("action-result", result);
    }
  });

  ipcMain.on("mqtt-status-changed", (event, mqttIsOnOff) => {
    console.info("mqtt-status-changed:", mqttIsOnOff);
    const timeutc = getTime();
    mqttIsOn = mqttIsOnOff;
    result =
      `${timeutc}:MQTT推送已${mqttIsOn ? "开启" : "关闭"}...<br>` + result;
    mainWindow.webContents.send("action-result", result);
  });

  ipcMain.on("restart-app", () => {
    app.relaunch();
    app.exit();
  });

  // 下载模板
  ipcMain.on("download-template", async (event) => {
    const templatePath = path.join(__dirname, "assets", "contacts.xlsx");

    // 弹出保存文件对话框
    const { filePath } = await dialog.showSaveDialog(mainWindow, {
      title: "保存模板",
      defaultPath: "template.xlsx",
      filters: [{ name: "Excel Files", extensions: ["xlsx"] }],
    });

    if (filePath) {
      // 将模板文件复制到用户选择的位置
      fs.copyFile(templatePath, filePath, (err) => {
        if (err) {
          console.error("Error copying the template file", err);
          result = `${getTime()}:下载模板失败: ${err.message}<br>` + result;
          mainWindow.webContents.send("action-result", result);
        } else {
          // 打开文件夹并选中文件
          shell.showItemInFolder(filePath);
          result = `${getTime()}:模板已下载到: ${filePath}<br>` + result;
          mainWindow.webContents.send("action-result", result);
        }
      });
    }
  });

  // 上传任务
  ipcMain.on("upload-excel", (event, filePath) => {
    try {
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(worksheet);
      jobs = data.filter((item) => item["name"] && item["id"] && item["text"]);

      result =
        `${getTime()}:读取Excel文件成功,待发送消息：${jobs.length}<br>` +
        result;
      mainWindow.webContents.send("action-result", result);

      jobs.forEach((item, index) => {
        const { id, name, text } = item;
        result = `${index + 1} ${name} ${text}<br>` + result;
      });

      result = "任务列表：<br>" + result;
      mainWindow.webContents.send("action-result", result);
    } catch (error) {
      console.error("Error reading the Excel file", error);
      result = `${getTime()}:读取Excel文件出错: ${error.message}<br>` + result;
      mainWindow.webContents.send("action-result", result);
    }
  });

  // 导出联系人
  ipcMain.on("export-contacts", async () => {
    try {
      const filePath = await exportContactsToExcel();
      result = `${getTime()}:联系人已导出完成...<br>` + result;
      mainWindow.webContents.send("action-result", result);
    } catch (error) {
      result = `${getTime()}:导出联系人时出错: ${error.message}<br>` + result;
      mainWindow.webContents.send("action-result", result);
    }
  });

  ipcMain.on("send-message", async (event, message) => {
    const timeutc = getTime();
    if (botProcess) {
      if (jobs.length > 0) {
        result = `${timeutc}:开始发送消息...<br>` + result;
        mainWindow.webContents.send("action-result", result);

        for (let i = 0; i < jobs.length; i++) {
          const item = jobs[i];
          const { id, name, text } = item;
          const contact = await bot.Contact.find({ id });
          if (contact) {
            await contact.say(text);
            result = `${timeutc}:${name} ${text}<br>` + result;
            mainWindow.webContents.send("action-result", result);
            // 随机延时200到1000毫秒
            await new Promise((resolve) =>
              setTimeout(resolve, Math.random() * 800 + 200)
            );
          } else {
            result = `${timeutc}:未找到联系人 ${name}<br>` + result;
            mainWindow.webContents.send("action-result", result);
          }
        }
        result = `${timeutc}:发送消息完成！<br>` + result;
        mainWindow.webContents.send("action-result", result);
      } else {
        result = `${timeutc}:没有待发送的消息...<br>` + result;
        mainWindow.webContents.send("action-result", result);
      }
    } else {
      result = `${timeutc}:Bot未启动...<br>` + result;
      mainWindow.webContents.send("action-result", result);
    }
  });

  ipcMain.on("chat-message", async (event, message) => {
    console.log("Received chat message:", message);
    chatMessage = message;
    result = `${getTime()}: ${message}<br>` + result;
    mainWindow.webContents.send("action-result", result);

    if (message === "是") {
      if (room) {
        await room.say(replyJson.message);
        result =
          `${getTime()}:消息已发送到群聊 ${replyJson.receiver}<br>` + result;
        mainWindow.webContents.send("action-result", result);
      } else {
        result = `${getTime()}:未找到群聊 ${replyJson.receiver}<br>` + result;
        mainWindow.webContents.send("action-result", result);
      }
      if (contact) {
        await contact.say(replyJson.message);
        result = `${getTime()}:消息已发送给 ${replyJson.receiver}<br>` + result;
        mainWindow.webContents.send("action-result", result);
      } else {
        result = `${getTime()}:未找到联系人 ${replyJson.receiver}<br>` + result;
        mainWindow.webContents.send("action-result", result);
      }
    } else if (message === "否") {
      room = null;
      contact = null;
      result = `${getTime()}:取消发送消息<br>` + result;
      mainWindow.webContents.send("action-result", result);
    } else {
      // 请求openai，处理收到的消息
      const prompt = `请根据用户的意图解析出结构化的json字符串，如果用户的意图是向某人发送信息，则返回对应的接收人姓名，否则返回空。
    以下是一个示例：
    输入：给张三发信息，告诉他明天上午9点到公司开会。
    输出json格式字符串：{"receiver": "张三","message": "明天上午9点到公司开会"}
    
    以下是需要正式解析的信息：
    输入：${message}
    输出json格式字符串：
    `;

      // 请求openai
      const { openai } = config;
      const { endpoint, key } = openai;
      const url = `${endpoint}/v1/chat/completions`;
      const headers = {
        "Content-Type": "application/json",
        Authorization: `Bearer ${key}`,
      };
      const data = {
        model: "gpt-3.5-turbo",
        messages: [
          { role: "system", content: "You are a helpful assistant." },
          { role: "user", content: prompt },
        ],
      };

      try {
        const response = await axios.post(url, data, { headers: headers });
        console.log(
          "openai response:",
          response.data.choices[0].message.content
        );
        const reply = response.data.choices[0].message.content;
        result = `${getTime()}:${reply}<br>` + result;
        mainWindow.webContents.send("action-result", result);
        try {
          replyJson = JSON.parse(reply);
        } catch (error) {
          result =
            `${getTime()}:解析openai返回的json失败: ${error.message}<br>` +
            result;
          mainWindow.webContents.send("action-result", result);
          result = `${getTime()}:尝试正则解析...<br>` + result;
          mainWindow.webContents.send("action-result", result);

          // 使用正则从reply中解析出receiver和message，reply的格式是：```json{"receiver": "大师", "message": "明天的会议取消了"}```
          const match = reply.match(/{(.+)}/);
          console.log("match:", match);
          result =
            `${getTime()}:正则解析返回中是否包含json字符串: ${match}<br>` +
            result;
          if (match) {
            const jsonStr = match[0];
            const json = JSON.parse(jsonStr);
            console.log("openai result:", json);
            replyJson = json;
          }
        }
        // 查找联系人并发送消息
        if (replyJson.receiver && replyJson.message) {
          contact = await bot.Contact.find({ name: replyJson.receiver });
          room = await bot.Room.find({ topic: replyJson.receiver });

          if (room) {
            result =
              `${getTime()}:找到群聊 ${await room.topic()},是否发送消息：<br>是：立即发送<br>否：取消发送<br>` +
              result;
            mainWindow.webContents.send("action-result", result);
          }

          if (contact) {
            result =
              `${getTime()}:找到联系人 ${contact.name()},是否发送消息：<br>是：立即发送<br>否：取消发送<br>` +
              result;
            mainWindow.webContents.send("action-result", result);
          }
        } else {
          result = `${getTime()}:未找到接收人或消息为空<br>` + result;
          mainWindow.webContents.send("action-result", result);
        }
      } catch (error) {
        console.error(error);
        result = `${getTime()}:请求openai失败: ${error.message}<br>` + result;
        mainWindow.webContents.send("action-result", result);
        // 还原输入框内容
        mainWindow.webContents.send("chat-message-result", message);
      }
    }
  });

  ipcMain.on("save-settings", (event, settings) => {
    console.log("Received settings:", settings);
    config.openai = {
      endpoint: settings.endpoint,
      key: settings.key,
    };
    try {
      // 保存配置文件
      fs.writeFileSync(
        path.join(__dirname, "assets", "config.json"),
        JSON.stringify(config, null, 2)
      );
      mainWindow.webContents.send("settings-saved", settings);
      result =
        `${getTime()}:Settings saved:${JSON.stringify(settings)}<br>` + result;
      mainWindow.webContents.send("action-result", result);
    } catch (error) {
      result = `${getTime()}:保存设置失败: ${error.message}<br>` + result;
      mainWindow.webContents.send("action-result", result);
    }
  });

  // bot.start()
  //     .then(() => {
  //         console.log('bot start success')
  //         const timeutc = new Date().toLocaleString();

  //         botProcess = true;
  //         result = `${timeutc}:Bot已启动！<br>` + result;
  //         mainWindow.webContents.send('action-result', result);
  //         return log.info('StarterBot', 'Starter Bot Started.')
  //     })
  //     .catch((err) => {
  //         const timeutc = new Date().toLocaleString();

  //         console.error('bot start error', err)
  //         result = `${timeutc}:启动Bot时发生错误...${err}<br>` + result;
  //         mainWindow.webContents.send('action-result', result);
  //     })
}

async function exportContactsToExcel() {
  result = `${getTime()}:正在导出联系人...<br>` + result;
  mainWindow.webContents.send("action-result", result);
  try {
    const contactList = await bot.Contact.findAll();
    const data = contactList
      .filter((contact) => contact.friend())
      .map((contact) => {
        result =
          `${getTime()}:正在导出联系人...${contact.id} ${contact.name()}<br>` +
          result;
        mainWindow.webContents.send("action-result", result);
        return {
          id: contact.id,
          name: contact.name(),
          alias: contact.alias() || "N/A", // 如果没有别名，显示 'N/A'
          text: "",
          state: "",
        };
      });

    // 转换数据为工作表
    const worksheet = xlsx.utils.json_to_sheet(data);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, "Contacts");

    // 生成Excel文件并保存
    const filePath = path.join(__dirname, "assets", "contacts.xlsx");
    xlsx.writeFile(workbook, filePath);
    result = `${getTime()}:联系人已导出到: ${filePath}<br>` + result;
    mainWindow.webContents.send("action-result", result);
    console.log("Contacts have been exported to Excel.");
    return filePath;
  } catch (error) {
    result = `${getTime()}:导出联系人时出错: ${error.message}<br>` + result;
    mainWindow.webContents.send("action-result", result);
    console.error("Failed to export contacts:", error);
    throw error; // 抛出错误供上层处理
  }
}

function checkWeChat() {
  // 检查 WeChat 是否启动的命令，根据你的系统情况适当修改
  const checkWeChatCommand = "tasklist | findstr WeChat.exe";

  exec(checkWeChatCommand, (err, stdout, stderr) => {
    let resultVersion = "";
    if (err || stderr) {
      resultVersion = "无法检查 WeChat 状态...";
    } else if (stdout) {
      resultVersion = "WeChat 正在运行...";
      // 这里可以添加获取 WeChat 版本的代码
      const filePath = "C:\\Program Files (x86)\\Tencent\\WeChat\\WeChat.exe";
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
        wechatVersion = stdout.trim();
        console.log(resultVersion);
        result = `${getTime()}:${resultVersion}<br>` + result;
        mainWindow.webContents.send("action-result", result);
        try {
          mainWindow.webContents.send("wechat-check-result", resultVersion);
          // console.log('send wechat-version');
        } catch (e) {
          console.error(e);
        }
      });
    } else {
      resultVersion = "WeChat 未运行...";
    }
    // 将检查结果发送到渲染进程
    mainWindow.webContents.send("wechat-check-result", resultVersion);
  });
}

app.whenReady().then(createWindow);

app.on("window-all-closed", function () {
  if (process.platform !== "darwin") app.quit();
});
