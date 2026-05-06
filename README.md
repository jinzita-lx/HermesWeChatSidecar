# Hermes WeChat Sidecar (Windows)

PC WeChat 端的 sidecar，使用 [wxauto](https://github.com/cluic/wxauto) 控制已登录的微信桌面客户端，
通过 WebSocket 连接 Linux 上的 `hermes-wechat-adapter` 与 Hermes Agent 互通。

## 架构

```
WeChat PC <—— wxauto ——> sidecar.py <== WS ==> Linux adapter (10.0.0.2:8787 / SSH-tunnelled 127.0.0.1:8787) <—> Hermes
```

## 前置条件

1. PC 微信已登录，窗口可见（wxauto 通过 UI 自动化）。
2. Python 3.11（winget install Python.Python.3.11）。
3. 与 Linux adapter 网络可达。当前选择：SSH 隧道
   ```powershell
   ssh -N -L 8787:127.0.0.1:8787 root@<linux-host>
   ```
4. 从 Linux `/root/hermes-wechat-adapter/.env` 拷贝 `ADAPTER_AUTH_TOKEN` 到本地 `.env`。

## 安装

```powershell
cd C:\HermesWeChatSidecar
copy .env.example .env
notepad .env   # 填 ADAPTER_AUTH_TOKEN，并把 ADAPTER_WS_URL 中 token=REPLACE_ME 替换
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 运行

确保：
- 微信 PC 已登录并窗口可见
- SSH 隧道在跑：`ssh -N -L 8787:127.0.0.1:8787 root@<linux-host>`

```powershell
.\run.ps1
```

或直接：

```powershell
.\.venv\Scripts\python.exe -m src.main
```

## 测试

1. 启动 sidecar，看日志连上 WS。
2. 在微信「文件传输助手」里发 `/ping`。
3. 应该收到 `pong`（由 Linux adapter / Hermes 回发）。

## 文件清单

- `src/config.py` 加载 .env
- `src/dedup.py` 消息去重
- `src/wechat_provider.py` wxauto 封装
- `src/ws_client.py` WebSocket 客户端（心跳 + 断线重连）
- `src/command_executor.py` 处理 send_text / send_file / send_image
- `src/main.py` 入口
- `run.ps1` 启动脚本
