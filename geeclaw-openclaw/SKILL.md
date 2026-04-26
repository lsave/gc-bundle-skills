---
name: geeclaw-openclaw
description: "任何需要执行 openclaw 命令的场景，都必须通过本 skill 的脚本执行。脚本会寻找正确的openclaw命令，拒绝误用未知的 system-wide openclaw。适用于 gateway、status、health、doctor、config、skills、plugins、models、agents、channels 等 OpenClaw CLI 操作。"
metadata: {"openclaw": {"emoji": "⚙️"}}
---

# GeeClaw OpenClaw CLI

## Mandatory

不要直接执行 `openclaw ...`。

统一使用本 skill 的脚本：

### macOS / Linux

```bash
bash <skill_dir>/scripts/openclaw-mac.sh <command> [args...]
```

### Windows

```cmd
<skill_dir>\scripts\openclaw-win.cmd <command> [args...]
```

脚本只做三件事：

1. 寻找 GeeClaw wrapper
2. 拒绝未知的 system `openclaw`
3. 找到后直接执行

## 常见命令

```bash
bash <skill_dir>/scripts/openclaw-mac.sh status
bash <skill_dir>/scripts/openclaw-mac.sh health
bash <skill_dir>/scripts/openclaw-mac.sh doctor
bash <skill_dir>/scripts/openclaw-mac.sh config get gateway.port
bash <skill_dir>/scripts/openclaw-mac.sh skills list
bash <skill_dir>/scripts/openclaw-mac.sh plugins list
```

## 默认不要执行的命令

- `gateway run/start/stop/restart/install/uninstall`
- `daemon start/stop/restart/install/uninstall`
- `reset`
- `uninstall`

## 配置文件位置

GeeClaw 托管配置默认在：

```text
~/.openclaw-geeclaw/openclaw.json
```

## 故障处理

如果脚本提示“找不到 GeeClaw wrapper”：

- 不要退回去直接调用 bare `openclaw`
- 让用户提供 GeeClaw 安装位置，或使用应用内显示的 CLI 命令
