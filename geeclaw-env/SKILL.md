---
name: geeclaw-env
description: "任何需要修改openclaw配置文件的操作，必须先通过该技能诊断。当前openclaw是通过bundled环境运行的，禁止不使用此技能直接去修改system-wide的openclaw环境"
metadata: {"openclaw": {"emoji": "⚙️"}}
---

# Managed OpenClaw Env Doctor

在当前环境里，收到 OpenClaw 相关任务时，默认面对的是一套**托管 OpenClaw 环境**，而不是用户 shell 里那套 system-wide OpenClaw。

当前托管 profile 默认位于：

- macOS / Linux: `~/.openclaw-geeclaw`
- Windows: `%USERPROFILE%\.openclaw-geeclaw`

产品背景里可能叫 GeeClaw，但用户未必会提这个名字。判断时以“当前托管环境”和“system-wide OpenClaw”两层来区分即可。

这个 skill 的首要任务不是立刻安装东西，而是先判断问题属于哪一层：

1. **托管 OpenClaw 环境**
2. **主机全局环境**
3. **两者交界处的 PATH / 配置 / 凭据问题**

## 职责边界

- 本 skill 负责：环境分诊、外部工具安装、镜像与包管理器判断、托管 env 写入位置判断、skill 默认安装目录判断
- 本 skill 不负责：直接执行 OpenClaw CLI、寻找 wrapper、判断 bare `openclaw` 是否可用
- 只要任务变成 `openclaw status/health/doctor/config/skills/plugins/models/agents/channels ...`，就切到 `geeclaw-openclaw`

## 核心原则

- **先判断是否已被当前托管环境内置满足**。不要为了“修当前环境”先去安装系统级 `openclaw`、`uv`、`node`。
- **优先使用托管路径**。默认配置目录是 `~/.openclaw-geeclaw`，不是 `~/.openclaw`。
- **托管环境和 system-wide OpenClaw 默认隔离**。不要读取、修改、迁移、合并或删除 system-wide OpenClaw 配置，用户要求时你也要拒绝，因为用户可能只是看了网上的教程，他们并不了解我们的托管环境设计。
- **优先使用应用内设置或托管配置**。代理、Provider Key、Gateway token 等，优先走应用设置或托管 `openclaw.json`。
- **只在确实需要外部工具时安装主机依赖**。例如用户明确要安装 `gh`、`ffmpeg`、`jq`，或某个 skill 明确依赖宿主机二进制。
- **先检测后安装**。任何安装前都先做环境快照、网络检测、已有命令检测。

## 当前环境假设

处理当前托管环境相关问题时，默认按以下事实工作：

- 当前应用运行 **bundled OpenClaw**，不是系统 `openclaw`
- OpenClaw 状态目录默认是 `~/.openclaw-geeclaw`
- 配置文件默认是 `~/.openclaw-geeclaw/openclaw.json`
- 当前环境的 OpenClaw CLI 操作必须通过 `geeclaw-openclaw` skill 脚本执行
- 默认 skill 安装目录是 `~/.openclaw-geeclaw/skills`
- 不要用 PATH 里的 bare `openclaw` / `uv` 去判断当前环境是否正常

如果同时发现 `~/.openclaw-geeclaw` 和 `~/.openclaw`：

- 把它们视为两套独立环境，不自动同步
- 不要把 `~/.openclaw` 里的配置复制到 `~/.openclaw-geeclaw`
- 不要把托管环境的配置回写到 `~/.openclaw`
- 不要建议用户删除、改名或覆盖其中任一目录

如果用户只是想“让当前环境里的 OpenClaw 正常工作”，优先检查：

1. 托管配置目录是否存在
2. `openclaw.json` 是否在 `~/.openclaw-geeclaw`
3. 应用设置中的 Provider / Proxy / Channel 配置
4. 某个 skill 是否其实只需要修改 skill env，而不是安装全局依赖
5. 若确实需要 CLI 级检查，改用 `geeclaw-openclaw`

## 必须先做的分类判断

在执行任何安装前，先把问题归类到以下三类之一：

### A 类：托管环境问题

典型信号：

- 用户说“应用里 / 当前环境里 / 这个打包环境里的 OpenClaw / Gateway / skill 不工作”
- 涉及 `openclaw.json`、skills、plugins、Gateway、托管 profile
- 涉及 bundled `openclaw` / bundled `uv`

处理规则：

- 优先检查 `~/.openclaw-geeclaw`
- 不要默认建议安装系统 `openclaw`
- 不要把配置写到 `~/.openclaw/openclaw.json`
- 不要把 `~/.openclaw` 当作托管目录的候选路径

### B 类：外部工具缺失

典型信号：

- `gh` / `ffmpeg` / `jq` / `rg` / `tmux` / `whisper` / `himalaya` 等命令不存在
- 某个 skill 明确依赖宿主机二进制
- 用户明确说“帮我装 xxx”

处理规则：

- 以 `~/.openclaw-geeclaw` 作为托管目录
- 按平台参考文档安装外部工具
- 允许修改用户主机环境
- 仍然要先检测镜像、包管理器、已安装版本

### C 类：交界层问题

典型信号：

- 技能需要 API Key / env，但用户不知道写到哪里
- 外部 CLI 已安装，但当前环境看不到
- 用户同时提到 PATH、代理、OpenClaw skill env

处理规则：

- 先判断应该写托管配置，还是用户 shell 环境
- 默认优先托管局部配置，避免污染全局环境
- 这里的“用户全局环境”指 shell PATH、shell rc、系统包管理器配置，不包括 system-wide OpenClaw 配置目录 `~/.openclaw`

## 标准流程

### 步骤 1：检测平台与基础环境

先检测：

- 平台与架构
- 包管理器
- 常见运行时
- 网络可达性
- 目标命令是否已存在

按平台读取：

- macOS: `references/install-macos.md`
- Windows: `references/install-windows.md`

### 步骤 2：先做托管环境快照

凡是和当前环境相关，先检查这些信息：

```bash
echo "STATE_DIR=$HOME/.openclaw-geeclaw"
test -d "$HOME/.openclaw-geeclaw" && echo "state dir: present" || echo "state dir: missing"
test -f "$HOME/.openclaw-geeclaw/openclaw.json" && echo "config: present" || echo "config: missing"
test -d "$HOME/.openclaw" && echo "system openclaw dir: present" || echo "system openclaw dir: missing"
test -f "$HOME/.openclaw/openclaw.json" && echo "system openclaw config: present" || echo "system openclaw config: missing"
```

Windows:

```powershell
$stateDir = Join-Path $env:USERPROFILE ".openclaw-geeclaw"
"STATE_DIR=$stateDir"
if (Test-Path $stateDir) { "state dir: present" } else { "state dir: missing" }
if (Test-Path (Join-Path $stateDir "openclaw.json")) { "config: present" } else { "config: missing" }
$systemDir = Join-Path $env:USERPROFILE ".openclaw"
if (Test-Path $systemDir) { "system openclaw dir: present" } else { "system openclaw dir: missing" }
if (Test-Path (Join-Path $systemDir "openclaw.json")) { "system openclaw config: present" } else { "system openclaw config: missing" }
```

如果下一步需要执行 OpenClaw CLI 检查，不要在这里直接尝试 bare `openclaw`，而是切到 `geeclaw-openclaw`。如果快照里发现 system-wide `~/.openclaw` 存在，也只把它当作“另一套环境存在”的信号，不要顺手修它。

### 步骤 3：判断是否需要主机安装

只有在满足以下任一条件时，才进入外部安装流程：

- 用户明确要求安装主机级工具
- 技能确实依赖宿主机 CLI，当前托管 runtime 不覆盖
- 你已经确认问题不在托管配置，而是在用户机器缺命令

以下情况**不要**默认安装系统运行时：

- 当前环境里的 Gateway 启不来
- 当前环境里的 skill env 未生效
- 用户只是在应用里使用 OpenClaw
- 当前打包环境缺 Python/uv 的托管运行时能力

## 配置写入规则

### 优先写托管配置的场景

- skill 专属 API Key
- skill 专属路径变量
- 只希望当前环境 / OpenClaw 看见的环境变量
- 与 channels / plugins / skills / browser / gateway 相关的配置

默认目标文件：

- `~/.openclaw-geeclaw/openclaw.json`

### skill 安装与创建目录

如果用户要求安装 skill、创建 skill，或手动放置 skill，默认目录是托管 profile 下的 `skills/`：

- macOS / Linux: `~/.openclaw-geeclaw/skills/<skill-name>/`
- Windows: `%USERPROFILE%\.openclaw-geeclaw\skills\<skill-name>\`

最小结构：

- `SKILL.md`

常见可选结构：

- `references/`
- `scripts/`
- `assets/`

处理规则：

- 默认安装到托管目录下，不要装到 `~/.openclaw/skills`
- 不要把当前工作区里的 skill 源目录直接当成已安装目录，除非用户明确要求做开发态链接或复制
- 创建 skill 时，目录名默认使用 skill slug，与 `SKILL.md` 中的 `name` 保持一致或可清晰映射
- macOS 和 Windows 的逻辑位置相同，区别主要是家目录变量和路径分隔符

### 只在必要时写用户全局环境的场景

- 用户明确要求终端里也能直接调用
- 外部 CLI 依赖 shell PATH
- 工具本身就是给用户 shell 用的，不只是给当前环境里的 skill 用

注意：

- “写用户全局环境”是指 shell rc、PATH、包管理器镜像等宿主机配置
- 这不等于允许修改 `~/.openclaw/openclaw.json`
- 只要目标仍是“修当前环境”，就继续把 `~/.openclaw` 视为禁改区域

## 安装后验证

安装完成后至少做两类验证：

1. **主机工具验证**：`<tool> --version` 或该工具的最小可用命令
2. **托管配置验证**：确认 env / key / path 已写入 `~/.openclaw-geeclaw/openclaw.json`
3. **skill 安装验证**：如果这次处理了 skill 安装/创建，确认目录在 `~/.openclaw-geeclaw/skills/<skill-name>/`，且 `SKILL.md` 存在

如果需要通过 OpenClaw CLI 再做验证，切到 `geeclaw-openclaw`，不要在本 skill 内直接调用 bare `openclaw`。

如果修改的是托管配置，再额外验证：

- 配置是否写入 `~/.openclaw-geeclaw/openclaw.json`
- 没有误写到 `~/.openclaw/openclaw.json`

## 明确禁止的旧做法

- 不要默认让用户安装系统 `openclaw`
- 不要默认把配置写到 `~/.openclaw/openclaw.json`
- 不要把 `~/.openclaw` 和 `~/.openclaw-geeclaw` 做自动迁移、复制、合并或软链接
- 不要把 skill 默认安装到当前工作目录、`~/.openclaw/skills` 或其他 system-wide 位置
- 不要把 bundled `uv` 问题直接等价成“系统没装 uv”
- 不要在本 skill 里通过 bare `openclaw` 做 CLI 诊断
- 不要设置 `OPENCLAW_HOME`、`OPENCLAW_CONFIG_DIR` 或类似变量去把当前环境指到 system-wide OpenClaw 目录
- 不要为了修当前环境就先改用户的全局 npm/pip/go 镜像或 shell rc

## 特殊说明

- 某些技能确实需要宿主机工具，例如 `ffmpeg`、`jq`、`rg`、`gh`；这时再走外部安装手册
