# Windows 外部工具安装参考（托管 OpenClaw 环境）

这份文档只用于 **Windows 上安装主机级外部工具**。

如果问题是当前托管环境里的 OpenClaw / Gateway / skill 配置，请先检查托管环境，不要默认安装系统 `openclaw`、`node`、`uv`。

修当前环境时，默认只处理 `%USERPROFILE%\.openclaw-geeclaw`。如果机器上同时有 `%USERPROFILE%\.openclaw`，把它视为另一套 system-wide OpenClaw 环境，不要顺手改它。

## 0. 先判断是不是托管环境问题

```powershell
$stateDir = Join-Path $env:USERPROFILE ".openclaw-geeclaw"
"=== 托管环境快照 ==="
"STATE_DIR=$stateDir"
if (Test-Path $stateDir) { "state dir: present" } else { "state dir: missing" }
if (Test-Path (Join-Path $stateDir "openclaw.json")) { "config: present" } else { "config: missing" }
$systemDir = Join-Path $env:USERPROFILE ".openclaw"
if (Test-Path $systemDir) { "system openclaw dir: present" } else { "system openclaw dir: missing" }
if (Test-Path (Join-Path $systemDir "openclaw.json")) { "system openclaw config: present" } else { "system openclaw config: missing" }
```

如果你只是要修当前环境：

- 配置文件默认改 `%USERPROFILE%\.openclaw-geeclaw\openclaw.json`
- 不要改 `%USERPROFILE%\.openclaw\openclaw.json`
- 不要把 `%USERPROFILE%\.openclaw` 的配置迁移、复制、合并到 `%USERPROFILE%\.openclaw-geeclaw`
- 不要默认建议安装系统 `openclaw`
- 如果需要执行 `openclaw status/health/doctor/config/...`，改用 `geeclaw-openclaw`

## 1. 主机环境快照

仅当确实需要安装外部工具时执行：

```powershell
chcp 65001 >nul
"=== Windows 主机环境快照 ==="
$ver = [System.Environment]::OSVersion.Version
"Windows Build: $($ver.Major).$($ver.Minor).$($ver.Build)"
foreach ($pm in @("scoop", "winget", "choco")) {
  if (Get-Command $pm -ErrorAction SilentlyContinue) { "$pm: present" } else { "$pm: missing" }
}
foreach ($cmd in @("node", "npm", "python", "pip", "go", "uv", "git", "gh", "jq", "rg", "ffmpeg", "himalaya")) {
  if (Get-Command $cmd -ErrorAction SilentlyContinue) { "$cmd: present" } else { "$cmd: missing" }
}
```

## 2. 网络与镜像检测

```powershell
"=== 网络检测 ==="
@("https://github.com", "https://registry.npmjs.org", "https://pypi.org") | ForEach-Object {
  try {
    Invoke-WebRequest -Uri $_ -TimeoutSec 3 -UseBasicParsing | Out-Null
    "reachable: $_"
  } catch {
    "unreachable: $_"
  }
}
```

如果用户在中国大陆且网络慢：

- 只有在真的要装主机级 npm/pip 工具时才改对应镜像
- 不要为了修当前托管 runtime 去改用户全局镜像

## 3. 包管理器策略

优先顺序：

1. 已有 `scoop` 就用 `scoop`
2. 否则用 `winget`
3. 再不行才考虑 `choco`

不要在已有包管理器的系统上无谓再装新的包管理器。

## 4. 常用主机级工具安装

### `gh`

```powershell
scoop install gh
gh --version
```

如无 scoop：

```powershell
winget install GitHub.cli
gh --version
```

### `jq`

```powershell
scoop install jq
jq --version
```

### `rg`

```powershell
scoop install ripgrep
rg --version
```

### `ffmpeg`

```powershell
scoop install ffmpeg
ffmpeg -version
```

### `himalaya`

```powershell
winget install himalaya.himalaya
himalaya --version
```

## 5. 仅在外部工具明确需要时安装运行时

### Node.js + npm

只在这些情况安装：

- 用户明确要装依赖 npm 的主机 CLI
- 目标工具必须通过 npm 全局安装

```powershell
scoop install nodejs
node --version
npm --version
```

或：

```powershell
winget install OpenJS.NodeJS.LTS
node --version
npm --version
```

如确需国内镜像：

```powershell
npm config set registry https://registry.npmmirror.com
npm config get registry
```

### Python + pip

只在这些情况安装：

- 目标工具是 Python CLI
- skill 明确依赖系统 `python` / `pip`

```powershell
scoop install python
python --version
pip --version
```

如果 `python` 指向 Windows Store stub，需要先装真实 Python，再重开终端。

### `uv`

**重要**：不要为了当前打包环境本体去安装系统 `uv`。当前打包环境自带 `uv`，而且 Windows 下还可能配合随包 `node.exe` 运行内置 `openclaw`。

只在这些情况安装：

- 用户要在 PowerShell 里直接执行 `uv`
- 某个外部工具明确要求主机 `uv`

```powershell
scoop install uv
uv --version
```

或：

```powershell
winget install astral-sh.uv
uv --version
```

## 6. 托管配置与 skill 目录

如果最终需要给某个 skill 配 env，默认写到：

- `%USERPROFILE%\.openclaw-geeclaw\openclaw.json`

即使检测到 `%USERPROFILE%\.openclaw\openclaw.json` 也不要把它当作备选写入位置。

示例：

```json5
{
  skills: {
    entries: {
      "openai-image-gen": {
        env: {
          OPENAI_API_KEY: "sk-..."
        }
      }
    }
  }
}
```

不要写到旧路径：

- `%USERPROFILE%\.openclaw\openclaw.json`
- 不要为了“同步配置”去同时改两个 profile

如果用户要求安装 skill 或创建 skill，默认目录是：

- `%USERPROFILE%\.openclaw-geeclaw\skills\<skill-name>\`

最小要求：

- `%USERPROFILE%\.openclaw-geeclaw\skills\<skill-name>\SKILL.md`

不要默认装到这些位置：

- `%USERPROFILE%\.openclaw\skills`
- 当前工作区源码目录

## 7. 当前环境专用提醒

- 修当前环境里的 Gateway：先查托管配置和应用设置，不先装系统 `openclaw`
- 修当前环境里的 skill env：先改 `%USERPROFILE%\.openclaw-geeclaw\openclaw.json`
- 修当前环境里的 Python/uv 相关运行时：先判断 bundled runtime 是否已存在
- 若需要 OpenClaw CLI 级检查：切到 `geeclaw-openclaw`，不要直接试 bare `openclaw`
- 如果用户明确要修终端里的 system-wide `openclaw`，把那当成独立任务，不要混进当前环境修复流程

## 8. 用户手动操作指引模板

当需要用户自己执行命令时，按这个格式输出：

1. 说明为什么需要手动操作
2. 给出编号步骤
3. 每步只做一件事
4. 命令放代码块
5. 说明成功标志
6. 让用户完成后回复“已完成”
