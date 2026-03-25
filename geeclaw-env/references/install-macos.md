# macOS 外部工具安装参考（托管 OpenClaw 环境）

这份文档只用于 **macOS 上安装主机级外部工具**。

如果问题是当前托管环境里的 OpenClaw / Gateway / skill 配置，请先回到 `SKILL.md` 的托管环境检查流程，不要一上来安装系统 `openclaw`、`node`、`uv`。

修当前环境时，默认只处理 `~/.openclaw-geeclaw`。如果机器上同时有 `~/.openclaw`，把它视为另一套 system-wide OpenClaw 环境，不要顺手改它。

## 0. 先判断是不是托管环境问题

先检查：

```bash
echo "=== 托管环境快照 ==="
echo "STATE_DIR=$HOME/.openclaw-geeclaw"
test -d "$HOME/.openclaw-geeclaw" && echo "state dir: present" || echo "state dir: missing"
test -f "$HOME/.openclaw-geeclaw/openclaw.json" && echo "config: present" || echo "config: missing"
test -d "$HOME/.openclaw" && echo "system openclaw dir: present" || echo "system openclaw dir: missing"
test -f "$HOME/.openclaw/openclaw.json" && echo "system openclaw config: present" || echo "system openclaw config: missing"
```

如果你只是要修当前环境：

- 配置文件默认改 `~/.openclaw-geeclaw/openclaw.json`
- 不要改 `~/.openclaw/openclaw.json`
- 不要把 `~/.openclaw` 的配置迁移、复制、合并到 `~/.openclaw-geeclaw`
- 不要默认建议安装系统 `openclaw`
- 如果需要执行 `openclaw status/health/doctor/config/...`，改用 `geeclaw-openclaw`

## 1. 主机环境快照

仅当确实需要安装外部工具时执行：

```bash
echo "=== macOS 主机环境快照 ==="
echo "macOS $(sw_vers -productVersion 2>/dev/null || echo 'N/A'), $(uname -m)"
command -v brew >/dev/null 2>&1 && echo "brew: $(brew --version | head -1)" || echo "brew: missing"
for cmd in node npm python3 pip3 go uv git gh jq rg ffmpeg tmux whisper himalaya; do
  if command -v "$cmd" >/dev/null 2>&1; then
    echo "$cmd: present"
  else
    echo "$cmd: missing"
  fi
done
```

## 2. 网络与镜像检测

```bash
echo "=== 网络检测 ==="
for url in "https://github.com" "https://registry.npmjs.org" "https://pypi.org"; do
  if curl -sI --connect-timeout 3 "$url" >/dev/null 2>&1; then
    echo "reachable: $url"
  else
    echo "unreachable: $url"
  fi
done
```

如果用户在中国大陆且外网源慢：

- Homebrew 镜像只在**真的要用 brew 安装主机工具**时再配
- npm / pip 镜像只在**真的要用系统 npm / pip 安装**时再配
- 不要为了修当前托管 runtime 去改用户全局镜像

## 3. 包管理器策略

优先顺序：

1. 已有 `brew` 就直接用
2. 没有 `brew` 且用户明确需要主机级工具时，再安装 Homebrew
3. 需要 sudo 的步骤做不到时，给用户明确分步指引

Apple Silicon 上安装好 Homebrew 后，如命令找不到，补：

```bash
eval "$(/opt/homebrew/bin/brew shellenv)"
```

## 4. 常用主机级工具安装

### `gh`

```bash
brew install gh
gh --version
```

### `jq`

```bash
brew install jq
jq --version
```

### `rg`

```bash
brew install ripgrep
rg --version
```

### `ffmpeg`

```bash
brew install ffmpeg
ffmpeg -version
```

### `tmux`

```bash
brew install tmux
tmux -V
```

### `himalaya`

```bash
brew install himalaya
himalaya --version
```

## 5. 仅在外部工具明确需要时安装运行时

### Node.js + npm

只在这些情况安装：

- 用户明确要装依赖 npm 的主机 CLI
- 目标工具必须通过 npm 全局安装

```bash
brew install node
node --version
npm --version
```

如确需国内镜像：

```bash
npm config set registry https://registry.npmmirror.com
npm config get registry
```

### Python 3 + pip3

只在这些情况安装：

- 目标工具是 Python CLI
- skill 明确依赖系统 `python3` / `pip3`

```bash
brew install python
python3 --version
pip3 --version
```

如需验证不是 stub：

```bash
python3 -c "import sys; print(sys.version)"
```

### `uv`

**重要**：不要为了当前打包环境本体去安装系统 `uv`。当前打包环境自带 `uv`。

只在这些情况安装：

- 用户要在 shell 里直接执行 `uv`
- 某个外部工具明确要求主机 `uv`

```bash
brew install uv
uv --version
```

## 6. 托管配置与 skill 目录

如果最终需要给某个 skill 配 env，默认写到：

- `~/.openclaw-geeclaw/openclaw.json`

即使检测到 `~/.openclaw/openclaw.json` 也不要把它当作备选写入位置。

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

- `~/.openclaw/openclaw.json`
- 不要为了“同步配置”去同时改两个 profile

如果用户要求安装 skill 或创建 skill，默认目录是：

- `~/.openclaw-geeclaw/skills/<skill-name>/`

最小要求：

- `~/.openclaw-geeclaw/skills/<skill-name>/SKILL.md`

不要默认装到这些位置：

- `~/.openclaw/skills`
- 当前工作区源码目录

## 7. 当前环境专用提醒

- 修当前环境里的 Gateway：先查托管配置和应用设置，不先装系统 `openclaw`
- 修当前环境里的 skill env：先改 `~/.openclaw-geeclaw/openclaw.json`
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
