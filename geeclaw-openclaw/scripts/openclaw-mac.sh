#!/usr/bin/env bash
set -euo pipefail

PREFIX="[geeclaw-openclaw]"

debug_log() {
  if [ "${GEECLAW_OPENCLAW_DEBUG:-0}" = "1" ]; then
    echo "${PREFIX} $*" >&2
  fi
}

resolve_path() {
  local candidate="$1"
  [ -n "$candidate" ] || return 1

  if command -v realpath >/dev/null 2>&1; then
    realpath "$candidate" 2>/dev/null && return 0
  fi

  if command -v python3 >/dev/null 2>&1; then
    python3 -c 'import os, sys; print(os.path.realpath(sys.argv[1]))' "$candidate" 2>/dev/null && return 0
  fi

  printf '%s\n' "$candidate"
}

is_geeclaw_wrapper() {
  local candidate="$1"
  local resolved=""

  [ -f "$candidate" ] || return 1
  resolved="$(resolve_path "$candidate")"

  case "$resolved" in
    */GeeClaw.app/Contents/Resources/managed-bin/openclaw|/opt/GeeClaw/resources/managed-bin/openclaw)
      return 0
      ;;
  esac

  grep -qi "GeeClaw" "$candidate" 2>/dev/null
}

append_candidate() {
  local candidate="$1"
  [ -n "$candidate" ] || return 0
  CANDIDATES+=("$candidate")
}

find_wrapper() {
  local from_path=""
  local candidate=""

  if [ -n "${GEECLAW_OPENCLAW_WRAPPER:-}" ] && is_geeclaw_wrapper "${GEECLAW_OPENCLAW_WRAPPER}"; then
    debug_log "使用环境变量指定的 wrapper: ${GEECLAW_OPENCLAW_WRAPPER}"
    printf '%s\n' "${GEECLAW_OPENCLAW_WRAPPER}"
    return 0
  fi

  append_candidate "/Applications/GeeClaw.app/Contents/Resources/managed-bin/openclaw"
  append_candidate "${HOME}/Applications/GeeClaw.app/Contents/Resources/managed-bin/openclaw"
  append_candidate "/opt/GeeClaw/resources/managed-bin/openclaw"

  if command -v mdfind >/dev/null 2>&1; then
    while IFS= read -r app_path; do
      [ -n "$app_path" ] || continue
      append_candidate "${app_path}/Contents/Resources/managed-bin/openclaw"
    done < <(mdfind "kMDItemCFBundleIdentifier == 'app.dtminds.geeclaw'" 2>/dev/null || true)
  fi

  for candidate in "${CANDIDATES[@]:-}"; do
    if is_geeclaw_wrapper "$candidate"; then
      debug_log "使用已知 GeeClaw 安装路径中的 wrapper: $candidate"
      printf '%s\n' "$candidate"
      return 0
    fi
  done

  if command -v openclaw >/dev/null 2>&1; then
    from_path="$(command -v openclaw)"
    if is_geeclaw_wrapper "$from_path"; then
      debug_log "使用 PATH 中检测到的 GeeClaw wrapper: $from_path"
      printf '%s\n' "$from_path"
      return 0
    fi

    debug_log "忽略 PATH 中的非 GeeClaw openclaw: $from_path"
  fi

  return 1
}

main() {
  CANDIDATES=()

  if [ "$#" -eq 0 ]; then
    echo "${PREFIX} 用法: bash openclaw-mac.sh <command> [args...]" >&2
    exit 2
  fi

  local wrapper=""
  if ! wrapper="$(find_wrapper)"; then
    echo "${PREFIX} 错误: 未找到 GeeClaw wrapper。" >&2
    echo "${PREFIX} 不会回退到未知的 system openclaw。" >&2
    echo "${PREFIX} 可检查的常见路径:" >&2
    printf '%s\n' "${CANDIDATES[@]:-}" | sed "s/^/${PREFIX}   - /" >&2
    exit 1
  fi

  debug_log "最终执行: ${wrapper} $*"
  exec "$wrapper" "$@"
}

main "$@"
