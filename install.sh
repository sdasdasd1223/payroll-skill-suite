#!/usr/bin/env bash
set -eo pipefail
# 兼容macOS Bash 3.2。程序、凭据和用户记录分开，安装事务由同一Python入口执行。
package_root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
python_bin=''
for name in python python3; do
  candidate="$(command -v "$name" 2>/dev/null || true)"
  [[ -n "$candidate" && "$candidate" != *WindowsApps* ]] || continue
  if "$candidate" -c 'import sys;sys.exit(0 if sys.version_info >= (3,10) else 2)' >/dev/null 2>&1; then python_bin="$candidate"; break; fi
done
if [[ -z "$python_bin" ]]; then printf '未找到Python，请让Codex配置Python 3.10或以上。\n' >&2; exit 2; fi
args=()
if [[ $# -gt 0 && "$1" != --* ]]; then args+=(--codex-home "$1"); shift; fi
while [[ $# -gt 0 ]]; do
  case "$1" in
    --skip-diagnostic-schedule|--check-only) args+=("$1"); shift ;;
    --device-profile) [[ $# -ge 2 ]] || exit 2; args+=("$1" "$2"); shift 2 ;;
    *) printf '未知安装参数。\n' >&2; exit 2 ;;
  esac
done
export PYTHONUTF8=1
"$python_bin" "$package_root/install_suite.py" "${args[@]}"
