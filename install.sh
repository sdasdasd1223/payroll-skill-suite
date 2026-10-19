#!/usr/bin/env bash
set -eo pipefail

# Compatible with macOS default Bash 3.2.
if [[ $# -gt 1 ]]; then
  printf '用法: bash install.sh [CODEX_HOME 路径]\n' >&2
  exit 2
fi
package_root="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
source_root="$package_root/skills"
codex_home="${1:-${CODEX_HOME:-$HOME/.codex}}"
skills_root="$codex_home/skills"
skill_names=(payroll-operations attendance-calculator attendance-settlement social-insurance-fund)

validate_skill() {
  local root="$1" name="$2" file agent field
  file="$root/$name/SKILL.md"
  agent="$root/$name/agents/openai.yaml"
  if [[ ! -f "$file" || ! -f "$agent" ]]; then
    printf '安装包缺少 %s 的 SKILL.md 或 agents/openai.yaml\n' "$name" >&2
    return 1
  fi
  if ! grep -Eq "^name:[[:space:]]*$name[[:space:]]*$" "$file" ||
     ! grep -Eq '^description:[[:space:]]*.+$' "$file"; then
    printf '%s 的 SKILL.md 信息不完整\n' "$name" >&2
    return 1
  fi
  for field in display_name short_description default_prompt; do
    if ! grep -Eq "^[[:space:]]*$field:[[:space:]]*.+$" "$agent"; then
      printf '%s 的 agents/openai.yaml 缺少 %s\n' "$name" "$field" >&2
      return 1
    fi
  done
}

printf '\n==> 检查安装包\n'
for name in "${skill_names[@]}"; do validate_skill "$source_root" "$name"; done
mkdir -p "$skills_root" "$codex_home/skill-backups"
stage_dir="$(mktemp -d "$skills_root/.payroll-stage.XXXXXX")"
trap 'rm -rf -- "$stage_dir"' EXIT
for name in "${skill_names[@]}"; do
  cp -R "$source_root/$name" "$stage_dir/$name"
  validate_skill "$stage_dir" "$name"
done

existing=()
for name in "${skill_names[@]}"; do
  if [[ -e "$skills_root/$name" ]]; then existing+=("$name"); fi
done
backup_root=''
if [[ ${#existing[@]} -gt 0 ]]; then
  backup_root="$(mktemp -d "$codex_home/skill-backups/payroll-skill-suite_$(date +%Y%m%d_%H%M%S).XXXXXX")"
  printf '\n==> 备份电脑上的旧 Skill\n'
  for name in "${existing[@]}"; do
    cp -R "$skills_root/$name" "$backup_root/$name"
    printf '已备份：%s\n' "$name"
  done
fi

installed=()
rollback() {
  local item
  for item in "${installed[@]}"; do
    rm -rf -- "$skills_root/$item"
    if [[ -e "$stage_dir/.old-$item" ]]; then
      mv "$stage_dir/.old-$item" "$skills_root/$item"
    fi
  done
}

printf '\n==> 安装工资核算 Skill 套件\n'
for name in "${skill_names[@]}"; do
  if [[ -e "$skills_root/$name" ]]; then
    mv "$skills_root/$name" "$stage_dir/.old-$name"
  fi
  if ! mv "$stage_dir/$name" "$skills_root/$name"; then
    if [[ -e "$stage_dir/.old-$name" ]]; then
      mv "$stage_dir/.old-$name" "$skills_root/$name"
    fi
    rollback
    printf '安装中断，已尝试恢复旧版本。备份位置：%s\n' "$backup_root" >&2
    exit 1
  fi
  installed+=("$name")
  printf '已安装：%s\n' "$name"
done

printf '\n==> 核对安装结果\n'
for name in "${skill_names[@]}"; do
  if ! validate_skill "$skills_root" "$name"; then
    rollback
    printf '核对失败，已尝试恢复旧版本。备份位置：%s\n' "$backup_root" >&2
    exit 1
  fi
done
printf '\n安装成功。\n'
if [[ -n "$backup_root" ]]; then printf '旧版本备份位置：%s\n' "$backup_root"; fi
printf '请完全退出并重新启动 Codex，重启后 Skill 才会生效。\n'
