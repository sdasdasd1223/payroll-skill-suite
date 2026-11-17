"""將私有 GitHub 診斷事件驗證後保存到公司本機；沒有雲端寫入或刪除能力。"""
from __future__ import annotations
import argparse
import base64
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import uuid

from common import canonical, decrypt, digest, load_secret, lock, organization, private_dir, read_json, verify_sender

OWNER = 'sdasdasd1223'
REPOSITORY = OWNER + '/payroll-skill-diagnostics-private'
EVENT_PATH = re.compile(r'events/([0-9a-f]{2})/([0-9a-f-]{36})\.json\Z')
SHA = re.compile(r'[0-9a-f]{40}\Z')
MAX_BLOB = 65536


def stamp():
    return dt.datetime.now(dt.timezone.utc).isoformat()


def git_sha(raw):
    return hashlib.sha1(b'blob ' + str(len(raw)).encode() + b'\0' + raw).hexdigest()


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise RuntimeError('拒絕轉送管理員憑據')


class GitHubReader:
    """憑據僅在記憶體；此類別只提供固定主機的 GET。"""
    def __init__(self):
        env = dict(os.environ, GIT_TERMINAL_PROMPT='0', GCM_INTERACTIVE='Never')
        result = subprocess.run(
            ['git', 'credential', 'fill'], input='protocol=https\nhost=github.com\n\n',
            capture_output=True, text=True, env=env, timeout=25,
            creationflags=0x08000000 if os.name == 'nt' else 0)
        values = dict(line.split('=', 1) for line in result.stdout.splitlines() if '=' in line)
        self.token = values.get('password')
        if result.returncode or not self.token:
            raise RuntimeError('管理員登入不可用')
        self.opener = urllib.request.build_opener(NoRedirect)

    def get(self, suffix):
        if suffix != '/user' and not suffix.startswith('/repos/' + REPOSITORY):
            raise ValueError('只允許讀取指定診斷倉庫')
        request = urllib.request.Request('https://api.github.com' + suffix, headers={
            'Authorization': 'Bearer ' + self.token,
            'Accept': 'application/vnd.github+json',
            'X-GitHub-Api-Version': '2026-03-10',
            'User-Agent': 'company-payroll-local-archive'})
        with self.opener.open(request, timeout=25) as response:
            raw = response.read(16 * 1024 * 1024 + 1)
        if len(raw) > 16 * 1024 * 1024:
            raise ValueError('雲端索引超出本版處理上限')
        return json.loads(raw)

    def snapshot(self):
        if self.get('/user').get('login') != OWNER:
            raise ValueError('不是公司管理員帳戶')
        base = '/repos/' + REPOSITORY
        repo = self.get(base)
        if repo.get('full_name') != REPOSITORY or repo.get('private') is not True:
            raise ValueError('診斷倉庫不是預期的私有倉庫')
        commit = self.get(base + '/git/ref/heads/main')['object']['sha']
        if not SHA.fullmatch(commit):
            raise ValueError('雲端版本格式錯誤')
        tree_sha = self.get(base + '/git/commits/' + commit)['tree']['sha']
        if not SHA.fullmatch(tree_sha):
            raise ValueError('雲端目錄版本格式錯誤')
        tree = self.get(base + '/git/trees/' + tree_sha + '?recursive=1')
        if tree.get('truncated'):
            raise ValueError('雲端索引不完整，拒絕宣告歸檔完成')
        entries = []
        for entry in tree['tree']:
            path = entry['path']
            if path.startswith('events/') and entry['type'] != 'tree':
                match = EVENT_PATH.fullmatch(path)
                if not match or entry.get('mode') != '100644' or not SHA.fullmatch(entry['sha']):
                    raise ValueError('異常事件路徑或類型')
                identifier = str(uuid.UUID(match[2]))
                if identifier != match[2] or identifier[:2] != match[1] or entry.get('size', MAX_BLOB + 1) > MAX_BLOB:
                    raise ValueError('異常事件編號或大小')
                entries.append(dict(path=path, sha=entry['sha'], size=entry['size']))
        return commit, sorted(entries, key=lambda x: x['path'])

    def blob(self, sha):
        if not SHA.fullmatch(sha):
            raise ValueError('雲端檔案摘要錯誤')
        item = self.get('/repos/' + REPOSITORY + '/git/blobs/' + sha)
        if item.get('encoding') != 'base64' or item.get('size', MAX_BLOB + 1) > MAX_BLOB:
            raise ValueError('雲端事件過大或格式錯誤')
        raw = base64.b64decode(''.join(item['content'].split()), validate=True)
        if len(raw) > MAX_BLOB or len(raw) != item['size'] or git_sha(raw) != sha:
            raise ValueError('下載摘要不符')
        return raw


def safe_path(root, relative):
    path = root / relative
    if not path.resolve().is_relative_to(root) or path.resolve() == root:
        raise ValueError('歸檔路徑越界')
    for part in [path, *path.parents]:
        if part.is_symlink() or getattr(part, 'is_junction', lambda: False)():
            raise ValueError('歸檔路徑不可經過連結')
        if part == root:
            break
    return path


def atomic_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix='.writing-', dir=path.parent)
    try:
        with os.fdopen(fd, 'wb') as stream:
            stream.write(canonical(value))
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(name, path)
    finally:
        if Path(name).exists():
            Path(name).unlink()


def validate_event(raw, entry, org, registry, pem):
    if len(raw) != entry['size'] or git_sha(raw) != entry['sha']:
        raise ValueError('本機或下載的檔案摘要不符')
    envelope = json.loads(raw)
    sender = envelope['sender_device']
    # 已撤銷設備的歷史簽名仍可核對，不接受未登記設備。
    device = registry.get(sender)
    if not device or envelope.get('key_id') != org['key_id']:
        raise ValueError('來源設備未登記或管理金鑰不符')
    verify_sender(envelope, device['public_key'])
    event = decrypt(envelope, pem)
    identifier = entry['path'].rsplit('/', 1)[1][:-5]
    if event.get('event_id') != identifier or event.get('device_id') != sender:
        raise ValueError('事件身份不一致')


def archive(root, admin_home, reader, reserve_bytes=10 * 2**30):
    original = Path(root).absolute()
    root = original.resolve()
    if root == Path(root.anchor) or root != original:
        raise ValueError('必須指定獨立且不經連結的歸檔資料夾')
    safe_path(root, 'status.json')
    private_dir(root, harden=True)
    with lock(root) as acquired:
        if not acquired:
            raise RuntimeError('已有歸檔正在執行')
        try:
            org = organization(Path(admin_home) / 'organization.json')
            if org['repository'] != REPOSITORY:
                raise ValueError('管理員組織與歸檔倉庫不符')
            registry = read_json(Path(admin_home) / 'devices.json')
            pem = load_secret(Path(admin_home) / 'admin-key.secure')
            commit, entries = reader.snapshot()
            # 空間不足時先停止，保留先前歸檔和雲端資料。
            required = sum(item['size'] for item in entries) + 1024 * 1024
            if shutil.disk_usage(root).free < reserve_bytes + required:
                raise OSError('歸檔磁碟可用空間不足')
            added = 0
            receipts = []
            with tempfile.TemporaryDirectory(prefix='.pending-', dir=root) as temp:
                pending = []
                for entry in entries:
                    target = safe_path(root, entry['path'])
                    if target.exists():
                        raw = target.read_bytes()
                    else:
                        raw = reader.blob(entry['sha'])
                        staged = Path(temp) / Path(entry['path']).name
                        staged.write_bytes(raw)
                        pending.append((staged, target))
                    validate_event(raw, entry, org, registry, pem)
                    receipts.append(dict(path=entry['path'], sha256=digest(raw), bytes=len(raw)))
                # 全部下載及驗證通過後才放進正式事件目錄；已存在內容不覆寫。
                for staged, target in pending:
                    target.parent.mkdir(parents=True, exist_ok=True)
                    with target.open('xb') as stream:
                        stream.write(staged.read_bytes())
                        stream.flush()
                        os.fsync(stream.fileno())
                    if digest(target.read_bytes()) != digest(staged.read_bytes()):
                        raise OSError('落盤後校驗失敗')
                    added += 1
            metadata = {'organization': org, 'registered_devices': registry}
            metadata_raw = canonical(metadata)
            metadata_path = safe_path(root, 'metadata/' + digest(metadata_raw) + '.json')
            if metadata_path.exists() and metadata_path.read_bytes() != metadata_raw:
                raise ValueError('管理員公鑰名冊歸檔損壞')
            if not metadata_path.exists():
                atomic_json(metadata_path, metadata)
            result = {'schema': 'company-payroll-archive-1', 'state': 'success', 'at': stamp(),
                      'repository': REPOSITORY, 'cloud_commit': commit,
                      'verified_events': len(entries), 'new_events': added,
                      'metadata': metadata_path.relative_to(root).as_posix(),
                      'files': receipts, 'cloud_deleted': False,
                      'scope': 'encrypted_diagnostics_only'}
            receipt = safe_path(root, 'receipts/' + dt.datetime.now().strftime('%Y%m%d-%H%M%S-') + uuid.uuid4().hex[:8] + '.json')
            atomic_json(receipt, result)
            atomic_json(safe_path(root, 'last-success.json'), result)
            atomic_json(safe_path(root, 'status.json'), result)
            return result
        except Exception as exc:
            atomic_json(safe_path(root, 'status.json'), {'state': 'failed', 'at': stamp(),
                        'category': type(exc).__name__, 'cloud_deleted': False,
                        'message': '本次歸檔失敗；上次成功紀錄與雲端資料保留，需重試。'})
            raise


def main():
    parser = argparse.ArgumentParser(description='每日公司診斷加密歸檔')
    parser.add_argument('--config', required=True)
    args = parser.parse_args()
    try:
        config = read_json(args.config)
        # 憑據取得失敗也必須更新失敗狀態。
        class LazyReader:
            client = None
            def snapshot(self):
                self.client = GitHubReader()
                return self.client.snapshot()
            def blob(self, sha):
                return self.client.blob(sha)
        result = archive(config['archive_root'], config['admin_home'], LazyReader())
        print(json.dumps({'結果': '歸檔成功', '驗證事件數': result['verified_events'],
                          '新增事件數': result['new_events'], '範圍': '加密診斷紀錄，不含原件及完整聊天',
                          '刪除雲端資料': False}, ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({'結果': '歸檔未完成', '錯誤類別': type(exc).__name__,
                          '說明': '既有歸檔及雲端資料保留；不代表本次同步成功。'}, ensure_ascii=False))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
