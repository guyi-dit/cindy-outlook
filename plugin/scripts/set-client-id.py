#!/usr/bin/env python3
import json
import re
import sys
from pathlib import Path

UUID = re.compile(
    r'^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
)

def main():
    if len(sys.argv) != 2 or not UUID.match(sys.argv[1].strip()):
        print('用法: python3 scripts/set-client-id.py <Azure Application (client) ID>', file=sys.stderr)
        print('Application (client) ID 是 Azure 应用注册 Overview 页上的 GUID。', file=sys.stderr)
        return 2
    client_id = sys.argv[1].strip()
    ghost_path = Path(__file__).resolve().parents[1] / 'ghost.json'
    data = json.loads(ghost_path.read_text(encoding='utf-8'))
    secrets = data.get('network', {}).get('secrets') or []
    if not secrets or 'oauth' not in secrets[0]:
        print('ghost.json 缺少 oauth 配置', file=sys.stderr)
        return 1
    secrets[0]['oauth']['clientId'] = client_id
    ghost_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print('已写入 clientId:', client_id)
    print('下一步：在 Cindy 侧边栏「插件」里刷新本地市场「我的插件」，再打开 Outlook 详情页点「连接账号」。')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
