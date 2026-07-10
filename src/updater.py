from __future__ import annotations
import json, os, re, subprocess, sys, tempfile, urllib.request
from pathlib import Path


def _tuple(v: str):
    nums=[int(x) for x in re.findall(r'\d+', v)[:3]]
    return tuple((nums+[0,0,0])[:3])


def config_path(base_dir: Path) -> Path:
    return base_dir / 'update_config.json'


def check_for_update(current_version: str, base_dir: Path):
    p=config_path(base_dir)
    if not p.exists(): return None
    cfg=json.loads(p.read_text(encoding='utf-8'))
    repo=str(cfg.get('github_repo','')).strip()
    if not repo or repo.startswith('OWNER/'):
        return None
    url=f'https://api.github.com/repos/{repo}/releases/latest'
    req=urllib.request.Request(url,headers={'User-Agent':'Analitika-Updater'})
    with urllib.request.urlopen(req,timeout=8) as r:
        data=json.loads(r.read().decode('utf-8'))
    latest=data.get('tag_name') or data.get('name') or ''
    if _tuple(latest) <= _tuple(current_version): return None
    asset=None
    for item in data.get('assets',[]):
        name=str(item.get('name',''))
        if name.lower().endswith('.exe') and 'setup' in name.lower():
            asset=item; break
    return {'version':latest,'notes':data.get('body',''),'download_url':asset.get('browser_download_url') if asset else None,'page_url':data.get('html_url')}


def download_and_launch(url: str):
    target=Path(tempfile.gettempdir())/'Analitika_Update_Setup.exe'
    req=urllib.request.Request(url,headers={'User-Agent':'Analitika-Updater'})
    with urllib.request.urlopen(req,timeout=60) as r, target.open('wb') as f:
        f.write(r.read())
    if sys.platform.startswith('win'):
        os.startfile(str(target))  # type: ignore[attr-defined]
    else:
        subprocess.Popen([str(target)])
    return target
