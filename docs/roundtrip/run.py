#!/usr/bin/env python3
"""Reproducible local validation harness, not shipped product code.
python3 docs/roundtrip/run.py prepare|render
Artifacts in ignored .validation/roundtrip. Safe to resume render; results logged per file.
"""
from pathlib import Path
import argparse, hashlib, json, os, re, subprocess, sys, time

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / '.validation' / 'roundtrip'
MANIFEST = Path(__file__).with_name('inputs.json')

def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()

def run(args, timeout=90, env=None):
    try:
        p = subprocess.run(args, cwd=ROOT, capture_output=True, text=True,
                           timeout=timeout, env=env)
        return {'code': p.returncode, 'stdout': p.stdout, 'stderr': p.stderr}
    except subprocess.TimeoutExpired as e:
        return {'code': 'TIMEOUT', 'stdout': str(e.stdout or ''), 'stderr': str(e.stderr or '')}

def norm(s):
    return '\n'.join(line.rstrip(' \t') for line in s.replace('\r\n', '\n').split('\n'))

def prepare():
    rows = []
    inputs = json.loads(MANIFEST.read_text())
    for i, item in enumerate(inputs):
        src = ROOT / item['path']
        if sha(src) != item['sha256']:
            raise ValueError('Input changed: ' + item['path'])
        folder = OUT / f'{i:03d}'
        folder.mkdir(parents=True, exist_ok=True)
        clean = folder / 'input-clean' / src.name
        clean.parent.mkdir(exist_ok=True)
        if clean.exists():
            raise ValueError('Preparation exists; retain it or choose a new run directory')
        log = run(['node', 'cli.js', str(src), '--scrub', str(clean)])
        (folder / 'scrub.json').write_text(json.dumps(log, indent=2))
        row = {**item, 'index': i, 'status': 'PREPARED' if log['code'] == 0 and clean.exists() else 'SCRUB_FAILED'}
        if clean.exists():
            row['clean_sha256'] = sha(clean)
            row['clean_path'] = str(clean.relative_to(ROOT))
        rows.append(row)
    (OUT / 'prepared.json').write_text(json.dumps(rows, indent=2) + '\n')
    (OUT / 'preparation_environment.json').write_text(json.dumps({
        'commit': run(['git','rev-parse','HEAD']), 'prepared_at': time.time(),
        'source_sha256': {str(p.relative_to(ROOT)): sha(p) for p in sorted((ROOT/'src').glob('*.js'))},
        'harness_sha256': sha(Path(__file__)),
    }, indent=2) + '\n')
    print(json.dumps({s: sum(r['status'] == s for r in rows) for s in sorted({r['status'] for r in rows})}))

def render():
    rows = json.loads((OUT / 'prepared.json').read_text())
    env = dict(os.environ, DCONF_PROFILE='/dev/null', SAL_USE_VCLPLUGIN='svp')
    (OUT / 'runtime').mkdir(exist_ok=True)
    env['XDG_RUNTIME_DIR'] = str(OUT / 'runtime')
    (OUT / 'environment.json').write_text(json.dumps({'soffice': run(['soffice', '--version'], env=env), 'started': time.time()}, indent=2))
    results = []
    for row in rows:
        folder = OUT / f"{row['index']:03d}"
        result_file = folder / 'render.json'
        if result_file.exists():
            results.append(json.loads(result_file.read_text()))
            continue
        result = dict(row)
        if row['status'] == 'PREPARED':
            srcs = [ROOT / row['path'], ROOT / row['clean_path'] if 'clean_path' in row else folder / 'clean.docx']
            assert sha(srcs[0]) == row['sha256'] and sha(srcs[1]) == row['clean_sha256']
            captures = []
            for label, src in zip(('original', 'cleaned'), srcs):
                dest = folder / label
                dest.mkdir(exist_ok=True)
                profile = folder / (label + '-profile')
                profile.mkdir(exist_ok=True)
                # Disable macro execution and automatic external-link updates in this fresh profile.
                user = profile / 'user'
                user.mkdir(exist_ok=True)
                (user / 'registrymodifications.xcu').write_text('''<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item><item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>2</value></prop></item></oor:items>''')
                log = run(['soffice', '-env:UserInstallation=' + profile.as_uri(), '--headless', '--convert-to', 'pdf', '--outdir', str(dest), str(src)], env=env)
                pdf = dest / (src.stem + '.pdf')
                cap = {'conversion': log, 'pages': None, 'text': None}
                if log['code'] == 0 and pdf.exists():
                    info = run(['pdfinfo', str(pdf)])
                    txt = run(['pdftotext', '-layout', '-enc', 'UTF-8', str(pdf), '-'])
                    cap['pdfinfo'] = info
                    cap['text_log'] = {'code': txt['code'], 'stderr': txt['stderr']}
                    match = re.search(r'^Pages:\s+(\d+)', info['stdout'], re.M)
                    if info['code'] == 0 and match and int(match[1]) > 0 and txt['code'] == 0:
                        cap['pages'] = int(match[1])
                        cap['text'] = txt['stdout']
                        (dest / 'rendered.txt').write_text(txt['stdout'])
                captures.append(cap)
            result['captures'] = captures
            if any(c['text'] is None for c in captures):
                result['status'] = 'CONVERSION_FAILED'
            elif captures[0]['pages'] == captures[1]['pages'] and norm(captures[0]['text']) == norm(captures[1]['text']):
                result['status'] = 'PASS'
            else:
                result['status'] = 'REVIEW_REQUIRED'
        result_file.write_text(json.dumps(result, indent=2) + '\n')
        results.append(result)
        print(row['path'], result['status'], flush=True)
    (OUT / 'results.json').write_text(json.dumps(results, indent=2) + '\n')
    print(json.dumps({s: sum(r['status'] == s for r in results) for s in sorted({r['status'] for r in results})}))

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('mode', choices=['prepare','render'])
    parser.add_argument('--out', default='.validation/roundtrip')
    args=parser.parse_args()
    OUT=(ROOT/args.out).resolve()
    OUT.mkdir(parents=True, exist_ok=True)
    {'prepare': prepare, 'render': render}[args.mode]()
