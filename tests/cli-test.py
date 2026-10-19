"""CLI integration in real subprocesses (Node spawnSync is denied in this sandbox)."""
from pathlib import Path
import json, subprocess, tempfile
ROOT = Path(__file__).resolve().parents[1]
def command(*args):
    p = subprocess.run(args, cwd=ROOT, text=True, capture_output=True)
    assert p.returncode == 0, (args, p.stderr)
    return p.stdout
with tempfile.TemporaryDirectory(prefix='office-cli-') as directory:
    base=Path(directory)
    for kind in ('xlsx','pptx'):
        src=base/(kind+'.bin')
        dest=base/('clean.'+kind)
        code="import {makeOfficeFixture} from './tests/office-fixtures.js';import {writeFileSync} from 'node:fs';writeFileSync(process.argv[1],await makeOfficeFixture(process.argv[2]));"
        command('node','--input-type=module','-e',code,str(src),kind)
        before=json.loads(command('node','cli.js',str(src),'--json'))
        assert before['format']==kind
        cleaned=json.loads(command('node','cli.js',str(src),'--scrub',str(dest),'--only','identity,comments','--json'))
        assert cleaned['after']['metadata']['creator'] is None
        assert cleaned['after']['comments']==[]
        assert len(cleaned['after']['findings'])>0
        after=json.loads(command('node','cli.js',str(dest),'--json'))
        assert after['format']==kind
        print(kind+': renamed input, selected cleanup, output re-analysis PASS')
    bad=base/'unsupported.zip'
    bad.write_bytes(b'not an office file')
    result=subprocess.run(['node','cli.js',str(bad),'--json'],cwd=ROOT,capture_output=True,text=True)
    assert result.returncode != 0 and result.stderr and not result.stdout
    print('Unsupported input exits nonzero: PASS')
