import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
// Static release contract; actual browser enforcement is checked separately.
const html = readFileSync(new URL('../dist/doc-rinse.html', import.meta.url), 'utf8');
const csp = html.match(/<meta http-equiv="Content-Security-Policy" content="([^"]+)"/i)?.[1];
assert.ok(csp, 'built artifact must contain CSP');
for (const directive of ["default-src 'none'", "connect-src 'none'", "object-src 'none'", "base-uri 'none'", "form-action 'none'"])
  assert.ok(csp.split(';').map(s => s.trim()).includes(directive), directive);
assert.ok(html.indexOf('Content-Security-Policy') < html.indexOf('<script'));
const script = html.match(/<script[^>]*>([\s\S]*)<\/script>/)[1];
assert.doesNotMatch(script, /\bfetch\s*\(|\bXMLHttpRequest\b|\bWebSocket\b|\bsendBeacon\s*\(/);
assert.doesNotMatch(script, /get\(['"]sample['"]\)|dataset\.url/);
assert.match(html, /<title>Doc Rinse/);
assert.match(html, /Copyright \(c\) 2026 mothwake/);
assert.match(html, /Permission is hereby granted, free of charge/);
assert.doesNotMatch(html, /Palimpsest/);
assert.match(script, /data-sample/);
const samples = JSON.parse(script.match(/const SAMPLES = (\[[^\n]*\]);/)[1]);
assert.deepEqual(samples.map(s => s.file), ['00-synthetic-everything.docx']);
assert.equal(samples[0].b64, readFileSync(new URL('../samples/00-synthetic-everything.docx', import.meta.url)).toString('base64'));
console.log('Release contract passes: name, embedded samples, CSP, no URL-loading APIs');
