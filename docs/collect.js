// 홈페이지 그림 모으기 전용 작업 공간 (모듈 워커). 공고 분석(worker.js)과 따로 둔다:
// 공고 PDF 를 읽은 Python 은 메모리가 커서 그 뒤에 그림 부품(Pillow)이 멈추는 일이 있다.
// 쪽은 Python 이 동기로, 그림은 여기서 6장씩 동시에 받아 바로 Python 에 넘긴다 (쌓아 두지 않음).
import { loadPyodide } from 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/pyodide.js';
const VERSION = '2026.10.18b';
const PYODIDE = 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/';
const PILLOW = 'https://cdn.jsdelivr.net/pyodide/v0.28.0a3/full/pillow-11.2.1-cp313-cp313-pyodide_2025_0_wasm32.whl';
// 브라우저는 다른 사이트를 직접 못 읽어서 Cloudflare 중계(proxy/worker.js)를 거친다. 이 PC 시험 때는 wrangler dev
const LOCAL = ['127.0.0.1', 'localhost'].includes(location.hostname);
const PROXY = LOCAL ? 'http://127.0.0.1:8787' : 'https://bunyang-proxy.bunyang-overview.workers.dev';
const PY_FILES = ['collect.py', 'engine/__init__.py', 'engine/webcollect.py', 'engine/imgcut.py'];
const APP = '/home/pyodide/app';
const NL = String.fromCharCode(10);

let pyodide = null;
const status = msg => postMessage({status: msg});
const log = msg => postMessage({log: msg});

async function init() {
  status('그림 모으기 준비 중… (처음 한 번은 조금 걸립니다)');
  pyodide = await loadPyodide({indexURL: PYODIDE});
  await pyodide.loadPackage(PILLOW);
  pyodide.FS.mkdirTree(APP + '/engine');
  for (const f of PY_FILES) {
    const r = await fetch(`py/${f}?v=${VERSION}`);
    if (!r.ok) throw new Error(`py/${f} 을(를) 받지 못했습니다 (${r.status})`);
    pyodide.FS.writeFile(`${APP}/${f}`, await r.text());
  }
  await pyodide.runPythonAsync(['import sys', `sys.path.insert(0, '${APP}')`, 'import collect'].join(NL));
  const m = pyodide.globals.get('collect');
  m.set_logger(log);
  m.destroy();
  status('');
}
const ready = init();

function call(fn, ...args) {
  const m = pyodide.globals.get('collect');
  try {
    const f = m[fn];
    const out = f(...args);
    f.destroy && f.destroy();
    return out;
  } finally { m.destroy(); }
}
function bytes(fn, ...args) {
  const py = call(fn, ...args);
  const out = py.toJs();
  py.destroy && py.destroy();
  return out;
}
async function pool(items, n, fn) {
  let k = 0;
  const run = async () => { while (k < items.length) { const it = items[k++]; await fn(it); } };
  await Promise.all(Array.from({length: Math.min(n, items.length)}, run));
}

async function collect(home, draft) {
  const r1 = JSON.parse(call('collect_pages', home, PROXY, JSON.stringify(draft)));
  if (!r1.ok) return r1;
  let done = 0, failed = false;
  await pool(r1.items, 6, async it => {
    try {
      const res = await fetch(`${PROXY}?url=${encodeURIComponent(it.src)}&ref=${encodeURIComponent(it.ref)}`);
      if (res.ok) {
        const b = new Uint8Array(await res.arrayBuffer());
        if (b.length >= 15000) call('collect_add', it.i, b);
      }
    } catch (err) {
      if (!failed) { failed = true; log(`  그림 처리 실패 (${it.src.split('/').pop()}): ${String(err && err.message || err).slice(-200)}`); }
    }
    if (++done % 40 === 0) log(`  그림 ${done}/${r1.items.length}장 받음`);
  });
  const r2 = JSON.parse(call('collect_finish'));
  if (r2.ok) r2.list = r2.cands.map(c => ({...c, data: bytes('collect_image', c.id)}));
  return r2;
}

onmessage = async (e) => {
  const {id, args} = e.data;
  try {
    await ready;
    const result = await collect(args.home, args.draft);
    postMessage({id, result}, (result.list || []).map(c => c.data.buffer));
  } catch (err) {
    postMessage({id, result: {ok: false, error: '그림 모으기 중 오류: ' + String(err && err.message || err).split(NL).slice(-2).join(' ')}});
  }
};
