// 브라우저 안에서 Python(Pyodide)을 띄워 분양개요 엔진을 돌린다 (모듈 워커).
// 공고·양식·그림 파일은 이 브라우저 밖으로 나가지 않는다.
import { loadPyodide } from 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/pyodide.js';
const VERSION = '2026.10.23a';
// PyMuPDF(공고 PDF 읽기)가 들어 있는 Pyodide 묶음 (Python 3.13)
const PYODIDE = 'https://cdn.jsdelivr.net/npm/@bentopdf/pymupdf-wasm@0.11.16/assets/';
const PYMUPDF = PYODIDE + 'pymupdf-1.26.3-cp313-none-pyodide_2025_0_wasm32.whl';
const PY_FILES = ['xlsxio.py', 'build_web.py', 'bridge.py',
  'engine/__init__.py', 'engine/pdfdoc.py', 'engine/notice.py', 'engine/options.py', 'engine/draft.py'];
const APP = '/home/pyodide/app';
const IMG = '/home/pyodide/img';


let pyodide = null;
const status = msg => postMessage({status: msg});
const log = msg => postMessage({log: msg});

async function fetchText(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`${url} 을(를) 받지 못했습니다 (${r.status})`);
  return await r.text();
}

async function init() {
  status('Python 준비 중… (처음 한 번은 30초~1분, 그 뒤로는 빨라집니다)');
  pyodide = await loadPyodide({indexURL: PYODIDE});
  status('PDF 읽기 부품 받는 중… (약 25MB, 처음 한 번만)');
  await pyodide.loadPackage(['lxml', PYMUPDF]);
  status('엔진 받는 중…');
  pyodide.FS.mkdirTree(APP + '/engine');
  pyodide.FS.mkdirTree(IMG);
  for (const f of PY_FILES) pyodide.FS.writeFile(`${APP}/${f}`, await fetchText(`py/${f}?v=${VERSION}`));
  await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, '${APP}')
import bridge
`);
  const bridge = pyodide.globals.get('bridge');
  bridge.set_logger(log);
  bridge.destroy();
  status('');
  postMessage({ready: true});
}
const ready = init().catch(e => { postMessage({fatal: String(e && e.message || e)}); throw e; });

function call(fn, ...args) {
  const bridge = pyodide.globals.get('bridge');
  try {
    const f = bridge[fn];
    const out = f(...args);
    f.destroy && f.destroy();
    return out;
  } finally { bridge.destroy(); }
}
function bytes(fn, ...args) {
  const py = call(fn, ...args);
  const out = py.toJs();
  py.destroy && py.destroy();
  return out;
}

onmessage = async (e) => {
  const {id, cmd, args} = e.data;
  try {
    await ready;
    let result, transfer = [];
    if (cmd === 'analyze') {
      result = JSON.parse(call('analyze', args.notice, args.nname, args.template, args.tname));
    } else if (cmd === 'pdf_images') {
      const list = JSON.parse(call('pdf_images'));
      result = {ok: true, list: list.map(m => ({...m, data: bytes('pdf_image', m.id)}))};
      transfer = result.list.map(m => m.data.buffer);
    } else if (cmd === 'page') {
      const n = call('page_count');
      const p = Math.max(1, Math.min(n, args.page || 1));
      const data = bytes('render_page', p - 1, args.zoom || 2);
      result = {ok: true, page: p, pages: n, data};
      transfer = [data.buffer];
    } else if (cmd === 'build') {
      for (const im of args.images || []) pyodide.FS.writeFile(`${IMG}/${im.id}`, new Uint8Array(im.data));
      result = JSON.parse(call('build', JSON.stringify(args.draft)));
      if (result.ok) { result.xlsx = bytes('result'); transfer = [result.xlsx.buffer]; }
    } else {
      throw new Error('알 수 없는 명령: ' + cmd);
    }
    postMessage({id, result}, transfer);
  } catch (err) {
    postMessage({id, result: {ok: false, error: '처리 중 오류가 났습니다: ' + String(err && err.message || err).split('\n').slice(-3).join(' ')}});
  }
};
