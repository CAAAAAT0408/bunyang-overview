// 분양개요 자동작성 — 분양 홈페이지의 쪽(html)·그림을 대신 받아 주는 중계 (Cloudflare Worker)
// 웹 버전(GitHub 페이지)은 브라우저 보안 때문에 다른 사이트를 직접 읽지 못해서 이 중계를 거친다.
//   GET /?url=<주소>&ref=<보낸 쪽 주소>  →  그 주소의 내용 (html·그림만, 15MB 까지)
// 아무나 쓰는 열린 중계가 되지 않도록 웹 버전 주소(Origin)에서 부를 때만 응답한다.
const ALLOW_ORIGINS = ['https://caaaaat0408.github.io', 'http://127.0.0.1:8799', 'http://localhost:8799'];
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36';
const MAX = 15 * 1024 * 1024;
const OK_TYPE = /^(text\/html|application\/xhtml\+xml|text\/plain|image\/)/i;

export default {
  async fetch(req) {
    const origin = req.headers.get('Origin') || '';
    const cors = {
      'Access-Control-Allow-Origin': ALLOW_ORIGINS.includes(origin) ? origin : ALLOW_ORIGINS[0],
      'Access-Control-Expose-Headers': 'X-Final-Url, Content-Type',
      'Vary': 'Origin',
    };
    const fail = (status, msg) => new Response(msg, {status, headers: {...cors, 'Content-Type': 'text/plain; charset=utf-8'}});
    if (req.method === 'OPTIONS') {
      return new Response(null, {headers: {...cors, 'Access-Control-Allow-Methods': 'GET', 'Access-Control-Max-Age': '86400'}});
    }
    if (req.method !== 'GET') return fail(405, 'GET only');
    if (!ALLOW_ORIGINS.includes(origin)) return fail(403, 'forbidden');
    const q = new URL(req.url).searchParams;
    let target;
    try { target = new URL(q.get('url') || ''); } catch (e) { return fail(400, 'bad url'); }
    if (!/^https?:$/.test(target.protocol)) return fail(400, 'http(s) only');
    const headers = {'User-Agent': UA, 'Accept-Language': 'ko-KR,ko;q=0.9'};
    const ref = q.get('ref');
    if (ref && /^https?:\/\//i.test(ref)) headers['Referer'] = ref;
    let r;
    try {
      r = await fetch(target.toString(), {headers, redirect: 'follow', cf: {cacheTtl: 3600, cacheEverything: true}});
    } catch (e) {
      return fail(502, 'fetch failed: ' + e);
    }
    const ct = r.headers.get('Content-Type') || '';
    if (ct && !OK_TYPE.test(ct)) return fail(415, 'type not allowed: ' + ct);
    if (Number(r.headers.get('Content-Length') || 0) > MAX) return fail(413, 'too large');
    return new Response(r.body, {
      status: r.status,
      headers: {...cors, 'Content-Type': ct || 'application/octet-stream', 'X-Final-Url': r.url, 'Cache-Control': 'public, max-age=3600'},
    });
  },
};
