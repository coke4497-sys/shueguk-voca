/* paper.html 브라우저 검사 — 종합·주차별 시험지·개념 정리
   실행: NODE_PATH=$(npm root -g) node tools/paper-e2e.js  (저장소 루트에서) */
const { chromium } = require('playwright');
const http = require('http'), fs = require('fs'), path = require('path');
const ROOT = path.resolve(__dirname, '..');
let pass = 0, fail = 0;
const ok = (c, m) => { if (c) { pass++; } else { fail++; console.log('FAIL', m); } };
const srv = http.createServer((q, r) => {
  const f = path.join(ROOT, decodeURIComponent(q.url.split('?')[0]));
  if (!f.startsWith(ROOT) || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { r.writeHead(404); return r.end(); }
  r.writeHead(200, { 'Content-Type': f.endsWith('.json') ? 'application/json' : 'text/html; charset=utf-8' });
  fs.createReadStream(f).pipe(r);
}).listen(0);
(async () => {
  const port = srv.address().port;
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.route(/supabase|script\.google|fonts\.g/, r => r.abort());
  await p.goto(`http://127.0.0.1:${port}/paper.html`);
  const n = (sel) => p.locator(sel).count();
  const pick = async (g, f, t) => {
    await p.selectOption('#selGrade', g);
    await p.selectOption('#selFrom', String(f));
    await p.selectOption('#selTo', String(t));
  };

  // ① 종합 시험지 그대로
  ok(await p.locator('.kind.on').getAttribute('data-kind') === 'mix', '기본은 종합');
  ok(await p.isVisible('#fldMode') && await p.isVisible('#fldExtra'), '종합: 문항 구성·단답 칸 보임');
  await pick('h1', 1, 2);
  await p.click('#btnBuild');
  await p.waitForSelector('#sheet .qwrap');
  const d1 = JSON.parse(fs.readFileSync(path.join(ROOT, 'data/h1-1.json')));
  const d2 = JSON.parse(fs.readFileSync(path.join(ROOT, 'data/h1-2.json')));
  ok(await n('#sheet .q') === d1.questions.length + d2.questions.length + 10, '종합: 문항 수(+단답 10)');
  ok(await p.isVisible('#hwPanel'), '종합: H WORK 패널');
  ok(await n('.wk-page') === 0, '종합: 주차 페이지 없음');

  // ② 주차별 시험지
  await p.click('.kind[data-kind=weekly]');
  ok(!(await p.isVisible('#sheet')), '종류를 바꾸면 미리보기 비움');
  ok(!(await p.isVisible('#fldMode')) && !(await p.isVisible('#fldExtra')) && !(await p.isVisible('#fldSeed')), '주차별: 출제 옵션 숨김');
  ok((await p.textContent('#btnBuild')).includes('주차별'), '주차별: 버튼 이름');
  await pick('h1', 1, 3);
  await p.click('#btnBuild');
  await p.waitForSelector('.wk-page');
  ok(await n('.wk-page') === 3, '주차별: 3장');
  const d3 = JSON.parse(fs.readFileSync(path.join(ROOT, 'data/h1-3.json')));
  const counts = await p.$$eval('.wk-page', ps => ps.map(x => x.querySelectorAll('.q').length));
  ok(JSON.stringify(counts) === JSON.stringify([d1, d2, d3].map(d => d.questions.length)), '주차별: 주차마다 온라인 문항 그대로 ' + counts);
  const firstNums = await p.$$eval('.wk-page', ps => ps.map(x => x.querySelector('.qn').textContent));
  ok(firstNums.every(t => t === '01.'), '주차별: 번호가 주차마다 01부터');
  ok((await p.textContent('.wk-page .title')).includes('1주차 테스트'), '주차별: 제목');
  ok(await n('.answer-sheet .ans-block') === 3, '주차별: 정답 3묶음');
  const firstAns = await p.$$eval('.ans-block', bs => bs[0].querySelectorAll('td:not(.no)').length);
  ok(firstAns === d1.questions.length, '주차별: 1주차 정답 수');
  ok(!(await p.isVisible('#hwPanel')), '주차별: H WORK 패널 없음');
  ok(await p.isVisible('#btnDlQ') && await p.isVisible('#btnDlA') && !(await p.isVisible('#btnDlS')), '주차별: 문제지/정답지 버튼');
  // 인쇄 모양: 문제지만 = 3페이지 이상, 정답지에는 wk-page 숨김
  await p.emulateMedia({ media: 'print' });
  await p.evaluate(() => document.body.classList.add('print-a'));
  ok(await p.$$eval('.wk-page', ps => ps.every(x => getComputedStyle(x).display === 'none')), '정답지 인쇄: 문제 페이지 숨김');
  await p.evaluate(() => document.body.classList.replace('print-a', 'print-q'));
  ok(await p.$eval('.answer-sheet', x => getComputedStyle(x).display === 'none'), '문제지 인쇄: 정답 숨김');
  ok(await p.$$eval('.wk-page', ps => getComputedStyle(ps[1]).breakBefore === 'page'), '문제지 인쇄: 주차마다 새 페이지');
  const pdfQ = await p.pdf({ format: 'A4' });
  const pagesQ = (pdfQ.toString('latin1').match(/\/Type\s*\/Page[^s]/g) || []).length;
  ok(pagesQ >= 3, '문제지 PDF ' + pagesQ + '쪽');
  await p.evaluate(() => document.body.classList.remove('print-q'));
  await p.emulateMedia({ media: 'screen' });
  // 한 주차만
  await pick('h1', 5, 5);
  await p.click('#btnBuild');
  await p.waitForFunction(() => document.querySelectorAll('.wk-page').length === 1);
  ok((await p.textContent('#status')).includes('시험지 1장'), '주차별: 한 주차 상태 문구');

  // ③ 개념 정리
  await p.click('.kind[data-kind=study]');
  ok(!(await p.isVisible('#fldExpl')), '개념 정리: 해설 칸 숨김');
  await pick('m1', 3, 4);
  await p.click('#btnBuild');
  await p.waitForSelector('.st-e');
  const m3 = JSON.parse(fs.readFileSync(path.join(ROOT, 'data/m1-3.json')));
  const m4 = JSON.parse(fs.readFileSync(path.join(ROOT, 'data/m1-4.json')));
  const cnt = d => ['vocab', 'concept', 'idiom'].reduce((a, k) => a + ((d.study || {})[k] || []).length, 0);
  ok(await n('.wk-page') === 2, '개념 정리: 2장');
  ok(await n('.st-e') === cnt(m3) + cnt(m4), '개념 정리: 어휘 수');
  const cats = await p.$$eval('.wk-page:first-child .st-cat', xs => xs.map(x => x.firstChild.textContent));
  ok(cats[0] === '필수 어휘', '개념 정리: 순서 ' + cats);
  ok((await p.textContent('.st-e .st-w')).includes(m3.study.vocab[0].word), '개념 정리: 첫 어휘');
  ok(await p.isVisible('#btnDlS') && !(await p.isVisible('#btnDlQ')), '개념 정리: 다운로드 버튼');
  ok(await n('.q') === 0 && await n('.answer-sheet') === 0, '개념 정리: 문항·정답 없음');

  // ④ 최근 만든 시험 → 다시 열기가 종류까지 되살림
  const items = await p.$$eval('.hist-item .h-name', xs => xs.map(x => x.textContent));
  ok(items[0].includes('개념 정리') && items.some(t => t.includes('주차별 시험지')) && items.some(t => t.includes('종합')), '기록 3종 ' + items);
  const wi = items.findIndex(t => t.includes('주차별'));
  await p.click(`.hist-item .h-open[data-i="${wi}"]`);
  await p.waitForFunction(() => document.querySelector('.kind.on').dataset.kind === 'weekly' && document.querySelectorAll('.wk-page .q').length);
  ok(await n('.wk-page') === 1, '다시 열기: 주차별 5주차');
  // 파일 이름
  const title = await p.evaluate(() => { let t; window.print = () => { t = document.title; }; document.getElementById('btnDlQ').click(); return t; });
  ok(title === '슈국어휘_고1_5주_주차별_문제지', '파일 이름 ' + title);

  await b.close(); srv.close();
  console.log(`통과 ${pass} · 실패 ${fail}`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
