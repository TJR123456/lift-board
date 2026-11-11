// Saves PNG screenshots of board sections into ./screenshots (2x resolution).
// Needs the board server running (http://localhost:8765) and Google Chrome.
// Run:  node --experimental-websocket shots.mjs
import { spawn } from 'node:child_process';
import { mkdirSync, writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const OUT = join(HERE, 'screenshots');
const URL_ = 'http://localhost:8765/';
const CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const PORT = 9333;

// name, setup (runs in the page), selector to frame
const SHOTS = [
  ['1-move-up', `showView('board')`, 'header, #ready'],
  ['2-lift-detail', `showView('board'); openDetail(DATA.lifts.find((l) => l.current && l.key === 'ez-bar cable curl').id)`, '#detail'],
  ['3-day-push-tricep', `showView('board')`, '#routine > .day-block:nth-child(1)'],
  ['4-week-by-muscle', `showView('board')`, '#week-muscles'],
  ['5-volume-chart', `showView('volume')`, '#vol-chart'],
  ['6-muscle-heatmap', `showView('volume')`, '#vol-muscles'],
  ['7-volume-by-lift', `showView('volume')`, '#vol-workouts'],
  ['8-death-sets', `showView('volume')`, '#vol-death'],
];

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const chrome = spawn(CHROME, ['--headless=new', `--remote-debugging-port=${PORT}`, `--user-data-dir=${mkdtempSync(join(tmpdir(), 'liftshots-'))}`,
  '--hide-scrollbars', '--window-size=1280,900', 'about:blank'], { stdio: 'ignore' });

let ws, id = 0;
const pending = new Map();
const send = (method, params = {}) => new Promise((resolve, reject) => {
  const n = ++id;
  pending.set(n, { resolve, reject });
  ws.send(JSON.stringify({ id: n, method, params }));
});
const evaluate = async (expression) => {
  const r = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
  if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text);
  return r.result.value;
};

try {
  let target;
  for (let i = 0; i < 50 && !target; i++) {
    await sleep(200);
    try { target = (await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json()).find((t) => t.type === 'page'); } catch {}
  }
  if (!target) throw new Error('Chrome did not start');
  ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((r) => ws.addEventListener('open', r));
  ws.addEventListener('message', (e) => {
    const m = JSON.parse(e.data);
    if (m.id && pending.has(m.id)) { const p = pending.get(m.id); pending.delete(m.id); m.error ? p.reject(new Error(m.error.message)) : p.resolve(m.result); }
  });
  await send('Emulation.setDeviceMetricsOverride', { width: 1180, height: 900, deviceScaleFactor: 2, mobile: false });
  await send('Emulation.setEmulatedMedia', { features: [{ name: 'prefers-color-scheme', value: 'light' }] });
  await send('Page.navigate', { url: URL_ });
  await sleep(2500);
  await evaluate(`document.fonts.ready.then(() => true)`);
  mkdirSync(OUT, { recursive: true });

  for (const [name, setup, sel] of SHOTS) {
    await evaluate(`(async () => { ${setup}; await new Promise((r) => setTimeout(r, 600)); })()`);
    const box = await evaluate(`(() => {
      const els = [...document.querySelectorAll(${JSON.stringify(sel)})];
      const rs = els.map((e) => (e.closest('section') && !e.matches('header') && e.id !== 'detail' && e.id !== 'ready' && !e.matches('.day-block') ? e.closest('section') : e).getBoundingClientRect());
      const x = Math.min(...rs.map((r) => r.left)), y = Math.min(...rs.map((r) => r.top));
      const r = Math.max(...rs.map((r) => r.right)), b = Math.max(...rs.map((r) => r.bottom));
      return { x: x - 16, y: y + scrollY - 16, width: r - x + 32, height: b - y + 32 };
    })()`);
    const { data } = await send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true, clip: { ...box, scale: 1 } });
    writeFileSync(join(OUT, `${name}.png`), Buffer.from(data, 'base64'));
    console.log('saved', `${name}.png`);
  }

  // Daily summary cards come straight off the canvas at full size (1080 wide).
  await evaluate(`showView('summary')`);
  const picks = await evaluate(`[...document.querySelectorAll('#sum-pick option')].map((o) => [o.value, o.textContent])`);
  for (const [value] of picks) {
    const png = await evaluate(`(async () => { const p = document.querySelector('#sum-pick'); p.value = ${JSON.stringify(value)};
      await drawSummary(); return document.querySelector('#sum-canvas').toDataURL('image/png'); })()`);
    const [date, day] = value.split('|');
    const file = `summary ${date} ${day.replace(/^\S+\s+/, '').replace('BicepForearm', 'Bicep & Forearm')}.png`;
    writeFileSync(join(OUT, file), Buffer.from(png.split(',')[1], 'base64'));
    console.log('saved', file);
  }
} finally {
  ws?.close();
  chrome.kill();
}
