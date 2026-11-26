(() => {
  'use strict';
  const $ = (s) => document.querySelector(s);
  const PAL = ['#1F4FFF', '#E5243B', '#0A2A9E', '#FF6B7D', '#5C84FF', '#B01228', '#A9BEFF', '#FFB3BC'];
  const TYPES = [['bar', 'Bar'], ['hbar', 'Horizontal bar'], ['stackedBar', 'Stacked bar'], ['combo', 'Bar + line'],
    ['line', 'Line'], ['step', 'Step line'], ['area', 'Area'], ['stackedArea', 'Stacked area'],
    ['scatter', 'Scatter'], ['bubble', 'Bubble'], ['histogram', 'Histogram'],
    ['pie', 'Pie'], ['doughnut', 'Doughnut'], ['radar', 'Radar'], ['polarArea', 'Polar']];
  const SINGLE = ['pie', 'doughnut', 'polarArea'];
  const ONE = [...SINGLE, 'histogram'];
  const XY = ['scatter', 'bubble'];
  const MAX = 300, GRID = 'rgba(11,21,51,.08)';
  const PDFV = '3.11.174';
  const LIBS = {
    Chart: ['https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js'],
    Papa: ['https://cdn.jsdelivr.net/npm/papaparse@5.4.1/papaparse.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js'],
    XLSX: ['https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'],
    mammoth: ['https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js'],
    pdfjsLib: [`https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFV}/pdf.min.js`, `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFV}/build/pdf.min.js`]
  };
  const STOP = new Set('the and for with from are was were been this that these those you she they not but has have had his her its our your their will would can could may which what when who into than then also'.split(' '));

  const state = { rows: [], cols: [], numeric: [], type: 'bar', x: null, ys: [], agg: 'none', chart: null };
  const el = {
    drop: $('#drop'), file: $('#file'), studio: $('#studio'), types: $('#types'), x: $('#xcol'), ys: $('#ycols'),
    agg: $('#agg'), table: $('#table'), canvas: $('#chart'), msg: $('#msg'), reset: $('#reset'),
    download: $('#download'), name: $('#name')
  };

  const fail = (m) => Object.assign(new Error(m), { user: true });
  const setMsg = (m, info) => { el.msg.textContent = m || ''; el.msg.className = info ? 'info' : ''; };
  const toNum = (v) => {
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    if (typeof v !== 'string' || !v.trim()) return null;
    const n = Number(v.replace(/[,\s$€£%]/g, ''));
    return Number.isFinite(n) ? n : null;
  };

  /* Libraries load only when needed, so the upload box always works */
  const pending = {};
  function lib(name) {
    if (window[name]) return Promise.resolve(window[name]);
    if (pending[name]) return pending[name];
    pending[name] = (async () => {
      for (const src of LIBS[name]) {
        try {
          await new Promise((ok, no) => {
            const s = document.createElement('script');
            s.src = src; s.onload = ok; s.onerror = no;
            document.head.append(s);
          });
          if (!window[name]) continue;
          if (name === 'Chart') {
            Chart.defaults.font.family = 'system-ui, sans-serif';
            Chart.defaults.color = '#5b6584';
          }
          if (name === 'pdfjsLib') pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFV}/pdf.worker.min.js`;
          return window[name];
        } catch (e) { /* try next source */ }
      }
      delete pending[name];
      throw fail('Could not load a required library. Check your internet connection.');
    })();
    return pending[name];
  }

  /* ---------- Reading files ---------- */
  function flatten(o, p = '', out = {}) {
    for (const [k, v] of Object.entries(o)) {
      const key = p ? `${p}.${k}` : k;
      if (v && typeof v === 'object' && !Array.isArray(v)) flatten(v, key, out);
      else out[key] = Array.isArray(v) ? JSON.stringify(v) : v;
    }
    return out;
  }
  function fromJson(d) {
    if (d && !Array.isArray(d) && typeof d === 'object') d = Object.values(d).find(Array.isArray) || Object.entries(d).map(([key, value]) => ({ key, value }));
    if (!Array.isArray(d)) throw fail('Unsupported JSON structure.');
    if (d.length && Array.isArray(d[0])) {
      const [h, ...rest] = d;
      return rest.map((r) => Object.fromEntries(h.map((k, i) => [String(k), r[i]])));
    }
    return d.map((r) => (r && typeof r === 'object' ? flatten(r) : { value: r }));
  }
  function wordRows(text) {
    const counts = new Map();
    (text.toLowerCase().match(/[\p{L}\p{N}']+/gu) || []).forEach((w) => {
      if (w.length > 2 && !STOP.has(w)) counts.set(w, (counts.get(w) || 0) + 1);
    });
    if (!counts.size) throw fail('No readable text found in this file.');
    return [...counts].sort((a, b) => b[1] - a[1]).slice(0, 30).map(([Word, Count]) => ({ Word, Count }));
  }

  async function readFile(f) {
    const ext = f.name.split('.').pop().toLowerCase();
    if (ext === 'json') return fromJson(JSON.parse(await f.text()));
    if (ext === 'csv' || ext === 'tsv') {
      const P = await lib('Papa');
      return P.parse((await f.text()).replace(/^\uFEFF/, ''), { header: true, dynamicTyping: true, skipEmptyLines: true }).data;
    }
    if (ext === 'xlsx' || ext === 'xls') {
      const X = await lib('XLSX');
      const wb = X.read(await f.arrayBuffer(), { type: 'array' });
      return X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: null });
    }
    if (ext === 'docx') {
      const m = await lib('mammoth');
      return wordRows((await m.extractRawText({ arrayBuffer: await f.arrayBuffer() })).value);
    }
    if (ext === 'pdf') {
      const pdf = await (await lib('pdfjsLib')).getDocument({ data: await f.arrayBuffer() }).promise;
      let text = '';
      for (let i = 1; i <= Math.min(pdf.numPages, 100); i++) {
        const c = await (await pdf.getPage(i)).getTextContent();
        text += c.items.map((it) => it.str).join(' ') + ' ';
      }
      return wordRows(text);
    }
    if (ext === 'txt' || ext === 'md') return wordRows(await f.text());
    throw fail('Unsupported file type.');
  }

  async function load(file) {
    if (!file) return;
    setMsg('Reading file…', true);
    try {
      const rows = (await readFile(file)).filter((r) => r && Object.keys(r).length);
      if (!rows.length) throw fail('No data found in this file.');
      el.name.textContent = file.name;
      init(rows);
      setMsg('');
    } catch (e) {
      setMsg(e.user ? e.message : 'Could not read this file.');
    }
  }

  /* ---------- Controls ---------- */
  function chip(text, value, fn) {
    const b = document.createElement('button');
    b.type = 'button'; b.className = 'chip'; b.textContent = text; b.dataset.v = value;
    b.setAttribute('aria-pressed', 'false');
    b.addEventListener('click', fn);
    return b;
  }

  function init(rows) {
    const sample = rows.slice(0, 200);
    const cols = [...new Set(sample.flatMap((r) => Object.keys(r)))];
    const numeric = cols.filter((c) => {
      const v = sample.map((r) => r[c]).filter((x) => x !== null && x !== undefined && x !== '');
      return v.length && v.filter((x) => toNum(x) !== null).length / v.length >= 0.8;
    });
    const x = cols.find((c) => !numeric.includes(c)) || cols[0];
    const y = numeric.find((c) => c !== x);
    Object.assign(state, { rows, cols, numeric, x, ys: y ? [y] : [], agg: y ? 'none' : 'count', type: 'bar' });
    el.x.replaceChildren(...cols.map((c) => new Option(c, c)));
    el.ys.replaceChildren(...numeric.map((c) => chip(c, c, () => toggleY(c))));
    el.drop.hidden = true; el.studio.hidden = false; el.reset.hidden = false;
    renderTable(); sync(); draw();
  }

  function sync() {
    el.types.querySelectorAll('.chip').forEach((b) => b.setAttribute('aria-pressed', b.dataset.v === state.type));
    el.ys.querySelectorAll('.chip').forEach((b) => b.setAttribute('aria-pressed', state.ys.includes(b.dataset.v)));
    el.x.value = state.x; el.agg.value = state.agg;
  }

  function toggleY(c) {
    state.ys = ONE.includes(state.type) ? [c] : state.ys.includes(c) ? state.ys.filter((y) => y !== c) : [...state.ys, c];
    sync(); draw();
  }

  function renderTable() {
    const cell = (tag, t) => { const e = document.createElement(tag); e.textContent = t; return e; };
    const head = document.createElement('thead'), hr = document.createElement('tr');
    state.cols.forEach((c) => hr.append(cell('th', c)));
    head.append(hr);
    const body = document.createElement('tbody');
    state.rows.slice(0, 50).forEach((r) => {
      const tr = document.createElement('tr');
      state.cols.forEach((c) => tr.append(cell('td', r[c] ?? '')));
      body.append(tr);
    });
    el.table.replaceChildren(head, body);
  }

  /* ---------- Data and chart ---------- */
  function buildData() {
    const { rows, x, ys, agg } = state;
    if (agg === 'none') {
      const part = rows.slice(0, MAX);
      return { labels: part.map((r) => String(r[x] ?? '')), series: ys.map((y) => ({ name: y, data: part.map((r) => toNum(r[y])) })) };
    }
    const groups = new Map();
    for (const r of rows) {
      const k = String(r[x] ?? '');
      let g = groups.get(k);
      if (!g) groups.set(k, (g = { n: 0, sum: ys.map(() => 0), cnt: ys.map(() => 0) }));
      g.n++;
      ys.forEach((y, i) => { const v = toNum(r[y]); if (v !== null) { g.sum[i] += v; g.cnt[i]++; } });
    }
    const labels = [...groups.keys()].slice(0, MAX);
    if (agg === 'count') return { labels, series: [{ name: 'Count', data: labels.map((k) => groups.get(k).n) }] };
    return {
      labels,
      series: ys.map((y, i) => ({
        name: `${agg === 'avg' ? 'Average' : 'Sum'} ${y}`,
        data: labels.map((k) => { const g = groups.get(k); return agg === 'avg' ? (g.cnt[i] ? g.sum[i] / g.cnt[i] : null) : g.sum[i]; })
      }))
    };
  }

  const kill = () => { if (state.chart) { state.chart.destroy(); state.chart = null; } };

  function renderStats() {
    const box = $('#stats');
    if (!box) return;
    box.replaceChildren();
    const f = (n) => n.toLocaleString(undefined, { maximumFractionDigits: 2 });
    state.ys.forEach((y) => {
      const v = state.rows.map((r) => toNum(r[y])).filter((n) => n !== null);
      if (!v.length) return;
      const sum = v.reduce((a, b) => a + b, 0);
      const min = v.reduce((a, b) => Math.min(a, b), Infinity);
      const max = v.reduce((a, b) => Math.max(a, b), -Infinity);
      const card = document.createElement('div');
      card.className = 'stat';
      const t = document.createElement('b');
      t.textContent = y;
      card.append(t);
      [['Sum', sum], ['Average', sum / v.length], ['Min', min], ['Max', max]].forEach(([k, n]) => {
        const s = document.createElement('span');
        s.textContent = `${k}: ${f(n)}`;
        card.append(s);
      });
      box.append(card);
    });
  }

  async function draw() {
    const { type, x, ys, agg, rows } = state;
    setMsg(''); kill(); renderStats();
    if ((XY.includes(type) || type === 'histogram' || agg !== 'count') && !ys.length) return setMsg('Select a value column.');
    try { await lib('Chart'); } catch (e) { return setMsg(e.message); }
    kill();

    const radial = SINGLE.includes(type);
    const numX = state.numeric.includes(x);
    let config;

    if (XY.includes(type)) {
      const xv = (r, k) => (numX ? toNum(r[x]) : k + 1);
      let datasets;
      if (type === 'bubble') {
        const [y, z] = ys;
        const d = rows.slice(0, 1000).map((r, k) => ({ x: xv(r, k), y: toNum(r[y]), z: z ? toNum(r[z]) : null })).filter((p) => p.x !== null && p.y !== null);
        const zs = d.map((p) => p.z).filter((n) => n !== null);
        const lo = zs.reduce((a, b) => Math.min(a, b), Infinity);
        const hi = zs.reduce((a, b) => Math.max(a, b), -Infinity);
        datasets = [{
          label: z ? `${y} (size: ${z})` : y,
          data: d.map((p) => ({ x: p.x, y: p.y, r: p.z === null || hi === lo ? 8 : 4 + 20 * (p.z - lo) / (hi - lo) })),
          backgroundColor: PAL[0] + '99', borderColor: PAL[0]
        }];
      } else {
        datasets = ys.map((y, i) => ({
          label: y,
          data: rows.slice(0, 2000).map((r, k) => ({ x: xv(r, k), y: toNum(r[y]) })).filter((p) => p.x !== null && p.y !== null),
          backgroundColor: PAL[i % PAL.length] + 'CC', borderColor: PAL[i % PAL.length], pointRadius: 4
        }));
      }
      config = { type, data: { datasets } };
    } else if (type === 'histogram') {
      const v = rows.map((r) => toNum(r[ys[0]])).filter((n) => n !== null);
      if (!v.length) return setMsg('No numbers to plot.');
      const lo = v.reduce((a, b) => Math.min(a, b), Infinity);
      const hi = v.reduce((a, b) => Math.max(a, b), -Infinity);
      const bins = Math.min(30, Math.max(5, Math.ceil(Math.log2(v.length) + 1)));
      const w = (hi - lo) / bins || 1;
      const counts = Array(bins).fill(0);
      v.forEach((n) => { counts[Math.min(bins - 1, Math.floor((n - lo) / w))]++; });
      const f = (n) => +n.toFixed(2);
      const labels = counts.map((_, i) => `${f(lo + i * w)} to ${f(lo + (i + 1) * w)}`);
      config = { type: 'bar', data: { labels, datasets: [{ label: `Count of ${ys[0]}`, data: counts, backgroundColor: PAL[0], borderRadius: 2, barPercentage: 1, categoryPercentage: 1 }] } };
    } else {
      const { labels, series } = buildData();
      const lineLike = ['line', 'step', 'area', 'stackedArea', 'radar'];
      const datasets = series.map((s, i) => {
        const c = PAL[i % PAL.length];
        const dots = labels.length > 60 ? 0 : 3;
        if (radial) return { label: s.name, data: s.data, backgroundColor: labels.map((_, j) => PAL[j % PAL.length]), borderColor: '#fff', borderWidth: 2 };
        if (lineLike.includes(type)) {
          const fill = type !== 'line' && type !== 'step';
          return { label: s.name, data: s.data, borderColor: c, borderWidth: 2.5, tension: type === 'step' ? 0 : 0.3, stepped: type === 'step', spanGaps: true, fill, backgroundColor: fill ? c + (type === 'stackedArea' ? '99' : '33') : c, pointRadius: dots };
        }
        if (type === 'combo' && i > 0) return { type: 'line', label: s.name, data: s.data, borderColor: c, backgroundColor: c, borderWidth: 2.5, tension: 0.3, spanGaps: true, pointRadius: dots };
        return { label: s.name, data: s.data, backgroundColor: c, borderRadius: 4 };
      });
      const base = { hbar: 'bar', stackedBar: 'bar', combo: 'bar', area: 'line', stackedArea: 'line', step: 'line' }[type] || type;
      config = { type: base, data: { labels, datasets } };
    }

    const options = {
      responsive: true, maintainAspectRatio: false, animation: { duration: 400 },
      plugins: { legend: { display: radial || config.data.datasets.length > 1, labels: { usePointStyle: true } } }
    };
    if (!radial && type !== 'radar') {
      const stacked = type === 'stackedBar' || type === 'stackedArea';
      options.indexAxis = type === 'hbar' ? 'y' : 'x';
      options.scales = {
        x: { stacked, grid: { color: GRID }, title: { display: XY.includes(type), text: numX ? x : 'Row' } },
        y: { stacked, grid: { color: GRID }, beginAtZero: ['bar', 'hbar', 'stackedBar', 'combo', 'histogram'].includes(type) }
      };
    }
    config.options = options;
    state.chart = new Chart(el.canvas.getContext('2d'), config);
  }

  /* ---------- Events ---------- */
  TYPES.forEach(([id, label]) => el.types.append(chip(label, id, () => {
    state.type = id;
    if (ONE.includes(id)) state.ys = state.ys.slice(0, 1);
    sync(); draw();
  })));
  el.x.addEventListener('change', () => { state.x = el.x.value; draw(); });
  el.agg.addEventListener('change', () => { state.agg = el.agg.value; draw(); });

  el.file.addEventListener('change', () => { const f = el.file.files[0]; el.file.value = ''; load(f); });
  ['dragenter', 'dragover'].forEach((ev) => el.drop.addEventListener(ev, (e) => { e.preventDefault(); el.drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => el.drop.addEventListener(ev, (e) => { e.preventDefault(); el.drop.classList.remove('over'); }));
  el.drop.addEventListener('drop', (e) => load(e.dataTransfer.files[0]));
  ['dragover', 'drop'].forEach((ev) => window.addEventListener(ev, (e) => e.preventDefault()));

  el.reset.addEventListener('click', () => {
    kill(); setMsg('');
    el.studio.hidden = true; el.reset.hidden = true; el.drop.hidden = false;
  });

  el.download.addEventListener('click', () => {
    if (!state.chart) return;
    const out = document.createElement('canvas');
    out.width = el.canvas.width; out.height = el.canvas.height;
    const g = out.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, out.width, out.height); g.drawImage(el.canvas, 0, 0);
    const a = document.createElement('a');
    a.download = 'titronic-chart.png'; a.href = out.toDataURL('image/png'); a.click();
  });
})();
