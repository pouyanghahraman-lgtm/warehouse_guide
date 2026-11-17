/* =====================================================================
   انبار هوشمند — برداشت سفارش با نمایش موقعیت روی نقشه
   - دو فایل Excel نقشه:
       1) فایل بارکدها (warehouse_products-Barcode.xlsx) → بارکدها و هندسه خانه‌ها
       2) فایل لیوت (warehouse_layout_55x16.xlsx) → کد انبار که روی کارتن نوشته شده (مثل 720 P)
   - بارکد ۴ رقمی = پیشوند 626057300 + ۴ رقم؛ کدهای کامل همان‌طور استفاده می‌شوند
   - اقلام سفارش به ترتیب مسیر حرکت در انبار مرتب می‌شوند
   ===================================================================== */
(() => {
  'use strict';

  /* ---------- تنظیمات ---------- */
  const BARCODE_PREFIX = '626057300';
  const DEFAULT_MAP_URL = 'data/warehouse_products-Barcode.xlsx';
  const DEFAULT_MAP_NAME = 'warehouse_products-Barcode.xlsx';
  const DEFAULT_LAYOUT_URL = 'data/warehouse_layout_55x16.xlsx';
  const DEFAULT_LAYOUT_NAME = 'warehouse_layout_55x16.xlsx';
  const KEY_LAYOUT = 'wh.layout.v5';
  const KEY_ORDERS = 'wh.orders.v1';
  const KEY_THEME = 'wh.theme';
  const KEY_ME = 'wh.me.v1';
  const ZOOMS = [1, 2, 4, 6];
  const ZOOM_LABEL = ['۱×', '۲×', '۴×', '۶×'];
  const DEFAULT_ZOOM = 1;
  const TEXT_ZOOM = 2;    // از این زوم به بعد کد روی همه قفسه‌ها دیده می‌شود
  const ROUTE_BAND_M = 2; // طول هر باند مسیر (متر)
  const MAP_PAD = 24;     // حاشیه نقشه (واحد SVG)
  const DOOR_SPACE = 74;  // فضای زیر نقشه برای ورودی

  /* ---------- عناصر صفحه ---------- */
  const app = document.getElementById('app');
  const fileInput = document.getElementById('fileInput');
  const subtitle = document.getElementById('subtitle');
  const btnHome = document.getElementById('btnHome');
  const btnTheme = document.getElementById('btnTheme');
  const toastEl = document.getElementById('toast');

  /* ---------- ذخیره‌سازی (فقط روی همین مرورگر) ---------- */
  const store = {
    read(key, fallback) {
      try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : fallback; }
      catch (e) { return fallback; }
    },
    write(key, value) {
      try { localStorage.setItem(key, JSON.stringify(value)); return true; }
      catch (e) { return false; }
    }
  };

  /* ---------- وضعیت برنامه ---------- */
  const state = {
    layout: store.read(KEY_LAYOUT, null),
    orders: store.read(KEY_ORDERS, []),
    view: { name: 'home' },
    importTarget: null,
    zoom: DEFAULT_ZOOM,
    mapError: '',
    lastPinKey: '',
    me: store.read(KEY_ME, null),   // موقعیت کاربر: کلید «ردیف-ستون» خانه‌ی نقشه
    keepScroll: null
  };
  // لینک QR به شکل ?at=ردیف-ستون باز شود: موقعیت را ثبت کن و آدرس را تمیز کن
  try {
    const at = new URLSearchParams(location.search).get('at');
    if (at) {
      state.me = at.trim();
      store.write(KEY_ME, state.me);
      history.replaceState(null, '', location.pathname);
    }
  } catch (e) { /* بدون دسترسی به آدرس، نادیده بگیر */ }

  /* ---------- ابزارها ---------- */
  const DIGITS = {};
  '۰۱۲۳۴۵۶۷۸۹'.split('').forEach((d, i) => { DIGITS[d] = String(i); });
  '٠١٢٣٤٥٦٧٨٩'.split('').forEach((d, i) => { DIGITS[d] = String(i); });
  const toEn = (s) => String(s ?? '').replace(/[۰-۹٠-٩]/g, (d) => DIGITS[d]);
  const normFa = (s) => toEn(s).replace(/ي/g, 'ی').replace(/ك/g, 'ک').replace(/‌/g, ' ').replace(/\s+/g, ' ').trim();
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fa = (n) => Number(n).toLocaleString('fa-IR');
  const fmtM = (n) => Number(n).toLocaleString('fa-IR', { maximumFractionDigits: 1 });
  const uid = () => Math.random().toString(36).slice(2, 8) + Date.now().toString(36).slice(-4);
  const dateFa = (ts) => new Date(ts).toLocaleDateString('fa-IR', { year: 'numeric', month: 'short', day: 'numeric' });

  /** بارکد را به شکل کامل برمی‌گرداند: ۴ رقم یا کمتر → پیشوند + ۴ رقم */
  function normBarcode(v) {
    if (v === null || v === undefined || v === '') return '';
    let s = typeof v === 'number' ? String(Math.round(v)) : String(v);
    s = toEn(s).replace(/\D/g, '');
    if (!s) return '';
    return s.length <= 4 ? BARCODE_PREFIX + s.padStart(4, '0') : s;
  }

  function toast(msg, kind) {
    toastEl.textContent = msg;
    toastEl.className = 'toast' + (kind === 'error' ? ' error' : '');
    toastEl.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { toastEl.hidden = true; }, 3200);
  }

  /* ---------- تم روشن / تیره ---------- */
  function currentTheme() {
    const saved = document.documentElement.dataset.theme;
    if (saved === 'dark' || saved === 'light') return saved;
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  function toggleTheme() {
    const next = currentTheme() === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    store.write(KEY_THEME, next);
    render();
  }

  /* ---------- خواندن نقشه از Excel ---------- */
  function fillOf(cell) {
    const s = cell && cell.s;
    if (!s || !s.fgColor || !s.fgColor.rgb) return null;
    if (s.patternType && s.patternType !== 'solid') return null;
    let rgb = String(s.fgColor.rgb);
    if (rgb.length === 8) rgb = rgb.slice(2);
    if (rgb.length !== 6 || rgb.toUpperCase() === 'FFFFFF') return null;
    return '#' + rgb;
  }

  const sheetOf = (wb) => wb.Sheets[wb.SheetNames.find((n) => /نقشه/.test(normFa(n))) || wb.SheetNames[0]];

  /**
   * wb  : فایل بارکدها (هندسه خانه‌ها، بارکدها، فاصله‌ها)
   * wbL : فایل لیوت (کد انبار روی کارتن) — اختیاری؛ به همان خانه‌ها وصل می‌شود
   */
  function parseLayout(wb, sourceName, wbL, layoutName) {
    const ws = sheetOf(wb);
    if (!ws || !ws['!ref']) throw new Error('شیت نقشه خالی است.');
    const wsL = wbL ? sheetOf(wbL) : null;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const cellObj = (r, c) => ws[XLSX.utils.encode_cell({ r, c })];
    const val = (r, c) => { const cell = cellObj(r, c); return cell ? cell.v : null; };
    const num = (v) => { const n = parseFloat(toEn(v)); return Number.isFinite(n) ? n : null; };
    const labelAt = (r, c) => {
      const cl = wsL && wsL[XLSX.utils.encode_cell({ r, c })];
      return cl && cl.v !== null && cl.v !== undefined ? String(cl.v).replace(/\s*[\r\n]+\s*/g, ' ').trim() : '';
    };

    // ستون «فاصله از پایین»
    let yCol = -1;
    for (let r = 0; r <= Math.min(range.e.r, 10) && yCol < 0; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        if (/فاصله از پایین/.test(normFa(val(r, c)))) { yCol = c; break; }
      }
    }
    if (yCol < 0) yCol = 34; // AI

    // ردیف عرض‌ها (فاصله از چپ)
    let hdr = -1;
    for (let r = 0; r <= 10 && hdr < 0; r++) {
      let n = 0;
      for (let c = 1; c < yCol; c++) if (num(val(r, c)) !== null) n++;
      if (n >= 5) hdr = r;
    }
    if (hdr < 0) throw new Error('ردیف فاصله از چپ (۰٫۵ تا ۱۶ متر) پیدا نشد.');

    const xOf = {};
    for (let c = 1; c < yCol; c++) { const n = num(val(hdr, c)); if (n !== null) xOf[c] = n; }
    const knownCols = Object.keys(xOf).map(Number).sort((a, b) => a - b);
    for (let c = 1; c < yCol; c++) {
      if (xOf[c] !== undefined) continue;
      const left = [...knownCols].reverse().find((k) => k < c);
      const right = knownCols.find((k) => k > c);
      if (left !== undefined && right !== undefined) xOf[c] = (xOf[left] + xOf[right]) / 2;
    }

    const yKnown = {};
    for (let r = hdr + 1; r <= range.e.r; r++) { const n = num(val(r, yCol)); if (n !== null) yKnown[r] = n; }
    const knownRows = Object.keys(yKnown).map(Number).sort((a, b) => a - b);
    const rowM = (r) => {
      if (yKnown[r] !== undefined) return yKnown[r];
      const up = [...knownRows].reverse().find((k) => k < r);
      const down = knownRows.find((k) => k > r);
      if (up === undefined || down === undefined) return null;
      return yKnown[up] + ((yKnown[down] - yKnown[up]) * (r - up)) / (down - up);
    };

    /* ---- هندسه واقعی اکسل (پیکسل) ---- */
    const R0 = hdr, R1 = range.e.r, C0 = 0, C1 = range.e.c;
    const cols = ws['!cols'] || [];
    const rows = ws['!rows'] || [];
    const colPx = (c) => (cols[c] && cols[c].wpx) || (cols[c] && cols[c].width ? Math.round(cols[c].width * 7 + 5) : 64);
    const rowPx = (r) => (rows[r] && rows[r].hpx) || (rows[r] && rows[r].hpt ? rows[r].hpt * 4 / 3 : 20);
    const colX = [0];
    for (let c = C0; c <= C1; c++) colX.push(colX[colX.length - 1] + colPx(c));
    const rowY = [0];
    for (let r = R0; r <= R1; r++) rowY.push(rowY[rowY.length - 1] + rowPx(r));
    const X = (c) => colX[Math.min(Math.max(c - C0, 0), colX.length - 1)];
    const Y = (r) => rowY[Math.min(Math.max(r - R0, 0), rowY.length - 1)];

    const mergeTL = {};
    const covered = new Set();
    for (const m of ws['!merges'] || []) {
      for (let rr = m.s.r; rr <= m.e.r; rr++) {
        for (let cc = m.s.c; cc <= m.e.c; cc++) {
          if (rr === m.s.r && cc === m.s.c) continue;
          covered.add(rr + ':' + cc);
        }
      }
      mergeTL[m.s.r + ':' + m.s.c] = { rs: m.e.r - m.s.r + 1, cs: m.e.c - m.s.c + 1 };
    }

    const cells = [];
    const index = {};
    let maxX = 0;
    let maxY = 0;
    let labelCount = 0;
    let office = null;

    for (let r = R0; r <= R1; r++) {
      for (let c = C0; c <= C1; c++) {
        const key = r + ':' + c;
        if (covered.has(key)) continue;
        const cobj = cellObj(r, c);
        const v = cobj ? cobj.v : null;
        const text = v === null || v === undefined ? '' : (typeof v === 'number' ? String(Math.round(v)) : String(v));
        const fill = fillOf(cobj);
        const label = labelAt(r, c);
        if (!text.trim() && !fill && !label) continue;

        const span = mergeTL[key] || { rs: 1, cs: 1 };
        const x0 = X(c);
        const y0 = Y(r);
        const w = X(c + span.cs) - x0;
        const h = Y(r + span.rs) - y0;

        // مختصات متری و بارکدها فقط برای خانه‌های داخل بخش بارکد
        let mx = null;
        let my = null;
        const codes = [];
        const labels = [];
        if (r > hdr && c >= 1 && c < yCol && xOf[c] !== undefined) {
          const yy = rowM(r);
          if (yy !== null) { mx = xOf[c]; my = yy; }
        }
        if (mx !== null && text.trim()) {
          for (const raw of text.split(/[\r\n]+/)) {
            const t = toEn(raw).replace(/\s+/g, '');
            if (!t) continue;
            if (/^\d+$/.test(t)) codes.push(t.length <= 4 ? BARCODE_PREFIX + t.padStart(4, '0') : t);
            else labels.push(raw.trim());
          }
        }
        const uniq = [...new Set(codes)];
        const isBlue = !!(cobj && cobj.s && cobj.s.fgColor && cobj.s.fgColor.theme === 3);
        // ستون فاصله‌ها و ردیف عنوان، بلوک نیستند
        const isRuler = c === yCol || r === hdr;
        const kind = isRuler ? '' : uniq.length ? (isBlue ? 'pallet' : 'rack') : ((label || text.trim() || fill) ? 'zone' : '');
        const idx = cells.length;
        cells.push({
          r, c, x: mx, y: my,
          x0, y0, w, h,
          cx: x0 + w / 2, cy: y0 + h / 2,
          fill, kind,
          text: text.trim(),
          label: label || (uniq.length ? '' : text.trim()),
          codes: uniq, labels
        });
        if (/اداری/.test(normFa(label + ' ' + labels.join(' ')))) office = idx;
        for (const k of uniq) (index[k] ||= []).push(idx);
        labelCount += label ? 1 : 0;
        if (mx !== null) { maxX = Math.max(maxX, mx); maxY = Math.max(maxY, my); }
      }
    }
    if (!Object.keys(index).length) throw new Error('هیچ بارکدی روی نقشه پیدا نشد.');

    return {
      source: sourceName,
      layoutSource: layoutName || null,
      savedAt: Date.now(),
      width: maxX + 0.5,
      length: maxY + 1.5,
      geo: { W: colX[colX.length - 1], H: rowY[rowY.length - 1] },
      cells,
      index,
      office,
      stats: {
        barcodes: Object.keys(index).length,
        locations: cells.filter((c) => c.codes.length).length,
        codes: labelCount
      }
    };
  }

  /* ---------- خواندن فایل سفارش ---------- */
  function parseOrder(wb) {
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
    const h = rows.findIndex((r) => r.some((v) => /بارک|barcode/i.test(normFa(v))));
    const head = h >= 0 ? rows[h].map((v) => normFa(v)) : [];
    const find = (re) => head.findIndex((v) => re.test(v));
    let cB = find(/بارک|barcode/i);
    let cN = find(/نام|محصول|کالا|product|name/i);
    let cQ = find(/تعداد|مقدار|order\s*count|count|qty|quantity/i);
    if (cB < 0) cB = 0;
    if (cN < 0) cN = cB === 1 ? 2 : 1;
    if (cQ < 0) cQ = [0, 1, 2].find((i) => i !== cB && i !== cN) ?? 2;

    const items = [];
    for (let i = h + 1; i < rows.length; i++) {
      const row = rows[i];
      const barcode = normBarcode(row[cB]);
      if (!barcode) continue;
      const name = String(row[cN] ?? '').trim() || 'کالای بدون نام';
      const q = parseInt(toEn(row[cQ]).replace(/[^\d]/g, ''), 10);
      items.push({ barcode, name, need: q > 0 ? q : 1 });
    }
    if (!items.length) throw new Error('هیچ ردیفی با بارکد در فایل پیدا نشد. ستون «بارکد» را بررسی کنید.');
    return items;
  }

  /* ---------- نقشه: جست‌وجو و مسیر ---------- */
  const cellIdxFor = (bc) => (state.layout && state.layout.index[bc]) || [];
  const cellAt = (i) => state.layout.cells[i];
  // باند ۲ متری از پایین به بالا؛ در باندهای زوج x صعودی، در باندهای فرد نزولی
  const routeKey = (c) => {
    if (!c || c.x === null) return 1e9;
    const band = Math.floor(c.y / ROUTE_BAND_M);
    return band * 1000 + (band % 2 === 0 ? c.x : 100 - c.x);
  };
  const locsOf = (bc) => cellIdxFor(bc).slice().sort((a, b) => routeKey(cellAt(a)) - routeKey(cellAt(b)));
  const primaryIdxOf = (item) => {
    const locs = locsOf(item.barcode);
    if (!locs.length) return null;
    return locs[Math.min(item.loc || 0, locs.length - 1)];
  };

  /* ---------- سفارش ---------- */
  function buildOrder(name, rawItems) {
    const items = rawItems.map((it) => ({
      id: uid(), barcode: it.barcode, name: it.name, need: it.need,
      picked: null, status: 'pending', loc: 0
    }));
    const firstCell = (bc) => { const l = locsOf(bc); return l.length ? cellAt(l[0]) : null; };
    items.sort((a, b) => routeKey(firstCell(a.barcode)) - routeKey(firstCell(b.barcode)));
    return { id: uid(), name, createdAt: Date.now(), items, cursor: items[0] ? items[0].id : null };
  }

  const orderById = (id) => state.orders.find((o) => o.id === id);
  const currentItem = (order) => {
    const byCursor = order.items.find((i) => i.id === order.cursor);
    if (byCursor) return byCursor;
    return order.items.find((i) => i.status === 'pending') || null;
  };
  const nextPending = (order, fromId) => {
    const idx = order.items.findIndex((i) => i.id === fromId);
    const after = order.items.slice(idx + 1).find((i) => i.status === 'pending');
    return after || order.items.find((i) => i.status === 'pending') || null;
  };
  const orderStats = (order) => {
    const done = order.items.filter((i) => i.status === 'done').length;
    const partial = order.items.filter((i) => i.status === 'partial').length;
    const missing = order.items.filter((i) => i.status === 'missing').length;
    return { done, partial, missing, handled: done + partial + missing, total: order.items.length };
  };
  const saveOrders = () => store.write(KEY_ORDERS, state.orders);

  /* ---------- نقشه SVG (طراحی اختصاصی) ---------- */
  function mapSVG(primaryIdx, hitSet) {
    const L = state.layout;
    const W = L.geo.W;
    const H = L.geo.H;
    const P = MAP_PAD;
    const VW = W + 2 * P;
    const VH = H + DOOR_SPACE + 2 * P;
    const showText = state.zoom >= TEXT_ZOOM;
    const out = [];

    out.push(`<svg class="map-svg" viewBox="${-P} ${-P} ${VW} ${VH}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="نقشه انبار">`);
    out.push(`<defs>
      <linearGradient id="gBlue" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#8FBBFF"/><stop offset="1" stop-color="#3D7BEE"/></linearGradient>
      <linearGradient id="gGreen" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#6ED9A2"/><stop offset="1" stop-color="#22A463"/></linearGradient>
      <linearGradient id="gZone" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#F3F5F9"/><stop offset="1" stop-color="#DFE5EE"/></linearGradient>
      <linearGradient id="gFloor" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#FFFFFF"/><stop offset="1" stop-color="#F1F4F9"/></linearGradient>
      <pattern id="dots" width="16" height="16" patternUnits="userSpaceOnUse"><circle cx="2" cy="2" r="1.1" fill="#D6DEE9"/></pattern>
      <filter id="shadow" x="-10%" y="-10%" width="120%" height="130%"><feDropShadow dx="0" dy="2" stdDeviation="2" flood-color="#0F172A" flood-opacity="0.22"/></filter>
    </defs>`);

    // سطح انبار با نقطه‌های شبکه
    out.push(`<rect class="floor" x="${-P}" y="${-P}" width="${VW}" height="${VH}" rx="22"/>`);
    out.push(`<rect class="floor-dots" x="${-P}" y="${-P}" width="${VW}" height="${VH}" rx="22" fill="url(#dots)"/>`);
    out.push(`<rect class="wall" x="0" y="0" width="${W}" height="${H}" rx="10"/>`);

    // بلوک‌ها: پالت (آبی)، قفسه (سبز)، ناحیه‌ها (خاکستری)
    let blocks = '';
    let texts = '';
    L.cells.forEach((c, i) => {
      if (!c.kind) return;
      const inset = 1.6;
      const x = c.x0 + inset;
      const y = c.y0 + inset;
      const w = Math.max(1, c.w - 2 * inset);
      const h = Math.max(1, c.h - 2 * inset);
      const fill = c.kind === 'pallet' ? 'url(#gBlue)' : c.kind === 'rack' ? 'url(#gGreen)' : 'url(#gZone)';
      const isSel = i === primaryIdx;
      const isHit = !isSel && hitSet.has(i);
      const cls = `blk ${c.kind}${isSel ? ' sel' : isHit ? ' hit' : ''}`;
      blocks += `<rect class="${cls}" x="${x.toFixed(2)}" y="${y.toFixed(2)}" width="${w.toFixed(2)}" height="${h.toFixed(2)}" rx="${Math.min(5, h / 3).toFixed(2)}" fill="${fill}"/>`;
      if (c.kind !== 'zone' && h > 8) {
        blocks += `<rect x="${(x + 1.5).toFixed(2)}" y="${(y + 1.2).toFixed(2)}" width="${Math.max(0, w - 3).toFixed(2)}" height="${Math.min(2.2, h / 4).toFixed(2)}" rx="1" fill="#FFFFFF" opacity=".45"/>`;
      }
      const label = c.label || '';
      if (label && (showText || isSel || isHit || c.kind === 'zone')) {
        const fs = Math.max(4.5, Math.min(h * 0.46, 10, w * 0.22));
        const color = isSel ? '#0F172A' : c.kind === 'zone' ? '#475569' : '#0B1B33';
        texts += `<text x="${c.cx.toFixed(2)}" y="${c.cy.toFixed(2)}" font-size="${fs.toFixed(2)}" fill="${color}" text-anchor="middle" dominant-baseline="central" font-weight="700">${esc(label)}</text>`;
      }
    });
    out.push(`<g class="blocks" filter="url(#shadow)">${blocks}</g>`);
    out.push(`<g class="labels">${texts}</g>`);

    // نشانگر محل کالا
    if (primaryIdx !== null && primaryIdx !== undefined) {
      const p = L.cells[primaryIdx];
      const r = Math.max(12, W * 0.011);
      out.push(`<g class="pin">
        <circle class="pulse" cx="${p.cx.toFixed(2)}" cy="${p.cy.toFixed(2)}" r="${(r * 2.4).toFixed(1)}"/>
        <circle class="ring" cx="${p.cx.toFixed(2)}" cy="${p.cy.toFixed(2)}" r="${(r * 1.1).toFixed(1)}"/>
        <circle class="dot" cx="${p.cx.toFixed(2)}" cy="${p.cy.toFixed(2)}" r="${(r * 0.45).toFixed(1)}"/>
      </g>`);
    }

    // موقعیت فعلی کاربر (آبی، جدا از نشانگر کالا)
    const meIdx = meCellIndex();
    if (meIdx !== null) {
      const m = L.cells[meIdx];
      const r = Math.max(9, W * 0.008);
      out.push(`<g class="me-mark">
        <circle class="me-pulse" cx="${m.cx.toFixed(2)}" cy="${m.cy.toFixed(2)}" r="${(r * 2.2).toFixed(1)}"/>
        <circle class="me-ring" cx="${m.cx.toFixed(2)}" cy="${m.cy.toFixed(2)}" r="${(r * 0.9).toFixed(1)}"/>
        <circle class="me-dot" cx="${m.cx.toFixed(2)}" cy="${m.cy.toFixed(2)}" r="${(r * 0.45).toFixed(1)}"/>
      </g>`);
    }

    // ورودی انبار: کنار «اداری»، لبه پایین
    if (L.office !== null && L.office !== undefined) {
      const o = L.cells[L.office];
      const dw = Math.max(o.w * 2, 60);
      const dx = o.x0 + o.w / 2 - dw / 2;
      out.push(`<rect class="door" x="${dx.toFixed(2)}" y="${(H + 10).toFixed(2)}" width="${dw.toFixed(2)}" height="12" rx="6"/>`);
      out.push(`<text class="entry" x="${(dx + dw / 2).toFixed(2)}" y="${(H + 44).toFixed(2)}" text-anchor="middle">ورودی انبار</text>`);
    }
    out.push('</svg>');
    return out.join('');
  }

  /* ---------- موقعیت کاربر و جست‌وجو ---------- */
  function meCellIndex() {
    if (!state.layout || !state.me) return null;
    const i = state.layout.cells.findIndex((c) => c.kind && `${c.r}-${c.c}` === state.me);
    return i >= 0 ? i : null;
  }

  function setMe(key) {
    state.me = key;
    store.write(KEY_ME, key);
  }

  function cellName(c) {
    return c.label || (c.codes && c.codes[0]) || `ردیف ${fa(c.r + 1)} · ستون ${fa(c.c + 1)}`;
  }

  /** جست‌وجو: عدد → بارکد کامل (با پیشوند) یا برچسب دقیق؛ متن → بخشی از برچسب خانه‌ها */
  function findCells(q) {
    if (!state.layout) return [];
    const compact = normFa(q).replace(/\s/g, '').toLowerCase();
    if (!compact) return [];
    const out = new Set();
    if (/^\d+$/.test(compact)) {
      (state.layout.index[normBarcode(compact)] || []).forEach((i) => out.add(i));
    }
    const digits = /^\d+$/.test(compact);
    state.layout.cells.forEach((c, i) => {
      if (!c.kind) return;
      const labels = [c.label, ...(c.labels || [])].filter(Boolean)
        .map((s) => normFa(s).replace(/\s/g, '').toLowerCase());
      if (labels.some((l) => (digits ? l === compact : l.includes(compact)))) out.add(i);
    });
    return [...out].sort((a, b) => {
      const A = state.layout.cells[a];
      const B = state.layout.cells[b];
      return (A.y - B.y) || (A.x - B.x);
    });
  }

  function searchFocus() {
    const hits = findCells(state.view.q || '');
    return hits.includes(state.view.focus) ? state.view.focus : (hits[0] ?? null);
  }

  function meBar() {
    const mi = meCellIndex();
    const text = mi !== null
      ? `موقعیت شما: <b>${esc(cellName(cellAt(mi)))}</b>`
      : 'موقعیت شما مشخص نیست. روی نقشه بزنید یا QR کنار راهرو را اسکن کنید.';
    return `<div class="me-bar">${ICON.pin}<span>${text}</span>${mi !== null ? '<button class="btn small ghost" data-action="me-clear">پاک کردن</button>' : ''}</div>`;
  }

  const searchFormHTML = (q = '') => `
    <form class="search-form" id="searchForm">
      <input class="input" id="qInput" type="search" value="${esc(q)}" placeholder="کد کالا، مثلاً 280 یا 720P" autocomplete="off" enterkeyhint="search">
      <button class="btn primary" type="submit">جستجو</button>
    </form>`;

  function searchView(q) {
    const hits = findCells(q);
    const focus = searchFocus();
    subtitle.textContent = 'جستجوی کالا';
    const head = `<section class="order-head"><div><span class="eyebrow">جستجو</span><h2>${esc(q)}</h2></div></section>`;
    if (!state.layout) return `${head}${searchFormHTML(q)}<p class="warn-text">نقشه بارگذاری نشده است.</p>`;
    if (!hits.length) return `${head}${searchFormHTML(q)}${meBar()}<div class="card empty">چیزی با «${esc(q)}» روی نقشه پیدا نشد.</div>`;

    const list = hits.map((i) => {
      const c = cellAt(i);
      return `<button class="hit-item ${i === focus ? 'on' : ''}" data-action="search-focus" data-i="${i}">
        <b>${esc(cellName(c))}</b>
        <span>${fmtM(c.x)} متر از چپ · ${fmtM(c.y)} متر از پایین</span>
      </button>`;
    }).join('');

    return `${head}${searchFormHTML(q)}
      <article class="card map-card">
        <div class="map-tools">
          <h3>${fa(hits.length)} محل پیدا شد</h3>
          <button class="btn small ghost" data-action="zoom">بزرگ‌نمایی ${ZOOM_LABEL[state.zoom]}</button>
        </div>
        ${meBar()}
        <div class="map-wrap" id="mapWrap">
          <div class="map-inner" style="width:${ZOOMS[state.zoom] * 100}%">${mapSVG(focus, new Set(hits))}</div>
        </div>
        <div class="legend">
          <span><i class="lg-pallet"></i>پالت</span>
          <span><i class="lg-rack"></i>قفسه</span>
          <span><i class="lg-pin"></i>محل کالا</span>
          <span><i class="lg-me"></i>شما</span>
        </div>
        <p class="hint">برای ثبت موقعیت خود، روی نقشه بزنید.</p>
      </article>
      <div class="hit-list">${list}</div>`;
  }

  /* ---------- قالب‌ها ---------- */
  const ICON = {
    upload: '<svg viewBox="0 0 24 24" width="24" height="24"><path d="M12 16V4m0 0-4 4m4-4 4 4M4 16v3a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-3" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    pin: '<svg viewBox="0 0 24 24" width="16" height="16"><path d="M12 22s7-6.2 7-12a7 7 0 1 0-14 0c0 5.8 7 12 7 12Zm0-9.2a2.8 2.8 0 1 1 0-5.6 2.8 2.8 0 0 1 0 5.6Z" fill="currentColor"/></svg>',
    trash: '<svg viewBox="0 0 24 24" width="18" height="18"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12M9 7V4h6v3" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    box: '<svg viewBox="0 0 24 24" width="22" height="22"><path d="M12 2 3 6.5v11L12 22l9-4.5v-11L12 2Zm0 2.2 6.8 3.4L12 11 5.2 7.6 12 4.2ZM5 9.3l6 3v7.3l-6-3V9.3Zm8 10.3v-7.3l6-3v7.3l-6 3Z" fill="currentColor"/></svg>',
    map: '<svg viewBox="0 0 24 24" width="22" height="22"><path d="M9 4 3 6v14l6-2 6 2 6-2V4l-6 2-6-2Zm0 0v14m6-12v14" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>',
    plus: '<svg viewBox="0 0 24 24" width="22" height="22"><path d="M12 5v14M5 12h14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/></svg>',
    check: '<svg viewBox="0 0 24 24" width="32" height="32"><path d="m5 12.5 4.5 4.5L19 7.5" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    alert: '<svg viewBox="0 0 24 24" width="22" height="22"><path d="M12 3 2 20h20L12 3Zm0 6v5m0 3h.01" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    moon: '<svg viewBox="0 0 24 24" width="20" height="20"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5Z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>',
    sun: '<svg viewBox="0 0 24 24" width="20" height="20"><circle cx="12" cy="12" r="4" fill="none" stroke="currentColor" stroke-width="2"/><path d="M12 2v2m0 16v2M4.9 4.9l1.4 1.4m11.4 11.4 1.4 1.4M2 12h2m16 0h2M4.9 19.1l1.4-1.4m11.4-11.4 1.4-1.4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>'
  };

  const STATUS_TEXT = { pending: 'در انتظار', done: 'برداشت شد', partial: 'کمتر از نیاز', missing: 'پیدا نشد' };

  function mapStatusCard() {
    if (state.layout) {
      const s = state.layout.stats;
      const files = [state.layout.source, state.layout.layoutSource].filter(Boolean).join(' + ');
      return `
        <section class="hero">
          <h2>نقشه انبار آماده است</h2>
          <p>${esc(files)}</p>
          <div class="stats">
            <div class="stat"><b>${fa(s.barcodes)}</b><span>بارکد</span></div>
            <div class="stat"><b>${fa(s.locations)}</b><span>محل ثبت‌شده</span></div>
            <div class="stat"><b>${fa(s.codes)}</b><span>تعداد برچسب</span></div>
          </div>
        </section>`;
    }
    return `
      <div class="notice bad">${ICON.alert}<div>
        <b>نقشه بارگذاری نشده است.</b><br>${esc(state.mapError || 'فایل‌های نقشه را ایمپورت کنید.')}
      </div></div>`;
  }

  function homeView() {
    const orderCards = state.orders.length
      ? state.orders.map((o) => {
          const s = orderStats(o);
          const pct = s.total ? Math.round((s.handled / s.total) * 100) : 0;
          return `
          <article class="card order-card">
            <div class="order-row" data-action="open" data-id="${o.id}" role="button" tabindex="0">
              <div class="order-ico">${ICON.box}</div>
              <div class="order-meta">
                <h3>${esc(o.name)}</h3>
                <p>${fa(s.handled)} از ${fa(s.total)} قلم · ${dateFa(o.createdAt)}</p>
              </div>
              <button class="del" data-action="delete" data-id="${o.id}" aria-label="حذف سفارش">${ICON.trash}</button>
            </div>
            <div class="bar"><i style="width:${pct}%"></i></div>
          </article>`;
        }).join('')
      : `<div class="card empty">${ICON.box}<br>هنوز سفارشی ساخته نشده است.<br>فایل سفارش را ایمپورت کنید یا سفارش دستی بسازید.</div>`;

    return `
      ${mapStatusCard()}
      ${state.layout ? `<section class="card search-card"><h3>جستجوی کالا یا کد</h3>${searchFormHTML()}${meBar()}</section>` : ''}
      <div class="action-grid">
        <button class="btn primary" data-action="import-order">${ICON.upload}<span>ایمپورت فایل سفارش</span></button>
        <button class="btn" data-action="manual">${ICON.plus}<span>سفارش دستی</span></button>
        <button class="btn" data-action="import-map">${ICON.map}<span>ایمپورت نقشه و لیوت</span></button>
        ${state.layout ? '' : `<button class="btn" data-action="retry-default">${ICON.map}<span>بارگذاری نقشه پیش‌فرض</span></button>`}
      </div>
      <div class="section-title"><span>سفارش‌ها</span><span>${fa(state.orders.length)}</span></div>
      ${orderCards}`;
  }

  function orderView(order) {
    const s = orderStats(order);
    const pct = s.total ? Math.round((s.handled / s.total) * 100) : 0;
    const cur = currentItem(order);
    subtitle.textContent = 'برداشت سفارش';

    const header = `
      <section class="order-head">
        <div><span class="eyebrow">سفارش</span><h2>${esc(order.name)}</h2></div>
        <div class="progress">
          <div class="progress-row"><span>${fa(s.handled)} از ${fa(s.total)} قلم</span><span>${fa(pct)}٪</span></div>
          <div class="bar"><i style="width:${pct}%"></i></div>
        </div>
      </section>`;

    if (!cur) {
      return `${header}${summaryView(order, s)}${itemsView(order)}`;
    }

    const idxInOrder = order.items.indexOf(cur);
    const locs = locsOf(cur.barcode);
    const primaryIdx = primaryIdxOf(cur);
    const primaryLoc = primaryIdx !== null && primaryIdx !== undefined ? cellAt(primaryIdx) : null;
    const curLoc = Math.min(cur.loc || 0, Math.max(0, locs.length - 1));
    const picked = cur.picked ?? cur.need;
    const isPending = cur.status === 'pending';

    const locBlock = !state.layout
      ? `<p class="warn-text">نقشه بارگذاری نشده است.</p>`
      : primaryLoc
        ? `<div class="loc">
             ${primaryLoc.label ? `<div class="code-box"><span>کد روی کارتن</span><b>${esc(primaryLoc.label)}</b></div>` : ''}
             <div class="loc-pills">
               <span class="pill">${ICON.pin} ${fmtM(primaryLoc.x)} متر از چپ</span>
               <span class="pill">${ICON.pin} ${fmtM(primaryLoc.y)} متر از پایین</span>
             </div>
             ${locs.length > 1 ? `<div class="chips">${locs.map((idx, k) =>
               `<button class="chip ${k === curLoc ? 'on' : ''}" data-action="set-loc" data-i="${k}">محل ${fa(k + 1)}${cellAt(idx).label ? ' · ' + esc(cellAt(idx).label) : ''}</button>`).join('')}</div>` : ''}
           </div>`
        : `<p class="warn-text">این بارکد روی نقشه ثبت نشده است. کالا را دستی پیدا کنید یا نقشه را بررسی کنید.</p>`;

    const focus = `
      <article class="card focus">
        <div class="focus-top">
          <span class="badge">قلم ${fa(idxInOrder + 1)} از ${fa(order.items.length)}</span>
          <span class="status s-${cur.status}">${STATUS_TEXT[cur.status]}</span>
        </div>
        <h3 class="item-name">${esc(cur.name)}</h3>
        <span class="barcode">${esc(cur.barcode)}</span>
        <div class="need"><span>تعداد موردنیاز برای برداشت</span><strong>${fa(cur.need)}</strong><small>عدد</small></div>
        ${locBlock}
      </article>`;

    const mapCard = state.layout ? `
      <article class="card map-card">
        <div class="map-tools">
          <h3>موقعیت در نقشه</h3>
          <button class="btn small ghost" data-action="zoom">بزرگ‌نمایی ${ZOOM_LABEL[state.zoom]}</button>
        </div>
        ${meBar()}
        <div class="map-wrap" id="mapWrap">
          <div class="map-inner" style="width:${ZOOMS[state.zoom] * 100}%">${mapSVG(primaryIdx ?? null, new Set(locs))}</div>
        </div>
        <div class="legend">
          <span><i class="lg-pallet"></i>پالت</span>
          <span><i class="lg-rack"></i>قفسه</span>
          <span><i class="lg-pin"></i>محل کالا</span>
          <span><i class="lg-me"></i>شما</span>
          <span><i class="lg-door"></i>ورودی</span>
        </div>
      </article>` : '';

    const picker = isPending ? `
      <article class="card picker">
        <div class="stepper">
          <span class="lbl">تعداد برداشت‌شده</span>
          <div class="stepper-ctrl">
            <button data-action="step" data-d="-1" aria-label="کم کردن">−</button>
            <output>${fa(picked)}</output>
            <button data-action="step" data-d="1" aria-label="اضافه کردن">+</button>
          </div>
        </div>
        <div class="actions two">
          <button class="btn success" data-action="confirm">تایید برداشت</button>
          <button class="btn danger" data-action="missing">پیدا نشد</button>
        </div>
      </article>` : `
      <article class="card picker">
        <div class="actions">
          <button class="btn" data-action="reopen">بازگرداندن به لیست برداشت</button>
        </div>
      </article>`;

    return `${header}${focus}${mapCard}${picker}${itemsView(order)}`;
  }

  function itemsView(order) {
    const cur = currentItem(order);
    const rows = order.items.map((it, i) => {
      const locs = locsOf(it.barcode);
      const code = locs.length && cellAt(locs[0]).label ? ' · ' + cellAt(locs[0]).label : '';
      const sub = state.layout
        ? (locs.length ? `${fa(locs.length)} محل${code}` : 'بدون موقعیت روی نقشه')
        : 'نقشه بارگذاری نشده';
      return `
        <div class="item-row ${cur && cur.id === it.id ? 'current' : ''}" data-action="focus" data-id="${it.id}" role="button" tabindex="0">
          <span class="n">${fa(i + 1)}</span>
          <div class="t"><b>${esc(it.name)}</b><small>${esc(it.barcode)} · ${sub}</small></div>
          <span class="s ${it.status}">${it.status === 'pending' ? fa(it.need) + ' عدد' : (it.status === 'missing' ? STATUS_TEXT.missing : fa(it.picked ?? 0) + ' از ' + fa(it.need))}</span>
        </div>`;
    }).join('');
    return `<div class="section-title"><span>همه اقلام (به ترتیب مسیر)</span><span>${fa(order.items.length)}</span></div>
            <div class="items">${rows}</div>`;
  }

  function summaryView(order, s) {
    subtitle.textContent = 'خلاصه برداشت';
    return `
      <article class="card done-card">
        <div class="check">${ICON.check}</div>
        <h2>برداشت این سفارش تمام شد</h2>
        <p>همه اقلام بررسی شدند. اقلام کم‌تر از نیاز یا پیدا‌نشده را در لیست زیر ببینید.</p>
        <div class="kpis">
          <div><b style="color:var(--ok)">${fa(s.done)}</b><span>برداشت شد</span></div>
          <div><b style="color:var(--warn)">${fa(s.partial)}</b><span>کمتر از نیاز</span></div>
          <div><b style="color:var(--bad)">${fa(s.missing)}</b><span>پیدا نشد</span></div>
        </div>
        <div class="actions" style="width:100%;margin-top:6px">
          <button class="btn" data-action="review-all">مرور همه اقلام</button>
        </div>
      </article>`;
  }

  function manualView() {
    subtitle.textContent = 'سفارش دستی';
    return `
      <section class="card" style="display:grid;gap:12px">
        <div>
          <h2 style="font-size:17px;font-weight:800">سفارش دستی</h2>
          <p style="color:var(--muted);font-size:13px;margin-top:4px">بارکد کامل یا ۴ رقم آخر بارکد را وارد کنید.</p>
        </div>
        <label class="field">نام سفارش
          <input id="mName" class="input" value="سفارش جدید">
        </label>
        <div id="mRows" style="display:grid;gap:10px">${mRowHTML()}</div>
        <div class="actions two">
          <button class="btn" data-action="add-row">${ICON.plus}<span>افزودن کالا</span></button>
          <button class="btn primary" data-action="save-manual">ساخت سفارش</button>
        </div>
      </section>`;
  }

  const mRowHTML = () => `
    <div class="mrow">
      <input class="input inp-bc" inputmode="numeric" placeholder="بارکد یا ۴ رقم آخر">
      <input class="input inp-qty" type="number" min="1" value="1" aria-label="تعداد">
      <input class="input inp-name" placeholder="نام کالا (اختیاری)">
      <div class="rm"><button class="btn small danger" data-action="rm-row">حذف</button></div>
    </div>`;

  /* ---------- رندر اصلی ---------- */
  function render() {
    const v = state.view;
    btnHome.hidden = v.name === 'home';
    const dark = currentTheme() === 'dark';
    btnTheme.innerHTML = dark ? ICON.sun : ICON.moon;
    btnTheme.setAttribute('aria-label', dark ? 'تم روشن' : 'تم تیره');
    let html;
    if (v.name === 'order') {
      const order = orderById(v.id);
      if (!order) { state.view = { name: 'home' }; return render(); }
      html = orderView(order);
    } else if (v.name === 'search') {
      html = searchView(v.q || '');
    } else if (v.name === 'manual') {
      html = manualView();
    } else {
      subtitle.textContent = 'مدیریت و برداشت سفارش';
      html = homeView();
    }
    app.innerHTML = html;
    afterRender();
  }

  function afterRender() {
    const wrap = document.getElementById('mapWrap');
    if (state.keepScroll) { // بعد از ثبت موقعیت با کلیک روی نقشه، جای اسکرول حفظ شود
      const k = state.keepScroll;
      state.keepScroll = null;
      if (wrap) { wrap.scrollTop = k.top; wrap.scrollLeft = k.left; }
      return;
    }
    if (!wrap) return;
    let pIdx = null;
    let key = '';
    if (state.view.name === 'order') {
      const order = orderById(state.view.id);
      const cur = order && currentItem(order);
      if (!cur) return;
      pIdx = primaryIdxOf(cur);
      key = `${order.id}|${cur.id}|${pIdx}|${state.zoom}`;
    } else if (state.view.name === 'search') {
      pIdx = searchFocus();
      key = `s|${state.view.q}|${pIdx}|${state.zoom}`;
    } else {
      return;
    }
    if (pIdx === null || pIdx === undefined) return;
    if (key === state.lastPinKey) return; // روی انتخاب دیگری اسکرول نشود
    state.lastPinKey = key;
    requestAnimationFrame(() => {
      const svg = wrap.querySelector('svg');
      if (!svg) return;
      const c = cellAt(pIdx);
      const { W, H } = state.layout.geo;
      const sw = svg.getBoundingClientRect().width;
      const sh = svg.getBoundingClientRect().height;
      const vw = W + 2 * MAP_PAD;
      const vh = H + DOOR_SPACE + 2 * MAP_PAD;
      wrap.scrollTop = Math.max(0, ((c.cy + MAP_PAD) / vh) * sh - wrap.clientHeight / 2);
      wrap.scrollLeft = Math.max(0, ((c.cx + MAP_PAD) / vw) * sw - wrap.clientWidth / 2);
    });
  }

  /* ---------- عملیات ---------- */
  function changeItem(order, item, patch) {
    Object.assign(item, patch);
    saveOrders();
  }

  function goNextAfter(order, item) {
    const n = nextPending(order, item.id);
    order.cursor = n ? n.id : null;
  }

  function confirmItem() {
    const order = orderById(state.view.id);
    const item = currentItem(order);
    if (!item) return;
    const picked = item.picked ?? item.need;
    changeItem(order, item, { picked, status: picked >= item.need ? 'done' : 'partial' });
    goNextAfter(order, item);
    saveOrders();
    toast(picked >= item.need ? 'برداشت ثبت شد' : `ثبت شد؛ ${fa(item.need - picked)} عدد کم است`);
    render();
  }

  function markMissing() {
    const order = orderById(state.view.id);
    const item = currentItem(order);
    if (!item) return;
    changeItem(order, item, { picked: 0, status: 'missing' });
    goNextAfter(order, item);
    saveOrders();
    toast('کالا پیدا نشد؛ به قلم بعدی رفتیم', 'error');
    render();
  }

  function reopenItem() {
    const order = orderById(state.view.id);
    const item = currentItem(order);
    if (!item) return;
    changeItem(order, item, { picked: null, status: 'pending' });
    order.cursor = item.id;
    saveOrders();
    render();
  }

  const readWB = async (file) => XLSX.read(await file.arrayBuffer(), { type: 'array', cellStyles: true });
  const isLayoutName = (name) => /layout|لیوت|لی‌اوت|لی اوت/i.test(name);

  async function handleFiles(files) {
    if (typeof XLSX === 'undefined') throw new Error('کتابخانه Excel بارگذاری نشده است. اتصال اینترنت را بررسی کنید.');
    if (state.importTarget === 'map') {
      const layoutFile = files.find((f) => isLayoutName(f.name)) || null;
      const barcodeFile = files.find((f) => !isLayoutName(f.name)) || null;
      if (!barcodeFile) throw new Error('فایل بارکدها (نقشه اصلی) را هم انتخاب کنید.');
      const wbP = await readWB(barcodeFile);
      const wbL = layoutFile ? await readWB(layoutFile) : null;
      state.layout = parseLayout(wbP, barcodeFile.name, wbL, layoutFile ? layoutFile.name : null);
      store.write(KEY_LAYOUT, state.layout);
      state.mapError = '';
      state.lastPinKey = '';
      toast(`نقشه ایمپورت شد: ${fa(state.layout.stats.barcodes)} بارکد${layoutFile ? '، با کدهای لیوت' : ''}`);
      state.view = { name: 'home' };
    } else {
      const wb = await readWB(files[0]);
      const raw = parseOrder(wb);
      const name = files[0].name.replace(/\.(xlsx|xls)$/i, '');
      const order = buildOrder(name, raw);
      state.orders.unshift(order);
      saveOrders();
      state.view = { name: 'order', id: order.id };
      toast(`سفارش ساخته شد: ${fa(order.items.length)} قلم`);
    }
    render();
  }

  async function loadDefaultLayout() {
    if (typeof XLSX === 'undefined') {
      state.mapError = 'کتابخانه Excel بارگذاری نشده است. اتصال اینترنت را بررسی کنید.';
      return render();
    }
    try {
      const res = await fetch(DEFAULT_MAP_URL, { cache: 'no-cache' });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const wbP = XLSX.read(await res.arrayBuffer(), { type: 'array', cellStyles: true });
      let wbL = null;
      try {
        const resL = await fetch(DEFAULT_LAYOUT_URL, { cache: 'no-cache' });
        if (resL.ok) wbL = XLSX.read(await resL.arrayBuffer(), { type: 'array', cellStyles: true });
      } catch (e) { wbL = null; }
      state.layout = parseLayout(wbP, DEFAULT_MAP_NAME, wbL, wbL ? DEFAULT_LAYOUT_NAME : null);
      store.write(KEY_LAYOUT, state.layout);
      state.mapError = '';
    } catch (e) {
      state.mapError = 'نقشه پیش‌فرض باز نشد. برنامه را از طریق سرور محلی اجرا کنید یا با دکمه «ایمپورت نقشه» فایل‌ها را انتخاب کنید.';
    }
    render();
  }

  function saveManual() {
    const name = (document.getElementById('mName').value || '').trim() || 'سفارش دستی';
    const rows = [...document.querySelectorAll('#mRows .mrow')];
    const items = [];
    for (const row of rows) {
      const barcode = normBarcode(row.querySelector('.inp-bc').value);
      if (!barcode) continue;
      const q = parseInt(toEn(row.querySelector('.inp-qty').value), 10);
      items.push({ barcode, name: row.querySelector('.inp-name').value.trim() || 'کالای بدون نام', need: q > 0 ? q : 1 });
    }
    if (!items.length) { toast('حداقل یک بارکد وارد کنید', 'error'); return; }
    const order = buildOrder(name, items);
    state.orders.unshift(order);
    saveOrders();
    state.view = { name: 'order', id: order.id };
    render();
  }

  /* ---------- رویدادها ---------- */
  document.addEventListener('click', (e) => {
    // کلیک روی نقشه = ثبت موقعیت من (نزدیک‌ترین خانه‌ی کد‌دار)
    const svgEl = e.target.closest ? e.target.closest('.map-svg') : null;
    if (svgEl && state.layout) {
      const ctm = svgEl.getScreenCTM();
      if (ctm) {
        const pt = svgEl.createSVGPoint();
        pt.x = e.clientX;
        pt.y = e.clientY;
        const p = pt.matrixTransform(ctm.inverse());
        let best = -1;
        let bd = Infinity;
        state.layout.cells.forEach((c, i) => {
          if (!c.kind) return;
          const d = (c.cx - p.x) ** 2 + (c.cy - p.y) ** 2;
          if (d < bd) { bd = d; best = i; }
        });
        if (best >= 0) {
          const c = state.layout.cells[best];
          const near = bd <= (Math.max(c.w, c.h) * 1.5) ** 2;
          if (near) {
            const wrap = document.getElementById('mapWrap');
            state.keepScroll = wrap ? { top: wrap.scrollTop, left: wrap.scrollLeft } : null;
            setMe(`${c.r}-${c.c}`);
            toast('موقعیت شما ثبت شد');
            render();
          }
        }
      }
      return;
    }

    const el = e.target.closest('[data-action]');
    if (!el) return;
    const a = el.dataset.action;
    switch (a) {
      case 'search-focus':
        state.view.focus = Number(el.dataset.i);
        state.lastPinKey = '';
        return render();
      case 'me-clear':
        setMe(null);
        state.lastPinKey = '';
        return render();
      case 'open':
        state.view = { name: 'order', id: el.dataset.id };
        state.lastPinKey = '';
        return render();
      case 'delete': {
        const o = orderById(el.dataset.id);
        if (o && confirm(`سفارش «${o.name}» حذف شود؟`)) {
          state.orders = state.orders.filter((x) => x.id !== o.id);
          saveOrders();
          toast('سفارش حذف شد');
          render();
        }
        return;
      }
      case 'theme':
        return toggleTheme();
      case 'home':
        state.view = { name: 'home' };
        return render();
      case 'manual':
        state.view = { name: 'manual' };
        return render();
      case 'import-map':
        state.importTarget = 'map';
        fileInput.multiple = true;
        return fileInput.click();
      case 'import-order':
        state.importTarget = 'order';
        fileInput.multiple = false;
        return fileInput.click();
      case 'retry-default':
        return loadDefaultLayout();
      case 'confirm':
        return confirmItem();
      case 'missing':
        return markMissing();
      case 'reopen':
        return reopenItem();
      case 'focus': {
        const order = orderById(state.view.id);
        if (order) { order.cursor = el.dataset.id; state.lastPinKey = ''; }
        return render();
      }
      case 'step': {
        const order = orderById(state.view.id);
        const item = currentItem(order);
        if (!item) return;
        const cur = item.picked ?? item.need;
        changeItem(order, item, { picked: Math.max(0, cur + Number(el.dataset.d)) });
        return render();
      }
      case 'set-loc': {
        const order = orderById(state.view.id);
        const item = currentItem(order);
        if (!item) return;
        changeItem(order, item, { loc: Number(el.dataset.i) });
        state.lastPinKey = '';
        return render();
      }
      case 'zoom':
        state.zoom = (state.zoom + 1) % ZOOMS.length;
        state.lastPinKey = '';
        return render();
      case 'review-all': {
        const order = orderById(state.view.id);
        if (order) { order.cursor = order.items[0].id; state.lastPinKey = ''; }
        return render();
      }
      case 'add-row':
        document.getElementById('mRows').insertAdjacentHTML('beforeend', mRowHTML());
        return;
      case 'rm-row': {
        const rows = document.querySelectorAll('#mRows .mrow');
        if (rows.length > 1) el.closest('.mrow').remove();
        return;
      }
      case 'save-manual':
        return saveManual();
      default:
        return;
    }
  });

  document.addEventListener('keydown', (e) => {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('[role="button"][data-action]')) {
      e.preventDefault();
      e.target.click();
    }
  });

  document.addEventListener('submit', (e) => {
    if (e.target.id !== 'searchForm') return;
    e.preventDefault();
    const q = (document.getElementById('qInput').value || '').trim();
    if (!q) return;
    state.view = { name: 'search', q };
    state.lastPinKey = '';
    render();
  });

  fileInput.addEventListener('change', async () => {
    const files = [...(fileInput.files || [])];
    fileInput.value = '';
    if (!files.length) return;
    try { await handleFiles(files); }
    catch (err) { toast(err.message || 'خطا در خواندن فایل', 'error'); }
  });

  /* ---------- شروع ---------- */
  render();
  if (!state.layout) loadDefaultLayout();
})();
