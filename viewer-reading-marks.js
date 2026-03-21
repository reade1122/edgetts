/**
 * PDF 區塊標記（對齊語音報讀系統個人版）
 * 標記序列化：PDF Subject = PDF_MARKS_DATA:{base64(json)}
 */
(function viewerReadingMarks() {
  const marks = {};
  let markIdCounter = 0;
  let isMarkingMode = false;
  let isManualMarkingEnabled = false;
  let currentDraw = null;
  let originalPdfBytes = null;
  let isCloudFileMode = false;
  let markDrawBound = false;
  let pdfPasswordHash = null;
  let currentSessionPassword = null;
  // 雲端（無法回寫）時：進入標記模式先備份 marks，取消標記時還原避免留下修改痕跡
  let cloudMarksBackup = null;
  let cloudMarkIdCounterBackup = 0;
  /** 已對該份 PDF（fingerprint）做過「軟重新載入」，避免 pagesloaded 迴圈 */
  let markViewerSoftReloadFingerprint = null;

  const VIEWER_IDB_DB_NAME = 'pdfjs-viewer';
  const VIEWER_IDB_STORE_NAME = 'files';
  const VIEWER_IDB_LAST_FILE_KEY = 'last-opened-file';
  /** 由 IndexedDB last-opened-file 紀錄的 markSourceCloud 快取（與 PDF 同存，不用 localStorage） */
  let idbMarkSourceCloudCached = false;

  function openViewerDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(VIEWER_IDB_DB_NAME, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(VIEWER_IDB_STORE_NAME)) {
          db.createObjectStore(VIEWER_IDB_STORE_NAME);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  function mkey(pageNum) {
    return String(pageNum);
  }

  function getApp() {
    return window.PDFViewerApplication;
  }

  async function refreshIdbMarkSourceCloudCache() {
    try {
      const db = await openViewerDB();
      const rec = await new Promise((resolve, reject) => {
        const tx = db.transaction(VIEWER_IDB_STORE_NAME, 'readonly');
        const req = tx.objectStore(VIEWER_IDB_STORE_NAME).get(VIEWER_IDB_LAST_FILE_KEY);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      });
      db.close();
      idbMarkSourceCloudCached = !!(rec && rec.markSourceCloud);
    } catch (e) {
      idbMarkSourceCloudCached = false;
    }
  }

  function updateCloudMode() {
    const app = getApp();
    const url = (app && app.url) || '';
    const fromHttpUrl = url.startsWith('http://') || url.startsWith('https://');
    // 本機挑檔／blob 開檔時 app.url 為 blob: 或 file:；open({data}) 時 url 常為空但绝不可沿用
    // IDB「上次檔」的 markSourceCloud，否則會誤判雲端而軟重載，變成本機也要 F5。
    if (url.startsWith('blob:') || url.startsWith('file:')) {
      idbMarkSourceCloudCached = false;
      isCloudFileMode = false;
      return;
    }
    isCloudFileMode = fromHttpUrl || idbMarkSourceCloudCached;
  }

  async function refreshOriginalPdfBytes() {
    const app = getApp();
    if (!app?.pdfDocument) {
      originalPdfBytes = null;
      return;
    }
    try {
      const data = await app.pdfDocument.getData();
      originalPdfBytes = data.slice(0);
    } catch (e) {
      console.warn('無法取得 PDF 位元組', e);
      originalPdfBytes = null;
    }
  }

  async function savePdfBytesToViewerIdb(uint8) {
    const app = getApp();
    const name =
      (app && app._docFilename) ||
      (app && app.url && !app.url.startsWith('blob:') ? app.url.split('/').pop() : null) ||
      'document.pdf';
    const buffer = uint8.buffer.slice(uint8.byteOffset, uint8.byteOffset + uint8.byteLength);
    const file = new File([buffer], name, { type: 'application/pdf' });
    const db = await openViewerDB();
    try {
      const existing = await new Promise((resolve, reject) => {
        const tx = db.transaction(VIEWER_IDB_STORE_NAME, 'readonly');
        const req = tx.objectStore(VIEWER_IDB_STORE_NAME).get(VIEWER_IDB_LAST_FILE_KEY);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => reject(req.error);
      });
      const buf = await file.arrayBuffer();
      await new Promise((resolve, reject) => {
        const tx = db.transaction(VIEWER_IDB_STORE_NAME, 'readwrite');
        const store = tx.objectStore(VIEWER_IDB_STORE_NAME);
        store.put(
          {
            name: file.name,
            type: file.type || 'application/pdf',
            data: buf,
            markSourceCloud: existing?.markSourceCloud === true,
          },
          VIEWER_IDB_LAST_FILE_KEY
        );
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
    } finally {
      db.close();
    }
  }

  function getMarkLayerViewport(pageNum) {
    const layer = document.querySelector(`.mark-layer[data-page-num="${pageNum}"]`);
    if (!layer) return null;
    const scale = parseFloat(layer.dataset.scale);
    if (!scale || scale <= 0) return null;
    const viewportWidth = parseFloat(layer.style.width) || layer.clientWidth;
    const viewportHeight = parseFloat(layer.style.height) || layer.clientHeight;
    if (!viewportWidth || !viewportHeight) return null;
    return { scale, viewportWidth, viewportHeight };
  }

  function viewportToPdfCoords(x, y, width, height, pageNum) {
    const v = getMarkLayerViewport(pageNum);
    if (!v) return null;
    const { scale, viewportHeight } = v;
    return {
      pdfL: x / scale,
      pdfT: (viewportHeight - y) / scale,
      pdfR: (x + width) / scale,
      pdfB: (viewportHeight - y - height) / scale,
    };
  }

  function pdfToViewportCoords(pdfL, pdfT, pdfR, pdfB, pageNum) {
    const v = getMarkLayerViewport(pageNum);
    if (!v) return null;
    const { scale, viewportHeight } = v;
    return {
      left: pdfL * scale,
      top: viewportHeight - pdfT * scale,
      width: (pdfR - pdfL) * scale,
      height: (pdfT - pdfB) * scale,
    };
  }

  function ensureMarkLayer(pageNum) {
    const pageDiv = document.querySelector(`#viewerContainer .page[data-page-number="${pageNum}"]`);
    if (!pageDiv) return null;
    let layer = pageDiv.querySelector(':scope > .mark-layer');
    if (!layer) {
      layer = document.createElement('div');
      layer.className = 'mark-layer';
      layer.dataset.pageNum = String(pageNum);
    }
    // 一律置於 .page 最後，避免 PDF.js 重繪後把 textLayer 等插到標記層之上，首開檔點擊被文字層攔截而無法朗讀
    pageDiv.appendChild(layer);
    const app = getApp();
    const pv = app?.pdfViewer?.getPageView(pageNum - 1);
    const vp = pv?.viewport;
    if (vp) {
      layer.dataset.scale = String(vp.scale);
      layer.style.width = `${vp.width}px`;
      layer.style.height = `${vp.height}px`;
    }
    layer.style.position = 'absolute';
    layer.style.left = '0';
    layer.style.top = '0';
    layer.style.zIndex = '8000';
    layer.style.overflow = 'visible';
    if (isMarkingMode) {
      layer.classList.add('marking-mode');
      layer.style.pointerEvents = 'auto';
      layer.style.display = 'block';
    } else {
      layer.classList.remove('marking-mode');
      layer.style.pointerEvents = 'none';
    }
    return layer;
  }

  /** PDF.js 若在標記層之後又插入 textLayer，會變成 false，需再 bump 一次 */
  function isMarkLayerLastInPage(pageNum) {
    const pageDiv = document.querySelector(`#viewerContainer .page[data-page-number="${pageNum}"]`);
    if (!pageDiv) return true;
    const layer = pageDiv.querySelector(':scope > .mark-layer');
    if (!layer) return true;
    return pageDiv.lastElementChild === layer;
  }

  function syncAllMarkLayersSize() {
    const app = getApp();
    const n = app?.pagesCount || 0;
    for (let p = 1; p <= n; p++) ensureMarkLayer(p);
  }

  function spanOverlapFraction(frameRect, spanRect) {
    const il = Math.max(frameRect.left, spanRect.left);
    const ir = Math.min(frameRect.right, spanRect.right);
    const it = Math.max(frameRect.top, spanRect.top);
    const ib = Math.min(frameRect.bottom, spanRect.bottom);
    const iw = ir - il;
    const ih = ib - it;
    if (iw <= 0 || ih <= 0) return 0;
    const area = iw * ih;
    const sa = Math.max(1, spanRect.width * spanRect.height);
    return area / sa;
  }

  function iouClientRects(a, b) {
    const il = Math.max(a.left, b.left);
    const ir = Math.min(a.right, b.right);
    const it = Math.max(a.top, b.top);
    const ib = Math.min(a.bottom, b.bottom);
    const iw = Math.max(0, ir - il);
    const ih = Math.max(0, ib - it);
    const inter = iw * ih;
    if (inter <= 0) return 0;
    const areaA = Math.max(1e-6, (a.right - a.left) * (a.bottom - a.top));
    const areaB = Math.max(1e-6, (b.right - b.left) * (b.bottom - b.top));
    return inter / (areaA + areaB - inter);
  }

  function normalizeSpanTextForDedupe(span) {
    return (span.textContent || '').replace(/\s+/g, ' ').trim();
  }

  /**
   * PDF 常以多層文字疊加模擬粗體；textLayer 會出現多個位置與內容幾乎相同的 span。
   * 在擷取框內只保留閱讀序較前的一個，避免「國小國小國小國小」。
   */
  function dedupeBoldOverlaySpans(spans) {
    const sorted = [...spans].sort((a, b) => {
      const ra = a.getBoundingClientRect();
      const rb = b.getBoundingClientRect();
      const dy = ra.top - rb.top;
      if (Math.abs(dy) > 6) return dy;
      return ra.left - rb.left;
    });
    const kept = [];
    for (const span of sorted) {
      const r = span.getBoundingClientRect();
      const t = normalizeSpanTextForDedupe(span);
      if (!t) continue;
      let duplicate = false;
      for (const k of kept) {
        const kr = k.getBoundingClientRect();
        const kt = normalizeSpanTextForDedupe(k);
        if (t !== kt) continue;
        if (iouClientRects(r, kr) >= 0.5) {
          duplicate = true;
          break;
        }
        const cxa = (r.left + r.right) / 2;
        const cya = (r.top + r.bottom) / 2;
        const cxb = (kr.left + kr.right) / 2;
        const cyb = (kr.top + kr.bottom) / 2;
        const dh = Math.min(r.height, kr.height);
        if (Math.hypot(cxa - cxb, cya - cyb) < Math.max(2.5, dh * 0.06)) {
          duplicate = true;
          break;
        }
      }
      if (!duplicate) kept.push(span);
    }
    return kept;
  }

  /**
   * 若單一 span 內仍為連續重複片段（無空格銜接），收斂為一段。
   */
  function collapseAdjacentDuplicateCjkRuns(s) {
    if (!s || s.length < 2) return s;
    let t = s;
    let prev = '';
    let guard = 0;
    const cjkSpan = String.raw`(?:[\u3000-\u9fff\u2460-\u2473])`;
    while (prev !== t && guard++ < 24) {
      prev = t;
      t = t.replace(
        new RegExp(`(${cjkSpan}{2,}?)(?:\\1)+`, 'gu'),
        '$1'
      );
      t = t.replace(new RegExp(`(${cjkSpan})(?:\\1)+`, 'gu'), '$1');
    }
    return t;
  }

  function extractTextFromMark(mark, pageNum) {
    const markLayer = document.querySelector(`.mark-layer[data-page-num="${pageNum}"]`);
    const textLayer = document.querySelector(`.page[data-page-number="${pageNum}"] .textLayer`);
    if (!markLayer || !textLayer) return '';
    const lr = markLayer.getBoundingClientRect();
    const left = lr.left + mark.x;
    const top = lr.top + mark.y;
    const right = left + mark.width;
    const bottom = top + mark.height;
    const frameRect = { left, top, right, bottom };

    const rawSpans = Array.from(textLayer.querySelectorAll('span')).filter(span => {
      const sr = span.getBoundingClientRect();
      return !(sr.right < left || sr.left > right || sr.bottom < top || sr.top > bottom);
    });
    const spans = dedupeBoldOverlaySpans(rawSpans);
    spans.sort((a, b) => {
      const ra = a.getBoundingClientRect();
      const rb = b.getBoundingClientRect();
      const dy = ra.top - rb.top;
      if (Math.abs(dy) > 6) return dy;
      return ra.left - rb.left;
    });

    let out = '';
    for (const span of spans) {
      const sr = span.getBoundingClientRect();
      const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
      let node;
      while ((node = walker.nextNode())) {
        const txt = node.nodeValue || '';
        for (let i = 0; i < txt.length; i++) {
          const range = document.createRange();
          range.setStart(node, i);
          range.setEnd(node, i + 1);
          const rects = range.getClientRects();
          for (const r of rects) {
            if (spanOverlapFraction(frameRect, r) >= 0.45) {
              out += txt[i];
            }
          }
        }
      }
    }
    out = out.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
    out = collapseAdjacentDuplicateCjkRuns(out);
    return out.replace(/\s+/g, ' ').trim();
  }

  async function hashPassword(password) {
    if (!password) return null;
    const encoder = new TextEncoder();
    const data = encoder.encode(password);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  }

  async function verifyPassword(inputPassword) {
    if (!pdfPasswordHash) return true;
    if (!inputPassword) return false;
    const inputHash = await hashPassword(inputPassword);
    return inputHash === pdfPasswordHash;
  }

  async function promptPassword(
    title = '輸入密碼',
    text = '此 PDF 標記已受密碼保護，請輸入密碼以繼續'
  ) {
    await ensureSwalLocal();
    const result = await window.Swal.fire({
      title,
      text,
      input: 'password',
      inputPlaceholder: '請輸入密碼',
      showCancelButton: true,
      confirmButtonText: '確定',
      cancelButtonText: '取消',
      customClass: { popup: 'swal-high-z-index' },
      inputValidator: value => {
        if (!value) return '請輸入密碼';
        return undefined;
      },
    });
    if (result.isConfirmed && result.value) return result.value;
    return null;
  }

  async function checkPasswordProtection() {
    if (!pdfPasswordHash) return true;
    if (currentSessionPassword) {
      if (await verifyPassword(currentSessionPassword)) return true;
      currentSessionPassword = null;
    }
    const password = await promptPassword('輸入密碼', '此 PDF 標記已受密碼保護，請輸入密碼以編輯標記');
    if (!password) return false;
    if (await verifyPassword(password)) {
      currentSessionPassword = password;
      return true;
    }
    await ensureSwalLocal();
    window.Swal.fire({
      icon: 'error',
      title: '密碼錯誤',
      text: '輸入的密碼不正確',
      timer: 2000,
      showConfirmButton: false,
      customClass: { popup: 'swal-high-z-index' },
    });
    return false;
  }

  function normalizeKeywordsArray(kw) {
    let keywordsArray = kw || [];
    if (typeof keywordsArray === 'string') {
      keywordsArray = keywordsArray ? [keywordsArray] : [];
    } else if (!Array.isArray(keywordsArray)) {
      keywordsArray = [];
    }
    return keywordsArray;
  }

  function readPasswordHashFromKeywords(keywordsArray) {
    const passwordKeyword = keywordsArray.find(k => String(k).includes('PDF_MARKS_PASSWORD:'));
    if (passwordKeyword) {
      const m = String(passwordKeyword).match(/PDF_MARKS_PASSWORD:([a-f0-9]{64})/);
      if (m) return m[1];
    }
    return null;
  }

  async function ensureSwalLocal() {
    if (window.Swal) return;
    await new Promise(resolve => {
      const s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js';
      s.onload = () => resolve();
      s.onerror = () => resolve();
      document.head.appendChild(s);
    });
  }

  async function waitForPDFLibReady(maxWaitMs = 4000) {
    if (typeof PDFLib !== 'undefined') return true;
    const start = Date.now();
    while (Date.now() - start < maxWaitMs) {
      await new Promise(resolve => setTimeout(resolve, 80));
      if (typeof PDFLib !== 'undefined') return true;
    }
    return false;
  }

  function hasMarksData() {
    return Object.keys(marks).some(k => Array.isArray(marks[k]) && marks[k].length > 0);
  }

  /** 從目前載入的 PDF 重新同步標記（雲端／首次開檔時序問題時，在「手動標記」點選時再讀一次） */
  async function reloadMarksFromPdfDocument() {
    await refreshIdbMarkSourceCloudCache();
    updateCloudMode();
    await refreshOriginalPdfBytes();
    await waitForPDFLibReady(4500);
    await loadMarksFromPdf();
    if (!hasMarksData()) {
      await new Promise(r => setTimeout(r, 280));
      await refreshOriginalPdfBytes();
      await loadMarksFromPdf();
    }
  }

  async function saveMarksToPdf(showNotification = true) {
    const app = getApp();
    if (!app?.pdfDocument || !originalPdfBytes) {
      if (showNotification) {
        await ensureSwalLocal();
        window.Swal?.fire({ icon: 'warning', title: '無法保存', text: '請先載入 PDF' });
      }
      return;
    }
    if (isCloudFileMode) {
      if (showNotification) {
        await ensureSwalLocal();
        window.Swal?.fire({
          icon: 'info',
          title: '不支援雲端檔案保存',
          html:
            '<div style="text-align:left;line-height:1.6">' +
            '<p style="margin:0 0 10px 0">目前開啟的是<strong>雲端／網址 PDF</strong>，基於瀏覽器安全限制，無法將標記寫回遠端檔案。</p>' +
            '<p style="margin:0">請改為<strong>下載 PDF 後以本機開啟</strong>，或使用先前已快取於本機的副本，即可使用「保存標記」。</p>' +
            '</div>',
          confirmButtonText: '我知道了',
          customClass: { popup: 'swal-high-z-index' },
        });
      }
      return;
    }
    if (typeof PDFLib === 'undefined') {
      if (showNotification) {
        await ensureSwalLocal();
        window.Swal?.fire({ icon: 'error', title: '缺少 PDF-Lib', text: '請重新整理頁面' });
      }
      return;
    }
    try {
      if (showNotification) {
        await ensureSwalLocal();
        window.Swal?.fire({
          title: '正在保存',
          text: '正在將標記寫入 PDF…',
          allowOutsideClick: false,
          didOpen: () => window.Swal.showLoading(),
        });
      }
      const { PDFDocument } = PDFLib;
      const pdfDocLib = await PDFDocument.load(originalPdfBytes);
      const marksJson = JSON.stringify(marks);
      const encodedMarks = btoa(unescape(encodeURIComponent(marksJson)));
      pdfDocLib.setSubject(`PDF_MARKS_DATA:${encodedMarks}`);
      let keywordsArray = normalizeKeywordsArray(pdfDocLib.getKeywords());
      if (pdfPasswordHash) {
        keywordsArray = keywordsArray.filter(k => !String(k).includes('PDF_MARKS_PASSWORD:'));
        keywordsArray.push(`PDF_MARKS_PASSWORD:${pdfPasswordHash}`);
        pdfDocLib.setKeywords(keywordsArray);
      } else {
        keywordsArray = keywordsArray.filter(k => !String(k).includes('PDF_MARKS_PASSWORD:'));
        pdfDocLib.setKeywords(keywordsArray);
      }
      const nextBytes = await pdfDocLib.save();
      const u8 = nextBytes instanceof Uint8Array ? nextBytes : new Uint8Array(nextBytes);
      originalPdfBytes = u8.slice(0);
      await savePdfBytesToViewerIdb(originalPdfBytes);
      if (showNotification) {
        window.Swal?.fire({
          icon: 'success',
          title: '已保存',
          text: '標記已寫入 PDF 中繼資料',
          timer: 1400,
          showConfirmButton: false,
        });
      } else {
        window.Swal?.close?.();
      }
    } catch (e) {
      console.error(e);
      await ensureSwalLocal();
      window.Swal?.fire({ icon: 'error', title: '保存失敗', text: e.message || String(e) });
    }
  }

  function applyLoadedMarksObject(loaded) {
    Object.keys(marks).forEach(k => delete marks[k]);
    Object.keys(loaded || {}).forEach(pk => {
      marks[pk] = loaded[pk];
      if (Array.isArray(marks[pk])) {
        marks[pk].forEach(m => {
          if (m.replacementText === undefined) m.replacementText = '';
          if (m.locked === undefined) m.locked = false;
        });
      }
    });
    const flat = Object.values(marks).flat();
    markIdCounter = flat.length ? Math.max(...flat.map(x => x.id)) + 1 : 0;
  }

  function parseMarksFromSubject(subject) {
    if (!subject || !String(subject).startsWith('PDF_MARKS_DATA:')) return null;
    const encoded = String(subject).substring('PDF_MARKS_DATA:'.length);
    const json = decodeURIComponent(escape(atob(encoded)));
    return JSON.parse(json);
  }

  async function loadMarksFromPdf() {
    const app = getApp();
    // 優先走 PDF.js metadata：不依賴 PDFLib 載入時序，避免首次載入抓不到已存在標記
    try {
      const meta = await app?.pdfDocument?.getMetadata?.();
      const subject =
        meta?.info?.Subject ||
        meta?.info?.subject ||
        meta?.metadata?.get?.('dc:subject') ||
        '';
      const loadedFromMeta = parseMarksFromSubject(subject);
      if (loadedFromMeta) {
        applyLoadedMarksObject(loadedFromMeta);
        return;
      }
    } catch (e) {
      // metadata 讀取失敗時再退回 PDFLib 路徑
    }

    if (!originalPdfBytes || typeof PDFLib === 'undefined') return;
    try {
      const { PDFDocument } = PDFLib;
      const pdfDocLib = await PDFDocument.load(originalPdfBytes);
      let keywordsArray = normalizeKeywordsArray(pdfDocLib.getKeywords());
      const ph = readPasswordHashFromKeywords(keywordsArray);
      pdfPasswordHash = ph;
      if (!ph) currentSessionPassword = null;

      const subject = pdfDocLib.getSubject() || '';
      const loaded = parseMarksFromSubject(subject);
      if (!loaded) {
        Object.keys(marks).forEach(k => delete marks[k]);
        markIdCounter = 0;
        return;
      }
      applyLoadedMarksObject(loaded);
    } catch (e) {
      console.warn('載入標記失敗', e);
      Object.keys(marks).forEach(k => delete marks[k]);
      markIdCounter = 0;
      pdfPasswordHash = null;
      currentSessionPassword = null;
    }
  }

  function deleteMarkBlock(markId, pageNum) {
    const k = mkey(pageNum);
    if (!marks[k]) return;
    marks[k] = marks[k].filter(m => m.id !== markId);
    if (marks[k].length === 0) delete marks[k];
    renderMarksPage(pageNum);
    saveMarksToPdf(false);
  }

  /** 與 viewer-reading 框選辨識相同：canvas 內部像素與 CSS 顯示尺寸對齊後裁切 */
  function cropPageCanvasByClientRect(pageCanvas, cropClientRect) {
    const canvasRect = pageCanvas.getBoundingClientRect();
    const interLeft = Math.max(cropClientRect.left, canvasRect.left);
    const interRight = Math.min(cropClientRect.right, canvasRect.right);
    const interTop = Math.max(cropClientRect.top, canvasRect.top);
    const interBottom = Math.min(cropClientRect.bottom, canvasRect.bottom);
    const interWidth = interRight - interLeft;
    const interHeight = interBottom - interTop;
    if (interWidth <= 5 || interHeight <= 5) return null;
    const scaleX = pageCanvas.width / canvasRect.width;
    const scaleY = pageCanvas.height / canvasRect.height;
    const sx = (interLeft - canvasRect.left) * scaleX;
    const sy = (interTop - canvasRect.top) * scaleY;
    const sw = interWidth * scaleX;
    const sh = interHeight * scaleY;
    const out = document.createElement('canvas');
    out.width = Math.max(64, Math.round(sw));
    out.height = Math.max(64, Math.round(sh));
    const ctx = out.getContext('2d');
    ctx.drawImage(pageCanvas, sx, sy, sw, sh, 0, 0, out.width, out.height);
    return out;
  }

  function canvasToBase64Png(canvas) {
    return canvas.toDataURL('image/png').split(',')[1];
  }

  function isCanvasAllWhite(canvas) {
    const ctx = canvas.getContext('2d');
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height).data;
    for (let i = 0; i < imageData.length; i += 4) {
      if (
        !(
          imageData[i] === 255 &&
          imageData[i + 1] === 255 &&
          imageData[i + 2] === 255 &&
          imageData[i + 3] === 255
        )
      ) {
        return false;
      }
    }
    return true;
  }

  function unlockLocalSpeechForIOS() {
    try {
      if (window.speechSynthesis && typeof window.speechSynthesis.resume === 'function') {
        window.speechSynthesis.resume();
      }
    } catch (e) {
      /* ignore */
    }
  }

  function getRatePercentFromMarkUi() {
    const speedInput = document.getElementById('speakSpeed');
    return Math.max(-50, Math.min(50, parseFloat(speedInput?.value || '0') || 0));
  }

  function getMarkClientFrameRect(mark, pageNum) {
    const markLayer = document.querySelector(`.mark-layer[data-page-num="${pageNum}"]`);
    if (!markLayer) return null;
    const lr = markLayer.getBoundingClientRect();
    return {
      left: lr.left + mark.x,
      top: lr.top + mark.y,
      right: lr.left + mark.x + mark.width,
      bottom: lr.top + mark.y + mark.height,
    };
  }

  function speakMarkWithHighlight(mark, pageNum, text, extra = {}) {
    const frame = getMarkClientFrameRect(mark, pageNum);
    const select = document.getElementById('voiceSelect');
    const voiceName = select ? select.value : 'local-zh-female';
    const ratePercent = getRatePercentFromMarkUi();
    if (typeof window.speakMarkRegionWithHighlight === 'function' && frame) {
      return window.speakMarkRegionWithHighlight(pageNum, frame, text, {
        voiceName,
        ratePercent,
        onEnd: extra.onEnd,
      });
    }
    sendTextToTTSForMark(text, extra.onEnd || null);
  }

  function sendTextToTTSForMark(text, onEnd) {
    const select = document.getElementById('voiceSelect');
    const voiceName = select ? select.value : 'local-zh-female';
    const ratePercent = getRatePercentFromMarkUi();
    if (window.sendTextToTTS) {
      window.sendTextToTTS(text, onEnd || null, { voiceName, ratePercent });
    }
  }

  function normalizeOriginalTextForDisplay(text) {
    const raw = (text || '').trim();
    if (!raw) return '';
    const normalizeFn = window.normalizeCircledDigitsForTts;
    if (typeof normalizeFn === 'function') {
      try {
        return (normalizeFn(raw) || '').trim();
      } catch (_) {
        return raw;
      }
    }
    return raw;
  }

  async function editMarkReplacementText(markId, pageNum) {
    const k = mkey(pageNum);
    const mark = marks[k]?.find(m => m.id === markId);
    if (!mark) return;

    // 以「編輯視窗當下看到的原始文字」作為後續朗讀基準（即使未按儲存替換文字也要落檔）。
    const textBeforeEdit = normalizeOriginalTextForDisplay(mark.text || '');
    const extractedNow = normalizeOriginalTextForDisplay(extractTextFromMark(mark, pageNum) || '');
    let displayText = extractedNow || textBeforeEdit || '(無文字)';
    const effectiveOriginalText = displayText === '(無文字)' ? '' : displayText;
    const originalTextChanged = effectiveOriginalText !== textBeforeEdit;
    if (originalTextChanged) {
      mark.text = effectiveOriginalText;
    }
    if (displayText !== '(無文字)') {
      displayText = displayText.replace(/^[\s\t\u3000]+/, '');
    }

    await ensureSwalLocal();

    const replacementPreviewUi = { reset: () => {} };

    const result = await window.Swal.fire({
      title: '編輯替換文字',
      html: `
            <div style="text-align: left; margin-bottom: 20px; width: 100%; box-sizing: border-box;">
                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; width: 95%; box-sizing: border-box;">
                    <label style="font-weight: bold; font-size: 18px; color: #333; margin: 0;">原始文字：</label>
                    <div style="display: flex; gap: 8px;">
                        <button id="aiImageRecognizeBtn" type="button" style="padding: 6px 12px; background-color: #6f42c1; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: normal; white-space: nowrap; pointer-events: auto; z-index: 10000; position: relative;" title="以 AI 辨識標記區域內的圖片，並將結果填入原始文字">AI圖片辨識</button>
                        <button id="copyOriginalTextBtn" type="button" onclick="window.copyOriginalToReplacement && window.copyOriginalToReplacement()" style="padding: 6px 12px; background-color: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: normal; white-space: nowrap; pointer-events: auto; z-index: 10000; position: relative;">複製到替換文字</button>
                    </div>
                </div>
                <div id="originalTextDisplay" style="width: 95%; margin-left: 0; margin-right: auto; padding: 12px 12px 12px 0; background-color: #f5f5f5; border: 1px solid #ddd; border-radius: 6px; min-height: 50px; max-height: 300px; overflow-y: auto; word-wrap: break-word; word-break: break-all; font-size: 15px; line-height: 1.8; color: #555; box-sizing: border-box; white-space: pre-wrap; font-family: inherit; cursor: pointer;" title="點擊此處發音">${displayText.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div>
            </div>
            <div style="text-align: left; width: 100%; box-sizing: border-box;">
                <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; width: 95%; box-sizing: border-box;">
                    <label for="replacementTextInput" style="font-weight: bold; font-size: 18px; color: #333; margin: 0;">替換文字（留空則使用原始文字）：</label>
                    <button id="previewReplacementTextBtn" type="button" style="padding: 6px 12px; background-color: #28a745; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: normal; white-space: nowrap; pointer-events: auto; z-index: 10000; position: relative;">替換文字試聽</button>
                </div>
                <textarea id="replacementTextInput" class="swal2-textarea" placeholder="輸入替換文字" style="width: 95%; margin-left: 0; margin-right: auto; min-height: 80px; max-height: 200px; resize: vertical; box-sizing: border-box; font-size: 15px; padding: 12px; border: 1px solid #ccc; border-radius: 6px; word-wrap: break-word; word-break: break-word; white-space: pre-wrap; overflow-wrap: break-word; overflow-x: hidden; overflow-y: auto; line-height: 1.6; font-family: inherit;">${(mark.replacementText || '').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</textarea>
            </div>
        `,
      width: '90%',
      maxWidth: '600px',
      showCancelButton: true,
      confirmButtonText: '儲存',
      cancelButtonText: '取消',
      customClass: { popup: 'swal-high-z-index' },
      didOpen: () => {
        const input = document.getElementById('replacementTextInput');
        if (input) {
          input.focus();
          input.select();
        }

        window.copyOriginalToReplacement = function () {
          const replacementInput = document.getElementById('replacementTextInput');
          const copyBtn = document.getElementById('copyOriginalTextBtn');
          if (replacementInput && copyBtn) {
            replacementInput.value = displayText;
            replacementInput.focus();
            const originalText = copyBtn.textContent;
            copyBtn.textContent = '已複製！';
            copyBtn.style.backgroundColor = '#28a745';
            setTimeout(() => {
              copyBtn.textContent = originalText;
              copyBtn.style.backgroundColor = '#007bff';
            }, 1500);
          }
        };

        window.runAiImageRecognizeForEdit = async function () {
          const originalDisplay = document.getElementById('originalTextDisplay');
          const aiBtn = document.getElementById('aiImageRecognizeBtn');
          if (!originalDisplay || !aiBtn) return;
          const markLayer = document.querySelector(`.mark-layer[data-page-num="${pageNum}"]`);
          if (!markLayer) {
            window.Swal.fire({
              icon: 'warning',
              title: '無法取得標記',
              text: '找不到該頁的標記層，請確認該頁已載入',
              customClass: { popup: 'swal-high-z-index' },
            });
            return;
          }
          const pageCanvas = document.querySelector(
            `#viewerContainer .page[data-page-number="${pageNum}"] canvas`
          );
          if (!pageCanvas) {
            window.Swal.fire({
              icon: 'warning',
              title: '無法辨識',
              text: '找不到該頁 PDF 畫布',
              customClass: { popup: 'swal-high-z-index' },
            });
            return;
          }
          let cropRect;
          const markBlockEl = markLayer.querySelector(`[data-mark-id="${markId}"]`);
          if (markBlockEl) {
            cropRect = markBlockEl.getBoundingClientRect();
          } else {
            const layerRect = markLayer.getBoundingClientRect();
            cropRect = {
              left: layerRect.left + mark.x,
              top: layerRect.top + mark.y,
              width: mark.width,
              height: mark.height,
              right: layerRect.left + mark.x + mark.width,
              bottom: layerRect.top + mark.y + mark.height,
            };
          }
          const croppedCanvas = cropPageCanvasByClientRect(pageCanvas, cropRect);
          if (!croppedCanvas) {
            window.Swal.fire({
              icon: 'warning',
              title: '無法裁切',
              text: '標記區域與畫布無交集',
              customClass: { popup: 'swal-high-z-index' },
            });
            return;
          }
          const base64Image = canvasToBase64Png(croppedCanvas);
          if (isCanvasAllWhite(croppedCanvas)) {
            window.Swal.fire({
              icon: 'warning',
              title: '沒有圖片內容',
              text: '標記區域內沒有可辨識的圖片',
              customClass: { popup: 'swal-high-z-index' },
            });
            return;
          }
          let apiKey = null;
          try {
            apiKey = window.loadApiKeyFromIndexedDB
              ? await window.loadApiKeyFromIndexedDB()
              : '';
          } catch (_) {
            apiKey = '';
          }
          if (!apiKey || !String(apiKey).trim()) {
            window.Swal.fire({
              icon: 'warning',
              title: '請先設定 API KEY',
              text: '請在「報讀系統進階設定」中填寫 Gemini API KEY 以使用 AI 圖片辨識',
              customClass: { popup: 'swal-high-z-index' },
            });
            return;
          }
          aiBtn.disabled = true;
          aiBtn.textContent = '辨識中...';
          try {
            window.Swal.fire({
              title: '正在辨識圖片',
              text: '正在使用 AI 辨識圖片內容…（Gemini）',
              allowOutsideClick: false,
              customClass: { popup: 'swal-high-z-index' },
              didOpen: () => {
                window.Swal.showLoading();
              },
            });
            const geminiFunc = window.geminiImageDescribe;
            if (!geminiFunc) {
              window.Swal.close();
              throw new Error('Gemini 辨識函數未載入（請確認已載入 viewer-reading.js）');
            }
            const aiText = await geminiFunc(base64Image);
            window.Swal.close();
            if (aiText && aiText.trim()) {
              const recognizedText = aiText.trim();
              displayText = recognizedText;
              originalDisplay.textContent = displayText;
              const replacementInput = document.getElementById('replacementTextInput');
              if (replacementInput) replacementInput.value = recognizedText;
              mark.replacementText = recognizedText;
              await saveMarksToPdf(false);
              aiBtn.textContent = '已填入並保存';
              aiBtn.style.backgroundColor = '#28a745';
              setTimeout(() => {
                aiBtn.textContent = 'AI圖片辨識';
                aiBtn.style.backgroundColor = '#6f42c1';
              }, 2000);
            } else {
              window.Swal.fire({
                icon: 'warning',
                title: '辨識結果為空',
                text: 'AI 無法辨識圖片內容',
                customClass: { popup: 'swal-high-z-index' },
              });
            }
          } catch (err) {
            window.Swal.close();
            console.error('AI圖片辨識錯誤:', err);
            window.Swal.fire({
              icon: 'error',
              title: '辨識失敗',
              text: err && err.message ? err.message : 'AI 辨識過程中發生錯誤',
              customClass: { popup: 'swal-high-z-index' },
            });
          } finally {
            aiBtn.disabled = false;
            if (aiBtn.textContent === '辨識中...') aiBtn.textContent = 'AI圖片辨識';
          }
        };

        setTimeout(() => {
          const aiRecognizeBtn = document.getElementById('aiImageRecognizeBtn');
          if (aiRecognizeBtn) {
            aiRecognizeBtn.addEventListener('click', e => {
              e.preventDefault();
              e.stopPropagation();
              if (window.runAiImageRecognizeForEdit) window.runAiImageRecognizeForEdit();
            });
          }
          const copyBtn = document.getElementById('copyOriginalTextBtn');
          if (copyBtn) {
            copyBtn.addEventListener('click', e => {
              e.preventDefault();
              e.stopPropagation();
              window.copyOriginalToReplacement();
            });
          }

          const originalTextDisplayEl = document.getElementById('originalTextDisplay');
          if (originalTextDisplayEl) {
            originalTextDisplayEl.addEventListener('click', e => {
              e.preventDefault();
              e.stopPropagation();
              unlockLocalSpeechForIOS();
              replacementPreviewUi.reset();
              if (displayText && displayText !== '(無文字)') {
                void speakMarkWithHighlight(mark, pageNum, displayText);
              }
            });
          }

          const previewBtn = document.getElementById('previewReplacementTextBtn');
          const PREVIEW_LABEL_IDLE = '替換文字試聽';
          const PREVIEW_LABEL_CANCEL = '取消試聽';
          if (previewBtn) {
            replacementPreviewUi.reset = () => {
              previewBtn.textContent = PREVIEW_LABEL_IDLE;
              previewBtn.style.backgroundColor = '#28a745';
            };
            previewBtn.addEventListener('click', e => {
              e.preventDefault();
              e.stopPropagation();
              if (previewBtn.textContent === PREVIEW_LABEL_CANCEL) {
                if (typeof window.stopCurrentSpeechNow === 'function') {
                  window.stopCurrentSpeechNow();
                }
                replacementPreviewUi.reset();
                return;
              }
              const replacementInput = document.getElementById('replacementTextInput');
              if (replacementInput) {
                const replacementText = replacementInput.value.trim();
                const textToRead = replacementText || displayText;
                if (textToRead && textToRead !== '(無文字)') {
                  previewBtn.textContent = PREVIEW_LABEL_CANCEL;
                  previewBtn.style.backgroundColor = '#dc3545';
                  // 編輯視窗中的「試聽」僅用於預覽，不需要留下朗讀高亮顏色。
                  sendTextToTTSForMark(textToRead, () => {
                    if (previewBtn.textContent === PREVIEW_LABEL_CANCEL) {
                      replacementPreviewUi.reset();
                    }
                  });
                } else {
                  previewBtn.textContent = '無文字';
                  previewBtn.style.backgroundColor = '#dc3545';
                  setTimeout(() => {
                    replacementPreviewUi.reset();
                  }, 1500);
                }
              }
            });
          }
        }, 100);
      },
      preConfirm: () => {
        const input = document.getElementById('replacementTextInput');
        return input ? input.value.trim() : '';
      },
      didDestroy: () => {
        try {
          if (typeof window.stopCurrentSpeechNow === 'function') {
            window.stopCurrentSpeechNow();
          }
          replacementPreviewUi.reset();
        } catch (e) {
          /* ignore */
        }
        try {
          delete window.copyOriginalToReplacement;
          delete window.runAiImageRecognizeForEdit;
        } catch (e) {
          /* ignore */
        }
      },
    });

    if (result.isConfirmed) {
      mark.replacementText = result.value;
      if (displayText && displayText !== '(無文字)') {
        mark.text = displayText;
      }
      await saveMarksToPdf(false);
      window.Swal.fire({
        icon: 'success',
        title: '已儲存',
        text: '替換文字已儲存',
        timer: 1500,
        showConfirmButton: false,
        customClass: { popup: 'swal-high-z-index' },
      });
    } else if (originalTextChanged) {
      // 就算未儲存替換文字，也要把本次編輯視窗確認過的原始文字落檔。
      await saveMarksToPdf(false);
    }

    const markLayerCleanup = document.querySelector(`.mark-layer[data-page-num="${pageNum}"]`);
    if (markLayerCleanup) {
      markLayerCleanup.querySelectorAll('.mark-block').forEach(block => {
        block.classList.remove('dragging');
        delete block.dataset.justHandledDrag;
      });
    }
    if (isMarkingMode) {
      setTimeout(() => renderMarksPage(pageNum), 50);
    }
  }

  async function readMarkBlock(markId, pageNum) {
    const k = mkey(pageNum);
    const mark = marks[k]?.find(m => m.id === markId);
    if (!mark) return;
    let text = (mark.replacementText || '').trim();
    // 朗讀以「已儲存文字」為準（編輯替換文字視窗中的內容），避免即時重新擷取造成跨行誤讀。
    if (!text) text = (mark.text || '').trim();
    if (!text) {
      text = extractTextFromMark(mark, pageNum);
      mark.text = text;
      await saveMarksToPdf(false);
    }
    if (!text) {
      await ensureSwalLocal();
      window.Swal?.fire({ icon: 'info', title: '無文字', text: '此區塊未偵測到可朗讀文字' });
      return;
    }
    const select = document.getElementById('voiceSelect');
    const voiceName = select ? select.value : 'local-zh-female';
    const speedInput = document.getElementById('speakSpeed');
    const ratePercent = Math.max(
      -50,
      Math.min(50, parseFloat(speedInput?.value || '0') || 0)
    );
    const frame = getMarkClientFrameRect(mark, pageNum);
    if (typeof window.speakMarkRegionWithHighlight === 'function' && frame) {
      await window.speakMarkRegionWithHighlight(pageNum, frame, text, { voiceName, ratePercent });
    } else if (window.sendTextToTTS) {
      await window.sendTextToTTS(text, null, { voiceName, ratePercent });
    }
  }

  function bindResizeDrag(markBlockEl, markId, pageNum, markRef) {
    const markLayer = document.querySelector(`.mark-layer[data-page-num="${pageNum}"]`);
    if (!markLayer) return;

    const markState = {
      isDragging: false,
      isResizing: false,
      resizeDirection: null,
      startX: 0,
      startY: 0,
      startLeft: 0,
      startTop: 0,
      startWidth: 0,
      startHeight: 0,
      hasMoved: false,
    };
    let moveHandler;
    let endHandler;
    let escHandler;

    function removeGlobals() {
      if (moveHandler) document.removeEventListener('mousemove', moveHandler);
      if (moveHandler) document.removeEventListener('touchmove', moveHandler);
      if (endHandler) document.removeEventListener('mouseup', endHandler);
      if (endHandler) document.removeEventListener('touchend', endHandler);
      if (escHandler) document.removeEventListener('keydown', escHandler);
      moveHandler = endHandler = escHandler = null;
    }

    function handleMove(e) {
      if (!markState.isDragging && !markState.isResizing) return;
      e.preventDefault();
      const cx = (e.touches ? e.touches[0].clientX : e.clientX);
      const cy = (e.touches ? e.touches[0].clientY : e.clientY);
      const dx = cx - markState.startX;
      const dy = cy - markState.startY;
      if (Math.abs(dx) > 5 || Math.abs(dy) > 5) markState.hasMoved = true;

      if (markState.isResizing && markState.resizeDirection) {
        let nl = markState.startLeft;
        let nt = markState.startTop;
        let nw = markState.startWidth;
        let nh = markState.startHeight;
        const d = markState.resizeDirection;
        if (d.includes('n')) {
          nh = Math.max(10, markState.startHeight - dy);
          nt = markState.startTop + (markState.startHeight - nh);
        }
        if (d.includes('s')) nh = Math.max(10, markState.startHeight + dy);
        if (d.includes('w')) {
          nw = Math.max(10, markState.startWidth - dx);
          nl = markState.startLeft + (markState.startWidth - nw);
        }
        if (d.includes('e')) nw = Math.max(10, markState.startWidth + dx);
        markBlockEl.style.left = `${nl}px`;
        markBlockEl.style.top = `${nt}px`;
        markBlockEl.style.width = `${nw}px`;
        markBlockEl.style.height = `${nh}px`;
      } else if (markState.isDragging) {
        const lr = markLayer.getBoundingClientRect();
        const nl = markState.startLeft + dx;
        const nt = markState.startTop + dy;
        const w = parseFloat(markBlockEl.style.width);
        const h = parseFloat(markBlockEl.style.height);
        const maxL = lr.width - w;
        const maxT = lr.height - h;
        markBlockEl.style.left = `${Math.max(0, Math.min(nl, maxL))}px`;
        markBlockEl.style.top = `${Math.max(0, Math.min(nt, maxT))}px`;
      }
    }

    async function handleEnd() {
      removeGlobals();
      const fl = parseFloat(markBlockEl.style.left);
      const ft = parseFloat(markBlockEl.style.top);
      const fw = parseFloat(markBlockEl.style.width);
      const fh = parseFloat(markBlockEl.style.height);
      const posCh =
        Math.abs(fl - markState.startLeft) > 1 ||
        Math.abs(ft - markState.startTop) > 1 ||
        Math.abs(fw - markState.startWidth) > 1 ||
        Math.abs(fh - markState.startHeight) > 1;

      if ((markState.isDragging || markState.isResizing) && (posCh || markState.hasMoved)) {
        markBlockEl.dataset.suppressNextMarkRead = '1';
        const data = marks[mkey(pageNum)]?.find(m => m.id === markId);
        if (data) {
          data.x = fl;
          data.y = ft;
          data.width = fw;
          data.height = fh;
          const p = viewportToPdfCoords(fl, ft, fw, fh, pageNum);
          if (p) {
            data.pdfL = p.pdfL;
            data.pdfT = p.pdfT;
            data.pdfR = p.pdfR;
            data.pdfB = p.pdfB;
          }
          if (markState.isResizing) data.text = extractTextFromMark(data, pageNum);
          await saveMarksToPdf(false);
        }
      }

      markState.isDragging = markState.isResizing = false;
      markState.resizeDirection = null;
      markState.hasMoved = false;
      markBlockEl.classList.remove('dragging');
    }

    function bindGlobals() {
      if (moveHandler) return;
      moveHandler = e => handleMove(e);
      endHandler = () => {
        handleEnd();
      };
      escHandler = e => {
        if (e.key === 'Escape') handleEnd();
      };
      document.addEventListener('mousemove', moveHandler);
      document.addEventListener('touchmove', moveHandler, { passive: false });
      document.addEventListener('mouseup', endHandler);
      document.addEventListener('touchend', endHandler);
      document.addEventListener('keydown', escHandler);
    }

    function startResize(e, direction) {
      if (markRef.locked) return;
      e.preventDefault();
      e.stopPropagation();
      markState.isResizing = true;
      markState.isDragging = false;
      markState.resizeDirection = direction;
      const br = markBlockEl.getBoundingClientRect();
      const lr = markLayer.getBoundingClientRect();
      const cx = e.touches ? e.touches[0].clientX : e.clientX;
      const cy = e.touches ? e.touches[0].clientY : e.clientY;
      markState.startX = cx;
      markState.startY = cy;
      markState.startLeft = br.left - lr.left;
      markState.startTop = br.top - lr.top;
      markState.startWidth = br.width;
      markState.startHeight = br.height;
      markBlockEl.classList.add('dragging');
      bindGlobals();
    }

    function startDrag(e) {
      if (markRef.locked) return;
      if (
        e.target.classList.contains('resize-handle') ||
        e.target.closest('.mark-toolbar') ||
        e.target.classList.contains('mark-delete') ||
        e.target.classList.contains('mark-edit') ||
        e.target.classList.contains('mark-lock')
      ) {
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      markState.isDragging = true;
      markState.isResizing = false;
      markState.hasMoved = false;
      const br = markBlockEl.getBoundingClientRect();
      const lr = markLayer.getBoundingClientRect();
      const cx = e.touches ? e.touches[0].clientX : e.clientX;
      const cy = e.touches ? e.touches[0].clientY : e.clientY;
      markState.startX = cx;
      markState.startY = cy;
      markState.startLeft = br.left - lr.left;
      markState.startTop = br.top - lr.top;
      markBlockEl.classList.add('dragging');
      bindGlobals();
    }

    ['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se'].forEach(dir => {
      const h = markBlockEl.querySelector(`.resize-handle.${dir}`);
      if (!h) return;
      h.addEventListener('mousedown', ev => startResize(ev, dir));
      h.addEventListener('touchstart', ev => startResize(ev, dir), { passive: false });
    });

    markBlockEl.addEventListener('mousedown', startDrag);
    markBlockEl.addEventListener('touchstart', startDrag, { passive: false });
  }

  function addToolbarAndInteractions(markBlockEl, markId, pageNum) {
    const k = mkey(pageNum);
    let markData = marks[k]?.find(m => m.id === markId);
    if (!markData) return;

    const toolbar = document.createElement('div');
    toolbar.className = 'mark-toolbar';
    if ((parseFloat(markBlockEl.style.width) || 0) < 80) toolbar.classList.add('small-box');

    const lockBtn = document.createElement('div');
    lockBtn.className = 'mark-lock';
    lockBtn.textContent = markData.locked ? '🔒' : '🔓';
    if (markData.locked) {
      lockBtn.classList.add('locked');
      markBlockEl.classList.add('locked');
    }
    lockBtn.title = markData.locked
      ? '解鎖標記（解鎖後可移動、調整和刪除）'
      : '鎖定標記（鎖定後不可移動、調整和刪除）';
    lockBtn.addEventListener('click', async e => {
      e.preventDefault();
      e.stopPropagation();
      markData = marks[k]?.find(m => m.id === markId);
      if (!markData) return;
      markData.locked = !markData.locked;
      if (markData.locked) {
        markBlockEl.classList.add('locked');
        lockBtn.textContent = '🔒';
        lockBtn.classList.add('locked');
        lockBtn.title = '解鎖標記（解鎖後可移動、調整和刪除）';
      } else {
        markBlockEl.classList.remove('locked');
        lockBtn.textContent = '🔓';
        lockBtn.classList.remove('locked');
        lockBtn.title = '鎖定標記（鎖定後不可移動、調整和刪除）';
      }
      await saveMarksToPdf(false);
    });
    toolbar.appendChild(lockBtn);

    const editBtn = document.createElement('div');
    editBtn.className = 'mark-edit';
    editBtn.textContent = '✎';
    editBtn.title = '編輯替換文字';
    editBtn.addEventListener('click', e => {
      e.preventDefault();
      e.stopPropagation();
      editMarkReplacementText(markId, pageNum);
    });
    toolbar.appendChild(editBtn);

    const delBtn = document.createElement('div');
    delBtn.className = 'mark-delete';
    delBtn.textContent = '×';
    delBtn.title = '刪除此標記';
    delBtn.addEventListener('click', async e => {
      e.preventDefault();
      e.stopPropagation();
      if (markData.locked) {
        ensureSwalLocal().then(() =>
          window.Swal?.fire({
            icon: 'warning',
            title: '無法刪除',
            text: '請先解鎖',
            timer: 1800,
            showConfirmButton: false,
          })
        );
        return;
      }
      if (!isCloudFileMode) {
        await ensureSwalLocal();
        const confirmResult = await window.Swal?.fire({
          icon: 'warning',
          title: '確認刪除標記',
          text: '此操作會刪除此區塊標記，確定要繼續嗎？',
          showCancelButton: true,
          confirmButtonText: '刪除',
          cancelButtonText: '取消',
          customClass: { popup: 'swal-high-z-index' },
        });
        if (!confirmResult?.isConfirmed) return;
      }
      deleteMarkBlock(markId, pageNum);
    });
    toolbar.appendChild(delBtn);
    markBlockEl.appendChild(toolbar);

    if (isMarkingMode) {
      ['nw', 'n', 'ne', 'w', 'e', 'sw', 's', 'se'].forEach(dir => {
        const h = document.createElement('div');
        h.className = `resize-handle ${dir}`;
        markBlockEl.appendChild(h);
      });
      bindResizeDrag(markBlockEl, markId, pageNum, markData);
    }

    markBlockEl.addEventListener('click', e => {
      if (e.target.closest('.mark-toolbar') || e.target.classList.contains('resize-handle')) {
        return;
      }
      if (markBlockEl.dataset.suppressNextMarkRead) {
        delete markBlockEl.dataset.suppressNextMarkRead;
        return;
      }
      if (markBlockEl.classList.contains('dragging')) return;
      e.stopPropagation();
      readMarkBlock(markId, pageNum);
    });
  }

  function layoutMarkBlockFromData(markBlock, mark, pageNum) {
    let left = mark.x;
    let top = mark.y;
    let width = mark.width;
    let height = mark.height;
    if (
      typeof mark.pdfL === 'number' &&
      typeof mark.pdfT === 'number' &&
      typeof mark.pdfR === 'number' &&
      typeof mark.pdfB === 'number'
    ) {
      const c = pdfToViewportCoords(mark.pdfL, mark.pdfT, mark.pdfR, mark.pdfB, pageNum);
      if (c) {
        left = c.left;
        top = c.top;
        width = c.width;
        height = c.height;
        mark.x = left;
        mark.y = top;
        mark.width = width;
        mark.height = height;
      }
    }
    markBlock.style.left = `${left}px`;
    markBlock.style.top = `${top}px`;
    markBlock.style.width = `${width}px`;
    markBlock.style.height = `${height}px`;
  }

  function renderMarksPage(pageNum) {
    const layer = ensureMarkLayer(pageNum);
    if (!layer) return;
    if (layer.querySelector('.mark-block.dragging')) return;

    layer.querySelectorAll('.mark-block').forEach(el => el.remove());
    const k = mkey(pageNum);
    const list = marks[k];
    if (!list || !list.length) return;

    const markArea = m => (m.width || 0) * (m.height || 0);
    const sorted = [...list].sort((a, b) => markArea(b) - markArea(a));

    sorted.forEach(mark => {
      const el = document.createElement('div');
      el.className = 'mark-block';
      el.dataset.markId = String(mark.id);
      layoutMarkBlockFromData(el, mark, pageNum);

      const label = document.createElement('div');
      label.className = 'mark-label';
      label.textContent = `區塊 ${mark.id + 1}`;
      el.appendChild(label);

      addToolbarAndInteractions(el, mark.id, pageNum);
      applyMarkVisualMode(el);
      layer.appendChild(el);
    });
  }

  function renderMarksAllPages() {
    const app = getApp();
    const n = app?.pagesCount || 0;
    for (let p = 1; p <= n; p++) renderMarksPage(p);
  }

  /**
   * 以記憶體中的位元組關閉再開啟一次 PDF（不重整整頁 HTML），效果接近使用者按 F5 重載 viewer。
   * 用於雲端／IDB 首開時頁面節點與圖層尚未對齊、標記層出不來的情況。
   */
  async function softReloadViewerWithCurrentPdfBytes() {
    const app = getApp();
    if (!app?.pdfDocument) return false;
    try {
      const raw = await app.pdfDocument.getData();
      const copy = new Uint8Array(raw.length);
      copy.set(raw);
      const originalUrl = app._docFilename || app.url || 'document.pdf';
      const page = app.pdfViewer?.currentPageNumber || 1;
      await app.close();
      await app.open({ data: copy, originalUrl });
      await new Promise(resolve => {
        const eb = app.eventBus;
        if (!eb?._on) {
          resolve();
          return;
        }
        let done = false;
        const finish = () => {
          if (done) return;
          done = true;
          try {
            eb._off('pagesloaded', onPl);
          } catch (_) {
            /* ignore */
          }
          clearTimeout(tid);
          resolve();
        };
        const onPl = () => finish();
        const tid = setTimeout(finish, 8000);
        eb._on('pagesloaded', onPl);
      });
      try {
        if (app.pdfViewer && page > 0) app.pdfViewer.currentPageNumber = page;
      } catch (_) {
        /* ignore */
      }
      return true;
    } catch (e) {
      console.warn('軟重新載入 PDF 失敗', e);
      return false;
    }
  }

  /**
   * documentloaded 時常只有第一頁就緒；延遲補一次同步／繪製（軟重新載入後仍需要）。
   */
  function scheduleMarksDomRetries() {
    const tick = () => {
      if (!getApp()?.pdfDocument) return;
      syncAllMarkLayersSize();
      renderMarksAllPages();
      refreshAllMarkVisuals();
    };
    tick();
    requestAnimationFrame(() => tick());
    setTimeout(tick, 120);
  }

  function applyMarkVisualMode(block) {
    const toolbar = block.querySelector('.mark-toolbar');
    const label = block.querySelector('.mark-label');
    const handles = block.querySelectorAll('.resize-handle');
    if (isMarkingMode && isManualMarkingEnabled) {
      block.style.opacity = '1';
      block.style.pointerEvents = 'auto';
      block.style.removeProperty('border');
      block.style.removeProperty('background-color');
      block.setAttribute('tabindex', '0');
      if (label) {
        label.style.removeProperty('display');
        label.style.removeProperty('visibility');
        label.style.removeProperty('opacity');
      }
      if (toolbar) {
        toolbar.style.display = 'flex';
        toolbar.style.removeProperty('visibility');
        toolbar.style.removeProperty('opacity');
      }
      handles.forEach(h => {
        h.style.removeProperty('display');
      });
    } else {
      block.removeAttribute('tabindex');
      block.style.opacity = '0';
      block.style.pointerEvents = 'auto';
      block.style.border = 'none';
      block.style.backgroundColor = 'transparent';
      if (label) label.style.display = 'none';
      if (toolbar) toolbar.style.display = 'none';
      handles.forEach(h => {
        h.style.display = 'none';
      });
    }
  }

  function refreshAllMarkVisuals() {
    document.querySelectorAll('.mark-block').forEach(applyMarkVisualMode);
  }

  function ensureMarkBlocksRenderedForHit() {
    if (!hasMarksData()) return;
    const blocks = document.querySelectorAll('#viewerContainer .mark-layer .mark-block');
    if (blocks.length === 0) {
      syncAllMarkLayersSize();
      renderMarksAllPages();
      refreshAllMarkVisuals();
      return;
    }
    for (const b of blocks) {
      const r = b.getBoundingClientRect();
      if (r.width < 1.5 || r.height < 1.5) {
        syncAllMarkLayersSize();
        renderMarksAllPages();
        refreshAllMarkVisuals();
        return;
      }
    }
  }

  /**
   * 一般閱讀（未開「手動標記」）時標記框 opacity=0，且 .mark-layer 可能為 pointer-events:none，
   * 首次開雲端檔 textLayer 常壓在子元素之上，點得到文字卻點不到 .mark-block。
   * 用座標找最小面積的命中框；在 document capture 攔截，早於 #viewerContainer 內各層。
   */
  function findTopMarkBlockUnderPoint(cx, cy) {
    const blocks = document.querySelectorAll('#viewerContainer .mark-layer .mark-block');
    let best = null;
    let bestArea = Infinity;
    for (const block of blocks) {
      try {
        if (getComputedStyle(block).pointerEvents === 'none') continue;
        const r = block.getBoundingClientRect();
        if (r.width < 2 || r.height < 2) continue;
        if (cx < r.left || cx > r.right || cy < r.top || cy > r.bottom) continue;
        const area = r.width * r.height;
        if (area < bestArea) {
          bestArea = area;
          best = block;
        }
      } catch (_) {
        /* ignore */
      }
    }
    return best;
  }

  let markHitRecentKey = '';
  let markHitRecentTs = 0;

  function tryHandleMarkHitAtClientPoint(e) {
    if (isMarkingMode || isManualMarkingEnabled) return;
    if (e.button !== 0 && e.button !== undefined) return;
    // 有任何前景彈窗時，不處理底層標記點擊，避免按彈窗按鈕時誤觸發朗讀。
    const apiKeyModal = document.getElementById('apiKeyModal');
    const apiKeyModalOpen = !!(apiKeyModal && getComputedStyle(apiKeyModal).display !== 'none');
    const swalContainerOpen = !!document.querySelector('.swal2-container:not([style*="display: none"])');
    if (apiKeyModalOpen || swalContainerOpen) return;
    const cx = e.clientX;
    const cy = e.clientY;
    const vc = document.getElementById('viewerContainer');
    if (!vc) return;
    const vr = vc.getBoundingClientRect();
    if (cx < vr.left || cx > vr.right || cy < vr.top || cy > vr.bottom) return;

    const stack = document.elementsFromPoint(cx, cy);
    const topEl = stack[0];
    if (topEl?.closest?.('.mark-toolbar, .resize-handle')) return;
    if (topEl?.closest?.('#readingFrame, #pendingReadButton')) return;

    ensureMarkBlocksRenderedForHit();
    const block = findTopMarkBlockUnderPoint(cx, cy);
    if (!block) return;
    if (getComputedStyle(block).pointerEvents === 'none') return;
    const layer = block.closest('.mark-layer');
    const pageNum = parseInt(layer?.dataset?.pageNum, 10);
    const markId = parseInt(block.dataset?.markId, 10);
    if (!Number.isFinite(pageNum) || !Number.isFinite(markId)) return;

    const dedupKey = `${markId}|${Math.round(cx / 8)}|${Math.round(cy / 8)}`;
    const now = Date.now();
    if (dedupKey === markHitRecentKey && now - markHitRecentTs < 550) {
      e.preventDefault();
      e.stopPropagation();
      return;
    }
    markHitRecentKey = dedupKey;
    markHitRecentTs = now;

    e.preventDefault();
    e.stopPropagation();
    void readMarkBlock(markId, pageNum);
  }

  function onDocumentMarkHitCapture(e) {
    if (e.type !== 'click') return;
    tryHandleMarkHitAtClientPoint(e);
  }

  function onDocumentMarkPointerUpCapture(e) {
    if (e.type !== 'pointerup') return;
    if (!e.isPrimary) return;
    if (e.pointerType === 'mouse' && e.button !== 0) return;
    tryHandleMarkHitAtClientPoint(e);
  }

  function onPointerDownStartMark(e) {
    if (!isMarkingMode || !isManualMarkingEnabled) return;
    if (e.target.closest('.mark-block') || e.target.closest('.mark-toolbar')) return;
    const pageDiv = e.target.closest?.('.page[data-page-number]');
    if (!pageDiv) return;
    if (!e.target.closest('#viewerContainer')) return;

    const pageNum = parseInt(pageDiv.getAttribute('data-page-number'), 10);
    const markLayer = ensureMarkLayer(pageNum);
    if (!markLayer) return;

    e.preventDefault();
    e.stopPropagation();

    const lr = markLayer.getBoundingClientRect();
    const cx = e.touches ? e.touches[0].clientX : e.clientX;
    const cy = e.touches ? e.touches[0].clientY : e.clientY;
    const sx = cx - lr.left;
    const sy = cy - lr.top;

    const block = document.createElement('div');
    block.className = 'mark-block';
    block.style.left = `${sx}px`;
    block.style.top = `${sy}px`;
    block.style.width = '0px';
    block.style.height = '0px';
    const label = document.createElement('div');
    label.className = 'mark-label';
    block.appendChild(label);
    markLayer.appendChild(block);

    currentDraw = { pageNum, startX: sx, startY: sy, block };

    function move(ev) {
      if (!currentDraw) return;
      const r = markLayer.getBoundingClientRect();
      const x = (ev.touches ? ev.touches[0].clientX : ev.clientX) - r.left;
      const y = (ev.touches ? ev.touches[0].clientY : ev.clientY) - r.top;
      const w = Math.abs(x - currentDraw.startX);
      const h = Math.abs(y - currentDraw.startY);
      const lx = Math.min(currentDraw.startX, x);
      const ly = Math.min(currentDraw.startY, y);
      block.style.left = `${lx}px`;
      block.style.top = `${ly}px`;
      block.style.width = `${w}px`;
      block.style.height = `${h}px`;
    }

    async function end() {
      document.removeEventListener('mousemove', move, true);
      document.removeEventListener('touchmove', move, true);
      document.removeEventListener('mouseup', end, true);
      document.removeEventListener('touchend', end, true);
      if (!currentDraw) return;
      const w = parseFloat(block.style.width);
      const h = parseFloat(block.style.height);
      const lx = parseFloat(block.style.left);
      const ly = parseFloat(block.style.top);
      currentDraw = null;
      if (w <= 10 || h <= 10) {
        block.remove();
        return;
      }
      const id = markIdCounter++;
      label.textContent = `區塊 ${id + 1}`;
      block.dataset.markId = String(id);
      const pk = mkey(pageNum);
      if (!marks[pk]) marks[pk] = [];
      const pcoords = viewportToPdfCoords(lx, ly, w, h, pageNum);
      const newMark = {
        id,
        x: lx,
        y: ly,
        width: w,
        height: h,
        text: '',
        replacementText: '',
        locked: false,
        ...(pcoords || {}),
      };
      newMark.text = extractTextFromMark(newMark, pageNum) || '';
      marks[pk].push(newMark);
      block.remove();
      renderMarksPage(pageNum);
      await saveMarksToPdf(false);
    }

    document.addEventListener('mousemove', move, true);
    document.addEventListener('touchmove', move, { capture: true, passive: false });
    document.addEventListener('mouseup', end, true);
    document.addEventListener('touchend', end, true);
  }

  function bindManualMarkingPointerHandlers() {
    const vc = document.getElementById('viewerContainer');
    if (!vc || markDrawBound) return;
    vc.addEventListener('mousedown', onPointerDownStartMark, true);
    vc.addEventListener('touchstart', onPointerDownStartMark, { capture: true, passive: false });
    markDrawBound = true;
  }

  function unbindManualMarkingPointerHandlers() {
    const vc = document.getElementById('viewerContainer');
    if (!vc || !markDrawBound) return;
    vc.removeEventListener('mousedown', onPointerDownStartMark, true);
    vc.removeEventListener('touchstart', onPointerDownStartMark, true);
    markDrawBound = false;
  }

  function setTextLayersPointerEvents(val) {
    document.querySelectorAll('#viewerContainer .textLayer').forEach(tl => {
      tl.style.pointerEvents = val;
    });
  }

  function goToPageAndWaitForRender(pageNum, timeoutMs = 2000) {
    const app = getApp();
    if (!app?.eventBus || !app.pdfViewer) return Promise.resolve();
    return new Promise(resolve => {
      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        clearTimeout(tid);
        app.eventBus._off('pagerendered', onRender);
        resolve();
      };
      const tid = setTimeout(finish, timeoutMs);
      const onRender = e => {
        if (e.pageNumber === pageNum) finish();
      };
      app.eventBus._on('pagerendered', onRender);
      app.pdfViewer.currentPageNumber = pageNum;
      try {
        app.pdfViewer.scrollPageIntoView({ pageNumber: pageNum });
      } catch (_) {
        /* ignore */
      }
    });
  }

  function setMarkingToolbarExtrasVisible(visible) {
    const ids = [
      'aiMarkButtonToolbar',
      'setPasswordButtonToolbar',
      'saveMarkedPdfButtonToolbar',
      'downloadMarkedPdfButtonToolbar',
      'clearMarksButtonToolbar',
    ];
    ids.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      if (visible) {
        el.style.display = 'inline-block';
        el.style.visibility = 'visible';
      } else {
        el.style.display = 'none';
      }
    });
  }

  async function setPasswordForPdf() {
    await ensureSwalLocal();
    const Swal = window.Swal;
    const app = getApp();
    if (!app?.pdfDocument || !originalPdfBytes) {
      Swal.fire({
        icon: 'warning',
        title: '無法設定',
        text: '請先載入 PDF',
        customClass: { popup: 'swal-high-z-index' },
      });
      return;
    }
    if (typeof PDFLib === 'undefined') {
      Swal.fire({
        icon: 'error',
        title: '庫未載入',
        text: 'PDF-lib 尚未載入，請重新整理頁面',
        customClass: { popup: 'swal-high-z-index' },
      });
      return;
    }
    const passwordResult = await Swal.fire({
      title: '設定密碼保護',
      html: `
            <div style="text-align: left; margin-bottom: 15px;">
                <p style="font-size: 13px; color: #666; margin-bottom: 15px;">設定密碼後，日後編輯、刪除或增加標記時需要輸入密碼</p>
            </div>
            <div style="text-align: left; width: 100%; box-sizing: border-box;">
                <label for="passwordInput" style="display: block; margin-bottom: 8px; font-weight: bold; font-size: 18px; color: #333;">密碼（留空則移除密碼保護）：</label>
                <input type="password" id="passwordInput" class="swal2-input" placeholder="輸入密碼" style="width: 95%; font-size: 14px; box-sizing: border-box;">
            </div>
            <div style="text-align: left; margin-top: 15px; width: 100%; box-sizing: border-box;">
                <label for="confirmPasswordInput" style="display: block; margin-bottom: 8px; font-weight: bold; font-size: 18px; color: #333;">確認密碼：</label>
                <input type="password" id="confirmPasswordInput" class="swal2-input" placeholder="再次輸入密碼" style="width: 95%; font-size: 14px; box-sizing: border-box;">
            </div>
        `,
      width: '70%',
      maxWidth: '450px',
      showCancelButton: true,
      confirmButtonText: '確定',
      cancelButtonText: '取消',
      customClass: { popup: 'swal-high-z-index' },
      didOpen: () => document.getElementById('passwordInput')?.focus(),
      preConfirm: () => {
        const passwordInput = document.getElementById('passwordInput');
        const confirmPasswordInput = document.getElementById('confirmPasswordInput');
        const password = passwordInput ? passwordInput.value : '';
        const confirmPassword = confirmPasswordInput ? confirmPasswordInput.value : '';
        if (password && password !== confirmPassword) {
          Swal.showValidationMessage('兩次輸入的密碼不一致');
          return false;
        }
        return password;
      },
    });
    if (!passwordResult.isConfirmed) return;
    try {
      Swal.fire({
        title: '正在設定',
        text: '正在設定密碼保護...',
        allowOutsideClick: false,
        didOpen: () => Swal.showLoading(),
        customClass: { popup: 'swal-high-z-index' },
      });
      const { PDFDocument } = PDFLib;
      const pdfDocLib = await PDFDocument.load(originalPdfBytes);
      let passwordHash = null;
      if (passwordResult.value && String(passwordResult.value).trim()) {
        const plain = String(passwordResult.value).trim();
        passwordHash = await hashPassword(plain);
        currentSessionPassword = plain;
      } else {
        currentSessionPassword = null;
      }
      let keywordsArray = normalizeKeywordsArray(pdfDocLib.getKeywords());
      if (passwordHash) {
        keywordsArray = keywordsArray.filter(k => !String(k).includes('PDF_MARKS_PASSWORD:'));
        keywordsArray.push(`PDF_MARKS_PASSWORD:${passwordHash}`);
        pdfDocLib.setKeywords(keywordsArray);
        pdfPasswordHash = passwordHash;
      } else {
        keywordsArray = keywordsArray.filter(k => !String(k).includes('PDF_MARKS_PASSWORD:'));
        pdfDocLib.setKeywords(keywordsArray);
        pdfPasswordHash = null;
        currentSessionPassword = null;
      }
      const pdfBytes = await pdfDocLib.save();
      const u8 = pdfBytes instanceof Uint8Array ? pdfBytes : new Uint8Array(pdfBytes);
      originalPdfBytes = u8.slice(0);
      await savePdfBytesToViewerIdb(originalPdfBytes);
      Swal.fire({
        icon: 'success',
        title: '設定成功',
        text: passwordHash ? '密碼保護已設定' : '密碼保護已移除',
        timer: 1500,
        showConfirmButton: false,
        customClass: { popup: 'swal-high-z-index' },
      });
    } catch (error) {
      console.error('設定密碼失敗:', error);
      Swal.fire({
        icon: 'error',
        title: '設定失敗',
        text: '無法設定密碼保護: ' + (error.message || String(error)),
        customClass: { popup: 'swal-high-z-index' },
      });
    }
  }

  async function downloadMarkedPdf() {
    await ensureSwalLocal();
    const Swal = window.Swal;
    if (!originalPdfBytes) {
      Swal.fire({
        icon: 'warning',
        title: '無法下載',
        text: '請先保存標記到 PDF',
        customClass: { popup: 'swal-high-z-index' },
      });
      return;
    }
    if (isCloudFileMode) {
      Swal.fire({
        icon: 'warning',
        title: '雲端檔案無法匯出保存',
        text:
          '雲端檔案無法回寫保存標記，因此也無法匯出帶標記的 PDF。請使用「取消標記」回到原狀，或改用本地 PDF 檔案。',
        customClass: { popup: 'swal-high-z-index' },
      });
      return;
    }
    try {
      await saveMarksToPdf(true);
      const blob = new Blob([originalPdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const app = getApp();
      let fileName = (app && app._docFilename) || 'document.pdf';
      if (!fileName.includes('已標示')) {
        const lastDotIndex = fileName.lastIndexOf('.');
        if (lastDotIndex > 0) {
          fileName =
            fileName.substring(0, lastDotIndex) + '_已標示' + fileName.substring(lastDotIndex);
        } else {
          fileName = `${fileName}_已標示`;
        }
      }
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      Swal.fire({
        icon: 'success',
        title: '匯出成功',
        text: 'PDF 已匯出到本地（標記已寫入檔案中繼資料，重新載入時會自動恢復）',
        timer: 2000,
        showConfirmButton: false,
        customClass: { popup: 'swal-high-z-index' },
      });
    } catch (error) {
      console.error('匯出 PDF 失敗:', error);
      Swal.fire({
        icon: 'error',
        title: '匯出失敗',
        text: '無法匯出 PDF: ' + (error.message || String(error)),
        customClass: { popup: 'swal-high-z-index' },
      });
    }
  }

  async function clearAllMarks() {
    await ensureSwalLocal();
    const Swal = window.Swal;
    const r = await Swal.fire({
      icon: 'warning',
      title: '確認清除',
      text: isCloudFileMode
        ? '雲端檔案模式：只能暫時清除畫面上的標記，無法真正清除/保存回 PDF。確定要繼續嗎？'
        : '確定要清除所有標記嗎？',
      showCancelButton: true,
      confirmButtonText: '確定',
      cancelButtonText: '取消',
      customClass: { popup: 'swal-high-z-index' },
    });
    if (!r.isConfirmed) return;
    Object.keys(marks).forEach(k => delete marks[k]);
    markIdCounter = 0;
    renderMarksAllPages();
    refreshAllMarkVisuals();
    if (!isCloudFileMode && typeof PDFLib !== 'undefined' && originalPdfBytes) {
      await saveMarksToPdf(true);
    } else if (isCloudFileMode) {
      await Swal.fire({
        icon: 'warning',
        title: '雲端檔案無法真正清除',
        text: '此操作僅影響本次編輯畫面；請按「取消標記」回復原狀。',
        timer: 2200,
        showConfirmButton: false,
        customClass: { popup: 'swal-high-z-index' },
      });
    }
    await Swal.fire({
      icon: 'success',
      title: '已清除',
      text: isCloudFileMode ? '畫面上的標記已暫時清除' : '所有標記已清除',
      timer: 1500,
      showConfirmButton: false,
      customClass: { popup: 'swal-high-z-index' },
    });
  }

  function syncMarkIdCounterFromMarks() {
    const flat = Object.values(marks).flat();
    markIdCounter = flat.length ? Math.max(...flat.map(x => x.id)) + 1 : 0;
  }

  async function aiMarkPdf() {
    await ensureSwalLocal();
    const Swal = window.Swal;
    const app = getApp();
    if (!app?.pdfDocument || !originalPdfBytes) {
      Swal.fire({
        icon: 'warning',
        title: '無法標記',
        text: '請先載入 PDF',
        customClass: { popup: 'swal-high-z-index' },
      });
      return;
    }
    const apiKey =
      typeof window.loadApiKeyFromIndexedDB === 'function'
        ? await window.loadApiKeyFromIndexedDB()
        : null;
    if (!apiKey) {
      Swal.fire({
        icon: 'warning',
        title: '需要 API KEY',
        text: '請先設定 Gemini API KEY 才能使用 AI 標記',
        customClass: { popup: 'swal-high-z-index' },
      });
      return;
    }
    const hasExistingMarks =
      Object.keys(marks).length > 0 &&
      Object.values(marks).some(
        pageMarks => Array.isArray(pageMarks) && pageMarks.length > 0
      );
    let isSupplementMode = false;
    if (hasExistingMarks) {
      const confirmResult = await Swal.fire({
        icon: 'info',
        title: '選擇標記模式',
        html: `
                <p>PDF 中已有標記，請選擇：</p>
                <p style="margin-top: 15px;">
                    <strong>補充模式</strong>：AI 將模仿現有標記風格，只補充未標記的區域，保留所有現有標記<br>
                    <strong>重新標記</strong>：刪除所有現有標記，重新進行 AI 標記
                </p>
                <p style="margin-top: 12px; color:#8a5a00; line-height:1.6;">
                    <strong>提醒：</strong>AI 標記無法保證 100% 完整。遇到圖片、掃描件或特殊排版格式時，可能有部分內容無法被 AI 正確框選，請再以手動標記補齊。
                </p>
            `,
        showCancelButton: true,
        confirmButtonText: '補充模式',
        cancelButtonText: '重新標記',
        confirmButtonColor: '#28a745',
        cancelButtonColor: '#dc3545',
        customClass: { popup: 'swal-high-z-index' },
      });
      if (confirmResult.isConfirmed) {
        isSupplementMode = true;
      } else if (confirmResult.dismiss === Swal.DismissReason.cancel) {
        /* 重新標記 */
      } else {
        return;
      }
    }
    let isCancelled = false;
    const swalPromise = Swal.fire({
      title: 'AI 標記中',
      text: '正在使用 AI 分析 PDF 內容並生成標記…',
      allowOutsideClick: false,
      showCancelButton: true,
      cancelButtonText: '取消',
      showConfirmButton: false,
      didOpen: () => {
        Swal.showLoading();
        document.querySelector('.swal2-cancel')?.addEventListener('click', () => {
          isCancelled = true;
        });
      },
      customClass: { popup: 'swal-high-z-index' },
    });
    swalPromise.then(result => {
      if (result.dismiss === Swal.DismissReason.cancel) isCancelled = true;
    });
    try {
      if (!isSupplementMode) {
        Object.keys(marks).forEach(k => delete marks[k]);
        markIdCounter = 0;
      }
      const existingMarksInfo = {};
      if (isSupplementMode) {
        for (const pageNumStr in marks) {
          const pm = marks[pageNumStr];
          if (pm && Array.isArray(pm) && pm.length > 0) {
            existingMarksInfo[pageNumStr] = pm.map(mark => ({
              text: mark.text || '',
              description: mark.replacementText ? `替換為：${mark.replacementText}` : '原始文字',
            }));
          }
        }
      }
      const pdfDocument = app.pdfDocument;
      const totalPages = pdfDocument.numPages;
      let totalCreatedMarksCount = 0;
      const originalCurrentPage = app.pdfViewer.currentPageNumber;
      for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
        if (isCancelled) {
          Swal.close();
          app.pdfViewer.currentPageNumber = originalCurrentPage;
          await goToPageAndWaitForRender(originalCurrentPage, 800);
          await Swal.fire({
            icon: 'info',
            title: '已取消',
            text: `已取消 AI 標記，已處理 ${pageNum - 1} / ${totalPages} 頁`,
            timer: 2000,
            showConfirmButton: false,
            customClass: { popup: 'swal-high-z-index' },
          });
          return;
        }
        Swal.update({
          title: 'AI 標記中',
          text: `正在處理第 ${pageNum} / ${totalPages} 頁…`,
          allowOutsideClick: false,
          showCancelButton: true,
          cancelButtonText: '取消',
          showConfirmButton: false,
        });
        await goToPageAndWaitForRender(pageNum);
        const page = await pdfDocument.getPage(pageNum);
        const textContent = await page.getTextContent();
        let fullText = '';
        textContent.items.forEach(item => {
          if (item.str) fullText += item.str + ' ';
        });
        if (isCancelled) {
          Swal.close();
          app.pdfViewer.currentPageNumber = originalCurrentPage;
          await goToPageAndWaitForRender(originalCurrentPage, 800);
          await Swal.fire({
            icon: 'info',
            title: '已取消',
            text: `已取消 AI 標記，已處理 ${pageNum - 1} / ${totalPages} 頁`,
            timer: 2000,
            showConfirmButton: false,
            customClass: { popup: 'swal-high-z-index' },
          });
          return;
        }
        if (!fullText.trim()) {
          console.log(`第 ${pageNum} 頁沒有文字內容，跳過`);
          continue;
        }
        const existingHint =
          isSupplementMode && existingMarksInfo[mkey(pageNum)]?.length
            ? `\n\n已有標記範例（請勿重複框選這些區域）：\n${JSON.stringify(
                existingMarksInfo[mkey(pageNum)],
                null,
                2
              )}`
            : '';
        const prompt = `請分析以下PDF頁面的文字內容，根據語意和邏輯結構，識別出應該被標記的重要區塊。

**標記規則：**

1. **考卷類內容識別**：
   - 如果內容包含題號（如「1.」「2.」「第1題」等）和選擇題選項（如「(A)」「(B)」「(C)」「(D)」或「A.」「B.」「C.」「D.」或「①」「②」「③」「④」等），則判定為考卷類內容
   
2. **考卷類內容的標記方式**：
   - 將每一題的**題幹和所有選項**標記為**同一個區塊**
   - 例如：第1題包含題幹「下列哪個選項正確？」和選項「(A)選項1 (B)選項2 (C)選項3 (D)選項4」，則將整個內容（從題號到最後一個選項）標記為一個區塊
   - 每題獨立標記，不要將多題合併成一個區塊
   - 確保題幹和所有選項都在同一個區塊內

3. **一般內容的標記方式**（非考卷類，如文章、說明、段落等）：
   - 以正常段落為單位標記
   - 每個完整的段落、標題、重要說明等標記為一個區塊
   - 保持語意完整性

請以JSON格式返回標記建議，格式如下：
{
  "marks": [
    {
      "text": "要標記的文字內容（考卷：完整包含題幹和所有選項；一般內容：完整段落）",
      "description": "這個標記的說明"
    }
  ]
}

文字內容：
${fullText.trim()}${existingHint}

請只返回JSON格式，不要包含其他說明文字。`;
        if (isCancelled) {
          Swal.close();
          app.pdfViewer.currentPageNumber = originalCurrentPage;
          await goToPageAndWaitForRender(originalCurrentPage, 800);
          await Swal.fire({
            icon: 'info',
            title: '已取消',
            text: `已取消 AI 標記，已處理 ${pageNum - 1} / ${totalPages} 頁`,
            timer: 2000,
            showConfirmButton: false,
            customClass: { popup: 'swal-high-z-index' },
          });
          return;
        }
        try {
          const response = await fetch(
            'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' +
              apiKey,
            {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                contents: [{ role: 'user', parts: [{ text: prompt }] }],
              }),
            }
          );
          if (!response.ok) {
            const errorText = await response.text();
            console.error(`第 ${pageNum} 頁 Gemini API 回應錯誤:`, response.status, errorText);
            continue;
          }
          const data = await response.json();
          if (data.error) {
            console.error(`第 ${pageNum} 頁 Gemini API 錯誤:`, data.error);
            continue;
          }
          if (!data.candidates || data.candidates.length === 0) {
            console.warn(`第 ${pageNum} 頁 API 回應中沒有結果`);
            continue;
          }
          if (isCancelled) {
            Swal.close();
            app.pdfViewer.currentPageNumber = originalCurrentPage;
            await goToPageAndWaitForRender(originalCurrentPage, 800);
            await Swal.fire({
              icon: 'info',
              title: '已取消',
              text: `已取消 AI 標記，已處理 ${pageNum - 1} / ${totalPages} 頁`,
              timer: 2000,
              showConfirmButton: false,
              customClass: { popup: 'swal-high-z-index' },
            });
            return;
          }
          const aiResponse = data.candidates[0].content.parts[0].text;
          let markSuggestions = [];
          try {
            const jsonMatch = aiResponse.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
              const jsonData = JSON.parse(jsonMatch[0]);
              if (jsonData.marks && Array.isArray(jsonData.marks)) {
                markSuggestions = jsonData.marks;
              }
            }
          } catch (parseError) {
            console.error(`第 ${pageNum} 頁 解析 AI 回應失敗:`, parseError);
            continue;
          }
          if (markSuggestions.length === 0) {
            console.log(`第 ${pageNum} 頁 AI 無法識別出需要標記的內容`);
            continue;
          }
          const textLayer = document.querySelector(
            `#viewerContainer .page[data-page-number="${pageNum}"] .textLayer`
          );
          if (!textLayer) {
            console.warn(`第 ${pageNum} 頁 找不到文字層`);
            continue;
          }
          const textLayerRect = textLayer.getBoundingClientRect();
          const markLayer = document.querySelector(`.mark-layer[data-page-num="${pageNum}"]`);
          if (!markLayer) {
            console.warn(`第 ${pageNum} 頁 找不到標記層`);
            continue;
          }
          const markLayerRect = markLayer.getBoundingClientRect();
          const spans = textLayer.querySelectorAll('span');
          let pageCreatedMarksCount = 0;
          const spanData = [];
          for (let i = 0; i < spans.length; i++) {
            const span = spans[i];
            const rect = span.getBoundingClientRect();
            spanData.push({
              span,
              text: span.textContent || '',
              x: rect.left - textLayerRect.left,
              y: rect.top - textLayerRect.top,
              right: rect.right - textLayerRect.left,
              bottom: rect.bottom - textLayerRect.top,
              centerX: (rect.left + rect.right) / 2 - textLayerRect.left,
              centerY: (rect.top + rect.bottom) / 2 - textLayerRect.top,
            });
          }
          let fullTextFromSpans = '';
          spanData.forEach(data => {
            fullTextFromSpans += data.text;
          });
          const pk = mkey(pageNum);
          for (const suggestion of markSuggestions) {
            const targetText = (suggestion.text || '').trim();
            if (!targetText) continue;
            const targetIndex = fullTextFromSpans.indexOf(targetText);
            const tryAddMark = (startSpanIndex, endSpanIndex) => {
              if (startSpanIndex === -1 || endSpanIndex === -1) return;
              const relevantSpans = [];
              for (let i = startSpanIndex; i <= endSpanIndex; i++) {
                relevantSpans.push(spanData[i]);
              }
              let minX = Infinity;
              let minY = Infinity;
              let maxX = -Infinity;
              let maxY = -Infinity;
              relevantSpans.forEach(data => {
                minX = Math.min(minX, data.x);
                minY = Math.min(minY, data.y);
                maxX = Math.max(maxX, data.right);
                maxY = Math.max(maxY, data.bottom);
              });
              const x = minX;
              const y = minY;
              const width = maxX - minX;
              const height = maxY - minY;
              const markX = x + (textLayerRect.left - markLayerRect.left);
              const markY = y + (textLayerRect.top - markLayerRect.top);
              if (width <= 10 || height <= 10) return;
              if (isSupplementMode && marks[pk]) {
                let overlaps = false;
                for (const existingMark of marks[pk]) {
                  const overlapX = Math.max(
                    0,
                    Math.min(markX + width, existingMark.x + existingMark.width) -
                      Math.max(markX, existingMark.x)
                  );
                  const overlapY = Math.max(
                    0,
                    Math.min(markY + height, existingMark.y + existingMark.height) -
                      Math.max(markY, existingMark.y)
                  );
                  const overlapArea = overlapX * overlapY;
                  const newMarkArea = width * height;
                  const existingMarkArea = existingMark.width * existingMark.height;
                  if (overlapArea > newMarkArea * 0.3 || overlapArea > existingMarkArea * 0.3) {
                    overlaps = true;
                    break;
                  }
                }
                if (overlaps) {
                  console.log(`第 ${pageNum} 頁 跳過與現有標記重疊的標記:`, targetText.substring(0, 20));
                  return;
                }
              }
              if (!marks[pk]) marks[pk] = [];
              const mark = {
                id: markIdCounter++,
                x: markX,
                y: markY,
                width,
                height,
                text: targetText,
                replacementText: '',
                locked: false,
              };
              const pdfCoords = viewportToPdfCoords(markX, markY, width, height, pageNum);
              if (pdfCoords) {
                mark.pdfL = pdfCoords.pdfL;
                mark.pdfT = pdfCoords.pdfT;
                mark.pdfR = pdfCoords.pdfR;
                mark.pdfB = pdfCoords.pdfB;
              }
              marks[pk].push(mark);
              pageCreatedMarksCount++;
            };
            if (targetIndex === -1) {
              const targetTextNoSpace = targetText.replace(/\s+/g, '');
              const fullTextNoSpace = fullTextFromSpans.replace(/\s+/g, '');
              const fuzzyIndex = fullTextNoSpace.indexOf(targetTextNoSpace);
              if (fuzzyIndex === -1) continue;
              let charCount = 0;
              let startSpanIndex = -1;
              let endSpanIndex = -1;
              for (let i = 0; i < spanData.length; i++) {
                charCount += spanData[i].text.replace(/\s+/g, '').length;
                if (startSpanIndex === -1 && charCount > fuzzyIndex) startSpanIndex = i;
                if (charCount >= fuzzyIndex + targetTextNoSpace.length) {
                  endSpanIndex = i;
                  break;
                }
              }
              tryAddMark(startSpanIndex, endSpanIndex);
            } else {
              let charCount = 0;
              let startSpanIndex = -1;
              let endSpanIndex = -1;
              for (let i = 0; i < spanData.length; i++) {
                charCount += spanData[i].text.length;
                if (startSpanIndex === -1 && charCount > targetIndex) startSpanIndex = i;
                if (charCount >= targetIndex + targetText.length) {
                  endSpanIndex = i;
                  break;
                }
              }
              tryAddMark(startSpanIndex, endSpanIndex);
            }
          }
          totalCreatedMarksCount += pageCreatedMarksCount;
          console.log(`第 ${pageNum} 頁創建了 ${pageCreatedMarksCount} 個標記`);
        } catch (pageError) {
          console.error(`第 ${pageNum} 頁處理失敗:`, pageError);
        }
      }
      app.pdfViewer.currentPageNumber = originalCurrentPage;
      await goToPageAndWaitForRender(originalCurrentPage, 2000);
      if (isCancelled) return;
      Swal.close();
      syncMarkIdCounterFromMarks();
      renderMarksAllPages();
      refreshAllMarkVisuals();
      if (totalCreatedMarksCount > 0) {
        // AI 完成後立即切到可視/可編輯狀態，避免需先取消標記再重進才看得到區塊。
        isManualMarkingEnabled = true;
        setTextLayersPointerEvents('none');
        document.querySelectorAll('.mark-layer').forEach(l => {
          l.classList.add('marking-mode');
          l.style.pointerEvents = 'auto';
        });
        bindManualMarkingPointerHandlers();
        renderMarksAllPages();
        refreshAllMarkVisuals();

        await saveMarksToPdf(true);
        await Swal.fire({
          icon: 'success',
          title: isSupplementMode ? 'AI 補充標記完成' : 'AI 標記完成',
          text: isSupplementMode
            ? `已成功處理 ${totalPages} 頁，共補充 ${totalCreatedMarksCount} 個新標記，現有標記已保留，您可以手動編輯或調整這些標記`
            : `已成功處理 ${totalPages} 頁，共創建 ${totalCreatedMarksCount} 個標記，您可以手動編輯或調整這些標記`,
          timer: 2000,
          showConfirmButton: false,
          customClass: { popup: 'swal-high-z-index' },
        });
      } else {
        await Swal.fire({
          icon: 'warning',
          title: '無法創建標記',
          text: `已處理 ${totalPages} 頁，但無法創建任何標記`,
          timer: 2000,
          showConfirmButton: false,
          customClass: { popup: 'swal-high-z-index' },
        });
      }
    } catch (error) {
      console.error('AI 標記失敗:', error);
      Swal.close();
      await Swal.fire({
        icon: 'error',
        title: 'AI 標記失敗',
        text: '無法完成 AI 標記: ' + (error.message || String(error)),
        customClass: { popup: 'swal-high-z-index' },
      });
    }
  }

  async function enterMarkingMode() {
    const app = getApp();
    if (!app?.pdfDocument) {
      ensureSwalLocal().then(() =>
        window.Swal?.fire({ icon: 'warning', title: '無法標記', text: '請先開啟 PDF' })
      );
      return;
    }
    // 使用者點按「標記區塊」當下再同步一次來源判斷，避免首次載入尚未走完 documentloaded 初始化造成漏判
    await refreshIdbMarkSourceCloudCache();
    updateCloudMode();
    cloudMarksBackup = null;
    cloudMarkIdCounterBackup = 0;
    if (isCloudFileMode) {
      // 進入標記模式後，使用者的任何修改都只允許在畫面上操作；取消標記必須回到原狀
      cloudMarksBackup = JSON.parse(JSON.stringify(marks));
      cloudMarkIdCounterBackup = markIdCounter;
      ensureSwalLocal().then(() =>
        window.Swal?.fire({
          icon: 'warning',
          title: '網路 PDF',
          html: '<div style="text-align:left">可試用標記介面，但<strong>無法寫回檔案</strong>。請改用本機或已快取之 PDF 以保存標記。</div>',
          confirmButtonText: '我知道了',
        })
      );
    }
    isMarkingMode = true;
    isManualMarkingEnabled = false;
    document.body.classList.add('marking-mode-active');
    const apiKeyModal = document.getElementById('apiKeyModal');
    if (apiKeyModal) apiKeyModal.style.display = 'none';
    const bar = document.getElementById('markingToolbar');
    if (bar) {
      bar.classList.add('show');
      bar.style.display = 'flex';
    }
    const mb = document.getElementById('markBlockButton');
    if (mb) mb.textContent = '取消標記';
    syncAllMarkLayersSize();
    renderMarksAllPages();
    refreshAllMarkVisuals();
    const mm = document.getElementById('manualMarkButtonToolbar');
    if (mm) mm.style.display = 'inline-block';
    const cx = document.getElementById('cancelMarkingButtonToolbar');
    if (cx) {
      cx.style.display = 'inline-block';
      cx.style.visibility = 'visible';
    }
    setMarkingToolbarExtrasVisible(true);
  }

  async function enableManualMarking() {
    if (!isMarkingMode) return;
    // 第一次啟用手動標記、或仍無任何標記時：強制從 PDF 再載入（等同資料層「重新整理」）
    if (!isManualMarkingEnabled || !hasMarksData()) {
      try {
        await reloadMarksFromPdfDocument();
      } catch (_) {
        /* ignore */
      }
    }
    isManualMarkingEnabled = true;
    setTextLayersPointerEvents('none');
    syncAllMarkLayersSize();
    renderMarksAllPages();
    refreshAllMarkVisuals();
    document.querySelectorAll('.mark-layer').forEach(l => {
      l.classList.add('marking-mode');
      l.style.pointerEvents = 'auto';
    });
    bindManualMarkingPointerHandlers();
  }

  function exitMarkingMode(shouldReload = false) {
    isMarkingMode = false;
    isManualMarkingEnabled = false;
    currentSessionPassword = null;
    document.body.classList.remove('marking-mode-active');
    unbindManualMarkingPointerHandlers();
    setTextLayersPointerEvents('auto');
    const bar = document.getElementById('markingToolbar');
    if (bar) {
      bar.classList.remove('show');
      bar.style.display = 'none';
    }
    const mb = document.getElementById('markBlockButton');
    if (mb) mb.textContent = '標記區塊';
    const cx = document.getElementById('cancelMarkingButtonToolbar');
    if (cx) cx.style.display = 'none';
    const mm = document.getElementById('manualMarkButtonToolbar');
    if (mm) mm.style.display = 'none';
    setMarkingToolbarExtrasVisible(false);
    document.querySelectorAll('.mark-layer').forEach(l => {
      l.classList.remove('marking-mode');
      l.style.pointerEvents = 'none';
    });

    // 雲端模式下取消標記：還原先前備份的 marks（刪除/編輯/新增都應捨棄）
    if (cloudMarksBackup) {
      Object.keys(marks).forEach(k => delete marks[k]);
      Object.assign(marks, JSON.parse(JSON.stringify(cloudMarksBackup)));
      markIdCounter = cloudMarkIdCounterBackup;
      cloudMarksBackup = null;
      cloudMarkIdCounterBackup = 0;
    }

    // 無論是否雲端，都重建一次標記節點與事件綁定，避免編輯流程後殘留舊 DOM 導致點擊不發音
    renderMarksAllPages();
    refreshAllMarkVisuals();

    if (shouldReload) {
      setTimeout(() => {
        window.location.reload();
      }, 0);
    }
  }

  async function toggleMarkingModeFromUi() {
    if (isMarkingMode) {
      exitMarkingMode();
      return;
    }
    const ok = await checkPasswordProtection();
    if (!ok) return;
    await enterMarkingMode();
  }

  function onDocumentLoaded() {
    void (async () => {
      await reloadMarksFromPdfDocument();
      scheduleMarksDomRetries();
    })();
  }

  /** viewer-main 先 await GAS 時，PDF 常在綁定 eventBus 之前就觸發 pagesloaded，導致永不跑標記／軟重載；與事件共用此流程並於綁定後補跑。 */
  let markPagesPipelineBusy = false;
  let markPagesPipelineQueued = false;

  async function runPagesLoadedMarkPipeline() {
    if (markPagesPipelineBusy) {
      markPagesPipelineQueued = true;
      return;
    }
    if (!getApp()?.pdfDocument) return;
    markPagesPipelineBusy = true;
    try {
      await refreshIdbMarkSourceCloudCache();
      updateCloudMode();
      const appPl = getApp();
      const fp = appPl?.pdfDocument?.fingerprints?.[0] ?? '';
      const cloudish = isCloudFileMode;

      await refreshOriginalPdfBytes();
      await waitForPDFLibReady(4500);
      await loadMarksFromPdf();
      if (cloudish && !hasMarksData()) {
        await new Promise(r => setTimeout(r, 240));
        await refreshOriginalPdfBytes();
        await loadMarksFromPdf();
      }
      scheduleMarksDomRetries();

      const n = appPl?.pdfViewer?.pagesCount ?? 0;
      for (let p = 1; p <= n; p++) {
        if (!isMarkLayerLastInPage(p)) {
          ensureMarkLayer(p);
          renderMarksPage(p);
        }
      }
      refreshAllMarkVisuals();

      const shouldSoft = cloudish && fp && markViewerSoftReloadFingerprint !== fp;
      if (shouldSoft) {
        const ok = await softReloadViewerWithCurrentPdfBytes();
        if (ok) markViewerSoftReloadFingerprint = fp;
      }
    } finally {
      markPagesPipelineBusy = false;
      if (markPagesPipelineQueued) {
        markPagesPipelineQueued = false;
        void runPagesLoadedMarkPipeline();
      }
    }
  }

  function initMarkFeatureBindings() {
    if (!initMarkFeatureBindings._uiBound) {
      initMarkFeatureBindings._uiBound = true;
      document.getElementById('markBlockButton')?.addEventListener('click', toggleMarkingModeFromUi);
      document.getElementById('cancelMarkingButtonToolbar')?.addEventListener('click', () => {
        if (isMarkingMode) exitMarkingMode(true);
      });
      document.getElementById('manualMarkButtonToolbar')?.addEventListener('click', e => {
        e.preventDefault();
        void enableManualMarking();
      });
      document.getElementById('aiMarkButtonToolbar')?.addEventListener('click', () => {
        void aiMarkPdf();
      });
      document.getElementById('setPasswordButtonToolbar')?.addEventListener('click', () => {
        void setPasswordForPdf();
      });
      document.getElementById('saveMarkedPdfButtonToolbar')?.addEventListener('click', () => {
        void saveMarksToPdf(true);
      });
      document.getElementById('downloadMarkedPdfButtonToolbar')?.addEventListener('click', () => {
        void downloadMarkedPdf();
      });
      document.getElementById('clearMarksButtonToolbar')?.addEventListener('click', () => {
        void clearAllMarks();
      });
    }

    const app = getApp();
    const eb = app?.eventBus;
    if (eb && !initMarkFeatureBindings._on) {
      initMarkFeatureBindings._on = true;
      eb._on('documentloaded', onDocumentLoaded);
      eb._on('pagesloaded', () => {
        void runPagesLoadedMarkPipeline();
      });
      const tryCatchupPagesLoadedPipeline = () => {
        const a = getApp();
        if (!a?.pdfDocument) return;
        const num = a.pdfDocument.numPages || 0;
        const cnt = a.pdfViewer?.pagesCount ?? 0;
        if (num > 0 && cnt >= num) {
          void runPagesLoadedMarkPipeline();
        }
      };
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          tryCatchupPagesLoadedPipeline();
          setTimeout(tryCatchupPagesLoadedPipeline, 400);
          setTimeout(tryCatchupPagesLoadedPipeline, 1200);
        });
      });
      const bumpMarkLayerAfterCanvas = e => {
        const pn = e?.pageNumber;
        if (!pn) return;
        ensureMarkLayer(pn);
        renderMarksPage(pn);
        refreshAllMarkVisuals();
      };
      // 僅在「文字層真的插在標記層後面」時再 bump；避免本機檔每頁多刷兩三次 renderMarksPage 造成異常。
      // 不監聽 annotationlayerrendered：表單／連結註解與標記層順序交錯時，額外 appendChild 易引發本機行為異常。
      const bumpMarkLayerIfCoveredByTextLayer = e => {
        const pn = e?.pageNumber;
        if (!pn || e.error) return;
        if (isMarkLayerLastInPage(pn)) return;
        ensureMarkLayer(pn);
        renderMarksPage(pn);
        refreshAllMarkVisuals();
      };
      eb._on('pagerendered', bumpMarkLayerAfterCanvas);
      eb._on('textlayerrendered', bumpMarkLayerIfCoveredByTextLayer);
      const rescale = () => {
        syncAllMarkLayersSize();
        renderMarksAllPages();
        refreshAllMarkVisuals();
      };
      eb._on('scalechanging', rescale);
      eb._on('rotationchanging', rescale);
    }
    // 不在此呼叫 onDocumentLoaded：若 pdf 已載入會與 runPagesLoadedMarkPipeline 並行搶改 marks；改由 pagesloaded、documentloaded 事件與綁定後 rAF 補跑處理。
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initMarkFeatureBindings);
  } else {
    initMarkFeatureBindings();
  }
  setTimeout(initMarkFeatureBindings, 1200);

  document.addEventListener('click', onDocumentMarkHitCapture, true);
  document.addEventListener('pointerup', onDocumentMarkPointerUpCapture, true);
})();
