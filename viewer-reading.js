// 框選朗讀：與 PDF.js viewer 整合

(function () {
  const viewerContainer = document.getElementById('viewerContainer');
  if (!viewerContainer) return;

  // 建立框與控制元件
  const frame = document.createElement('div');
  frame.id = 'readingFrame';
  frame.innerHTML = `
    <div id="readingLabel">框選朗讀</div>
    <button id="readingExpandX"></button>
    <button id="readingExpandY"></button>
  `;
  document.body.appendChild(frame);

  // 讓左右/上下使用同一個字元渲染，左右用旋轉達成（iPad 上可避免字型替換不一致）
  const expandXBtn = frame.querySelector('#readingExpandX');
  const expandYBtn = frame.querySelector('#readingExpandY');
  if (expandXBtn) expandXBtn.innerHTML = '<span class="reading-arrow reading-arrow-x">↕</span>';
  if (expandYBtn) expandYBtn.innerHTML = '<span class="reading-arrow reading-arrow-y">↕</span>';

  // 若目前頁面沒有 voiceSelect，下方自動建立一個與「語音報讀系統個人版」相同選單
  if (!document.getElementById('voiceSelect')) {
    const voiceSelect = document.createElement('select');
    voiceSelect.id = 'voiceSelect';
    voiceSelect.style.position = 'fixed';
    // 與語速控制橫向並列，位於其左側，多預留寬度避免重疊
    voiceSelect.style.top = '70px';
    voiceSelect.style.right = '450px';
    voiceSelect.style.zIndex = '3100';
    voiceSelect.style.fontSize = '16px';
    voiceSelect.style.height = '40px';
    voiceSelect.style.width = '260px';
    voiceSelect.style.minWidth = '220px';
    voiceSelect.style.maxWidth = '32vw';
    voiceSelect.style.padding = '6px 10px';
    voiceSelect.style.borderRadius = '8px';
    voiceSelect.innerHTML = `
      <option value="zh-TW-HsiaoChenNeural">雲端：中文 曉臻 (女性)</option>
      <option value="zh-TW-YunJheNeural">雲端：中文 允哲 (男性)</option>
      <option value="zh-TW-HsiaoYuNeural">雲端：中文 曉雨 (女性)</option>
      <option value="en-US-JennyNeural">雲端：英文 Jenny (女性)</option>
      <option value="en-US-GuyNeural">雲端：英文 Guy (男性)</option>
      <option value="en-US-AriaNeural">雲端：英文 Aria (女性)</option>
      <option value="en-US-ChristopherNeural">雲端：英文 Christ (男性)</option>
      <option value="local-zh-female">本機：中文 (瀏覽器)</option>
      <option value="local-en-female">本機：英文 (瀏覽器)</option>
    `;
    document.body.appendChild(voiceSelect);
  }

  // 若目前頁面沒有 speakSpeed，建立與原頁面相容的語速控制（-50 ~ +50）
  if (!document.getElementById('speakSpeed')) {
    const speedWrap = document.createElement('div');
    speedWrap.id = 'readingSpeedWrap';
    speedWrap.style.position = 'fixed';
    speedWrap.style.top = '70px';
    speedWrap.style.right = '150px';
    speedWrap.style.zIndex = '3100';
    speedWrap.style.background = 'rgba(255,255,255,0.9)';
    speedWrap.style.padding = '2px 6px';
    speedWrap.style.borderRadius = '6px';
    speedWrap.style.display = 'flex';
    speedWrap.style.alignItems = 'center';
    speedWrap.style.gap = '6px';
    speedWrap.style.minHeight = '0';

    const speedText = document.createElement('span');
    speedText.textContent = '語速';
    speedText.style.fontSize = '14px';

    const speedInput = document.createElement('input');
    speedInput.type = 'range';
    speedInput.id = 'speakSpeed';
    speedInput.min = '-50';
    speedInput.max = '50';
    speedInput.value = '-30';
    speedInput.style.width = '130px';
    speedInput.style.height = '32px';
    speedInput.style.webkitAppearance = 'none';
    speedInput.style.appearance = 'none';
    speedInput.style.background = 'transparent';
    speedInput.style.cursor = 'pointer';
    speedInput.style.touchAction = 'pan-y';

    const speedStyleId = 'reading-speed-slider-style';
    if (!document.getElementById(speedStyleId)) {
      const st = document.createElement('style');
      st.id = speedStyleId;
      st.textContent = `
#speakSpeed::-webkit-slider-runnable-track{
  height:10px;
  background:#cfd4db;
  border-radius:999px;
}
#speakSpeed::-webkit-slider-thumb{
  -webkit-appearance:none;
  appearance:none;
  width:28px;
  height:28px;
  margin-top:-9px;
  border-radius:50%;
  background:#2e7dff;
  border:1px solid #1f5fcc;
}
#speakSpeed::-moz-range-track{
  height:10px;
  background:#cfd4db;
  border-radius:999px;
}
#speakSpeed::-moz-range-thumb{
  width:28px;
  height:28px;
  border-radius:50%;
  background:#2e7dff;
  border:1px solid #1f5fcc;
}
      `;
      document.head.appendChild(st);
    }

    const speedPercent = document.createElement('span');
    speedPercent.id = 'speedPercent';
    speedPercent.textContent = '-30%';
    speedPercent.style.fontSize = '14px';
    speedPercent.style.minWidth = '48px';

    speedInput.addEventListener('input', () => {
      const v = parseFloat(speedInput.value || '0');
      speedPercent.textContent = `${v > 0 ? '+' : ''}${v}%`;
    });

    speedWrap.appendChild(speedText);
    speedWrap.appendChild(speedInput);
    speedWrap.appendChild(speedPercent);
    document.body.appendChild(speedWrap);
  }
  ensureTtsPlaybackBar();
  setTimeout(layoutTtsPlaybackBar, 0);
  setTimeout(layoutTtsPlaybackBar, 300);

  let dragging = false;
  let resizing = false;
  let resizeDir = ''; // 'x' 或 'y'
  let startX = 0, startY = 0;
  let startLeft = 0, startTop = 0, startWidth = 0, startHeight = 0;
  const MIN_FRAME_WIDTH = 18;
  const MIN_FRAME_HEIGHT = 18;

  // 預設位置與大小：依照實際工具列高度與按鈕尺寸計算，
  // 讓「網頁初始」與「之後點觸控朗讀」看到的框一致。
  function resetFrameToDefaultPosition() {
    const label = document.getElementById('readingLabel');
    const labelWidth = label ? label.offsetWidth : 120;
    const labelHeight = label ? label.offsetHeight : 40;
    const toolbar = document.getElementById('toolbarContainer');
    const toolbarRect = toolbar ? toolbar.getBoundingClientRect() : null;
    frame.style.display = 'block';
    // 靠左一點，避免貼邊
    frame.style.left = '10px';
    // 以實際工具列高度為基準，再多往下 48px，確保完全不遮到工具列且視覺上合理
    if (toolbarRect) {
      frame.style.top = `${toolbarRect.bottom + 48}px`;
    } else {
      frame.style.top = `${labelHeight * 4 + 40}px`;
    }
    // 預設框固定為 40x40px
    frame.style.width = '40px';
    frame.style.height = '40px';
  }

  // 讓外部 onload 也能呼叫預設定位
  window.resetReadingFrameDefault = resetFrameToDefaultPosition;

  function showFrame() {
    if (frame.style.display === 'block') return;
    const vw = window.innerWidth || document.documentElement.clientWidth;
    const vh = window.innerHeight || document.documentElement.clientHeight;
    const w = Math.max(260, vw * 0.45);
    const h = Math.max(120, vh * 0.25);
    frame.style.left = (vw * 0.2) + 'px';
    frame.style.top = (vh * 0.25) + 'px';
    frame.style.width = w + 'px';
    frame.style.height = h + 'px';
    frame.style.display = 'block';
    frame.classList.add('active');
  }

  function isRectHit(frameRect, rect) {
    const interLeft = Math.max(rect.left, frameRect.left);
    const interRight = Math.min(rect.right, frameRect.right);
    const interTop = Math.max(rect.top, frameRect.top);
    const interBottom = Math.min(rect.bottom, frameRect.bottom);
    const interWidth = interRight - interLeft;
    const interHeight = interBottom - interTop;
    if (interWidth <= 0 || interHeight <= 0) return false;

    const rectArea = (rect.width || 1) * (rect.height || 1);
    const overlapArea = interWidth * interHeight;
    const overlapRatio = overlapArea / rectArea;
    return overlapRatio >= 0.55;
  }

  /** PDF 文字層常把 ① 拆成「○」「１」等不同 span；若以字元座標全域排序會與閱讀順序錯亂。改依 span 順序 + span 內字元順序。 */
  function compareSpanReadingOrder(a, b) {
    const ra = a.getBoundingClientRect();
    const rb = b.getBoundingClientRect();
    const dy = ra.top - rb.top;
    if (Math.abs(dy) > 6) return dy;
    return ra.left - rb.left;
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

  // 與「標記區塊」相同：粗體通常是重疊 span，先做 span 層級去重。
  function dedupeBoldOverlaySpans(spans) {
    const sorted = [...spans].sort(compareSpanReadingOrder);
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

  // 與「標記區塊」一致：收斂連續重複片段，避免粗體疊字重複送到 TTS。
  function collapseAdjacentDuplicateCjkRuns(s) {
    if (!s || s.length < 2) return s;
    let t = s;
    let prev = '';
    let guard = 0;
    const cjkSpan = String.raw`(?:[\u3000-\u9fff\u2460-\u2473])`;
    while (prev !== t && guard++ < 24) {
      prev = t;
      t = t.replace(new RegExp(`(${cjkSpan}{2,}?)(?:\\1)+`, 'gu'), '$1');
      t = t.replace(new RegExp(`(${cjkSpan})(?:\\1)+`, 'gu'), '$1');
    }
    return t;
  }

  // 收集框內文字，並同時回傳被框住的字元矩形
  function collectTextAndSpansInFrame() {
    const frameRect = frame.getBoundingClientRect();
    const spans = Array.from(document.querySelectorAll('#viewerContainer .textLayer span'));
    const hitRects = [];
    const seenHighlightKeys = new Set();
    const orderedGlyphs = [];
    const seenGlyphPosByChar = new Map();

    const sortedSpans = dedupeBoldOverlaySpans(spans
      .filter(span => {
        const spanRect = span.getBoundingClientRect();
        return !(
          spanRect.right < frameRect.left ||
          spanRect.left > frameRect.right ||
          spanRect.bottom < frameRect.top ||
          spanRect.top > frameRect.bottom
        );
      })
      .sort(compareSpanReadingOrder));

    sortedSpans.forEach(span => {
      const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
      let textNode;
      while ((textNode = walker.nextNode())) {
        const txt = textNode.nodeValue || '';
        if (!txt) continue;

        for (let i = 0; i < txt.length; i++) {
          const range = document.createRange();
          range.setStart(textNode, i);
          range.setEnd(textNode, i + 1);
          const rectList = range.getClientRects();
          let hit = false;
          let representativeRect = null;

          for (const r of rectList) {
            if (!isRectHit(frameRect, r)) continue;
            hit = true;
            representativeRect ||= r;
            const rectKey = `${Math.round(r.left)}|${Math.round(r.top)}|${Math.round(r.width)}|${Math.round(r.height)}`;
            if (!seenHighlightKeys.has(rectKey)) {
              seenHighlightKeys.add(rectKey);
              const textLayerEl = span.closest('.textLayer');
              const layerRect = textLayerEl?.getBoundingClientRect();
              const pageEl = span.closest('.page');
              const pageNumber = parseInt(pageEl?.dataset?.pageNumber || '0', 10) || 0;
              const relLeft = layerRect ? (r.left - layerRect.left) / (layerRect.width || 1) : 0;
              const relTop = layerRect ? (r.top - layerRect.top) / (layerRect.height || 1) : 0;
              const relWidth = layerRect ? r.width / (layerRect.width || 1) : 0;
              const relHeight = layerRect ? r.height / (layerRect.height || 1) : 0;
              hitRects.push({
                left: r.left,
                top: r.top,
                width: r.width,
                height: r.height,
                pageNumber,
                layerClass: 'textLayer',
                relLeft,
                relTop,
                relWidth,
                relHeight,
              });
            }
          }

          if (hit && representativeRect) {
            const ch = txt[i];
            const gx = representativeRect.left;
            const gy = representativeRect.top;
            const prior = seenGlyphPosByChar.get(ch) || [];
            const duplicatedByPosition = prior.some(p =>
              Math.abs(p.x - gx) <= 3 && Math.abs(p.y - gy) <= 3
            );
            if (!duplicatedByPosition) {
              prior.push({ x: gx, y: gy });
              seenGlyphPosByChar.set(ch, prior);
              orderedGlyphs.push({
                ch,
                x: gx,
                y: gy,
                w: representativeRect.width,
                h: representativeRect.height,
              });
            }
          }
        }
      }
    });

    // 粗體/陰影：連續同字元且座標極接近只留一個（保持串接順序）
    const deduped = [];
    for (const g of orderedGlyphs) {
      const prev = deduped[deduped.length - 1];
      if (
        prev &&
        prev.ch === g.ch &&
        Math.abs(prev.x - g.x) <= 3 &&
        Math.abs(prev.y - g.y) <= 3
      ) {
        continue;
      }
      deduped.push(g);
    }

    const text = collapseAdjacentDuplicateCjkRuns(deduped.map(g => g.ch).join(''));
    return { text, rects: hitRects };
  }

  function handleStart(e) {
    const rect = frame.getBoundingClientRect();
    const point = e.touches ? e.touches[0] : e;
    startX = point.clientX;
    startY = point.clientY;
    startLeft = rect.left;
    startTop = rect.top;
    startWidth = rect.width;
    startHeight = rect.height;

    // 每次開始拖曳框時，先清除前一次的高亮
    if (typeof clearHighlights === 'function') {
      clearHighlights();
    }

    dragging = true;
    resizing = false;
    resizeDir = '';
    e.preventDefault();
  }

  frame.addEventListener('mousedown', (e) => {
    const target = e.target;
    if (target.id === 'readingLabel' || target.id === 'readingExpandX' || target.id === 'readingExpandY') {
      return; // 這些由各自監聽處理
    }
    handleStart(e);
  });

  // 觸控拖曳（iOS / 平板）
  frame.addEventListener('touchstart', (e) => {
    const target = e.target;
    if (target.id === 'readingLabel' || target.id === 'readingExpandX' || target.id === 'readingExpandY') {
      return;
    }
    handleStart(e);
  }, { passive: false });

  function handleMove(e) {
    if (!dragging && !resizing) return;
    const point = e.touches ? e.touches[0] : e;
    const dx = point.clientX - startX;
    const dy = point.clientY - startY;

    const vw = window.innerWidth || document.documentElement.clientWidth;
    const vh = window.innerHeight || document.documentElement.clientHeight;

    if (dragging) {
      let newLeft = startLeft + dx;
      let newTop = startTop + dy;
      newLeft = Math.max(0, Math.min(newLeft, vw - frame.offsetWidth));
      newTop = Math.max(0, Math.min(newTop, vh - frame.offsetHeight));
      frame.style.left = newLeft + 'px';
      frame.style.top = newTop + 'px';
    } else if (resizing) {
      if (resizeDir === 'x') {
        const newWidth = Math.max(MIN_FRAME_WIDTH, startWidth + dx);
        frame.style.width = Math.min(vw - startLeft, newWidth) + 'px';
      } else if (resizeDir === 'y') {
        const newHeight = Math.max(MIN_FRAME_HEIGHT, startHeight + dy);
        frame.style.height = Math.min(vh - startTop, newHeight) + 'px';
      }
    }
  }

  document.addEventListener('mousemove', (e) => {
    handleMove(e);
  });

  document.addEventListener('touchmove', (e) => {
    const beforeDragging = dragging;
    const beforeResizing = resizing;
    handleMove(e);
    // 只有在真的在拖曳/拉伸框時才攔截觸控事件，避免影響 iOS 上的滑軌等元件
    if (beforeDragging || beforeResizing) {
      e.preventDefault();
    }
  }, { passive: false });

  document.addEventListener('mouseup', () => {
    dragging = false;
    resizing = false;
    resizeDir = '';
  });

  document.addEventListener('touchend', () => {
    dragging = false;
    resizing = false;
    resizeDir = '';
  });

  let lastHighlighted = [];
  let currentReadToken = 0;
  let frameReading = false;
  let fullPageReading = false;
  let fullPageReadToken = 0;
  let fullPagePaused = false;
  let pauseWaiters = [];
  let touchSelectEnabled = false;
  let touchReadToken = 0;
  let markBlockReadToken = 0;
  let pendingGlyphs = [];
  let pendingGlyphKeySet = new Set();
  let pendingReadButton = null;
  let pendingSelectionReading = false;
  const readHighlightModelByPage = new Map();
  let isImageRecognizing = false;

  function getRatePercentFromUI() {
    const speedInput = document.getElementById('speakSpeed');
    const v = speedInput ? parseFloat(speedInput.value || '0') : 0;
    return Math.max(-50, Math.min(50, Number.isFinite(v) ? v : 0));
  }

  // --- PDF 快取（比照語音報讀系統個人版，用 IndexedDB 儲存與清除） ---
  const VIEWER_IDB_DB_NAME = 'pdfjs-viewer';
  const VIEWER_IDB_STORE_NAME = 'files';
  const VIEWER_IDB_LAST_FILE_KEY = 'last-opened-file';

  function openViewerDBLocal() {
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

  // 動態載入 SweetAlert2，供圖片辨識時顯示動態視窗使用（比照語音報讀系統個人版）
  let swalLoaded = false;
  function ensureSwal() {
    return new Promise(resolve => {
      if (swalLoaded || window.Swal) {
        swalLoaded = true;
        resolve();
        return;
      }
      const script = document.createElement('script');
      script.src = 'https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js';
      script.onload = () => {
        swalLoaded = true;
        resolve();
      };
      script.onerror = () => resolve(); // 載入失敗就靜默略過，不影響主要功能
      document.head.appendChild(script);
    });
  }

  // --- IndexedDB：比照語音報讀系統個人版，使用 pdfDB/settings 儲存 geminiApiKey ---
  const GEMINI_DB_NAME = 'pdfDB';
  const GEMINI_DB_VERSION = 3;

  function openGeminiDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(GEMINI_DB_NAME, GEMINI_DB_VERSION);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (db.objectStoreNames.contains('pdfFiles')) {
          db.deleteObjectStore('pdfFiles');
        }
        db.createObjectStore('pdfFiles', { keyPath: 'id' });
        if (!db.objectStoreNames.contains('settings')) {
          db.createObjectStore('settings', { keyPath: 'id' });
        }
      };
      request.onsuccess = (e) => resolve(e.target.result);
      request.onerror = (e) => reject(e.target.error);
    });
  }

  async function saveGeminiApiKey(key) {
    try {
      const db = await openGeminiDB();
      const transaction = db.transaction('settings', 'readwrite');
      const store = transaction.objectStore('settings');
      const dataToStore = { id: 'geminiApiKey', value: key || '' };
      const putRequest = store.put(dataToStore);
      await new Promise((resolve, reject) => {
        putRequest.onsuccess = () => {
          transaction.oncomplete = () => resolve();
          transaction.onerror = () => reject(transaction.error);
        };
        putRequest.onerror = () => reject(putRequest.error);
      });
    } catch (e) {
      console.error('保存 API KEY 到 IndexedDB 失敗：', e);
    }
  }

  async function loadGeminiApiKey() {
    try {
      const db = await openGeminiDB();
      const transaction = db.transaction('settings', 'readonly');
      const store = transaction.objectStore('settings');
      const request = store.get('geminiApiKey');
      const value = await new Promise((resolve, reject) => {
        request.onsuccess = (e) => {
          if (e.target.result) {
            resolve(e.target.result.value);
          } else {
            resolve(null);
          }
        };
        request.onerror = (e) => reject(e.target.error);
      });
      return value || '';
    } catch (e) {
      console.error('從 IndexedDB 讀取 API KEY 失敗：', e);
      return '';
    }
  }

  async function geminiImageDescribe(base64Image) {
    try {
      let apiKey = await loadGeminiApiKey();
      // 這裡假設已經先檢查過 apiKey 是否存在；若為空，直接丟錯
      if (!apiKey) {
        throw new Error('尚未設定 Gemini API KEY');
      }
      const response = await fetch(
        'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [
              {
                role: 'user',
                parts: [
                  {
                    text:
                      '你是一個給學習障礙學生報讀考題的報讀員，不要產生 markdown 的 * 符號，請先讀出圖片中的所有文字，遇到科學或數學符號請轉成中文文字的讀法，遇到數學分數時以中文讀法表示，如:幾分之幾，數學式子無須再特別說明它是一個公式，如果這些文字與圖片內容有關，請合併描述。然後用簡單句子描述圖片中的物件，不要贅述，不要描述場景、整體佈局或風格，不要使用「圖片中」、「顯示了」、「圖片顯示」等表述，直接描述內容即可。如果文字部分已能呈現全部，描述完不要再重複陳述圖片元素，如:弟弟吃了八分之五盒雞塊，不要再說「圖片中顯示了一行中文字和一個分數」等贅述。遇到英文字請如實呈現，不要把X讀成艾克斯。'
                  },
                  {
                    inline_data: {
                      mime_type: 'image/png',
                      data: base64Image
                    }
                  }
                ]
              }
            ]
          })
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('Gemini API 回應錯誤:', response.status, errorText);
        throw new Error(`API 請求失敗 (${response.status}): ${errorText}`);
      }

      const data = await response.json();
      if (data.error) {
        console.error('Gemini API 錯誤:', data.error);
        throw new Error(data.error.message || 'API 返回錯誤');
      }
      if (data.candidates && data.candidates.length > 0) {
        const text = data.candidates[0].content.parts[0].text;
        return text;
      }
      throw new Error('API 回應中沒有辨識結果');
    } catch (error) {
      console.error('Gemini 圖片辨識錯誤:', error);
      throw error;
    }
  }

  /* 供 viewer-reading-marks.js「編輯替換文字」與個人版相容（AI 圖片辨識） */
  if (typeof window !== 'undefined') {
    window.geminiImageDescribe = geminiImageDescribe;
    window.loadApiKeyFromIndexedDB = loadGeminiApiKey;
  }

  async function recognizeImageInFrame() {
    if (isImageRecognizing) return null;

    // 先確認是否有 API KEY，沒有就提示使用者設定
    const apiKey = await loadGeminiApiKey();
    if (!apiKey) {
      await ensureSwal();
      if (window.Swal) {
        window.Swal.fire({
          icon: 'warning',
          title: '尚未設定 API KEY',
          text: '請先點選左側齒輪按鈕，在「報讀系統進階設定」中輸入 Gemini API KEY 後再進行圖片辨識。',
          backdrop: false,
          customClass: { popup: 'swal-high-z-index' },
        });
      } else {
        alert('尚未設定 API KEY，請先點選左側齒輪按鈕設定 Gemini API KEY。');
      }
      return null;
    }
    const pageNumber =
      window.PDFViewerApplication?.pdfViewer?.currentPageNumber ||
      parseInt(document.getElementById('pageNumber')?.value || '1', 10) ||
      1;
    const pageCanvas = document.querySelector(
      `.page[data-page-number="${pageNumber}"] canvas`
    );
    if (!pageCanvas) return null;

    const frameRect = frame.getBoundingClientRect();
    const canvasRect = pageCanvas.getBoundingClientRect();

    const interLeft = Math.max(frameRect.left, canvasRect.left);
    const interRight = Math.min(frameRect.right, canvasRect.right);
    const interTop = Math.max(frameRect.top, canvasRect.top);
    const interBottom = Math.min(frameRect.bottom, canvasRect.bottom);
    const interWidth = interRight - interLeft;
    const interHeight = interBottom - interTop;
    if (interWidth <= 5 || interHeight <= 5) {
      return null;
    }

    const scaleX = pageCanvas.width / canvasRect.width;
    const scaleY = pageCanvas.height / canvasRect.height;
    const sx = (interLeft - canvasRect.left) * scaleX;
    const sy = (interTop - canvasRect.top) * scaleY;
    const sw = interWidth * scaleX;
    const sh = interHeight * scaleY;

    const cropCanvas = document.createElement('canvas');
    cropCanvas.width = Math.max(64, Math.round(sw));
    cropCanvas.height = Math.max(64, Math.round(sh));
    const ctx = cropCanvas.getContext('2d');
    ctx.drawImage(
      pageCanvas,
      sx,
      sy,
      sw,
      sh,
      0,
      0,
      cropCanvas.width,
      cropCanvas.height
    );

    const base64Image = cropCanvas
      .toDataURL('image/png')
      .split(',')[1];

    isImageRecognizing = true;
    let usingSwal = false;
    try {
      await ensureSwal();
      if (window.Swal) {
        usingSwal = true;
        window.Swal.fire({
          title: '圖片辨識中',
          text: '正在辨識框選區域的圖片，請稍候…',
          allowOutsideClick: false,
          allowEscapeKey: false,
          backdrop: false, // 背後網頁保持可見
          customClass: { popup: 'swal-high-z-index' },
          didOpen: () => {
            window.Swal.showLoading();
          },
        });
      }
      const text = await geminiImageDescribe(base64Image);
      if (usingSwal) {
        window.Swal.close();
      }
      return text;
    } catch (e) {
      if (usingSwal && window.Swal) {
        window.Swal.close();
        window.Swal.fire({
          icon: 'error',
          title: '圖片辨識失敗',
          text: e && e.message ? e.message : '呼叫圖片辨識服務時發生錯誤。',
          backdrop: false,
          customClass: { popup: 'swal-high-z-index' },
          timer: 2500,
          showConfirmButton: false,
        });
      }
      throw e;
    } finally {
      isImageRecognizing = false;
    }
  }

  function clearHighlights() {
    lastHighlighted.forEach(el => el.remove());
    lastHighlighted = [];
  }

  function convertCurrentRectHighlightsToRead() {
    if (!lastHighlighted.length) return;
    lastHighlighted.forEach(el => {
      if (!el || !el.classList) return;
      el.classList.remove('reading-rect-highlight');
      el.classList.add('reading-read-highlight');
    });
    // 已轉成「已讀高亮」後，從暫存清單移除，避免被 clearHighlights() 刪掉。
    lastHighlighted = [];
  }

  function drawRectHighlights(rects) {
    rects.forEach(r => {
      const el = document.createElement('div');
      el.className = 'reading-rect-highlight';
      const textLayerEl = r.pageNumber
        ? document.querySelector(`.page[data-page-number="${r.pageNumber}"] .textLayer`)
        : null;
      if (textLayerEl) {
        el.style.left = `${(r.relLeft || 0) * 100}%`;
        el.style.top = `${(r.relTop || 0) * 100}%`;
        el.style.width = `${Math.max(0.1, (r.relWidth || 0) * 100)}%`;
        el.style.height = `${Math.max(0.1, (r.relHeight || 0) * 100)}%`;
        textLayerEl.insertBefore(el, textLayerEl.firstChild);
      } else {
        el.style.left = `${r.left}px`;
        el.style.top = `${r.top}px`;
        el.style.width = `${Math.max(1, r.width)}px`;
        el.style.height = `${Math.max(1, r.height)}px`;
        document.body.appendChild(el);
      }
      lastHighlighted.push(el);
    });
  }

  async function readWithHighlight() {
    // 為這次朗讀建立唯一 token
    const myToken = ++currentReadToken;
    stopCurrentSpeechNow();
    const { text, rects } = collectTextAndSpansInFrame();
    let cleaned = (text || '').trim();
    clearHighlights();
    if (cleaned) {
      drawRectHighlights(rects);
    }
    const select = document.getElementById('voiceSelect');
    const voiceName = select ? select.value : 'local-zh-female';
    const ratePercent = getRatePercentFromUI();
    if (window.sendTextToTTS) {
      // 若框內沒有偵測到文字，改用圖片辨識
      if (!cleaned) {
        try {
          const imgText = await recognizeImageInFrame();
          cleaned = (imgText || '').trim();
        } catch (e) {
          cleaned = '';
        }
        if (!cleaned) {
          frameReading = false;
          const labelEl0 = document.getElementById('readingLabel');
          if (labelEl0) labelEl0.textContent = '框選朗讀';
          return;
        }
      }
      frameReading = true;
      const labelEl = document.getElementById('readingLabel');
      if (labelEl) labelEl.textContent = '取消朗讀';
      await window.sendTextToTTS(
        cleaned,
        () => {
          // 只有當這次朗讀仍是最新一次時，才清除高亮
          if (myToken === currentReadToken) {
            clearHighlights();
            frameReading = false;
            const labelEl2 = document.getElementById('readingLabel');
            if (labelEl2) labelEl2.textContent = '框選朗讀';
          }
        },
        {
          voiceName,
          ratePercent,
        }
      );
    }
  }

  function stopCurrentSpeechNow() {
    markBlockReadToken++;
    clearHighlights();
    try {
      if (window.speechSynthesis) {
        window.speechSynthesis.cancel();
      }
    } catch (e) {}
    try {
      const audio = document.getElementById('readingAudioPlayer');
      if (audio) {
        audio.pause();
        audio.currentTime = 0;
        audio.removeAttribute('src');
        try { audio.load(); } catch (e) {}
      }
    } catch (e) {}
    window.setTtsPlaybackState?.('idle');
  }

  if (typeof window !== 'undefined') {
    window.stopCurrentSpeechNow = stopCurrentSpeechNow;
  }

  function getPageGlyphs(pageNumber) {
    const pageLayer = document.querySelector(`.page[data-page-number="${pageNumber}"] .textLayer`);
    if (!pageLayer) return [];
    const glyphs = [];

    for (const el of pageLayer.querySelectorAll('span')) {
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let textNode;
      while ((textNode = walker.nextNode())) {
        const txt = textNode.nodeValue || '';
        if (!txt) continue;

        for (let i = 0; i < txt.length; i++) {
          const range = document.createRange();
          range.setStart(textNode, i);
          range.setEnd(textNode, i + 1);
          const rect = range.getBoundingClientRect();
          if (!rect || rect.width <= 0 || rect.height <= 0) continue;

          glyphs.push({
            ch: txt[i],
            x: rect.left,
            y: rect.top,
            w: rect.width,
            h: rect.height,
          });
        }
      }
    }

    // 依閱讀順序排序
    glyphs.sort((a, b) => {
      const dy = a.y - b.y;
      if (Math.abs(dy) > 2) return dy;
      return a.x - b.x;
    });

    // 粗體/陰影重疊去重：同字元且位置接近(容差 3px)只留一個
    const deduped = [];
    for (const g of glyphs) {
      const duplicated = deduped.some(d =>
        d.ch === g.ch &&
        Math.abs(d.x - g.x) <= 3 &&
        Math.abs(d.y - g.y) <= 3
      );
      if (!duplicated) deduped.push(g);
    }

    return deduped.map((g, idx) => ({ ...g, idx, pageNumber }));
  }

  function getCurrentPageGlyphs() {
    const pageNumber = window.PDFViewerApplication?.pdfViewer?.currentPageNumber ||
      parseInt(document.getElementById('pageNumber')?.value || '1', 10) || 1;
    return getPageGlyphs(pageNumber);
  }

  function buildRectMetaFromGlyph(g) {
    const textLayerEl = document.querySelector(`.page[data-page-number="${g.pageNumber}"] .textLayer`);
    const layerRect = textLayerEl?.getBoundingClientRect();
    if (!textLayerEl || !layerRect) return null;
    return {
      pageNumber: g.pageNumber,
      left: g.x,
      top: g.y,
      width: g.w,
      height: g.h,
      relLeft: (g.x - layerRect.left) / (layerRect.width || 1),
      relTop: (g.y - layerRect.top) / (layerRect.height || 1),
      relWidth: g.w / (layerRect.width || 1),
      relHeight: g.h / (layerRect.height || 1),
    };
  }

  function buildSentenceSegmentsFromGlyphs(glyphs) {
    const raw = glyphs.map(g => g.ch).join('');
    const regex = /[^。！？!?；;.;]+[。！？!?；;.]?/g;
    const segments = [];
    let match;
    while ((match = regex.exec(raw)) !== null) {
      const segmentRaw = match[0] || '';
      const text = segmentRaw.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
      if (!text) continue;

      const start = match.index;
      const end = start + segmentRaw.length;
      const rects = glyphs.slice(start, end)
        .filter(g => (g.ch || '').trim() !== '')
        .map(g => buildRectMetaFromGlyph(g))
        .filter(r => !!r);
      segments.push({ text, rects });
    }
    return segments;
  }

  async function speakSegment(text, voiceName, ratePercent, token) {
    return new Promise(resolve => {
      if (token !== fullPageReadToken) {
        resolve(false);
        return;
      }
      if (!window.sendTextToTTS) {
        resolve(false);
        return;
      }
      window.sendTextToTTS(text, () => {
        resolve(token === fullPageReadToken);
      }, {
        voiceName,
        ratePercent,
      });
    });
  }

  function updateFullPageButtonText() {
    const btn = document.getElementById('fullPageReadButton');
    if (!btn) return;
    const span = btn.querySelector('span');
    if (span) {
      span.textContent = fullPageReading ? '取消朗讀' : '整頁朗讀';
    }
    btn.title = fullPageReading ? '取消朗讀' : '整頁朗讀';

    const pauseBtn = document.getElementById('fullPagePauseButton');
    if (pauseBtn) {
      pauseBtn.style.display = fullPageReading ? '' : 'none';
      const pauseSpan = pauseBtn.querySelector('span');
      if (pauseSpan) {
        pauseSpan.textContent = fullPagePaused ? '繼續' : '暫停';
      }
      pauseBtn.title = fullPagePaused ? '繼續朗讀' : '暫停朗讀';
    }
  }

  function resolvePauseWaiters() {
    const waiters = pauseWaiters;
    pauseWaiters = [];
    waiters.forEach(fn => fn());
  }

  async function waitIfPaused(token) {
    while (fullPagePaused && token === fullPageReadToken) {
      await new Promise(resolve => pauseWaiters.push(resolve));
    }
  }

  function togglePauseResume() {
    if (!fullPageReading) return;
    fullPagePaused = !fullPagePaused;

    try {
      const audio = document.getElementById('readingAudioPlayer');
      if (audio && !audio.paused) {
        if (fullPagePaused) {
          audio.pause();
        }
      } else if (audio && audio.paused && !fullPagePaused) {
        audio.play().catch(() => {});
      }
    } catch (e) {}

    try {
      if (window.speechSynthesis) {
        if (fullPagePaused && typeof window.speechSynthesis.pause === 'function') {
          window.speechSynthesis.pause();
        } else if (!fullPagePaused && typeof window.speechSynthesis.resume === 'function') {
          window.speechSynthesis.resume();
        }
      }
    } catch (e) {}

    if (!fullPagePaused) {
      resolvePauseWaiters();
    }
    window.setTtsPlaybackState?.(fullPagePaused ? 'paused' : 'playing');
    updateFullPageButtonText();
  }

  function findGlyphIndexAtPoint(glyphs, x, y) {
    let bestIdx = -1;
    let bestDist = Number.POSITIVE_INFINITY;
    let bestArea = Number.POSITIVE_INFINITY;
    for (let i = 0; i < glyphs.length; i++) {
      const g = glyphs[i];
      const left = g.x, top = g.y, right = g.x + g.w, bottom = g.y + g.h;
      if (x >= left && x <= right && y >= top && y <= bottom) {
        const cx = left + g.w / 2;
        const cy = top + g.h / 2;
        const dist = Math.abs(cx - x) + Math.abs(cy - y);
        const area = Math.max(1, g.w * g.h);
        // 有重疊時，優先挑最近且框較小者，避免大框蓋住小字框導致難點選
        if (dist < bestDist || (Math.abs(dist - bestDist) < 0.5 && area < bestArea)) {
          bestDist = dist;
          bestArea = area;
          bestIdx = i;
        }
        continue;
      }
      const cx = left + g.w / 2;
      const cy = top + g.h / 2;
      const dist = Math.abs(cx - x) + Math.abs(cy - y);
      if (dist < bestDist) {
        bestDist = dist;
        bestIdx = i;
      }
    }
    return bestIdx;
  }

  function wordFromPoint(glyphs, x, y) {
    if (!glyphs.length) return null;
    const idx = findGlyphIndexAtPoint(glyphs, x, y);
    if (idx < 0) return null;

    const g = glyphs[idx];
    const text = (g.ch || '').trim();
    if (!text) return null;
    const rectMeta = buildRectMetaFromGlyph(g) || { left: g.x, top: g.y, width: g.w, height: g.h };
    return {
      text,
      rects: [rectMeta],
      idx: g.idx,
      pageNumber: g.pageNumber,
      key: `${text}|${Math.round(g.x)}|${Math.round(g.y)}|${Math.round(g.w)}|${Math.round(g.h)}`,
    };
  }

  // 以觸控點為基準，取得同一行、同一垂直觸控線上的字元群組
  function buildTouchGlyphGroup(glyphs, seedWord, touchX) {
    if (!seedWord) return [];
    const pageGlyphs = glyphs.filter(g => g.pageNumber === seedWord.pageNumber);
    const seed = pageGlyphs.find(g => g.idx === seedWord.idx);
    if (!seed) return [seedWord];

    const sameLineTolerance = 3;
    const touchXTolerance = 1.5; // 觸控線左右一點點的誤差範圍

    const lineGlyphs = pageGlyphs.filter(g => Math.abs(g.y - seed.y) <= sameLineTolerance);
    const group = lineGlyphs.filter(g => {
      const left = g.x;
      const right = g.x + g.w;
      return touchX >= left - touchXTolerance && touchX <= right + touchXTolerance;
    });

    group.sort((a, b) => a.idx - b.idx);
    return group.map(g => ({
      text: (g.ch || '').trim(),
      idx: g.idx,
      pageNumber: g.pageNumber,
      key: `${(g.ch || '').trim()}|${Math.round(g.x)}|${Math.round(g.y)}|${Math.round(g.w)}|${Math.round(g.h)}`,
      rects: [buildRectMetaFromGlyph(g) || { left: g.x, top: g.y, width: g.w, height: g.h }],
    })).filter(w => w.text);
  }

  function ensurePendingReadButton() {
    if (pendingReadButton) return pendingReadButton;
    const btn = document.createElement('button');
    btn.id = 'pendingReadButton';
    btn.textContent = '開始朗讀';
    btn.style.display = 'none';
    // 讓「開始朗讀 / 取消朗讀」按鈕在平板上更好點擊
    btn.style.position = 'fixed';
    btn.style.zIndex = '3200';
    btn.style.padding = '12px 26px';
    btn.style.fontSize = '22px';
    btn.style.fontWeight = '600';
    btn.style.borderRadius = '18px';
    btn.style.border = '1px solid rgba(0,0,0,0.3)';
    btn.style.backgroundColor = 'rgba(255,255,255,0.96)';
    btn.style.boxShadow = '0 2px 6px rgba(0,0,0,0.25)';
    btn.addEventListener('click', async () => {
      if (pendingSelectionReading) {
        pendingSelectionReading = false;
        touchReadToken++;
        stopCurrentSpeechNow();
        // 取消朗讀時，視為放棄本次選取：清除當前高亮與選取狀態
        clearHighlights();
        resetPendingSelection(true);
        updatePendingReadButtonText();
        return;
      }
      await speakPendingSelection();
    });
    document.body.appendChild(btn);
    pendingReadButton = btn;
    return btn;
  }

  function updatePendingReadButtonText() {
    if (!pendingReadButton) return;
    pendingReadButton.textContent = pendingSelectionReading ? '取消朗讀' : '開始朗讀';
    pendingReadButton.title = pendingSelectionReading ? '取消朗讀' : '開始朗讀';
  }

  function updatePendingReadButtonPosition() {
    const btn = ensurePendingReadButton();
    if (pendingGlyphs.length === 0) {
      btn.style.display = 'none';
      return;
    }
    const last = pendingGlyphs[pendingGlyphs.length - 1];
    const left = last.rects?.[0]?.left ?? last.x ?? 0;
    const top = last.rects?.[0]?.top ?? last.y ?? 0;
    btn.style.display = 'block';
    btn.style.left = `${Math.max(8, left)}px`;
    btn.style.top = `${Math.max(8, top - 52)}px`;
    updatePendingReadButtonText();
  }

  function resetPendingSelection(clearVisual = true) {
    pendingGlyphs = [];
    pendingGlyphKeySet.clear();
    if (clearVisual) {
      clearHighlights();
    }
    if (pendingReadButton) {
      pendingReadButton.style.display = 'none';
    }
  }

  function markReadHighlights(rects) {
    rects.forEach(r => {
      const el = document.createElement('div');
      el.className = 'reading-read-highlight';
      const textLayerEl = r.pageNumber
        ? document.querySelector(`.page[data-page-number="${r.pageNumber}"] .textLayer`)
        : null;
      if (textLayerEl) {
        el.style.left = `${(r.relLeft || 0) * 100}%`;
        el.style.top = `${(r.relTop || 0) * 100}%`;
        el.style.width = `${Math.max(0.1, (r.relWidth || 0) * 100)}%`;
        el.style.height = `${Math.max(0.1, (r.relHeight || 0) * 100)}%`;
        textLayerEl.insertBefore(el, textLayerEl.firstChild);
      } else {
        el.style.left = `${r.left}px`;
        el.style.top = `${r.top}px`;
        el.style.width = `${Math.max(1, r.width)}px`;
        el.style.height = `${Math.max(1, r.height)}px`;
        document.body.appendChild(el);
      }
    });
  }

  function clearReadHighlights() {
    document.querySelectorAll('.reading-read-highlight').forEach(el => el.remove());
  }

  /** 工具列「清除顏色標示」：朗讀中的暫存高亮 + 已讀黃／綠標 + 內部追蹤模型一併清掉 */
  function clearAllReadingColorMarkers() {
    clearHighlights();
    readHighlightModelByPage.clear();
    clearReadHighlights();
    document.querySelectorAll('.reading-rect-highlight, .reading-read-highlight').forEach(el => el.remove());
    lastHighlighted = [];
  }

  let ttsPlaybackState = 'idle'; // idle | playing | paused

  function setTtsPlaybackState(state) {
    ttsPlaybackState = state;
    const textEl = document.getElementById('ttsPlaybackStateText');
    const toggleBtn = document.getElementById('ttsToggleBtn');
    if (textEl) {
      if (state === 'playing') textEl.textContent = '播放中';
      else if (state === 'paused') textEl.textContent = '已暫停';
      else textEl.textContent = '已停止';
    }
    if (toggleBtn) {
      toggleBtn.textContent = state === 'paused' ? '繼續' : '暫停';
      toggleBtn.title = state === 'paused' ? '繼續播放' : '暫停播放';
    }
  }
  if (typeof window !== 'undefined') {
    window.setTtsPlaybackState = setTtsPlaybackState;
  }

  function layoutTtsPlaybackBar() {
    const bar = document.getElementById('ttsPlaybackBar');
    const voiceSelect = document.getElementById('voiceSelect');
    const speedWrap = document.getElementById('readingSpeedWrap');
    if (!bar || !voiceSelect || !speedWrap) return;

    const GAP_BAR_VOICE = 10; // 維持原本 bar 與發音選單間距
    const GAP_VOICE_SPEED = 24; // 視覺上接近目前配置
    const VOICE_SHIFT_X = 20; // 發音選單微調向右
    const TOP = 70;
    const PAD = 8;

    const bw = bar.offsetWidth || 250;
    const vw = voiceSelect.offsetWidth || 300;
    const sw = speedWrap.offsetWidth || 260;
    const groupW = bw + GAP_BAR_VOICE + vw + GAP_VOICE_SPEED + sw;

    let groupLeft = Math.round((window.innerWidth - groupW) / 2);
    groupLeft = Math.max(PAD, Math.min(groupLeft, Math.max(PAD, window.innerWidth - PAD - groupW)));

    const barLeft = groupLeft;
    const voiceLeft = barLeft + bw + GAP_BAR_VOICE + VOICE_SHIFT_X;
    const speedLeft = voiceLeft + vw + GAP_VOICE_SPEED;

    bar.style.top = `${TOP}px`;
    bar.style.left = `${Math.round(barLeft)}px`;
    bar.style.right = 'auto';

    voiceSelect.style.top = `${TOP}px`;
    voiceSelect.style.left = `${Math.round(voiceLeft)}px`;
    voiceSelect.style.right = 'auto';

    speedWrap.style.top = `${TOP}px`;
    speedWrap.style.left = `${Math.round(speedLeft)}px`;
    speedWrap.style.right = 'auto';
  }

  function ensureTtsPlaybackBar() {
    if (document.getElementById('ttsPlaybackBar')) return;
    const bar = document.createElement('div');
    bar.id = 'ttsPlaybackBar';
    bar.style.position = 'fixed';
    bar.style.zIndex = '3100';
    bar.style.display = 'flex';
    bar.style.alignItems = 'center';
    bar.style.gap = '8px';
    bar.style.padding = '6px 10px';
    bar.style.background = 'rgba(25,27,31,0.92)';
    bar.style.border = '1px solid rgba(255,255,255,0.22)';
    bar.style.borderRadius = '999px';
    bar.style.boxShadow = '0 2px 8px rgba(0,0,0,0.28)';
    bar.style.fontSize = '12px';
    bar.innerHTML = `
      <span id="ttsPlaybackStateText" style="min-width:48px;color:#fff;">已停止</span>
      <button id="ttsToggleBtn" type="button" style="padding:2px 10px;border-radius:999px;border:1px solid rgba(255,255,255,0.35);background:#2e7dff;color:#fff;">暫停</button>
      <button id="ttsStopBtn" type="button" style="padding:2px 10px;border-radius:999px;border:1px solid rgba(255,255,255,0.35);background:#4a4f58;color:#fff;">停止</button>
    `;
    document.body.appendChild(bar);

    document.getElementById('ttsToggleBtn')?.addEventListener('click', () => {
      const nextPaused = ttsPlaybackState !== 'paused';
      try {
        const audio = document.getElementById('readingAudioPlayer');
        if (audio) {
          if (nextPaused && !audio.paused) audio.pause();
          if (!nextPaused && audio.paused) audio.play().catch(() => {});
        }
      } catch (_) {
        /* ignore */
      }
      try {
        if (window.speechSynthesis) {
          if (nextPaused && typeof window.speechSynthesis.pause === 'function') {
            window.speechSynthesis.pause();
          } else if (!nextPaused && typeof window.speechSynthesis.resume === 'function') {
            window.speechSynthesis.resume();
          }
        }
      } catch (_) {
        /* ignore */
      }
      if (fullPageReading) {
        fullPagePaused = nextPaused;
        if (!nextPaused) resolvePauseWaiters();
        updateFullPageButtonText();
      }
      window.setTtsPlaybackState?.(nextPaused ? 'paused' : 'playing');
    });
    document.getElementById('ttsStopBtn')?.addEventListener('click', () => {
      fullPageReading = false;
      fullPagePaused = false;
      pendingSelectionReading = false;
      frameReading = false;
      fullPageReadToken++;
      touchReadToken++;
      markBlockReadToken++;
      resolvePauseWaiters();
      stopCurrentSpeechNow();
      updateFullPageButtonText();
      updatePendingReadButtonText();
    });
    requestAnimationFrame(layoutTtsPlaybackBar);
    window.addEventListener('resize', layoutTtsPlaybackBar);
  }

  if (typeof window !== 'undefined') {
    window.clearAllReadingColorMarkers = clearAllReadingColorMarkers;
  }

  function appendReadHighlightModel(glyphs) {
    for (const g of glyphs) {
      const page = g.pageNumber;
      if (!page) continue;
      const arr = readHighlightModelByPage.get(page) || [];
      if (!arr.some(x => x.idx === g.idx && x.ch === g.text)) {
        arr.push({ idx: g.idx, ch: g.text });
      }
      readHighlightModelByPage.set(page, arr);
    }
  }

  function rebuildReadHighlightsFromModel() {
    clearReadHighlights();
    for (const [pageNumber, models] of readHighlightModelByPage.entries()) {
      const glyphs = getPageGlyphs(pageNumber);
      if (!glyphs.length) continue;
      for (const m of models) {
        let g = glyphs[m.idx];
        if (!g || g.ch !== m.ch) {
          let found = null;
          for (let d = 1; d <= 4 && !found; d++) {
            if (glyphs[m.idx - d]?.ch === m.ch) found = glyphs[m.idx - d];
            if (glyphs[m.idx + d]?.ch === m.ch) found = glyphs[m.idx + d];
          }
          g = found;
        }
        if (!g) continue;
        const rect = buildRectMetaFromGlyph(g);
        if (rect) markReadHighlights([rect]);
      }
    }
  }

  function bindModelRepaintEvents() {
    const app = window.PDFViewerApplication;
    const eb = app?.eventBus;
    if (!eb || bindModelRepaintEvents._bound) return false;
    bindModelRepaintEvents._bound = true;
    const rerender = () => rebuildReadHighlightsFromModel();
    eb._on('pagerendered', rerender);
    eb._on('scalechanging', rerender);
    eb._on('rotationchanging', rerender);
    eb._on('updateviewarea', rerender);
    return true;
  }

  function sortGlyphsByReadingOrder(glyphList) {
    const arr = [...glyphList];
    arr.sort((a, b) => {
      const ay = a.rects?.[0]?.top ?? a.y ?? 0;
      const by = b.rects?.[0]?.top ?? b.y ?? 0;
      const ax = a.rects?.[0]?.left ?? a.x ?? 0;
      const bx = b.rects?.[0]?.left ?? b.x ?? 0;
      const dy = ay - by;
      // 同一行容差：3px
      if (Math.abs(dy) > 3) return dy;
      return ax - bx;
    });
    return arr;
  }

  function glyphIntersectsClientFrame(g, frame) {
    const gl = g.x;
    const gt = g.y;
    const gr = g.x + g.w;
    const gb = g.y + g.h;
    const il = Math.max(gl, frame.left);
    const ir = Math.min(gr, frame.right);
    const it = Math.max(gt, frame.top);
    const ib = Math.min(gb, frame.bottom);
    const iw = ir - il;
    const ih = ib - it;
    if (iw <= 0 || ih <= 0) return false;
    // 與標記文字抽取門檻一致：需覆蓋到字元框一定比例，避免只擦到下緣就被高亮。
    const overlapArea = iw * ih;
    const glyphArea = Math.max(1e-6, (gr - gl) * (gb - gt));
    return overlapArea / glyphArea >= 0.45;
  }

  function buildTextLayerRectFromClientFrame(pageNumber, frameRect) {
    const textLayerEl = document.querySelector(
      `.page[data-page-number="${pageNumber}"] .textLayer`
    );
    const layerRect = textLayerEl?.getBoundingClientRect();
    if (!textLayerEl || !layerRect) return null;
    const il = Math.max(frameRect.left, layerRect.left);
    const ir = Math.min(frameRect.right, layerRect.right);
    const it = Math.max(frameRect.top, layerRect.top);
    const ib = Math.min(frameRect.bottom, layerRect.bottom);
    if (ir <= il || ib <= it) return null;
    return {
      pageNumber,
      left: il,
      top: it,
      width: ir - il,
      height: ib - it,
      relLeft: (il - layerRect.left) / (layerRect.width || 1),
      relTop: (it - layerRect.top) / (layerRect.height || 1),
      relWidth: (ir - il) / (layerRect.width || 1),
      relHeight: (ib - it) / (layerRect.height || 1),
    };
  }

  /**
   * 標記區塊朗讀：與觸控選取朗讀相同，朗讀中為 .reading-rect-highlight，結束後改為 .reading-read-highlight。
   */
  async function speakMarkRegionWithHighlight(pageNumber, frameRect, text, options = {}) {
    bindModelRepaintEvents();
    touchReadToken++;
    currentReadToken++;
    stopCurrentSpeechNow();
    const token = markBlockReadToken;
    const cleaned = (text || '')
      .replace(/[\r\n]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (!cleaned) return;

    const glyphs = getPageGlyphs(pageNumber);
    const orderedGlyphs = [];
    for (const g of glyphs) {
      if (!(g.ch || '').trim()) continue;
      if (!glyphIntersectsClientFrame(g, frameRect)) continue;
      const rectMeta = buildRectMetaFromGlyph(g);
      if (!rectMeta) continue;
      orderedGlyphs.push({
        text: (g.ch || '').trim(),
        idx: g.idx,
        pageNumber: g.pageNumber,
        key: `mark|${g.pageNumber}|${g.idx}`,
        rects: [rectMeta],
      });
    }
    const sorted = sortGlyphsByReadingOrder(orderedGlyphs);
    let rects = sorted.flatMap(g => g.rects);
    if (rects.length === 0) {
      const fallback = buildTextLayerRectFromClientFrame(pageNumber, frameRect);
      if (fallback) rects = [fallback];
    }
    if (rects.length) drawRectHighlights(rects);

    const voiceName =
      options.voiceName ||
      (document.getElementById('voiceSelect')
        ? document.getElementById('voiceSelect').value
        : 'local-zh-female');
    const ratePercent =
      typeof options.ratePercent === 'number' ? options.ratePercent : getRatePercentFromUI();

    if (!window.sendTextToTTS) return;
    const onEndExtra = typeof options.onEnd === 'function' ? options.onEnd : null;
    await window.sendTextToTTS(
      cleaned,
      () => {
        if (token !== markBlockReadToken) return;
        convertCurrentRectHighlightsToRead();
        if (sorted.length > 0) {
          appendReadHighlightModel(sorted);
        }
        if (onEndExtra) {
          try {
            onEndExtra();
          } catch (e) {
            /* ignore */
          }
        }
      },
      { voiceName, ratePercent }
    );
  }

  window.speakMarkRegionWithHighlight = speakMarkRegionWithHighlight;

  async function speakPendingSelection() {
    if (!touchSelectEnabled || pendingGlyphs.length === 0) return;
    const token = ++touchReadToken;
    pendingSelectionReading = true;
    updatePendingReadButtonText();
    const orderedGlyphs = sortGlyphsByReadingOrder(pendingGlyphs);
    const text = orderedGlyphs.map(g => g.text).join('').trim();
    if (!text) {
      pendingSelectionReading = false;
      updatePendingReadButtonText();
      resetPendingSelection(true);
      return;
    }
    const rects = orderedGlyphs.flatMap(g => g.rects);
    stopCurrentSpeechNow();
    clearHighlights();
    drawRectHighlights(rects);

    const select = document.getElementById('voiceSelect');
    const voiceName = select ? select.value : 'local-zh-female';
    const ratePercent = getRatePercentFromUI();

    if (window.sendTextToTTS) {
      await window.sendTextToTTS(text, () => {
        if (token === touchReadToken) {
          appendReadHighlightModel(orderedGlyphs);
          rebuildReadHighlightsFromModel();
          clearHighlights();
          pendingSelectionReading = false;
          updatePendingReadButtonText();
          resetPendingSelection(false);
        }
      }, {
        voiceName,
        ratePercent,
      });
    }
  }

  function markWordAtPoint(clientX, clientY) {
    if (!touchSelectEnabled) return;
    const glyphs = getCurrentPageGlyphs();
    const word = wordFromPoint(glyphs, clientX, clientY);
    if (!word) return;
    const seedSelected = pendingGlyphKeySet.has(word.key);

    // 以觸控點那一個字為唯一單位：已選就只取消這一個字，未選就只選這一個字。
    if (seedSelected) {
      pendingGlyphKeySet.delete(word.key);
      pendingGlyphs = pendingGlyphs.filter(g => g.key !== word.key);
    } else {
      pendingGlyphKeySet.add(word.key);
      pendingGlyphs.push(word);
    }
    clearHighlights();
    drawRectHighlights(pendingGlyphs.flatMap(g => g.rects));
    updatePendingReadButtonPosition();
  }

  // 標籤：直接執行框選朗讀（高亮 + 發音）
  const label = frame.querySelector('#readingLabel');
  label.addEventListener('click', async () => {
    // 若已在朗讀中，改為「取消朗讀」
    if (frameReading) {
      frameReading = false;
      currentReadToken++;
      stopCurrentSpeechNow();
      clearHighlights();
      label.textContent = '框選朗讀';
      return;
    }
    await readWithHighlight();
  });

  // 進階設定（API KEY）相關事件綁定
  (function initApiKeySettings() {
    const btn = document.getElementById('apiKeySettingsButton');
    const modal = document.getElementById('apiKeyModal');
    const input = document.getElementById('apiKeyInput');
    const saveBtn = document.getElementById('saveApiKeyButton');
    const deleteBtn = document.getElementById('deleteApiKeyButton');
    const cancelBtn = document.getElementById('cancelApiKeyButton');
    const helpBtn = document.getElementById('apiKeyHelpButton');
    const apiKeyHelpPopup = document.getElementById('apiKeyHelpPopup');
    const apiKeyHelpCloseButton = document.getElementById('apiKeyHelpCloseButton');
    const apiKeyHelpImage = document.getElementById('apiKeyHelpImage');
    const apiKeyImageFullscreen = document.getElementById('apiKeyImageFullscreen');
    const apiKeyImageFullscreenClose = document.getElementById('apiKeyImageFullscreenClose');
    const scanQRButton = document.getElementById('scanQRButton');
    const qrScannerContainer = document.getElementById('qrScannerContainer');
    const qrVideo = document.getElementById('qrVideo');
    const qrCanvas = document.getElementById('qrCanvas');
    const stopScanButton = document.getElementById('stopScanButton');

    if (!btn || !modal || !input || !saveBtn || !deleteBtn || !cancelBtn) return;

    // 記住目前 IndexedDB 中已儲存的 API KEY（用於顯示「已設定」提示）
    let currentStoredApiKey = '';

    function maskApiKey(key) {
      if (!key) return '';
      if (key.length <= 8) return '已設定：********';
      const start = key.slice(0, 4);
      const end = key.slice(-4);
      return `已設定：${start}...${end}`;
    }

    const openModal = async () => {
      if (document.body.classList.contains('marking-mode-active')) {
        await ensureSwal();
        if (window.Swal) {
          window.Swal.fire({
            icon: 'info',
            title: '標記模式中',
            text: '請先離開「標記區塊」模式，再開啟進階設定。',
            backdrop: false,
            customClass: { popup: 'swal-high-z-index' },
          });
        } else {
          alert('請先離開「標記區塊」模式，再開啟進階設定。');
        }
        return;
      }
      const storedKey = await loadGeminiApiKey();
      currentStoredApiKey = storedKey || '';
      // 輸入時一律隱藏明碼，用星號顯示
      input.type = 'password';
      input.value = '';
      if (currentStoredApiKey) {
        // 欄位內以淡色 placeholder 顯示「已設定：xxxx...xxxx」
        input.placeholder = maskApiKey(currentStoredApiKey);
      } else {
        input.placeholder = '請輸入或掃描 QR Code';
      }
      modal.style.display = 'flex';
    };
    const closeModal = () => {
      modal.style.display = 'none';
    };

    btn.addEventListener('click', openModal);
    cancelBtn.addEventListener('click', closeModal);
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeModal();
    });
    saveBtn.addEventListener('click', async () => {
      const newKey = input.value.trim();
      const finalKey = newKey || currentStoredApiKey || '';
      await saveGeminiApiKey(finalKey);
      currentStoredApiKey = finalKey;
      alert('API KEY 已儲存。');
      closeModal();
    });
    deleteBtn.addEventListener('click', async () => {
      input.value = '';
      currentStoredApiKey = '';
      await saveGeminiApiKey('');
      alert('API KEY 已刪除。');
      closeModal();
    });

    // QR Code 掃描相關：比照語音報讀系統個人版.html
    let qrStream = null;
    let qrScanInterval = null;

    async function requestQRVideoStreamWithFallback() {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        throw new Error('BROWSER_UNSUPPORTED');
      }
      // iOS/Safari 偏好後鏡頭失敗時，需退回較寬鬆條件。
      const candidates = [
        { video: { facingMode: { exact: 'environment' } }, audio: false },
        { video: { facingMode: { ideal: 'environment' } }, audio: false },
        { video: true, audio: false },
      ];
      let lastError = null;
      for (const constraints of candidates) {
        try {
          return await navigator.mediaDevices.getUserMedia(constraints);
        } catch (err) {
          lastError = err;
        }
      }
      throw lastError || new Error('CAMERA_FAILED');
    }

    function stopQRScan() {
      try {
        if (qrScanInterval) {
          clearInterval(qrScanInterval);
          qrScanInterval = null;
        }
        if (qrStream) {
          qrStream.getTracks().forEach(track => track.stop());
          qrStream = null;
        }
      } catch (e) {
        console.error('停止 QR 掃描時發生錯誤:', e);
      }
      if (qrVideo) qrVideo.srcObject = null;
      if (qrScannerContainer) qrScannerContainer.style.display = 'none';
    }

    if (scanQRButton && qrScannerContainer && qrVideo && qrCanvas) {
      scanQRButton.addEventListener('click', async () => {
        try {
          if (!window.isSecureContext) {
            throw new Error('INSECURE_CONTEXT');
          }
          // 請求相機權限（含 iOS/Safari fallback）
          qrStream = await requestQRVideoStreamWithFallback();
          qrVideo.srcObject = qrStream;
          qrVideo.setAttribute('playsinline', 'true');
          qrVideo.setAttribute('autoplay', 'true');
          qrVideo.setAttribute('muted', 'true');
          qrVideo.muted = true;
          qrScannerContainer.style.display = 'block';
          await qrVideo.play();

          const videoWidth = qrVideo.videoWidth || qrVideo.clientWidth || 640;
          const videoHeight = qrVideo.videoHeight || qrVideo.clientHeight || 480;
          qrCanvas.width = videoWidth;
          qrCanvas.height = videoHeight;
          const ctx = qrCanvas.getContext('2d');

          qrScanInterval = setInterval(() => {
            if (qrVideo.readyState === qrVideo.HAVE_ENOUGH_DATA) {
              ctx.drawImage(qrVideo, 0, 0, videoWidth, videoHeight);
              const imageData = ctx.getImageData(0, 0, videoWidth, videoHeight);
              const code = window.jsQR
                ? window.jsQR(imageData.data, imageData.width, imageData.height)
                : null;
              if (code && code.data) {
                input.value = code.data;
                stopQRScan();
                if (window.Swal) {
                  window.Swal.fire({
                    icon: 'success',
                    title: '掃描成功',
                    text: 'QR Code 已讀取',
                    timer: 1500,
                    showConfirmButton: false,
                    backdrop: false,
                    customClass: { popup: 'swal-high-z-index' },
                  });
                } else {
                  alert('掃描成功，已填入 API KEY。');
                }
              }
            }
          }, 120);
        } catch (error) {
          console.error('無法訪問相機：', error);
          let message = '請允許瀏覽器訪問相機權限，或手動輸入 API KEY';
          if (error && error.message === 'INSECURE_CONTEXT') {
            message = 'iOS 需在安全連線(HTTPS 或 localhost)才能使用相機。';
          } else if (error && error.message === 'BROWSER_UNSUPPORTED') {
            message = '目前瀏覽器不支援相機存取，請改用 Safari/Chrome 最新版。';
          } else if (error && (error.name === 'NotAllowedError' || error.name === 'SecurityError')) {
            message = '尚未授權相機，請到瀏覽器網站設定允許相機權限後重試。';
          } else if (error && (error.name === 'NotFoundError' || error.name === 'OverconstrainedError')) {
            message = '找不到可用相機，已嘗試切換相機模式但仍失敗，請改用手動輸入。';
          }
          if (window.Swal) {
            window.Swal.fire({
              icon: 'error',
              title: '無法訪問相機',
              text: message,
              backdrop: false,
              customClass: { popup: 'swal-high-z-index' },
            });
          } else {
            alert(message);
          }
        }
      });
    }

    if (stopScanButton) {
      stopScanButton.addEventListener('click', stopQRScan);
    }

    // 二、PDF 快取：清除目前 pdfjs-viewer 的 IndexedDB 檔案（比照語音報讀系統個人版的「清除 PDF 快取」）
    const clearCacheButton = document.getElementById('clearCacheButton');
    if (clearCacheButton) {
      clearCacheButton.addEventListener('click', async () => {
        try {
          const db = await openViewerDBLocal();
          await new Promise((resolve, reject) => {
            const tx = db.transaction(VIEWER_IDB_STORE_NAME, 'readwrite');
            const store = tx.objectStore(VIEWER_IDB_STORE_NAME);
            store.delete(VIEWER_IDB_LAST_FILE_KEY);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
          });
          db.close();
          try {
            localStorage.removeItem('pdfjs.lastFile');
          } catch (_) {
            /* ignore */
          }
          try {
            const u = new URL(window.location.href);
            if (u.searchParams.has('file')) {
              u.searchParams.delete('file');
              window.history.replaceState({}, '', u.toString());
            }
          } catch (_) {
            /* ignore */
          }
          const pdfApp = window.PDFViewerApplication;
          if (pdfApp && typeof pdfApp.close === 'function') {
            await pdfApp.close();
          }
          await ensureSwal();
          if (window.Swal) {
            window.Swal.fire({
              icon: 'success',
              title: '已清除 PDF 快取',
              text: 'PDF 快取已清除。請使用「開啟檔案」選擇 PDF，無需重新整理頁面。',
              timer: 1500,
              showConfirmButton: false,
              backdrop: false,
              customClass: { popup: 'swal-high-z-index' },
            });
          } else {
            alert('PDF 快取已清除。請使用「開啟檔案」選擇 PDF，無需重新整理頁面。');
          }
        } catch (e) {
          console.error('清除 PDF 快取失敗：', e);
          await ensureSwal();
          if (window.Swal) {
            window.Swal.fire({
              icon: 'error',
              title: '清除快取失敗',
              text: '清除 PDF 快取時發生錯誤，請稍後再試。',
              backdrop: false,
              customClass: { popup: 'swal-high-z-index' },
            });
          } else {
            alert('清除 PDF 快取失敗，請稍後再試。');
          }
        }
      });
    }

    // 取得 API KEY 教學
    // API KEY 教學彈窗開關（照抄語音報讀系統個人版）
    if (helpBtn && apiKeyHelpPopup && apiKeyHelpCloseButton) {
      helpBtn.addEventListener('click', () => {
        apiKeyHelpPopup.style.display = 'flex';
      });

      apiKeyHelpCloseButton.addEventListener('click', () => {
        apiKeyHelpPopup.style.display = 'none';
      });

      apiKeyHelpPopup.addEventListener('click', (event) => {
        if (event.target === apiKeyHelpPopup) {
          apiKeyHelpPopup.style.display = 'none';
        }
      });
    }

    if (apiKeyHelpImage && apiKeyImageFullscreen && apiKeyImageFullscreenClose) {
      apiKeyHelpImage.addEventListener('click', () => {
        apiKeyImageFullscreen.style.display = 'flex';
      });

      apiKeyImageFullscreenClose.addEventListener('click', () => {
        apiKeyImageFullscreen.style.display = 'none';
      });

      apiKeyImageFullscreen.addEventListener('click', (event) => {
        if (event.target === apiKeyImageFullscreen) {
          apiKeyImageFullscreen.style.display = 'none';
        }
      });
    }
  })();

  const fullPageReadButton = document.getElementById('fullPageReadButton');
  const fullPagePauseButton = document.getElementById('fullPagePauseButton');
  const touchSelectToggle = document.getElementById('touchSelectToggle');
  fullPagePauseButton?.addEventListener('click', () => {
    togglePauseResume();
  });

  document.getElementById('clearReadHighlightsButton')?.addEventListener('click', () => {
    clearAllReadingColorMarkers();
  });

  touchSelectToggle?.addEventListener('change', () => {
    touchSelectEnabled = !!touchSelectToggle.checked;
    const htmlEl = document.documentElement;
    if (!touchSelectEnabled) {
      pendingSelectionReading = false;
      touchReadToken++;
      stopCurrentSpeechNow();
      resetPendingSelection(true);
      readHighlightModelByPage.clear();
      clearReadHighlights();
      htmlEl.classList.remove('touch-select-reading-active');
    } else {
      ensurePendingReadButton();
      bindModelRepaintEvents();
      htmlEl.classList.add('touch-select-reading-active');
      // 啟用觸控選取朗讀時，將框選朗讀框重設回預設位置，避免遮擋
      resetFrameToDefaultPosition();
    }
  });

  fullPageReadButton?.addEventListener('click', async () => {
    if (fullPageReading) {
      fullPageReading = false;
      fullPagePaused = false;
      fullPageReadToken++;
      resolvePauseWaiters();
      stopCurrentSpeechNow();
      clearHighlights();
      updateFullPageButtonText();
      return;
    }

    // 開始整頁朗讀時，將框選朗讀框重設回預設位置，避免與整頁朗讀視覺干擾
    resetFrameToDefaultPosition();

    const glyphs = getCurrentPageGlyphs();
    if (!glyphs.length) return;
    const segments = buildSentenceSegmentsFromGlyphs(glyphs);
    if (segments.length === 0) return;

    fullPageReading = true;
    fullPagePaused = false;
    const token = ++fullPageReadToken;
    updateFullPageButtonText();
    stopCurrentSpeechNow();
    clearHighlights();

    const select = document.getElementById('voiceSelect');
    const voiceName = select ? select.value : 'local-zh-female';
    const ratePercent = getRatePercentFromUI();

    for (const seg of segments) {
      if (token !== fullPageReadToken) break;
      await waitIfPaused(token);
      if (token !== fullPageReadToken) break;
      clearHighlights();
      if (seg.rects?.length) {
        drawRectHighlights(seg.rects);
      }
      const ok = await speakSegment(seg.text, voiceName, ratePercent, token);
      if (!ok) break;
    }

    if (token === fullPageReadToken) {
      fullPageReading = false;
      fullPagePaused = false;
      clearHighlights();
      updateFullPageButtonText();
    }
  });
  updateFullPageButtonText();

  // 觸控選取朗讀：啟用時，攔截 touch 事件避免 iPad 長按出現系統選字反白
  viewerContainer.addEventListener('touchstart', (e) => {
    if (!touchSelectEnabled) return;
    const t = e.touches && e.touches[0];
    if (!t) return;
    e.preventDefault(); // 阻止 iOS 觸發長按選字 / 選單
    markWordAtPoint(t.clientX, t.clientY);
  }, { passive: false });

  viewerContainer.addEventListener('touchmove', (e) => {
    if (!touchSelectEnabled) return;
    const t = e.touches && e.touches[0];
    if (!t) return;
    e.preventDefault(); // 避免移動中又被判定為選字拖曳
    markWordAtPoint(t.clientX, t.clientY);
  }, { passive: false });

  viewerContainer.addEventListener('touchend', (e) => {
    if (!touchSelectEnabled) return;
    e.preventDefault();
    updatePendingReadButtonPosition();
  }, { passive: false });

  // 滑鼠點選也啟用同樣機制
  viewerContainer.addEventListener('click', async (e) => {
    if (!touchSelectEnabled) return;
    markWordAtPoint(e.clientX, e.clientY);
  });

  viewerContainer.addEventListener('mousemove', (e) => {
    if (!touchSelectEnabled) return;
    if ((e.buttons & 1) !== 1) return;
    markWordAtPoint(e.clientX, e.clientY);
  });

  viewerContainer.addEventListener('mouseleave', () => {
    if (!touchSelectEnabled) return;
    updatePendingReadButtonPosition();
  });

  // 等 viewer 初始化後綁定「重算重繪」事件
  let bindRetry = 0;
  const bindTimer = setInterval(() => {
    bindRetry++;
    if (bindModelRepaintEvents() || bindRetry > 30) {
      clearInterval(bindTimer);
    }
  }, 300);

  // 右側 / 下方橘色按鈕：快速放大框
  const expandX = frame.querySelector('#readingExpandX');
  const expandY = frame.querySelector('#readingExpandY');

  // 橘色按鈕：按住 + 拖曳 來拉伸
  expandX.addEventListener('mousedown', (e) => {
    const rect = frame.getBoundingClientRect();
    startX = e.clientX;
    startY = e.clientY;
    startLeft = rect.left;
    startTop = rect.top;
    startWidth = rect.width;
    startHeight = rect.height;
    // 重新拉伸時也先清除前一次的高亮
    if (typeof clearHighlights === 'function') {
      clearHighlights();
    }
    dragging = false;
    resizing = true;
    resizeDir = 'x';
    e.stopPropagation();
    e.preventDefault();
  });

  expandY.addEventListener('mousedown', (e) => {
    const rect = frame.getBoundingClientRect();
    startX = e.clientX;
    startY = e.clientY;
    startLeft = rect.left;
    startTop = rect.top;
    startWidth = rect.width;
    startHeight = rect.height;
    // 重新拉伸時也先清除前一次的高亮
    if (typeof clearHighlights === 'function') {
      clearHighlights();
    }
    dragging = false;
    resizing = true;
    resizeDir = 'y';
    e.stopPropagation();
    e.preventDefault();
  });

  // 觸控拉伸
  expandX.addEventListener('touchstart', (e) => {
    const rect = frame.getBoundingClientRect();
    const point = e.touches[0];
    startX = point.clientX;
    startY = point.clientY;
    startLeft = rect.left;
    startTop = rect.top;
    startWidth = rect.width;
    startHeight = rect.height;
    if (typeof clearHighlights === 'function') {
      clearHighlights();
    }
    dragging = false;
    resizing = true;
    resizeDir = 'x';
    e.stopPropagation();
    e.preventDefault();
  }, { passive: false });

  expandY.addEventListener('touchstart', (e) => {
    const rect = frame.getBoundingClientRect();
    const point = e.touches[0];
    startX = point.clientX;
    startY = point.clientY;
    startLeft = rect.left;
    startTop = rect.top;
    startWidth = rect.width;
    startHeight = rect.height;
    if (typeof clearHighlights === 'function') {
      clearHighlights();
    }
    dragging = false;
    resizing = true;
    resizeDir = 'y';
    e.stopPropagation();
    e.preventDefault();
  }, { passive: false });
})();

// 簡化版 TTS：支援本機 speechSynthesis 與線上 Edge TTS（試算表替換須與語音報讀系統個人版一致）
(function () {
  const API_CONFIG = {
    BASE_URL:
      (typeof window !== 'undefined' &&
        window.PDF_VIEWER_CONFIG &&
        window.PDF_VIEWER_CONFIG.TTS_BASE_URL) ||
      'https://readtts-tts.hf.space',
  };

  let browserVoices = [];

  /** 與第一個 IIFE 內 getRatePercentFromUI 相同（第二個 IIFE 無法閉包到該函式，須自行複製） */
  function getRatePercentFromUILocal() {
    const speedInput = document.getElementById('speakSpeed');
    const v = speedInput ? parseFloat(speedInput.value || '0') : 0;
    return Math.max(-50, Math.min(50, Number.isFinite(v) ? v : 0));
  }

  let cachedGvizSheetId = null;
  let cachedGvizData = null;

  function initBrowserVoices() {
    if (!('speechSynthesis' in window)) return;
    browserVoices = window.speechSynthesis.getVoices() || [];
  }

  function findLocalVoice(langPrefix) {
    if (!browserVoices || !browserVoices.length) return null;
    const exact = browserVoices.find(v => v.lang && v.lang.toLowerCase().startsWith(langPrefix.toLowerCase()));
    return exact || browserVoices[0] || null;
  }

  function speakWithLocalTTS(text, ratePercent, langPrefix, callback) {
    if (!('speechSynthesis' in window)) {
      if (callback) callback();
      return;
    }
    try { window.speechSynthesis.cancel(); } catch (e) {}

    const voice = findLocalVoice(langPrefix);
    const utterance = new SpeechSynthesisUtterance(text);
    if (voice) utterance.voice = voice;

    if (langPrefix === 'zh') {
      utterance.lang = (voice && voice.lang) ? voice.lang : 'zh-TW';
    } else if (langPrefix === 'en') {
      utterance.lang = (voice && voice.lang) ? voice.lang : 'en-US';
    } else {
      utterance.lang = (voice && voice.lang) ? voice.lang : langPrefix;
    }

    // 本機語速：負值用較緩斜率（-30% 約略慢於正常），正值仍可明顯加速
    let rate;
    if (ratePercent <= 0) {
      rate = 1 + (ratePercent / 100) * 0.35; // -30% → 約 0.895、-50% → 約 0.825
    } else {
      rate = 1 + (ratePercent / 100) * 1.2; // +50% → 約 1.6
    }
    rate = Math.min(3, Math.max(0.45, rate));
    utterance.rate = rate;
    utterance.pitch = 1.0;
    utterance.volume = 1.0;

    utterance.onend = () => { if (callback) callback(); };
    utterance.onerror = () => { if (callback) callback(); };

    try {
      if (typeof window.speechSynthesis.resume === 'function') {
        window.speechSynthesis.resume();
      }
      window.speechSynthesis.speak(utterance);
    } catch (e) {
      if (callback) callback();
    }
  }

  function getSheetIdFromUrl() {
    const urlParams = new URLSearchParams(window.location.search);
    const id =
      urlParams.get('googlesheetid') ||
      urlParams.get('GoogleSheetId') ||
      urlParams.get('googleSheetId') ||
      '';
    return id.trim();
  }

  function cellToString(cell) {
    if (cell == null) return '';
    if (cell.v != null && cell.v !== '') return cell.v.toString();
    if (cell.f != null && cell.f !== '') return cell.f.toString();
    return '';
  }

  function fetchSheetData(sheetId) {
    const base = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?`;
    const query = encodeURIComponent('Select B, C format B "", C ""');
    const url = `${base}&tq=${query}`;
    return fetch(url)
      .then(res => res.text())
      .then(rep => {
        const m = rep.match(/google\.visualization\.Query\.setResponse\(([\s\S]+)\);\s*$/);
        if (!m) {
          console.warn('試算表：無法解析 gviz 回應（非預期格式或權限／未發布）');
          return null;
        }
        const data = JSON.parse(m[1]);
        if (data.status === 'error') {
          console.error('試算表 gviz 錯誤:', data.errors);
          return null;
        }
        console.log('確認欄位對應（第一行標題，將被跳過）:', data.table?.rows?.[0]?.c);
        return data;
      })
      .catch(err => {
        console.error('載入失敗:', err);
        return null;
      });
  }

  async function fetchSheetDataCached(sheetId) {
    if (!sheetId) return null;
    if (cachedGvizSheetId === sheetId && cachedGvizData) {
      return cachedGvizData;
    }
    const data = await fetchSheetData(sheetId);
    cachedGvizSheetId = sheetId;
    cachedGvizData = data;
    return data;
  }

  function replaceTextWithSheetData(text, data) {
    if (!data || !data.table || !data.table.rows) return text;
    const dataRows = data.table.rows;
    console.log('總行數（API 已自動跳過標題行）:', dataRows.length);
    dataRows.forEach((row, index) => {
      const original = cellToString(row.c?.[0]);
      const replacement = cellToString(row.c?.[1]);
      if (original && replacement) {
        console.log(
          `處理第 ${index + 2} 行（試算表第 ${index + 2} 行）: "${original}" -> "${replacement}"`
        );
        text = text.replace(new RegExp(original.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), replacement);
      }
    });
    return text;
  }

  /**
   * 阿拉伯數字外包圍（①、➀、⑴、⓪…）：還原成半形數字；㊀、㈠ 等還原成一、二…十。
   * 凡屬「去外圈／括號」還原出的數字或中文數字，前後加短停頓（全形逗號「，」）以利 TTS。
   * 須先替換再 NFKC：否則 NFKC 會把 ①→1、㊀→一，永遠偵測不到帶圈碼位而無法加停頓。
   * 「○」+ 數字（PDF 拆字）同樣在 NFKC 之前處理。
   */
  function normalizeCircledDigitsForTts(s) {
    if (!s) return s;
    let out = s;

    const fw = '０１２３４５６７８９';

    function ttsPauseAround(inner) {
      if (inner == null || inner === '') return inner;
      return `，${inner}，`;
    }

    function swapOne(codePoint, num) {
      const ch = String.fromCodePoint(codePoint);
      if (out.includes(ch)) {
        out = out.split(ch).join(ttsPauseAround(String(num)));
      }
    }

    function swapRange(startCode, count, firstNum) {
      for (let i = 0; i < count; i++) {
        swapOne(startCode + i, firstNum + i);
      }
    }

    function swapToLiteral(codePoint, literal) {
      const ch = String.fromCodePoint(codePoint);
      if (out.includes(ch)) {
        out = out.split(ch).join(ttsPauseAround(literal));
      }
    }

    // Enclosed alphanumerics：①–⑳、⓪、⓿、⑪–⑳（負圈）
    swapRange(0x2460, 20, 1);
    swapOne(0x24ea, 0);
    swapOne(0x24ff, 0);
    swapRange(0x24eb, 10, 11);

    // 兩重圓圈 1–9、10
    swapRange(0x24f5, 9, 1);
    swapOne(0x24fe, 10);

    // 括號數字 ⑴–⑽、⑾–⑳
    swapRange(0x2474, 10, 1);
    swapRange(0x247e, 10, 11);

    // 括號 21–35
    swapRange(0x3251, 15, 21);

    // Dingbats：❶–❿、➀–➈、➊–➓
    swapRange(0x2776, 10, 1);
    swapRange(0x2780, 10, 1);
    swapRange(0x278a, 10, 1);

    // ㈠–㈩、㊀–㊉：Unicode 語意為「括號／圓圈」包中文數字，僅還原為一、二…十
    const cjkNumZh = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
    for (let i = 0; i < 10; i++) {
      swapToLiteral(0x3220 + i, cjkNumZh[i]);
      swapToLiteral(0x3280 + i, cjkNumZh[i]);
    }

    // 「○ / 白圓」+ 半形或全形數字（PDF 拆字後常見）
    out = out.replace(/○\s*([0-9０-９])/g, (_, d) => {
      const n = fw.indexOf(d);
      const num = n >= 0 ? n : parseInt(d, 10);
      return Number.isFinite(num) ? ttsPauseAround(String(num)) : d;
    });

    try {
      out = out.normalize('NFKC');
    } catch (e) {
      /* ignore */
    }

    out = out.replace(/，{2,}/g, '，');

    return out;
  }

  if (typeof window !== 'undefined') {
    window.normalizeCircledDigitsForTts = normalizeCircledDigitsForTts;
  }

  window.sendTextToTTS = async function (selectedText, callback, options) {
    try {
      if (window.speechSynthesis) {
        window.speechSynthesis.cancel();
      }
      const audioPlayer = document.getElementById('readingAudioPlayer');
      if (audioPlayer) {
        audioPlayer.pause();
        audioPlayer.currentTime = 0;
        audioPlayer.oncanplay = null;
        audioPlayer.onended = null;
        audioPlayer.onerror = null;
        audioPlayer.removeAttribute('src');
        try {
          audioPlayer.load();
        } catch (e) {
          /* ignore */
        }
      }
    } catch (e) {
      console.warn('重置發音狀態:', e);
    }

    let cleanedText = (selectedText || '')
      .replace(/[\r\n]+/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/[<>&]/g, '')
      .trim();

    if (!cleanedText) {
      if (callback) callback();
      return;
    }

    cleanedText = normalizeCircledDigitsForTts(cleanedText);

    let filteredText = cleanedText;
    const sheetId = getSheetIdFromUrl();
    if (sheetId) {
      const data = await fetchSheetDataCached(sheetId);
      filteredText = replaceTextWithSheetData(cleanedText, data);
    }

    console.log('送出給 TTS 的文本:', filteredText);
    console.log('文本長度:', filteredText.length);

    if (!filteredText || !filteredText.length) {
      if (callback) callback();
      return;
    }

    const voiceSelectEl = document.getElementById('voiceSelect');
    const voiceName =
      (options && options.voiceName) ||
      (voiceSelectEl ? voiceSelectEl.value : 'local-zh-female');
    const ratePercent =
      options && typeof options.ratePercent === 'number'
        ? options.ratePercent
        : getRatePercentFromUILocal();

    // 本機語音
    if (voiceName === 'local-zh-female' || voiceName === 'local-en-female') {
      const langPrefix = voiceName === 'local-zh-female' ? 'zh' : 'en';
      window.setTtsPlaybackState?.('playing');
      speakWithLocalTTS(filteredText, ratePercent, langPrefix, () => {
        window.setTtsPlaybackState?.('idle');
        if (callback) callback();
      });
      return;
    }

    // 線上 Edge TTS
    const rateString = `${ratePercent >= 0 ? '+' : ''}${ratePercent}%`;
    const requestData = {
      text: filteredText,
      voice: voiceName,
      rate: rateString,
      volume: '+0%',
      pitch: '+0Hz',
    };

    try {
      const response = await fetch(API_CONFIG.BASE_URL + '/tts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestData),
      });
      if (!response.ok) throw new Error('TTS request failed');
      const data = await response.json();
      if (!data || !data.success || !data.audio_url) {
        window.setTtsPlaybackState?.('idle');
        if (callback) callback();
        return;
      }
      const audioUrl = API_CONFIG.BASE_URL + data.audio_url;
      let audio = document.getElementById('readingAudioPlayer');
      if (!audio) {
        audio = document.createElement('audio');
        audio.id = 'readingAudioPlayer';
        audio.style.position = 'fixed';
        audio.style.left = '50%';
        audio.style.bottom = '8px';
        audio.style.transform = 'translateX(-50%)';
        document.body.appendChild(audio);
      }
      audio.src = audioUrl;
      audio.onplay = () => window.setTtsPlaybackState?.('playing');
      audio.onpause = () => {
        if (!audio.ended && audio.currentTime > 0) window.setTtsPlaybackState?.('paused');
      };
      audio.onended = () => {
        window.setTtsPlaybackState?.('idle');
        if (callback) callback();
      };
      audio.onerror = () => {
        window.setTtsPlaybackState?.('idle');
        if (callback) callback();
      };
      window.setTtsPlaybackState?.('playing');
      await audio.play();
    } catch (e) {
      window.setTtsPlaybackState?.('idle');
      if (callback) callback();
    }
  };

  if ('speechSynthesis' in window) {
    window.speechSynthesis.onvoiceschanged = initBrowserVoices;
    initBrowserVoices();
    setTimeout(initBrowserVoices, 100);
    setTimeout(initBrowserVoices, 500);
    setTimeout(initBrowserVoices, 1000);
  }
})();

// 載入後立即顯示框與橘色「框選朗讀」標籤，方便使用者看到與拖曳
window.addEventListener('load', () => {
  const frame = document.getElementById('readingFrame');
  if (!frame) return;

  const app = window.PDFViewerApplication;
  const eb = app?.eventBus;

  /**
   * PDF.js：若 #viewerContainer 在第一次 update() 時 clientHeight 仍為 0，
   * getVisiblePages() 會全空並直接 return，畫面白底直到使用者捲動觸發 watchScroll → _scrollUpdate。
   * 因此在 pagesloaded / ResizeObserver / 延遲後重複補踢。
   */
  const kickPdfViewerVisibleRefresh = (opts = {}) => {
    const { withSyntheticPointer = false } = opts;
    const a = window.PDFViewerApplication;
    if (!a?.pdfDocument || !a.pdfViewer) return;
    const pv = a.pdfViewer;
    const el = pv.container;
    const pn = Math.max(1, pv.currentPageNumber || 1);
    try {
      pv.scrollPageIntoView({ pageNumber: pn });
    } catch (_) {
      /* ignore */
    }
    try {
      if (el) {
        const prevTop = el.scrollTop;
        const maxTop = Math.max(0, el.scrollHeight - el.clientHeight);
        if (maxTop > 0) {
          el.scrollTop = Math.min(prevTop + 1, maxTop);
          el.scrollTop = prevTop;
        } else {
          const prevLeft = el.scrollLeft;
          const maxLeft = Math.max(0, el.scrollWidth - el.clientWidth);
          if (maxLeft > 0) {
            el.scrollLeft = Math.min(prevLeft + 1, maxLeft);
            el.scrollLeft = prevLeft;
          }
        }
      }
    } catch (_) {
      /* ignore */
    }
    try {
      pv.focus();
    } catch (_) {
      /* ignore */
    }
    // 與內部 scroll 監聽相同入口（不依賴實際捲動事件）
    try {
      if (typeof pv._scrollUpdate === 'function') {
        pv._scrollUpdate();
      }
    } catch (_) {
      /* ignore */
    }
    try {
      pv.update();
    } catch (_) {
      /* ignore */
    }
    try {
      a.forceRendering();
    } catch (_) {
      /* ignore */
    }
    try {
      window.dispatchEvent(new Event('resize'));
    } catch (_) {
      /* ignore */
    }
    if (withSyntheticPointer && el) {
      try {
        const pageEl = el.querySelector('.page[data-page-number="' + pn + '"]') || el.querySelector('.page');
        const target = pageEl || el;
        const r = target.getBoundingClientRect();
        if (r.width >= 4 && r.height >= 4) {
          const x = Math.round(r.left + r.width / 2);
          const y = Math.round(r.top + r.height / 2);
          const base = {
            bubbles: true,
            cancelable: true,
            clientX: x,
            clientY: y,
            screenX: x,
            screenY: y,
            view: window,
          };
          target.dispatchEvent(new MouseEvent('mousedown', base));
          target.dispatchEvent(new MouseEvent('mouseup', base));
        }
      } catch (_) {
        /* ignore */
      }
    }
  };

  let layoutKickRaf = 0;
  const scheduleLayoutKick = (withPointer = false) => {
    if (layoutKickRaf) cancelAnimationFrame(layoutKickRaf);
    layoutKickRaf = requestAnimationFrame(() => {
      layoutKickRaf = 0;
      kickPdfViewerVisibleRefresh({ withSyntheticPointer: withPointer });
      requestAnimationFrame(() => kickPdfViewerVisibleRefresh({ withSyntheticPointer: withPointer }));
    });
  };

  if (eb && typeof eb._on === 'function') {
    eb._on('pagesloaded', () => {
      scheduleLayoutKick(false);
      [80, 200, 450].forEach(ms => setTimeout(() => scheduleLayoutKick(false), ms));
    });
    eb._on('documentloaded', () => {
      requestAnimationFrame(() => scheduleLayoutKick(false));
    });
  }

  // window load 若晚於 PDF 開啟，events 已錯過，補一次
  if (app?.pdfDocument && app.pdfViewer?.pagesCount > 0) {
    scheduleLayoutKick(false);
    [60, 180, 400].forEach(ms => setTimeout(() => scheduleLayoutKick(false), ms));
  }

  const viewerBox = document.getElementById('viewerContainer');
  if (viewerBox && typeof ResizeObserver !== 'undefined') {
    let lastArea = 0;
    let roT = 0;
    const ro = new ResizeObserver(() => {
      if (!window.PDFViewerApplication?.pdfDocument) return;
      const area = viewerBox.clientWidth * viewerBox.clientHeight;
      if (area < 4) return;
      if (area === lastArea) return;
      lastArea = area;
      if (roT) clearTimeout(roT);
      roT = setTimeout(() => {
        roT = 0;
        scheduleLayoutKick(false);
      }, 32);
    });
    ro.observe(viewerBox);
  }

  /**
   * 首次開啟與重新載入共用同一套「可互動」就緒條件，避免網路串流只先出 canvas、流程卻與 F5 後不一致。
   * 條件：pagesloaded + 目前頁 textlayerrendered + 該頁 canvas 已出現 + 最短停留（與後續 kick 銜接）。
   * 遮罩僅在 PDF 就緒後關閉：不設幾秒強制關閉，且 window.load 時若 eventBus 尚未建立則輪詢綁定（避免 600ms 誤關）。
   */
  const INITIAL_COVER_MIN_MS = 520;
  const COVER_BIND_MAX_WAIT_MS = 45 * 60 * 1000;
  const COVER_NO_FILE_PARAM_HIDE_MS = 650;

  const initialCover = document.getElementById('initialRenderCover');
  if (initialCover) {
    const hideCover = () => {
      if (!initialCover.isConnected) return;
      initialCover.removeAttribute('aria-busy');
      initialCover.style.pointerEvents = 'none';
      initialCover.style.transition = 'opacity 0.28s ease-out';
      initialCover.style.opacity = '0';
      setTimeout(() => {
        initialCover.remove();
        kickPdfViewerVisibleRefresh({ withSyntheticPointer: true });
        requestAnimationFrame(() => kickPdfViewerVisibleRefresh({ withSyntheticPointer: true }));
        setTimeout(() => scheduleLayoutKick(true), 80);
      }, 320);
    };
    const scheduleHideCover = () => {
      kickPdfViewerVisibleRefresh();
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          kickPdfViewerVisibleRefresh();
          setTimeout(() => {
            kickPdfViewerVisibleRefresh();
            hideCover();
          }, 140);
        });
      });
    };

    const coverBindT0 = Date.now();
    let hasFileParam = false;
    try {
      const u = new URL(window.location.href);
      hasFileParam = !!(u.searchParams.get('file') || '').trim();
    } catch (_) {
      hasFileParam = false;
    }
    if (!hasFileParam) {
      // 沒有 ?file= 代表不需等待遠端檔案流程，避免空白首頁長時間停留在載入遮罩。
      setTimeout(() => {
        if (initialCover.dataset.pdfjsCoverBound === '1') return;
        initialCover.dataset.pdfjsCoverBound = '1';
        scheduleHideCover();
      }, COVER_NO_FILE_PARAM_HIDE_MS);
      return;
    }
    const bindInitialCoverWhenReady = () => {
      if (initialCover.dataset.pdfjsCoverBound === '1') {
        return;
      }
      const ebNow = window.PDFViewerApplication?.eventBus;
      if (!ebNow || typeof ebNow._on !== 'function') {
        // 沒有 ?file= 時代表非遠端 bootstrap 場景，不要長時間停留在載入動畫。
        if (!hasFileParam && Date.now() - coverBindT0 > COVER_NO_FILE_PARAM_HIDE_MS) {
          initialCover.dataset.pdfjsCoverBound = '1';
          scheduleHideCover();
          return;
        }
        if (Date.now() - coverBindT0 > COVER_BIND_MAX_WAIT_MS) {
          initialCover.dataset.pdfjsCoverBound = '1';
          scheduleHideCover();
          return;
        }
        setTimeout(bindInitialCoverWhenReady, 100);
        return;
      }

      initialCover.dataset.pdfjsCoverBound = '1';

      let hidden = false;
      let textLayerFallbackTid = 0;
      const coverWait = {
        pagesLoaded: false,
        textLayerOk: false,
        firstCanvasAt: 0,
      };
      const syncPagesLoaded = () => {
        if (coverWait.pagesLoaded) return;
        try {
          const a = window.PDFViewerApplication;
          const num = a?.pdfDocument?.numPages;
          const cnt = a?.pdfViewer?.pagesCount;
          if (num && cnt && cnt >= num) coverWait.pagesLoaded = true;
        } catch (_) {}
      };
      const tryRevealCover = () => {
        if (hidden) return;
        syncPagesLoaded();
        if (!coverWait.pagesLoaded || !coverWait.textLayerOk || !coverWait.firstCanvasAt) return;
        const elapsed = Date.now() - coverWait.firstCanvasAt;
        if (elapsed < INITIAL_COVER_MIN_MS) {
          setTimeout(tryRevealCover, INITIAL_COVER_MIN_MS - elapsed + 20);
          return;
        }
        hidden = true;
        scheduleHideCover();
      };

      ebNow._on('pagesloaded', () => {
        coverWait.pagesLoaded = true;
        tryRevealCover();
      });
      ebNow._on('textlayerrendered', evt => {
        if (hidden) return;
        if (evt?.error) return;
        const a = window.PDFViewerApplication;
        const want = a?.pdfViewer?.currentPageNumber;
        if (want != null && evt?.pageNumber != null && evt.pageNumber === want) {
          if (textLayerFallbackTid) {
            clearTimeout(textLayerFallbackTid);
            textLayerFallbackTid = 0;
          }
          coverWait.textLayerOk = true;
          tryRevealCover();
        }
      });
      syncPagesLoaded();
      requestAnimationFrame(() => {
        syncPagesLoaded();
        tryRevealCover();
      });
      ebNow._on('pagerendered', evt => {
        if (hidden) return;
        if (evt?.cssTransform) return;
        const a = window.PDFViewerApplication;
        const viewer = a?.pdfViewer;
        const want = viewer?.currentPageNumber;
        if (want != null && evt && evt.pageNumber != null && evt.pageNumber !== want) {
          return;
        }
        if (!coverWait.firstCanvasAt) coverWait.firstCanvasAt = Date.now();
        if (textLayerFallbackTid) clearTimeout(textLayerFallbackTid);
        textLayerFallbackTid = setTimeout(() => {
          textLayerFallbackTid = 0;
          if (!hidden && !coverWait.textLayerOk) {
            coverWait.textLayerOk = true;
            tryRevealCover();
          }
        }, 880);
        tryRevealCover();
      });
    };

    bindInitialCoverWhenReady();

    if (app?.pdfDocument) {
      requestAnimationFrame(() => kickPdfViewerVisibleRefresh());
    }
  }
  // 延後一點點時間，確保工具列與按鈕尺寸已經穩定，再套用與觸控朗讀相同的預設位置與大小
  setTimeout(() => {
    if (typeof window.resetReadingFrameDefault === 'function') {
      window.resetReadingFrameDefault();
    } else {
      frame.style.display = 'block';
    }
  }, 50);
});

// 右上「工具」摺疊選單：開啟時自動貼齊視窗，若超出則往左(或往右)平移，確保可完整點擊。
window.addEventListener('load', () => {
  const toggleBtn = document.getElementById('secondaryToolbarToggleButton');
  const menu = document.getElementById('secondaryToolbar');
  if (!toggleBtn || !menu) return;

  function clampSecondaryToolbarIntoViewport() {
    if (menu.classList.contains('hidden')) return;
    const pad = 6;
    menu.style.transform = 'translateX(0)';
    const rect = menu.getBoundingClientRect();
    let shiftX = 0;
    if (rect.right > window.innerWidth - pad) {
      shiftX = rect.right - (window.innerWidth - pad);
    } else if (rect.left < pad) {
      shiftX = rect.left - pad;
    }
    if (shiftX !== 0) {
      menu.style.transform = `translateX(${-shiftX}px)`;
    }
  }

  toggleBtn.addEventListener('click', () => {
    requestAnimationFrame(() => clampSecondaryToolbarIntoViewport());
  });
  window.addEventListener('resize', () => clampSecondaryToolbarIntoViewport());
});

// 工具列分區由 PDF.js 內建 JS 控制；這裡補一層「窄螢幕動態配置」，
// 保證四顆自訂按鈕優先保留，不靠單純 CSS media query。
window.addEventListener('load', () => {
  const toolbarLeft = document.getElementById('toolbarViewerLeft');
  if (!toolbarLeft) return;

  const keepIds = [
    'openFileTopLeft',
    'fullPageReadButton',
    'touchSelectToggleLabel',
    'clearReadHighlightsButton',
  ];
  const candidateGetters = [
    () => document.getElementById('sidebarToggleButton'),
    () => document.getElementById('previous')?.closest('.toolbarHorizontalGroup'),
    () => document.getElementById('pageNumber')?.closest('.toolbarHorizontalGroup'),
  ];

  function setShown(el, shown) {
    if (!el) return;
    el.style.display = shown ? '' : 'none';
  }

  function relayoutTopToolbar() {
    // 搜尋與縮放已移到摺疊次選單：主工具列固定隱藏。
    const findWrap = document.getElementById('viewFindButton')?.closest('.toolbarButtonWithContainer');
    const toolbarMiddle = document.getElementById('toolbarViewerMiddle');
    setShown(findWrap, false);
    setShown(toolbarMiddle, false);

    // 先重置：候選項全部顯示，再視空間逐步收斂。
    candidateGetters.forEach(getEl => setShown(getEl(), true));
    toolbarLeft.querySelectorAll('.toolbarButtonSpacer').forEach(sp => {
      sp.style.width = '';
    });

    keepIds.forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      el.style.display = '';
      el.style.flex = '0 0 auto';
      el.style.minWidth = 'max-content';
      el.style.whiteSpace = 'nowrap';
    });

    // 若左區塊爆寬，先縮 spacer。
    if (toolbarLeft.scrollWidth > toolbarLeft.clientWidth + 1) {
      toolbarLeft.querySelectorAll('.toolbarButtonSpacer').forEach(sp => {
        sp.style.width = '8px';
      });
    }

    // 仍爆寬就依序收掉次要功能，保留四顆自訂按鈕。
    for (const getEl of candidateGetters) {
      if (toolbarLeft.scrollWidth <= toolbarLeft.clientWidth + 1) break;
      setShown(getEl(), false);
    }
  }

  relayoutTopToolbar();
  window.addEventListener('resize', relayoutTopToolbar);
});

// 將主工具列的「搜尋、縮放」收納到右上摺疊次選單。
window.addEventListener('load', () => {
  const secondaryContainer = document.getElementById('secondaryToolbarButtonContainer');
  if (!secondaryContainer) return;
  if (document.getElementById('secondaryFindTrigger')) return;

  const findWrap = document.getElementById('viewFindButton')?.closest('.toolbarButtonWithContainer');
  const toolbarMiddle = document.getElementById('toolbarViewerMiddle');
  if (findWrap) findWrap.style.display = 'none';
  if (toolbarMiddle) toolbarMiddle.style.display = 'none';

  const findBtn = document.getElementById('viewFindButton');
  const findInput = document.getElementById('findInput');
  const findPrevBtn = document.getElementById('findPreviousButton');
  const findNextBtn = document.getElementById('findNextButton');
  const zoomOutBtn = document.getElementById('zoomOutButton');
  const zoomInBtn = document.getElementById('zoomInButton');
  const scaleSelect = document.getElementById('scaleSelect');

  const separator = document.createElement('div');
  separator.className = 'horizontalToolbarSeparator';

  const mkBtn = (id, text, onClick) => {
    const btn = document.createElement('button');
    btn.id = id;
    btn.type = 'button';
    btn.className = 'toolbarButton labeled';
    btn.textContent = text;
    btn.addEventListener('click', onClick);
    return btn;
  };

  const secondaryFind = mkBtn('secondaryFindTrigger', '搜尋面板', () => {
    findBtn?.click();
  });
  const secondaryZoomOut = mkBtn('secondaryZoomOutTrigger', '縮小', () => {
    zoomOutBtn?.click();
  });
  const secondaryZoomIn = mkBtn('secondaryZoomInTrigger', '放大', () => {
    zoomInBtn?.click();
  });

  const searchWrap = document.createElement('div');
  searchWrap.id = 'secondarySearchWrap';
  searchWrap.style.padding = '6px 12px';
  searchWrap.style.display = 'flex';
  searchWrap.style.alignItems = 'center';
  searchWrap.style.gap = '6px';
  const secondarySearchInput = document.createElement('input');
  secondarySearchInput.id = 'secondaryFindInput';
  secondarySearchInput.className = 'toolbarField';
  secondarySearchInput.type = 'text';
  secondarySearchInput.placeholder = '搜尋文字';
  secondarySearchInput.style.flex = '1 1 auto';
  secondarySearchInput.style.minWidth = '120px';
  secondarySearchInput.style.height = '32px';
  const secondaryPrev = document.createElement('button');
  secondaryPrev.id = 'secondaryFindPrevTrigger';
  secondaryPrev.type = 'button';
  secondaryPrev.textContent = '▲';
  secondaryPrev.addEventListener('click', () => {
    findPrevBtn?.click();
  });
  const secondaryNext = document.createElement('button');
  secondaryNext.id = 'secondaryFindNextTrigger';
  secondaryNext.type = 'button';
  secondaryNext.textContent = '▼';
  secondaryNext.addEventListener('click', () => {
    findNextBtn?.click();
  });
  secondaryPrev.style.minWidth = '0';
  secondaryNext.style.minWidth = '0';
  secondaryPrev.style.padding = '0 4px';
  secondaryNext.style.padding = '0 4px';
  secondaryPrev.style.border = 'none';
  secondaryNext.style.border = 'none';
  secondaryPrev.style.background = 'transparent';
  secondaryNext.style.background = 'transparent';
  secondaryPrev.style.fontSize = '16px';
  secondaryNext.style.fontSize = '16px';
  secondaryPrev.style.lineHeight = '1';
  secondaryNext.style.lineHeight = '1';

  if (findInput) {
    secondarySearchInput.value = findInput.value || '';
    secondarySearchInput.addEventListener('input', () => {
      findInput.value = secondarySearchInput.value;
      findInput.dispatchEvent(new Event('input', { bubbles: true }));
    });
    secondarySearchInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        findNextBtn?.click();
      }
    });
    findInput.addEventListener('input', () => {
      if (secondarySearchInput.value !== findInput.value) {
        secondarySearchInput.value = findInput.value || '';
      }
    });
  }
  searchWrap.appendChild(secondarySearchInput);
  searchWrap.appendChild(secondaryPrev);
  searchWrap.appendChild(secondaryNext);

  const scaleWrap = document.createElement('div');
  scaleWrap.id = 'secondaryScaleWrap';
  scaleWrap.style.padding = '6px 12px';
  scaleWrap.style.display = 'flex';
  scaleWrap.style.alignItems = 'center';
  scaleWrap.style.gap = '8px';
  const scaleLabel = document.createElement('span');
  scaleLabel.textContent = '縮放';
  scaleLabel.style.minWidth = '3em';
  const secondaryScale = document.createElement('select');
  secondaryScale.id = 'secondaryScaleSelect';
  secondaryScale.className = 'toolbarField';
  secondaryScale.style.width = '120px';
  if (scaleSelect) {
    secondaryScale.innerHTML = scaleSelect.innerHTML;
    secondaryScale.value = scaleSelect.value;
    secondaryScale.addEventListener('change', () => {
      scaleSelect.value = secondaryScale.value;
      scaleSelect.dispatchEvent(new Event('change', { bubbles: true }));
    });
    scaleSelect.addEventListener('change', () => {
      secondaryScale.value = scaleSelect.value;
    });
  }
  scaleWrap.appendChild(scaleLabel);
  scaleWrap.appendChild(secondaryScale);

  secondaryContainer.appendChild(separator);
  secondaryContainer.appendChild(secondaryFind);
  secondaryContainer.appendChild(searchWrap);
  secondaryContainer.appendChild(secondaryZoomOut);
  secondaryContainer.appendChild(secondaryZoomIn);
  secondaryContainer.appendChild(scaleWrap);
});

