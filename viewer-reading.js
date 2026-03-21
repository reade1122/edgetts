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
    voiceSelect.style.top = '70px';
    voiceSelect.style.right = '12px';
    voiceSelect.style.left = 'auto';
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
    speedWrap.style.zIndex = '3100';
    // 實際位置由 layoutTtsPlaybackBar() 置於底部置中（與播放列、audio 同列）
    speedWrap.style.top = 'auto';
    speedWrap.style.bottom = '12px';
    speedWrap.style.left = '50%';
    speedWrap.style.right = 'auto';
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

  /** 字元框與框選／標記區重疊比例（相對於字元面積），標記區慣用 0.45、框選朗讀用 0.55 */
  function isRectHitFlexible(frameRect, rect, minRatio = 0.55) {
    const interLeft = Math.max(rect.left, frameRect.left);
    const interRight = Math.min(rect.right, frameRect.right);
    const interTop = Math.max(rect.top, frameRect.top);
    const interBottom = Math.min(rect.bottom, frameRect.bottom);
    const interWidth = interRight - interLeft;
    const interHeight = interBottom - interTop;
    if (interWidth <= 0 || interHeight <= 0) return false;

    const rectArea = (rect.width || 1) * (rect.height || 1);
    const overlapArea = interWidth * interHeight;
    return overlapArea / rectArea >= minRatio;
  }

  function isRectHit(frameRect, rect) {
    return isRectHitFlexible(frameRect, rect, 0.55);
  }

  /** 直書字元通常「高明顯大於寬」；橫書則相反或接近方形 */
  function isVerticalishSize(width, height) {
    const w = Math.max(0, Number(width) || 0);
    const h = Math.max(0, Number(height) || 0);
    return h > w * 1.2;
  }

  function isCjkLikeChar(ch) {
    if (ch == null || ch === '') return false;
    const cp = ch.codePointAt(0);
    return (
      (cp >= 0x3000 && cp <= 0x9fff) ||
      (cp >= 0x3040 && cp <= 0x30ff) ||
      (cp >= 0x3400 && cp <= 0x4dbf) ||
      (cp >= 0xff10 && cp <= 0xff19) ||
      (cp >= 0x2460 && cp <= 0x2473)
    );
  }

  /**
   * PDF.js 直排時，Range 量到的常是「很寬、很扁」的錯誤外框；維持中心點對調寬高以接近真實直書字元。
   */
  function normalizeCjkGlyphClientRect(rect, ch) {
    const w = rect.width || 0;
    const h = rect.height || 0;
    if (w <= 0 || h <= 0) {
      return {
        left: rect.left,
        top: rect.top,
        width: w,
        height: h,
        right: rect.right,
        bottom: rect.bottom,
      };
    }
    if (!isCjkLikeChar(ch)) {
      return {
        left: rect.left,
        top: rect.top,
        width: w,
        height: h,
        right: rect.right,
        bottom: rect.bottom,
      };
    }
    if (!isVerticalishSize(w, h) && isVerticalishSize(h, w)) {
      const cx = rect.left + w * 0.5;
      const cy = rect.top + h * 0.5;
      const nw = h;
      const nh = w;
      const left = cx - nw * 0.5;
      const top = cy - nh * 0.5;
      return {
        left,
        top,
        width: nw,
        height: nh,
        right: left + nw,
        bottom: top + nh,
      };
    }
    return {
      left: rect.left,
      top: rect.top,
      width: w,
      height: h,
      right: rect.right,
      bottom: rect.bottom,
    };
  }

  /** span 為一般橫向排版（非 PDF 直排旋轉／直排 writing-mode） */
  function spanLooksLikeHorizontalLayout(span) {
    if (!span) return false;
    const tStyle = getComputedStyle(span).transform;
    const rotDeg = rotationDegFromComputedTransform(tStyle);
    let a = ((rotDeg % 360) + 360) % 360;
    if (a > 180) a -= 360;
    const near90 = Math.abs(Math.abs(a) - 90) < 42;
    if (near90) return false;
    return !spanLooksVerticallyTypeset(span);
  }

  /**
   * 橫書下 Range 常同時回傳「整行殘影」的細長豎條與正常字元框；若優先選豎條，顿號「、」等會變超高窄條，排序／高亮錯亂。
   */
  function pickReasonableGlyphRectFromList(list, ch, span) {
    if (!list.length) return null;
    const horizontal = spanLooksLikeHorizontalLayout(span);
    // 橫書含英文／數字／帶圈碼：同樣排除「細長豎條」量測殘影（如 LED 的 D 下垂）
    if (horizontal) {
      const notSkinnyStick = list.filter(r => {
        const w = r.width || 0;
        const h = r.height || 0;
        if (w <= 0 || h <= 0) return false;
        return h <= w * 2.85;
      });
      const pool = notSkinnyStick.length ? notSkinnyStick : list;
      pool.sort((a, b) => a.width * a.height - b.width * b.height);
      return pool[0];
    }
    if (isCjkLikeChar(ch)) {
      const vert = list.filter(r => isVerticalishSize(r.width, r.height));
      const pool = vert.length ? vert : list;
      pool.sort((a, b) => a.width * a.height - b.width * b.height);
      return pool[0];
    }
    const sorted = [...list].sort((a, b) => a.width * a.height - b.width * b.height);
    return sorted[0];
  }

  /**
   * 同一字元可能回傳多個 client rect；取面積最小且（若可）符合直書比例者，避免整欄寬的扁條當成高亮。
   * @param {Element} [span] 若為橫書 span，勿優先選細長豎條。
   */
  function pickGlyphRectsFromRange(range, ch, span) {
    const raw = Array.from(range.getClientRects()).filter(r => r.width > 0 && r.height > 0);
    const list = raw.length ? raw.map(r => normalizeCjkGlyphClientRect(r, ch)) : [];
    if (!list.length) {
      const br = range.getBoundingClientRect();
      if (!br || br.width <= 0 || br.height <= 0) return [];
      const one = normalizeCjkGlyphClientRect(br, ch);
      const picked = pickReasonableGlyphRectFromList([one], ch, span);
      return picked ? [picked] : [];
    }
    const picked = pickReasonableGlyphRectFromList(list, ch, span);
    return picked ? [picked] : [];
  }

  /** 從 getComputedStyle(span).transform 取得主要旋轉角（度），供 PDF.js 直書 rotate(90deg) 偵測 */
  function rotationDegFromComputedTransform(transformStr) {
    if (!transformStr || transformStr === 'none') return 0;
    const rot = transformStr.match(/rotate\(\s*(-?[\d.]+)\s*deg\s*\)/);
    if (rot) return parseFloat(rot[1]);
    const m = transformStr.match(/matrix\(([^)]+)\)/);
    if (m) {
      const p = m[1].split(/[\s,]+/).map(parseFloat);
      if (p.length >= 4 && p.every(Number.isFinite)) {
        return (Math.atan2(p[1], p[0]) * 180) / Math.PI;
      }
    }
    return 0;
  }

  /** 從 matrix(a,b,c,d,…) 取得旋轉角（弧度），含 rotate+scaleX 合成矩陣 */
  function rotationRadFromTransformMatrix(transformStr) {
    if (!transformStr || transformStr === 'none') return 0;
    const m = transformStr.match(/matrix\(([^)]+)\)/);
    if (!m) return (rotationDegFromComputedTransform(transformStr) * Math.PI) / 180;
    const p = m[1].split(/[\s,]+/).map(parseFloat);
    if (p.length < 4 || !p.every(Number.isFinite)) return 0;
    return Math.atan2(p[1], p[0]);
  }

  /**
   * PDF.js 直書：常為 rotate(90deg)+scaleX，瀏覽器合成後仍是 matrix；用弧度判斷「明顯非水平字排」。
   */
  function transformLooksLikePdfVerticalSpan(span) {
    if (!span) return false;
    const t = getComputedStyle(span).transform;
    const rad = Math.abs(rotationRadFromTransformMatrix(t));
    const deg = (rad * 180) / Math.PI;
    if (deg > 90) return Math.abs(deg - 180) > 25;
    return deg > 25 && deg < 155;
  }

  /** span 上 getComputedStyle 的 writing-mode 為直排／側排（與 PDF transform 分開，部分檔案只靠樣式標直橫） */
  function computedWritingModeIsVertical(span) {
    if (!span) return false;
    try {
      const wm = (getComputedStyle(span).writingMode || '').toLowerCase();
      return wm.includes('vertical') || wm.includes('sideways');
    } catch (e) {
      return false;
    }
  }

  /** 文字層中的明確換行／分段（橫書多行段落常見；直書欄內較少出現在單一寬 span 內） */
  function spanTextContainsExplicitLineBreak(raw) {
    if (!raw || typeof raw !== 'string') return false;
    return /[\n\r\u2028\u2029]/.test(raw);
  }

  /**
   * 該 span 是否為直書排版語意：旋轉矩陣、writing-mode 直排、或直排用 text-combine-upright。
   * 供頁面掃描、標記取字、spanIsVertical 與 transform 偵測一致。
   */
  function spanLooksVerticallyTypeset(span) {
    if (!span) return false;
    if (transformLooksLikePdfVerticalSpan(span)) return true;
    if (computedWritingModeIsVertical(span)) return true;
    try {
      const tc = (getComputedStyle(span).textCombineUpright || '').toLowerCase();
      if (tc && tc !== 'none' && tc !== 'initial' && tc !== 'inherit') {
        return true;
      }
    } catch (e) {
      /* ignore */
    }
    return false;
  }

  /** CJK + 已判定為直書矩陣：若外框仍「橫寬>縱高」則再對調一次（scaleX 後 AABB 仍可能很扁） */
  function coerceTallNarrowCjkRect(rect, ch) {
    if (!rect || !isCjkLikeChar(ch)) return rect;
    let r = normalizeCjkGlyphClientRect(rect, ch);
    const w = r.width;
    const h = r.height;
    if (w <= 0 || h <= 0) return r;
    if (w > h * 1.06) {
      const cx = r.left + w * 0.5;
      const cy = r.top + h * 0.5;
      const nw = h;
      const nh = w;
      const left = cx - nw * 0.5;
      const top = cy - nh * 0.5;
      return {
        left,
        top,
        width: nw,
        height: nh,
        right: left + nw,
        bottom: top + nh,
      };
    }
    return r;
  }

  /**
   * 橫書：限制單字高亮外框勿成「貫穿整行的細長條」，並與 span 行高對齊（顿號、注音旁字常見）。
   */
  function clampHorizontalGlyphBBox(rect, span) {
    if (!rect || !span) return rect;
    const sr = span.getBoundingClientRect();
    const sh = sr.height;
    if (sh <= 1) return rect;
    const w = rect.width;
    const h = rect.height;
    if (w <= 0 || h <= 0) return rect;
    const cx = rect.left + w / 2;
    const cy = rect.top + h / 2;
    const lineCap = Math.min(sh * 1.12, sh + 6);
    let nw = w;
    let nh = h;
    if (nh > nw * 2.1) {
      nh = Math.min(nh, lineCap, Math.max(nw * 1.55, sh * 0.94));
      nw = Math.max(nw, nh * 0.36);
    }
    if (nw > nh * 2.75) {
      nw = Math.min(nw, Math.max(nh * 2.05, sh * 0.52));
      nh = Math.min(Math.max(nh, sh * 0.76), lineCap);
    }
    const left = cx - nw / 2;
    const top = cy - nh / 2;
    return {
      left,
      top,
      width: nw,
      height: nh,
      right: left + nw,
      bottom: top + nh,
    };
  }

  /**
   * PDF.js 直書：span 常帶 rotate(±90deg)+scaleX；Range 量到整段 span 的軸對齊外框（一條橫扁帶）。
   * - 多字元且單字矩形覆蓋 span 大半面積：沿 span **較長邊**均分（不再依賴 near90）。
   * - 單字元且為直書矩陣 span：一律以 span 外框為基準再 coerce 成窄高。
   * - 其餘：維持 normalize + 橫扁時改採 span 外框。
   */
  function refineGlyphRectWithSpan(span, candidateRect, charIndex, charsInNode, ch) {
    if (!candidateRect || !span) return candidateRect;
    const spanRect = span.getBoundingClientRect();
    const sArea = Math.max(1e-6, spanRect.width * spanRect.height);
    const cArea = Math.max(1, candidateRect.width * candidateRect.height);
    const normCand = normalizeCjkGlyphClientRect(candidateRect, ch);
    const tStyle = getComputedStyle(span).transform;
    const rotDeg = rotationDegFromComputedTransform(tStyle);
    let a = ((rotDeg % 360) + 360) % 360;
    if (a > 180) a -= 360;
    const near90 = Math.abs(Math.abs(a) - 90) < 42;
    const verticalSpan = near90 || spanLooksVerticallyTypeset(span);
    // 橫書 span 內多字元（含英文如 LED）亦做橫向均分，避免單字 Range 量到跨行殘影
    const wholeSpanHit =
      charsInNode > 1 && cArea >= sArea * 0.48 && (!verticalSpan || isCjkLikeChar(ch));

    if (wholeSpanHit) {
      const sr0 = normalizeCjkGlyphClientRect(spanRect, ch);
      const w0 = sr0.width;
      const h0 = sr0.height;
      const l0 = sr0.left;
      const t0 = sr0.top;
      if (h0 >= w0 * 1.04) {
        const slice = h0 / charsInNode;
        const top = a < 0 ? t0 + (charsInNode - 1 - charIndex) * slice : t0 + charIndex * slice;
        return coerceTallNarrowCjkRect(
          {
            left: l0,
            top,
            width: w0,
            height: slice,
            right: l0 + w0,
            bottom: top + slice,
          },
          ch
        );
      }
      if (w0 >= h0 * 1.04) {
        const slice = w0 / charsInNode;
        const left = a < 0 ? l0 + (charsInNode - 1 - charIndex) * slice : l0 + charIndex * slice;
        const sliced = coerceTallNarrowCjkRect(
          {
            left,
            top: t0,
            width: slice,
            height: h0,
            right: left + slice,
            bottom: t0 + h0,
          },
          ch
        );
        return verticalSpan ? sliced : clampHorizontalGlyphBBox(sliced, span);
      }
    }

    if (charsInNode === 1 && isCjkLikeChar(ch) && verticalSpan) {
      return coerceTallNarrowCjkRect(spanRect, ch);
    }

    if (charsInNode === 1 && isCjkLikeChar(ch)) {
      const flat = normCand.width > normCand.height * 1.12;
      if (flat && verticalSpan) {
        return coerceTallNarrowCjkRect(spanRect, ch);
      }
      if (flat && !verticalSpan) {
        return clampHorizontalGlyphBBox(normCand, span);
      }
    }

    if (verticalSpan && isCjkLikeChar(ch)) {
      return coerceTallNarrowCjkRect(normCand, ch);
    }

    if (!verticalSpan) {
      return clampHorizontalGlyphBBox(normCand, span);
    }

    return normCand;
  }

  /** 頁級直書推斷結果快取（縮放／旋轉／重繪後清除） */
  const verticalReadingModeByPage = new Map();

  /**
   * 頁面閱讀版面：vertical / horizontal / mixed（直橫並存，取字時逐 span／字元混合排序）
   * @typedef {{ mode: 'vertical'|'horizontal'|'mixed', vRatio: number, hRatio: number, pageVertical: boolean, newlineWideRatio?: number, geometryVerticalHint?: boolean|null }} PageReadingLayoutMeta
   */
  const pageReadingLayoutMetaByPage = new Map();

  const PAGE_LAYOUT_FALLBACK = {
    mode: 'horizontal',
    vRatio: 0,
    hRatio: 0,
    pageVertical: false,
    newlineWideRatio: 0,
    geometryVerticalHint: null,
  };

  function clearVerticalReadingModeCacheForPage(pageNumber) {
    if (pageNumber != null && pageNumber !== '') {
      const pn = Number(pageNumber);
      verticalReadingModeByPage.delete(pn);
      pageReadingLayoutMetaByPage.delete(pn);
    } else {
      verticalReadingModeByPage.clear();
      pageReadingLayoutMetaByPage.clear();
    }
  }

  /**
   * 與專案根目錄 **index.htm**（PDF.js 2.11 `renderTextLayer`）相同的直書幾何判斷：
   * 多數 span 外框「寬度 &lt; 高度×0.8」且各 span **x 中心**的標準差相對「平均字寬」不大 → 視為直書欄。
   *
   * **為何 web/index.html 以前較差**：全功能 viewer（PDF.js 4.x）常把直排字用 `transform: matrix(…)` 畫成
   * **橫向包圍盒**，`width ≥ height` 的 span 變多，僅統計「直排 transform」比例會低估直書；index.htm 則易得到窄高 span。
   * 此函式把舊頁有效的幾何訊號併入，與 transform／writing-mode 互補。
   *
   * @returns {boolean|null} true／false 有把握；null 樣本不足或不明
   */
  function geometrySuggestsVerticalTextLayerLegacyHeuristic(textLayer) {
    if (!textLayer) return null;
    const cfg = typeof window !== 'undefined' ? window.PDF_VIEWER_CONFIG || {} : {};
    const minSpans =
      typeof cfg.READING_GEOMETRY_VERTICAL_MIN_SPANS === 'number'
        ? cfg.READING_GEOMETRY_VERTICAL_MIN_SPANS
        : 10;
    const narrowTh =
      typeof cfg.READING_GEOMETRY_VERTICAL_NARROW_RATIO === 'number'
        ? cfg.READING_GEOMETRY_VERTICAL_NARROW_RATIO
        : 0.6;
    const xVarMult =
      typeof cfg.READING_GEOMETRY_VERTICAL_XVAR_AVGW_MULT === 'number'
        ? cfg.READING_GEOMETRY_VERTICAL_XVAR_AVGW_MULT
        : 3;

    const tl = textLayer.getBoundingClientRect();
    if (tl.width < 1 || tl.height < 1) return null;

    const spanData = [];
    for (const span of textLayer.querySelectorAll('span')) {
      const t = (span.textContent || '').replace(/\s/g, '');
      if (!t) continue;
      const rect = span.getBoundingClientRect();
      const w = rect.width;
      const h = rect.height;
      if (w <= 0.5 || h <= 0.5) continue;
      const centerX = rect.left + w * 0.5 - tl.left;
      spanData.push({ w, h, centerX });
      if (spanData.length >= 160) break;
    }

    if (spanData.length < minSpans) return null;

    let narrow = 0;
    let avgX = 0;
    let avgW = 0;
    const n = spanData.length;
    for (const d of spanData) {
      avgX += d.centerX;
      avgW += d.w;
      if (d.w < d.h * 0.8) narrow++;
    }
    avgX /= n;
    avgW /= n;

    let xVarSum = 0;
    for (const d of spanData) {
      const dx = d.centerX - avgX;
      xVarSum += dx * dx;
    }
    const xVar = Math.sqrt(xVarSum / n);
    const narrowRatio = narrow / n;

    if (narrowRatio > narrowTh && xVar < avgW * xVarMult) {
      return true;
    }
    if (narrowRatio < 0.22 && xVar > avgW * 4.5) {
      return false;
    }
    return null;
  }

  /**
   * 掃描單頁文字層，快取直／橫／混合與比例（供載入預掃、朗讀取字、標記一致使用）。
   */
  function getPageReadingLayoutMeta(pageNumber) {
    const pn = Number(pageNumber);
    if (!pn || !Number.isFinite(pn)) {
      return { ...PAGE_LAYOUT_FALLBACK };
    }
    if (pageReadingLayoutMetaByPage.has(pn)) {
      return pageReadingLayoutMetaByPage.get(pn);
    }
    const textLayer = document.querySelector(`.page[data-page-number="${pn}"] .textLayer`);
    if (!textLayer) {
      const emptyMeta = { ...PAGE_LAYOUT_FALLBACK };
      pageReadingLayoutMetaByPage.set(pn, emptyMeta);
      verticalReadingModeByPage.set(pn, false);
      return emptyMeta;
    }
    const cfg = typeof window !== 'undefined' ? window.PDF_VIEWER_CONFIG || {} : {};
    const vMin =
      typeof cfg.READING_VERTICAL_MIN_TRANSFORM_RATIO === 'number'
        ? cfg.READING_VERTICAL_MIN_TRANSFORM_RATIO
        : 0.38;
    const hMax =
      typeof cfg.READING_HORIZONTAL_SPAN_MAX_RATIO === 'number'
        ? cfg.READING_HORIZONTAL_SPAN_MAX_RATIO
        : 0.44;
    const mixV =
      typeof cfg.READING_MIXED_MIN_V_RATIO === 'number' ? cfg.READING_MIXED_MIN_V_RATIO : 0.16;
    const mixH =
      typeof cfg.READING_MIXED_MIN_H_RATIO === 'number' ? cfg.READING_MIXED_MIN_H_RATIO : 0.18;
    const newlineWideBoost =
      typeof cfg.READING_NEWLINE_WIDE_HORIZ_BOOST === 'number'
        ? cfg.READING_NEWLINE_WIDE_HORIZ_BOOST
        : 0.42;

    let n = 0;
    let vertTransforms = 0;
    let horizBoxSpans = 0;
    let newlineWideCount = 0;
    for (const span of textLayer.querySelectorAll('span')) {
      const rawFull = span.textContent || '';
      const t = rawFull.replace(/\s/g, '');
      if (!t) continue;
      n++;
      const vt = spanLooksVerticallyTypeset(span);
      if (vt) {
        vertTransforms++;
      } else {
        const sr = span.getBoundingClientRect();
        const wide = sr.width > 0 && sr.height > 0 && sr.width >= sr.height * 0.88;
        if (wide) {
          horizBoxSpans++;
          if (spanTextContainsExplicitLineBreak(rawFull)) {
            horizBoxSpans += newlineWideBoost;
            newlineWideCount++;
          }
        }
      }
      if (n >= 120) break;
    }
    const vRatio = n ? vertTransforms / n : 0;
    const hRatio = n ? Math.min(1.05, horizBoxSpans / n) : 0;

    let pageVertical = vRatio >= vMin && hRatio <= hMax && vRatio > hRatio;
    if (!pageVertical && vRatio >= 0.26 && hRatio <= 0.18) {
      pageVertical = true;
    }
    if (!pageVertical) {
      try {
        const wm = (getComputedStyle(textLayer).writingMode || '').toLowerCase();
        if (wm.includes('vertical') || wm.includes('sideways')) pageVertical = true;
      } catch (e) {
        /* ignore */
      }
    }

    let geometryVerticalHint = null;
    if (cfg.READING_GEOMETRY_VERTICAL_ENABLED !== false) {
      geometryVerticalHint = geometrySuggestsVerticalTextLayerLegacyHeuristic(textLayer);
      if (geometryVerticalHint === true) {
        pageVertical = true;
      } else if (geometryVerticalHint === false && vRatio < 0.22) {
        pageVertical = false;
      }
    }

    let mode = pageVertical ? 'vertical' : 'horizontal';
    if (vRatio >= mixV && hRatio >= mixH) {
      mode = 'mixed';
    } else if (pageVertical && hRatio >= 0.32) {
      mode = 'mixed';
    } else if (!pageVertical && vRatio >= 0.24) {
      mode = 'mixed';
    }

    const newlineWideRatio = n ? newlineWideCount / n : 0;
    const meta = {
      mode,
      vRatio,
      hRatio,
      pageVertical,
      newlineWideRatio,
      geometryVerticalHint,
    };
    pageReadingLayoutMetaByPage.set(pn, meta);
    verticalReadingModeByPage.set(pn, pageVertical);
    return meta;
  }

  /**
   * 依文字層推斷該頁是否為直書（與 edgetts 對照：edgetts 無頁級直書，橫書永遠 y→x）。
   * 舊版僅用「直排 span ≥14%」易把橫書頁（標題／注音旋轉）誤判為整頁直書，導致欄排序錯亂。
   * 改為：直排比例夠高 **且** 「明顯橫向外框」span 未過半；另保留 writing-mode: vertical 與強訊號備援。
   */
  function inferVerticalReadingModeForPage(pageNumber) {
    return getPageReadingLayoutMeta(pageNumber).pageVertical;
  }

  /**
   * 僅依「本次框選／標記收集到的字元」推斷是否為直書區塊，用於覆寫整頁推斷。
   * 教材 PDF 常在直書正文旁有大量橫向注音 span，導致整頁 hRatio 過高而被判橫書，
   * 多欄直書標記後「原始文字」會變成左欄→右欄（錯誤）；此處以字元外框高寬比與欄分群補救。
   * @returns {boolean|null} true/false 強制直／橫；null 表示沿用 inferVerticalReadingModeForPage
   */
  function inferVerticalReadingModeForCollectedGlyphs(xywhList) {
    if (!xywhList || xywhList.length < 4) return null;
    let n = 0;
    let tall = 0;
    let sumW = 0;
    const centers = [];
    for (const g of xywhList) {
      const w = Math.max(0, Number(g.w) || 0);
      const h = Math.max(0, Number(g.h) || 0);
      if (w < 1.5 || h < 1.5) continue;
      n++;
      sumW += w;
      if (isVerticalishSize(w, h)) tall++;
      centers.push({
        cx: (Number(g.x) || 0) + w * 0.5,
        cy: (Number(g.y) || 0) + h * 0.5,
      });
    }
    if (n < 4) return null;
    const tallRatio = tall / n;
    const avgW = sumW / n;

    // 區塊內多數字元外框為「高明顯大於寬」→ 直書欄排序（右→左、上→下）
    if (tallRatio >= 0.36) return true;

    // 至少兩欄：x 中心有明顯斷層，且仍有一定比例直向字框（含注音混排）
    if (centers.length >= 5) {
      centers.sort((a, b) => a.cx - b.cx);
      const thr = Math.max(7, avgW * 0.42);
      let colBreaks = 0;
      for (let i = 1; i < centers.length; i++) {
        if (centers[i].cx - centers[i - 1].cx > thr) colBreaks++;
      }
      const clusters = colBreaks + 1;
      if (clusters >= 2 && tallRatio >= 0.16) return true;
    }

    // 幾乎都是扁寬框 → 明確橫書（覆寫整頁直書误判時用）
    if (tallRatio <= 0.14) return false;

    return null;
  }

  /**
   * 標記／框選矩形內相交的 span：若正文為 PDF 直排 transform，字元外框常被 clamp 成橫向，
   * 子集「高>寬」比例會失真；改以 span 的 computed transform 當強訊號（優先於子集幾何）。
   * @returns {boolean|null} true=強制直書排序；false=強制橫書；null=交給子集／頁級
   */
  function inferVerticalFromIntersectedSpans(spans) {
    if (!spans || !spans.length) return null;
    let v = 0;
    let h = 0;
    let amb = 0;
    for (const span of spans) {
      const rawFull = span.textContent || '';
      const raw = rawFull.replace(/\s/g, '');
      if (!raw) continue;
      let n = 0;
      for (const _ of raw) n++;
      if (spanLooksVerticallyTypeset(span)) {
        v += n;
      } else {
        try {
          const sr = span.getBoundingClientRect();
          if (sr.width > 0 && sr.height > 0 && sr.width >= sr.height * 0.9) {
            let hw = n;
            if (spanTextContainsExplicitLineBreak(rawFull)) {
              hw += n * 0.32;
            }
            h += hw;
          } else {
            amb += n;
          }
        } catch (e) {
          amb += n;
        }
      }
    }
    const tot = v + h + amb;
    if (tot < 4) return null;
    if (v / tot >= 0.24) return true;
    if (h / tot >= 0.55 && v / tot < 0.12) return false;
    return null;
  }

  /**
   * 直書（繁中常見）：欄位由右而左，欄內由上而下。
   * 橫書：由上而下，同行由左而右。
   */
  function compareReadingOrderByClientRect(ra, rb) {
    const vertA = isVerticalishSize(ra.width, ra.height);
    const vertB = isVerticalishSize(rb.width, rb.height);
    if (vertA && vertB) {
      const dx = ra.left - rb.left;
      if (Math.abs(dx) > 8) return rb.left - ra.left;
      const dy = ra.top - rb.top;
      if (Math.abs(dy) > 6) return dy;
      return rb.left - ra.left;
    }
    const dy = ra.top - rb.top;
    if (Math.abs(dy) > 6) return dy;
    return ra.left - rb.left;
  }

  /**
   * 橫書：依本批字元高度中位數推算「同行」容差；外框高低不一時用左上角 y 易錯行。
   */
  function horizontalLineToleranceForGlyphXYWHs(items, toXYWH) {
    const hs = [];
    for (const it of items) {
      const g = toXYWH(it);
      const h = Math.max(1, g.h || 0);
      hs.push(h);
    }
    hs.sort((a, b) => a - b);
    if (!hs.length) return 10;
    const mid = hs[Math.floor(hs.length / 2)] || hs[0];
    return Math.min(24, Math.max(8, mid * 0.45));
  }

  /**
   * @param {boolean} forceVertical 由 inferVerticalReadingModeForPage 預先算出，避免 sort 比較器內重複掃 DOM
   * @param {number} [lineTolerance] 橫書專用：比較字元 **垂直中心** 時的同行容差（px）
   */
  function compareReadingOrderGlyphsWithMode(a, b, forceVertical, lineTolerance) {
    if (forceVertical) {
      const dx = a.x - b.x;
      if (Math.abs(dx) > 8) return b.x - a.x;
      const dy = a.y - b.y;
      if (Math.abs(dy) > 6) return dy;
      return b.x - a.x;
    }
    const tol =
      typeof lineTolerance === 'number' && lineTolerance > 0 ? lineTolerance : 10;
    const acy = a.y + (a.h || 0) * 0.5;
    const bcy = b.y + (b.h || 0) * 0.5;
    const dy = acy - bcy;
    if (Math.abs(dy) > tol) return dy;
    return a.x - b.x;
  }

  function medianOfPositive(nums) {
    const a = nums.filter(n => Number.isFinite(n) && n > 0).sort((x, y) => x - y);
    if (!a.length) return 0;
    return a[Math.floor(a.length / 2)];
  }

  /**
   * 直書多欄：欄內字元 x 中心常有數 px 抖動；固定 8px 會把同欄上下字判成「換欄」而順序錯亂。
   * 依字寬與 x 分佈最大空隙估「欄間距」下沿的自適應容差。
   */
  function verticalInterColumnTolerance(chunk, toXYWH) {
    if (!chunk.length) return 14;
    const ws = chunk.map(it => Math.max(1, toXYWH(it).w || 0));
    const medianW = medianOfPositive(ws) || 12;
    const cxs = chunk
      .map(it => {
        const g = toXYWH(it);
        return (Number(g.x) || 0) + Math.max(1, g.w || 0) * 0.5;
      })
      .sort((a, b) => a - b);
    if (cxs.length < 2) {
      return Math.min(44, Math.max(8, medianW * 0.78));
    }
    let maxGap = 0;
    for (let i = 1; i < cxs.length; i++) {
      maxGap = Math.max(maxGap, cxs[i] - cxs[i - 1]);
    }
    const tFromW = medianW * 0.72;
    const tFromGap = maxGap > medianW * 1.55 ? maxGap * 0.3 : 0;
    return Math.min(46, Math.max(7, Math.max(tFromW, tFromGap)));
  }

  /** 直書：不同欄比較 x（容差內視同欄）；欄由右→左；同欄上→下 */
  function compareReadingOrderGlyphsVerticalWithTolerance(a, b, colTol) {
    const acx = (Number(a.x) || 0) + (Number(a.w) || 0) * 0.5;
    const bcx = (Number(b.x) || 0) + (Number(b.w) || 0) * 0.5;
    const acy = (Number(a.y) || 0) + (Number(a.h) || 0) * 0.5;
    const bcy = (Number(b.y) || 0) + (Number(b.h) || 0) * 0.5;
    const tol = Math.max(6, Math.min(50, colTol || 14));
    if (Math.abs(acx - bcx) > tol) {
      return bcx - acx;
    }
    if (Math.abs(acy - bcy) > 0.35) {
      return acy - bcy;
    }
    return acx - bcx;
  }

  /** 無頁級資訊時視為橫書 */
  function compareReadingOrderGlyphs(a, b) {
    return compareReadingOrderGlyphsWithMode(a, b, false, undefined);
  }

  function chunkHasExplicitMixedSpanFlags(chunk) {
    let t = 0;
    let f = 0;
    for (const it of chunk) {
      if (it.spanIsVertical === true) t++;
      else if (it.spanIsVertical === false) f++;
    }
    return t > 0 && f > 0;
  }

  /**
   * 單一字元排序時是否依「直書欄」規則；無 span 標記時看頁面 layout 與外框比例。
   */
  function glyphPrefersVerticalSortForReading(it, toXYWH) {
    if (it.spanIsVertical === true) return true;
    if (it.spanIsVertical === false) return false;
    const g = toXYWH(it);
    const w = Math.max(0, g.w || 0);
    const h = Math.max(0, g.h || 0);
    const pn = Number(it.pageNumber) || 0;
    const meta = pn > 0 ? getPageReadingLayoutMeta(pn) : null;
    if (!meta || meta.mode === 'horizontal') {
      return isVerticalishSize(w, h);
    }
    if (meta.mode === 'vertical') {
      return true;
    }
    return isVerticalishSize(w, h);
  }

  function sortGlyphChunkHybridReading(chunk, toXYWH) {
    const vertPart = chunk.filter(it => glyphPrefersVerticalSortForReading(it, toXYWH));
    const colTolV = verticalInterColumnTolerance(vertPart.length >= 2 ? vertPart : chunk, toXYWH);
    const horizPart = chunk.filter(it => !glyphPrefersVerticalSortForReading(it, toXYWH));
    const lineTolH = horizontalLineToleranceForGlyphXYWHs(horizPart.length >= 2 ? horizPart : chunk, toXYWH);
    chunk.sort((ia, ib) => {
      const va = glyphPrefersVerticalSortForReading(ia, toXYWH);
      const vb = glyphPrefersVerticalSortForReading(ib, toXYWH);
      const a = toXYWH(ia);
      const b = toXYWH(ib);
      if (va && vb) {
        return compareReadingOrderGlyphsVerticalWithTolerance(a, b, colTolV);
      }
      if (!va && !vb) {
        return compareReadingOrderGlyphsWithMode(a, b, false, lineTolH);
      }
      return compareReadingOrderByClientRect(
        { left: a.x, top: a.y, width: a.w, height: a.h },
        { left: b.x, top: b.y, width: b.w, height: b.h }
      );
    });
  }

  /**
   * 單頁字元區塊排序（標記／框選／整頁共用邏輯）。
   * @param {{ useSubsetVerticalInference?: boolean, spanVerticalHint?: boolean|null, workSpansForHint?: Element[], pageNumber?: number, forceHybridSort?: boolean }} options
   */
  function sortSinglePageGlyphChunk(chunk, toXYWH, options = {}) {
    if (!chunk.length) return;
    const p = options.pageNumber ?? chunk[0]?.pageNumber ?? 0;
    let fv = p > 0 && inferVerticalReadingModeForPage(p);

    let sh = options.spanVerticalHint;
    const wsh = options.workSpansForHint;
    if (wsh && wsh.length && p > 0) {
      const sub = wsh.filter(
        s => parseInt(s.closest('.page')?.dataset?.pageNumber || '0', 10) === p
      );
      const local = inferVerticalFromIntersectedSpans(sub);
      if (local === true || local === false) {
        sh = local;
      }
    }
    if (sh === true) {
      fv = true;
    } else if (sh === false) {
      fv = false;
    } else if (options.useSubsetVerticalInference && chunk.length >= 4) {
      const hint = inferVerticalReadingModeForCollectedGlyphs(chunk.map(it => toXYWH(it)));
      if (hint !== null && hint !== undefined) {
        fv = hint;
      }
    }

    const meta = p > 0 ? getPageReadingLayoutMeta(p) : null;
    const useHybrid =
      options.forceHybridSort === true ||
      meta?.mode === 'mixed' ||
      chunkHasExplicitMixedSpanFlags(chunk);

    if (useHybrid) {
      sortGlyphChunkHybridReading(chunk, toXYWH);
      return;
    }
    if (fv) {
      const colTol = verticalInterColumnTolerance(chunk, toXYWH);
      chunk.sort((ia, ib) => {
        const a = toXYWH(ia);
        const b = toXYWH(ib);
        return compareReadingOrderGlyphsVerticalWithTolerance(a, b, colTol);
      });
    } else {
      const lineTol = horizontalLineToleranceForGlyphXYWHs(chunk, toXYWH);
      chunk.sort((ia, ib) => {
        const a = toXYWH(ia);
        const b = toXYWH(ib);
        return compareReadingOrderGlyphsWithMode(a, b, false, lineTol);
      });
    }
  }

  /**
   * 多頁字元各頁推斷直書後再排序接回。
   * @param {{ useSubsetVerticalInference?: boolean, spanVerticalHint?: boolean|null, workSpansForHint?: Element[] }} [options]
   *        workSpansForHint：與選區相交的 span，依頁分別做 transform 推斷（框選未帶 pageFilter 時仍有效）
   */
  function sortGlyphLikeListByReadingOrder(items, toXYWH, options = {}) {
    const byPage = new Map();
    for (const item of items) {
      const p = item.pageNumber || 0;
      if (!byPage.has(p)) byPage.set(p, []);
      byPage.get(p).push(item);
    }
    const pageNums = Array.from(byPage.keys()).sort((a, b) => a - b);
    const out = [];
    for (const p of pageNums) {
      const chunk = byPage.get(p);
      sortSinglePageGlyphChunk(chunk, toXYWH, { ...options, pageNumber: p });
      out.push(...chunk);
    }
    return out;
  }

  /** 避免直書窄字被強制 min 0.1% 寬高變成「橫向方塊」 */
  function highlightPercentDims(relWidth, relHeight) {
    const rw = Math.max(0, relWidth || 0) * 100;
    const rh = Math.max(0, relHeight || 0) * 100;
    const minW = rh > rw * 1.25 ? 0.02 : 0.1;
    const minH = rw > rh * 1.25 ? 0.02 : 0.1;
    return {
      w: Math.max(minW, rw),
      h: Math.max(minH, rh),
    };
  }

  /** PDF 文字層常把 ① 拆成「○」「１」等不同 span；若以字元座標全域排序會與閱讀順序錯亂。改依 span 順序 + span 內字元順序。 */
  function compareSpanReadingOrder(a, b) {
    const ra = a.getBoundingClientRect();
    const rb = b.getBoundingClientRect();
    return compareReadingOrderByClientRect(ra, rb);
  }

  if (typeof window !== 'undefined') {
    window.pdfViewerSpanReadingCompare = (a, b) => compareSpanReadingOrder(a, b);
    /** 供 viewer-reading-marks 等：與朗讀高亮同一套直書 rect 修正 */
    window.pdfViewerGlyphRectsForCharRange = (range, span, charIndex, charsInNode, ch) => {
      if (!range) return [];
      const picked = pickGlyphRectsFromRange(range, ch, span);
      const r0 = picked[0];
      if (!r0) return [];
      if (span && typeof span.getBoundingClientRect === 'function') {
        return [refineGlyphRectWithSpan(span, r0, charIndex, charsInNode, ch)];
      }
      return [r0];
    };
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

  /**
   * 將「螢幕座標」下的矩形轉成 .textLayer 內 left/top/width/height 所用的百分比（0~1）。
   * 必須與 PDF.js EditorUIManager.getSelectionBoxes（pdf.mjs）依 data-main-rotation 的公式一致；
   * 僅用 (x-layerX)/layerW 在 rotation 90/180/270 時會錯位，直書／旋轉頁看起來像一整條橫的。
   */
  function clientRectToTextLayerPercentages(textLayerEl, clientBox) {
    const layerRect = textLayerEl.getBoundingClientRect();
    const layerX = layerRect.left;
    const layerY = layerRect.top;
    const parentWidth = layerRect.width || 1;
    const parentHeight = layerRect.height || 1;
    const x = clientBox.left;
    const y = clientBox.top;
    const w = clientBox.width;
    const h = clientBox.height;
    const rot = textLayerEl.getAttribute('data-main-rotation') || '0';
    let relLeft;
    let relTop;
    let relWidth;
    let relHeight;
    switch (rot) {
      case '90':
        relLeft = (y - layerY) / parentHeight;
        relTop = 1 - (x + w - layerX) / parentWidth;
        relWidth = h / parentHeight;
        relHeight = w / parentWidth;
        break;
      case '180':
        relLeft = 1 - (x + w - layerX) / parentWidth;
        relTop = 1 - (y + h - layerY) / parentHeight;
        relWidth = w / parentWidth;
        relHeight = h / parentHeight;
        break;
      case '270':
        relLeft = 1 - (y + h - layerY) / parentHeight;
        relTop = (x - layerX) / parentWidth;
        relWidth = h / parentHeight;
        relHeight = w / parentWidth;
        break;
      default:
        relLeft = (x - layerX) / parentWidth;
        relTop = (y - layerY) / parentHeight;
        relWidth = w / parentWidth;
        relHeight = h / parentHeight;
        break;
    }
    return { relLeft, relTop, relWidth, relHeight };
  }

  if (typeof window !== 'undefined') {
    window.pdfViewerClientRectToTextLayerPercents = clientRectToTextLayerPercentages;
  }

  /** 與朗讀用 ordered/deduped 字元同一組座標，避免 hitRects 掃描順序與實際文本不一致 */
  function buildRectMetaFromCollectedGlyph(g) {
    const pageNumber = g.pageNumber;
    if (!pageNumber) return null;
    const textLayerEl = document.querySelector(
      `.page[data-page-number="${pageNumber}"] .textLayer`
    );
    const layerRect = textLayerEl?.getBoundingClientRect();
    if (!textLayerEl || !layerRect || layerRect.width < 1 || layerRect.height < 1) return null;
    const left = g.x;
    const top = g.y;
    const width = g.w;
    const height = g.h;
    const rel = clientRectToTextLayerPercentages(textLayerEl, { left, top, width, height });
    return {
      left,
      top,
      width,
      height,
      pageNumber,
      layerClass: 'textLayer',
      relLeft: rel.relLeft,
      relTop: rel.relTop,
      relWidth: rel.relWidth,
      relHeight: rel.relHeight,
    };
  }

  /**
   * 收集與 client 矩形相交的字元（不依 span DOM 順序串接）。
   * 直書多欄時 PDF 的 span 順序常與閱讀序不同；先收集再依頁級直書推斷＋幾何排序（與 getPageGlyphs 一致）。
   */
  function collectGlyphsIntersectingClientRect(frameRect, opts = {}) {
    const pageFilter = opts.pageNumber;
    const minOverlap = opts.minOverlapRatio ?? 0.55;

    const fr = {
      left: frameRect.left,
      top: frameRect.top,
      right: frameRect.right,
      bottom: frameRect.bottom,
    };

    const spans = Array.from(document.querySelectorAll('#viewerContainer .textLayer span')).filter(span => {
      if (pageFilter != null) {
        const pn = parseInt(span.closest('.page')?.dataset?.pageNumber || '0', 10);
        if (pn !== pageFilter) return false;
      }
      const spanRect = span.getBoundingClientRect();
      return !(
        spanRect.right < fr.left ||
        spanRect.left > fr.right ||
        spanRect.bottom < fr.top ||
        spanRect.top > fr.bottom
      );
    });

    const workSpans = dedupeBoldOverlaySpans(spans);
    const sortOpts = {
      useSubsetVerticalInference: true,
      workSpansForHint: workSpans,
    };
    const raw = [];

    for (const span of workSpans) {
      const pageNumber = parseInt(span.closest('.page')?.dataset?.pageNumber || '0', 10) || 0;
      if (!pageNumber) continue;

      const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
      let textNode;
      while ((textNode = walker.nextNode())) {
        const txt = textNode.nodeValue || '';
        if (!txt) continue;

        for (let i = 0; i < txt.length; i++) {
          const range = document.createRange();
          range.setStart(textNode, i);
          range.setEnd(textNode, i + 1);
          const rectList = pickGlyphRectsFromRange(range, txt[i], span);
          let representativeRect = null;

          for (const r0 of rectList) {
            const r = refineGlyphRectWithSpan(span, r0, i, txt.length, txt[i]);
            if (!r || r.width <= 0 || r.height <= 0) continue;
            const rr = {
              left: r.left,
              top: r.top,
              right: r.right,
              bottom: r.bottom,
              width: r.width,
              height: r.height,
            };
            if (!isRectHitFlexible(fr, rr, minOverlap)) continue;
            representativeRect = r;
            break;
          }

          if (representativeRect) {
            raw.push({
              ch: txt[i],
              x: representativeRect.left,
              y: representativeRect.top,
              w: representativeRect.width,
              h: representativeRect.height,
              pageNumber,
              spanIsVertical: spanLooksVerticallyTypeset(span),
            });
          }
        }
      }
    }

    const sortedRaw = sortGlyphLikeListByReadingOrder(
      raw,
      g => ({
        x: g.x,
        y: g.y,
        w: g.w,
        h: g.h,
      }),
      sortOpts
    );

    const deduped = [];
    for (const g of sortedRaw) {
      const duplicated = deduped.some(
        d => d.ch === g.ch && Math.abs(d.x - g.x) <= 3 && Math.abs(d.y - g.y) <= 3
      );
      if (!duplicated) deduped.push(g);
    }
    return deduped;
  }

  // 收集框內文字，並同時回傳被框住的字元矩形
  function collectTextAndSpansInFrame() {
    const frameRect = frame.getBoundingClientRect();
    const deduped = collectGlyphsIntersectingClientRect(frameRect, { minOverlapRatio: 0.55 });
    const text = collapseAdjacentDuplicateCjkRuns(deduped.map(g => g.ch).join(''));
    const rects = deduped.map(g => buildRectMetaFromCollectedGlyph(g)).filter(Boolean);
    return { text, rects };
  }

  if (typeof window !== 'undefined') {
    /** 標記區「原始文字」：與框選朗讀相同字元排序（直書右→左欄、欄內上→下） */
    window.pdfViewerExtractTextForMarkRect = (pageNum, frameRect) => {
      const fr = {
        left: frameRect.left,
        top: frameRect.top,
        right: frameRect.right,
        bottom: frameRect.bottom,
      };
      const deduped = collectGlyphsIntersectingClientRect(fr, {
        pageNumber: pageNum,
        minOverlapRatio: 0.45,
      });
      let out = deduped.map(g => g.ch).join('');
      out = collapseAdjacentDuplicateCjkRuns(out);
      return out.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
    };
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
        el.style.boxSizing = 'border-box';
        el.style.left = `${(r.relLeft || 0) * 100}%`;
        el.style.top = `${(r.relTop || 0) * 100}%`;
        const hp = highlightPercentDims(r.relWidth, r.relHeight);
        el.style.width = `${hp.w}%`;
        el.style.height = `${hp.h}%`;
        textLayerEl.insertBefore(el, textLayerEl.firstChild);
      } else {
        el.style.boxSizing = 'border-box';
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
          const picked = pickGlyphRectsFromRange(range, txt[i], el);
          const rawRect = picked[0];
          if (!rawRect || rawRect.width <= 0 || rawRect.height <= 0) continue;
          const rect = refineGlyphRectWithSpan(el, rawRect, i, txt.length, txt[i]);
          if (!rect || rect.width <= 0 || rect.height <= 0) continue;

          glyphs.push({
            ch: txt[i],
            x: rect.left,
            y: rect.top,
            w: rect.width,
            h: rect.height,
            pageNumber,
            spanIsVertical: spanLooksVerticallyTypeset(el),
          });
        }
      }
    }

    sortSinglePageGlyphChunk(glyphs, g => g, {
      pageNumber,
      workSpansForHint: Array.from(pageLayer.querySelectorAll('span')),
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
    const rel = clientRectToTextLayerPercentages(textLayerEl, {
      left: g.x,
      top: g.y,
      width: g.w,
      height: g.h,
    });
    return {
      pageNumber: g.pageNumber,
      left: g.x,
      top: g.y,
      width: g.w,
      height: g.h,
      relLeft: rel.relLeft,
      relTop: rel.relTop,
      relWidth: rel.relWidth,
      relHeight: rel.relHeight,
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
        el.style.boxSizing = 'border-box';
        el.style.left = `${(r.relLeft || 0) * 100}%`;
        el.style.top = `${(r.relTop || 0) * 100}%`;
        const hp = highlightPercentDims(r.relWidth, r.relHeight);
        el.style.width = `${hp.w}%`;
        el.style.height = `${hp.h}%`;
        textLayerEl.insertBefore(el, textLayerEl.firstChild);
      } else {
        el.style.boxSizing = 'border-box';
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
    const audio = document.getElementById('readingAudioPlayer');
    if (!bar || !voiceSelect || !speedWrap) return;

    const PAD_X = 12;
    const PAD_BOTTOM = 12;
    const TOP_VOICE = 70;
    const GAP = 12;

    // 語言／語音選單：右上角
    voiceSelect.style.top = `${TOP_VOICE}px`;
    voiceSelect.style.right = `${PAD_X}px`;
    voiceSelect.style.left = 'auto';
    voiceSelect.style.bottom = 'auto';

    // 底部置中：播放列 + 語速滑軌 +（若有）audio player
    const bw = bar.offsetWidth || 250;
    const sw = speedWrap.offsetWidth || 260;
    let aw = 0;
    if (audio && audio.isConnected) {
      const rw = audio.getBoundingClientRect().width;
      aw = rw >= 8 ? rw : 280;
    }
    const totalW = bw + GAP + sw + (aw ? GAP + aw : 0);
    let groupLeft = Math.round((window.innerWidth - totalW) / 2);
    groupLeft = Math.max(PAD_X, Math.min(groupLeft, window.innerWidth - PAD_X - totalW));

    bar.style.bottom = `${PAD_BOTTOM}px`;
    bar.style.top = 'auto';
    bar.style.left = `${groupLeft}px`;
    bar.style.right = 'auto';

    const speedLeft = groupLeft + bw + GAP;
    speedWrap.style.bottom = `${PAD_BOTTOM}px`;
    speedWrap.style.top = 'auto';
    speedWrap.style.left = `${speedLeft}px`;
    speedWrap.style.right = 'auto';

    if (audio) {
      const audioLeft = speedLeft + sw + GAP;
      audio.style.position = 'fixed';
      audio.style.bottom = `${PAD_BOTTOM}px`;
      audio.style.top = 'auto';
      audio.style.left = `${audioLeft}px`;
      audio.style.right = 'auto';
      audio.style.transform = 'none';
      audio.style.zIndex = '3100';
      audio.style.maxWidth = `${Math.min(360, Math.round(window.innerWidth * 0.42))}px`;
    }
  }

  if (typeof window !== 'undefined') {
    window.layoutTtsPlaybackBar = layoutTtsPlaybackBar;
  }

  function ensureTtsPlaybackBar() {
    if (document.getElementById('ttsPlaybackBar')) return;
    const bar = document.createElement('div');
    bar.id = 'ttsPlaybackBar';
    bar.style.position = 'fixed';
    bar.style.bottom = '12px';
    bar.style.top = 'auto';
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

  /** PDF 載入／文字層重繪後掃描各頁，預先填入 getPageReadingLayoutMeta（直／橫／混合）。 */
  function prefetchReadingLayoutForAllPages() {
    const app = window.PDFViewerApplication;
    const n = app?.pdfDocument?.numPages;
    if (!n || !Number.isFinite(n)) return;
    for (let p = 1; p <= n; p++) {
      const tl = document.querySelector(`.page[data-page-number="${p}"] .textLayer`);
      if (tl && tl.querySelector('span')) {
        getPageReadingLayoutMeta(p);
      }
    }
  }

  let _readingLayoutPrefetchTid = 0;
  function schedulePrefetchReadingLayoutDebounced() {
    clearTimeout(_readingLayoutPrefetchTid);
    _readingLayoutPrefetchTid = setTimeout(() => {
      prefetchReadingLayoutForAllPages();
      _readingLayoutPrefetchTid = 0;
    }, 180);
  }

  if (typeof window !== 'undefined') {
    window.pdfViewerGetPageReadingLayout = pageNum => {
      const m = getPageReadingLayoutMeta(pageNum);
      return { ...m };
    };
    window.pdfViewerPrefetchReadingLayouts = prefetchReadingLayoutForAllPages;
  }

  function bindModelRepaintEvents() {
    const app = window.PDFViewerApplication;
    const eb = app?.eventBus;
    if (!eb || bindModelRepaintEvents._bound) return false;
    bindModelRepaintEvents._bound = true;
    const rerender = () => rebuildReadHighlightsFromModel();
    const runPrefetch = () => schedulePrefetchReadingLayoutDebounced();
    eb._on('pagesloaded', runPrefetch);
    eb._on('textlayerrendered', runPrefetch);
    eb._on('pagerendered', (e) => {
      clearVerticalReadingModeCacheForPage(e?.pageNumber);
      rerender();
    });
    eb._on('scalechanging', () => {
      clearVerticalReadingModeCacheForPage('');
      rerender();
    });
    eb._on('rotationchanging', () => {
      clearVerticalReadingModeCacheForPage('');
      rerender();
    });
    eb._on('updateviewarea', rerender);
    schedulePrefetchReadingLayoutDebounced();
    return true;
  }

  function sortGlyphsByReadingOrder(glyphList) {
    if (!glyphList.length) return glyphList;
    return sortGlyphLikeListByReadingOrder(
      glyphList,
      g => ({
        x: g.rects?.[0]?.left ?? g.x ?? 0,
        y: g.rects?.[0]?.top ?? g.y ?? 0,
        w: g.rects?.[0]?.width ?? g.w ?? 0,
        h: g.rects?.[0]?.height ?? g.h ?? 0,
      }),
      { useSubsetVerticalInference: true }
    );
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
    const rel = clientRectToTextLayerPercentages(textLayerEl, {
      left: il,
      top: it,
      width: ir - il,
      height: ib - it,
    });
    return {
      pageNumber,
      left: il,
      top: it,
      width: ir - il,
      height: ib - it,
      relLeft: rel.relLeft,
      relTop: rel.relTop,
      relWidth: rel.relWidth,
      relHeight: rel.relHeight,
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
   * 粗體／疊層 PDF 文字常造成相鄰重複（如「標題標題」、HelloHello）；送 TTS 前收斂。
   * 與框選字元收集用的 CJK 規則一致，並加強英文單字、數字段的重複相接。
   */
  function collapseAdjacentDuplicateRunsForTts(s) {
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
    prev = '';
    guard = 0;
    while (prev !== t && guard++ < 16) {
      prev = t;
      t = t.replace(/([A-Za-z]{3,})(?:\1)+/g, '$1');
      t = t.replace(/([0-9０-９]{2,})(?:\1)+/g, '$1');
    }
    return t;
  }

  /**
   * 各類帶圈／括號選項碼（①、➊、⓵、⑴、❶、㈠、㊀、⒈…）皆還原為半形數字，TTS 前後加「，」停頓。
   * 須先替換再 NFKC：否則 NFKC 會先把部分符號打散，不利偵測。
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

    // ㈠–㈩、㊀–㊉：與 ① 相同，半形 1–10，TTS 為 ，1，
    swapRange(0x3220, 9, 1);
    swapOne(0x3229, 10);
    swapRange(0x3280, 9, 1);
    swapOne(0x3289, 10);

    // ⒈–⒚（數字+句點）
    swapRange(0x2488, 9, 1);
    swapOne(0x2491, 10);
    swapRange(0x2492, 10, 11);

    swapOne(0x1f10b, 0);
    swapOne(0x1f10c, 0);
    swapOne(0x1f101, 0);
    swapRange(0x1f102, 9, 1);

    function parseParensChineseNumeral(inner) {
      const d = {
        〇: 0,
        零: 0,
        一: 1,
        二: 2,
        三: 3,
        四: 4,
        五: 5,
        六: 6,
        七: 7,
        八: 8,
        九: 9,
      };
      if (!inner) return null;
      if (inner === '十') return 10;
      if (inner.length === 1 && d[inner] !== undefined) return d[inner];
      if (inner.length === 2 && inner[0] === '十' && d[inner[1]] !== undefined) return 10 + d[inner[1]];
      if (inner.length === 2 && inner[1] === '十' && d[inner[0]] !== undefined && d[inner[0]] > 0)
        return d[inner[0]] * 10;
      if (
        inner.length === 3 &&
        inner[1] === '十' &&
        d[inner[0]] !== undefined &&
        d[inner[2]] !== undefined
      )
        return d[inner[0]] * 10 + d[inner[2]];
      return null;
    }

    function replaceParenthesizedNumerals(str) {
      let t = str;
      t = t.replace(/[（(]\s*([〇零一二三四五六七八九十]{1,3})\s*[）)]/gu, (m, inner) => {
        const n = parseParensChineseNumeral(inner.replace(/\s/g, ''));
        return n != null ? ttsPauseAround(String(n)) : m;
      });
      t = t.replace(/[（(]\s*([0-9０-９]{1,2})\s*[）)]/gu, (m, numStr) => {
        const trimmed = numStr.replace(/\s/g, '');
        let ds = '';
        for (const ch of trimmed) {
          const idx = fw.indexOf(ch);
          if (idx >= 0) ds += String(idx);
          else if (ch >= '0' && ch <= '9') ds += ch;
          else return m;
        }
        const n = parseInt(ds, 10);
        if (!Number.isFinite(n) || n < 0 || n > 99) return m;
        return ttsPauseAround(String(n));
      });
      return t;
    }

    out = replaceParenthesizedNumerals(out);

    /** 半形／全形數字 → 半形數字字串；無法解析則 null */
    function digitCharToNumStr(d) {
      const idx = fw.indexOf(d);
      if (idx >= 0) return String(idx);
      const v = parseInt(d, 10);
      return Number.isFinite(v) ? String(v) : null;
    }

    // 白圓／空心圓等（含 U+25CB、U+3007 〇、●）；全形空格等置於圓與數字之間時亦須配對
    const circleLike = '[\u25CB\u3007\u25EF\u25E6\u25CF\u25CC]';
    const betweenDigitCircle =
      '[\\s\\u00a0\\u3000\\u2000-\\u200d\\u202f\\u205f\\ufeff，,、．.·\u00b7•\u2022]*';

    function replaceCircleDigitPairs(str) {
      const digitClass = '[\\d\\uFF10-\\uFF19]';
      let t = str;
      t = t.replace(new RegExp(`${circleLike}${betweenDigitCircle}(${digitClass})`, 'gu'), (_, d) => {
        const ns = digitCharToNumStr(d);
        return ns != null ? ttsPauseAround(ns) : _;
      });
      t = t.replace(new RegExp(`(${digitClass})${betweenDigitCircle}${circleLike}`, 'gu'), (_, d) => {
        const ns = digitCharToNumStr(d);
        return ns != null ? ttsPauseAround(ns) : _;
      });
      return t;
    }

    // 先去掉零寬字元，避免「2」與「○」視覺相鄰但字串不相鄰
    out = out.replace(/[\u200b-\u200d\ufeff]/g, '');

    // NFKC 前先處理拆字組合（避免 ② 等先被 NFKC 打成半形數字後無法套帶圈規則）
    out = replaceCircleDigitPairs(out);

    try {
      out = out.normalize('NFKC');
    } catch (e) {
      /* ignore */
    }

    // NFKC 後再掃一次：殘留的「2○」「〇3」等（區塊切分／排序後仍可能出現）
    out = replaceCircleDigitPairs(out);

    out = replaceParenthesizedNumerals(out);

    out = out.replace(/，{2,}/g, '，');

    return out;
  }

  /** 與 index2：先 TTS 正規化再套試算表（key 用 ① 或 TTS 後的 ，1， 皆可） */
  function applySheetReplacementsForReading(sourceText, data) {
    let out = normalizeCircledDigitsForTts(sourceText || '');
    if (!data?.table?.rows) return out;
    const fwDigits = '０１２３４５６７８９';
    data.table.rows.forEach((row, index) => {
      const original = cellToString(row.c?.[0]);
      const replacement = cellToString(row.c?.[1]);
      if (!original || !replacement) return;
      const origTrim = original.trim();
      const replTts = normalizeCircledDigitsForTts(replacement.trim());
      const keyTts = normalizeCircledDigitsForTts(origTrim);
      const isBareDigit =
        /^[0-9]$/.test(origTrim) || (origTrim.length === 1 && fwDigits.includes(origTrim));
      let replaced = false;
      if (!isBareDigit && keyTts && out.includes(keyTts)) {
        out = out.split(keyTts).join(replTts);
        replaced = true;
      } else if (!isBareDigit && origTrim && out.includes(origTrim)) {
        out = out.split(origTrim).join(replTts);
        replaced = true;
      }
      if (replaced) {
        console.log(`試算表替換（朗讀）第 ${index + 2} 行: "${original}" -> "${replacement}"`);
      }
    });
    return out;
  }

  if (typeof window !== 'undefined') {
    window.normalizeCircledDigitsForTts = normalizeCircledDigitsForTts;
    window.collapseAdjacentDuplicateRunsForTts = collapseAdjacentDuplicateRunsForTts;
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

    cleanedText = collapseAdjacentDuplicateRunsForTts(cleanedText);

    let filteredText;
    const sheetId = getSheetIdFromUrl();
    if (sheetId) {
      const data = await fetchSheetDataCached(sheetId);
      filteredText = applySheetReplacementsForReading(cleanedText, data);
    } else {
      filteredText = normalizeCircledDigitsForTts(cleanedText);
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
        audio.controls = true;
        audio.style.position = 'fixed';
        audio.style.zIndex = '3100';
        document.body.appendChild(audio);
      }
      audio.addEventListener(
        'loadedmetadata',
        () => {
          if (typeof window.layoutTtsPlaybackBar === 'function') window.layoutTtsPlaybackBar();
        },
        { once: true }
      );
      if (typeof window.layoutTtsPlaybackBar === 'function') {
        window.layoutTtsPlaybackBar();
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

    /** 載入失敗時須關閉遮罩，否則錯誤訊息被蓋住且看似永遠轉圈 */
    const tryBindDocumentErrorToHideCover = () => {
      if (initialCover.dataset.pdfjsDocErrorBound === '1') return;
      const ebErr = window.PDFViewerApplication?.eventBus;
      if (!ebErr || typeof ebErr._on !== 'function') {
        setTimeout(tryBindDocumentErrorToHideCover, 50);
        return;
      }
      initialCover.dataset.pdfjsDocErrorBound = '1';
      ebErr._on('documenterror', () => {
        try {
          initialCover.dataset.pdfjsCoverBound = '1';
          scheduleHideCover();
        } catch (_) {}
      });
    };
    tryBindDocumentErrorToHideCover();

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
      // 不可在此 return：會略過 load 回呼後段的 resetReadingFrameDefault，導致 #readingFrame 永遠 display:none
    } else {
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

      // IndexedDB／小檔本機開啟時，PDF 常在 window.load 之前就載入並 dispatch 完 pagesloaded、pagerendered；
      // 此時才在 load 裡註冊的 _on 會全部錯過 → 畫面永遠「載入 PDF」但主控台已顯示 PDF 已開啟。
      const PDF_PAGE_RENDER_FINISHED = 3; // 等同 viewer 內 RenderingStates.FINISHED
      const bootstrapCoverIfViewerAlreadyRendered = () => {
        if (hidden) return;
        syncPagesLoaded();
        try {
          const a = window.PDFViewerApplication;
          const pv = a?.pdfViewer;
          if (!a?.pdfDocument || !pv) return;
          const pn = Math.max(1, pv.currentPageNumber || 1);
          const pageView = typeof pv.getPageView === 'function' ? pv.getPageView(pn - 1) : null;
          const finished = pageView?.renderingState === PDF_PAGE_RENDER_FINISHED;
          const wrap = pageView?.div;
          const canvas = wrap?.querySelector?.('canvas');
          const canvasOk = !!(canvas && canvas.width >= 2 && canvas.height >= 2);
          if (finished && canvasOk) {
            if (!coverWait.firstCanvasAt) {
              coverWait.firstCanvasAt = Date.now();
            }
            coverWait.textLayerOk = true;
            tryRevealCover();
          }
        } catch (_) {}
      };
      [0, 50, 120, 300, 700, 1600, 3200].forEach(ms => {
        setTimeout(bootstrapCoverIfViewerAlreadyRendered, ms);
      });
      requestAnimationFrame(() => {
        requestAnimationFrame(bootstrapCoverIfViewerAlreadyRendered);
      });

      const COVER_STUCK_FORCE_MS = 5500;
      setTimeout(() => {
        if (hidden || !initialCover?.isConnected) return;
        bootstrapCoverIfViewerAlreadyRendered();
        if (hidden) return;
        try {
          if (!window.PDFViewerApplication?.pdfDocument) return;
          const vc = document.getElementById('viewerContainer');
          const anyCanvas = vc?.querySelector?.('.page canvas');
          if (anyCanvas && anyCanvas.width >= 2) {
            coverWait.pagesLoaded = true;
            if (!coverWait.firstCanvasAt) {
              coverWait.firstCanvasAt = Date.now();
            }
            coverWait.textLayerOk = true;
            tryRevealCover();
          }
        } catch (_) {}
      }, COVER_STUCK_FORCE_MS);
    };

    bindInitialCoverWhenReady();

    if (app?.pdfDocument) {
      requestAnimationFrame(() => kickPdfViewerVisibleRefresh());
    }
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

