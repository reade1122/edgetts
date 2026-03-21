/**
 * 須於 viewer.mjs 之前執行。比照語音報讀系統個人版：
 * - ?file=<GoogleDriveFileId>：經 GAS 代理下載 → 寫入 pdfjs-viewer IndexedDB → ?file=idb://last-opened-file
 * - ?file=https://...pdf：**優先** fetch 完整檔寫入同上 IndexedDB → idb://（與雲端硬碟一致，標記層／載入時序較穩）；
 *   若 CORS 等無法 fetch，則不變更網址，交由 PDF.js 直接 url 開啟。
 * - ?file=xxx.pdf：若 IndexedDB 內已有同名上次檔案，改為 idb:// 以符合 viewer 流程
 */
const IDB_DB_NAME = "pdfjs-viewer";
const IDB_STORE_NAME = "files";
const IDB_LAST_FILE_KEY = "last-opened-file";
const IDB_LAST_FILE_URL_MARKER = "idb://last-opened-file";

const gasUrl =
  (typeof window !== "undefined" &&
    window.PDF_VIEWER_CONFIG &&
    window.PDF_VIEWER_CONFIG.GAS_PROXY_URL) ||
  "https://script.google.com/macros/s/AKfycbyD5dVne4RQ0gbQm_wbT-9No18RsADs2I78tZ8NOEwF9-75QkeYRYdWD7diV2r1z94i/exec";

/** 錯誤提示用；雲端下載進度改由 viewer #initialRenderCover 顯示，與其他載入方式一致 */
function ensureSwal() {
  return new Promise((resolve) => {
    if (typeof window !== "undefined" && window.Swal) {
      resolve();
      return;
    }
    if (typeof document === "undefined") {
      resolve();
      return;
    }
    if (!document.querySelector('link[href*="sweetalert2"][rel="stylesheet"]')) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = "https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css";
      document.head.appendChild(link);
    }
    const existing = document.querySelector("script[data-pdf-bootstrap-swal]");
    if (existing) {
      const wait = () => (window.Swal ? resolve() : setTimeout(wait, 30));
      if (window.Swal) resolve();
      else wait();
      return;
    }
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js";
    script.dataset.pdfBootstrapSwal = "1";
    script.onload = () => resolve();
    script.onerror = () => resolve();
    document.head.appendChild(script);
  });
}

function closeSwalIfAny() {
  try {
    if (window.Swal) window.Swal.close();
  } catch (e) {
    /* ignore */
  }
}

function openViewerDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(IDB_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(IDB_STORE_NAME)) {
        db.createObjectStore(IDB_STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveFileToIndexedDB(file, markSourceCloud = false) {
  const db = await openViewerDB();
  try {
    const buffer = await file.arrayBuffer();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE_NAME, "readwrite");
      const store = tx.objectStore(IDB_STORE_NAME);
      store.put(
        {
          name: file.name,
          type: file.type || "application/pdf",
          data: buffer,
          markSourceCloud: !!markSourceCloud,
        },
        IDB_LAST_FILE_KEY
      );
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  } finally {
    db.close();
  }
}

/** 不重新下載，只更新 last-opened-file 的 markSourceCloud（例如 ?file=同名.pdf 捷徑視為本機快取） */
async function patchLastOpenedMarkSourceCloud(markSourceCloud) {
  const saved = await readFileFromIndexedDB();
  if (!saved?.data) return;
  const db = await openViewerDB();
  try {
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE_NAME, "readwrite");
      const store = tx.objectStore(IDB_STORE_NAME);
      store.put(
        {
          ...saved,
          markSourceCloud: !!markSourceCloud,
        },
        IDB_LAST_FILE_KEY
      );
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  } finally {
    db.close();
  }
}

async function readFileFromIndexedDB() {
  const db = await openViewerDB();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE_NAME, "readonly");
      const store = tx.objectStore(IDB_STORE_NAME);
      const req = store.get(IDB_LAST_FILE_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

function replaceFileQueryWithIdbMarker() {
  const u = new URL(window.location.href);
  u.searchParams.set("file", IDB_LAST_FILE_URL_MARKER);
  try {
    localStorage.setItem("pdfjs.lastFile", IDB_LAST_FILE_URL_MARKER);
  } catch (e) {
    /* ignore */
  }
  history.replaceState({}, "", u.toString());
}

/**
 * 雲端 PDF 已寫入 IndexedDB 並改為 idb:// 後，整頁重整一次（等同使用者 F5），
 * 並通知 viewer-main 勿在本輪繼續 import viewer（見 viewer-main.mjs）。
 */
function reloadOnceAfterCloudIdbWrite() {
  try {
    globalThis.__PDFJS_RELOADING_FOR_CLOUD_IDB__ = true;
  } catch (e) {
    /* ignore */
  }
  location.reload();
}

async function fetchWithTimeout(url, options = {}, timeoutMs = 6000000) {
  return Promise.race([
    fetch(url, options),
    new Promise((_, reject) => setTimeout(() => reject(new Error("請求超時")), timeoutMs)),
  ]);
}

function pickFilenameFromHttpResponse(url, headers) {
  const cd = headers && headers.get ? headers.get("Content-Disposition") : null;
  if (cd) {
    const star = /filename\*\s*=\s*(?:UTF-8''|)([^;\n]+)/i.exec(cd);
    if (star && star[1]) {
      try {
        return decodeURIComponent(star[1].trim().replace(/^["']|["']$/g, ""));
      } catch (_) {
        /* fall through */
      }
    }
    const quoted = /filename\s*=\s*"([^"]+)"/i.exec(cd);
    if (quoted && quoted[1]) return quoted[1];
    const plain = /filename\s*=\s*([^;\s]+)/i.exec(cd);
    if (plain && plain[1]) {
      try {
        return decodeURIComponent(plain[1].trim());
      } catch (_) {
        return plain[1].trim();
      }
    }
  }
  try {
    const u = new URL(url);
    const last = u.pathname.split("/").filter(Boolean).pop();
    if (last && /\.pdf$/i.test(last)) return decodeURIComponent(last);
  } catch (_) {
    /* ignore */
  }
  return "document.pdf";
}

/**
 * 將 http(s) PDF 先完整下載並寫入 last-opened-file；呼叫端再 replace 成 idb://。
 * markSourceCloud=true 與 GAS 下載相同。
 */
async function loadHttpUrlIntoIndexedDB(url) {
  let absUrl;
  try {
    absUrl = new URL(url, window.location.href).href;
  } catch (e) {
    absUrl = url;
  }
  let sameOrigin = false;
  try {
    sameOrigin = new URL(absUrl).origin === window.location.origin;
  } catch (e) {
    /* ignore */
  }
  const response = await fetchWithTimeout(
    absUrl,
    {
      method: "GET",
      mode: "cors",
      credentials: sameOrigin ? "include" : "omit",
      cache: "no-store",
    },
    6000000
  );
  if (!response.ok) {
    throw new Error(`HTTP ${response.status} ${response.statusText}`);
  }
  const blob = await response.blob();
  const type =
    blob.type && blob.type !== "application/octet-stream" ? blob.type : "application/pdf";
  const name = pickFilenameFromHttpResponse(absUrl, response.headers);
  const file = new File([blob], name, { type });
  await saveFileToIndexedDB(file, true);
}

async function loadPdfFromDrive(fileId) {
  const gasResponse = await fetchWithTimeout(
    `${gasUrl}?fileId=${encodeURIComponent(fileId)}`,
    { method: "GET", mode: "cors" },
    6000000
  );
  if (!gasResponse.ok) {
    throw new Error(`GAS 代理回應錯誤: ${gasResponse.status} ${gasResponse.statusText}`);
  }
  const rawText = await gasResponse.text();
  let data;
  try {
    data = JSON.parse(rawText);
  } catch {
    throw new Error(
      "GAS 回應不是 JSON（常見原因：尚未登入、部署權限非「任何人」，或 fileId 錯誤而回傳 HTML 登入頁）。"
    );
  }
  if (!data.fileContent || !data.fileName) {
    throw new Error("GAS 代理返回的資料格式錯誤（缺少 fileContent 或 fileName）");
  }
  const fileContent = atob(data.fileContent);
  const arrayBuffer = new ArrayBuffer(fileContent.length);
  const uint8Array = new Uint8Array(arrayBuffer);
  for (let i = 0; i < fileContent.length; i++) {
    uint8Array[i] = fileContent.charCodeAt(i);
  }
  const blob = new Blob([arrayBuffer], { type: "application/pdf" });
  const file = new File([blob], data.fileName, { type: "application/pdf" });
  await saveFileToIndexedDB(file, true);
}

await (async function bootstrap() {
  let fileParam;
  try {
    const params = new URLSearchParams(document.location.search);
    fileParam = (params.get("file") || "").trim();
    if (!fileParam || fileParam === IDB_LAST_FILE_URL_MARKER) {
      return;
    }
    if (fileParam.startsWith("http://") || fileParam.startsWith("https://")) {
      try {
        await loadHttpUrlIntoIndexedDB(fileParam);
        replaceFileQueryWithIdbMarker();
        console.log("[viewer-remote-pdf-bootstrap] http(s) PDF 已寫入 IndexedDB，即將重整後以 idb:// 開啟");
        reloadOnceAfterCloudIdbWrite();
        return;
      } catch (err) {
        console.warn(
          "[viewer-remote-pdf-bootstrap] 無法先下載至 IndexedDB（常見：伺服器未允許 CORS），改由 PDF.js 直接開啟 URL。",
          err
        );
      }
      return;
    }

    if (fileParam.toLowerCase().endsWith(".pdf")) {
      const saved = await readFileFromIndexedDB();
      if (saved?.data && saved.name === fileParam) {
        await patchLastOpenedMarkSourceCloud(false);
        replaceFileQueryWithIdbMarker();
      }
      return;
    }

    // 與本機／重整一致：僅使用 viewer.html 的 #initialRenderCover（載入 PDF 中…），不用 Swal 白窗
    await loadPdfFromDrive(fileParam);
    replaceFileQueryWithIdbMarker();
    console.log("[viewer-remote-pdf-bootstrap] 雲端硬碟 PDF 已寫入 IndexedDB，即將重整後以 idb:// 開啟");
    reloadOnceAfterCloudIdbWrite();
    return;
  } catch (e) {
    console.error("viewer-remote-pdf-bootstrap（比照雲端硬碟 GAS）:", fileParam, e);
    closeSwalIfAny();
    try {
      const u = new URL(window.location.href);
      const f = (u.searchParams.get("file") || "").trim();
      if (
        f &&
        f !== IDB_LAST_FILE_URL_MARKER &&
        !f.startsWith("http://") &&
        !f.startsWith("https://") &&
        !f.toLowerCase().endsWith(".pdf")
      ) {
        u.searchParams.delete("file");
        try {
          localStorage.removeItem("pdfjs.lastFile");
        } catch (_) {
          /* ignore */
        }
        history.replaceState({}, "", u.toString());
      }
    } catch (_) {
      /* ignore */
    }
    const msg =
      (e && e.message) ||
      String(e) ||
      "未知錯誤";
    const detail =
      "無法從 Google 雲端硬碟下載檔案。\n\n" +
      "可能的原因：\n" +
      "1. 檔案未公開分享（請確認檔案分享設定為「知道連結的使用者」）\n" +
      "2. GAS 代理服務異常\n" +
      "3. 網路連線問題\n\n" +
      "建議：\n" +
      "- 確認檔案已公開分享\n" +
      "- 檢查網路連線\n" +
      "- 稍後再試\n\n" +
      "技術訊息：" +
      msg;
    await ensureSwal();
    if (window.Swal) {
      window.Swal.fire({
        icon: "error",
        title: "下載失敗",
        text: detail,
      });
    } else {
      alert("無法從 Google 雲端硬碟經 GAS 載入 PDF：\n\n" + msg);
    }
  }
})();
