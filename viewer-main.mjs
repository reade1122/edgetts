/**
 * 單一進入點：強制在 PDF.js viewer 啟動「之前」跑完雲端 GAS bootstrap 的 top-level await。
 * 多個並列的 <script type="module"> 在部分瀏覽器不會等待彼此，會導致 ?file=<DriveId> 尚未寫入 IndexedDB 就開檔。
 * 雲端寫入 IDB 後 bootstrap 會 location.reload()：須跳過本輪後續 import，以免與卸載競態。
 */
await import("./viewer-remote-pdf-bootstrap.mjs");
if (globalThis.__PDFJS_RELOADING_FOR_CLOUD_IDB__) {
  /* 即將重整，不載入 pdf.mjs / viewer.mjs */
} else {
  await import("./build/pdf.mjs");
  await import("./viewer.mjs");
}
