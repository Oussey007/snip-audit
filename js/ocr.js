// OCR embarqué (Tesseract.js + français), utilisé uniquement quand la zone encadrée ne contient pas de texte (scan).
let workerP = null;
function loadScript(src) {
  return new Promise((res, rej) => {
    if (window.Tesseract) return res();
    const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => rej(new Error("Chargement OCR impossible"));
    document.head.appendChild(s);
  });
}
async function worker() {
  if (!workerP) workerP = (async () => {
    const base = new URL("../vendor/tesseract/", import.meta.url).href;
    await loadScript(base + "tesseract.min.js");
    return await window.Tesseract.createWorker("fra", 1, {
      workerPath: base + "worker.min.js",
      corePath: base + "core/",
      langPath: base + "lang/",
      gzip: true,
      cacheMethod: "none",
      logger: () => {}
    });
  })();
  return workerP;
}
export async function ocrCanvas(canvas) {
  const w = await worker();
  const { data } = await w.recognize(canvas);
  return (data.text || "").replace(/[ \t]+/g, " ").trim();
}
