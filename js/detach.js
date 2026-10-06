/* global Office */
// Fenêtre détachée de la liseuse (déplaçable sur un second écran).
// Excel : boîte de dialogue Office (DialogApi 1.2 : messageChild). Hors Office (tests) : window.open + postMessage.
// Les PDF sont transmis à la fenêtre par morceaux (la fenêtre n'a pas accès au classeur).
const CHUNK = 30000;

export function supported() {
  if (window.__MOCK__) return true;
  try { return !!(Office.context.requirements && Office.context.requirements.isSetSupported("DialogApi", "1.2")); } catch (e) { return false; }
}

export function open({ onMessage, onClosed, getB64 }) {
  const url = new URL("reader.html?v=10", location.href).href;
  return new Promise((resolve, reject) => {
    let sendRaw, closeRaw, closed = false;
    const handle = m => {
      if (m.t === "need") streamFile(m.name);
      else onMessage(m);
    };
    async function streamFile(name) {
      const b64 = await getB64(name);
      if (!b64) { sendRaw({ t: "missing", name }); return; }
      const n = Math.ceil(b64.length / CHUNK) || 1;
      for (let i = 0; i < n; i++) { if (closed) return; sendRaw({ t: "chunk", name, i, n, data: b64.slice(i * CHUNK, (i + 1) * CHUNK) }); if (i % 8 === 7) await new Promise(r => setTimeout(r, 0)); }
    }
    const finish = () => { if (closed) return; closed = true; onClosed && onClosed(); };
    const api = { send: m => !closed && sendRaw(m), close: () => { try { closeRaw(); } catch (e) {} finish(); }, get closed() { return closed; } };

    if (window.__MOCK__) {
      const w = window.open(url, "liseuse", "width=1400,height=900");
      if (!w) return reject(new Error("fenêtre bloquée par le navigateur"));
      window.__reader = w;
      sendRaw = m => w.postMessage(m, "*"); closeRaw = () => w.close();
      window.addEventListener("message", e => { if (e.source === w) handle(e.data); });
      const iv = setInterval(() => { if (w.closed) { clearInterval(iv); finish(); } }, 500);
      const ready = e => { if (e.source === w && e.data && e.data.t === "ready") { window.removeEventListener("message", ready); resolve(api); } };
      window.addEventListener("message", ready);
      return;
    }
    if (!supported()) return reject(new Error("cette version d'Excel ne permet pas de détacher la liseuse (DialogApi 1.2 requise : Microsoft 365 ou Excel 2021 et suivants)"));
    Office.context.ui.displayDialogAsync(url, { height: 85, width: 80, displayInIframe: false, promptBeforeOpen: false }, res => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(res.error ? res.error.message : "ouverture refusée"));
      const dlg = res.value; let resolved = false;
      sendRaw = m => dlg.messageChild(JSON.stringify(m)); closeRaw = () => dlg.close();
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, a => {
        let m; try { m = JSON.parse(a.message); } catch (e) { return; }
        if (m.t === "ready" && !resolved) { resolved = true; resolve(api); return; }
        handle(m);
      });
      dlg.addEventHandler(Office.EventType.DialogEventReceived, () => finish());
      setTimeout(() => { if (!resolved) { resolved = true; resolve(api); } }, 8000);
    });
  });
}
