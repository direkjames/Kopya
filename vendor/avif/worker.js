// Runs the AVIF encoder off the main thread so the page stays responsive.
// A fresh worker is used for every file and thrown away afterwards, which
// hands all of the encoder's memory back to the browser.
import createModule from './avif_enc.js';

let ready;

self.onmessage = async (e) => {
  const { buffer, width, height, options } = e.data;
  try {
    ready = ready || createModule({ noInitialRun: true });
    const mod = await ready;
    const out = mod.encode(new Uint8Array(buffer), width, height, options);
    if (!out) throw new Error('The AVIF encoder could not process this image.');
    self.postMessage({ ok: true, bytes: out.buffer }, [out.buffer]);
  } catch (err) {
    self.postMessage({ ok: false, message: String((err && err.message) || err) });
  }
};
