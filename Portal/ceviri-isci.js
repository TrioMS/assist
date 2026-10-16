/* Assist çevrimdışı çeviri işçisi (16.10.2026) — Mozilla Bergamot (Firefox Translations motoru, MPL-2.0).
   Çeviri TAMAMEN bu bilgisayarda yapılır. GÜVENLİK: sunucu bu dosyayı ağ izni olmayan CSP ile verir
   (default-src 'none'; connect-src yok) — işçi hiçbir adrese (Assist dahil) istek atamaz. Motor ve
   model baytları sayfadan postMessage ile gelir. Firefox'taki kullanım: translations-engine.worker.js.
   İletiler:  {tur:'baslat', wasm}  → {tur:'hazir'}
              {tur:'model', cift:'entr', dosyalar:{model, lex, vocab}}  → {tur:'model-tamam', cift}
              {tur:'cevir', id, cift:'entr'|['tren','ende'], metinler:[...]}  → {tur:'sonuc', id, metinler}
              {tur:'ag-dene'}  → {tur:'ag-sonuc', engellendi}  (ağ kilidinin denetimi)  */
'use strict';
importScripts('lib/bergamot/bergamot-translator.js');
const HIZA = { model: 256, lex: 64, vocab: 64, qualityModel: 64, srcvocab: 64, trgvocab: 64 };
let B = null, servis = null;
const modeller = {};

function wasmBaslat(wasmBinary) {
  return new Promise((coz, red) => {
    const m = loadBergamot({
      INITIAL_MEMORY: 234291200, print: () => {}, printErr: () => {},
      onAbort() { red(new Error('Çeviri motoru yüklenemedi.')); },
      onRuntimeInitialized: async () => { await Promise.resolve(); coz(m); },
      wasmBinary
    });
  });
}
function ayar(o) { const g = '            '; return '\n' + Object.entries(o).map(([k, v]) => `${g}${k}: ${v}\n`).join('') + g; }
function modelKur(cift, dosyalar) {
  const bellek = {};
  for (const [tur, buf] of Object.entries(dosyalar)) {
    const a = new B.AlignedMemory(buf.byteLength, HIZA[tur] || 64);
    a.getByteArrayView().set(new Uint8Array(buf)); bellek[tur] = a;
  }
  const sozluk = new B.AlignedMemoryList(); sozluk.push_back(bellek.vocab);
  const yapilandirma = ayar({ 'beam-size': '1', normalize: '1.0', 'word-penalty': '0', 'max-length-break': '128', 'mini-batch-words': '1024',
    workspace: '128', 'max-length-factor': '2.0', 'skip-cost': 'true', 'cpu-threads': '0', quiet: 'true', 'quiet-translation': 'true',
    'gemm-precision': 'int8shiftAlphaAll', alignment: 'soft' });
  modeller[cift] = new B.TranslationModel(cift.slice(0, 2), cift.slice(2), yapilandirma, bellek.model, bellek.lex || null, sozluk, null);
}
function cevir(cift, metinler) {
  const mesaj = new B.VectorString(), secenek = new B.VectorResponseOptions();
  const dolu = metinler.map(t => !!(t && t.trim()));
  metinler.forEach((t, i) => { if (dolu[i]) { mesaj.push_back(t); secenek.push_back({ qualityScores: false, alignment: false, html: false }); } });
  let yanit;
  try {
    if (!mesaj.size()) return metinler.map(() => '');
    yanit = Array.isArray(cift) ? servis.translateViaPivoting(modeller[cift[0]], modeller[cift[1]], mesaj, secenek) : servis.translate(modeller[cift], mesaj, secenek);
    let j = 0;
    return metinler.map((t, i) => dolu[i] ? yanit.get(j++).getTranslatedText() : t);
  } finally { mesaj.delete(); secenek.delete(); if (yanit) yanit.delete(); }
}
onmessage = async e => {
  const d = e.data || {};
  try {
    if (d.tur === 'baslat') { if (!B) { B = await wasmBaslat(d.wasm); servis = new B.BlockingService({ cacheSize: 0 }); } postMessage({ tur: 'hazir' }); }
    else if (d.tur === 'model') { if (!modeller[d.cift]) modelKur(d.cift, d.dosyalar); postMessage({ tur: 'model-tamam', cift: d.cift }); }
    else if (d.tur === 'cevir') postMessage({ tur: 'sonuc', id: d.id, metinler: cevir(d.cift, d.metinler) });
    else if (d.tur === 'ag-dene') {
      let engellendi = false;
      try { await fetch('/__saglik', { cache: 'no-store' }); } catch { engellendi = true; }
      postMessage({ tur: 'ag-sonuc', engellendi });
    }
  } catch (er) {   // hata, bekleyen isteğin anahtarıyla döner (çeviri: id; kurulum/model: yanıt türü)
    postMessage({ tur: 'hata', id: d.id || { baslat: 'hazir', model: 'model-tamam', 'ag-dene': 'ag-sonuc' }[d.tur], hata: String(er && er.message || er) });
  }
};
