// Luftprovet: ryms Word-filen i Google Dokument, utan Google? (niclasfohlin.se 2026-10-19, K-456; Niclas: "Borde inte
// du kunna bygga nu en enkel testmodul för Google testet som är mycket snabbare? ... En testmodul du och riggen kan
// dela?")
//
// Exakt hur Google ritar går inte att räkna ut: var en rad bryts avgörs av bråkdelar av en punkt i en sidmotor som inte
// är öppen. Men det som stoppar en fil i jämförelsen (jamforWordGoogle) är nästan alltid ett block som inte ryms: ett
// stycke från ett ställe där Word-filen själv börjar en ny sida (sidankaret) till nästa, som tar fler sidor i Google än i
// Word, så att nästa ankare inte börjar sin sida och en sida kan bli tom. Det händer bara när blockets sista sida är så
// full i Word att det Google lägger till inte ryms. Reglerna (src/regler.js) tar resten: typsnitten, tecknen, ordens
// bredd och formerna som Google ritar annorlunda.
//
// Provet läser Words pdf (som den som anropar ofta redan har) och Word-filen. Varje sida ritas i 36 dpi i gråskala
// (pdftoppm), och innehållets nederkant är den lägsta raden med mörka punkter ovanför sidfoten. Textytans nederkant är
// sidans höjd minus den nedre marginalen ur Word-filens sektioner, för den sektion vars sidfotsavstånd stämmer med var
// sidfoten står på sidan (en fil kan ha flera marginaler på samma sidstorlek). Luften är skillnaden. Ett block är trångt
// när luften på dess sista sida, minus det Google lägger till, är mindre än marginalen.
//
// Det Google lägger till vid ett blocks slut (driften) mättes 2026-10-19 i niclasfohlin.se:s 63 Word-filer från samma
// dag, 1 742 block: hälften högst 4 pt, 99 procent högst 20 pt, och som mest 32 pt för ett block på en sida, 50 för två
// och 18 för fler. Inget block hade mindre luft än sin drift, och den minsta marginalen var 2 pt. Men metodriggen har sett
// större: lektionskort med 40 och 49 pt luft blev två sidor i Google, och en lärarhandledning med 36 pt var på väg över
// (riggens TILL-SAJTEN samma dag). De felen var lagade innan sajtens par gjordes, så de syns inte i mätningen. Utan annat
// underlag räknas driften därför som 50 pt för ett block på en eller två sidor och 24 för längre (STANDARDDRIFT), så att
// ett block behöver 56 pt luft. Efter varje jämförelse med Google sparar jamforWordGoogle blockens uppmätta drift i
// <ut>/<namn>-drift.json, och nästa luftprov av samma fil använder den för block med samma ankare och längd, när
// blockets luft är högst LIKA_LUFT pt annorlunda än när driften mättes. Har innehållet växt, som med fler rader i en
// tabell, där Google avrundar varje rad, räknas standarddriften igen. Så blir provet säkrare för varje jämförelse, och en
// fil som ändras lite går igenom utan Google.
import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import JSZip from 'jszip';

export const DPI = 36;
const PT = 72 / DPI;
// En punkt är mörk under det här gråvärdet (0 svart, 255 vitt): text, linjer och bilder, inte en ljus bakgrundston.
const MORK = 190;
/** Driften för ett block utan uppmätt drift, efter blockets längd i sidor (2026-10-19, se ovan). */
export const STANDARDDRIFT = { 1: 50, 2: 50, fler: 24 };
/** Så mycket får blockets luft skilja från när driften mättes, i punkter, för att den uppmätta driften ska gälla. */
export const LIKA_LUFT = 12;
/** Så mycket luft ska finnas kvar när driften är dragen, i punkter. */
export const MARGINAL = 6;
export const standarddrift = (k) => STANDARDDRIFT[k] ?? STANDARDDRIFT.fler;

const avkoda = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** Word-filens sektioner: sidans bredd och höjd, den nedre marginalen och sidfotens avstånd, i punkter. */
export async function sektioner(docx) {
  const xml = await (await JSZip.loadAsync(docx)).file('word/document.xml').async('string');
  const ut = [];
  for (const m of xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)) {
    const t = (re) => Number(m[0].match(re)?.[1] ?? NaN) / 20;
    ut.push({ b: t(/<w:pgSz\b[^>]*\bw:w="(\d+)"/), h: t(/<w:pgSz\b[^>]*\bw:h="(\d+)"/), botten: t(/<w:pgMar\b[^>]*\bw:bottom="(-?\d+)"/), fot: t(/<w:pgMar\b[^>]*\bw:footer="(\d+)"/) || 0 });
  }
  return ut;
}

// Sidorna i pdf:en som gråskala: för varje rad antalet mörka punkter.
function sidrader(pdf) {
  const tmp = mkdtempSync(join(tmpdir(), 'wordparitet-luft-'));
  try {
    execFileSync('pdftoppm', ['-r', String(DPI), '-gray', pdf, join(tmp, 's')], { maxBuffer: 1 << 28 });
    const nr = (f) => Number(f.match(/-(\d+)\.pgm$/)?.[1] ?? 0);
    return readdirSync(tmp).filter((f) => f.endsWith('.pgm')).sort((a, b) => nr(a) - nr(b)).map((f) => {
      const buf = readFileSync(join(tmp, f));
      // P5, bredd, höjd, största värdet, ett blanktecken, sedan punkterna.
      let i = 0;
      const tok = [];
      while (tok.length < 4) {
        while (/\s/.test(String.fromCharCode(buf[i]))) i++;
        let s = '';
        while (!/\s/.test(String.fromCharCode(buf[i]))) s += String.fromCharCode(buf[i++]);
        tok.push(s);
      }
      i++;
      const b = Number(tok[1]);
      const h = Number(tok[2]);
      const rader = new Array(h);
      for (let y = 0; y < h; y++) {
        let n = 0;
        const o = i + y * b;
        for (let x = 0; x < b; x++) if (buf[o + x] < MORK) n++;
        rader[y] = n;
      }
      return { b, h, rader };
    });
  } finally {
    rmSync(tmp, { recursive: true, force: true });
  }
}

/**
 * Sidornas mått i en pdf av Word-filen: innehållets nederkant, textytans nederkant och luften mellan dem, i punkter
 * uppifrån, och om sidan är tom (inget mörkt i textytan: en sida med bara en bild är inte tom). sekt är Word-filens
 * sektioner (sektioner()). Sidfoten är den nedersta klungan av mörka rader, under ett glapp på minst 8 pt, och
 * sektionen är den med sidans storlek vars sidfotsavstånd ligger närmast; en sida utan sidfot får sektionen utan sidfot.
 */
export function sidmatt(pdf, sekt) {
  return sidrader(pdf).map((s) => {
    const H = s.h * PT;
    const B = s.b * PT;
    const mojliga = sekt.filter((x) => Math.abs(x.h - H) < 4 && Math.abs(x.b - B) < 4);
    let lagst = s.h - 1;
    while (lagst > 0 && !s.rader[lagst]) lagst--;
    let glapp = 0;
    let topp = lagst;
    for (let y = lagst; y > s.h * 0.85; y--) {
      if (s.rader[y]) { topp = y; glapp = 0; } else if (++glapp >= 4) break;
    }
    const fotAvstand = H - (lagst + 1) * PT;
    const medFot = mojliga.filter((x) => x.fot > 0);
    let val = medFot.length ? medFot.reduce((a, x) => (Math.abs(x.fot - fotAvstand) < Math.abs(a.fot - fotAvstand) ? x : a)) : undefined;
    const utanFot = mojliga.find((x) => x.fot === 0);
    if (utanFot && (!val || Math.abs(val.fot - fotAvstand) > 8)) val = utanFot;
    val ??= mojliga[0] ?? { botten: 56.7, fot: 28.35 };
    let c = val.fot > 0 ? topp - 1 : lagst;
    while (c > 0 && !s.rader[c]) c--;
    const botten = (c + 1) * PT;
    return { hojd: H, botten, textBotten: H - val.botten, luft: H - val.botten - botten, tom: c < s.h * 0.12 };
  });
}

/**
 * Blocken i Word-filen: från varje sidankare till nästa, med ankarets text, blockets första sida (från 0), antal sidor
 * och luften på dess sista sida. ankare och iWord kommer ur jämförelsens sidankare och ankarsidor (src/jamfor.js), och
 * matt ur sidmatt().
 */
export function block(ankare, iWord, matt) {
  const nasta = (i) => iWord.slice(i + 1).find((x) => x >= 0) ?? matt.length;
  const ut = [];
  ankare.forEach((a, i) => {
    if (iWord[i] < 0) return;
    const k = nasta(i) - iWord[i];
    const sista = matt[iWord[i] + k - 1];
    if (k <= 0 || !sista) return;
    ut.push({ ankare: a, fran: iWord[i], k, luft: Math.round(sista.luft) });
  });
  return ut;
}

/**
 * Är blocken trånga? drift är förra jämförelsens uppmätta drift ({ ankare, k, luft, drift }, ur <namn>-drift.json); ett
 * block med samma ankare och längd och nästan samma luft (LIKA_LUFT) får sin uppmätta drift, de andra STANDARDDRIFT.
 * Svarar { lage: 'ryms' | 'google', tranga } där tranga är de trånga blocken, med för lite luft, med driften de räknades
 * med och varifrån den kom.
 */
export function bedom(blocken, drift = [], { marginal = MARGINAL } = {}) {
  const kand = new Map(drift.map((d) => [`${d.ankare}|${d.k}`, d]));
  const tranga = [];
  for (const b of blocken) {
    const forut = kand.get(`${b.ankare}|${b.k}`);
    const matt = forut && (forut.luft === undefined || Math.abs(forut.luft - b.luft) <= LIKA_LUFT) ? forut.drift : undefined;
    const d = matt ?? standarddrift(b.k);
    if (b.luft - (matt === undefined ? Math.max(d, 0) : d) < marginal) tranga.push({ ...b, drift: d, kalla: matt === undefined ? 'standard' : 'uppmätt' });
  }
  return { lage: tranga.length ? 'google' : 'ryms', tranga };
}
