// Jämförelserna: ritar Google Dokument en Word-fil som Word gör (jamforWordGoogle), och ändrar en ändring i Word-koden
// hur Word-filerna ser ut i Word (wordFacit och wordJamfor). Båda läser pdf:erna med Popplers pdftotext och pdftoppm
// (de ska finnas i PATH) och ritar bilder med sharp, som den som anropar har installerat.
//
// jamforWordGoogle gör för en Word-fil två pdf:er: Words egen (src/word.js), sparad efter filens innehåll så att en
// oförändrad fil inte görs om, och Google Dokuments (src/google.js). Det här stoppar (NEJ):
//
//   blocken      varje ställe där Word-filen själv börjar en ny sida (sektion, sidbrytning före) börjar en sida också i
//                Google, och blocket fram till nästa tar inte fler sidor i Google än i Word (sidankare)
//   tomma sidor  ingen sida är tom i Google när Words sida med samma nummer har text
//   texten       inget tecken i Word saknas i Google; ekvationerna räknas inte, Google ritar dem utan text i pdf:en
//   ordbrytning  inget ord bryts mitt i, utan bindestreck, i någon av dem ("Personbeskrivni" och "ng")
//
// Skillnader i den fria texten inne i ett block blir obs, med första sidan som skiljer sig: Word håller ihop en tabell
// vars rader har "håll ihop med nästa", Google delar den, så Google kan få färre sidor med samma innehåll. Ett
// översiktsark med Words sidor över Googles läses sida för sida.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { basename, join } from 'node:path';
import JSZip from 'jszip';
import { googlePdf } from './google.js';
import { wordPdf, wordPdfSync } from './word.js';

const sharpen = async () => (await import('sharp')).default;
const sha = (b) => createHash('sha256').update(b).digest('hex');
// Gränsen för en sida: under 90 procent av bokstäverna gemensamma skiljer sig sidan.
const SIDLIKHET = 0.9;
// Sidnumret står som "Sida 12 av 31" och på boksidorna som "s. 42" sist i sidfoten; bara numret skiljer när Word och
// Google bryter sidorna olika.
export const SIDNUMMER = /Sida \d+ av \d+|(?<= s\.) \d+(?=\s*$)/gm;

/**
 * Word-filens innehåll utan det som skiljer två byggen av samma fil åt: byggets tidsstämpel (docProps/core.xml),
 * nyckeln som det inbäddade typsnittet är förvrängt med (fontTable.xml, word/fonts/) och bildernas byte. SVG jämförs med
 * LF och PNG och JPEG efter pixlarna, eftersom en bild får andra byte i en annan miljö (2026-09-30: 13 av 49 filer
 * skilde sig annars mellan byggen på Windows och Linux av samma kod). Svarar en kontrollsumma.
 */
export async function innehall(buf) {
  const sharp = await sharpen();
  const zip = await JSZip.loadAsync(buf);
  const delar = [];
  const bilder = [];
  for (const n of Object.keys(zip.files).sort()) {
    if (zip.files[n].dir || /^docProps\/core\.xml$|^word\/fontTable\.xml$|^word\/fonts\/|\.rels$/.test(n)) continue;
    const data = await zip.file(n).async('nodebuffer');
    if (/^word\/media\/.+\.(png|jpe?g)$/i.test(n)) {
      const { data: px, info } = await sharp(data).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
      bilder.push(`${info.width}x${info.height}:${sha(px)}`);
    } else if (n.startsWith('word/media/')) bilder.push(sha(Buffer.from(data.toString('latin1').replace(/\r\n/g, '\n'), 'latin1')));
    else delar.push(`${n}:${sha(data)}`);
  }
  return sha([...delar, ...bilder.sort()].join('\n'));
}

const ord = (text) => text.normalize('NFC').toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
const antal = (lista) => lista.reduce((m, o) => m.set(o, (m.get(o) ?? 0) + 1), new Map());
const bokstaver = (text) => antal(text.normalize('NFC').toLowerCase().match(/\p{L}/gu) ?? []);
const tecken = (text) => antal(text.normalize('NFC').toLowerCase().match(/[\p{L}\p{N}]/gu) ?? []);
const avkoda = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

// Texten i Word-filen: orden, ekvationernas tecken, tabellernas rubrikrader och sidhuvudets och sidfotens stycken.
async function wordText(buf) {
  const zip = await JSZip.loadAsync(buf);
  let text = '';
  let ekvationer = '';
  const rubrikrader = new Set();
  const sidhuvud = new Set();
  for (const d of Object.keys(zip.files).filter((n) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(n))) {
    const xml = await zip.file(d).async('string');
    for (const m of xml.matchAll(/<m:oMath\b[\s\S]*?<\/m:oMath>/g)) ekvationer += ` ${[...m[0].matchAll(/<m:t(?:\s[^>]*)?>([^<]*)<\/m:t>/g)].map((t) => t[1]).join(' ')}`;
    // Det som upprepas på varje sida och därför står olika många gånger när Word och Google bryter sidorna olika:
    // cellerna i tabellernas rubrikrader och styckena i sidhuvud och sidfot. Varje cell och stycke räknas för sig, och
    // sidhuvudets och sidfotens stycken delas vid tabben (adressen och, efter en tabb, Sida N av M).
    const texten = (x) => avkoda((x.match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, '')).join('')).normalize('NFC').toLowerCase();
    const delar = /^word\/(header|footer)/.test(d) ? xml.split(/<\/w:p>|<w:tab\/>/) : [...xml.matchAll(/<w:tr><w:trPr>(?:(?!<\/w:trPr>).)*<w:tblHeader\b(?:(?!<\/w:trPr>).)*<\/w:trPr>((?:(?!<\/w:tr>).)*)<\/w:tr>/gs)].flatMap((m) => m[1].split(/<\/w:tc>/));
    for (const x of delar) {
      const t = texten(x);
      const b = (t.match(/\p{L}/gu) ?? []).join('');
      if (b.length >= 4) rubrikrader.add(`${b}|${(t.match(/\p{N}/gu) ?? []).join('')}`);
      if (b.length >= 4 && /^word\/(header|footer)/.test(d) && !/<w:instrText|<w:fldSimple/.test(x)) sidhuvud.add((t.match(/[\p{L}\p{N}]/gu) ?? []).join(''));
    }
    text += ' ' + xml.replace(/<m:oMath\b[\s\S]*?<\/m:oMath>/g, ' ').replace(/<w:instrText\b[^>]*>[^<]*<\/w:instrText>/g, ' ')
      .replace(/<\/w:p>|<\/w:tc>|<w:tab\/>|<w:br\/>/g, ' ').replace(/<w:noBreakHyphen\/>/g, '-')
      .replace(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g, '$1').replace(/<[^>]+>/g, '');
  }
  return {
    ord: new Set(ord(avkoda(text))),
    ekvationer: tecken(avkoda(ekvationer)),
    rubrikrader: [...rubrikrader].map((r) => { const [bokst, siffror] = r.split('|'); return { bokst, siffror }; }),
    sidhuvud: [...sidhuvud].sort((a, b) => b.length - a.length),
  };
}
const forekomster = (bokst, rad) => { let n = 0; for (let i = bokst.indexOf(rad); i >= 0; i = bokst.indexOf(rad, i + rad.length)) n++; return n; };
function likhet(a, b) {
  let gemensamt = 0;
  let summa = 0;
  for (const [c, n] of a) { gemensamt += Math.min(n, b.get(c) ?? 0); summa += n; }
  for (const [, n] of b) summa += n;
  return summa ? (2 * gemensamt) / summa : 1;
}
// Ett ord som bryts mitt i utan bindestreck: en del som inte är ett ord i filen men blir det med nästa del. Bara ord på
// minst sex bokstäver, som är de som bryts i en smal kolumn.
function brutnaOrd(sidorna, iFilen) {
  const t = ord(sidorna.join(' '));
  const brutna = new Set();
  for (let i = 0; i < t.length - 1; i++) {
    const hel = t[i] + t[i + 1];
    if (hel.length >= 6 && t[i].length > 1 && !iFilen.has(t[i]) && iFilen.has(hel)) brutna.add(hel);
  }
  return brutna;
}
// Sidankarna: de första bokstäverna efter varje ställe där Word-filen själv börjar en ny sida, en ny sektion eller ett
// stycke med sidbrytning före. Ankaret tar bara text fram till nästa sådant ställe.
async function sidankare(buf) {
  const xml = await (await JSZip.loadAsync(buf)).file('word/document.xml').async('string');
  const starter = [0];
  for (const m of xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr><\/w:pPr><\/w:p>/g)) starter.push(m.index + m[0].length);
  for (const m of xml.matchAll(/<w:p><w:pPr>(?:(?!<\/w:pPr>).)*<w:pageBreakBefore\/>/gs)) starter.push(m.index);
  const sorterade = [...new Set(starter)].sort((a, b) => a - b);
  return sorterade.map((i, k) => {
    const text = avkoda((xml.slice(i, Math.min(sorterade[k + 1] ?? Infinity, i + 30000)).match(/<w:t(?:\s[^>]*)?>[^<]*<\/w:t>/g) ?? []).map((t) => t.replace(/<[^>]+>/g, '')).join(''));
    return (text.normalize('NFC').toLowerCase().match(/[\p{L}\p{N}]/gu) ?? []).join('').slice(0, 48);
  }).filter((a) => a.length >= 8);
}
// Står samma text överst på flera sidor väljs i Google den sida som ligger närmast där ankaret borde stå. Ett ankare
// står överst på en sida när sidans första tecken innehåller det, eller nästan alla dess tecken i en annan ordning.
function ankarsidor(bokstaverna, ankare, iWord) {
  const bokst = bokstaverna.map((b) => b.slice(0, 130));
  const nastanAlla = (topp, a) => {
    const kvar = antal([...topp]);
    let traff = 0;
    for (const c of a) if ((kvar.get(c) ?? 0) > 0) { kvar.set(c, kvar.get(c) - 1); traff++; }
    return traff >= 0.95 * a.length;
  };
  const ut = [];
  let p = 0;
  let forra = -1;
  ankare.forEach((a, i) => {
    const kandidater = [];
    for (let s = p + (i > 0 && a === ankare[i - 1] ? 1 : 0); s < bokst.length; s++) if (bokst[s].includes(a) || nastanAlla(bokst[s].slice(0, a.length + 20), a)) kandidater.push(s);
    let hittad = kandidater[0] ?? -1;
    if (iWord && iWord[i] >= 0 && kandidater.length > 1) {
      const vantat = forra >= 0 ? ut[forra] + (iWord[i] - iWord[forra]) : iWord[i];
      hittad = kandidater.reduce((b, s) => (Math.abs(s - vantat) < Math.abs(b - vantat) ? s : b));
    }
    ut.push(hittad);
    if (hittad >= 0) { p = hittad; if (!iWord || iWord[i] >= 0) forra = i; }
  });
  return ut;
}

// Översiktsarket: tio sidor i bredd, Words sidor överst och Googles under, par för par.
async function oversikt(wordPdfFil, googlePdfFil, ark, tmp) {
  const sharp = await sharpen();
  const bilder = async (pdf, prefix) => {
    const mapp = join(tmp, `.sidor-${prefix}`);
    rmSync(mapp, { recursive: true, force: true });
    mkdirSync(mapp, { recursive: true });
    execFileSync('pdftoppm', ['-r', '20', '-png', pdf, join(mapp, 's')]);
    const ut = await Promise.all(readdirSync(mapp).sort().map(async (f) => ({ data: readFileSync(join(mapp, f)), matt: await sharp(join(mapp, f)).metadata() })));
    rmSync(mapp, { recursive: true, force: true });
    return ut;
  };
  const [w, g] = [await bilder(wordPdfFil, 'w'), await bilder(googlePdfFil, 'g')];
  const alla = [...w, ...g];
  const cb = Math.max(...alla.map((b) => b.matt.width)) + 6;
  const ch = Math.max(...alla.map((b) => b.matt.height)) + 6;
  const k = 10;
  const par = Math.ceil(Math.max(w.length, g.length) / k);
  const lager = [];
  for (const [rad, lista] of [[0, w], [1, g]]) lista.forEach((b, i) => lager.push({ input: b.data, left: (i % k) * cb, top: Math.floor(i / k) * (2 * ch + 16) + rad * ch }));
  await sharp({ create: { width: k * cb, height: par * (2 * ch + 16), channels: 3, background: '#666666' } }).composite(lager).png().toFile(ark);
}

/**
 * Jämför hur Word och Google Dokument ritar Word-filen fil. ut är mappen för pdf:erna och översiktsarket; Words pdf
 * sparas efter innehållet i ut/word/. ram är det som står på varje sida och inte räknas som sidans text (sidnumren och
 * sidhuvudets och sidfotens rader), som ett reguljärt uttryck med flaggorna gm. Med igen jämförs den senast hämtade
 * Google-pdf:en en gång till, utan uppladdning, om den är gjord ur samma innehåll. sidor ger bilder av de sidorna.
 * Svarar { namn, lage: 'ok' | 'obs' | 'NEJ', problem, obs, beskrivning, ark, lank, provade, ankare, sidorWord, sidorGoogle, bilder }.
 */
export async function jamforWordGoogle(fil, { ut, mappId, inloggning, igen = false, ram = SIDNUMMER, sidnummer = SIDNUMMER, sidor = [], visaAnkare = false } = {}) {
  if (!ut) throw new Error('jamforWordGoogle: ange ut, mappen för pdf:erna.');
  mkdirSync(join(ut, 'word'), { recursive: true });
  const namn = basename(fil, '.docx');
  const docx = readFileSync(fil);
  const gPdf = join(ut, `${namn}-google.pdf`);
  const gNyckel = `${gPdf}.nyckel`;
  const nyckel = await innehall(docx);
  const ateranvand = igen && existsSync(gPdf) && existsSync(gNyckel) && readFileSync(gNyckel, 'utf8') === nyckel;
  // Words pdf efter innehållet, och äldre versioner av samma fil tas bort.
  const wMapp = join(ut, 'word');
  const wPdf = join(wMapp, `${namn}-${nyckel.slice(0, 16)}.pdf`);
  const word = existsSync(wPdf) ? Promise.resolve(wPdf) : (async () => {
    for (const f of readdirSync(wMapp)) if (f.startsWith(`${namn}-`) && /^[0-9a-f]{16}\.pdf$/.test(f.slice(namn.length + 1))) rmSync(join(wMapp, f));
    return wordPdf(fil, wPdf);
  })();
  // Word och Google arbetar samtidigt: Word på datorn, Google på nätet.
  const [, g] = await Promise.all([word, ateranvand ? { pdf: readFileSync(gPdf) } : googlePdf(docx, { namn, mappId, inloggning })]);
  writeFileSync(gPdf, g.pdf);
  writeFileSync(gNyckel, nyckel);
  const iFilen = await wordText(docx);
  const pdftext = (pdf, layout = true) => execFileSync('pdftotext', ['-enc', 'UTF-8', ...(layout ? ['-layout'] : []), pdf, '-'], { encoding: 'utf8', maxBuffer: 1 << 26 });
  const sidtexter = (pdf, layout) => pdftext(pdf, layout).split('\f').slice(0, -1).map((t) => t.replace(ram, ''));
  const ankartexter = (pdf) => pdftext(pdf).split('\f').slice(0, -1).map((t) => iFilen.sidhuvud.reduce((b, h) => b.split(h).join(''), (t.replace(sidnummer, '').normalize('NFC').toLowerCase().match(/[\p{L}\p{N}]/gu) ?? []).join('')));
  // Sidorna uppifrån och ned (-layout) för allt utom ordbrytningen, som läses i läsordning.
  const w = sidtexter(wPdf);
  const gs = sidtexter(gPdf);
  const problem = [];
  const obs = [];
  const ankare = await sidankare(docx);
  const pw = ankarsidor(ankartexter(wPdf), ankare);
  const pg = ankarsidor(ankartexter(gPdf), ankare, pw);
  const nasta = (p, i, slut) => p.slice(i + 1).find((x) => x >= 0) ?? slut;
  const utanSida = [];
  const langre = [];
  ankare.forEach((a, i) => {
    if (pw[i] < 0) return;
    if (pg[i] < 0) { utanSida.push(`"${a.slice(0, 16)}…" (Words sida ${pw[i] + 1})`); return; }
    const iWord = nasta(pw, i, w.length) - pw[i];
    const iGoogle = nasta(pg, i, gs.length) - pg[i];
    if (iGoogle > iWord) langre.push(`blocket från Words sida ${pw[i] + 1} tar ${iGoogle} sidor i Google mot ${iWord}`);
  });
  const ankarlista = visaAnkare ? ankare.map((a, i) => `ankare ${String(i + 1).padStart(2)}: Word s${pw[i] + 1}, Google s${pg[i] + 1}  ${a}`) : [];
  if (utanSida.length) problem.push(`börjar inte en sida i Google: ${utanSida.slice(0, 4).join(', ')}${utanSida.length > 4 ? ` och ${utanSida.length - 4} till` : ''}`);
  if (langre.length) problem.push(langre.slice(0, 4).join(', ') + (langre.length > 4 ? ` och ${langre.length - 4} till` : ''));
  const tomma = [];
  const olika = [];
  for (let s = 0; s < gs.length; s++) {
    const gb = bokstaver(gs[s]);
    const wb = bokstaver(w[s] ?? '');
    if (gs[s].trim().length < 3 && (w[s] === undefined || w[s].trim().length >= 3)) tomma.push(s + 1);
    else if (w[s] !== undefined && likhet(wb, gb) < SIDLIKHET) olika.push(`${s + 1} (${Math.round(likhet(wb, gb) * 100)} %)`);
  }
  if (tomma.length) problem.push(`tom sida i Google ${tomma.join(', ')}`);
  if (w.length !== gs.length) obs.push(`Word ${w.length} sidor, Google ${gs.length}`);
  if (olika.length) obs.push(`${w.length === gs.length ? 'annan text på sidan' : 'första sidan som skiljer sig:'} ${(w.length === gs.length ? olika : olika.slice(0, 1)).join(', ')}`);
  // Texten: bokstäver och siffror i Words pdf som saknas i Googles, utan sidnumren; det som upprepas per sida dras av så
  // många gånger som det står fler gånger i Word, och ekvationerna räknas inte.
  const [helW, helG] = [pdftext(wPdf), pdftext(gPdf)];
  const utanSidnummer = (t) => t.replace(sidnummer, (m) => (m.startsWith('Sida') ? 'Sida av' : ''));
  const iWord = tecken(utanSidnummer(helW));
  const iGoogle = tecken(utanSidnummer(helG));
  const bokstW = (helW.normalize('NFC').toLowerCase().match(/\p{L}/gu) ?? []).join('');
  const bokstG = (helG.normalize('NFC').toLowerCase().match(/\p{L}/gu) ?? []).join('');
  for (const rad of iFilen.rubrikrader) {
    const extra = forekomster(bokstW, rad.bokst) - forekomster(bokstG, rad.bokst);
    if (extra > 0) for (const c of rad.bokst + rad.siffror) iWord.set(c, (iWord.get(c) ?? 0) - extra);
  }
  const saknas = [...iWord].map(([c, n]) => [c, n - (iFilen.ekvationer.get(c) ?? 0) - (iGoogle.get(c) ?? 0)]).filter(([, d]) => d > 0);
  if (saknas.length) problem.push(`saknas i Google: ${saknas.slice(0, 12).map(([c, d]) => `${c}${d > 1 ? ` ×${d}` : ''}`).join(' ')}${saknas.length > 12 ? ` och ${saknas.length - 12} till` : ''}`);
  const iW = brutnaOrd(sidtexter(wPdf, false), iFilen.ord);
  const iG = brutnaOrd(sidtexter(gPdf, false), iFilen.ord);
  const brutna = [...new Set([...iW, ...iG])].map((o) => `${o} (${iW.has(o) && iG.has(o) ? 'Word och Google' : iW.has(o) ? 'Word' : 'Google'})`);
  if (brutna.length) problem.push(`bryts mitt i ordet: ${brutna.join(', ')}`);
  const ark = join(ut, `${namn}-oversikt.png`);
  await oversikt(wPdf, gPdf, ark, ut);
  const bilder = [];
  for (const s of sidor) {
    for (const [pdf, vem] of [[wPdf, 'word'], [gPdf, 'google']]) {
      const bas = join(ut, `${namn}-${vem}-s${s}`);
      execFileSync('pdftoppm', ['-r', '60', '-f', String(s), '-l', String(s), '-png', '-singlefile', pdf, bas]);
      bilder.push({ sida: s, vem, fil: `${bas}.png` });
    }
  }
  const lage = problem.length ? 'NEJ' : obs.length ? 'obs' : 'ok';
  const provade = pw.filter((p) => p >= 0).length;
  const beskrivning = [...problem, ...obs].join('; ') || `${w.length} sidor i Word och i Google, samma text på varje sida`;
  return { namn, lage, problem, obs, beskrivning, ark, lank: g.lank, provade, ankare: ankare.length, ankarlista, sidorWord: w.length, sidorGoogle: gs.length, bilder, wordPdf: wPdf, googlePdf: gPdf };
}

// ---------------------------------------------------------------- Word före och efter
// Sidans tecken sorterade, utan sidnumret och med kända byten: samma tecken på samma sida betyder att sidbrytningarna
// står kvar. Ordningen räknas inte, eftersom pdftotext kan läsa kortens rutnät i en annan ordning när något flyttar en
// bråkdel av en punkt.
const sidtecken = (pdf, byten) => execFileSync('pdftotext', ['-enc', 'UTF-8', pdf, '-'], { encoding: 'utf8', maxBuffer: 1 << 26 }).split('\f').slice(0, -1)
  .map((s) => byten.reduce((t, [fran, till]) => t.replaceAll(fran, till), s.replace(/Sida \d+ av \d+/g, '')))
  .map((s) => [...s.replace(/\s+/g, '')].sort().join(''));
const GRANS = 0.0005;

/** Words pdf av Word-filen i facitmappen, inför en ändring. Svarar antalet sidor. */
export function wordFacit(fil, facitMapp) {
  mkdirSync(facitMapp, { recursive: true });
  const pdf = join(facitMapp, basename(fil).replace(/\.docx$/, '.pdf'));
  wordPdfSync(fil, pdf);
  return sidtecken(pdf, []).length;
}

/**
 * Jämför Words pdf av Word-filen nu med facit, sida för sida: texten (samma antal sidor och samma tecken på varje sida
 * betyder att varje sidbrytning står kvar) och bilden (pdftoppm i 40 dpi; en sida där mer än 0,05 procent av punkterna
 * skiljer sig räknas som ändrad). De tre mest ändrade sidorna ritas som facit, nu och skillnaden i bildMapp. byten är
 * tecken som ändringen bytt med avsikt, [['☐', '□']]. Svarar { lage: 'ok' | 'NEJ' | '?', delar, sidor }.
 */
export async function wordJamfor(fil, { facitMapp, nuMapp, bildMapp, byten = [] }) {
  const sharp = await sharpen();
  const namn = basename(fil).replace(/\.docx$/, '');
  mkdirSync(nuMapp, { recursive: true });
  const pdf = join(nuMapp, `${namn}.pdf`);
  wordPdfSync(fil, pdf);
  const gammal = join(facitMapp, `${namn}.pdf`);
  if (!existsSync(gammal)) return { lage: '?', delar: ['inget facit'], sidor: sidtecken(pdf, byten).length };
  const [a, b] = [sidtecken(gammal, byten), sidtecken(pdf, byten)];
  const olika = a.map((s, i) => (s === b[i] ? 0 : i + 1)).filter(Boolean);
  const sidbilder = (p, mapp) => {
    rmSync(mapp, { recursive: true, force: true });
    mkdirSync(mapp, { recursive: true });
    execFileSync('pdftoppm', ['-r', '40', '-gray', '-png', p, join(mapp, 's')]);
    return readdirSync(mapp).sort().map((f) => join(mapp, f));
  };
  const ba = sidbilder(gammal, join(nuMapp, '.a'));
  const bb = sidbilder(pdf, join(nuMapp, '.b'));
  const res = [];
  for (let i = 0; i < Math.min(ba.length, bb.length); i++) {
    const [ma, mb] = await Promise.all([sharp(ba[i]).metadata(), sharp(bb[i]).metadata()]);
    const w = Math.min(ma.width, mb.width);
    const h = Math.min(ma.height, mb.height);
    const [pa, pb] = await Promise.all([ba[i], bb[i]].map((p) => sharp(p).extract({ left: 0, top: 0, width: w, height: h }).greyscale().raw().toBuffer()));
    const diff = Buffer.alloc(w * h, 255);
    let antalOlika = 0;
    for (let j = 0; j < pa.length; j++) if (Math.abs(pa[j] - pb[j]) > 48) { antalOlika++; diff[j] = 0; }
    res.push({ sida: i + 1, andel: antalOlika / (w * h), w, h, diff, a: ba[i], b: bb[i] });
  }
  mkdirSync(bildMapp, { recursive: true });
  for (const f of readdirSync(bildMapp)) if (f.startsWith(`${namn}-s`)) rmSync(join(bildMapp, f));
  for (const r of [...res].sort((x, y) => y.andel - x.andel).slice(0, 3).filter((r) => r.andel > GRANS)) {
    await sharp({ create: { width: r.w * 3 + 8, height: r.h, channels: 3, background: '#999999' } }).composite([
      { input: await sharp(r.a).extract({ left: 0, top: 0, width: r.w, height: r.h }).png().toBuffer(), left: 0, top: 0 },
      { input: await sharp(r.b).extract({ left: 0, top: 0, width: r.w, height: r.h }).png().toBuffer(), left: r.w + 4, top: 0 },
      { input: await sharp(r.diff, { raw: { width: r.w, height: r.h, channels: 1 } }).png().toBuffer(), left: 2 * r.w + 8, top: 0 },
    ]).png().toFile(join(bildMapp, `${namn}-s${r.sida}.png`));
  }
  rmSync(join(nuMapp, '.a'), { recursive: true, force: true });
  rmSync(join(nuMapp, '.b'), { recursive: true, force: true });
  const ibild = res.filter((r) => r.andel > GRANS);
  const delar = [];
  if (a.length !== b.length) delar.push(`${a.length} → ${b.length} sidor`);
  if (olika.length) delar.push(`annan text på sidan ${olika.slice(0, 12).join(', ')}${olika.length > 12 ? ' …' : ''}`);
  if (ibild.length) delar.push(`ändrad bild på ${ibild.length} sidor (${ibild.slice(0, 8).map((r) => `s${r.sida} ${(r.andel * 100).toFixed(2)} %`).join(', ')}${ibild.length > 8 ? ' …' : ''})`);
  return { lage: delar.length ? 'NEJ' : 'ok', delar, sidor: b.length };
}
