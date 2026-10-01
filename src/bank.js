// Mätbänken: hur ritar Word och Google Dokument en form? En ny form i en Word-fil (en ny sorts kort, tabell, rad eller
// linje) prövas här innan den byggs in. Ett prov (bank/*.mjs) bygger varianter av formen med docx, Word och Google
// Dokument ritar samma fil, och bänken mäter varje variants radsteg (avståndet mellan etiketterna Ar0, Ar1 …) i båda.
// Det som skiljer blir en regel i src/regler.js, en byggsten och en rad i REGLER.md. Proven är mätningarna bakom
// reglerna och förebilder: kopiera det prov som liknar formen och byt varianterna.
//
// Varje prov har sina uppmätta steg som `vantat`, i git. Bänken säger ÄNDRAT när ett steg skiljer mer än 0,1 pt från
// dem, så att en regel som slutat gälla syns när proven körs om efter att docx eller Google bytt version; med facit
// skrivs de nya stegen in i provet.
//
// Ett prov exporterar `namn`, `varianter` ({ A: { text, barn: (e) => Barn[] }, … }, där e(i) ger etiketten för rad i,
// "Ar0", som ska stå först i sin rad eller sitt stycke; en variant heter med versaler), om det behövs `typsnitt` (docx
// fonts) och `vantat` ({ A: [Word, Google] } i punkter). Varje variant följs av ett stycke "mellan". Kräver Word
// (src/word.js), Google-inloggningen (src/google.js) och Popplers pdftotext och pdftoppm. Proven bygger dokumentet med
// samma kopia av docx som bänken, så att filen går att öppna i Word.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as docx from 'docx';
import { googlePdf } from './google.js';
import { wordPdfSync } from './word.js';

// Popplers pdftotext ger ordens lägen (-bbox); den i Git för Windows (xpdf) gör det inte.
function popplerPdftotext() {
  try {
    const pdftoppm = execFileSync(process.platform === 'win32' ? 'where' : 'which', ['pdftoppm'], { encoding: 'utf8' }).split(/\r?\n/)[0].trim();
    const fil = join(dirname(pdftoppm), process.platform === 'win32' ? 'pdftotext.exe' : 'pdftotext');
    if (existsSync(fil)) return fil;
  } catch { /* ingen pdftoppm i PATH */ }
  throw new Error('Popplers pdftotext saknas (den ska ligga bredvid pdftoppm).');
}

/**
 * Kör provet provFil i Word och Google och mäter radstegen. ut är mappen där Word-filen, pdf:erna, resultatet och
 * bilden av första sidan hamnar. Med facit skrivs de uppmätta stegen in i provet som vantat. Svarar { rader, andrade,
 * nyttFacit, bild }; andrade är antalet varianter som skiljer mer än 0,1 pt från vantat.
 */
export async function korBank(provFil, { ut, facit = false, inloggning } = {}) {
  const prov = await import(pathToFileURL(resolve(provFil)).href);
  const { namn, varianter, typsnitt, vantat } = prov;
  if (!namn || !varianter) throw new Error(`${provFil} ska exportera namn och varianter.`);
  const mapp = ut ?? join(process.cwd(), 'matbank', namn);
  mkdirSync(mapp, { recursive: true });
  const PDFTOTEXT = popplerPdftotext();
  // Samma standard som Word-filerna på sajten: Calibri 11 pt, svenska, A4 med 1 cm marginal.
  const barn = [];
  for (const [v, variant] of Object.entries(varianter)) {
    barn.push(...variant.barn((i) => `${v}r${i}`));
    barn.push(new docx.Paragraph({ spacing: { before: 0, after: 0 }, children: [new docx.TextRun('mellan')] }));
  }
  const dokument = new docx.Document({
    ...(typsnitt ? { fonts: typsnitt } : {}),
    styles: { default: { document: { run: { font: 'Calibri', size: 22, language: { value: 'sv-SE' } } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 567, bottom: 567, left: 567, right: 567 } } }, children: barn }],
  });
  const fil = join(mapp, `${namn}.docx`);
  writeFileSync(fil, await docx.Packer.toBuffer(dokument));
  const pdf = { word: join(mapp, `${namn}-word.pdf`), google: join(mapp, `${namn}-google.pdf`) };
  try {
    wordPdfSync(fil, pdf.word);
  } catch (e) {
    throw new Error(`Word kunde inte göra pdf:en av ${fil}. Oftast går en annan Word-körning samtidigt; kör en i taget. Words svar: ${String(e.stderr || e.message).trim().split('\n').slice(0, 6).join(' ')}`);
  }
  writeFileSync(pdf.google, (await googlePdf(readFileSync(fil), { namn: `matbank ${namn}`, inloggning })).pdf);

  // Etiketternas lägen, sida för sida: yMin är överkanten av ordets ruta, så radsteget är skillnaden mellan två
  // etiketter på samma sida. Word och Google ger typsnittet olika övre kant i pdf:en; därför jämförs bara steg.
  const etiketter = (pdfFil) => {
    const xml = execFileSync(PDFTOTEXT, ['-bbox', '-enc', 'UTF-8', pdfFil, '-'], { encoding: 'utf8', maxBuffer: 1 << 26 });
    const sidor = xml.split('<page ').slice(1);
    const lagen = {};
    sidor.forEach((s, nr) => {
      for (const m of s.matchAll(/<word xMin="[\d.]+" yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">([A-Z]+)r(\d+)<\/word>/g)) {
        (lagen[m[2]] ??= []).push({ sida: nr + 1, y: Number(m[1]), i: Number(m[3]) });
      }
    });
    return { lagen, sidor: sidor.length };
  };
  const matt = { word: etiketter(pdf.word), google: etiketter(pdf.google) };
  const steg = (l = []) => {
    const s = l.slice(1).map((b, k) => (b.sida === l[k].sida && b.i === l[k].i + 1 ? b.y - l[k].y : null)).filter((x) => x !== null);
    return s.length ? { medel: s.reduce((a, b) => a + b, 0) / s.length, min: Math.min(...s), max: Math.max(...s), n: s.length } : null;
  };
  const tal = (x) => (x ? `${x.medel.toFixed(2).padStart(6)} (${x.min.toFixed(2)}–${x.max.toFixed(2)})` : '     –           ');
  const rader = [`Mätbänken: ${namn}, radsteg i punkter (medel, minst–störst). Word ${matt.word.sidor} sidor, Google ${matt.google.sidor}.`, ''];
  const nyttFacit = {};
  let andrade = 0;
  for (const [v, variant] of Object.entries(varianter)) {
    const [w, g] = [steg(matt.word.lagen[v]), steg(matt.google.lagen[v])];
    const skillnad = w && g ? `${(g.medel - w.medel >= 0 ? '+' : '')}${(g.medel - w.medel).toFixed(2)}` : '';
    nyttFacit[v] = [w ? Number(w.medel.toFixed(2)) : null, g ? Number(g.medel.toFixed(2)) : null];
    const fore = vantat?.[v];
    const avvik = fore && [0, 1].some((k) => (fore[k] === null) !== (nyttFacit[v][k] === null) || Math.abs((fore[k] ?? 0) - (nyttFacit[v][k] ?? 0)) > 0.1);
    if (avvik) andrade++;
    const mot = !vantat || facit ? '' : !fore ? '  (nytt)' : avvik ? `  ÄNDRAT, förut Word ${fore[0]} och Google ${fore[1]}` : '';
    rader.push(`${v.padEnd(2)}  Word ${tal(w)}  Google ${tal(g)}  ${skillnad.padStart(6)}  ${variant.text ?? ''}${mot}`);
  }
  writeFileSync(join(mapp, 'resultat.txt'), rader.join('\n') + '\n');

  // Första sidan i Word och Google bredvid varandra, för att se formen och inte bara talen.
  let bild;
  try {
    const sharp = (await import('sharp')).default;
    const bilder = {};
    for (const vem of ['word', 'google']) {
      const bas = join(mapp, `${namn}-${vem}-s1`);
      execFileSync('pdftoppm', ['-r', '60', '-f', '1', '-l', '1', '-png', '-singlefile', pdf[vem], bas]);
      bilder[vem] = `${bas}.png`;
    }
    const [mw, mg] = await Promise.all([sharp(bilder.word).metadata(), sharp(bilder.google).metadata()]);
    bild = join(mapp, `${namn}-word-google.png`);
    await sharp({ create: { width: mw.width + mg.width + 12, height: Math.max(mw.height, mg.height), channels: 3, background: '#999999' } })
      .composite([{ input: bilder.word, left: 0, top: 0 }, { input: bilder.google, left: mw.width + 12, top: 0 }])
      .png().toFile(bild);
  } catch { /* utan sharp blir det ingen bild */ }

  // Med facit skrivs de uppmätta stegen in i provet, i stället för en tidigare rad med vantat.
  if (facit) {
    const kalla = readFileSync(resolve(provFil), 'utf8');
    const rad = `export const vantat = { ${Object.entries(nyttFacit).map(([v, [w, g]]) => `${v}: [${w}, ${g}]`).join(', ')} };`;
    const ny = /^export const vantat = .*;\r?$/m.test(kalla) ? kalla.replace(/^export const vantat = .*;(\r?)$/m, `${rad}$1`) : `${kalla.replace(/\s*$/, '')}\n\n// Uppmätta radsteg i punkter, [Word, Google] (wordparitet bank ${namn} --facit).\n${rad}\n`;
    writeFileSync(resolve(provFil), ny);
  }
  return { rader, andrade, nyttFacit, bild, mapp };
}
