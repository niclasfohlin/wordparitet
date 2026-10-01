#!/usr/bin/env node
// Kommandona för wordparitet.
//
//   wordparitet regler <fil.docx> …                  regelprovet; stannar med kod 1 vid brott
//   wordparitet bank <prov | fil.mjs> [--facit]      mätbänken i Word och Google (bank/<prov>.mjs)
//   wordparitet google <fil.docx> … [--mapp <namn>] [--igen] [--sidor 1,21] [--ut <mapp>]
//                                                    Word mot Google, med översiktsark
//
// Google-inloggningen anges med miljövariabeln WORDPARITET_GOOGLE (sökvägen till en modul med token()).
import { existsSync, readFileSync } from 'node:fs';
import { basename, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const [kommando, ...args] = process.argv.slice(2);
const flagga = (n) => (args.includes(n) ? args[args.indexOf(n) + 1] : undefined);
const filer = args.filter((a, i) => !a.startsWith('--') && !(i > 0 && ['--mapp', '--sidor', '--ut'].includes(args[i - 1])));

if (kommando === 'regler') {
  const { provaWordfil, brottSomText } = await import('../src/regler.js');
  let fel = 0;
  for (const fil of filer) {
    const brott = await provaWordfil(readFileSync(fil), { filnamn: basename(fil) });
    if (brott.length) { fel++; console.error(`${basename(fil)}:\n  ${brottSomText(brott).join('\n  ')}`); }
  }
  if (fel) { console.error(`\nWord-reglerna: ${fel} av ${filer.length} Word-filer använder något som Google Dokument ritar annorlunda än Word (REGLER.md).`); process.exit(1); }
  console.log(`Word-reglerna: alla ${filer.length} Word-filer klarar reglerna för Word och Google Dokument (REGLER.md).`);
} else if (kommando === 'bank') {
  const { korBank } = await import('../src/bank.js');
  const prov = filer[0];
  const fil = prov && (existsSync(prov) ? resolve(prov) : join(fileURLToPath(new URL('../bank/', import.meta.url)), `${prov.replace(/\.mjs$/, '')}.mjs`));
  if (!fil || !existsSync(fil)) { console.error('Ange ett prov: wordparitet bank <namn i bank/ | fil.mjs> [--facit]'); process.exit(1); }
  const { rader, andrade, bild } = await korBank(fil, { facit: args.includes('--facit'), ut: flagga('--ut') });
  console.log(rader.join('\n'));
  if (bild) console.log(`\nWord till vänster, Google till höger: ${bild}`);
  if (andrade && !args.includes('--facit')) { console.error(`\nmätbänken: ${andrade} varianter mäter annorlunda än provets vantat. Läs om regeln fortfarande gäller; --facit skriver de nya stegen.`); process.exit(1); }
} else if (kommando === 'google') {
  const { jamforWordGoogle } = await import('../src/jamfor.js');
  const { testmappen } = await import('../src/google.js');
  const ut = resolve(flagga('--ut') ?? 'wordparitet-google');
  const mappId = flagga('--mapp') ? await testmappen(flagga('--mapp')) : undefined;
  const sidor = flagga('--sidor') ? flagga('--sidor').split(',').map(Number) : [];
  let fel = 0;
  for (const fil of filer) {
    const r = await jamforWordGoogle(fil, { ut, mappId, igen: args.includes('--igen'), sidor });
    if (r.lage === 'NEJ') fel++;
    console.log(`${r.lage.padEnd(3)}  ${r.namn}: ${r.beskrivning}`);
    console.log(`     översikt (Word över Google): ${r.ark}${r.lank && mappId ? `\n     i testmappen: ${r.lank}` : ''}`);
  }
  if (fel) process.exit(1);
} else {
  console.error('Kommandon: regler <fil.docx> …, bank <prov> [--facit], google <fil.docx> … [--mapp <namn>]. Se README.md.');
  process.exit(1);
}
