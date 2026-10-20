// Luftprovet (src/luft.js) utan Word och Google: bedömningen av blocken, blocken ur ankarna, sektionerna i en Word-fil
// och sidornas mått i en liten pdf som provet skriver själv (kräver Popplers pdftoppm, som jämförelsen redan kräver).
// Körs med npm test.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import * as docx from 'docx';
import { bedom, block, LIKA_LUFT, MARGINAL, sektioner, sidmatt, standarddrift, STANDARDDRIFT } from '../src/luft.js';

const { Document, Packer, Paragraph } = docx;

test('ett block med mer luft än standarddriften och marginalen ryms', () => {
  const r = bedom([{ ankare: 'aaaaaaaa', fran: 0, k: 1, luft: STANDARDDRIFT[1] + MARGINAL }]);
  assert.equal(r.lage, 'ryms');
});

test('en fil utan block har inte mätts och går till Google', () => {
  const r = bedom([]);
  assert.equal(r.lage, 'google');
  assert.equal(r.inga, true);
  assert.equal(r.tranga.length, 0);
});

test('ett block med mindre luft än standarddriften och marginalen går till Google', () => {
  const r = bedom([{ ankare: 'aaaaaaaa', fran: 0, k: 1, luft: STANDARDDRIFT[1] + MARGINAL - 1 }]);
  assert.equal(r.lage, 'google');
  assert.equal(r.tranga[0].kalla, 'standard');
});

test('uppmätt drift för samma ankare och längd går före standarddriften, också när den är negativ', () => {
  const b = { ankare: 'aaaaaaaa', fran: 0, k: 1, luft: 5 };
  assert.equal(bedom([b], [{ ankare: 'aaaaaaaa', k: 1, drift: -2 }]).lage, 'ryms');
  assert.equal(bedom([b], [{ ankare: 'aaaaaaaa', k: 1, drift: 0 }]).lage, 'google');
  // Ett block som har blivit längre sedan mätningen får standarddriften för sin nya längd.
  assert.equal(bedom([{ ...b, k: 2, luft: 30 }], [{ ankare: 'aaaaaaaa', k: 1, drift: -2 }]).tranga[0].kalla, 'standard');
});

test('uppmätt drift gäller bara när blockets luft är nästan densamma som när den mättes', () => {
  const b = { ankare: 'aaaaaaaa', fran: 0, k: 1, luft: 20 };
  assert.equal(bedom([b], [{ ankare: 'aaaaaaaa', k: 1, luft: 20 + LIKA_LUFT, drift: 4 }]).lage, 'ryms');
  // Innehållet har växt, som en tabell med fler rader: standarddriften gäller igen.
  const r = bedom([b], [{ ankare: 'aaaaaaaa', k: 1, luft: 21 + LIKA_LUFT, drift: 4 }]);
  assert.equal(r.lage, 'google');
  assert.equal(r.tranga[0].kalla, 'standard');
});

test('ett lektionskort med 49 pt luft, som metodriggen såg välta i Google, går till Google utan uppmätt drift', () => {
  assert.equal(bedom([{ ankare: 'lektionskortet', fran: 0, k: 1, luft: 49 }]).lage, 'google');
});

test('standarddriften efter blockets längd', () => {
  assert.equal(standarddrift(1), STANDARDDRIFT[1]);
  assert.equal(standarddrift(2), STANDARDDRIFT[2]);
  assert.equal(standarddrift(7), STANDARDDRIFT.fler);
});

test('blocken går från ankare till ankare, med luften på blockets sista sida', () => {
  const matt = [10, 20, 30].map((luft) => ({ luft }));
  assert.deepEqual(block(['forstaankaret', 'andraankaret', 'saknas i word'], [0, 2, -1], matt), [
    { ankare: 'forstaankaret', fran: 0, k: 2, luft: 20 },
    { ankare: 'andraankaret', fran: 2, k: 1, luft: 30 },
  ]);
});

test('sektionerna läses ur Word-filen i punkter', async () => {
  const buf = await Packer.toBuffer(new Document({ sections: [{ properties: { page: { margin: { bottom: 1134, footer: 567 } } }, children: [new Paragraph('Hej')] }] }));
  const [s] = await sektioner(buf);
  assert.equal(s.botten, 56.7);
  assert.equal(s.fot, 28.35);
  assert.ok(Math.abs(s.h - 841.9) < 1);
});

// En pdf med två A4: en svart rektangel vars nederkant står 442 pt från sidans överkant, och en tom sida.
function pdfMedRektangel() {
  const strom = '0 0 0 rg 72 400 100 50 re f';
  const obj = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R 5 0 R] /Count 2 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Contents 4 0 R >>',
    `<< /Length ${strom.length} >>\nstream\n${strom}\nendstream`,
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] >>',
  ];
  let ut = '%PDF-1.4\n';
  const lagen = [];
  obj.forEach((o, i) => { lagen.push(ut.length); ut += `${i + 1} 0 obj\n${o}\nendobj\n`; });
  const xref = ut.length;
  ut += `xref\n0 ${obj.length + 1}\n0000000000 65535 f \n${lagen.map((l) => `${String(l).padStart(10, '0')} 00000 n \n`).join('')}`;
  ut += `trailer\n<< /Size ${obj.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return ut;
}
const harPdftoppm = (() => { try { execFileSync('pdftoppm', ['-v'], { stdio: 'ignore' }); return true; } catch { return false; } })();

test('sidornas mått: innehållets nederkant, luften och den tomma sidan', { skip: !harPdftoppm && 'pdftoppm saknas' }, () => {
  const tmp = mkdtempSync(join(tmpdir(), 'wordparitet-luftprov-'));
  try {
    const pdf = join(tmp, 'prov.pdf');
    writeFileSync(pdf, pdfMedRektangel(), 'latin1');
    const [forsta, andra] = sidmatt(pdf, [{ b: 595.28, h: 841.89, botten: 56.7, fot: 0 }]);
    assert.ok(Math.abs(forsta.botten - 442) <= 3, `nederkanten ${forsta.botten}`);
    assert.ok(Math.abs(forsta.luft - (841.89 - 56.7 - 442)) <= 3, `luften ${forsta.luft}`);
    assert.equal(forsta.tom, false);
    assert.equal(andra.tom, true);
  } finally {
    rmSync(tmp, { recursive: true, force: true });
  }
});
