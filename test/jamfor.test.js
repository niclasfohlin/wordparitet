// Jämförelsen av Words och Googles pdf utan Word och Google: ordbrytningen, texten som saknas och ankarna, med texter
// som pdf:erna ger dem. Körs med npm test.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as docx from 'docx';
import { ankarsidor, brutnaOrd, NARA, saknadeTecken, wordText } from '../src/jamfor.js';

const { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } = docx;

test('ett ord som bryts mellan två rader är brutet', () => {
  const iFilen = new Set(['personbeskrivning', 'och', 'sedan']);
  assert.deepEqual([...brutnaOrd(['Personbeskrivni\nng och sedan\n'], iFilen)], ['personbeskrivning']);
});

test('ett ord som pdf:en läser i två delar på samma rad är inte brutet (vikkortens glesa ord)', () => {
  const iFilen = new Set(['paraply', 'brandbil', 'och']);
  assert.deepEqual([...brutnaOrd(['parapl y och brandbi l\nparapl y\n'], iFilen)], []);
});

const iFilen = (o = {}) => ({ rubrikrader: [], kortaRubriker: [], ekvationer: new Map(), ...o });

test('en rubrikrad som Word upprepar en gång mer saknas inte i Google, också med en kort cell', () => {
  const helW = 'Vecka Bas Medel Avancerad\nEn rad text\n\fVecka Bas Medel Avancerad\nMer text\n';
  const helG = 'Vecka Bas Medel Avancerad\nEn rad text\nMer text\n';
  const rubrikrader = [{ bokst: 'vecka', siffror: '' }, { bokst: 'medel', siffror: '' }, { bokst: 'avancerad', siffror: '' }];
  assert.deepEqual(saknadeTecken(helW, helG, iFilen({ rubrikrader, kortaRubriker: [{ rad: 'veckabasmedelavancerad', bokst: 'bas', siffror: '' }] })), []);
  // Utan raden för de korta cellerna stod b kvar som saknat, som i De fyra räknesätten i grupp 2026-10-18.
  assert.ok(saknadeTecken(helW, helG, iFilen({ rubrikrader })).some(([c]) => c === 'b'));
});

test('en rubrikrad som pdftotext flätar ihop uppifrån och ned räknas i läsordning (Kortskrivning 2026-10-27)', () => {
  // Uppifrån och ned står cellerna sida vid sida och den första cellens sista ord på nästa rad, i Word två gånger.
  const helW = 'Det eleverna behöver   Så kan läraren säga\nförstå\nText\n\fDet eleverna behöver   Så kan läraren säga\nförstå\nMer\n';
  const helG = 'Det eleverna behöver förstå   Så kan läraren säga\nText\nMer\n';
  const lasordning = { w: 'Det eleverna behöver förstå\nSå kan läraren säga\nText\n\fDet eleverna behöver förstå\nSå kan läraren säga\nMer\n', g: 'Det eleverna behöver förstå\nSå kan läraren säga\nText\nMer\n' };
  const rubrikrader = [{ bokst: 'detelevernabehöverförstå', siffror: '' }, { bokst: 'såkanlärarensäga', siffror: '' }];
  assert.deepEqual(saknadeTecken(helW, helG, iFilen({ rubrikrader }), undefined, lasordning), []);
  // Utan läsordningen hittades den första cellen inte, och bokstäver stod kvar som saknade.
  assert.ok(saknadeTecken(helW, helG, iFilen({ rubrikrader })).length > 0);
});

test('text som verkligen saknas i Google syns fortfarande', () => {
  assert.deepEqual(saknadeTecken('Hej på dig\n', 'på dig\n', iFilen()).map(([c]) => c).sort(), ['e', 'h', 'j']);
});

test('wordText hittar de korta cellerna i en rubrikrad', async () => {
  const cell = (t) => new TableCell({ width: { size: 2000, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun(t)] })] });
  const tabell = new Table({ rows: [new TableRow({ tableHeader: true, children: ['Vecka', 'Bas', 'Medel'].map(cell) }), new TableRow({ children: ['1', 'a', 'b'].map(cell) })] });
  const buf = await Packer.toBuffer(new Document({ sections: [{ children: [tabell, new Paragraph('')] }] }));
  const t = await wordText(buf);
  assert.deepEqual(t.kortaRubriker, [{ rad: 'veckabasmedel', bokst: 'bas', siffror: '' }]);
  assert.deepEqual(t.rubrikrader.map((r) => r.bokst).sort(), ['medel', 'vecka']);
});

// Ankarna: sidornas tecken som ankartexter ger dem, bokstäver och siffror utan mellanrum. Marginalens Ord i texten i
// faktatexterna står uppifrån och ned (-layout) på samma rad som rubriken, och i läsordning efter stycket.
const fyll = (n) => 'fyllnadstext'.repeat(Math.ceil(n / 12)).slice(0, n);
const rubrik = 'fördjupningmaskinensomgjordeluftdensomskabopåmar';
// Uppifrån och ned i Googles pdf: marginalens rubrik och ord först, sedan rubriken och stycket med marginalens
// förklaringar emellan.
const iMarginalen = 'orditextenroverettfordonsomkörutanförare' + 'fördjupning' + 'instrumentverktyg' + 'maskinensomgjordeluft' + 'densomskabopåmars';

test('ett ankare som marginalen delar uppifrån och ned hittas i läsordning', () => {
  const layout = [`förstasidan${fyll(200)}`, `${iMarginalen}${fyll(200)}`];
  const lasordning = [`förstasidan${fyll(200)}`, `${rubrik}orditextenrover${fyll(200)}`];
  const ankare = ['förstasidan', rubrik];
  assert.deepEqual(ankarsidor(layout, ankare, undefined, lasordning), [0, 1]);
  // Utan läsordningen hittas det inte, och det ger inget nej: ett ankare som Word inte hittar prövas inte i Google.
  assert.deepEqual(ankarsidor(layout, ankare), [0, -1]);
});

test('nästan alla tecken långt fram tar inte sökningen förbi de andra ankarna', () => {
  // Ankare 2 står överst på sidan 2 men går inte att läsa där. Långt fram står nästan alla dess tecken överst på en
  // annan sida, i en annan ordning. Förut togs den sidan, och ankare 3 på sidan 3 hittades aldrig.
  const a1 = 'förstasidanbörjarhärmedenlångrubrikochnågotmerab';
  const a2 = 'modellpassetproblemochlösningförskräpetruntjorde';
  const a3 = 'lärarenssidamedeltext6underdinafötterjordenärbyg';
  const sidor = Array.from({ length: 40 }, () => fyll(200));
  sidor[0] = a1 + fyll(150);
  sidor[2] = [...a2].map((c) => `${c}xyz`).join('') + fyll(50);
  sidor[3] = a3 + fyll(150);
  sidor[2 + NARA + 20] = [...a2].reverse().join('') + fyll(150);
  assert.deepEqual(ankarsidor(sidor, [a1, a2, a3]), [0, -1, 3]);
  // Nära den väntade sidan räcker nästan alla tecken, som förut.
  const nara = [...sidor];
  nara[2] = [...a2].reverse().join('') + fyll(150);
  assert.deepEqual(ankarsidor(nara, [a1, a2, a3]), [0, 2, 3]);
});

test('i Google väljs bland flera träffar den som ligger närmast Words sida', () => {
  const a1 = 'förstasidanbörjarhärmedenlångrubrikochnågotmerab';
  const a2 = 'samtalaomtextenfrågaochsvarsomståråsidanavtexten';
  const sidor = Array.from({ length: 12 }, () => fyll(200));
  sidor[0] = a1 + fyll(150);
  sidor[3] = a2 + fyll(150);
  sidor[7] = a2 + fyll(150);
  assert.deepEqual(ankarsidor(sidor, [a1, a2], [0, 7]), [0, 7]);
  assert.deepEqual(ankarsidor(sidor, [a1, a2], [0, 3]), [0, 3]);
});
