// Jämförelsen av Words och Googles pdf utan Word och Google: ordbrytningen och texten som saknas, med texter som pdf:erna
// ger dem. Körs med npm test.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as docx from 'docx';
import { brutnaOrd, saknadeTecken, wordText } from '../src/jamfor.js';

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
