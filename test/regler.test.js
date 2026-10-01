// Regelprovet mot små Word-filer: en fil som klarar alla regler, och för varje regel en fil som bryter den. Körs utan
// nätverk, Word och Google (npm test), före varje commit i båda riggarna.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import * as docx from 'docx';
import { provaWordfil, REGLER } from '../src/regler.js';

const {
  AlignmentType, BorderStyle, Document, HeightRule, LevelFormat, LineRuleType, Packer, Paragraph, Table, TableCell,
  TableLayoutType, TableRow, TabStopType, TextRun, WidthType, LeaderType,
} = docx;

const ingen = { style: BorderStyle.NONE, size: 0, color: 'auto' };
const punkt = () => new Paragraph({ spacing: { before: 0, after: 0, line: 240 }, run: { size: 2, font: 'Calibri' }, children: [new TextRun({ text: '', size: 2, font: 'Calibri' })] });
const cell = (barn, o = {}) => new TableCell({ width: { size: 4000, type: WidthType.DXA }, children: barn, ...o });
const tabell = (rader, o = {}) => new Table({ width: { size: 4000, type: WidthType.DXA }, columnWidths: [4000], layout: TableLayoutType.FIXED, rows: rader, ...o });
// Ett dokument som Word-filerna på sajten: Calibri, svenska, och ett stycke på en punkt sist, så att avsnittet slutar rätt.
async function fil(barn, o = {}) {
  const dokument = new Document({
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
    ...o,
    sections: [{ children: [...barn, punkt()] }],
  });
  return Packer.toBuffer(dokument);
}
const brutna = async (data, undantag) => (await provaWordfil(data, { filnamn: 'prov.docx', undantag })).map((b) => b.regel);

test('en fil som följer reglerna har inga brott', async () => {
  const data = await fil([
    new Paragraph({ children: [new TextRun('Läs texten och svara på frågorna – högt, två gånger.')] }),
    new Paragraph({ spacing: { line: 300 }, children: [new TextRun({ text: 'Radavstånd som multipel.', font: 'Andika' })] }),
    tabell([new TableRow({ children: [cell([new Paragraph('Cell')], { borders: { top: { style: BorderStyle.SINGLE, size: 4, color: '999999' }, bottom: ingen, left: ingen, right: ingen } })] })]),
    punkt(),
  ]);
  assert.deepEqual(await brutna(data), []);
});

test('numreringens start räknas inte som kant', async () => {
  const data = await fil([new Paragraph({ numbering: { reference: 'lista', level: 0 }, children: [new TextRun('Ett')] })], {
    numbering: { config: [{ reference: 'lista', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START }] }] },
  });
  assert.ok(!(await brutna(data)).includes('kantstil som Google saknar'));
});

const fall = {
  'exakt eller minsta radavstånd': () => [new Paragraph({ spacing: { line: 300, lineRule: LineRuleType.EXACT }, children: [new TextRun('Exakt')] })],
  'kantstil som Google saknar': () => [tabell([new TableRow({ children: [cell([new Paragraph('Dubbel')], { borders: { top: { style: BorderStyle.DOUBLE, size: 12, color: '000000' }, bottom: ingen, left: ingen, right: ingen } })] })]), punkt()],
  'typsnitt som Google saknar': () => [new Paragraph({ children: [new TextRun({ text: 'Times', font: 'Times New Roman' })] })],
  'tecken som inte är prövade i Word och Google': () => [new Paragraph({ children: [new TextRun('☐ Bockruta')] })],
  'teckenavstånd': () => [new Paragraph({ children: [new TextRun({ text: 'Spärrad', characterSpacing: 20 })] })],
  'tabell i en cell utan ett stycke på en punkt före och efter': () => [tabell([new TableRow({ children: [cell([tabell([new TableRow({ children: [cell([new Paragraph('Inre')])] })]), new Paragraph('')])] })]), punkt()],
  'rad med satt höjd och cellmarginal upptill, eller nedtill mer än kantens bildpunkter': () => [tabell([new TableRow({ height: { value: 600, rule: HeightRule.ATLEAST }, children: [cell([new Paragraph('Rad')], { margins: { top: 80, bottom: 80 } })] })]), punkt()],
  'tabbstopp som Google saknar': () => [new Paragraph({ tabStops: [{ type: TabStopType.RIGHT, position: 4000, leader: LeaderType.DOT }], children: [new TextRun('Prickar')] })],
};
for (const [regel, barn] of Object.entries(fall)) {
  test(`bryter: ${regel}`, async () => {
    assert.ok(REGLER.some((r) => r.namn === regel), `regeln "${regel}" finns inte i REGLER`);
    assert.ok((await brutna(await fil(barn()))).includes(regel));
  });
}

test('dokumentet slutar med en tabell', async () => {
  const data = await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Före'), tabell([new TableRow({ children: [cell([new Paragraph('Sist')])] })])] }] }));
  assert.ok((await brutna(data)).includes('dokumentet slutar med en tabell'));
});

test('ett undantag med skäl släpper igenom ett tecken i en fil', async () => {
  const data = await fil([new Paragraph({ children: [new TextRun('187 ✓')] })]);
  assert.ok((await brutna(data)).includes('tecken som inte är prövade i Word och Google'));
  const undantag = [{ regel: 'tecken som inte är prövade i Word och Google', fil: 'prov', tecken: '✓', skal: 'ett tecken i en tabellcell' }];
  assert.ok(!(await brutna(data, undantag)).includes('tecken som inte är prövade i Word och Google'));
});
