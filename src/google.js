// Google Dokuments återgivning av en Word-fil: Word-filen laddas upp till Google Drive, Google Dokument gör om den till
// ett eget dokument, och dokumentet hämtas som pdf. Så ser filen ut för läraren som sparar den i Drive och öppnar den.
//
// Inloggningen anges av den som anropar: { inloggning } är antingen ett objekt med token() (som ger en åtkomstnyckel med
// behörigheten drive.file) eller sökvägen till en CommonJS- eller ESM-modul som har token(). Utan den läses sökvägen ur
// miljövariabeln WORDPARITET_GOOGLE. Nycklar skrivs aldrig ut. Utan testmapp är dokumentet tillfälligt och går till
// papperskorgen efter hämtningen. I en testmapp ligger det kvar under filens namn, och en ny version skriver över samma
// dokument med samma länk; ett dokument i testmappen läggs aldrig i papperskorgen, eftersom någon kan ha det öppet
// (2026-09-30 blev bilderna varningstrianglar för den som läste).
import { createRequire } from 'node:module';
import { existsSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

const API = 'https://www.googleapis.com/drive/v3/files';
const UPP = 'https://www.googleapis.com/upload/drive/v3/files';
const DOKUMENT = 'application/vnd.google-apps.document';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const nycklar = new Map();
async function hamtaToken(inloggning) {
  const kalla = inloggning ?? process.env.WORDPARITET_GOOGLE;
  if (!kalla) throw new Error('Google-inloggningen saknas: ange { inloggning } eller miljövariabeln WORDPARITET_GOOGLE (sökvägen till en modul med token()).');
  if (typeof kalla === 'object' && typeof kalla.token === 'function') return kalla.token();
  if (!existsSync(kalla)) throw new Error(`Google-inloggningen finns inte: ${kalla}`);
  let modul;
  try { modul = createRequire(import.meta.url)(kalla); } catch { modul = await import(pathToFileURL(kalla).href); }
  return (modul.token ?? modul.default?.token)();
}
async function huvud(inloggning) {
  const nyckel = typeof inloggning === 'string' || !inloggning ? inloggning ?? process.env.WORDPARITET_GOOGLE ?? '' : inloggning;
  if (!nycklar.has(nyckel)) nycklar.set(nyckel, { Authorization: `Bearer ${await hamtaToken(inloggning)}` });
  return { nyckel, auth: nycklar.get(nyckel) };
}
// Nyckeln gäller omkring en timme; svarar Google 401 hämtas en ny en gång (en lång körning med Word kan ta längre tid).
async function anrop(url, init = {}, inloggning, igen = true) {
  const { nyckel, auth } = await huvud(inloggning);
  const svar = await fetch(url, { ...init, headers: { ...auth, ...(init.headers ?? {}) } });
  if (svar.status === 401 && igen) { nycklar.delete(nyckel); return anrop(url, init, inloggning, false); }
  if (!svar.ok) throw new Error(`${init.method ?? 'GET'} ${url.split('?')[0]} svarade ${svar.status}: ${(await svar.text()).replace(/\s+/g, ' ').slice(0, 300)}`);
  return svar;
}
const sok = async (q, inloggning) => (await (await anrop(`${API}?q=${encodeURIComponent(`${q} and trashed=false`)}&fields=files(id,name)`, {}, inloggning)).json()).files;
const papperskorgen = (id, inloggning) => anrop(`${API}/${id}`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ trashed: true }) }, inloggning);

/** Testmappens id; mappen skapas om den inte finns. */
export async function testmappen(namn, { inloggning } = {}) {
  const finns = (await sok(`name='${namn.replace(/'/g, "\\'")}' and mimeType='application/vnd.google-apps.folder'`, inloggning))[0]?.id;
  if (finns) return finns;
  return (await (await anrop(API, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: namn, mimeType: 'application/vnd.google-apps.folder' }) }, inloggning)).json()).id;
}

/**
 * Gör om Word-filen (en Buffer) till ett Google-dokument och hämtar det som pdf. Med mappId ligger dokumentet kvar i
 * mappen under namnet, och en ny version skriver över samma dokument; en uppdatering har ingen metadata (föräldern får
 * inte anges där), bara den nya Word-filen. Utan mappId går dokumentet till papperskorgen, där det går att hämta
 * tillbaka i 30 dagar. Svarar { pdf, lank }.
 */
export async function googlePdf(docx, { namn = 'fil', mappId, inloggning } = {}) {
  const forra = mappId ? (await sok(`name='${namn.replace(/'/g, "\\'")}' and '${mappId}' in parents`, inloggning))[0] : undefined;
  const meta = forra ? {} : { name: mappId ? namn : `wordparitet ${namn} (tas bort)`, mimeType: DOKUMENT, ...(mappId ? { parents: [mappId] } : {}) };
  const grans = `wordparitet-${Date.now().toString(36)}`;
  const kropp = Buffer.concat([
    Buffer.from(`--${grans}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n--${grans}\r\nContent-Type: ${DOCX}\r\n\r\n`),
    Buffer.from(docx),
    Buffer.from(`\r\n--${grans}--`),
  ]);
  const { id, webViewLink } = await (await anrop(
    `${UPP}${forra ? `/${forra.id}` : ''}?uploadType=multipart&fields=id,webViewLink`,
    { method: forra ? 'PATCH' : 'POST', headers: { 'Content-Type': `multipart/related; boundary=${grans}` }, body: kropp },
    inloggning,
  )).json();
  try {
    return { pdf: Buffer.from(await (await anrop(`${API}/${id}/export?mimeType=application/pdf`, {}, inloggning)).arrayBuffer()), lank: webViewLink };
  } finally {
    if (!mappId) await papperskorgen(id, inloggning);
  }
}
