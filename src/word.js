// Words egen återgivning av en Word-fil: Word (COM, ps/word-pdf.ps1) exporterar filen som pdf, som när läraren skriver ut
// den. Kräver Windows med Word. Bara en Word-körning åt gången: skriptet startar Word och stänger det efteråt, så två
// samtidiga körningar (ett prov i bänken, en jämförelse, en mätning i ett bygge) stänger varandras Word.
import { execFile, execFileSync } from 'node:child_process';
import { promisify } from 'node:util';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const SKRIPT = join(dirname(fileURLToPath(import.meta.url)), '..', 'ps', 'word-pdf.ps1');
const kor = promisify(execFile);
const argument = (fil, pdf) => ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', SKRIPT, fil, pdf];

/** Gör pdf:en ut av Word-filen fil med Word. Väntar högst timeout millisekunder (fem minuter). */
export async function wordPdf(fil, ut, { timeout = 300000 } = {}) {
  await kor('powershell', argument(fil, ut), { timeout });
  return ut;
}
/** Samma som wordPdf, men väntar i samma tråd. */
export function wordPdfSync(fil, ut, { timeout = 300000 } = {}) {
  execFileSync('powershell', argument(fil, ut), { stdio: 'pipe', timeout, encoding: 'utf8' });
  return ut;
}
/** Sökvägen till ps/word-pdf.ps1, för den som vill köra skriptet själv. */
export const WORD_PDF_SKRIPT = SKRIPT;
