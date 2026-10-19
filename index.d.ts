// Typerna för wordparitet. Modulen är JavaScript (ESM) utan byggsteg; typerna står här.

export interface Regel {
  namn: string;
  skal: string;
  prova(xml: string, o?: { tillatna?: Set<string> }): number[];
}
export interface Undantag {
  /** Regelns namn, som i REGLER. */
  regel: string;
  /** Början av filnamnet som undantaget gäller; utan fil gäller det alla filer. */
  fil?: string;
  /** Ett tecken som får stå i filen trots regeln om tecken. */
  tecken?: string;
  /** Varför undantaget finns. */
  skal: string;
}
export interface Brott {
  regel: string;
  skal: string;
  /** Delen i Word-filen, till exempel word/document.xml. */
  del: string;
  antal: number;
  /** Texten närmast efter det första stället. */
  utdrag: string;
}
export const REGLER: Regel[];
export const TYPSNITT: Set<string>;
export const VANLIGA: Set<string>;
export const SARSKILDA: Record<string, string>;
export const FORBJUDNA_TECKEN: Record<string, string>;
export function kantensBildpunkter(pt: number): number;
export function utdrag(xml: string, index: number): string;
export function provaWordfil(data: Uint8Array | ArrayBuffer, o?: { filnamn?: string; undantag?: Undantag[] }): Promise<Brott[]>;
export function brottSomText(brott: Brott[]): string[];

/** En inloggning med åtkomstnyckel för Google Drive (behörigheten drive.file), eller sökvägen till en modul som har token(). */
export type Inloggning = { token(): Promise<string> } | string;
export function testmappen(namn: string, o?: { inloggning?: Inloggning }): Promise<string>;
export function googlePdf(docx: Uint8Array, o?: { namn?: string; mappId?: string; inloggning?: Inloggning }): Promise<{ pdf: Buffer; lank?: string }>;

export function wordPdf(fil: string, ut: string, o?: { timeout?: number }): Promise<string>;
export function wordPdfSync(fil: string, ut: string, o?: { timeout?: number }): string;
export const WORD_PDF_SKRIPT: string;

export const SIDNUMMER: RegExp;
export function innehall(buf: Uint8Array): Promise<string>;
export interface Jamforelse {
  namn: string;
  lage: 'ok' | 'obs' | 'NEJ';
  problem: string[];
  obs: string[];
  beskrivning: string;
  /** Översiktsarket med Words sidor över Googles. */
  ark: string;
  lank?: string;
  provade: number;
  ankare: number;
  ankarlista: string[];
  sidorWord: number;
  sidorGoogle: number;
  bilder: { sida: number; vem: 'word' | 'google'; fil: string }[];
  wordPdf: string;
  googlePdf: string;
  /** Blockens drift: hur mycket längre ned Googles sista sida i blocket slutar än Words, i punkter. Sparas också i <ut>/<namn>-drift.json. */
  drift: Drift[];
}
export interface JamforVal {
  ut: string;
  mappId?: string;
  inloggning?: Inloggning;
  igen?: boolean;
  /** Det som står på varje sida och inte räknas som sidans text, med flaggorna gm. */
  ram?: RegExp;
  sidnummer?: RegExp;
  sidor?: number[];
  visaAnkare?: boolean;
  /** Words pdf av just den här Word-filen, när den som anropar redan har den. */
  wordPdf?: string;
}
export function jamforWordGoogle(fil: string, o: JamforVal): Promise<Jamforelse>;
/** Flera Word-filer mot Google, högst samtidigt åt gången (förval 4); wordPdf ger Words pdf per fil. */
export function jamforFlera(filer: string[], o: Omit<JamforVal, 'wordPdf'> & { samtidigt?: number; wordPdf?: (fil: string) => string | undefined; vid?: (svar: Jamforelse) => void }): Promise<Jamforelse[]>;
export function sidankare(buf: Uint8Array): Promise<string[]>;

/** Luftprovet (src/luft.js). */
export interface Drift { ankare: string; k: number; luft: number; drift: number }
export interface Block { ankare: string; fran: number; k: number; luft: number }
export interface Sidmatt { hojd: number; botten: number; textBotten: number; luft: number; tom: boolean }
export interface Sektion { b: number; h: number; botten: number; fot: number }
export const DPI: number;
export const STANDARDDRIFT: { 1: number; 2: number; fler: number };
export const MARGINAL: number;
export function standarddrift(k: number): number;
export function sektioner(docx: Uint8Array): Promise<Sektion[]>;
export function sidmatt(pdf: string, sekt: Sektion[]): Sidmatt[];
export function block(ankare: string[], iWord: number[], matt: Sidmatt[]): Block[];
export function bedom(blocken: Block[], drift?: Drift[], o?: { marginal?: number }): { lage: 'ryms' | 'google'; tranga: (Block & { drift: number; kalla: 'standard' | 'uppmätt' })[] };
export function luftprov(fil: string, o: { ut: string; wordPdf?: string; sidnummer?: RegExp; marginal?: number }): Promise<{ namn: string; lage: 'ryms' | 'google'; tranga: (Block & { drift: number; kalla: 'standard' | 'uppmätt' })[]; block: number; sidor: number; uppmatt: boolean; wordPdf: string }>;
export function wordFacit(fil: string, facitMapp: string): number;
export function wordJamfor(fil: string, o: { facitMapp: string; nuMapp: string; bildMapp: string; byten?: [string, string][] }): Promise<{ lage: 'ok' | 'NEJ' | '?'; delar: string[]; sidor: number }>;

/** Mätbänken, i 'wordparitet/bank'. */
export function korBank(provFil: string, o?: { ut?: string; facit?: boolean; inloggning?: Inloggning }): Promise<{ rader: string[]; andrade: number; nyttFacit: Record<string, [number | null, number | null]>; bild?: string; mapp: string }>;
