// wordparitet: Word och Google Dokument, delat mellan niclasfohlin.se och metodriggen. Mätbänken hämtas för sig, med
// import { korBank } from 'wordparitet/bank', eftersom den kräver docx.
export { REGLER, TYPSNITT, VANLIGA, SARSKILDA, FORBJUDNA_TECKEN, kantensBildpunkter, provaWordfil, brottSomText, utdrag } from './regler.js';
export { googlePdf, testmappen } from './google.js';
export { wordPdf, wordPdfSync, WORD_PDF_SKRIPT } from './word.js';
export { SIDNUMMER, innehall, jamforWordGoogle, wordFacit, wordJamfor } from './jamfor.js';
