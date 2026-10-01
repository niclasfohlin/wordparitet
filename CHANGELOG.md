# Ändringar

## 0.1.0 (2026-10-01)

Första versionen, steg 1: kunskapen och proven, flyttade från niclasfohlin.se (Niclas 2026-10-01: "dela hela din kunskap och hur man jämför docx och drive ... som en modul eller liknande ni båda kan bygga ut och göra bättre").

- **Reglerna** (`src/regler.js`, `REGLER.md`): fjorton regler med skäl, ur sajtens `scripts/wordregler.mjs` och METODER.md. Den senaste är kantstilen: Google ritar en dubbel kant som en enkel, mätt i `bank/dubbellinje.mjs` (K-157). Undantagen anges nu av den som provar, med skäl, i stället för att stå i regeln.
- **Regelprovet** som en funktion, `provaWordfil`, och kommandot `wordparitet regler`. Det ger samma svar som sajtens skript: sajtens 53 Word-filer går igenom med sajtens två undantag för bocken (✓) i Problemlösning i grupp, och riggens 20 Word-filer har de fyra slags brott riggen själv listade (radavstånd, rader med satt höjd och cellmarginal, tabeller i celler utan stycken runt och tabeller i celler i rader med satt höjd).
- **Mätbänken** (`src/bank.js`, `wordparitet bank`) med sajtens fem prov och deras facit: `radavstand`, `styckeavstand`, `radhojd`, `tabellrader` och `dubbellinje`. Typsnitten som proven bäddar in ligger i `bank/typsnitt/`, med licenserna.
- **Jämförelserna** (`src/jamfor.js`): Word mot Google (`jamforWordGoogle`, ur sajtens `googleprov.mjs`) och Word före mot efter (`wordFacit`, `wordJamfor`, ur sajtens `wordjmf.mjs`). Det som bara gällde sajten är inställningar: sidhuvudets och sidfotens rader (`ram`), sidnumrens form (`sidnummer`), mapparna och testmappen.
- **Google och Word** (`src/google.js`, `src/word.js`, `ps/word-pdf.ps1`): inloggningen anges av den som anropar.
- **Proven** (`test/regler.test.js`): tolv prov utan nätverk.
