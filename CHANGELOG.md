# Ändringar

## 0.2.0 (2026-10-18)

- **Ordbrytningen** (`brutnaOrd` i `src/jamfor.js`): ett ord räknas som brutet bara när första delen står sist på en rad och andra delen först på nästa. Två delar på samma rad är ett glest ritat ord: vikkorten i Ljudlek i grupp ritar ordet över ljudens prickar, och pdf:en läser "parapl y", som stoppade sajtens Första ljudet två gånger (niclasfohlin.se 2026-10-17, sajtens K-394). Prövat i `test/jamfor.test.js` och med sajtens googleprov på Första ljudet, som nu går igenom.
- **Texten som saknas** (`saknadeTecken`, utbruten ur `jamforWordGoogle`): en cell i en rubrikrad med färre än fyra bokstäver, som Bas i raden Vecka, Bas, Medel och Avancerad, räknas med hela raden. Står raden fler gånger i Word än i Google dras de korta cellernas tecken av, som de långa cellernas redan drogs av. Förut sa provet att ett b saknades i Google när Word upprepade raden på en sida mer (niclasfohlin.se 2026-10-18, De fyra räknesätten i grupp). Text som verkligen saknas syns som förut; proven i `test/jamfor.test.js` visar båda.
- `wordText`, `brutnaOrd` och `saknadeTecken` exporteras ur `src/jamfor.js`, så att proven når dem utan Word och Google.

Och det som var ej utgivet sedan 0.1.0:

- **Word i telefonen, tabellcellerna** (`REGLER.md`): ett rutnät och andra mönster av många rutor ritas som en bild och inte som en tabell med en cell per ruta. Metodriggens fil med De fyra räknesätten i grupp gick från 7 331 tabellceller till 752 och blev i telefonen "snabbare och funkar direkt" (Niclas 2026-10-05). Från metodriggen 2026-10-05 (metodriggens 7d80262). Sajtens `rutnatBild` gör samma sak (7 454 till 890 celler) och ritar linjerna med crispEdges, så att en förstorad reservbild får jämna linjer. Regelprovet räknar inte cellerna ännu.
- **Word i telefonen** (`REGLER.md`): Word för Android ritar ett inbäddat typsnitt först när det är inläst, så texten i elevens typsnitt syns efter en stund, längre i en stor fil. I typsnittet finns inget att rätta. Från niclasfohlin.se och metodriggen 2026-10-05, prövat av Niclas i telefonen med riggens `prov-typsnitt-mobil.docx` och sajtens fil.
- **Regeln om keepNext före en sidbrytning** (`REGLER.md`): ett stycke som hänger ihop med nästa står inte direkt före ett stycke med sidbrytning före, eftersom Google flyttar stycket till en egen sida. Från niclasfohlin.se 2026-10-04 (De fyra räknesätten i grupp, ramens text före veckorna), mätt med sajtens googleprov före och efter. Regelprovet prövar det inte ännu.

## 0.1.0 (2026-10-01)

Första versionen, steg 1: kunskapen och proven, flyttade från niclasfohlin.se (Niclas 2026-10-01: "dela hela din kunskap och hur man jämför docx och drive ... som en modul eller liknande ni båda kan bygga ut och göra bättre").

- **Reglerna** (`src/regler.js`, `REGLER.md`): fjorton regler med skäl, ur sajtens `scripts/wordregler.mjs` och METODER.md. Den senaste är kantstilen: Google ritar en dubbel kant som en enkel, mätt i `bank/dubbellinje.mjs` (K-157). Undantagen anges nu av den som provar, med skäl, i stället för att stå i regeln.
- **Regelprovet** som en funktion, `provaWordfil`, och kommandot `wordparitet regler`. Det ger samma svar som sajtens skript: sajtens 53 Word-filer går igenom med sajtens två undantag för bocken (✓) i Problemlösning i grupp, och riggens 20 Word-filer har de fyra slags brott riggen själv listade (radavstånd, rader med satt höjd och cellmarginal, tabeller i celler utan stycken runt och tabeller i celler i rader med satt höjd).
- **Mätbänken** (`src/bank.js`, `wordparitet bank`) med sajtens fem prov och deras facit: `radavstand`, `styckeavstand`, `radhojd`, `tabellrader` och `dubbellinje`. Typsnitten som proven bäddar in ligger i `bank/typsnitt/`, med licenserna.
- **Jämförelserna** (`src/jamfor.js`): Word mot Google (`jamforWordGoogle`, ur sajtens `googleprov.mjs`) och Word före mot efter (`wordFacit`, `wordJamfor`, ur sajtens `wordjmf.mjs`). Det som bara gällde sajten är inställningar: sidhuvudets och sidfotens rader (`ram`), sidnumrens form (`sidnummer`), mapparna och testmappen.
- **Google och Word** (`src/google.js`, `src/word.js`, `ps/word-pdf.ps1`): inloggningen anges av den som anropar.
- **Proven** (`test/regler.test.js`): tolv prov utan nätverk.
