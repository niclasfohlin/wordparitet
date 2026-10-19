# Ändringar

## 0.3.3 (2026-10-19)

- **Tildet `~` är prövat** och står bland de särskilda tecknen (`SARSKILDA` i `src/regler.js`). Det är tecknet för Lyssnarens takt i metodriggens Högläsning i par, i kompendiet, lathunden och lektionskortet. pdfminer på Words och Googles pdf:er av niclasfohlin.se:s tre Word-filer och riggens kompendium: tildet står i samma typsnitt som texten runt det i båda, Calibri, Calibri kursiv och Arial, så raden blir inte högre. Förut stoppade regeln om tecken som inte är prövade filerna.

## 0.3.2 (2026-10-19)

- **Standarddriften höjs** till 50 pt för block på en och två sidor och 24 för längre (`STANDARDDRIFT` i `src/luft.js`): metodriggen har sett lektionskort med 40 och 49 pt luft välta i Google, och sådana fel var lagade innan sajtens mätning gjordes. Ett block utan uppmätt drift behöver nu 56 pt luft.
- **Uppmätt drift gäller bara när blocket har nästan samma luft som när den mättes**, högst 12 pt skillnad (`LIKA_LUFT`). Har innehållet växt, som med fler tabellrader, räknas standarddriften igen.

## 0.3.1 (2026-10-19)

- **Driften för ett block som Google packar tätare**: tar blocket färre sidor i Google än i Word räknas driften över sidgränsen, som ett negativt tal, och sparas. Förut sparades ingen drift för sådana block, och luftprovet skickade samma fem av niclasfohlin.se:s 63 filer till Google varje gång.

## 0.3.0 (2026-10-19)

- **Luftprovet** (`luftprov` i `src/jamfor.js`, `src/luft.js`, kommandot `wordparitet luft`): avgör utan Google om blocken mellan sidankarna ryms i Google Dokument, ur luften på blockets sista sida i Words pdf och driften, det Google lägger till vid ett blocks slut. Standarddriften är mätt i niclasfohlin.se:s 63 Word-filer från 2026-10-19, 1 742 block (REGLER.md under Luftprovet). Niclas samma dag: "Borde inte du kunna bygga nu en enkel testmodul för Google testet som är mycket snabbare? ... En testmodul du och riggen kan dela?"
- **Driften sparas** efter varje jämförelse i `<ut>/<namn>-drift.json`, och `jamforWordGoogle` svarar den som `drift`. Nästa luftprov av samma fil använder den.
- **Words pdf kan ges in** (`wordPdf` i `jamforWordGoogle` och `luftprov`), så att den som redan har den, som niclasfohlin.se:s validering, slipper Word. Word körs annars en åt gången.
- **Flera filer samtidigt** (`jamforFlera`, `wordparitet google --samtidigt 4`): ett fel i en fil blir ett NEJ för den och stoppar inte de andra. Översiktsarkets sidor ritas i en egen mapp per fil.
- **Tillfälliga fel hos Google** (429 och 5xx) försöks igen högst tre gånger (`src/google.js`). 2026-10-19 svarade exporten 500 för en av 63 filer.
- **Tomma sidor** (niclasfohlin.se K-368): en sida i Google räknas som tom bara när den varken har text eller något annat i textytan, så att ett blad med bara en bild inte larmar.
- `sidankare` och luftprovets delar (`sektioner`, `sidmatt`, `block`, `bedom`, `STANDARDDRIFT`, `MARGINAL`) exporteras, och proven i `test/luft.test.js` prövar dem utan Word och Google.

Och det som var ej utgivet sedan 0.2.0:

- **Ankarna i jämförelsen med Google** (`ankarsidor` i `src/jamfor.js`): ett ankare hittas överst på en sida uppifrån och ned (`-layout`) eller i läsordning. Uppifrån och ned skriver pdftotext en spalt i marginalen på samma rad som texten bredvid, så att en rubrik och styckets början kommer isär, och i Googles pdf kan marginalens rubrik stå först. I läsordning står de ihop i båda pdf:erna. En ungefärlig träff, nästan alla tecken i en annan ordning, godtas bara högst `NARA` (10) sidor från den sida där ankaret borde stå. Förut kunde en sådan träff långt fram ta sökningen förbi resten av ankarna: i metodriggens kompendium Faktatextsamtal i grupp (281 sidor, marginalens Ord i texten bredvid faktatexterna) hittades 82 av 141 ankare i Word, ett av dem på fel sida, och svaret blev NEJ fast Word och Google hade samma sidor. Nu hittas alla 141 i båda, på samma sidor, och svaret är ok. Temafilen Vulkaner och jordbävningar gick från NEJ till ok på samma sätt. Prövat med tre nya prov i `test/jamfor.test.js` och med 30 jämförelser som riggen hade sparade, gammal och ny version sida vid sida utan uppladdning: de 28 som var ok är ok, och de två falska nej är ok. Från metodriggen 2026-10-19.
- `ankarsidor` och `NARA` exporteras ur `src/jamfor.js`, så att proven når dem utan Word och Google.

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
