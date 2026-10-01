# wordparitet

Word och Google Dokument, delat mellan niclasfohlin.se och metodriggen. Modulen samlar det som får en Word-fil att se likadan ut i Word och i Google Dokument: reglerna med sina skäl, regelprovet, mätbänken och jämförelserna. Ingen av riggarna äger den; båda använder och utvecklar den (Niclas 2026-10-01: "Ni båda ska kunna använda den och även kunna utveckla den vid behov").

Reglerna och arbetsgången står i [REGLER.md](REGLER.md), och vad som ändrats i [CHANGELOG.md](CHANGELOG.md).

## Delarna

| Del | Vad den gör | Kräver |
|---|---|---|
| `provaWordfil` (`src/regler.js`), `wordparitet regler` | Prövar en Word-fil mot reglerna och ger brotten, med undantag som har skäl | jszip |
| `korBank` (`wordparitet/bank`), `wordparitet bank`, proven i `bank/` | Mätbänken: Word och Google ritar varianter av en form, och radstegen mäts i båda | docx, Word, Google, Poppler |
| `jamforWordGoogle` (`src/jamfor.js`), `wordparitet google` | Jämför Words och Googles sidor av en Word-fil och ritar ett översiktsark | Word, Google, Poppler, sharp |
| `wordFacit`, `wordJamfor` (`src/jamfor.js`) | Ändrar en ändring i koden hur Word ritar filen? Före och efter, i text och bild | Word, Poppler, sharp |
| `googlePdf`, `testmappen` (`src/google.js`) | Google Dokuments pdf av en Word-fil, tillfälligt eller i en testmapp | Google |
| `wordPdf` (`src/word.js`, `ps/word-pdf.ps1`) | Words pdf av en Word-fil | Windows med Word |

Byggstenarna (radavstånd, tabellrader i hela bildpunkter, stycken runt tabeller i celler, kolumnbredd efter längsta ord, dubbla linjer och boksidan) flyttar in i steg 2, som en fabrik `byggstenar(docx)` som får den som anropars docx.

## Så används den

Sajten pekar på en tagg, så att Netlify får samma version som är prövad: `"wordparitet": "https://github.com/niclasfohlin/wordparitet/archive/refs/tags/v0.1.0.tar.gz"`. Riggen pekar på klonen bredvid sig: `"wordparitet": "file:../wordparitet"`.

```js
import { provaWordfil, brottSomText } from 'wordparitet';
const brott = await provaWordfil(readFileSync('fil.docx'), { filnamn: 'fil.docx', undantag: [] });
```

Kommandona: `npx wordparitet regler <fil.docx> …`, `npx wordparitet bank <prov> [--facit]`, `npx wordparitet google <fil.docx> … [--mapp <namn>]`.

Google-delarna behöver en inloggning med behörigheten drive.file: `{ inloggning }` (ett objekt med `token()` eller sökvägen till en modul med `token()`), eller miljövariabeln `WORDPARITET_GOOGLE`. Nycklar skrivs aldrig ut och ligger aldrig i förrådet. Sajten har inloggningen; riggen når inte Google och ber sajten köra en mätning.

## Arbetssättet

- `npm test` körs utan nätverk, före varje commit. Varje regel prövas mot en liten Word-fil som bryter den och en som klarar den.
- Båda riggarna committar i klonen `C:\wordparitet`, med `git status` före varje ändring och små commits med meddelanden på svenska.
- En ny skillnad mäts i bänken och blir en regel med skäl, ett prov och en byggsten. Versionen höjs, `CHANGELOG.md` säger vad, varför och var det är mätt, den som höjer taggar och pushar, och den andra riggen får ett meddelande.
- Ett undantag från en regel anges av den som provar, alltid med skäl, aldrig tyst.
