# Tildeling af sygeplejeartikler til terminale borgere (§ 26)

Automatisering der tildeler sygeplejeartikler til terminale borgere i Odense Kommune i henhold til Ældrelovens § 26.

## Hvad gør robotten?

1. **Henter opgaver** fra aktivitetslisten "Opgaver - 6 mdr tilbage til 6 mdr frem" i KMD Nexus under organisationen "Sygeplejehjælpemidler"
2. **Filtrerer** på opgaver tildelt robotbrugeren og med `§26` i opgavebeskrivelsen
3. **Fylder arbejdskøen** med relevante opgaver inkl. borgerens CPR-nummer og opgave-id
4. **Validerer** for hvert køelement om borgeren opfylder tildelingskriterierne:
   - Borgeren bor i Odense Kommune
   - Borgeren har en gyldig terminalerklæring
   - Borgeren bor ikke på plejehjem eller bosted (jf. regelsæt)
   - Borgeren har ikke en aktiv indsats under en af de accepterede paragraffer (§ 83, § 94, § 95, § 119, § 138 m.fl.)
5. **Opretter** forløb, sagsnotat og indsats ("Sygeplejeartikler - ÆL § 26") i KMD Nexus, hvis alle kriterier er opfyldt
6. **Genererer og sender** et afgørelsesbrev som Digital Post til borgeren via SBSip
7. **Uploader** afgørelsesdokumentet til KMD Nexus med tagget "ÆL § 26"
8. **Opretter en opgave til personalet**, hvis et af kriterierne ikke er opfyldt, med angivelse af årsagen
9. **Afslutter** den udløsende opgave

## Forudsætninger

- Python ≥ 3.13
- [`uv`](https://docs.astral.sh/uv/) til pakkehåndtering
- Adgang til **Automation Server** (arbejdskø)
- Adgang til **KMD Nexus** (produktion)
- Adgang til **SBSip** (Digital Post)
- Adgang til **Datafordeler** (adresseoplysninger) inkl. gyldige certifikater
- En **Odense SQL Server**-konto til tracking

## Installation

```sh
uv sync
```

## Konfiguration

Kopiér `.env.example` til `.env` og udfyld følgende:

| Variabel | Beskrivelse |
|---|---|
| `EXCEL_MAPPING_PATH` | Sti til `Regelsæt.xlsx` (kan også angives via `--excel-file`) |
| `LETTER_TEMPLATE_PATH` | Sti til Word-brevskabelonen (kan også angives via `--word-template`) |
| `CERTIFIKATER` | Sti til mappe med Datafordeler-certifikater (standard: `/certifikater`) |
| *(Automation Server-variabler)* | Ifølge `automation-server-client`-dokumentationen |

Credentials til KMD Nexus, SBSip og SQL Server hentes automatisk fra Automation Server under kørsel.

## Kørsel

```sh
# Fyld arbejdskøen med opgaver fra KMD Nexus
uv run python main.py --queue

# Behandl arbejdskøen
uv run python main.py --excel-file "Regelsæt.xlsx" --word-template "Tildeling af sygeplejeartikler til terminale borgere.docx"
```

### Argumenter

| Argument | Beskrivelse |
|---|---|
| `--excel-file <sti>` | Tilsidesæt stien til `Regelsæt.xlsx` |
| `--word-template <sti>` | Tilsidesæt stien til Word-brevskabelonen |
| `--queue` | Fyld arbejdskøen og afslut (kør ingen behandling) |

## Inputfiler

Følgende filer er nødvendige for kørsel, men er **ikke** inkluderet i repositoriet:

| Fil | Beskrivelse |
|---|---|
| `Regelsæt.xlsx` | Liste over plejehjemsliste og bosteder. Borgere tilknyttet disse ekskluderes fra automatisk tildeling. |
| `Tildeling af sygeplejeartikler til terminale borgere.docx` | Word-brevskabelon med flettefelter til generering af afgørelsesbrev. |

## Afhængigheder

| Pakke | Formål |
|---|---|
| `automation-server-client` | Arbejdskø-håndtering |
| `kmd-nexus-client` | Integration med KMD Nexus |
| `odk-tools` | Aktivitetssporing |
| `sbsip` | Afsendelse af Digital Post |
| `datafordeler` | Opslag af adresseoplysninger til Digital Post |
| `openpyxl` | Læsning af Excel-regelsæt |
| `httpx` | Konvertering af Word-skabelon til PDF via render-service |

## Persondatasikkerhed

Robotten behandler følsomme personoplysninger på vegne af Odense Kommune, herunder CPR-numre og helbredsoplysninger (terminalerklæringer), der udgør særlige kategorier jf. GDPR art. 9.

- Ingen personoplysninger må lægges i dette repository — hverken som testdata, i kode eller i kommentarer
- `input/`-mappen er ekskluderet via `.gitignore` og må aldrig committes
- Legitimationsoplysninger håndteres udelukkende via miljøvariabler (`.env`) og Automation Server Credentials
- Den genererede PDF slettes lokalt, når den er uploadet til KMD Nexus og afsendt til borgeren

