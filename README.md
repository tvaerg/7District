# 7D Portfolio Aggregation Engine

Det här programmet tar in månads- och kvartalsrapporter från portföljbolagen (PDF, PowerPoint, Excel, Word) och bygger automatiskt en färdig, branded styrelse-deck i PowerPoint, med en komplett heatmap och en uppdateringsslide per bolag.

Du kör ett kommando och får tillbaka en .pptx som du granskar och redigerar direkt i PowerPoint. Programmet föreslår RAG-färger (grön/gul/röd) och text, du har sista ordet.

## 1. Vad du behöver först

Innan du börjar, se till att du har:

**Python 3.10 eller senare** installerat. Kontrollera med:

```
python --version
```

Om du inte har Python, ladda ner det från https://www.python.org/downloads/ (på Windows: kryssa i "Add Python to PATH" under installationen).

**Git** installerat, för att hämta koden. Kontrollera med:

```
git --version
```

Saknas det, ladda ner från https://git-scm.com/downloads.

**Två egna API-nycklar** (se avsnitt 4). Programmet anropar två AI-tjänster, en för att läsa PDF:er och en för analysen. Du måste sätta upp dina egna nycklar, de gamla nycklarna från utvecklingen ska inte användas i drift.

## 2. Hämta koden från GitHub

Det enklaste sättet på Windows är **GitHub Desktop**. Det sköter inloggningen åt dig grafiskt, så du slipper allt strul med lösenord och tokens i terminalen. Ladda ner från https://desktop.github.com, logga in med ditt GitHub-konto, och klona repot via "File > Clone repository". Adressen till repot får du av Tom. Behöver du åtkomst till ett privat repo, be Tom bjuda in dig som collaborator först.

Föredrar du terminalen går det också. Öppna en terminal (Terminal på Mac, eller PowerShell/Command Prompt på Windows), gå till mappen där du vill ha projektet och kör:

```
git clone <REPO_URL> 7d-engine
cd 7d-engine
```

Byt ut `<REPO_URL>` mot adressen till repot (den ser ut ungefär som `https://github.com/<anvandare>/<repo>.git`).

Viktigt om repot är privat och du använder terminalen: GitHub accepterar **inte** ditt vanliga lösenord när Git frågar efter inloggning. Du måste antingen använda GitHub Desktop (rekommenderas, se ovan) eller låta Tom skapa en Personal Access Token åt dig som du klistrar in i stället för lösenordet. Om inloggningen nekas trots att du skrev rätt lösenord är det här orsaken.

Senare, när det kommit uppdateringar i koden, hämtar du dem med `git pull` i terminalen, eller med "Pull origin" i GitHub Desktop.

## 3. Installera Python-beroenden

Det är bäst att köra programmet i en egen, isolerad Python-miljö så att det inte krockar med annat på datorn.

I projektmappen, skapa och aktivera en virtuell miljö:

**Mac / Linux:**

```
python -m venv venv
source venv/bin/activate
```

**Windows (PowerShell):**

```
python -m venv venv
venv\Scripts\Activate.ps1
```

Om PowerShell svarar med ett fel om att skript är blockerade ("running scripts is disabled on this system"), kör först denna rad och sedan activate-kommandot igen:

```
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Den gäller bara det aktuella terminalfönstret och ändrar ingenting permanent på datorn.

**Windows (Command Prompt), enklaste alternativet:** vill du slippa PowerShells skriptregler helt, öppna Command Prompt i stället och kör:

```
python -m venv venv
venv\Scripts\activate.bat
```

När miljön är aktiv ser du `(venv)` längst fram i terminalraden. Installera sedan alla beroenden:

```
pip install -r requirements.txt
```

Varje gång du öppnar en ny terminal för att köra programmet, aktivera miljön igen med samma activate-kommando som ovan.

## 4. Skaffa dina API-nycklar

Programmet använder två tjänster. Du behöver ett konto och en nyckel hos båda.

**A) Anthropic (Claude / Opus), analyslagret**

1. Gå till https://console.anthropic.com och skapa eller logga in på ett konto.
2. Lägg till en betalningsmetod under Billing (analysen kostar pengar per körning, se avsnitt 9).
3. Gå till API Keys, klicka Create Key, och kopiera nyckeln. Den börjar med `sk-ant-`.
4. Spara nyckeln direkt, du kan bara se den en gång.

**B) Google Gemini, PDF-inläsningen**

1. Gå till https://aistudio.google.com/app/apikey och logga in med ett Google-konto.
2. Klicka Create API key och kopiera nyckeln.
3. Spara den.

Behandla båda nycklarna som lösenord. Lägg aldrig in dem direkt i koden och dela dem aldrig i e-post eller chatt. De ska bara ligga i `.env`-filen (nästa avsnitt), som aldrig läggs upp på GitHub.

## 5. Skapa din .env-fil

Programmet läser nycklarna från en fil som heter `.env` i projektmappen. Den filen skapar du själv en gång.

Innehållet ska se ut exakt så här, med dina egna nycklar inklistrade:

```
ANTHROPIC_API_KEY=sk-ant-din-nyckel-har
GEMINI_API_KEY=din-gemini-nyckel-har
```

Det säkraste sättet att skapa filen är från terminalen, så att den garanterat får rätt namn:

**Windows (PowerShell):**

```
@"
ANTHROPIC_API_KEY=sk-ant-din-nyckel-har
GEMINI_API_KEY=din-gemini-nyckel-har
"@ | Out-File -FilePath .env -Encoding ascii
```

**Mac / Linux:**

```
cat > .env << 'EOF'
ANTHROPIC_API_KEY=sk-ant-din-nyckel-har
GEMINI_API_KEY=din-gemini-nyckel-har
EOF
```

Varför terminalkommandot rekommenderas på Windows: skapar du filen manuellt i Anteckningar lägger Windows ofta på en dold `.txt`-ändelse, så att filen i själva verket heter `.env.txt` i stället för `.env`. Då hittar programmet inte nycklarna och kraschar med felet "ANTHROPIC_API_KEY missing", trots att allt ser rätt ut. Eftersom Utforskaren döljer filändelser som standard ser du inte misstaget. Skapar du filen med kommandot ovan slipper du hela problemet.

Ett par saker att veta:

- Inga citationstecken och inga mellanslag runt `=`-tecknet.
- Filen ska ligga i projektmappens rot, alltså bredvid `sevend_pipeline.py`.
- `.env` ska aldrig checkas in i Git. Kontrollera att den står i `.gitignore`. Om den inte gör det, lägg till en rad med bara `.env` i `.gitignore`.

**Valfria inställningar**

Om du senare vill byta vilken AI-modell som används kan du lägga till dessa rader i `.env`. De har bra standardvärden, så du behöver normalt inte röra dem:

```
CLAUDE_MODEL=claude-opus-4-7
GEMINI_MODEL=models/gemini-3.1-pro-preview
```

## 6. Mappstruktur

Programmet förväntar sig en liten mappstruktur. Den ska redan finnas i repot, men för säkerhets skull:

```
7d-engine/
├── .env                      <- dina nycklar (skapar du själv, se avsnitt 5)
├── requirements.txt
├── sevend_pipeline.py        <- programmet du kör
├── sevend_*.py               <- ovriga moduler
├── assets/
│   └── 7d_template.pptx      <- 7D:s PowerPoint-mall (decken byggs ovanpå denna)
├── input/                    <- lagg bolagens rapporter har (valfritt, se nedan)
└── output/                   <- de fardiga deckarna hamnar har
```

Mallen i `assets/7d_template.pptx` är viktig. Det är den som ger decken rätt logotyper, färger och layout. Rör den inte om du inte medvetet vill ändra designen.

## 7. Köra programmet

Lägg ihop alla månadens rapportfiler och peka programmet mot dem. Du anger också vilken period det gäller med `--cycle`.

Grundkommandot:

```
python sevend_pipeline.py <filer> --cycle "April 2026"
```

Ett verkligt exempel, där alla filer ligger i `input/`-mappen:

```
python sevend_pipeline.py input/*.pdf input/*.pptx input/*.xlsx input/*.docx --cycle "April 2026"
```

Du kan också peka ut filer var de än ligger:

```
python sevend_pipeline.py "MR 03-2026.pdf" "Kvartalsrapport_Kvaser.pptx" --cycle "April 2026"
```

En notering för Windows: stjärnan (`input/*.pdf`) fungerar inte alltid likadant i PowerShell som på Mac/Linux. Om en körning inte hittar några filer trots att de ligger i `input/`, lista filerna var för sig i kommandot i stället för att använda stjärnan.

Programmet identifierar automatiskt vilket bolag varje fil tillhör genom att läsa innehållet, du behöver alltså inte döpa om filerna. Flera filer för samma bolag (t.ex. en resultatrapport och en balansrapport) slås ihop automatiskt.

**Tillval**

- `--cycle "April 2026"` (obligatoriskt), perioden, syns på decken.
- `--ceo-update "..."`, en valfri VD-rad som hamnar på heatmap-sliden.
- `--out output/min_deck.pptx`, styr filnamnet på resultatet. Utelämnar du det får decken automatiskt ett namn med period och tidsstämpel, så att varje körning sparas i stället för att skriva över den förra.
- `--template assets/7d_template.pptx`, sökväg till mallen, behöver normalt inte anges.

När körningen är klar skrivs den färdiga .pptx-filen till `output/`-mappen, och sökvägen visas längst ner i terminalen.

## 8. Granska och redigera resultatet

Öppna den färdiga decken i PowerPoint. Allt är vanliga, redigerbara element:

- Texten ändrar du direkt i PowerPoint.
- RAG-färgerna i heatmappen är riktiga tabellceller. Håller du inte med om en färg klickar du på cellen och byter fyllningsfärg.
- Motiveringen till varje färg som programmet föreslog ligger sparad i sidans talanteckningar (speaker notes), som ett spår att gå tillbaka till.

Programmet föreslår, du bestämmer. Inget av det automatiken sätter är låst.

Bredvid varje deck skapas också en `..._debug.json`-fil. Den är till för felsökning (se avsnitt 10) och behöver inte skickas vidare.

## 9. Vad en körning kostar

Varje körning gör betalda AI-anrop: ett analysanrop per bolag som rapporterat (Opus) och ett läsanrop per PDF (Gemini). En normal månadskörning av hela portföljen landar typiskt på några få dollar totalt. Du ser den faktiska förbrukningen i Anthropics och Googles respektive konsoler. Lägg gärna en månadsbudget/spend-gräns i Anthropic-konsolen så du har koll.

## 10. Felsökning

**"ANTHROPIC_API_KEY missing" eller liknande nyckel-fel**
`.env`-filen hittas inte eller saknar nyckeln. Kontrollera att filen heter exakt `.env`, ligger i projektmappens rot, och innehåller rätt rader utan mellanslag runt `=`. På Windows: kontrollera att filen inte i själva verket heter `.env.txt` (se avsnitt 5). Slå på "Filnamnstillägg" i Utforskaren under fliken Visa om du vill se den verkliga ändelsen.

**PowerShell vägrar aktivera den virtuella miljön**
Felet "running scripts is disabled on this system" betyder att skript är blockerade. Kör `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` först, eller använd Command Prompt med `venv\Scripts\activate.bat` (se avsnitt 3).

**Git nekar inloggning fast lösenordet är rätt**
GitHub accepterar inte vanligt lösenord i terminalen. Använd GitHub Desktop, eller en Personal Access Token från Tom (se avsnitt 2).

**Ett bolag saknas eller hamnar som "No report"**
Programmet kunde inte koppla någon fil till bolaget. Kontrollera att filen verkligen var med i kommandot. Om ett bolag konsekvent inte känns igen kan en alias behöva läggas till i `sevend_registry.py`, hör av dig till Tom så hjälper han till.

**En körning ser konstig ut**
Öppna `..._debug.json` som skapades bredvid decken. Den visar per bolag exakt vilka filer som bidrog, vilka siffror som lästes ut, vad analysen föreslog, och vad eventuell QA ändrade. Det är den snabbaste vägen att se var något gick snett.

**`pip install` eller `python` hittas inte**
Den virtuella miljön är förmodligen inte aktiverad. Kör activate-kommandot från avsnitt 3 igen, du ska se `(venv)` i terminalraden.

**Mallen ser fel ut i resultatet**
Kontrollera att `assets/7d_template.pptx` finns och inte har ändrats av misstag.

## 11. Löpande underhåll

Det enda du normalt behöver underhålla är portföljregistret i `sevend_registry.py`. Där ligger 7D:s egen styrningsdata som inte kommer från rapporterna: vilka bolag som finns, vilken kategori de tillhör (Core / Non-core / Under divestment), lead/support-initialer, exit-timing och investeringstes.

När portföljen ändras (ett bolag avyttras, ett nytt tillkommer) uppdaterar du den filen. Strukturen är väl kommenterad. Är du osäker, hör av dig till Tom innan en månadskörning så går ni igenom det tillsammans.

Frågor under överlämningen: kontakta Tom.
