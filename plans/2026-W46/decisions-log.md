# Decisions log 2026-W46: Invoke-Kvittohantering.ps1

Överlämning från Arkitekt-sessionen (2026-11-12) till Utförar-sessionen.

## Läs detta först

- Granskningen gjordes först på scriptet som det klistrades in i chatten. Filen ligger nu
  i repot som `Kvittohantering/Invoke-Kvittohantering.ps1` (commit 193a327). Radnummer i
  raderna **Rad** gäller den versionen.
- Uppföljning samma dag med pwsh 7.4.6 och PSScriptAnalyzer 1.23.0: enskilda påståenden
  testkördes med lokala snuttar, utan nätverk mot kedjorna. Poster som testats har raden
  **Testkört**. Scriptet självt är inte kört. Testmiljön är Linux i tidszon UTC, så
  macOS-delarna (rättigheter, nyckelring, Playwright) är inte testade.
- PSScriptAnalyzer på hela scriptet: inga tolkningsfel, 3 anmärkningar (se E-04).
  Diffsnuttarna för D-02 och D-03 ger inga anmärkningar.
- **Underlag** per post: `kod` = följer direkt av koden, `kod+.NET` = följer av koden plus
  dokumenterat .NET/Playwright-beteende som inte kunnat köras här, `antagande` = beror på
  hur kedjan faktiskt svarar.
- Hårda regler som gäller alla poster: inga nya cookies/tokens/kvittodata på disk, råkvitton
  i RawPath raderas eller skrivs aldrig över, svenska kommentarer, PS 7.4+, ingen inloggning
  eller nedladdning under arbetet. Poster som rör vid en regel har raden **Flagga**.
- Ta posterna i den ordning de står. Alla har status `föreslagen`.

## Akut: regression efter 1.3.0

### L-04 Låst Playwright-version går inte att installera på Python 3.9
- Status: genomförd i 1.4.1 (versionskontrollen testad lokalt; installation på Python 3.9 ej testad, ingen 3.9 här) | Område: leveranskedja | Allvarlighet: hög (blockerar `-AutoCookie` och Lidl-PDF) | Underlag: användarens felutskrift + PyPI-metadata
- Var: `$script:PlaywrightVersion = '1.65.1'` (rad 156), `Initialize-Kvittowebblasare` (rad 525-544)
- Symptom: `No matching distribution found for playwright==1.65.1`; pip listar versioner till och med 1.60.0.
- Orsak: Playwright 1.61.0 och senare kräver Python 3.10 (`requires_python >=3.10`), 1.60.0
  är sista versionen för 3.9. venv:en är skapad med `python3` från PATH, som på macOS
  normalt är Xcodes Python 3.9. Versionen 1.65.1 finns på PyPI, men pip visar bara det som
  passar tolken. Markörfilen saknade versionen, så 1.3.0 tvingade fram en ominstallation
  av en installation som fungerade.
- Åtgärd nu (egen liten PR, 1.4.1): lås till 1.60.0 och ge ett begripligt fel om tolken är för gammal.

```diff
-$script:PlaywrightVersion = '1.65.1'
+# 1.60.0 är sista versionen som stöder Python 3.9 (Xcodes python3). 1.61+ kräver Python 3.10.
+$script:PlaywrightVersion   = '1.60.0'
+$script:PlaywrightMinPython = [version]'3.9'
```
```diff
         if ($LASTEXITCODE -ne 0) { throw "python3 -m venv misslyckades (kod $LASTEXITCODE)" }
     }
+    # pip listar bara versioner som passar tolken, så felet blir annars obegripligt
+    $pyVersion = [version](& $python -c 'import sys; print("%d.%d" % sys.version_info[:2])')
+    if ($pyVersion -lt $script:PlaywrightMinPython) {
+        throw "Playwright $($script:PlaywrightVersion) kräver Python $($script:PlaywrightMinPython)+, men venv har $pyVersion."
+    }
     Write-Information 'Installerar Playwright...' -InformationAction Continue
```
- Klart när: `-AutoCookie` startar på en Mac där `python3 --version` ger 3.9.x, utan att
  AppDir raderas först.
- Lärdom för kommande låsningar: kontrollera vald version med venv:ens egen pip
  (`& $python -m pip index versions playwright`), inte mot PyPI i allmänhet.
- Uppföljning (medel, eget beslut): Python 3.9 är utgånget och 1.60.0 fryser webbläsaren
  som används för BankID. Se F-07.

## Att verifiera innan något annat (V)

### V-01 Klistringsartefakter i underlaget
- Status: avklarad, inget fel i filen | Allvarlighet: låg (kvarvarande del)
- Kontrollerat i repot: rad 596 och 698 har `'www.willys.se'` / `'www.lidl.se'`. Rad 109 har
  blanksteg, punkt, U+00A0 och U+202F som osynliga literaltecken.
- Kvarstår: skriv de två osynliga tecknen som `\u00A0\u202F`, så att de syns i diffar och
  överlever kopiering. Kan göras tillsammans med K-02.

## Prioritet 1: hög

### D-01 Normaliseringen skriver över månadsfiler med ofullständigt underlag
- Status: genomförd i 1.0.1 (testad lokalt med -SkipDownload: saknad pdftotext lämnar DestinationPath orörd; spärrad månad och manad.* skrivs inte) | Område: dataintegritet | Underlag: kod
- Var: huvudflödet, steg 2 (`if (-not $pdfToText) { Write-Warning ... }` och `catch`-blocken
  runt `ConvertFrom-WillysKvitto`/`ConvertFrom-LidlKvitto`)
- Problem: alla månadsmappar skrivs om från noll vid varje körning. Saknas pdftotext hoppas
  samtliga Willys-kvitton över med en varning, och sedan skrivs `kvitton.json`, `varor.csv`
  och `manad.*` för hela historiken utan Willys. Samma sak i liten skala när ett enskilt
  kvitto inte går att tolka.
- Åtgärd:
  1. Saknas pdftotext och det finns Willys-PDF:er: `throw` innan något skrivs.
  2. Samla misslyckade filer. Månaden går att läsa ur filnamnets datumprefix; hoppa över
     skrivning av de månaderna och av `manad.*`, och varna med fillistan.

```diff
     $pdfToText = Find-PdfToText -Hint $PdfToTextPath
     if (-not $pdfToText) {
-        Write-Warning "pdftotext saknas - $($pdfFiles.Count) Willys-kvitton hoppas över. Installera med: brew install poppler"
+        # Avbryt före skrivning: annars skrivs alla månadsfiler om utan Willys-kvitton
+        throw "pdftotext saknas - $($pdfFiles.Count) Willys-kvitton kan inte tolkas. Installera med: brew install poppler"
     }
-    else {
-        Write-Verbose "Använder pdftotext: $pdfToText"
-        foreach ($pdf in $pdfFiles) {
-            try {
-                $receipts.Add((ConvertFrom-WillysKvitto -File $pdf -PdfToText $pdfToText))
-            }
-            catch {
-                Write-Warning "Willys $($pdf.Name): $($PSItem.Exception.Message)"
-            }
+    Write-Verbose "Använder pdftotext: $pdfToText"
+    foreach ($pdf in $pdfFiles) {
+        try {
+            $receipts.Add((ConvertFrom-WillysKvitto -File $pdf -PdfToText $pdfToText))
+        }
+        catch {
+            # Månaden (ur filnamnet) spärras för skrivning, så att en tidigare bra fil står kvar
+            [void]$blockedMonths.Add($pdf.Name.Substring(0, 7))
+            Write-Warning "Willys $($pdf.Name): $($PSItem.Exception.Message)"
         }
     }
```
  `$blockedMonths` är ett `HashSet[string]` som skapas före Willys-blocket, fylls på samma
  sätt i Lidl-`catch`, och kontrolleras i månadsloopen (`if ($blockedMonths.Contains($group.Name)) { varna; continue }`)
  samt före `manad.*`.
- Rad: 1234.
- Klart när: körning utan pdftotext lämnar DestinationPath orörd.

### D-02 Avbruten nedladdning lämnar en halv fil under slutnamnet
- Status: genomförd i 1.1.0 (testad lokalt med stubbad Invoke-KvittoRequest) | Område: dataintegritet | Underlag: kod
- Var: `Save-WillysKvitto` (`Invoke-KvittoRequest ... -OutFile $fileName`), `Save-LidlKvitto`
  (`WriteAllText($jsonFile ...)`)
- Problem: PDF:en strömmas direkt till slutnamnet. Bryts körningen eller anslutningen
  mitt i ligger filen kvar, nästa körning räknar den som "fanns redan", och pdftotext
  misslyckas på den för alltid. Det generella `catch`-blocket städar inte heller.
- Åtgärd: ladda ner till `<namn>.part`, verifiera, byt namn utan överskrivning.

```diff
         try {
-            Invoke-KvittoRequest -Uri $uri -Headers $headers -Session $session -OutFile $fileName
-            $bytes = [System.IO.File]::ReadAllBytes($fileName)
-            $magic = $bytes.Length -ge 4 ? [System.Text.Encoding]::ASCII.GetString($bytes, 0, 4) : ''
-            if ($magic -ne '%PDF') {
-                [System.IO.File]::Delete($fileName)
+            # Ladda ner till .part och byt namn först när filen är verifierad. Ett avbrott
+            # lämnar då aldrig en halv PDF under slutnamnet, som nästa körning skulle hoppa över.
+            $partFile = "$fileName.part"
+            Invoke-KvittoRequest -Uri $uri -Headers $headers -Session $session -OutFile $partFile
+            $head = [byte[]]::new(4)
+            $fs   = [System.IO.File]::OpenRead($partFile)
+            try { $read = $fs.Read($head, 0, 4) } finally { $fs.Dispose() }
+            if ($read -lt 4 -or [System.Text.Encoding]::ASCII.GetString($head) -ne '%PDF') {
+                [System.IO.File]::Delete($partFile)
                 throw [System.UnauthorizedAccessException]::new('Willys returnerade inte en PDF - sessionen har troligen gått ut.')
             }
+            [System.IO.File]::Move($partFile, $fileName)   # utan överskrivning: kastar om slutnamnet finns
             $result.Nya++
```
  Lidl: skriv HTML först (bara om den inte finns), sedan JSON till `.part` + `Move`. JSON-filen
  är markören för "klart", så den ska komma sist.
- Flagga: `.part`-filer är tillfälliga filer i RawPath. De skrivs över vid nästa försök, men
  de är aldrig färdiga råkvitton. Slutnamn skrivs aldrig över (`Move` utan overwrite).
- Rad: 647 (Willys), 752 (Lidl).
- Klart när: Ctrl+C mitt i en nedladdning ger ingen `*.pdf`/`*.json` för det kvittot.

### D-03 Två kvitton med samma datum, butik och belopp: det andra hämtas aldrig
- Status: genomförd i 1.1.0 (testad lokalt med stubbad Invoke-KvittoRequest) | Område: dataintegritet | Underlag: kod (hur ofta det händer: antagande)
- Var: `Save-WillysKvitto` och `Save-LidlKvitto` (`if (Test-Path -LiteralPath ...) { $result.Fanns++; continue }`)
- Problem: filnamnet är enda identiteten. Två köp samma dag i samma butik på samma summa
  ger samma namn, det andra räknas som "fanns redan" och försvinner tyst.
- Åtgärd: räknare per körning; andra förekomsten i listan får `_2`, tredje `_3`. Listan är
  sorterad, så tilldelningen blir densamma mellan körningar. Befintliga filer behåller namn.

```diff
+    $nameCount = @{}   # samma datum+butik+belopp flera gånger i listan -> _2, _3 ...
     foreach ($t in $withReceipt) {
 ...
-        $fileName = Join-Path $OutputDir "${date}_${store}_${amount}kr.pdf"
+        $baseName = "${date}_${store}_${amount}kr"
+        $nameCount[$baseName] = 1 + [int]$nameCount[$baseName]
+        if ($nameCount[$baseName] -gt 1) { $baseName += "_$($nameCount[$baseName])" }
+        $fileName = Join-Path $OutputDir "$baseName.pdf"
```
  Följdändring i `ConvertFrom-WillysKvitto`: filnamnsregexen ska sluta på `kr(?:_\d+)?\.pdf$`.
  Samma räknare i `Save-LidlKvitto`.
- Rad: 639 (Willys), 746 (Lidl). Testkört: räknaren fungerar under StrictMode (`x, x, y, x` ger `x, x_2, y, x_3`).
- Flagga: filnamnsschemat utökas (bakåtkompatibelt). Stabiliteten förutsätter att båda
  kvittona ligger i samma datumintervall, vilket de gör eftersom datumet ingår i namnet.

### S-01 Cookie-värde kan skrivas ut i varning
- Status: föreslagen | Område: säkerhet | Underlag: kod, bekräftat i pwsh | Rad: 372
- Var: `New-CookieSession` (`Write-Warning "Cookie '$name' kunde inte läggas till: $($PSItem.Exception.Message)"`)
- Problem: `CookieContainer.Add` kastar `CookieException` med texten
  `The 'Value'='<värdet>' part of the cookie is invalid.` när värdet innehåller komma eller
  semikolon utan citattecken. Värdet hamnar då i terminalen och i en eventuell transcript.
- Åtgärd:

```diff
-            Write-Warning "Cookie '$name' kunde inte läggas till: $($PSItem.Exception.Message)"
+            # Inte Exception.Message: CookieException citerar cookiens värde i klartext
+            Write-Warning "Cookie '$name' kunde inte läggas till (ogiltigt tecken i värdet)."
```
- Testkört: `[System.Net.CookieContainer]::new().Add([System.Net.Cookie]::new('a','HEMLIGT,VARDE','/','example.com'))`
  ger meddelandet `The 'Value'='HEMLIGT,VARDE' part of the cookie is invalid.` Samma för semikolon.
- Bieffekt att känna till: sådana cookies skickas inte alls. Är det en sessionscookie
  ser det ut som utgången session.

## Prioritet 2: medel

### U-01 Lidl-HTML renderas med script och nätverk påslaget
- Status: genomförd i 1.3.0 (testad lokalt) | Område: opålitlig data | Underlag: kod
- Var: Python `render_receipt` / `cmd_pdf` (`page.set_content(doc, wait_until="load")`)
- Problem: innehållet i `<pre>` klistras in oescapat. `<script>`, `<img src=...>`, `<link>`
  och `<meta refresh>` i kvittotexten körs respektive hämtas. Sidan har ingen åtkomst till
  inloggningsprofilen (ny kontext, about:blank), så skadan begränsas till externa anrop,
  ändrat kvittoinnehåll och hängande `load`.
- Åtgärd:

```diff
-            page = browser.new_page()
+            # Kvitto-HTML är opålitlig data: inga script och inget nätverk vid rendering
+            ctx = browser.new_context(java_script_enabled=False, offline=True)
+            ctx.route("**/*", lambda route: route.abort())
+            page = ctx.new_page()
```
- Rad: 245.
- Testa att `set_content` och `bounding_box` fungerar med JS avstängt (Playwright kör sin
  egen kod via CDP, så det ska fungera, men det är inte kört här).
- Relaterat: `re.search(r"<pre>(.*?)</pre>")` tar bara första `<pre>`. Har Lidl flera block
  blir PDF:en avkortad (antagande, kontrollera mot ett sparat kvitto).

### S-02 Cookies i miljövariabler ärvs av alla barnprocesser
- Status: genomförd i 1.3.0 (testad lokalt) | Område: säkerhet | Underlag: kod
- Var: huvudflödet (`$willysCookie = $env:WILLYS_COOKIE`), alla `& $python`, `& $PdfToText`
- Problem: pip, Playwrights installationsprogram, Chromium och pdftotext (som tolkar
  opålitliga PDF:er) startas med cookies i sin miljö.
- Åtgärd: läs värdena, ta bort variablerna ur processmiljön under körningen och återställ
  dem i `finally` (samma mönster som `PLAYWRIGHT_BROWSERS_PATH`).

### S-03 AppDir skapas med standardrättigheter; profilen är i praktiken okrypterad
- Status: genomförd i 1.3.0 (testad lokalt) | Område: säkerhet | Underlag: kod (rättigheter), antagande (Playwrights standardflaggor)
- Var: `Initialize-Kvittowebblasare` (`CreateDirectory($script:AppDir)`)
- Problem: mappen får umask-rättigheter (normalt 755). Playwright startar Chromium med
  mock-nyckelring, så profilens cookies skyddas inte av macOS nyckelring.
- Åtgärd: `[System.IO.Directory]::CreateDirectory($script:AppDir, [System.IO.UnixFileMode]'UserRead, UserWrite, UserExecute')`
  och `[System.IO.File]::SetUnixFileMode` på befintlig mapp vid varje körning. Nämn i .NOTES
  att profilen följer med i Time Machine.

### L-01 Playwright är olåst, och en andra webbläsare uppgraderar bort den första
- Status: genomförd i 1.3.0 (testad lokalt) | Område: leveranskedja | Underlag: kod
- Var: `Initialize-Kvittowebblasare` (`pip install ... --upgrade playwright`, markörfil per webbläsare)
- Problem: versionen bestäms av installationsdagen. Installeras webkit senare uppgraderas
  paketet, Chromium-bygget matchar inte längre, men `.installerad-chromium` säger att allt
  är klart. Omvänt uppdateras webbläsaren aldrig, fast den används för BankID-inloggning.
- Åtgärd: konstant `$script:PlaywrightVersion`, `pip install "playwright==$v"`, skriv
  versionen i markörfilen och installera om när den skiljer sig. Hashlåsning: se F-03.

### K-01 Lidl-datum kan hamna på fel dag och månad
- Status: genomförd i 1.2.0 (testad lokalt) | Område: korrekthet | Underlag: kod, bekräftat i pwsh (beteendet); antagande (vilket format Lidl skickar) | Rad: 1003
- Var: `ConvertFrom-LidlKvitto` (`$dt.ToString('yyyy-MM-dd')`) jämfört med `Save-LidlKvitto` (`.ToLocalTime()`)
- Problem: `ConvertFrom-Json` ger `DateTime` med `Kind=Utc` för strängar som slutar på `Z`.
  Normaliseringen formaterar då UTC-tid, medan filnamnet bygger på lokal tid. Ett köp
  00:30 den första hamnar i föregående månad, och alla tider blir 1-2 timmar fel.
- Åtgärd: konvertera alltid till svensk tid före formatering, i både nedladdning och
  normalisering: `[System.TimeZoneInfo]::ConvertTime($dto, [System.TimeZoneInfo]::FindSystemTimeZoneById('Europe/Stockholm'))`.
  Det löser också D-05 (filnamn beror på datorns tidszon).
- Testkört (dator i UTC) med `ConvertFrom-Json`:
  - `...T22:30:00Z` blir `Kind=Utc` och formateras `2026-04-30 22:30`; i Stockholmstid är det `2026-05-01 00:30`, alltså annan månad.
  - `...T00:30:00+02:00` blir `Kind=Local` i datorns tidszon, så även datum med offset beror på var scriptet körs.
  - Utan tidszon blir det `Kind=Unspecified` och lämnas orört.

### K-02 Beloppsregexen saknar vänstergräns och förstår bara ASCII-minus framför talet
- Status: genomförd i 1.2.0 (testad lokalt) | Område: korrekthet | Underlag: kod, bekräftat i pwsh (.NET-regex) | Rad: 109-110
- Var: `$script:AmountPos`, `$script:Amount`
- Testresultat:
  - `123456,78` ger `23456,78` (första siffran tappas)
  - `Kapslar 100 129,00` ger `100 129,00` (100 129 kr i stället för 129)
  - `PRODUKT  12 199,00` ger namn `PRODUKT`, total `12 199,00`
  - `12,90-` och `−12,90` (U+2212) ger `12,90` utan tecken
- Åtgärd: lägg `(?<![\d,.])` först i `AmountPos`; tillåt `[-−]` som prefix och normalisera
  U+2212 i `ConvertFrom-SvensktTal`. Avslutande minus behövs inte för Willys/Lidl i dag
  men troligen för ICA (antagande), så gör det samtidigt som N-03.
  Sammanslagningen `100 129,00` går inte att lösa i regexen så länge Willys skriver
  `1 118,34`; kontrollsumman fångar den (se D-04 om att göra kontrollsumman synlig).
- Testkört: alla fyra resultaten ovan är desamma i .NET. Med `(?<![\d,.])` ger `123456,78`
  ingen träff alls (i stället för fel belopp), `0493686  19,90` ger `19,90` och `1 118,34`
  är oförändrat. `[double]::TryParse` klarar inte U+2212, så tecknet måste bytas före tolkning.

### D-04 Formatändring hos en kedja skriver över bra data med tomma kvitton
- Status: genomförd i 1.4.0 (testad lokalt) | Område: dataintegritet/korrekthet | Underlag: kod
- Var: `ConvertFrom-LidlKvitto` (`$reLine`), `ConvertFrom-WillysKvitto` (avdelare `^-+$`), kontrollsumman i huvudflödet
- Problem: Lidl-regexen kräver `id` före `class` och exakt klassnamn. Ändras markupen får
  varje kvitto noll varor, det blir en DIFF-varning per kvitto, och sedan skrivs alla
  månadsfiler över ändå. Kontrollsummans resultat sparas inte, så `kvitton.json` visar inte
  vilka kvitton som är opålitliga.
- Åtgärd: lägg `check_diff` (tal, 0 = stämmer) på varje kvitto i `kvitton.json`. Behandla
  kvitto med total men noll varor som tolkningsfel, så att D-01:s månadsspärr gäller.
- Flagga: nytt fält i det normaliserade schemat.

### D-05 Dubbletter mellan körningar
- Status: genomförd i 1.4.0 (testad lokalt) | Område: dataintegritet | Underlag: kod
- Var: filnamnsbygget i båda `Save-*`, steg 2 saknar dubblettkontroll
- Problem: namnet beror på datorns tidszon (`ToLocalTime`), butikens namn och hur beloppet
  formateras. Ändras något av dem hämtas samma kvitto igen under nytt namn och räknas dubbelt.
- Åtgärd: K-01 tar tidszonen. Lägg dubblettvarning i steg 2 på (kedja, datum, tid, total)
  och ta med bara det första. Radera aldrig råfilen.

### K-03 `$t.storeName` i felhanteringen kastar under StrictMode
- Status: genomförd i 1.2.0 (testad lokalt) | Område: korrekthet | Underlag: kod
- Var: `Save-WillysKvitto`, sista `catch` (`Write-Warning "Willys $date $($t.storeName): ..."`)
- Problem: saknas egenskapen kastar StrictMode inne i `catch`, och hela scriptet avbryts.
- Åtgärd: `$(Get-PropertyValue $t 'storeName')`, som Lidl-funktionen redan gör.
- Rad: 663. Testkört: `"$($t.storeName)"` på ett objekt utan egenskapen kastar under `Set-StrictMode -Version Latest`.

### K-04 Nätverksfel i listfasen avbryter hela scriptet före normaliseringen
- Status: genomförd i 1.2.0 (testad lokalt) | Område: korrekthet | Underlag: kod
- Var: `Invoke-KvittoRequest` (fångar bara `HttpResponseException`), huvudflödet (fångar bara `UnauthorizedAccessException`)
- Problem: timeout eller DNS-fel försöks inte om och fångas inte, så steg 1b och 2 körs
  aldrig. `Retry-After` används utan tak; servern kan beordra en mycket lång väntan.
- Åtgärd: försök om även `HttpRequestException`/`TaskCanceledException`, sätt tak på
  väntan (t.ex. 60 s), och fånga övriga fel per kedja i huvudflödet med varning.

### K-05 Lidl-listan avbryts efter sida 1 om `totalCount` saknas
- Status: genomförd i 1.2.0 (testad lokalt) | Område: korrekthet | Underlag: kod (att fältet kan försvinna: antagande)
- Var: `Save-LidlKvitto` (`if ($allOlder -or $seen -ge $total) { break }`)
- Problem: saknat fält blir 0, och `$seen -ge 0` är alltid sant.
- Åtgärd: `($total -gt 0 -and $seen -ge $total)`.
- Rad: 725. Testkört: `[int]($null ?? 0)` ger 0 och `10 -ge 0` är sant.

### K-06 Bara den ena miljövariabeln satt: den andra kedjan hoppas över tyst
- Status: genomförd i 1.2.0 (testad lokalt) | Område: korrekthet | Underlag: kod
- Var: huvudflödet (`if (IsNullOrWhiteSpace($willysCookie) -and IsNullOrWhiteSpace($lidlCookie))`)
- Åtgärd: fråga per kedja som saknar värde. Görs enklast i kedjetabellen (N-01).

### D-06 Lidls "råkvitto" är omserialiserat
- Status: föreslagen | Område: dataintegritet | Underlag: kod
- Var: `Save-LidlKvitto` (`ConvertTo-Json -InputObject $detail -Depth 20`)
- Problem: svaret tolkas och skrivs ut igen. Datumsträngar blir `DateTime` och skrivs i
  annat format, och djup över 20 kapas. Filen är inte det kedjan skickade.
- Åtgärd: hämta med `Invoke-WebRequest`, spara `.Content` oförändrat, tolka en kopia för
  kontrollen. Gäller bara nya nedladdningar.
- Flagga: samma data som i dag, men byte-exakt. Befintliga filer rörs inte.
- Rad: 752. Testkört: `"2026-05-01T00:30:00+02:00"` skrivs tillbaka som `"2026-04-30T22:30:00+00:00"` på en dator i UTC.

## Prioritet 3: låg

### U-02 Beloppet i filnamnet saneras inte
- Status: föreslagen | Område: opålitlig data | Underlag: kod
- Var: båda `Save-*` (`"$(...)" -replace '[,\s]', '_'`); frågeparametrarna `storeId`, `source`, `memberCardNumber` i Willys-URI:n
- Problem: butiksnamnet går genom `Get-SafeName`, men beloppet gör det inte. Kommer det som
  sträng med `/` kan sökvägen lämna mappen (kräver i praktiken att en mellanmapp finns).
  Negativa belopp (`-45.5kr`) matchar inte filnamnsregexen i `ConvertFrom-WillysKvitto`.
- Åtgärd: `-replace '[^\d.-]', '_'` på beloppet, `[Uri]::EscapeDataString` på
  frågeparametrarna, `-?` i filnamnsregexen.

### U-03 Formelinjektion i varor.csv
- Status: föreslagen | Underlag: kod
- Var: `Export-KvittoMapp`
- Problem: varunamn som börjar med `=`, `+`, `-` eller `@` tolkas som formel i Excel.
- Åtgärd: prefixa sådana namn med `'` i CSV-raden (inte i JSON).

### S-04 Verbose kan visa medlemsnummer och kvittoreferens
- Status: föreslagen att avskrivas | Underlag: motbevisat i pwsh 7.4.6
- Testkört mot lokal server: `Invoke-WebRequest`/`Invoke-RestMethod -Verbose` skriver
  `Requested HTTP/1.1 GET with 0-byte payload` utan URI. Ingen åtgärd behövs.

### S-05 Cookie-strängar lever kvar i `$auto` och sessionsobjekten
- Status: föreslagen | Underlag: kod
- Problem: `$willysCookie = $null` släpper en referens av flera. Strängar är oföränderliga
  i .NET, så minnet går ändå inte att tömma. Kommentaren "Släpp cookies ur minnet" lovar mer
  än koden håller.
- Åtgärd: nolla även `$auto`, justera kommentaren. Ingen funktionell ändring.

### L-02 Hjälparen körs från skrivbar mapp med mappen först i sys.path
- Status: föreslagen | Underlag: kod
- Var: `Invoke-Kvittohjalpare`
- Åtgärd: starta med `& $python -I $helper ...` (isolerat läge: ingen `PYTHON*`-miljö,
  scriptmappen läggs inte i sökvägen). Ta bort raden som raderar `kvittocookie.py` när den
  gjort sitt. Alternativet utan fil på disk står i F-02.

### L-03 python3 och pdftotext via PATH
- Status: föreslagen | Underlag: kod
- Åtgärd: prova kända absoluta sökvägar först (`/usr/bin/python3`, `/opt/homebrew/bin/...`),
  PATH som reserv, och skriv vald sökväg med `Write-Verbose` även för python.

### D-07 Icke-atomära skrivningar i DestinationPath
- Status: föreslagen | Underlag: kod
- Problem: avbrott mitt i en skrivning ger en trasig fil och `manad.*` som inte matchar
  månadsmapparna. Nästa körning lagar det.
- Åtgärd: skriv till temporär fil i samma mapp och `File.Move(..., $true)`.

### D-08 -WhatIf: inga bestående skrivningar (ingen åtgärd krävs)
- Status: föreslagen (notering) | Underlag: kod
- Genomgånget: alla skrivningar till RawPath, DestinationPath och AppDir ligger bakom
  `ShouldProcess`. Två saker att känna till: utan `-SkipDownload` görs listanropen mot
  kedjorna ändå, och `ConvertFrom-WillysKvitto` skapar och raderar temporära filer.
- Åtgärd: nämn listanropen i hjälptexten för `-WhatIf`.

### K-07 Falska DIFF-varningar vid öresavrundning
- Status: föreslagen | Underlag: antagande
- Problem: kontantköp avrundas till hel krona; då skiljer totalen upp till 0,50 från radsumman.
- Åtgärd: kontrollera mot ett kontantkvitto innan något ändras.

### K-08 Rabattrader som kan tappas i Willys-tolkningen
- Status: föreslagen | Underlag: kod (vägarna finns), antagande (om kvitton ser ut så)
- Var: `ConvertFrom-WillysKvitto`
- Problem: oindragen rad som börjar med `Rabatt`/`Willys Plus`/`Prisned` hoppas över helt,
  och rabatt efter en viktvara utan viktrad läggs på en platshållare som sedan kastas.
- Åtgärd: lägg sådana rader som egen post med negativt belopp i stället för `continue`.

## Effektivitet och förkortning (E)

### E-01 Alla original kopieras om vid varje körning
- Status: föreslagen | Underlag: kod
- Var: `Export-KvittoMapp` (`File.Copy(..., $true)`)
- Åtgärd: hoppa över kopiering när målet finns med samma längd och ändringstid. Minskar
  också synkning mot iCloud/Paperless.

### E-02 Dubbelarbete kring Lidl-filerna
- Status: föreslagen | Underlag: kod
- Problem: `$lidlRawDir` och `$lidlDir` är samma sökväg, fillistan med samma filter byggs
  i `ConvertTo-LidlPdf` och i steg 2, och hela PDF:en läses för fyra byte (löses av D-02).
- Åtgärd: en variabel, en hjälpfunktion `Get-LidlRawFile`.

### E-03 Upprepningar som kan tas bort
- Status: föreslagen | Underlag: kod
- `-InformationAction Continue` på ca 15 ställen: sätt `$InformationPreference = 'Continue'` överst.
- `Get-SafeName`: `åäöÅÄÖ` är redan med i `\w`.
- Blocken för daterade och odaterade kvitton i steg 2 gör samma sak: en lokal funktion.
- `Write-CsvFile -Header` används bara för tom fil; behåll, men ta rubrikerna från första raden när rader finns (görs redan) och dokumentera det.
- Kör PSScriptAnalyzer efter varje ändring.

### E-04 PSScriptAnalyzer-anmärkningar i nuvarande version
- Status: föreslagen | Underlag: kört (PSScriptAnalyzer 1.23.0, standardregler)
- `PSUseBOMForUnicodeEncodedFile` (varning): filen har åäö men ingen BOM. pwsh 7 läser
  UTF-8 utan BOM, så det är ofarligt med `#Requires -Version 7.4`. Välj: lägg till BOM
  eller undanta regeln i en `PSScriptAnalyzerSettings.psd1`.
- `PSUseOutputTypeCorrectly` (information, rad 294-295): `ConvertTo-Array` saknar
  `[OutputType([object[]])]`.

## Nya kedjor: Hemköp och ICA (N)

Görs efter prioritet 1, eftersom fler kedjor mångdubblar effekten av D-01 till D-03.

### N-01 Kedjetabell i stället för hårdkodade Willys/Lidl-grenar
- Status: föreslagen | Underlag: kod
- Problem i dagens kod som stoppar nya kedjor:
  - `Export-KvittoMapp`: `($r.chain -eq 'Willys') ? 'willys' : 'lidl'` skickar varje ny kedja till `lidl/`.
  - Huvudflödet har en kopierad gren per kedja; `-Chain`, miljövariabler och Python-`SITES` listar kedjorna var för sig.
- Åtgärd: `$script:Kedjor`, en post per kedja med nyckel (ASCII: `willys`, `lidl`, `hemkop`, `ica`),
  visningsnamn (`Hemköp`), värdnamn, miljövariabel, råmapp, nedladdningsfunktion och
  tolkningsfunktion. Huvudflödet loopar över tabellen (Lidl först). Löser K-06.

### N-02 Hemköp via samma Axfood-kod som Willys
- Status: föreslagen | Underlag: antagande (Hemköp och Willys ägs av Axfood; att API:t är identiskt är inte bekräftat)
- Bekräfta först (användaren, i webbläsarens nätverksflik på https://www.hemkop.se/mina-sidor/kop):
  finns `…/axfood/rest/v1/account/pagedOrderBonusCombined` och `…/digitalreceipt/…`, och
  heter listan `loyaltyTransactionsInPage`?
- Åtgärd om ja:
  - `Save-WillysKvitto` blir `Save-AxfoodKvitto -Kedja <post ur tabellen>`; värdnamn och Referer ur tabellen.
  - `ConvertFrom-WillysKvitto` blir `ConvertFrom-AxfoodKvitto`; orden `Willys Plus`, `Med Willys`
    och sparat-raden blir kedjespecifika mönster i tabellen. Jämför mot en riktig Hemköp-PDF innan mönstren skrivs.
  - Python `SITES["hemkop"]` med start-URL `https://www.hemkop.se/mina-sidor/kop`.
  - `-Chain` får `Hemkop`; miljövariabel `HEMKOP_COOKIE`; råmapp `RawPath/hemkop`.
- Flagga: ny råmapp och ny miljövariabel för cookie. Inget nytt slags data på disk.

### N-03 ICA: API:t måste kartläggas innan design
- Status: föreslagen (blockerad på underlag) | Underlag: antagande
- Läge: inget offentligt dokumenterat API för köphistoriken hittades. Det som finns är
  inofficiella projekt som slutade fungera när ICA ändrade sitt API i april 2024.
- Behövs från användaren (från https://www.ica.se/mitt-konto/kophistorik/, nätverksfliken):
  ett listanrop och ett kvittoanrop med URL, metod, **namnen** på request-headers (inga
  värden) och svarets struktur med personuppgifter borttagna.
- Det avgör tre saker:
  1. Räcker cookies, eller krävs en bearer-token i `Authorization`? I så fall måste
     Python-hjälparen fånga den ur webbläsarens anrop. **Flagga:** ännu en token på stdout
     och i minnet (inte på disk).
  2. Kommer kvittot som JSON med rader (enkel tolkning, ingen pdftotext) eller som PDF?
  3. Hur länge lever sessionen, dvs. var i ordningen ICA ska hämtas.
- Förberedelse som kan göras nu: K-02 (avslutande minus), N-01.

## Separata förslag (F)

Större ändringar. Tas inte som fynd; kräver eget beslut.

### F-01 Fast svensk tidszon överallt
- Status: föreslagen. För: tar K-01 och halva D-05. Mot: befintliga filnamn kan skilja om
  de skapats i annan tidszon (då uppstår en engångsdubblett som D-05 varnar för).

### F-02 Python-hjälparen: egen fil i repot eller `python -I -c`
- Status: föreslagen.
- Egen `.py` bredvid scriptet. För: går att linta, testa och granska i diff. Mot: två filer att distribuera.
- `-c $script:HelperPy`. För: ingen fil skrivs till AppDir. Mot: sämre felspårning, fortfarande inbäddad.

### F-03 Hashlåst installation
- Status: föreslagen. `requirements.txt` med `--require-hashes` för playwright och dess beroenden.
  För: skydd mot utbytt paket. Mot: filen måste genereras om vid varje versionsbyte och
  hashar är plattformsberoende.

### F-04 Inkrementell normalisering
- Status: föreslagen. Tolka bara månader med nya eller ändrade råfiler; `-Force` för allt.
  För: snabbare, färre skrivningar. Mot: en fix i tolkningen slår inte igenom på gamla
  månader utan `-Force`.

### F-05 Pester-tester med fixtures
- Status: föreslagen. För: skyddar belopps- och kvittotolkningen vid N-02/N-03.
  Mot/**Flagga:** fixtures är kvittodata på disk i repot; de måste anonymiseras för hand.

### F-06 Dela upp scriptet i modul + tunt startscript
- Status: föreslagen. För: testbart, kedjorna i egna filer. Mot: ett script blir flera;
  görs lämpligast först efter N-01.

### F-07 Kräv Python 3.10+ för venv:en
- Status: föreslagen. Leta upp nyaste `python3.1x` (Homebrew, python.org) före `python3`,
  och bygg om venv:en när dess tolk är äldre än 3.10. Då kan Playwright-låsningen höjas
  förbi 1.60.0.
  För: webbläsaren som används för BankID kan hållas uppdaterad. Mot: Xcodes python3
  räcker inte längre, så användaren måste installera Python (`brew install python`);
  venv och webbläsare laddas ner på nytt en gång.
