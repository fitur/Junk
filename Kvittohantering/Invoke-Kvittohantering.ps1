<#
.SYNOPSIS
    Hämtar digitala kvitton från Willys och Lidl, normaliserar dem och sorterar dem per år/månad.

.DESCRIPTION
    Steg 1 - Nedladdning (hoppas över med -SkipDownload)
        Frågar efter cookies (dolt) om de inte finns i miljövariablerna WILLYS_COOKIE /
        LIDL_COOKIE. Lämna tomt för att hoppa över en kedja. Lidl hämtas först eftersom
        Lidls token bara lever ca 30 minuter.
        Råkvitton sparas i RawPath/willys (PDF) och RawPath/lidl (JSON + HTML).
        Redan nedladdade kvitton hoppas över.

    Steg 1b - PDF av Lidl-kvitton (hoppas över med -SkipPdf)
        Varje Lidl-kvitto som saknar PDF renderas till en kvittoformad PDF med sökbar text
        (RawPath/lidl/<namn>.pdf), redo för Paperless. Kräver samma engångsinstallation av
        Playwright + Chromium som -AutoCookie och görs även för tidigare nedladdade kvitton.

    Steg 2 - Normalisering
        Tolkar ALLA råkvitton i RawPath och skriver per månad till
        DestinationPath/<år>/<MM-Månad>/ (t.ex. 2026/05-Maj):
            kvitton.json   - gemensamt schema (numeriskt)
            varor.csv      - en rad per vara (svenskt talformat, UTF-8 med BOM)
            <original>     - kopior av råkvittona
        I DestinationPath-roten: manad.json + manad.csv (översikt per månad och kedja).
        Varje kvitto kontrolleras: radsummor + rabatter ska matcha totalen.

    Returnerar ett objekt per skriven månadsmapp.

.EXAMPLE
    ./Invoke-Kvittohantering.ps1
    Frågar efter cookies, hämtar de senaste tre månaderna och normaliserar allt.

.EXAMPLE
    ./Invoke-Kvittohantering.ps1 -AutoCookie
    Öppnar ett webbläsarfönster där du loggar in (BankID för Willys). Inloggningen sparas i
    en egen profil, så nästa körning kräver ofta ingen eller bara en snabb inloggning.
    Första körningen installerar Playwright + Chromium i ~/Library/Application Support/Kvittohantering.

.EXAMPLE
    ./Invoke-Kvittohantering.ps1 -AutoCookie -Chain Lidl
    Hämtar bara Lidl.

.EXAMPLE
    ./Invoke-Kvittohantering.ps1 -FromDate 2024-01-01 -Verbose
    Hämtar all historik sedan 2024 och visar varje kvitto.

.EXAMPLE
    ./Invoke-Kvittohantering.ps1 -SkipDownload -WhatIf
    Visar vilka månadsmappar som skulle skrivas, utan nedladdning och utan att skriva.

.NOTES
    Version:   1.3.0
    Changelog:
        1.3.0 - S-02: cookies i miljövariabler ärvs inte av barnprocesser. S-03: AppDir med rättigheter 700.
                L-01: Playwright låst till 1.65.1, version i markörfilen. U-01: Lidl-HTML renderas utan
                script och nätverk.
        1.2.0 - K-01: alla datum/tider i svensk tid (Europe/Stockholm). K-02: vänstergräns i beloppsregexen och
                U+2212 som minus. K-03: StrictMode-säker felhantering i Willys. K-04: omförsök vid nätverksfel,
                tak 60 s på väntan, fel per kedja stoppar inte normaliseringen. K-05: Lidl-paginering utan
                totalCount. K-06: frågar per kedja som saknar cookie.
        1.1.0 - D-02: nedladdning via .part + Move, ingen halv fil under slutnamnet.
                D-03: samma datum/butik/belopp får suffix _2, _3 (filnamnsschemat utökas bakåtkompatibelt).
                S-01: cookiens värde skrivs inte ut i varning.
        1.0.1 - D-01: avbryter om pdftotext saknas och spärrar skrivning av månader (samt manad.*)
                med otolkbara råkvitton, så att tidigare bra filer inte skrivs över.
        1.0.0 - Utgångspunkt: befintligt script, ingen funktionell ändring.

    Kräver PowerShell 7.4+ (??, ternary, Read-Host -MaskInput, ConvertTo-Csv -UseQuotes).
    Willys-PDF:er kräver pdftotext från poppler: brew install poppler
    Cookies innehåller inloggningstoken - de läses dolt och sparas aldrig på disk av skriptet.
    Med -AutoCookie sparar webbläsarprofilen i AppDir inloggningen, precis som en vanlig webbläsare.
    Avinstallera -AutoCookie genom att radera ~/Library/Application Support/Kvittohantering.
    Mappen får rättigheterna 700. Profilen innehåller inloggningscookies som i praktiken inte är
    krypterade (Playwright använder mock-nyckelring), och mappen följer med i Time Machine-backuper.
    Playwright är låst till en version ($script:PlaywrightVersion); en ändring ominstallerar webbläsaren.
    WILLYS_COOKIE / LIDL_COOKIE tas bort ur processmiljön under körningen och återställs efteråt,
    så att barnprocesser (pip, Chromium, pdftotext) inte ärver dem.
#>

#Requires -Version 7.4

[CmdletBinding(SupportsShouldProcess)]
param(
    # Äldsta kvittodatum att hämta (standard: första dagen tre månader bakåt)
    [datetime]$FromDate = (Get-Date -Day 1).Date.AddMonths(-3),

    # Senaste kvittodatum att hämta
    [datetime]$ToDate = (Get-Date).Date,

    # Mapp för råkvitton (undermappar willys/ och lidl/)
    [ValidateNotNullOrEmpty()]
    [string]$RawPath = (Join-Path -Path $HOME -ChildPath 'Documents/Skript/Kvittonedladdning/Kvitton'),

    # Målmapp; normaliserade kvitton skrivs till <år>/<MM-Månad>/
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath = (Join-Path -Path $HOME -ChildPath 'Documents/Kvitton/Livsmedel'),

    # Sökväg till pdftotext; letas upp automatiskt om den utelämnas
    [string]$PdfToTextPath,

    # Hoppa över nedladdning och normalisera bara det som redan finns i RawPath
    [switch]$SkipDownload,

    # Hämta inloggningen via ett webbläsarfönster (Playwright) i stället för att klistra in cookies
    [switch]$AutoCookie,

    # Skapa inte PDF:er av Lidl-kvitton (görs annars för alla Lidl-kvitton som saknar PDF)
    [switch]$SkipPdf,

    # Kedjor att hämta
    [ValidateSet('Willys', 'Lidl')]
    [string[]]$Chain = @('Willys', 'Lidl'),

    # Webbläsarmotor för -AutoCookie (webkit = Safaris motor, reserv om chromium blockeras)
    [ValidateSet('chromium', 'webkit')]
    [string]$Browser = 'chromium'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --------------------------------------------------------------------------
# Konstanter
# --------------------------------------------------------------------------
$script:Inv = [System.Globalization.CultureInfo]::InvariantCulture

# Beloppsmönster: antingen grupperat (1 118,34 / 1.001,38) eller ogrupperat upp till
# 5 siffror (1118,34 - Willys skriver totaler så). Inga blanksteg inne i ogrupperade
# tal, så artikelnummer (0493686  19,90) kan inte klistras ihop med ett pris.
# Vänstergräns: talet får inte börja mitt i ett längre tal (123456,78 ska inte ge 23456,78).
$script:AmountPos = '(?<![\d,.])(?:\d{1,3}(?:[ .\u00A0\u202F]\d{3})+|\d{1,5}),\d{2}'
# Minustecken framför talet: ASCII-minus eller U+2212
$script:Minus     = '[-\u2212]'
$script:Amount    = $script:Minus + '?' + $script:AmountPos

# Alla datum och tider hanteras i svensk tid, oavsett datorns tidszon (filnamn och normalisering ska matcha)
$script:SvenskTid = [System.TimeZoneInfo]::FindSystemTimeZoneById('Europe/Stockholm')

$script:Manader = @('', 'Januari', 'Februari', 'Mars', 'April', 'Maj', 'Juni',
    'Juli', 'Augusti', 'September', 'Oktober', 'November', 'December')

$script:Utf8NoBom = [System.Text.UTF8Encoding]::new($false)
$script:Utf8Bom   = [System.Text.UTF8Encoding]::new($true)

$script:UserAgent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15'

# Allt för -AutoCookie (venv, webbläsare, profil) samlas här - utanför Dokument,
# som ofta synkas till iCloud. Radera mappen för att avinstallera.
$script:AppDir = Join-Path -Path $HOME -ChildPath 'Library/Application Support/Kvittohantering'

# Låst Playwright-version: paketet och webbläsarbygget hör ihop. Höj här medvetet; markörfilen
# innehåller versionen, så en ändring ger en ominstallation vid nästa körning.
$script:PlaywrightVersion = '1.65.1'

# AppDir innehåller webbläsarprofilen (cookies, i praktiken okrypterade): bara ägaren (700)
$script:AppDirMode = [System.IO.UnixFileMode]'UserRead, UserWrite, UserExecute'

# Python-hjälpare (Playwright) med två kommandon:
#   cookies - inloggning via webbläsarfönster (-AutoCookie)
#   pdf     - renderar Lidl-kvitton (HTML) till PDF
# Skrivs till AppDir vid varje körning, så att den alltid matchar skriptets version.
# stdout = enbart JSON, stderr = meddelanden till användaren.
$script:HelperPy = @'
import argparse, html, json, math, re, sys, time
from datetime import date, timedelta
from playwright.sync_api import sync_playwright, Error as PlaywrightError

def log(msg):
    print(msg, file=sys.stderr, flush=True)

def willys_probe():
    today = date.today()
    return ("https://www.willys.se/axfood/rest/v1/account/pagedOrderBonusCombined"
            f"?fromDate={(today - timedelta(days=30)).isoformat()}&toDate={today.isoformat()}"
            "&currentPage=0&pageSize=1")

LIDL_HISTORY = ("https://www.lidl.se/mre/purchase-history"
                "?client_id=SwedenRetailClient&country_code=se&language=sv-SE&page=1")

SITES = {
    "willys": {
        "start": "https://www.willys.se/mina-kop",
        "probe": willys_probe,
        "headers": {"Accept": "application/json", "X-Requested-With": "XMLHttpRequest",
                    "Referer": "https://www.willys.se/mina-kop"},
        "expect": "loyaltyTransactionsInPage",
        "cookie_url": "https://www.willys.se/axfood/rest/v1/account/",
        "hint": "Logga in på Willys (BankID) i webbläsarfönstret...",
    },
    "lidl": {
        "start": LIDL_HISTORY,
        "probe": lambda: "https://www.lidl.se/mre/api/v1/tickets?country=SE&page=1",
        "headers": {"Accept": "application/json", "Referer": LIDL_HISTORY},
        "expect": "items",
        "cookie_url": "https://www.lidl.se/mre/api/v1/tickets",
        "hint": "Logga in på Lidl i webbläsarfönstret...",
    },
}

def logged_in(ctx, site):
    # Inloggad = API:t svarar 200 med förväntad JSON (inte en omdirigering till inloggning)
    try:
        r = ctx.request.get(site["probe"](), headers=site["headers"], max_redirects=0, timeout=15000)
        if r.status != 200:
            return False
        data = r.json()
        return isinstance(data, dict) and site["expect"] in data
    except Exception:
        return False

def cmd_cookies(a):
    out = {}
    with sync_playwright() as p:
        kw = {"user_data_dir": a.profile, "headless": False, "locale": "sv-SE", "no_viewport": True}
        if a.browser == "chromium":
            # Dölj automationsflaggan som vissa sajter blockerar på
            kw["args"] = ["--disable-blink-features=AutomationControlled"]
            kw["ignore_default_args"] = ["--enable-automation"]
        ctx = getattr(p, a.browser).launch_persistent_context(**kw)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            for name in a.chains:
                site = SITES[name]
                page.goto(site["start"], wait_until="domcontentloaded")
                deadline = time.monotonic() + a.timeout
                hinted = False
                while not logged_in(ctx, site):
                    if not hinted:
                        log(site["hint"])
                        hinted = True
                    if time.monotonic() > deadline:
                        log(f"Tidsgräns: inte inloggad på {name} inom {a.timeout} s")
                        return 2
                    page.wait_for_timeout(3000)
                cookies = ctx.cookies(site["cookie_url"])
                out[name] = "; ".join(f"{c['name']}={c['value']}" for c in cookies)
                log(f"{name}: inloggad")
        except PlaywrightError as e:
            log(f"Webbläsarfel (stängdes fönstret?): {e}")
            return 3
        finally:
            try:
                ctx.close()
            except PlaywrightError:
                pass
    print(json.dumps(out))
    return 0

# Kvittoformad sida: smal, höjd efter innehållet, monospace som på pappersrullen
RECEIPT_CSS = """
@page { margin: 0 }
html, body { margin: 0; padding: 0; background: #fff }
.kvitto { display: inline-block; padding: 8mm 6mm; color: #000;
          font-family: Menlo, Monaco, 'DejaVu Sans Mono', 'Courier New', monospace; font-size: 9pt }
.kvitto h1 { margin: 0 0 4mm 0; font-size: 13pt; text-align: center; letter-spacing: 3px }
.kvitto pre { margin: 0; font-family: inherit; font-size: inherit; line-height: 1.3; white-space: pre }
"""

def render_receipt(page, job):
    src = open(job["html"], encoding="utf-8").read()
    m = re.search(r"<pre>(.*?)</pre>", src, re.S | re.I)
    body = m.group(1) if m else html.escape(src)
    doc = ("<!doctype html><html><head><meta charset='utf-8'>"
           f"<title>{html.escape(job['title'])}</title><style>{RECEIPT_CSS}</style></head>"
           f"<body><div class='kvitto'><h1>LIDL</h1><pre>{body}</pre></div></body></html>")
    page.set_content(doc, wait_until="load")
    box = page.locator(".kvitto").bounding_box()
    # Sidan får exakt innehållets storlek (+ marginal mot avrundning) = en enda sida
    page.pdf(path=job["pdf"], width=f"{math.ceil(box['width']) + 4}px",
             height=f"{math.ceil(box['height']) + 4}px", print_background=True)

def cmd_pdf(a):
    jobs = json.load(open(a.jobs, encoding="utf-8"))
    ok = failed = 0
    with sync_playwright() as p:
        # channel="chromium" = fullversionen i headless-läge (headless shell installeras inte)
        browser = p.chromium.launch(channel="chromium", headless=True)
        try:
            # Kvitto-HTML är opålitlig data: inga script och inget nätverk vid rendering
            ctx = browser.new_context(java_script_enabled=False, offline=True)
            ctx.route("**/*", lambda route: route.abort())
            page = ctx.new_page()
            for job in jobs:
                try:
                    render_receipt(page, job)
                    ok += 1
                except Exception as e:
                    failed += 1
                    log(f"PDF misslyckades för {job['html']}: {e}")
        finally:
            browser.close()
    print(json.dumps({"ok": ok, "failed": failed}))
    return 0 if ok or not jobs else 4

def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("cookies")
    c.add_argument("--profile", required=True)
    c.add_argument("--chains", nargs="+", required=True, choices=list(SITES))
    c.add_argument("--browser", default="chromium", choices=["chromium", "webkit"])
    c.add_argument("--timeout", type=int, default=300)
    d = sub.add_parser("pdf")
    d.add_argument("--jobs", required=True, help="JSON-fil med [{html, pdf, title}]")
    a = ap.parse_args()
    return cmd_cookies(a) if a.cmd == "cookies" else cmd_pdf(a)

if __name__ == "__main__":
    sys.exit(main())
'@

# ==========================================================================
# Allmänna hjälpfunktioner
# ==========================================================================
function Get-PropertyValue {
    # StrictMode-säker egenskapsläsning
    [CmdletBinding()]
    param([AllowNull()]$InputObject, [Parameter(Mandatory)][string]$Name)

    if ($null -eq $InputObject) { return $null }
    $prop = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $null }
    return $prop.Value
}

function ConvertTo-Array {
    # Gör om null/enstaka objekt till en riktig array (undviker fantomrader från @($null))
    [CmdletBinding()]
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) { return , @() }
    return , @($InputObject)
}

function ConvertFrom-SvensktTal {
    # Svenskt talformat -> double. Komma = decimal; punkt/blanksteg = tusental.
    # '1 234,56' -> 1234.56 ; '1.001,38' -> 1001.38 ; tomt -> $null
    [CmdletBinding()]
    [OutputType([double])]
    param([AllowNull()][AllowEmptyString()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    # \s i .NET täcker även nbsp (U+00A0) och smalt mellanslag (U+202F)
    $s = ($Text -replace '\u2212', '-') -replace '\s', ''
    if ($s -eq '-') { return $null }
    if ($s.Contains(',')) { $s = $s.Replace('.', '').Replace(',', '.') }
    $value = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $script:Inv, [ref]$value)) {
        return $value
    }
    return $null
}

function Format-Krona {
    # Belopp -> '1604,79'. InvariantCulture + ersätt punkt, eftersom sv-SE i .NET (ICU)
    # använder U+2212 som minustecken, vilket Excel inte tolkar som ett tal.
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    return ([double]$Value).ToString('F2', $script:Inv).Replace('.', ',')
}

function Format-Tal {
    # Allmänt tal (t.ex. antal/vikt) -> '1,152'
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    return ([double]$Value).ToString('0.###', $script:Inv).Replace('.', ',')
}

function ConvertTo-SvenskTid {
    # Tidpunkt -> väggklocktid i Europa/Stockholm
    [CmdletBinding()]
    [OutputType([datetime])]
    param([Parameter(Mandatory)][DateTimeOffset]$Value)

    return [System.TimeZoneInfo]::ConvertTime($Value, $script:SvenskTid).DateTime
}

function Get-SafeName {
    # Filnamnssäker butiksnamn (behåller åäö - samma schema som tidigare nedladdningar)
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Name)

    return (($Name -replace '[^\w\såäöÅÄÖ-]', '') -replace '\s+', '_')
}

# ==========================================================================
# HTTP
# ==========================================================================
function New-CookieSession {
    # Bygger en WebSession från en rå Cookie-rad ("a=1; b=2; ...")
    [CmdletBinding()]
    [OutputType([Microsoft.PowerShell.Commands.WebRequestSession])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Skapar endast ett objekt i minnet.')]
    param(
        [Parameter(Mandatory)][string]$CookieString,
        [Parameter(Mandatory)][string]$Domain
    )

    $session = [Microsoft.PowerShell.Commands.WebRequestSession]::new()
    $jar     = [System.Net.CookieContainer]::new()
    foreach ($part in ($CookieString.Trim() -replace '^Cookie:\s*', '' -split ';\s*')) {
        $idx = $part.IndexOf('=')
        if ($idx -lt 1) { continue }
        $name  = $part.Substring(0, $idx).Trim()
        $value = $part.Substring($idx + 1).Trim()
        try {
            $jar.Add([System.Net.Cookie]::new($name, $value, '/', $Domain))
        }
        catch {
            # Inte Exception.Message: CookieException citerar cookiens värde i klartext
            Write-Warning "Cookie '$name' kunde inte läggas till (ogiltigt tecken i värdet)."
        }
    }
    $session.Cookies = $jar
    return $session
}

function Invoke-KvittoRequest {
    # GET med retry och back-off för 429/5xx (respekterar Retry-After).
    # 401/403 kastar UnauthorizedAccessException = sessionen har gått ut.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Headers,
        [Parameter(Mandatory)][Microsoft.PowerShell.Commands.WebRequestSession]$Session,
        # Om angiven sparas svaret direkt till fil (binärt) i stället för att tolkas
        [string]$OutFile,
        [ValidateRange(0, 10)][int]$MaxRetries = 4
    )

    for ($attempt = 1; ; $attempt++) {
        try {
            $params = @{ Uri = $Uri; Headers = $Headers; WebSession = $Session; ErrorAction = 'Stop' }
            if ($OutFile) {
                Invoke-WebRequest @params -OutFile $OutFile | Out-Null
                return
            }
            return Invoke-RestMethod @params
        }
        catch [Microsoft.PowerShell.Commands.HttpResponseException] {
            $status = [int]$PSItem.Exception.Response.StatusCode
            if ($status -in 401, 403) {
                throw [System.UnauthorizedAccessException]::new("HTTP $status - sessionen har gått ut eller cookien är ogiltig.")
            }
            if (($status -eq 429 -or $status -ge 500) -and $attempt -le $MaxRetries) {
                $delay = [math]::Pow(2, $attempt)
                $retryAfter = $PSItem.Exception.Response.Headers.RetryAfter
                if ($retryAfter) {
                    if ($null -ne $retryAfter.Delta) { $delay = $retryAfter.Delta.TotalSeconds }
                    elseif ($null -ne $retryAfter.Date) { $delay = [math]::Max(1, ($retryAfter.Date - [DateTimeOffset]::UtcNow).TotalSeconds) }
                }
                $delay = [math]::Min(60, $delay)   # tak: servern kan beordra en orimligt lång väntan
                Write-Verbose "HTTP $status - försök $attempt av $MaxRetries, väntar $delay s"
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
        catch [System.Net.Http.HttpRequestException], [System.Threading.Tasks.TaskCanceledException] {
            # Nätverksfel utan HTTP-svar (DNS, anslutning, timeout): försök om med back-off.
            # HttpResponseException ärver från HttpRequestException men fångas av blocket ovan.
            if ($attempt -le $MaxRetries) {
                $delay = [math]::Min(60, [math]::Pow(2, $attempt))
                Write-Verbose "Nätverksfel - försök $attempt av $MaxRetries, väntar $delay s"
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Assert-JsonSvar {
    # Inloggningssidor kommer som HTML med status 200. Invoke-RestMethod ger då en
    # sträng eller (om HTML:en råkar vara giltig XML) ett XmlDocument. Kräv därför
    # att den förväntade JSON-egenskapen finns, annars räknas sessionen som utgången.
    [CmdletBinding()]
    param(
        [AllowNull()]$Response,
        [Parameter(Mandatory)][string]$ExpectedProperty,
        [Parameter(Mandatory)][string]$Chain
    )
    if ($null -eq $Response -or $Response -is [string] -or $Response -is [System.Xml.XmlNode] -or
        $null -eq $Response.PSObject.Properties[$ExpectedProperty]) {
        throw [System.UnauthorizedAccessException]::new("$Chain svarade inte med förväntad JSON - sessionen har troligen gått ut.")
    }
}

function Get-JwtUtgang {
    # Läser exp-claim ur en JWT (utan signaturkontroll - används bara för förvarning)
    [CmdletBinding()]
    [OutputType([datetime])]
    param([Parameter(Mandatory)][string]$Token)

    $payload = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
    $claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
    $exp = Get-PropertyValue $claims 'exp'
    if ($null -eq $exp) { return $null }
    return [DateTimeOffset]::FromUnixTimeSeconds([long]$exp).LocalDateTime
}

# ==========================================================================
# -AutoCookie: inloggning via webbläsarfönster (Playwright)
# ==========================================================================
function Initialize-Kvittowebblasare {
    # Engångsinstallation: venv + Playwright + webbläsare i AppDir. Returnerar sökväg till python.
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([string])]
    param([Parameter(Mandatory)][ValidateSet('chromium', 'webkit')][string]$Browser)

    $python = Join-Path $script:AppDir 'venv/bin/python'
    $marker = Join-Path $script:AppDir "browsers/.installerad-$Browser"
    if ((Test-Path -LiteralPath $python) -and (Test-Path -LiteralPath $marker) -and
        ([System.IO.File]::ReadAllText($marker).Trim() -eq $script:PlaywrightVersion)) { return $python }

    $sysPython = Get-Command -Name 'python3' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $sysPython) { throw 'python3 saknas. Installera Xcode Command Line Tools: xcode-select --install' }

    if (-not $PSCmdlet.ShouldProcess($script:AppDir, "Installera Playwright och $Browser (engångsåtgärd, ca 150-250 MB)")) {
        throw 'Installationen av webbläsaren avbröts.'
    }
    # S-03: bara ägaren når mappen (profilen med inloggning ligger här)
    [void][System.IO.Directory]::CreateDirectory($script:AppDir, $script:AppDirMode)

    if (-not (Test-Path -LiteralPath $python)) {
        Write-Information 'Skapar Python-miljö...' -InformationAction Continue
        & $sysPython.Source -m venv (Join-Path $script:AppDir 'venv') | Out-Host
        if ($LASTEXITCODE -ne 0) { throw "python3 -m venv misslyckades (kod $LASTEXITCODE)" }
    }
    Write-Information 'Installerar Playwright...' -InformationAction Continue
    & $python -m pip install --quiet --disable-pip-version-check "playwright==$($script:PlaywrightVersion)" | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "pip install playwright misslyckades (kod $LASTEXITCODE)" }

    Write-Information "Laddar ner $Browser..." -InformationAction Continue
    $installArgs = @('-m', 'playwright', 'install', $Browser)
    if ($Browser -eq 'chromium') { $installArgs += '--no-shell' }   # hoppa över headless-varianten
    & $python @installArgs | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "playwright install $Browser misslyckades (kod $LASTEXITCODE)" }

    [System.IO.File]::WriteAllText($marker, $script:PlaywrightVersion, $script:Utf8NoBom)
    return $python
}

function Invoke-Kvittohjalpare {
    # Kör Python-hjälparen (installerar vid behov). Returnerar dess stdout som JSON-objekt.
    # stderr (instruktioner, fel) visas direkt i terminalen.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        # Webbläsarmotor som måste vara installerad för kommandot
        [Parameter(Mandatory)][ValidateSet('chromium', 'webkit')][string]$Browser,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    $previous = $env:PLAYWRIGHT_BROWSERS_PATH
    try {
        # Webbläsarna läggs i AppDir i stället för ~/Library/Caches/ms-playwright
        $env:PLAYWRIGHT_BROWSERS_PATH = Join-Path $script:AppDir 'browsers'
        $python = Initialize-Kvittowebblasare -Browser $Browser
        # Gäller även en mapp som skapats tidigare med standardrättigheter
        [System.IO.File]::SetUnixFileMode($script:AppDir, $script:AppDirMode)

        $helper = Join-Path $script:AppDir 'kvittohjalpare.py'
        [System.IO.File]::WriteAllText($helper, $script:HelperPy, $script:Utf8NoBom)
        [System.IO.File]::Delete((Join-Path $script:AppDir 'kvittocookie.py'))   # tidigare namn

        $stdout = & $python $helper @Arguments
        if ($LASTEXITCODE -ne 0) { throw "Webbläsarhjälparen ($($Arguments[0])) misslyckades med kod $LASTEXITCODE." }
    }
    finally {
        $env:PLAYWRIGHT_BROWSERS_PATH = $previous
    }
    return (($stdout -join '') | ConvertFrom-Json)
}

function Get-KvittoCookieAutomatisk {
    # Öppnar ett webbläsarfönster med beständig profil, väntar tills du är inloggad
    # och returnerar cookie-raderna som { willys = '...'; lidl = '...' }.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string[]]$Chain,
        [Parameter(Mandatory)][ValidateSet('chromium', 'webkit')][string]$Browser
    )

    # Willys först (BankID tar längst), sedan Lidl så att Lidl-tokenen är så färsk som möjligt
    $order = @('willys', 'lidl') | Where-Object { $_ -in $Chain }
    Write-Information 'Öppnar webbläsarfönster - logga in om du blir ombedd.' -InformationAction Continue
    $helperArgs = @('cookies', '--profile', (Join-Path $script:AppDir 'profil'), '--browser', $Browser, '--chains') + $order
    return Invoke-Kvittohjalpare -Browser $Browser -Arguments $helperArgs
}

function ConvertTo-LidlPdf {
    # Skapar <namn>.pdf bredvid varje Lidl-kvitto (<namn>.json) som saknar PDF.
    # Returnerar antal skapade PDF:er.
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([int])]
    param([Parameter(Mandatory)][string]$LidlDir)

    $pending = @(Get-ChildItem -LiteralPath $LidlDir -Filter '*.json' -File |
            Where-Object { $_.Name -ne 'kvittolista.json' -and
                -not (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($_.FullName, '.pdf'))) } |
            Sort-Object -Property Name)
    if ($pending.Count -eq 0) { return 0 }
    if (-not $PSCmdlet.ShouldProcess($LidlDir, "Skapa $($pending.Count) PDF:er av Lidl-kvitton")) { return 0 }

    $jobs = [System.Collections.Generic.List[object]]::new()
    foreach ($js in $pending) {
        $htmlPath = [System.IO.Path]::ChangeExtension($js.FullName, '.html')
        if (-not (Test-Path -LiteralPath $htmlPath)) {
            # HTML saknas (äldre nedladdning) - återskapa från JSON
            $receipt = Get-PropertyValue (Get-Content -LiteralPath $js.FullName -Raw -Encoding utf8 | ConvertFrom-Json) 'ticket'
            $markup  = Get-PropertyValue $receipt 'htmlPrintedReceipt'
            if (-not $markup) { Write-Warning "Lidl $($js.Name): saknar kvittotext, ingen PDF skapad"; continue }
            [System.IO.File]::WriteAllText($htmlPath, [string]$markup, $script:Utf8NoBom)
        }
        $jobs.Add([pscustomobject]@{
                html  = $htmlPath
                pdf   = [System.IO.Path]::ChangeExtension($js.FullName, '.pdf')
                title = 'Lidl ' + $js.BaseName.Replace('_', ' ')   # PDF-titel, syns i Paperless
            })
    }
    if ($jobs.Count -eq 0) { return 0 }

    Write-Information "Skapar $($jobs.Count) Lidl-PDF:er..." -InformationAction Continue
    $jobFile = [System.IO.Path]::GetTempFileName()
    try {
        [System.IO.File]::WriteAllText($jobFile, (ConvertTo-Json -InputObject @($jobs) -Depth 20), $script:Utf8NoBom)
        $result = Invoke-Kvittohjalpare -Browser chromium -Arguments @('pdf', '--jobs', $jobFile)
    }
    finally {
        [System.IO.File]::Delete($jobFile)
    }
    if ($result.failed -gt 0) { Write-Warning "$($result.failed) Lidl-PDF:er kunde inte skapas (se meddelanden ovan)." }
    return [int]$result.ok
}

# ==========================================================================
# Nedladdning - Willys
# ==========================================================================
function Save-WillysKvitto {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CookieString,
        [Parameter(Mandatory)][string]$OutputDir,
        [Parameter(Mandatory)][datetime]$From,
        [Parameter(Mandatory)][datetime]$To
    )

    $session = New-CookieSession -CookieString $CookieString -Domain 'www.willys.se'
    $headers = @{
        'User-Agent'       = $script:UserAgent
        'Accept'           = 'application/json, text/plain, */*'
        'Accept-Language'  = 'sv-SE,sv;q=0.9'
        'Referer'          = 'https://www.willys.se/mina-kop'
        'X-Requested-With' = 'XMLHttpRequest'
    }
    $fromStr = $From.ToString('yyyy-MM-dd', $script:Inv)
    $toStr   = $To.ToString('yyyy-MM-dd', $script:Inv)
    $pageSize = 10

    # Steg 1: transaktionslista (paginerad)
    $all = [System.Collections.Generic.List[object]]::new()
    for ($page = 0; $page -lt 500; $page++) {
        $uri  = "https://www.willys.se/axfood/rest/v1/account/pagedOrderBonusCombined?fromDate=$fromStr&toDate=$toStr&currentPage=$page&pageSize=$pageSize"
        $resp = Invoke-KvittoRequest -Uri $uri -Headers $headers -Session $session
        Assert-JsonSvar -Response $resp -ExpectedProperty 'loyaltyTransactionsInPage' -Chain 'Willys'
        $tx = ConvertTo-Array (Get-PropertyValue $resp 'loyaltyTransactionsInPage')
        foreach ($t in $tx) { $all.Add($t) }
        Write-Verbose "Willys sida $page : $($tx.Count) transaktioner"
        if ($tx.Count -lt $pageSize) { break }
    }

    # Sorterad så att suffixen (_2, _3) tilldelas lika vid varje körning, oavsett API:ts ordning
    $withReceipt = @($all | Where-Object {
            (Get-PropertyValue $_ 'digitalReceiptAvailable') -and (Get-PropertyValue $_ 'digitalReceiptReference')
        } | Sort-Object -Property { [long](Get-PropertyValue $_ 'bookingDate') }, { [string](Get-PropertyValue $_ 'digitalReceiptReference') })
    $result = [pscustomobject]@{ Kedja = 'Willys'; Hittade = $withReceipt.Count; Nya = 0; Fanns = 0; Misslyckade = 0 }
    if ($withReceipt.Count -eq 0) { return $result }
    if (-not $PSCmdlet.ShouldProcess($OutputDir, "Ladda ner upp till $($withReceipt.Count) Willys-kvitton")) { return $result }

    [void][System.IO.Directory]::CreateDirectory($OutputDir)
    $nameCount = @{}   # samma datum+butik+belopp flera gånger i listan -> _2, _3 ...
    $i = 0
    foreach ($t in $withReceipt) {
        $i++
        Write-Progress -Activity 'Willys' -Status "Kvitto $i av $($withReceipt.Count)" -PercentComplete ($i * 100 / $withReceipt.Count)

        # Filnamnsschemat är detsamma som tidigare nedladdningar, så befintliga filer känns igen
        $date     = (ConvertTo-SvenskTid ([DateTimeOffset]::FromUnixTimeMilliseconds([long](Get-PropertyValue $t 'bookingDate')))).ToString('yyyy-MM-dd', $script:Inv)
        $store    = Get-SafeName (Get-PropertyValue $t 'storeName')
        $amount   = "$(Get-PropertyValue $t 'amount')" -replace '[,\s]', '_'
        $baseName = "${date}_${store}_${amount}kr"
        $nameCount[$baseName] = 1 + [int]$nameCount[$baseName]
        if ($nameCount[$baseName] -gt 1) { $baseName += "_$($nameCount[$baseName])" }
        $fileName = Join-Path $OutputDir "$baseName.pdf"

        if (Test-Path -LiteralPath $fileName) { $result.Fanns++; continue }

        $source  = (Get-PropertyValue $t 'receiptSource') ?? 'aws'
        $refEnc  = [Uri]::EscapeDataString((Get-PropertyValue $t 'digitalReceiptReference'))
        $uri     = "https://www.willys.se/axfood/rest/order/orders/digitalreceipt/$refEnc" +
                   "?date=$date&storeId=$(Get-PropertyValue $t 'storeCustomerId')&source=$source&memberCardNumber=$(Get-PropertyValue $t 'memberCardNumber')"

        try {
            # Ladda ner till .part och byt namn först när filen är verifierad. Ett avbrott
            # lämnar då aldrig en halv PDF under slutnamnet, som nästa körning skulle hoppa över.
            $partFile = "$fileName.part"
            Invoke-KvittoRequest -Uri $uri -Headers $headers -Session $session -OutFile $partFile
            $head = [byte[]]::new(4)
            $fs   = [System.IO.File]::OpenRead($partFile)
            try { $read = $fs.Read($head, 0, 4) } finally { $fs.Dispose() }
            if ($read -lt 4 -or [System.Text.Encoding]::ASCII.GetString($head) -ne '%PDF') {
                [System.IO.File]::Delete($partFile)
                throw [System.UnauthorizedAccessException]::new('Willys returnerade inte en PDF - sessionen har troligen gått ut.')
            }
            [System.IO.File]::Move($partFile, $fileName)   # utan överskrivning: kastar om slutnamnet finns
            $result.Nya++
            Write-Verbose "Willys: $(Split-Path $fileName -Leaf)"
        }
        catch [System.UnauthorizedAccessException] {
            Write-Progress -Activity 'Willys' -Completed
            throw
        }
        catch {
            $result.Misslyckade++
            Write-Warning "Willys $date $(Get-PropertyValue $t 'storeName'): $($PSItem.Exception.Message)"
        }
        Start-Sleep -Milliseconds 300
    }
    Write-Progress -Activity 'Willys' -Completed
    return $result
}

# ==========================================================================
# Nedladdning - Lidl
# ==========================================================================
function Save-LidlKvitto {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$CookieString,
        [Parameter(Mandatory)][string]$OutputDir,
        [Parameter(Mandatory)][datetime]$From,
        [Parameter(Mandatory)][datetime]$To
    )

    # Förvarna om token redan gått ut eller snart går ut
    $mToken = [regex]::Match($CookieString, '(?:^|;\s*)authToken=([^;]+)')
    if (-not $mToken.Success) { throw 'Lidl-cookien saknar authToken - kopiera hela Cookie-raden från ett mre/api-anrop.' }
    try {
        $exp = Get-JwtUtgang -Token $mToken.Groups[1].Value
        if ($exp) {
            $left = ($exp - (Get-Date)).TotalMinutes
            if ($left -le 0) { throw [System.UnauthorizedAccessException]::new("Lidl-token gick ut $($exp.ToString('HH:mm')) - hämta en ny cookie.") }
            Write-Verbose ("Lidl-token giltig i {0:N0} min till" -f $left)
            if ($left -lt 3) { Write-Warning ("Lidl-token går ut om {0:N1} min - nedladdningen kan avbrytas." -f $left) }
        }
    }
    catch [System.UnauthorizedAccessException] { throw }
    catch { Write-Verbose "Kunde inte läsa token-utgång: $($PSItem.Exception.Message)" }

    $session = New-CookieSession -CookieString $CookieString -Domain 'www.lidl.se'
    $headers = @{
        'User-Agent'      = $script:UserAgent
        'Accept'          = 'application/json, text/plain, */*'
        'Accept-Language' = 'sv-SE,sv;q=0.9'
        'Referer'         = 'https://www.lidl.se/mre/purchase-history?client_id=SwedenRetailClient&country_code=se&language=sv-SE&page=1'
    }

    # Steg 1: kvittolista. API:t saknar datumfilter men listan är sorterad nyast först,
    # så pagineringen avbryts när en sida bara innehåller kvitton äldre än From.
    $inRange = [System.Collections.Generic.List[object]]::new()
    $seen    = 0
    for ($page = 1; $page -lt 500; $page++) {
        $resp  = Invoke-KvittoRequest -Uri "https://www.lidl.se/mre/api/v1/tickets?country=SE&page=$page" -Headers $headers -Session $session
        Assert-JsonSvar -Response $resp -ExpectedProperty 'items' -Chain 'Lidl'
        $items = ConvertTo-Array (Get-PropertyValue $resp 'items')
        $total = [int]((Get-PropertyValue $resp 'totalCount') ?? 0)
        if ($items.Count -eq 0) { break }
        $seen += $items.Count

        $allOlder = $true
        foreach ($tk in $items) {
            $day = (ConvertTo-SvenskTid ([DateTimeOffset](Get-PropertyValue $tk 'date'))).Date
            if ($day -ge $From.Date) { $allOlder = $false }
            if ($day -ge $From.Date -and $day -le $To.Date) { $inRange.Add($tk) }
        }
        Write-Verbose "Lidl sida $page : $($items.Count) kvitton ($seen av $total)"
        if ($allOlder -or ($total -gt 0 -and $seen -ge $total)) { break }
    }

    # Sorterad så att suffixen (_2, _3) tilldelas lika vid varje körning, oavsett API:ts ordning
    $inRange = [System.Collections.Generic.List[object]]::new(
        [object[]]@($inRange | Sort-Object -Property { ([DateTimeOffset](Get-PropertyValue $_ 'date')).UtcTicks }, { [string](Get-PropertyValue $_ 'id') }))
    $result = [pscustomobject]@{ Kedja = 'Lidl'; Hittade = $inRange.Count; Nya = 0; Fanns = 0; Misslyckade = 0 }
    if ($inRange.Count -eq 0) { return $result }
    if (-not $PSCmdlet.ShouldProcess($OutputDir, "Ladda ner upp till $($inRange.Count) Lidl-kvitton")) { return $result }

    [void][System.IO.Directory]::CreateDirectory($OutputDir)
    $nameCount = @{}   # samma datum+butik+belopp flera gånger i listan -> _2, _3 ...
    $i = 0
    foreach ($tk in $inRange) {
        $i++
        Write-Progress -Activity 'Lidl' -Status "Kvitto $i av $($inRange.Count)" -PercentComplete ($i * 100 / $inRange.Count)

        # Filnamnsschemat är detsamma som tidigare nedladdningar, så befintliga filer känns igen
        $date     = (ConvertTo-SvenskTid ([DateTimeOffset](Get-PropertyValue $tk 'date'))).ToString('yyyy-MM-dd', $script:Inv)
        $store    = Get-SafeName (Get-PropertyValue $tk 'store')
        $amount   = "$(Get-PropertyValue $tk 'totalAmount')" -replace '[,\s]', '_'
        $baseName = "${date}_${store}_${amount}kr"
        $nameCount[$baseName] = 1 + [int]$nameCount[$baseName]
        if ($nameCount[$baseName] -gt 1) { $baseName += "_$($nameCount[$baseName])" }
        $jsonFile = Join-Path $OutputDir "$baseName.json"
        $htmlFile = Join-Path $OutputDir "$baseName.html"

        if (Test-Path -LiteralPath $jsonFile) { $result.Fanns++; continue }

        try {
            $id     = [Uri]::EscapeDataString([string](Get-PropertyValue $tk 'id'))
            $detail = Invoke-KvittoRequest -Uri "https://www.lidl.se/mre/api/v1/tickets/${id}?country=SE&languageCode=sv-SE" -Headers $headers -Session $session
            Assert-JsonSvar -Response $detail -ExpectedProperty 'ticket' -Chain 'Lidl'
            # HTML först (bara om den inte finns), sedan JSON sist: JSON-filen är markören för "klart".
            # Båda skrivs till .part och byts namn utan överskrivning, så ett avbrott lämnar ingen halv slutfil.
            $html = Get-PropertyValue (Get-PropertyValue $detail 'ticket') 'htmlPrintedReceipt'
            if ($html -and -not (Test-Path -LiteralPath $htmlFile)) {
                [System.IO.File]::WriteAllText("$htmlFile.part", [string]$html, $script:Utf8NoBom)
                [System.IO.File]::Move("$htmlFile.part", $htmlFile)
            }
            [System.IO.File]::WriteAllText("$jsonFile.part", (ConvertTo-Json -InputObject $detail -Depth 20), $script:Utf8NoBom)
            [System.IO.File]::Move("$jsonFile.part", $jsonFile)
            $result.Nya++
            Write-Verbose "Lidl: $baseName"
        }
        catch [System.UnauthorizedAccessException] {
            Write-Progress -Activity 'Lidl' -Completed
            throw
        }
        catch {
            $result.Misslyckade++
            Write-Warning "Lidl $date $(Get-PropertyValue $tk 'store'): $($PSItem.Exception.Message)"
        }
        Start-Sleep -Milliseconds 200
    }
    Write-Progress -Activity 'Lidl' -Completed
    return $result
}

# ==========================================================================
# Normalisering - gemensamma byggstenar
# ==========================================================================
function New-KvittoRad {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '',
        Justification = 'Skapar endast ett objekt i minnet.')]
    param(
        [AllowNull()][string]$Name,
        [AllowNull()]$Quantity,
        [AllowNull()]$UnitPrice,
        [AllowNull()]$LineTotal,
        [switch]$Deposit
    )
    return [pscustomobject]@{
        name       = $Name
        quantity   = $Quantity
        unit_price = $UnitPrice
        line_total = $LineTotal
        discount   = $null
        deposit    = [bool]$Deposit
    }
}

function Find-PdfToText {
    # VS Code startat från Dock kan sakna Homebrew i PATH, därför kontrolleras
    # även Homebrews standardsökvägar.
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyString()][string]$Hint)

    if ($Hint) {
        if (Test-Path -LiteralPath $Hint -PathType Leaf) { return $Hint }
        throw "pdftotext hittades inte på angiven sökväg: $Hint"
    }
    $cmd = Get-Command -Name 'pdftotext' -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($cmd) { return $cmd.Source }
    foreach ($candidate in '/opt/homebrew/bin/pdftotext', '/usr/local/bin/pdftotext') {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    }
    return $null
}

# ==========================================================================
# Normalisering - Willys (PDF via pdftotext -layout)
# ==========================================================================
function ConvertFrom-WillysKvitto {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)][string]$PdfToText
    )

    # Skriv till temporär fil med explicit UTF-8, så konsolens kodning inte spelar roll
    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        & $PdfToText -enc UTF-8 -layout $File.FullName $tmp 2>$null
        if ($LASTEXITCODE -ne 0) { throw "pdftotext avslutades med kod $LASTEXITCODE" }
        $text = [System.IO.File]::ReadAllText($tmp, $script:Utf8NoBom)
    }
    finally {
        [System.IO.File]::Delete($tmp)
    }
    # Dela även på sidbrytning (\f), som pdftotext sätter mellan sidor
    $lines = $text -split '\r\n|\n|\r|\f'

    $A  = $script:Amount
    $AP = $script:AmountPos
    # Varulistans rader skiljs åt av indraget, inte av ordval:
    #   0-1 blanksteg            = vara (vissa varor har ett inledande blanksteg)
    #   indrag + negativt belopp = rabatt (Rabatt, Willys Plus, Prisnedsättning, Nytt pris ...)
    #   indrag + "+"             = pant
    $reArt      = [regex]::new('^ ?(?<name>\S.*?)\s{2,}(?:(?<qty>\d+)\s*st?\s*\*\s*(?<unit>[\d,]+)\s+)?(?<total>' + $A + ')\s*$')
    $reDisc     = [regex]::new('^\s+\S.*?\s(?<amt>' + $script:Minus + $AP + ')\s*$')
    $reDeposit  = [regex]::new('^\s+\+.*?(?<amt>' + $A + ')\s*$')
    $reWeight   = [regex]::new('^\s+(?<kg>[\d,]+)kg\s*\*\s*(?<unit>[\d,]+)kr/kg\s+(?<total>' + $A + ')')
    $reNameOnly = [regex]::new('^ ?(?<name>\p{Lu}.*?)\s*$')
    $reTotal    = [regex]::new('^\s*Totalt\s+(?<t>' + $A + ')\s*SEK')
    $reSavings  = [regex]::new('har du sparat:\s+(?<s>' + $A + ')')
    $reDate     = [regex]::new('^(?<d>\d{4}-\d{2}-\d{2})\s+(?<ti>\d{2}:\d{2}):\d{2}')
    $reCard     = [regex]::new('(\*{4,}\d{3,4})')

    # Butik = första icke-tomma raden
    $store = $null
    foreach ($ln in ($lines | Select-Object -First 5)) {
        if ($ln.Trim()) { $store = $ln.Trim(); break }
    }

    $items   = [System.Collections.Generic.List[object]]::new()
    $total   = $null
    $savings = $null
    $date    = $null
    $time    = $null
    $payment = $null
    $inItems = $false

    foreach ($ln in $lines) {
        $trimmed = $ln.Trim()

        # Streckad avdelare startar varulistan
        if ($trimmed.Length -gt 10 -and $trimmed -match '^-+$') { $inItems = $true; continue }

        $m = $reTotal.Match($ln)
        if ($m.Success) { $total = ConvertFrom-SvensktTal $m.Groups['t'].Value; $inItems = $false; continue }
        # Skyddsnät: varulistan slutar vid första "Totalt"-raden även om beloppet inte gick
        # att tolka - annars räknas moms- och betalrader som varor
        if ($trimmed.StartsWith('Totalt', [StringComparison]::Ordinal)) { $inItems = $false; continue }
        $m = $reSavings.Match($ln)
        if ($m.Success) { $savings = ConvertFrom-SvensktTal $m.Groups['s'].Value; continue }
        $m = $reDate.Match($ln)
        if ($m.Success) { $date = $m.Groups['d'].Value; $time = $m.Groups['ti'].Value; continue }
        if ($ln.StartsWith('Visa', [StringComparison]::Ordinal) -or
            $ln.StartsWith('Mastercard', [StringComparison]::Ordinal) -or
            $ln.Contains('Kontokort')) {
            $mc = $reCard.Match($ln)
            if ($mc.Success) { $payment = 'Kort ' + $mc.Groups[1].Value }
        }

        if (-not $inItems) { continue }
        $last = $items.Count -gt 0 ? $items[$items.Count - 1] : $null

        # Pant -> egen post
        $m = $reDeposit.Match($ln)
        if ($m.Success -and $last) {
            $amt = ConvertFrom-SvensktTal $m.Groups['amt'].Value
            $items.Add((New-KvittoRad -Name 'Pant' -Quantity 1 -UnitPrice $amt -LineTotal $amt -Deposit))
            continue
        }

        # Rabatt -> knyts till föregående vara
        $m = $reDisc.Match($ln)
        if ($m.Success -and $last) {
            $last.discount = ($last.discount ?? 0) + (ConvertFrom-SvensktTal $m.Groups['amt'].Value)
            continue
        }

        # Viktrad -> fyller föregående namnrad som saknar belopp
        $m = $reWeight.Match($ln)
        if ($m.Success -and $last -and $null -eq $last.line_total) {
            $last.quantity   = ConvertFrom-SvensktTal $m.Groups['kg'].Value
            $last.unit_price = ConvertFrom-SvensktTal $m.Groups['unit'].Value
            $last.line_total = ConvertFrom-SvensktTal $m.Groups['total'].Value
            continue
        }

        # Vanlig artikelrad
        $m = $reArt.Match($ln)
        if ($m.Success -and -not ($trimmed.StartsWith('Moms') -or $trimmed.StartsWith('Totalt') -or $trimmed.StartsWith('Med Willys'))) {
            $name = $m.Groups['name'].Value.Trim()
            if (-not $name -or $name.StartsWith('Rabatt') -or $name.StartsWith('Willys Plus') -or $name.StartsWith('Prisned')) {
                continue
            }
            $qty  = (ConvertFrom-SvensktTal $m.Groups['qty'].Value) ?? 1
            $unit = ConvertFrom-SvensktTal $m.Groups['unit'].Value
            $lt   = ConvertFrom-SvensktTal $m.Groups['total'].Value
            if ($null -eq $unit -and $qty -ne 1 -and $null -ne $lt) {
                $unit = [math]::Round($lt / $qty, 2)
            }
            $items.Add((New-KvittoRad -Name $name -Quantity $qty -UnitPrice ($unit ? $unit : $lt) -LineTotal $lt))
            continue
        }

        # Namnrad utan belopp (viktvara) -> platshållare som fylls av viktraden
        $m = $reNameOnly.Match($ln)
        if ($m.Success -and -not $ln.StartsWith('  ') -and $ln -notmatch '\d,\d{2}' -and -not $ln.Contains('*')) {
            $items.Add((New-KvittoRad -Name $m.Groups['name'].Value.Trim()))
        }
    }

    # Släng platshållare som aldrig fick ett belopp
    $kept = @($items | Where-Object { $null -ne $_.line_total })

    # Filnamnet kommer från Willys API (datum_butik_beloppkr.pdf) och är tillförlitligt:
    # - butik: alltid därifrån (korrekta åäö, samma format oavsett kvittots årgång)
    # - datum/total: reserv om PDF-texten saknar dem
    $mf = [regex]::Match($File.Name.Normalize([System.Text.NormalizationForm]::FormC),
        '^(?<d>\d{4}-\d{2}-\d{2})_(?<s>.+)_(?<a>\d+(?:\.\d+)?)kr(?:_\d+)?\.pdf$')
    if ($mf.Success) {
        $store = $mf.Groups['s'].Value.Replace('_', ' ')
        if (-not $date) { $date = $mf.Groups['d'].Value }
        if ($null -eq $total) {
            $total = [double]::Parse($mf.Groups['a'].Value, $script:Inv)
            Write-Verbose "Willys $($File.Name): total från filnamnet ($total)"
        }
    }

    return [pscustomobject]@{
        chain          = 'Willys'
        store          = $store
        date           = $date
        time           = $time
        currency       = 'SEK'
        total          = $total
        savings        = $savings
        payment_method = $payment
        source_file    = $File.Name
        items          = $kept
    }
}

# ==========================================================================
# Normalisering - Lidl (htmlPrintedReceipt ur JSON)
# ==========================================================================
function ConvertFrom-LidlKvitto {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)

    $data   = Get-Content -LiteralPath $File.FullName -Raw -Encoding utf8 | ConvertFrom-Json
    $ticket = Get-PropertyValue $data 'ticket'
    if ($null -eq $ticket) { throw 'Saknar ticket-objekt' }

    $storeObj = Get-PropertyValue $ticket 'store'
    $store    = ($storeObj -is [string]) ? $storeObj : (Get-PropertyValue $storeObj 'name')

    # PS7:s ConvertFrom-Json gör om ISO-datumsträngar till [datetime] automatiskt
    $date = $null
    $time = $null
    $dtRaw = Get-PropertyValue $ticket 'date'
    if ($dtRaw) {
        $dt = $null
        if ($dtRaw -is [datetime]) { $dt = $dtRaw }
        else {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse([string]$dtRaw, $script:Inv, [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
                $dt = $parsed
            }
        }
        if ($dt) {
            # Utc/Local -> svensk tid. Unspecified = saknar tidszon i källan, antas redan vara svensk tid.
            if ($dt.Kind -ne [System.DateTimeKind]::Unspecified) { $dt = ConvertTo-SvenskTid ([DateTimeOffset]$dt) }
            $date = $dt.ToString('yyyy-MM-dd', $script:Inv); $time = $dt.ToString('HH:mm', $script:Inv)
        }
    }

    $raw   = [string](Get-PropertyValue $ticket 'htmlPrintedReceipt')
    $items = [System.Collections.Generic.List[object]]::new()
    $A     = $script:Amount

    $reLine = [regex]::new('<span id="purchase_list_line_\d+"[^>]*class="(?<cls>article|discount)"[^>]*>(?<inner>.*?)</span>',
        [System.Text.RegularExpressions.RegexOptions]::Singleline)

    foreach ($m in $reLine.Matches($raw)) {
        $tag   = $m.Value
        $inner = [System.Net.WebUtility]::HtmlDecode($m.Groups['inner'].Value)
        $amts  = [regex]::Matches($inner, $A)

        if ($m.Groups['cls'].Value -eq 'article') {
            # Viktvarans andrarad ("1,152 kg x 24,90 SEK/kg") hoppas över
            if ($inner -match 'kg\s*x\s*[\d,]+\s*SEK/kg' -or $inner -match '\d+\s*x\s*[\d,]+\s*SEK') { continue }

            $md = [regex]::Match($tag, 'data-art-description="([^"]*)"')
            if (-not $md.Success) { continue }
            $mq = [regex]::Match($tag, 'data-art-quantity="([^"]*)"')
            $mu = [regex]::Match($tag, 'data-unit-price="([^"]*)"')

            $lineTotal = $amts.Count -gt 0 ? (ConvertFrom-SvensktTal $amts[$amts.Count - 1].Value) : $null
            $items.Add((New-KvittoRad `
                -Name ([System.Net.WebUtility]::HtmlDecode($md.Groups[1].Value)) `
                -Quantity ($mq.Success ? (ConvertFrom-SvensktTal $mq.Groups[1].Value) : 1) `
                -UnitPrice ($mu.Success ? (ConvertFrom-SvensktTal $mu.Groups[1].Value) : $lineTotal) `
                -LineTotal $lineTotal))
        }
        elseif ($amts.Count -gt 0 -and $items.Count -gt 0) {
            # Rabattrad -> knyts till föregående vara
            $last = $items[$items.Count - 1]
            $last.discount = ($last.discount ?? 0) + (ConvertFrom-SvensktTal $amts[$amts.Count - 1].Value)
        }
    }

    $decoded = [System.Net.WebUtility]::HtmlDecode($raw)
    $mt = [regex]::Match($decoded, 'ATT BETALA\s+(' + $A + ')')
    $ms = [regex]::Match($decoded, 'Total Lidl Plus-rabatt\s+(' + $A + ')')
    $mp = [regex]::Match($decoded, 'xxxx xxxx xxxx (\d{3,4})')

    return [pscustomobject]@{
        chain          = 'Lidl'
        store          = $store
        date           = $date
        time           = $time
        currency       = 'SEK'
        total          = $mt.Success ? (ConvertFrom-SvensktTal $mt.Groups[1].Value) : $null
        savings        = $ms.Success ? (ConvertFrom-SvensktTal $ms.Groups[1].Value) : $null
        payment_method = $mp.Success ? ('Kort *' + $mp.Groups[1].Value) : $null
        source_file    = $File.Name
        items          = @($items)
    }
}

# ==========================================================================
# Normalisering - skrivning (.NET, så att ShouldProcess bara frågar en gång per mapp)
# ==========================================================================
function Write-CsvFile {
    # Semikolonseparerad CSV med UTF-8 BOM (krävs för åäö i Excel för Mac)
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string[]]$Header,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows
    )
    $lines = if ($Rows.Count -gt 0) {
        @($Rows | ConvertTo-Csv -Delimiter ';' -UseQuotes AsNeeded)
    }
    else {
        @($Header -join ';')
    }
    [System.IO.File]::WriteAllText($Path, (($lines -join "`r`n") + "`r`n"), $script:Utf8Bom)
}

function Export-KvittoMapp {
    # Skriver kvitton.json + varor.csv och kopierar originalkvitton till mappen
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Receipts,
        [Parameter(Mandatory)][string]$Source
    )

    [void][System.IO.Directory]::CreateDirectory($Folder)

    $json = ConvertTo-Json -InputObject @($Receipts) -Depth 20
    [System.IO.File]::WriteAllText((Join-Path $Folder 'kvitton.json'), $json, $script:Utf8NoBom)

    $rows = foreach ($r in $Receipts) {
        foreach ($it in $r.items) {
            [pscustomobject]@{
                kedja        = $r.chain
                datum        = $r.date
                butik        = $r.store
                vara         = $it.name
                antal        = Format-Tal $it.quantity
                apris        = Format-Krona $it.unit_price
                radsumma     = Format-Krona $it.line_total
                rabatt       = Format-Krona $it.discount
                kvitto_total = Format-Krona $r.total
            }
        }
    }
    Write-CsvFile -Path (Join-Path $Folder 'varor.csv') -Rows @($rows) `
        -Header 'kedja', 'datum', 'butik', 'vara', 'antal', 'apris', 'radsumma', 'rabatt', 'kvitto_total'

    # Originalkvitton: Willys-PDF, Lidl-JSON + läsbar HTML
    foreach ($r in $Receipts) {
        $sub = ($r.chain -eq 'Willys') ? 'willys' : 'lidl'
        $src = Join-Path (Join-Path $Source $sub) $r.source_file
        if (Test-Path -LiteralPath $src -PathType Leaf) {
            [System.IO.File]::Copy($src, (Join-Path $Folder $r.source_file), $true)
        }
        if ($r.chain -eq 'Lidl') {
            # Läsbar HTML + PDF (för Paperless) om de finns
            foreach ($ext in '.html', '.pdf') {
                $extra = [System.IO.Path]::ChangeExtension($src, $ext)
                if (Test-Path -LiteralPath $extra -PathType Leaf) {
                    [System.IO.File]::Copy($extra, (Join-Path $Folder (Split-Path $extra -Leaf)), $true)
                }
            }
        }
    }
}

# ##########################################################################
# HUVUDFLÖDE
# ##########################################################################

# Absoluta sökvägar: .NET-metoderna (File.WriteAllText, File.Copy) löser annars
# relativa sökvägar mot processens arbetskatalog, som i VS Code ofta skiljer sig från $PWD.
$RawPath         = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RawPath)
$DestinationPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationPath)

if ($FromDate -gt $ToDate) { throw "FromDate ($($FromDate.ToString('yyyy-MM-dd'))) ligger efter ToDate ($($ToDate.ToString('yyyy-MM-dd')))." }

# S-02: cookies i miljövariabler hämtas till variabler och tas bort ur processmiljön, så att
# barnprocesser (pip, Playwright, Chromium, pdftotext) inte ärver dem. Återställs i finally.
$envCookies = @{}
foreach ($envName in 'WILLYS_COOKIE', 'LIDL_COOKIE') {
    $envCookies[$envName] = [System.Environment]::GetEnvironmentVariable($envName)
    [System.Environment]::SetEnvironmentVariable($envName, $null)
}

try {
    # --------------------------------------------------------------------------
    # Steg 1: Nedladdning
    # --------------------------------------------------------------------------
    if (-not $SkipDownload) {
        # Cookies hålls bara i minnet. Källa i prioritetsordning:
        # -AutoCookie (webbläsarfönster) > miljövariabler > dold inmatning
        $willysCookie = $null
        $lidlCookie   = $null
        if ($AutoCookie) {
            if ($PSCmdlet.ShouldProcess('webbläsarfönster', "Hämta inloggning för $($Chain -join ' och ')")) {
                $auto = Get-KvittoCookieAutomatisk -Chain $Chain -Browser $Browser
                $willysCookie = Get-PropertyValue $auto 'willys'
                $lidlCookie   = Get-PropertyValue $auto 'lidl'
            }
        }
        else {
            $willysCookie = $envCookies['WILLYS_COOKIE']
            $lidlCookie   = $envCookies['LIDL_COOKIE']
            # Fråga per kedja som saknar värde (utan inmatning, t.ex. schemalagt, varnar vi i stället)
            $saknas = @(
                if ('Willys' -in $Chain -and [string]::IsNullOrWhiteSpace($willysCookie)) { 'Willys' }
                if ('Lidl' -in $Chain -and [string]::IsNullOrWhiteSpace($lidlCookie)) { 'Lidl' }
            )
            if ($saknas.Count -gt 0 -and [Console]::IsInputRedirected) {
                Write-Warning "Cookie saknas för $($saknas -join ' och ') (WILLYS_COOKIE / LIDL_COOKIE) - kedjan hoppas över."
            }
            elseif ($saknas.Count -gt 0) {
                Write-Information 'Klistra in Cookie-raden från webbläsarens nätverksflik. Tomt = hoppa över kedjan.' -InformationAction Continue
                if ('Willys' -in $saknas) { $willysCookie = Read-Host -Prompt 'Willys-cookie (från willys.se/mina-kop)' -MaskInput }
                if ('Lidl' -in $saknas)   { $lidlCookie   = Read-Host -Prompt 'Lidl-cookie (från ett mre/api-anrop, hämta precis innan)' -MaskInput }
            }
        }
        if ('Willys' -notin $Chain) { $willysCookie = $null }
        if ('Lidl' -notin $Chain)   { $lidlCookie   = $null }

        $downloads = [System.Collections.Generic.List[object]]::new()

        # Lidl först - dess token lever bara ca 30 minuter
        if (-not [string]::IsNullOrWhiteSpace($lidlCookie)) {
            Write-Information 'Hämtar Lidl-kvitton...' -InformationAction Continue
            try {
                $downloads.Add((Save-LidlKvitto -CookieString $lidlCookie -OutputDir (Join-Path $RawPath 'lidl') -From $FromDate -To $ToDate))
            }
            catch [System.UnauthorizedAccessException] {
                Write-Warning "Lidl avbröts: $($PSItem.Exception.Message) Redan hämtade kvitton är sparade - hämta ny cookie och kör om."
            }
            catch {
                Write-Warning "Lidl avbröts av ett fel: $($PSItem.Exception.Message) Normaliseringen fortsätter."
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($willysCookie)) {
            Write-Information 'Hämtar Willys-kvitton...' -InformationAction Continue
            try {
                $downloads.Add((Save-WillysKvitto -CookieString $willysCookie -OutputDir (Join-Path $RawPath 'willys') -From $FromDate -To $ToDate))
            }
            catch [System.UnauthorizedAccessException] {
                Write-Warning "Willys avbröts: $($PSItem.Exception.Message) Redan hämtade kvitton är sparade - hämta ny cookie och kör om."
            }
            catch {
                Write-Warning "Willys avbröts av ett fel: $($PSItem.Exception.Message) Normaliseringen fortsätter."
            }
        }

        # Släpp cookies ur minnet så snart de inte behövs
        $willysCookie = $null
        $lidlCookie   = $null

        foreach ($d in $downloads) {
            Write-Information ('{0}: {1} i perioden, {2} nya, {3} fanns redan, {4} misslyckade' -f
                $d.Kedja, $d.Hittade, $d.Nya, $d.Fanns, $d.Misslyckade) -InformationAction Continue
        }
    }

    # --------------------------------------------------------------------------
    # Steg 1b: PDF av Lidl-kvitton (för Paperless) - även för tidigare nedladdade
    # --------------------------------------------------------------------------
    $lidlRawDir = Join-Path $RawPath 'lidl'
    if (-not $SkipPdf -and (Test-Path -LiteralPath $lidlRawDir -PathType Container)) {
        try {
            $created = ConvertTo-LidlPdf -LidlDir $lidlRawDir
            if ($created -gt 0) { Write-Information "Skapade $created Lidl-PDF:er." -InformationAction Continue }
        }
        catch {
            Write-Warning "Lidl-PDF:er kunde inte skapas: $($PSItem.Exception.Message) Normaliseringen fortsätter."
        }
    }

    # --------------------------------------------------------------------------
    # Steg 2: Normalisering av ALLA råkvitton
    # --------------------------------------------------------------------------
    if (-not (Test-Path -LiteralPath $RawPath -PathType Container)) {
        Write-Warning "Råkvittomappen finns inte: $RawPath"
        return
    }

    $receipts = [System.Collections.Generic.List[object]]::new()

    # Månader (yyyy-MM, eller 'Odaterade') där minst ett råkvitto inte gick att tolka. Allt skrivs om
    # från noll, så en månad med ofullständigt underlag får inte skrivas: en tidigare bra fil ska stå kvar.
    $blockedMonths = [System.Collections.Generic.HashSet[string]]::new()
    $failedFiles   = [System.Collections.Generic.List[string]]::new()

    function Block-KvittoManad {
        # Spärrar månaden (ur filnamnets datumprefix) för skrivning och minns filen till varningen.
        param([Parameter(Mandatory)][string]$FileName)

        $month = if ($FileName -match '^(\d{4}-\d{2})') { $Matches[1] } else { 'Odaterade' }
        [void]$blockedMonths.Add($month)
        $failedFiles.Add($FileName)
    }

    # Willys
    $willysDir = Join-Path $RawPath 'willys'
    $pdfFiles  = @(if (Test-Path -LiteralPath $willysDir) {
            Get-ChildItem -LiteralPath $willysDir -Filter '*.pdf' -File | Sort-Object -Property Name
        })
    if ($pdfFiles.Count -gt 0) {
        $pdfToText = Find-PdfToText -Hint $PdfToTextPath
        if (-not $pdfToText) {
            # Avbryt före skrivning: annars skrivs alla månadsfiler om utan Willys-kvitton
            throw "pdftotext saknas - $($pdfFiles.Count) Willys-kvitton kan inte tolkas. Installera med: brew install poppler"
        }
        Write-Verbose "Använder pdftotext: $pdfToText"
        foreach ($pdf in $pdfFiles) {
            try {
                $receipts.Add((ConvertFrom-WillysKvitto -File $pdf -PdfToText $pdfToText))
            }
            catch {
                Block-KvittoManad -FileName $pdf.Name
                Write-Warning "Willys $($pdf.Name): $($PSItem.Exception.Message)"
            }
        }
    }

    # Lidl
    $lidlDir   = Join-Path $RawPath 'lidl'
    $jsonFiles = @(if (Test-Path -LiteralPath $lidlDir) {
            Get-ChildItem -LiteralPath $lidlDir -Filter '*.json' -File |
                Where-Object { $_.Name -ne 'kvittolista.json' } |
                Sort-Object -Property Name
        })
    foreach ($js in $jsonFiles) {
        try {
            $receipts.Add((ConvertFrom-LidlKvitto -File $js))
        }
        catch {
            Block-KvittoManad -FileName $js.Name
            Write-Warning "Lidl $($js.Name): $($PSItem.Exception.Message)"
        }
    }

    if ($failedFiles.Count -gt 0) {
        Write-Warning ("{0} råkvitton gick inte att tolka. Berörda månader ({1}) skrivs inte och manad.json/manad.csv uppdateras inte. Filer: {2}" -f
            $failedFiles.Count, (($blockedMonths | Sort-Object) -join ', '), ($failedFiles -join ', '))
    }

    if ($receipts.Count -eq 0) {
        Write-Warning "Inga kvitton att normalisera i $RawPath"
        return
    }
    Write-Information "Normaliserar $($receipts.Count) kvitton..." -InformationAction Continue

    # Kontrollsumma per kvitto: rader + rabatter mot total
    foreach ($r in $receipts) {
        $radsum = 0.0
        foreach ($it in $r.items) { $radsum += ($it.line_total ?? 0) + ($it.discount ?? 0) }
        $diff = [math]::Round($radsum - ($r.total ?? 0), 2)
        if ([math]::Abs($diff) -ge 0.01) {
            Write-Warning ('DIFF {0} kr: {1} {2} {3} ({4}) - kontrollera mot originalet' -f
                (Format-Krona $diff), $r.date, $r.chain, $r.store, $r.source_file)
        }
    }

    # Skriv per år/månad
    $dated   = @($receipts | Where-Object { $_.date } | Sort-Object -Property date -Descending)
    $undated = @($receipts | Where-Object { -not $_.date })

    foreach ($group in ($dated | Group-Object -Property { $_.date.Substring(0, 7) } | Sort-Object -Property Name)) {
        $year   = $group.Name.Substring(0, 4)
        $mnum   = [int]$group.Name.Substring(5, 2)
        $mname  = '{0:D2}-{1}' -f $mnum, $script:Manader[$mnum]
        $folder = Join-Path (Join-Path $DestinationPath $year) $mname

        if ($blockedMonths.Contains($group.Name)) {
            Write-Warning "Hoppar över ${year}/${mname}: minst ett råkvitto för månaden gick inte att tolka. Befintliga filer lämnas orörda."
            continue
        }

        if ($PSCmdlet.ShouldProcess($folder, "Skriv $($group.Count) normaliserade kvitton")) {
            Export-KvittoMapp -Folder $folder -Receipts @($group.Group) -Source $RawPath
            [pscustomobject]@{ Manad = "$year/$mname"; Kvitton = $group.Count; Mapp = $folder }
        }
    }

    if ($undated.Count -gt 0 -and $blockedMonths.Contains('Odaterade')) {
        Write-Warning "Hoppar över Odaterade: ett råkvitto utan läsbart datum i filnamnet gick inte att tolka. Befintliga filer lämnas orörda."
    }
    elseif ($undated.Count -gt 0) {
        $folder = Join-Path $DestinationPath 'Odaterade'
        Write-Warning "$($undated.Count) kvitton saknar tolkbart datum och läggs i $folder"
        if ($PSCmdlet.ShouldProcess($folder, "Skriv $($undated.Count) odaterade kvitton")) {
            Export-KvittoMapp -Folder $folder -Receipts $undated -Source $RawPath
            [pscustomobject]@{ Manad = 'Odaterade'; Kvitton = $undated.Count; Mapp = $folder }
        }
    }

    # Översikt per månad och kedja
    $overview = @($dated |
            Group-Object -Property { '{0}|{1}' -f $_.date.Substring(0, 7), $_.chain } |
            ForEach-Object {
                $first = $_.Group[0]
                $sumTotal = 0.0; $sumSaved = 0.0; $nItems = 0
                foreach ($r in $_.Group) {
                    $sumTotal += $r.total ?? 0
                    $sumSaved += $r.savings ?? 0
                    $nItems   += @($r.items).Count
                }
                [pscustomobject]@{
                    manad         = $first.date.Substring(0, 7)
                    kedja         = $first.chain
                    antal_kvitton = $_.Count
                    summa_total   = [math]::Round($sumTotal, 2)
                    summa_sparat  = [math]::Round($sumSaved, 2)
                    antal_varor   = $nItems
                }
            } |
            Sort-Object -Property manad, kedja -Descending)

    if ($blockedMonths.Count -gt 0) {
        # Översikten räknas på alla kvitton; med ofullständigt underlag skulle den bli fel
        Write-Warning 'Hoppar över manad.json och manad.csv: underlaget är ofullständigt (se varning ovan).'
    }
    elseif ($PSCmdlet.ShouldProcess($DestinationPath, 'Skriv månadsöversikt (manad.json + manad.csv)')) {
        [void][System.IO.Directory]::CreateDirectory($DestinationPath)
        [System.IO.File]::WriteAllText((Join-Path $DestinationPath 'manad.json'),
            (ConvertTo-Json -InputObject $overview -Depth 20), $script:Utf8NoBom)

        $csvRows = foreach ($o in $overview) {
            [pscustomobject]@{
                manad         = $o.manad
                kedja         = $o.kedja
                antal_kvitton = $o.antal_kvitton
                summa_total   = Format-Krona $o.summa_total
                summa_sparat  = Format-Krona $o.summa_sparat
                antal_varor   = $o.antal_varor
            }
        }
        Write-CsvFile -Path (Join-Path $DestinationPath 'manad.csv') -Rows @($csvRows) `
            -Header 'manad', 'kedja', 'antal_kvitton', 'summa_total', 'summa_sparat', 'antal_varor'
    }
}
finally {
    foreach ($envName in $envCookies.Keys) {
        [System.Environment]::SetEnvironmentVariable($envName, $envCookies[$envName])
    }
}
