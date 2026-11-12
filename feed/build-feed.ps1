param(
  [string]$OutPath = "feed.xml",
  [string]$YmlUrl = "https://insklad-nt.ru/tstore/yml/4f1af8184cf87bc6f9b0026b48823f09.yml",
  [string]$CsvUrl = "https://insklad-nt.ru/feed-fb.csv",
  [string]$CachePath = "",
  [int]$MaxFetches = 300
)

$ErrorActionPreference = "Stop"
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch { }

if ($CachePath -eq "") {
  if ($PSScriptRoot) { $CachePath = Join-Path $PSScriptRoot "enrich-cache.json" }
  else { $CachePath = Join-Path (Get-Location) "enrich-cache.json" }
}

$ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"
$client = New-Object System.Net.Http.HttpClient
$client.Timeout = [TimeSpan]::FromSeconds(60)
$client.DefaultRequestHeaders.Add("User-Agent", $ua)

function Get-Text([string]$url) {
  $sep = if ($url.Contains("?")) { "&" } else { "?" }
  $resp = $client.GetAsync($url + $sep + "cb=" + [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()).Result
  if (-not $resp.IsSuccessStatusCode) { throw ("HTTP " + [int]$resp.StatusCode + " for " + $url) }
  $bytes = $resp.Content.ReadAsByteArrayAsync().Result
  return [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Esc([string]$s) {
  if ($null -eq $s) { return "" }
  $s = $s -replace '&', '&amp;'
  $s = $s -replace '<', '&lt;'
  $s = $s -replace '>', '&gt;'
  $s = $s -replace '"', '&quot;'
  $s = $s -replace "'", '&apos;'
  return $s
}
function Price([string]$p) {
  if (-not $p) { return "0" }
  $clean = ($p -replace "[^\d.,]", "").Replace(",", "")
  if ($clean -eq "") { return "0" }
  return $clean
}

# Tilda renders the description as HTML with inline formatting; its own YML strips those
# tags (span, a, u, b, ...) and the outer <div>, keeping only block markup.
function Clean-Desc([string]$s) {
  if ($null -eq $s) { return "" }
  $s = [regex]::Replace($s, '(?i)</?(span|a|u|b|strong|i|em|s|small|big|sub|sup|mark|font)\b[^>]*>', '')
  $s = [regex]::Replace($s, '(?i)^\s*<div[^>]*>', '')
  $s = [regex]::Replace($s, '(?i)</div>\s*$', '')
  $s = $s.Replace("&nbsp;", " ").Replace("#nbsp;", " ")
  return $s.Trim()
}

# Extract the "var product = {...}" JSON object from a public Tilda product page.
function Get-ProductJson([string]$html) {
  $marker = "var product = "
  $i = $html.IndexOf($marker)
  if ($i -lt 0) { return $null }
  $start = $html.IndexOf("{", $i)
  if ($start -lt 0) { return $null }
  $depth = 0
  $inStr = $false
  $esc = $false
  for ($k = $start; $k -lt $html.Length; $k++) {
    $c = $html[$k]
    if ($inStr) {
      if ($esc) { $esc = $false }
      elseif ($c -eq "\") { $esc = $true }
      elseif ($c -eq '"') { $inStr = $false }
    } else {
      if ($c -eq '"') { $inStr = $true }
      elseif ($c -eq "{") { $depth++ }
      elseif ($c -eq "}") {
        $depth--
        if ($depth -eq 0) { return $html.Substring($start, $k - $start + 1) }
      }
    }
  }
  return $null
}

function ConvertTo-CacheEntry($p) {
  $src = @()
  if ($p.characteristics) { $src = @($p.characteristics) }
  elseif ($p.json_chars) {
    try { $src = @($p.json_chars | ConvertFrom-Json) } catch { $src = @() }
  }

  $chars = @()
  foreach ($c in $src) {
    $t = ([string]$c.title).Trim()
    if ($t -eq "") { continue }
    $chars += , @($t, [string]$c.value)
  }

  $pics = @()
  foreach ($g in @($p.gallery)) {
    if ($null -ne $g -and [string]$g.img -ne "") { $pics += [string]$g.img }
  }

  $cat = ""
  foreach ($u in @($p.partuids)) {
    if ([string]$u -ne "") { $cat = [string]$u; break }
  }

  $desc = Clean-Desc ([string]$p.text)
  if ($desc -eq "") { $desc = Clean-Desc ([string]$p.descr) }

  $pm = 0.0
  try { $pm = [double]$p.pack_m } catch { $pm = 0.0 }

  return @{
    cat   = $cat
    w     = $pm
    t     = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    chars = $chars
    pics  = $pics
    desc  = $desc
  }
}

# --- enrichment cache (committed to the repo, avoids re-fetching unchanged products)
$cache = New-Object System.Collections.Specialized.OrderedDictionary
if (Test-Path $CachePath) {
  try {
    $cacheRaw = [System.IO.File]::ReadAllText($CachePath, [System.Text.Encoding]::UTF8)
    if ($cacheRaw.Trim() -ne "") {
      $cacheObj = $cacheRaw | ConvertFrom-Json
      foreach ($prop in $cacheObj.PSObject.Properties) {
        $v = $prop.Value
        $cache[$prop.Name] = @{
          cat   = [string]$v.cat
          w     = [double]$v.w
          t     = $(try { [long]$v.t } catch { 0 })
          chars = @($v.chars)
          pics  = @($v.pics)
          desc  = [string]$v.desc
        }
      }
    }
  } catch {
    Write-Output ("CACHE_READ_ERROR=" + $_.Exception.Message)
    $cache = New-Object System.Collections.Specialized.OrderedDictionary
  }
}
$cacheSizeBefore = $cache.Count

Write-Output "== downloading =="
$yml = Get-Text $YmlUrl
$csv = Get-Text $CsvUrl
Write-Output ("YML_BYTES=" + $yml.Length + " CSV_BYTES=" + $csv.Length)

# --- existing offer ids
$ymlIds = @{}
foreach ($m in [regex]::Matches($yml, '<offer\b[^>]*\bid="([^"]+)"')) {
  $ymlIds[$m.Groups[1].Value] = $true
}
Write-Output ("YML_OFFERS=" + $ymlIds.Count)

# --- categories: name -> id and known id set
$catByName = @{}
$catIds = @{}
foreach ($m in [regex]::Matches($yml, '<category id="([^"]+)"[^>]*>([^<]*)</category>')) {
  $cid = $m.Groups[1].Value
  $catIds[$cid] = $true
  $name = $m.Groups[2].Value.Trim().ToLower()
  if (-not $catByName.ContainsKey($name)) { $catByName[$name] = $cid }
}
$rootCat = "637007060011"
if (-not $catIds.ContainsKey($rootCat)) {
  $rootCat = @($catIds.Keys)[0]
}
Write-Output ("CATEGORIES=" + $catByName.Count + " ROOT=" + $rootCat)

# --- manual fallback: csv product_type is empty for these, category matched by title
$manualCats = @(
  @{ Pattern = "Захват магнитный"; CatId = "625505081193" }  # Индикаторы часового типа, стойки магнитные
)

# --- parse csv
$rows = $csv | ConvertFrom-Csv
Write-Output ("CSV_ROWS=" + @($rows).Count)

$sb = New-Object System.Text.StringBuilder
$added = 0
$noCategory = 0
$enriched = 0
$fromCache = 0
$fallback = 0
$fetches = 0
$seen = @{}
$nowSec = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
$maxCacheAge = 30 * 86400   # refresh product data at most once a month

foreach ($r in $rows) {
  $id = [string]$r.id
  if ($id -eq "" -or $ymlIds.ContainsKey($id) -or $seen.ContainsKey($id)) { continue }
  $seen[$id] = $true

  # --- full product data: fresh cache first, then the public product page
  $enrich = $null
  $cachedEntry = $null
  if ($cache.Contains($id)) {
    $cachedEntry = $cache[$id]
    $ts = 0
    try { $ts = [long]$cachedEntry.t } catch { $ts = 0 }
    if ($ts -gt 0 -and ($nowSec - $ts) -lt $maxCacheAge) {
      $enrich = $cachedEntry
      $fromCache++
    }
  }
  if (-not $enrich -and $fetches -lt $MaxFetches -and [string]$r.link -ne "") {
    $fetches++
    try {
      $html = Get-Text ([string]$r.link)
      $raw = Get-ProductJson $html
      if ($raw) {
        $p = $raw | ConvertFrom-Json
        if ([string]$p.uid -eq $id) {
          $enrich = ConvertTo-CacheEntry $p
          $cache[$id] = $enrich
          $enriched++
        }
      }
    } catch {
      # keep going: this offer falls back to cached/CSV data and is retried next run
    }
  }
  if (-not $enrich -and $null -ne $cachedEntry) {
    # refresh failed: better stale product data than none
    $enrich = $cachedEntry
    $fromCache++
  }

  # --- category: product partuid (same ids as YML) -> csv product_type -> manual -> root
  $catId = $rootCat
  $resolved = $false
  if ($enrich -and [string]$enrich.cat -ne "" -and $catIds.ContainsKey([string]$enrich.cat)) {
    $catId = [string]$enrich.cat
    $resolved = $true
  }
  if (-not $resolved) {
    $typeName = ([string]$r.product_type -replace "\\", "").Trim().ToLower()
    if ($typeName -ne "" -and $catByName.ContainsKey($typeName)) {
      $catId = $catByName[$typeName]
    } else {
      $hit = $null
      foreach ($mc in $manualCats) {
        if ([string]$r.title -match $mc.Pattern) { $hit = $mc.CatId; break }
      }
      if ($hit) { $catId = $hit } else { $noCategory++ }
    }
  }

  $avail = ([string]$r.availability -eq "in stock")
  $availableAttr = if ($avail) { ' available="true"' } else { ' available="false"' }

  # --- description / pictures / weight / params
  $desc = ""
  $pics = @()
  $weight = ""
  $chars = @()

  if ($enrich) {
    $desc = [string]$enrich.desc
    $pics = @($enrich.pics)
    $chars = @($enrich.chars)
    $pm = 0.0
    try { $pm = [double]$enrich.w } catch { $pm = 0.0 }
    if ($pm -gt 0) {
      $weight = ($pm / 1000.0).ToString("0.####", [System.Globalization.CultureInfo]::InvariantCulture)
    }
  }

  if ($desc -eq "") {
    $desc = Clean-Desc ([string]$r.description)
    $fallback++
  }
  if ($pics.Count -eq 0 -and [string]$r.image_link -ne "") {
    $pics = @([string]$r.image_link)
  }

  $descCdata = $desc.Replace("]]>", "]]]]><![CDATA[>")

  # vendor: keep only real brands, skip site-domain placeholders like "insklad-nt.ru"
  $brand = [string]$r.brand
  $brandOut = if ($brand -ne "" -and $brand -notmatch "\.") { $brand } else { "" }

  [void]$sb.AppendLine("`t`t<offer id=`"$(Esc $id)`"$availableAttr>")
  [void]$sb.AppendLine("`t`t`t<name>$(Esc ([string]$r.title))</name>")
  if ($brandOut -ne "") {
    [void]$sb.AppendLine("`t`t`t<vendor>$(Esc $brandOut)</vendor>")
  }
  if ($desc -ne "") {
    [void]$sb.AppendLine("`t`t`t<description><![CDATA[$descCdata]]></description>")
  }
  foreach ($pic in $pics) {
    if ([string]$pic -ne "") {
      [void]$sb.AppendLine("`t`t`t<picture>$(Esc ([string]$pic))</picture>")
    }
  }
  if ([string]$r.link -ne "") {
    [void]$sb.AppendLine("`t`t`t<url>$(Esc ([string]$r.link))</url>")
  }
  # no <count>: an out-of-stock offer must not claim "0 pcs"
  [void]$sb.AppendLine("`t`t`t<price>$(Price ([string]$r.price))</price>")
  [void]$sb.AppendLine("`t`t`t<currencyId>RUR</currencyId>")
  [void]$sb.AppendLine("`t`t`t<categoryId>$catId</categoryId>")
  if ($weight -ne "") {
    [void]$sb.AppendLine("`t`t`t<weight>$weight</weight>")
  }
  foreach ($c in $chars) {
    if ($null -eq $c) { continue }
    $t = [string]$c[0]
    $v = [string]$c[1]
    if ($t -eq "") { continue }
    [void]$sb.AppendLine("`t`t`t<param name=`"$(Esc $t)`">$(Esc $v)</param>")
  }
  [void]$sb.AppendLine("`t`t</offer>")
  $added++
}

$idx = $yml.LastIndexOf("</offers>")
if ($idx -lt 0) { throw "No </offers> in YML" }

# --- safety guards: never publish a truncated/broken feed
if ($ymlIds.Count -lt 4000) { throw ("YML too small: " + $ymlIds.Count + " offers") }
if (@($rows).Count -lt 1000) { throw ("CSV too small: " + @($rows).Count + " rows") }
if ($added -gt 2000) { throw ("Too many added offers: " + $added) }

$out = $yml.Substring(0, $idx) + $sb.ToString() + $yml.Substring($idx)

# --- refresh feed date (Moscow time)
$dateStr = [DateTime]::UtcNow.AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss") + "+03:00"
$out = [regex]::Replace($out, '<yml_catalog date="[^"]*"', '<yml_catalog date="' + $dateStr + '"', 1)

$dir = Split-Path -Parent $OutPath
if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

# --- validate before writing
$doc = New-Object System.Xml.XmlDocument
$doc.LoadXml($out)
$offerNodes = $doc.SelectNodes("//offer")
$falseCount = $doc.SelectNodes("//offer[@available='false']").Count
if ($offerNodes.Count -lt 4500) { throw ("Resulting feed too small: " + $offerNodes.Count + " offers") }

$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutPath, $out, $utf8)

# --- save the enrichment cache (compact, readable, no \uXXXX escaping; works on PS 5.1 and 7)
function JsonStr([string]$s) {
  if ($null -eq $s) { return "" }
  $out = New-Object System.Text.StringBuilder
  foreach ($ch in $s.ToCharArray()) {
    $code = [int]$ch
    if ($ch -eq '"') { [void]$out.Append('\"') }
    elseif ($ch -eq "\") { [void]$out.Append('\\') }
    elseif ($ch -eq "`n") { [void]$out.Append('\n') }
    elseif ($ch -eq "`r") { [void]$out.Append('\r') }
    elseif ($ch -eq "`t") { [void]$out.Append('\t') }
    elseif ($code -lt 32) { [void]$out.Append(('\u{0:x4}' -f $code)) }
    else { [void]$out.Append($ch) }
  }
  return $out.ToString()
}

if ($cache.Count -gt 0) {
  try {
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $jp = New-Object System.Text.StringBuilder
    [void]$jp.Append("{")
    $firstKey = $true
    foreach ($key in $cache.Keys) {
      if (-not $firstKey) { [void]$jp.Append(",") }
      $firstKey = $false
      $e = $cache[$key]
      [void]$jp.Append('"' + (JsonStr ([string]$key)) + '":{"cat":"' + (JsonStr ([string]$e.cat)) + '","w":')
      $w = 0.0
      try { $w = [double]$e.w } catch { $w = 0.0 }
      [void]$jp.Append($w.ToString("0.####", $inv))
      $ts = 0
      try { $ts = [long]$e.t } catch { $ts = 0 }
      [void]$jp.Append(',"t":' + $ts.ToString($inv))
      [void]$jp.Append(',"chars":[')
      $firstChar = $true
      foreach ($c in @($e.chars)) {
        if ($null -eq $c) { continue }
        if (-not $firstChar) { [void]$jp.Append(",") }
        $firstChar = $false
        [void]$jp.Append('["' + (JsonStr ([string]$c[0])) + '","' + (JsonStr ([string]$c[1])) + '"]')
      }
      [void]$jp.Append('],"pics":[')
      $firstPic = $true
      foreach ($pic in @($e.pics)) {
        if ($null -eq $pic -or [string]$pic -eq "") { continue }
        if (-not $firstPic) { [void]$jp.Append(",") }
        $firstPic = $false
        [void]$jp.Append('"' + (JsonStr ([string]$pic)) + '"')
      }
      [void]$jp.Append('],"desc":"' + (JsonStr ([string]$e.desc)) + '"}')
    }
    [void]$jp.Append("}")
    $cacheDir = Split-Path -Parent $CachePath
    if ($cacheDir -and -not (Test-Path $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }
    [System.IO.File]::WriteAllText($CachePath, $jp.ToString(), $utf8)
    Write-Output ("CACHE_SAVED=" + $cache.Count + " (was " + $cacheSizeBefore + ")")
  } catch {
    Write-Output ("CACHE_WRITE_ERROR=" + $_.Exception.Message)
  }
}

Write-Output ("ADDED=" + $added + " NO_CATEGORY=" + $noCategory)
Write-Output ("ENRICHED_NEW=" + $enriched + " FROM_CACHE=" + $fromCache + " CSV_FALLBACK_DESC=" + $fallback + " FETCHES=" + $fetches)
Write-Output ("OUT_PATH=" + (Resolve-Path $OutPath))
Write-Output ("OUT_BYTES=" + (Get-Item $OutPath).Length)
Write-Output ("OFFERS_TOTAL=" + $offerNodes.Count + " AVAILABLE_FALSE=" + $falseCount)
Write-Output "DONE"
