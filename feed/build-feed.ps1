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
  $s = $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&apos;'
  return $s
}

function Price([string]$p) {
  if (-not $p) { return "0" }
  $clean = ($p -replace "[^\d.,]", "").Replace(",", "")
  return if ($clean -eq "") { "0" } else { $clean }
}

function Clean-Desc([string]$s) {
  if ($null -eq $s) { return "" }
  $s = [regex]::Replace($s, '(?i)</?(span|a|u|b|strong|i|em|s|small|big|sub|sup|mark|font)\b[^>]*>', '')
  $s = [regex]::Replace($s, '(?i)^\s*<div[^>]*>', '')
  $s = [regex]::Replace($s, '(?i)</div>\s*$', '')
  return $s.Replace("&nbsp;", " ").Replace("#nbsp;", " ").Trim()
}

function Get-ProductJson([string]$html) {
  $marker = "var product = "
  $i = $html.IndexOf($marker)
  if ($i -lt 0) { return $null }
  $start = $html.IndexOf("{", $i)
  if ($start -lt 0) { return $null }
  $depth = 0; $inStr = $false; $esc = $false
  for ($k = $start; $k -lt $html.Length; $k++) {
    $c = $html[$k]
    if ($inStr) {
      if ($esc) { $esc = $false } elseif ($c -eq "\") { $esc = $true } elseif ($c -eq '"') { $inStr = $false }
    } else {
      if ($c -eq '"') { $inStr = $true } elseif ($c -eq "{") { $depth++ } elseif ($c -eq "}") { $depth--; if ($depth -eq 0) { return $html.Substring($start, $k - $start + 1) } }
    }
  }
  return $null
}

function ConvertTo-CacheEntry($p) {
  $src = if ($p.characteristics) { @($p.characteristics) } elseif ($p.json_chars) { try { @($p.json_chars | ConvertFrom-Json) } catch { @() } } else { @() }
  $chars = @()
  foreach ($c in $src) { $t = ([string]$c.title).Trim(); if ($t -ne "") { $chars += , @($t, [string]$c.value) } }
  $pics = @(); foreach ($g in @($p.gallery)) { if ($null -ne $g -and [string]$g.img -ne "") { $pics += [string]$g.img } }
  $cat = ""; foreach ($u in @($p.partuids)) { if ([string]$u -ne "") { $cat = [string]$u; break } }
  $desc = Clean-Desc ([string]$p.text); if ($desc -eq "") { $desc = Clean-Desc ([string]$p.descr) }
  $pm = 0.0; try { $pm = [double]$p.pack_m } catch { }
  return @{ cat = $cat; w = $pm; t = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds(); chars = $chars; pics = $pics; desc = $desc }
}

$cache = New-Object System.Collections.Specialized.OrderedDictionary
if (Test-Path $CachePath) {
  try {
    $cacheRaw = [System.IO.File]::ReadAllText($CachePath, [System.Text.Encoding]::UTF8)
    if ($cacheRaw.Trim() -ne "") {
      $cacheObj = $cacheRaw | ConvertFrom-Json
      foreach ($prop in $cacheObj.PSObject.Properties) {
        $v = $prop.Value
        $cache[$prop.Name] = @{ cat = [string]$v.cat; w = [double]$v.w; t = $(try{[long]$v.t}catch{0}); chars = @($v.chars); pics = @($v.pics); desc = [string]$v.desc }
      }
    }
  } catch { }
}

Write-Output "== downloading =="
$yml = Get-Text $YmlUrl
$csv = Get-Text $CsvUrl

$ymlIds = @{}
foreach ($m in [regex]::Matches($yml, '<offer\b[^>]*\bid="([^"]+)"')) { $ymlIds[$m.Groups[1].Value] = $true }

$catByName = @{}; $catIds = @{}
foreach ($m in [regex]::Matches($yml, '<category id="([^"]+)"[^>]*>([^<]*)</category>')) {
  $cid = $m.Groups[1].Value; $catIds[$cid] = $true; $name = $m.Groups[2].Value.Trim().ToLower()
  if (-not $catByName.ContainsKey($name)) { $catByName[$name] = $cid }
}
$rootCat = "637007060011"
if (-not $catIds.ContainsKey($rootCat)) { $rootCat = @($catIds.Keys)[0] }

$rows = $csv | ConvertFrom-Csv
$sb = New-Object System.Text.StringBuilder
$seen = @{}; $fetches = 0; $nowSec = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

foreach ($r in $rows) {
  $id = [string]$r.id
  if ($id -eq "" -or $ymlIds.ContainsKey($id) -or $seen.ContainsKey($id)) { continue }
  $seen[$id] = $true

  $enrich = $null; $cachedEntry = $null
  if ($cache.Contains($id)) {
    $cachedEntry = $cache[$id]; $ts = 0; try{$ts=[long]$cachedEntry.t}catch{}
    if ($ts -gt 0 -and ($nowSec - $ts) -lt 2592000) { $enrich = $cachedEntry }
  }
  if (-not $enrich -and $fetches -lt 300 -and [string]$r.link -ne "") {
    $fetches++; try {
      $html = Get-Text ([string]$r.link); $raw = Get-ProductJson $html
      if ($raw) { $p = $raw | ConvertFrom-Json; if ([string]$p.uid -eq $id) { $enrich = ConvertTo-CacheEntry $p; $cache[$id] = $enrich } }
    } catch { }
  }
  if (-not $enrich -and $null -ne $cachedEntry) { $enrich = $cachedEntry }

  $catId = $rootCat
  if ($enrich -and [string]$enrich.cat -ne "" -and $catIds.ContainsKey([string]$enrich.cat)) { $catId = [string]$enrich.cat }
  else {
    $typeName = ([string]$r.product_type -replace "\\", "").Trim().ToLower()
    if ($typeName -ne "" -and $catByName.ContainsKey($typeName)) { $catId = $catByName[$typeName] }
  }

  $avail = ([string]$r.availability -eq "in stock")
  $availableAttr = if ($avail) { ' available="true"' } else { ' available="false"' }

  $desc = ""; $pics = @(); $weight = ""; $chars = @()
  if ($enrich) {
    $desc = [string]$enrich.desc; $pics = @($enrich.pics); $chars = @($enrich.chars)
    $pm = 0.0; try { $pm = [double]$enrich.w } catch { }
    if ($pm -gt 0) { $weight = ($pm / 1000.0).ToString("0.####", [System.Globalization.CultureInfo]::InvariantCulture) }
  }
  if ($desc -eq "") { $desc = Clean-Desc ([string]$r.description) }
  if ($pics.Count -eq 0 -and [string]$r.image_link -ne "") { $pics = @([string]$r.image_link) }

  [void]$sb.AppendLine("`t`t<offer id=`"$(Esc $id)`"$availableAttr>")
  [void]$sb.AppendLine("`t`t`t<name>$(Esc ([string]$r.title))</name>")
  if ([string]$r.brand -notmatch "\.") { [void]$sb.AppendLine("`t`t`t<vendor>$(Esc ([string]$r.brand))</vendor>") }
  if ($desc -ne "") { [void]$sb.AppendLine("`t`t`t<description><![CDATA[$($desc.Replace(']]>', ']]]]><![CDATA[>'))]]></description>") }
  foreach ($pic in $pics) { if ([string]$pic -ne "") { [void]$sb.AppendLine("`t`t`t<picture>$(Esc ([string]$pic))</picture>") } }
  if ([string]$r.link -ne "") { [void]$sb.AppendLine("`t`t`t<url>$(Esc ([string]$r.link))</url>") }
  
  # count only for available offers; out-of-stock ones keep available="false" without count
  if ($avail) { [void]$sb.AppendLine("`t`t`t<count>1</count>") }
  [void]$sb.AppendLine("`t`t`t<price>$(Price ([string]$r.price))</price>")
  [void]$sb.AppendLine("`t`t`t<currencyId>RUR</currencyId>")
  [void]$sb.AppendLine("`t`t`t<categoryId>$catId</categoryId>")
  if ($weight -ne "") { [void]$sb.AppendLine("`t`t`t<weight>$weight</weight>") }
  foreach ($c in $chars) { if ($null -eq $c) { continue }; $t = [string]$c[0]; $v = [string]$c[1]; if ($t -ne "") { [void]$sb.AppendLine("`t`t`t<param name=`"$(Esc $t)`">$(Esc $v)</param>") } }
  [void]$sb.AppendLine("`t`t</offer>")
}

$idx = $yml.LastIndexOf("</offers>")
$out = $yml.Substring(0, $idx) + $sb.ToString() + $yml.Substring($idx)
$dateStr = [DateTime]::UtcNow.AddHours(3).ToString("yyyy-MM-ddTHH:mm:ss") + "+03:00"
$out = [regex]::Replace($out, '<yml_catalog date="[^"]*"', '<yml_catalog date="' + $dateStr + '"', 1)

$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutPath, $out, $utf8)

# Сохранение кэша
if ($cache.Count -gt 0) {
  $jp = New-Object System.Text.StringBuilder; [void]$jp.Append("{")
  $firstKey = $true; foreach ($key in $cache.Keys) {
    if (-not $firstKey) { [void]$jp.Append(",") }; $firstKey = $false; $e = $cache[$key]
    [void]$jp.Append('"' + $key + '":{"cat":"' + $e.cat + '","w":' + ([double]$e.w).ToString("0.####", [System.Globalization.CultureInfo]::InvariantCulture) + ',"t":' + $e.t + ',"chars":[')
    $fc = $true; foreach ($c in @($e.chars)) { if (-not $fc) { [void]$jp.Append(",") }; $fc = $false; [void]$jp.Append('["' + (Esc $c[0]) + '","' + (Esc $c[1]) + '"]') }
    [void]$jp.Append('],"pics":[')
    $fp = $true; foreach ($p in @($e.pics)) { if (-not $fp) { [void]$jp.Append(",") }; $fp = $false; [void]$jp.Append('"' + $p + '"') }
    [void]$jp.Append('],"desc":"' + (Esc $e.desc) + '"}')
  }
  [void]$jp.Append("}"); [System.IO.File]::WriteAllText($CachePath, $jp.ToString(), $utf8)
}
Write-Output "DONE"
