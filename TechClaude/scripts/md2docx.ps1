param([string]$Src, [string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Esc($s){ $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' }

# разбор строки на куски с разметкой {+add+} {-del-} {~note~} {!hl!}
function Runs($line){
  $res = @()
  $rx = [regex]'\{(\+|\-|~|!)(.*?)\1\}'
  $pos = 0
  foreach($m in $rx.Matches($line)){
    if($m.Index -gt $pos){ $res += ,@('n', $line.Substring($pos, $m.Index - $pos)) }
    $res += ,@($m.Groups[1].Value, $m.Groups[2].Value)
    $pos = $m.Index + $m.Length
  }
  if($pos -lt $line.Length){ $res += ,@('n', $line.Substring($pos)) }
  ,$res
}

function RunXml($kind, $text, $baseSize, $italic){
  $rpr = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="' + $baseSize + '"/>'
  switch($kind){
    '+' { $rpr += '<w:color w:val="1F7A33"/><w:u w:val="single"/><w:b/>' }
    '-' { $rpr += '<w:color w:val="C00000"/><w:strike/>' }
    '~' { $rpr += '<w:color w:val="7F7F7F"/><w:i/>' }
    '!' { $rpr += '<w:highlight w:val="yellow"/>' }
  }
  if($italic){ $rpr += '<w:i/>' }
  $t = Esc $text
  '<w:r><w:rPr>' + $rpr + '</w:rPr><w:t xml:space="preserve">' + $t + '</w:t></w:r>'
}

function CellPara($text, $isHead){
  $runs = Runs $text
  $xml = foreach($r in $runs){
    $x = RunXml $r[0] $r[1] 18 $false
    if($isHead -and $r[0] -eq 'n'){ $x = $x.Replace('<w:sz w:val="18"/>', '<w:sz w:val="18"/><w:b/>') }
    $x
  }
  '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>' + ($xml -join '') + '</w:p>'
}
function TableXml($rows){
  $head = $true
  $trs = foreach($cells in $rows){
    $tcs = foreach($c in $cells){
      '<w:tc><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4" w:color="999999"/><w:left w:val="single" w:sz="4" w:color="999999"/><w:bottom w:val="single" w:sz="4" w:color="999999"/><w:right w:val="single" w:sz="4" w:color="999999"/></w:tcBorders></w:tcPr>' + (CellPara $c $head) + '</w:tc>'
    }
    $row = '<w:tr>' + ($tcs -join '') + '</w:tr>'
    $head = $false
    $row
  }
  '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="autofit"/></w:tblPr>' + ($trs -join '') + '</w:tbl><w:p/>'
}

$body = New-Object Text.StringBuilder
$tableRows = @()
$lines = @(Get-Content -LiteralPath $Src -Encoding UTF8)
foreach($raw in $lines){
  $line = $raw.TrimEnd()
  if($line.StartsWith('|')){
    $cells = @($line.Trim('|') -split '\|' | ForEach-Object { $_.Trim() })
    if(($cells -join '') -match '^[-: ]+$'){ continue }
    $tableRows += ,$cells
    continue
  } elseif($tableRows.Count){
    [void]$body.Append((TableXml $tableRows)); $tableRows = @()
  }
  if($line -eq ''){ [void]$body.Append("<w:p/>"); continue }
  $size = 22; $align = 'both'; $spaceBefore = 0; $italic = $false; $bold = $false
  if($line -like '# *'){ $line = $line.Substring(2); $size = 28; $align='center'; $bold=$true; $spaceBefore=240 }
  elseif($line -like '## *'){ $line = $line.Substring(3); $size = 24; $align='left'; $bold=$true; $spaceBefore=240 }
  elseif($line -like '### *'){ $line = $line.Substring(4); $size = 22; $align='left'; $bold=$true; $spaceBefore=120 }
  elseif($line -like '> *'){ $line = $line.Substring(2); $size = 20; $italic=$true }
  $ppr = '<w:pPr><w:spacing w:before="' + $spaceBefore + '" w:after="60" w:line="276" w:lineRule="auto"/><w:jc w:val="' + $align + '"/></w:pPr>'
  $runs = Runs $line
  $xml = foreach($r in $runs){
    $k = $r[0]; $txt = $r[1]
    $x = RunXml $k $txt $size $italic
    if($bold -and $k -eq 'n'){ $x = $x.Replace('<w:sz w:val="' + $size + '"/>', '<w:sz w:val="' + $size + '"/><w:b/>') }
    $x
  }
  [void]$body.Append("<w:p>$ppr" + ($xml -join '') + "</w:p>")
}
if($tableRows.Count){ [void]$body.Append((TableXml $tableRows)) }

$docXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>$($body.ToString())
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>
"@

$ct = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"@
$rels = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"@

$tmp = Join-Path ([IO.Path]::GetTempPath()) ("docx_" + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force "$tmp\_rels","$tmp\word" | Out-Null
$enc = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText("$tmp\[Content_Types].xml", $ct, $enc)
[IO.File]::WriteAllText("$tmp\_rels\.rels", $rels, $enc)
[IO.File]::WriteAllText("$tmp\word\document.xml", $docXml, $enc)
if(Test-Path -LiteralPath $Out){ Remove-Item -LiteralPath $Out -Force }
[IO.Compression.ZipFile]::CreateFromDirectory($tmp, $Out)
Remove-Item $tmp -Recurse -Force
"OK: $Out ($([math]::Round((Get-Item -LiteralPath $Out).Length/1KB)) КБ)"
