param([string]$Path, [int]$Cols = 3, [switch]$ShowRu)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$z = [IO.Compression.ZipFile]::OpenRead($Path)
$s = New-Object IO.StreamReader($z.GetEntry('word/document.xml').Open())
$xml = $s.ReadToEnd(); $s.Close(); $z.Dispose()
$doc = [xml]$xml
$ns = New-Object Xml.XmlNamespaceManager($doc.NameTable)
$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
function CellText($cell){
  $ps = $cell.SelectNodes('.//w:p', $ns)
  $out = foreach($p in $ps){ ($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '' }
  ((($out | Where-Object { $_.Trim() -ne '' }) -join ' ') -replace '\s+',' ').Trim()
}
$rowNo = 0
foreach($tbl in $doc.SelectNodes('//w:tbl', $ns)){
  foreach($tr in $tbl.SelectNodes('./w:tr', $ns)){
    $cells = @(foreach($tc in $tr.SelectNodes('./w:tc', $ns)){ CellText $tc })
    if($cells.Count -lt 2){ continue }
    $rowNo++
    $old = $cells[0]; $new = $cells[1]
    $ru = if($cells.Count -ge 3){ $cells[2] } else { '' }
    if($old -eq $new){ continue }
    $ow = $old -split '\s+' | Where-Object { $_ }
    $nw = $new -split '\s+' | Where-Object { $_ }
    $cmp = Compare-Object $ow $nw -SyncWindow 40
    $rem = ($cmp | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object { $_.InputObject }) -join ' '
    $add = ($cmp | Where-Object { $_.SideIndicator -eq '=>' } | ForEach-Object { $_.InputObject }) -join ' '
    if(-not $rem -and -not $add){ continue }
    "--- строка $rowNo"
    "  БЫЛО(нач): " + $old.Substring(0,[Math]::Min(220,$old.Length))
    if($rem){ "  УБРАНО:  " + $rem.Substring(0,[Math]::Min(700,$rem.Length)) }
    if($add){ "  ДОБАВЛЕНО: " + $add.Substring(0,[Math]::Min(700,$add.Length)) }
    if($ShowRu -and $ru){ "  RU: " + $ru.Substring(0,[Math]::Min(900,$ru.Length)) }
  }
}
