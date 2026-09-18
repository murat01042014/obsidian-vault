param([string]$Path)
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
  (($out | Where-Object { $_ -ne '' }) -join ' / ')
}
$body = $doc.SelectSingleNode('//w:body', $ns)
$i = 0
foreach($node in $body.ChildNodes){
  if($node.LocalName -eq 'tbl'){
    $i++; "=== ТАБЛИЦА $i ==="
    foreach($tr in $node.SelectNodes('./w:tr', $ns)){
      $cells = foreach($tc in $tr.SelectNodes('./w:tc', $ns)){ CellText $tc }
      ($cells -join ' || ')
    }
  } elseif($node.LocalName -eq 'p'){
    $t = ($node.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ''
    if($t.Trim()){ "[текст] " + $t.Trim() }
  }
}
