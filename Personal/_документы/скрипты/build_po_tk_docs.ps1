$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# --- список ПО ТК: берём актуальный файл Мурата (должности, подразделения, порядок)
$src = 'D:\ТАЗА\Список сотрудников ПО Таза Қазақстан.docx'
$out = 'D:\Кадры\Список ПО Таза Қазақстан - документы для оформления (2026-11-02).docx'

$fs = New-Object IO.FileStream($src, 'Open', 'Read', 'ReadWrite'); $z = New-Object IO.Compression.ZipArchive($fs)
$e = $z.GetEntry('word/document.xml'); $sr = New-Object IO.StreamReader($e.Open(), [Text.Encoding]::UTF8); $x = $sr.ReadToEnd(); $sr.Close(); $z.Dispose(); $fs.Dispose()
$list = @(); $dept = ''
foreach ($tr in [regex]::Matches($x, '<w:tr[ >].*?</w:tr>')) {
  $cells = @(foreach ($tc in [regex]::Matches($tr.Value, '<w:tc>.*?</w:tc>')) { [Net.WebUtility]::HtmlDecode((($tc.Value -replace '</w:p><w:p[ >]', ' / ' -replace '<[^>]+>', '').Trim())) })
  if ($cells.Count -eq 1) { if ($cells[0] -notlike 'Проектный офис «*') { $dept = $cells[0] } ; continue }
  if ($cells.Count -ge 3 -and $cells[0] -match '^\d+$') { $list += [pscustomobject]@{ n = [int]$cells[0]; pos = $cells[1]; name = $cells[2]; dept = $dept; start = ($cells[5] -replace '^(новая,\s*)?с\s*', '') } }
}
# даты выхода, уточнённые Муратом (в файле списка могут отличаться)
foreach ($p in $list) { if ($p.name -like 'Өміртай*') { $p.start = '03.11.2026' } }
foreach ($p in $list) { if ($p.name -like 'Сағынұлы*') { $p.start = '09.11.2026' } }
foreach ($p in $list) { if ($p.name -like 'Тленчиева*') { $p.start = '12.2026' } }   # выход в декабре, день уточняется (Мурат, 03.11.2026)   # выход с 9 ноября (Мурат, 03.11.2026)
"list rows: " + $list.Count

# --- данные из удостоверений (папка «Проекты ТД») и проверка папок Drive «Taza KZ HR документы» на 02.11.2026
$all  = 'Резюме; диплом; трудовая книжка; справка о счёте; справка об отсутствии судимости; справка об отсутствии коррупционных преступлений'
$info = [ordered]@{
  'Серикжанова' = @{ func = $true; fio = 'Серикжанова Айгерим Тельжановна'; iin = '941127400263'; td = $true;  miss = $all.Replace('Резюме; ', 'Резюме; ') }
  'Өміртай'     = @{ func = $true; fio = 'Өміртай Амалбек Даркенұлы';     iin = '940410301578'; td = $true;  miss = 'Справка об отсутствии судимости; справка об отсутствии коррупционных преступлений' }
  'Сағынұлы'    = @{ fio = 'Сағынұлы Санжар';              iin = '921128301118'; td = $true;  miss = 'Справка о счёте; справка об отсутствии судимости; справка об отсутствии коррупционных преступлений' }
  'Абенов'      = @{ func = $true; fio = 'Абенов Алибек Даниярулы';          iin = '941123351359'; td = $true;  miss = 'Справка об отсутствии судимости; справка об отсутствии коррупционных преступлений' }
  'Қуандық'     = @{ func = $true; fio = 'Қуандық Жасұлан Бостандықұлы';     iin = '970615300440'; td = $true;  miss = 'Трудовая книжка' }
  'Жандилдаев'  = @{ func = $true; fio = 'Жанділдаев Бақдаулет Болатұлы';    iin = '981208300214'; td = $true;  miss = $all }
  'Шербаева'    = @{ func = $true; fio = 'Шербаева Бибинур Байжигит қизи';   iin = '960320401579'; td = $true;  miss = 'Справка о счёте; справка об отсутствии судимости; справка об отсутствии коррупционных преступлений' }
  'Тастан'      = @{ func = $true; fio = 'Тастан Нұрбол Ерболұлы';           iin = '900505302684'; td = $true;  miss = '' }
  'Талғат'      = @{ func = $true; fio = 'Талғат Айман Талғатқызы';          iin = '980414450472'; td = $true;  miss = 'Трудовая книжка' }
  'Найла'       = @{ func = $true; fio = 'Мухамадиева Найля Азаматовна';     iin = '060205651277'; td = $true; miss = 'Диплом; трудовая книжка' }
  'Бектемиров'  = @{ fio = 'Бектемиров Амир Русланович';       iin = '000703550427'; td = $true; miss = 'Трудовая книжка' }
  'Жумабай'     = @{ func = $true; fio = 'Жумабай Бейбарыс Ғалымұлы';        iin = '970621300433'; td = $true; miss = $all }
}

function Esc([string]$s) { $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') }
$widths = @(450, 2400, 2100, 2400, 1400, 1250, 700, 850, 900, 3020)   # twips; сумма = 15470 (A4 альбомная)

function Cell($text, $w, $o) {
  $fill = if ($o.fill) { '<w:shd w:val="clear" w:color="auto" w:fill="' + $o.fill + '"/>' } else { '' }
  $span = if ($o.span) { '<w:gridSpan w:val="' + $o.span + '"/>' } else { '' }
  $jc = if ($o.jc) { $o.jc } else { 'left' }
  $b = if ($o.b) { '<w:b/>' } else { '' }
  $it = if ($o.i) { '<w:i/>' } else { '' }
  $col = if ($o.color) { '<w:color w:val="' + $o.color + '"/>' } else { '' }
  $sz = if ($o.sz) { $o.sz } else { 18 }
  '<w:tc><w:tcPr><w:tcW w:w="' + $w + '" w:type="dxa"/>' + $span + $fill + '<w:vAlign w:val="center"/></w:tcPr><w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="' + $jc + '"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + $b + $it + $col + '<w:sz w:val="' + $sz + '"/><w:szCs w:val="' + $sz + '"/></w:rPr><w:t xml:space="preserve">' + (Esc $text) + '</w:t></w:r></w:p></w:tc>'
}
function Row($cells, $header) {
  '<w:tr><w:trPr><w:cantSplit/>' + $(if ($header) { '<w:tblHeader/>' } else { '' }) + '</w:trPr>' + ($cells -join '') + '</w:tr>'
}
$green = 'D9EAD3'; $red = 'F4CCCC'; $grey = '7F7F7F'
$extraDocs = 'справка из психоневрологического диспансера; справка из наркологического диспансера'

$tbl = ''
$hdr = @('№', 'Должность', 'Структурное подразделение', 'ФИО (по удостоверению личности)', 'ИИН', 'Дата выхода на работу', 'Есть уд. л.', 'Есть проект ТД', 'Функционал в ТД', 'Не сдано из перечня')
$cells = @(); for ($i = 0; $i -lt 10; $i++) { $cells += Cell $hdr[$i] $widths[$i] @{ fill = '1B5E3B'; color = 'FFFFFF'; b = $true; jc = $(if ($i -ge 5 -and $i -le 8) { 'center' } else { 'left' }) } }
$tbl += Row $cells $true

function PersonRow($n, $pos, $dept, $fio, $iin, $start, $hasId, $hasTd, $miss, $fioGrey, $func, $noExtra) {
  $idTxt = if ($hasId) { 'да' } else { 'нет' }
  $tdTxt = if ($hasTd -is [string]) { $hasTd } elseif ($hasTd) { 'да' } else { 'нет' }
  $idFill = if ($hasId) { $green } else { $red }
  $tdFill = if ($hasTd -is [string]) { 'FFF2CC' } elseif ($hasTd) { $green } else { $red }
  $missFill = if ($miss -eq 'нет ни одного документа') { $red } else { $null }
  $fioFill = if ($miss -eq '' -or $miss -eq 'Трудовая книжка') { $green } else { $null }   # всё сдано или осталась только трудовая книжка (по исходным 8 пунктам перечня)
  if ($miss -ne '-' -and $miss -ne 'нет ни одного документа' -and -not $noExtra) { $miss = (($miss, $extraDocs) | Where-Object { $_ -ne '' }) -join '; '; $miss = $miss -replace '(судимости|преступлений|диспансера)(?! \(с eGov\))', '$1 (с eGov)' }   # справки с eGov помечаются; психоневрологическая и наркологическая справки: добавлены 03.11.2026, пока не сданы ни у кого
  $funcFill = if ($func -eq 'да') { $green } elseif ($func -eq 'нет') { $red } else { $null }
  Row @(
    (Cell $n $widths[0] @{ jc = 'center' }),
    (Cell $pos $widths[1] @{}),
    (Cell $dept $widths[2] @{}),
    (Cell $fio $widths[3] @{ color = $(if ($fioGrey) { $grey } else { $null }); i = $fioGrey; b = (-not $fioGrey); fill = $fioFill }),
    (Cell $iin $widths[4] @{}),
    (Cell $start $widths[5] @{ jc = 'center' }),
    (Cell $idTxt $widths[6] @{ jc = 'center'; fill = $idFill }),
    (Cell $tdTxt $widths[7] @{ jc = 'center'; fill = $tdFill }),
    (Cell $func $widths[8] @{ jc = 'center'; fill = $funcFill }),
    (Cell $miss $widths[9] @{ fill = $missFill })
  ) $false
}

$withId = 0; $withTd = 0; $complete = 0; $tdNoId = 0; $funcYes = 0
# три блока по дате выхода на работу: ноябрь 2026, декабрь 2026, февраль 2027
$blocks = @(
  @{ title = 'Выход на работу в ноябре 2026 года'; m = '\.11\.2026$' },
  @{ title = 'Выход на работу в декабре 2026 года'; m = '^(\d\d\.)?12\.2026$' },
  @{ title = 'Выход на работу в феврале 2027 года'; m = '^(\d\d\.)?02\.2027$' }
)
$num = 0
foreach ($bl in $blocks) {
  $tbl += Row @((Cell $bl.title 15470 @{ span = 10; fill = 'C9DDCF'; b = $true; sz = 20 })) $false
  $rows = @($list | Where-Object { $_.start -match $bl.m })
  if ($rows.Count -eq 0) { $tbl += Row @((Cell 'нет' 15470 @{ span = 10; i = $true; color = $grey })) $false; continue }
  foreach ($p in $rows) {
    $num++
    $key = $null
    foreach ($k in $info.Keys) { if ($p.name -like "$k*") { $key = $k; break } }
    if ($key) {
      $d = $info[$key]
      $tbl += PersonRow $num $p.pos $p.dept $d.fio $d.iin $p.start $true $d.td $d.miss $false $(if ($d.func) { 'да' } else { 'нет' }) $([bool]$d.noextra)
      $withId++; if ($d.td) { $withTd++ }; if ($d.miss -eq '') { $complete++ }; if ($d.func) { $funcYes++ }
    }
    elseif ($p.name -eq 'вакансия') {
      $tbl += PersonRow $num $p.pos $p.dept 'вакансия' '-' $p.start $false $false '-' $true '-'
    }
    elseif ($p.name -like 'Дутбаева*') {
      # черновик ТД с данными (№ и дата удостоверения, адрес, функционал) получен 03.11.2026; скана удостоверения нет
      $tbl += PersonRow $num $p.pos $p.dept 'Дутпаева Алия Кумархановна' '-' $p.start $false 'да, данные из черновика' 'нет ни одного документа' $true 'да'
      $tdNoId++; $funcYes++
    }
    elseif ($p.name -like 'Нурканат*') {
      # аутсорс: вместо ТД оформляется договор, проект ТД не нужен
      $tbl += PersonRow $num $p.pos $p.dept $p.name '-' $p.start $false 'нет (аутсорс)' 'нет ни одного документа' $true '-'
    }
    else {
      if ($p.start -match '\.11\.2026$' -or $p.name -like 'Тленчиева*') {
        # ноябрь: проект ТД собран без данных удостоверения (D:\Кадры\ТД предитоговые\... без уд.л.); февраль - ТД не делаем
        $tbl += PersonRow $num $p.pos $p.dept $p.name '-' $p.start $false 'да, без уд.л.' 'нет ни одного документа' $true 'нет'
        $tdNoId++
      } else {
        $tbl += PersonRow $num $p.pos $p.dept $p.name '-' $p.start $false $false 'нет ни одного документа' $true '-'
      }
    }
  }
}
$other = @($list | Where-Object { $s = $_.start; -not ($blocks | Where-Object { $s -match $_.m }) })
if ($other.Count) { throw ('дата выхода вне блоков: ' + (($other | % { $_.name + ' ' + $_.start }) -join '; ')) }$tblXml = '<w:tbl><w:tblPr><w:tblW w:w="15470" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/><w:left w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/><w:right w:val="single" w:sz="4" w:space="0" w:color="7F7F7F"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' + (($widths | ForEach-Object { '<w:gridCol w:w="' + $_ + '"/>' }) -join '') + '</w:tblGrid>' + $tbl + '</w:tbl>'

function Para($text, $sz, $bold, $after, $color) {
  $b = if ($bold) { '<w:b/>' } else { '' }
  $c = if ($color) { '<w:color w:val="' + $color + '"/>' } else { '' }
  '<w:p><w:pPr><w:spacing w:before="0" w:after="' + $after + '"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + $b + $c + '<w:sz w:val="' + $sz + '"/><w:szCs w:val="' + $sz + '"/></w:rPr><w:t xml:space="preserve">' + (Esc $text) + '</w:t></w:r></w:p>'
}
$head = (Para 'Проектный офис «Таза Қазақстан»: документы для оформления сотрудников' 30 $true 60 '1B5E3B') +
        (Para 'Перечень документов: 1) резюме; 2) удостоверение личности (электронно); 3) диплом (бакалавриат, магистратура); 4) копии сертификатов (при наличии); 5) трудовая книжка (бумажная, электронная); 6) справка о наличии счёта (предпочтительно Халык, БЦК); 7) справка об отсутствии судимости (с eGov); 8) справка из психоневрологического диспансера (с eGov); 9) справка из наркологического диспансера (с eGov); 10) справка об отсутствии коррупционных преступлений (с eGov). Сертификаты в графу «Не сдано» не включены (по наличию). По состоянию на 03.11.2026: папка «Проекты ТД» и Google Диск «Taza KZ HR документы».' 17 $false 160 $null)
$tail = (Para '' 12 $false 60 $null) +
        (Para ('Удостоверения личности: ' + $withId + ' из ' + $list.Count + ' позиций списка. Проект ТД: ' + $withTd + ' с данными удостоверения и ' + $tdNoId + ' без уд.л. (ФИО по списку, данные удостоверения и адрес - пустые поля). Функционал прописан в ' + $funcYes + ' из ' + ($withTd + $tdNoId) + ' проектов ТД.') 18 $true 0 $null)

$sect = '<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="680" w:right="680" w:bottom="680" w:left="680" w:header="400" w:footer="400" w:gutter="0"/></w:sectPr>'
$docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $head + $tblXml + $tail + $sect + '</w:body></w:document>'
$chk = New-Object Xml.XmlDocument; $chk.LoadXml($docXml)

$ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'
$rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
$core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Список ПО Таза Қазақстан - документы для оформления</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">2026-11-02T10:00:00Z</dcterms:created></cp:coreProperties>'
$app = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application><Company>CACF</Company></Properties>'

if (Test-Path -LiteralPath $out) { Copy-Item -LiteralPath $out -Destination ($out -replace '\.docx$', ' (копия до пересборки).docx') -Force }
$tmp = $out + '.tmp'; if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
$zip = [IO.Compression.ZipFile]::Open($tmp, 'Create')
foreach ($en in @(@('[Content_Types].xml', $ct), @('_rels/.rels', $rels), @('word/document.xml', $docXml), @('docProps/core.xml', $core), @('docProps/app.xml', $app))) {
  $ent = $zip.CreateEntry($en[0]); $sw = New-Object IO.StreamWriter($ent.Open(), (New-Object Text.UTF8Encoding $false)); $sw.Write($en[1]); $sw.Close()
}
$zip.Dispose()
Move-Item -LiteralPath $tmp -Destination $out -Force
"built: $out ; withId=$withId withTd=$withTd complete=$complete"
