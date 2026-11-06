$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tpl = 'C:\Users\User\Documents\письма\Предложение CACF.docx'
$out = 'D:\Кадры\Справки\Справка с места работы - Қаражан Е.Н. (Япония, исх. 02-397 от 06.11.2026).docx'
Copy-Item -LiteralPath $tpl -Destination $out -Force

# a - выравнивание; ind - отступ слева, см; fl - красная строка (1 = да); sa - интервал после, пт; sz - кегль; b - жирный
$spec = @(
  @{ t = 'Исх. № 02/397 от 06.11.2026 г.'; a = 'left'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'Консульский отдел Посольства Японии в Республике Казахстан'; a = 'left'; ind = 8.5; fl = 0; sa = 14; sz = 11; b = $true },
  @{ t = 'СПРАВКА'; a = 'center'; ind = 0; fl = 0; sa = 1; sz = 13; b = $true },
  @{ t = 'с места работы'; a = 'center'; ind = 0; fl = 0; sa = 10; sz = 11 },
  @{ t = 'Дана Қаражан Ерлік Нариманұлы, 30 марта 1978 года рождения, ИИН 780330300085, паспорт № N18308567, в том, что он работает в Центрально-Азиатском климатическом фонде в должности Главного исполнительного директора, руководителя Проектного офиса для Центральной Азии по изменению климата и зелёной энергетике, на основании трудового договора № 1-1 от 18 февраля 2025 года, с 18 февраля 2025 года и по настоящее время. Работа в Фонде является основным местом работы.'; a = 'both'; ind = 0; fl = 1; sa = 8; sz = 11 },
  @{ t = 'На период служебной командировки в г. Токио (Япония) с 16 по 23 ноября 2026 года за ним сохраняются место работы, должность и заработная плата. По возвращении он приступает к исполнению своих обязанностей.'; a = 'both'; ind = 0; fl = 1; sa = 8; sz = 11 },
  @{ t = 'Справка выдана для предъявления в Консульский отдел Посольства Японии в Республике Казахстан.'; a = 'both'; ind = 0; fl = 1; sa = 14; sz = 11 },
  @{ t = 'CERTIFICATE OF EMPLOYMENT'; a = 'center'; ind = 0; fl = 0; sa = 8; sz = 13; b = $true },
  @{ t = 'This is to certify that Mr Karazhan Yerlik Narimanuly, born on 30 March 1978, IIN 780330300085, passport No. N18308567, is employed by the Central Asia Climate Foundation as Chief Executive Officer, Head of the Project Office for Central Asia on Climate Change and Green Energy, under Employment Contract No. 1-1 dated 18 February 2025, from 18 February 2025 to date. The Foundation is his main place of work.'; a = 'both'; ind = 0; fl = 1; sa = 8; sz = 11 },
  @{ t = 'During his official business trip to Tokyo (Japan) from 16 to 23 November 2026, his position, workplace and salary are retained. Upon return he will resume his duties.'; a = 'both'; ind = 0; fl = 1; sa = 8; sz = 11 },
  @{ t = 'Issued for submission to the Consular Section of the Embassy of Japan in the Republic of Kazakhstan.'; a = 'both'; ind = 0; fl = 1; sa = 16; sz = 11 },
  @{ t = 'Заместитель Главного исполнительного директора'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11; b = $true },
  @{ t = "по административным вопросам`tМ.Е. Темирбеков"; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11; b = $true },
  @{ t = 'Deputy Chief Executive Officer for Administrative Affairs'; a = 'left'; ind = 0; fl = 0; sa = 14; sz = 10 },
  @{ t = 'М.П.'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 }
)

function Esc([string]$s) { $s.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;') }

function RunsXml([string]$text, [double]$sz, [bool]$bold) {
  $fonts = '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/>'
  $szv = [int]($sz * 2)
  $b = if ($bold) { '<w:b/>' } else { '' }
  $sb = ''
  $sub = $text -split "`t"
  for ($k = 0; $k -lt $sub.Count; $k++) {
    if ($k -gt 0) { $sb += '<w:r><w:rPr>' + $fonts + $b + '<w:sz w:val="' + $szv + '"/><w:szCs w:val="' + $szv + '"/></w:rPr><w:tab/></w:r>' }
    if ($sub[$k] -ne '') {
      $sb += '<w:r><w:rPr>' + $fonts + $b + '<w:sz w:val="' + $szv + '"/><w:szCs w:val="' + $szv + '"/></w:rPr><w:t xml:space="preserve">' + (Esc $sub[$k]) + '</w:t></w:r>'
    }
  }
  return $sb
}

$body = ''
foreach ($s in $spec) {
  $sz = [double]$s.sz
  $ind = [int]([double]$s.ind * 567)
  $fl = if ($s.fl) { 709 } else { 0 }
  $after = [int]([double]$s.sa * 20)
  $szv = [int]($sz * 2)
  $tabs = if ($s.t -like "*`t*") { '<w:tabs><w:tab w:val="left" w:pos="6000"/></w:tabs>' } else { '' }
  $ppr = '<w:pPr>' + $tabs + '<w:spacing w:before="0" w:after="' + $after + '" w:line="276" w:lineRule="auto"/><w:ind w:left="' + $ind + '" w:right="0" w:firstLine="' + $fl + '"/><w:jc w:val="' + $s.a + '"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="' + $szv + '"/><w:szCs w:val="' + $szv + '"/></w:rPr></w:pPr>'
  $body += '<w:p>' + $ppr + (RunsXml $s.t $sz ([bool]$s.b)) + '</w:p>'
}

$zip = [IO.Compression.ZipFile]::Open($out, 'Update')
try {
  $entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
  $sr = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
  $xml = $sr.ReadToEnd(); $sr.Close()
  $i = $xml.LastIndexOf('<w:sectPr')
  $j = $xml.IndexOf('</w:body>')
  $sect = $xml.Substring($i, $j - $i)
  $head = $xml.Substring(0, $xml.IndexOf('<w:body>') + 8)
  $new = $head + $body + $sect + '</w:body></w:document>'
  $entry.Delete()
  $e2 = $zip.CreateEntry('word/document.xml')
  $sw = New-Object IO.StreamWriter($e2.Open(), (New-Object Text.UTF8Encoding $false))
  $sw.Write($new); $sw.Close()
}
finally { $zip.Dispose() }
"docx built, paragraphs=" + $spec.Count

# подпись Мурата (png) поверх строки с фамилией + свойство «Организация» = CACF + PDF
$hadWord = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
$sig = 'D:\Шаблоны\Подпись\Подпись Темирбеков М.Е. (png).png'
$pdf = [IO.Path]::ChangeExtension($out, 'pdf')
$w = New-Object -ComObject Word.Application
$w.Visible = $false; $w.DisplayAlerts = 0
try {
  $d = $w.Documents.Open($out)
  $r = $d.Content
  $fd = $r.Find; $fd.ClearFormatting()
  $ok = $fd.Execute('Темирбеков', $true, $false, $false, $false, $false, $true, 0, $false)
  $x = $r.Information(5); $y = $r.Information(6)
  "name at x=$x y=$y"
  $left = [single]($x + 4); $top = [single]($y - 34)
  $s = $d.Shapes.AddPicture($sig, $false, $true, $left, $top, [single]90, [single]40, $r)
  $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1
  $s.Left = $left; $s.Top = $top
  $s.WrapFormat.Type = 3
  $s.LockAspectRatio = $true
  $bt = [System.__ComObject]
  $it = $bt.InvokeMember('Item', 'GetProperty', $null, $d.BuiltInDocumentProperties, @('Company'))
  $bt.InvokeMember('Value', 'SetProperty', $null, $it, @('CACF'))
  "pages: " + $d.ComputeStatistics(2)
  $d.Save()
  $d.SaveAs($pdf, 17)
  $d.Close($false)
} finally { if (-not $hadWord) { $w.Quit() } }
"done: $pdf"
