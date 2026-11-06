$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tpl = 'C:\Users\User\Documents\письма\Предложение CACF.docx'
$out = 'D:\Кадры\Справки\Письмо поддержки - Inkar Karazhan (виза в Японию, проект 2026-11-06).docx'
Copy-Item -LiteralPath $tpl -Destination $out -Force

# a - выравнивание; ind - отступ слева, см; fl - красная строка (1 = да); sa - интервал после, пт; sz - кегль; b - жирный
$spec = @(
  @{ t = 'Ref. No. 02/398 dated 06 November 2026'; a = 'left'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'The Consular Section, Consulate-General of Japan in Busan, Republic of Korea'; a = 'left'; ind = 8.5; fl = 0; sa = 10; sz = 11; b = $true },
  @{ t = 'Subject: Letter of Support and Financial Guarantee for Ms Inkar Karazhan - short-stay visa to Japan, 17-22 November 2026'; a = 'both'; ind = 0; fl = 0; sa = 8; sz = 11; b = $true },
  @{ t = 'Dear Sir / Madam,'; a = 'left'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'I, Yerlik Karazhan, a citizen of the Republic of Kazakhstan, holder of passport No. N18308567, Chief Executive Officer of the Central Asia Climate Foundation (Astana, Republic of Kazakhstan), respectfully submit this letter in support of the short-stay visa application of my daughter, Ms Inkar Karazhan.'; a = 'both'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'Applicant''s particulars'; a = 'left'; ind = 0; fl = 0; sa = 3; sz = 11; b = $true },
  @{ t = 'Full name: KARAZHAN Inkar'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Date of birth: 15 January 2008'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Nationality: Republic of Kazakhstan'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Passport No.: N17344632, valid until 13 November 2034'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Relationship to me: daughter'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Current status: student, resident in the Republic of Korea'; a = 'left'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'Purpose and circumstances of the visit'; a = 'left'; ind = 0; fl = 0; sa = 3; sz = 11; b = $true },
  @{ t = 'I will be in Tokyo from 17 to 22 November 2026 at the invitation of the Japan Bank for International Cooperation (JBIC), to take part in the “Raisina Tokyo” Roundtable and the Second Meeting of the High-Level Empowered Board on Green Transition, held on 18 and 19 November 2026. My own visa application is lodged with the Embassy of Japan in the Republic of Kazakhstan, supported by a Letter of Invitation and a Letter of Guarantee issued by JBIC.'; a = 'both'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'My daughter resides in the Republic of Korea for her studies, and we see each other rarely. She will therefore join me in Tokyo for the whole of my stay, sharing my accommodation throughout, and will return to the Republic of Korea immediately afterwards to resume her studies.'; a = 'both'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'Her travel is already booked and paid for by me:'; a = 'left'; ind = 0; fl = 0; sa = 2; sz = 11 },
  @{ t = 'Arrival: 17 November 2026 - Korean Air KE 2131, Busan (PUS) 16:20 → Tokyo Narita (NRT) 18:30'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Departure: 22 November 2026 - Korean Air KE 2132, Tokyo Narita (NRT) 19:40 → Busan (PUS) 22:00'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Order No. AA9NXGN9 · PNR XYHDQT · Ticket No. 180-3985813294'; a = 'left'; ind = 0; fl = 0; sa = 4; sz = 11 },
  @{ t = 'Our flights are on the same days: I arrive at Narita on 17 November at 17:55 and depart on 22 November at 19:35, so we travel to and from Japan together.'; a = 'both'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'Accommodation'; a = 'left'; ind = 0; fl = 0; sa = 3; sz = 11; b = $true },
  @{ t = '17-20 November 2026 - my accommodation, arranged by the organisers; my daughter stays with me in the same room: ANA InterContinental Tokyo, 1-12-33 Akasaka, Minato-ku, Tokyo 107-0052, Japan, Tel. +81 3-3505-1111.'; a = 'left'; ind = 0; fl = 0; sa = 3; sz = 11 },
  @{ t = '20-22 November 2026 - booked by me for two guests, my daughter and myself: APA Hotel Higashi Shinjuku Kabukicho Tower, 2-31-12 Kabukicho, Shinjuku-ku, Tokyo 160-0021, Japan, Tel. +81 3 5291 9311; Booking.com confirmation No. 5035.657.373; one double room, two adults, check-in 20 November, check-out 22 November.'; a = 'left'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'Financial guarantee'; a = 'left'; ind = 0; fl = 0; sa = 3; sz = 11; b = $true },
  @{ t = 'I hereby confirm that I assume full financial responsibility for my daughter''s travel to and stay in Japan. Her return air ticket has already been purchased and paid for by me. All further costs - accommodation, meals, local transport, medical insurance and any incidental expenses - are likewise borne entirely by me. She will not seek employment in Japan and will not become a charge on public funds.'; a = 'both'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'I further confirm that she will leave Japan on 22 November 2026, before the expiry of her authorised period of stay, and will return to the Republic of Korea to continue her education.'; a = 'both'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'Enclosures'; a = 'left'; ind = 0; fl = 0; sa = 3; sz = 11; b = $true },
  @{ t = '1. Copy of my passport (bio-data page)'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '2. Copy of my Letter of Invitation and Letter of Guarantee issued by JBIC'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '3. My flight itinerary to and from Japan (16-23 November 2026)'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '4. My daughter''s e-ticket, Busan-Tokyo-Busan, 17-22 November 2026 (Order No. AA9NXGN9)'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '5. Confirmation of my accommodation at ANA InterContinental Tokyo, 17-20 November 2026'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '6. Hotel reservation for two guests, 20-22 November 2026 (Booking.com No. 5035.657.373)'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '7. Document evidencing our relationship (birth certificate)'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '8. Certificate of employment, Ref. No. 02/397 dated 6 November 2026'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = '9. Bank statement / certificate of funds'; a = 'left'; ind = 0; fl = 0; sa = 6; sz = 11 },
  @{ t = 'I remain at your disposal should any further information or documentation be required, and I thank you for your kind consideration of my daughter''s application.'; a = 'both'; ind = 0; fl = 0; sa = 8; sz = 11 },
  @{ t = 'Yours faithfully,'; a = 'left'; ind = 0; fl = 0; sa = 24; sz = 11 },
  @{ t = 'Yerlik Karazhan'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11; b = $true },
  @{ t = 'Chief Executive Officer'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 },
  @{ t = 'Central Asia Climate Foundation'; a = 'left'; ind = 0; fl = 0; sa = 0; sz = 11 }
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
  $bt = [System.__ComObject]
  $it = $bt.InvokeMember('Item', 'GetProperty', $null, $d.BuiltInDocumentProperties, @('Company'))
  $bt.InvokeMember('Value', 'SetProperty', $null, $it, @('CACF'))
  "pages: " + $d.ComputeStatistics(2)
  $d.Save()
  $d.SaveAs($pdf, 17)
  $d.Close($false)
} finally { if (-not $hadWord) { $w.Quit() } }
"done: $pdf"
