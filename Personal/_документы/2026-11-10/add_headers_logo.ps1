$ErrorActionPreference='Stop'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
$f='D:\Кадры\График отпусков и остатки на 2026 (2026-11-10).xlsx'
$logo='D:\Шаблоны\Логотип\Логотип ЦАКФ (png).png'
$sign='D:\Шаблоны\Подпись\Подпись Темирбеков М.Е. (png).png'
$pdf='C:\Users\User\AppData\Local\Temp\claude\0fa2f3b7e932577a\0b7af0df-9350-47f3-8436-e9c3027c6b67\scratchpad\grafik_check.pdf'
$log='D:\Кадры\_исходники\add_headers_logo.log'; Remove-Item $log -ErrorAction SilentlyContinue
function L($m){ Add-Content $log ((Get-Date -Format 'HH:mm:ss')+' '+$m) -Encoding UTF8 }
function Pic($ws,$path,[double]$l,[double]$t,[double]$w,[double]$h){
  [void]$ws.Shapes.AddPicture($path,[int]0,[int]-1,[single]$l,[single]$t,[single]$w,[single]$h)
}
$xl=New-Object -ComObject Excel.Application
Setp $xl 'Visible' $false; Setp $xl 'DisplayAlerts' $false
try{
 $wb=$xl.Workbooks.Open($f)
 $w1=$wb.Worksheets.Item('График отпусков 2026')
 $w2=$wb.Worksheets.Item('Остатки на 10.11.2026')
 $w3=$wb.Worksheets.Item('Отпуска 2025')
 foreach($ws in @($w1,$w2,$w3)){ if($ws.Shapes.Count -gt 0){ L ('на листе '+$ws.Name+' уже есть рисунки: '+$ws.Shapes.Count); throw 'shapes exist' } }

 # ---------- Лист 1: График ----------
 [void]$w1.Activate(); Setp $xl.ActiveWindow 'FreezePanes' $false
 [void]$w1.Range('A1:A9').EntireRow.Insert()
 for($r=1;$r -le 5;$r++){ Setp $w1.Rows.Item($r) 'RowHeight' ([double]16) }
 Setp $w1.Rows.Item(6) 'RowHeight' ([double]8)
 Setp $w1.Rows.Item(7) 'RowHeight' ([double]24)
 Setp $w1.Rows.Item(8) 'RowHeight' ([double]18)
 Setp $w1.Rows.Item(9) 'RowHeight' ([double]8)
 $w1.Range('H1').Formula='УТВЕРЖДАЮ'
 $w1.Range('H2').Formula='Главный исполнительный директор,'
 $w1.Range('H3').Formula='Руководитель Проектного офиса ЦАКФ'
 $w1.Range('H4').Formula='_________________ Е.Н. Қаражан'
 $w1.Range('H5').Formula='«____» ______________ 2026 г.'
 Setp $w1.Range('H1').Font 'Bold' $true
 Setp $w1.Range('H1:H5') 'HorizontalAlignment' ([int]-4131)
 [void]$w1.Range('A7:H7').Merge(); [void]$w1.Range('A8:H8').Merge()
 $w1.Range('A7').Formula='График отпусков ЦАКФ на 2026 год'
 $w1.Range('A8').Formula='по состоянию на 10.11.2026'
 Setp $w1.Range('A7').Font 'Bold' $true; Setp $w1.Range('A7').Font 'Size' ([double]16)
 Setp $w1.Range('A8').Font 'Size' ([double]11)
 Setp $w1.Range('A7:A8') 'HorizontalAlignment' ([int]-4108)
 Setp $w1.Range('A7:A8') 'VerticalAlignment' ([int]-4108)
 $top=$w1.Range('A1').Top
 Pic $w1 $logo ($w1.Range('A1').Left+4) ($top+2) 150 60
 L 'лист 1: шапка'
 # подпись внизу
 $lastRow=10; while($w1.Cells.Item($lastRow+1,2).Text -ne ''){ $lastRow++ }
 $sr=$lastRow+3
 Setp $w1.Rows.Item($sr) 'RowHeight' ([double]40)
 [void]$w1.Range("A${sr}:E${sr}").Merge()
 $w1.Range("A$sr").Formula='Заместитель Главного исполнительного директора по административным вопросам'
 Setp $w1.Range("A$sr") 'HorizontalAlignment' ([int]-4131); Setp $w1.Range("A$sr") 'VerticalAlignment' ([int]-4107); Setp $w1.Range("A$sr") 'WrapText' $true
 [void]$w1.Range("F${sr}:G${sr}").Merge()
 $w1.Range("F$sr").Formula='_______________'
 Setp $w1.Range("F$sr") 'HorizontalAlignment' ([int]-4108); Setp $w1.Range("F$sr") 'VerticalAlignment' ([int]-4107)
 $w1.Range("H$sr").Formula='М.Е. Темирбеков'
 Setp $w1.Range("H$sr") 'HorizontalAlignment' ([int]-4131); Setp $w1.Range("H$sr") 'VerticalAlignment' ([int]-4107)
 $fl=$w1.Range("F$sr").Left; $ft=$w1.Range("F$sr").Top
 Pic $w1 $sign ($fl+6) ($ft+1) 80 36
 L "лист 1: подпись в строке $sr (последняя строка таблицы $lastRow); A8='"+$w1.Range('A8').Formula+"' A7='"+$w1.Range('A7').Formula+"'"
 Setp $w1.PageSetup 'PrintArea' ('$A$1:$H$'+$sr)
 Setp $w1.PageSetup 'Orientation' ([int]2)

 # ---------- Лист 2: Остатки ----------
 [void]$w2.Range('A1:A4').EntireRow.Insert()
 Setp $w2.Rows.Item(1) 'RowHeight' ([double]22); Setp $w2.Rows.Item(2) 'RowHeight' ([double]22); Setp $w2.Rows.Item(3) 'RowHeight' ([double]18); Setp $w2.Rows.Item(4) 'RowHeight' ([double]8)
 [void]$w2.Range('C1:L2').Merge(); [void]$w2.Range('C3:L3').Merge()
 $w2.Range('C1').Formula='Остатки отпусков ЦАКФ на 2026 год'
 $w2.Range('C3').Formula='по состоянию на 10.11.2026, с учётом переноса на 2027 год'
 Setp $w2.Range('C1').Font 'Bold' $true; Setp $w2.Range('C1').Font 'Size' ([double]16)
 Setp $w2.Range('C1:C3') 'HorizontalAlignment' ([int]-4108); Setp $w2.Range('C1:C3') 'VerticalAlignment' ([int]-4108)
 Pic $w2 $logo ($w2.Range('A1').Left+4) ($w2.Range('A1').Top+2) 100 40
 L 'лист 2: шапка'

 # ---------- Лист 3: Отпуска 2025 ----------
 [void]$w3.Range('A1:A4').EntireRow.Insert()
 Setp $w3.Rows.Item(1) 'RowHeight' ([double]22); Setp $w3.Rows.Item(2) 'RowHeight' ([double]22); Setp $w3.Rows.Item(3) 'RowHeight' ([double]18); Setp $w3.Rows.Item(4) 'RowHeight' ([double]8)
 [void]$w3.Range('C1:H2').Merge(); [void]$w3.Range('C3:H3').Merge()
 $w3.Range('C1').Formula='Отпуска ЦАКФ за 2025 год'
 $w3.Range('C3').Formula='для расчёта переноса неиспользованных дней в 2026 год'
 Setp $w3.Range('C1').Font 'Bold' $true; Setp $w3.Range('C1').Font 'Size' ([double]16)
 Setp $w3.Range('C1:C3') 'HorizontalAlignment' ([int]-4108); Setp $w3.Range('C1:C3') 'VerticalAlignment' ([int]-4108)
 Pic $w3 $logo ($w3.Range('A1').Left+4) ($w3.Range('A1').Top+2) 100 40
 L 'лист 3: шапка'

 $xl.Calculate()
 [void]$w1.Activate()
 # контроль значений
 $out=@()
 foreach($ws in @($w1,$w2)){ $ur=$ws.UsedRange; for($rr=1;$rr -le [Math]::Min(30,$ur.Rows.Count+0);$rr++){ $l=@(); for($c=1;$c -le 16;$c++){$l+=$ws.Cells.Item($rr,$c).Text}; $out+=($ws.Name.Substring(0,4)+' '+$rr+' | '+($l -join ' ; ')) } }
 $out | Set-Content 'D:\Кадры\_исходники\add_headers_logo_check.txt' -Encoding UTF8
 $wb.Save(); L 'сохранено'
 # PDF для визуальной проверки (первый лист)
 [void]$w1.ExportAsFixedFormat(0, $pdf)
 L 'pdf'
 $wb.Close($false)
}catch{ L ('ОШИБКА: '+$_.Exception.Message+' @'+$_.InvocationInfo.ScriptLineNumber) }
finally{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
