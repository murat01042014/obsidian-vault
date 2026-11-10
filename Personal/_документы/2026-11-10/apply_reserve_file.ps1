$ErrorActionPreference='Stop'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
$f='D:\Кадры\График отпусков и остатки на 2026 (2026-11-10).xlsx'
$logo='D:\Шаблоны\Логотип\Логотип ЦАКФ (png).png'
$pdf='C:\Users\User\AppData\Local\Temp\claude\0fa2f3b7e932577a\0b7af0df-9350-47f3-8436-e9c3027c6b67\scratchpad\sverka_check.pdf'
$log='D:\Кадры\_исходники\apply_reserve_file.log'; Remove-Item $log -ErrorAction SilentlyContinue
function L($m){ Add-Content $log ((Get-Date -Format 'HH:mm:ss')+' '+$m) -Encoding UTF8 }
Add-Type -Namespace W -Name U -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$xl=New-Object -ComObject Excel.Application
$p=[uint32]0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$p); L "PID $p"
Setp $xl 'Visible' $false; Setp $xl 'DisplayAlerts' $false; Setp $xl 'AutomationSecurity' ([int]3)
try{
 $wb=$xl.Workbooks.Open($f,[int]0,$false)
 $w2=$wb.Worksheets.Item('Остатки на 10.11.2026'); $w3=$wb.Worksheets.Item('Отпуска 2025')
 if($wb.Worksheets.Count -ne 3){ throw 'неожиданное число листов' }

 # ---- 2025: данные по файлу резерва ----
 $w3.Range('D5').Formula='Мес. начисления с 01.03.2025 (по файлу резерва)'
 $w3.Range('D6').Formula='10'; $w3.Range('D7').Formula='10'; $w3.Range('D8').Formula='9.9'
 $w3.Range('F12').Formula='9'
 $w3.Range('H5').Formula='Остаток на 31.12.2025 = перенос в 2026 (по файлу резерва)'
 $w3.Range('H6').Formula='=G6'; $w3.Range('H7').Formula='=G7'; $w3.Range('H8').Formula='=G8'
 $w3.Range('I5').Formula='Сверх лимита 5 дней'
 $w3.Range('I6').Formula='=MAX(0,G6-5)'; $w3.Range('I7').Formula='=MAX(0,G7-5)'; $w3.Range('I8').Formula='=MAX(0,G8-5)'
 Setp $w3.Range('D6:D8') 'NumberFormat' '0.0'
 Setp $w3.Range('E6:I8') 'NumberFormat' '0.00'
 $w3.Range('H5').Copy($w3.Range('I5')) | Out-Null
 $w3.Range('H6:H8').Copy($w3.Range('I6:I8')) | Out-Null
 $w3.Range('I6').Formula='=MAX(0,G6-5)'; $w3.Range('I7').Formula='=MAX(0,G7-5)'; $w3.Range('I8').Formula='=MAX(0,G8-5)'
 $w3.Range('I5').Formula='Сверх лимита 5 дней'
 Setp $w3.Range('I6:I8') 'NumberFormat' '0.00'
 Setp $w3.Columns.Item(8) 'ColumnWidth' ([double]22); Setp $w3.Columns.Item(9) 'ColumnWidth' ([double]16)
 Setp $w3.Rows.Item(5) 'RowHeight' ([double]62)
 L 'лист 2025 исправлен'

 # ---- Остатки: подпись столбца ----
 $w2.Range('F5').Formula='Перенос с 2025 в 2026 (по файлу резерва)'
 $xl.Calculate()
 L 'остатки пересчитаны'

 # ---- Новый лист: сверка ----
 $ws=$wb.Worksheets.Add([Type]::Missing,$wb.Worksheets.Item($wb.Worksheets.Count)); $ws.Name='Сверка с файлом резерва'
 Setp $ws.Rows.Item(1) 'RowHeight' ([double]22); Setp $ws.Rows.Item(2) 'RowHeight' ([double]22); Setp $ws.Rows.Item(3) 'RowHeight' ([double]18); Setp $ws.Rows.Item(4) 'RowHeight' ([double]8)
 [void]$ws.Range('C1:G2').Merge()
 $ws.Range('C1').Formula='Сверка с файлом «Расчет резерва на 31.12.2025 года - коррект»'
 $ws.Range('C3').Formula='основной источник: файл резерва; расхождения с графиком и приказами 2026 года'
 Setp $ws.Range('C1').Font 'Bold' $true; Setp $ws.Range('C1').Font 'Size' ([double]14)
 Setp $ws.Range('C1') 'HorizontalAlignment' ([int]-4108); Setp $ws.Range('C1') 'VerticalAlignment' ([int]-4108); Setp $ws.Range('C1') 'WrapText' $true
 Setp $ws.Range('C3:G3') 'HorizontalAlignment' ([int]7)
 [void]$ws.Shapes.AddPicture($logo,[int]0,[int]-1,[single]($ws.Range('A1').Left+4),[single]($ws.Range('A1').Top+2),[single]100,[single]40)
 $h=@('№','Сотрудник','Показатель','Файл резерва','График и приказы 2026','Разница','Где в файле резерва / в документах','Принято в расчёте')
 for($i=0;$i -lt 8;$i++){ $ws.Cells.Item(5,$i+1).Formula=$h[$i] }
 $rows=@(
  @('Қаражан Е.Н.','Остаток на 31.12.2025, раб. дн.','4,67','5,0 (был лимит 5)','-0,33','лист «Расчет для резерва на 31.12.2025», K21','по файлу: 4,67'),
  @('Темирбеков М.Е.','Остаток на 31.12.2025, раб. дн.','4,67','5,0 (был лимит 5)','-0,33','лист «Расчет для резерва на 31.12.2025», K22','по файлу: 4,67'),
  @('Оразбаева М.К.','Остаток на 31.12.2025, раб. дн.','6,5','5,0 (был лимит 5)','+1,5','K23; превышает лимит остатка 5 дней','по файлу: 6,5; сверх лимита 1,5'),
  @('Қаражан Е.Н.','Использовано в 2025 (приказ 6-4, 25.08-05.09.2025)','9 (в сумме 12 с приказом 6-18)','10','-1','F21:G21; 01.09.2025 нерабочий','по файлу: 9'),
  @('Қаражан Е.Н.','Начислено за 2025','16,67 (10 мес. с 01.03.2025)','18,33 (11 мес. с 18.02.2025)','-1,67','E21; февраль 2025 в резерв не включён (лист «Пересчет февраль 2025»)','по файлу'),
  @('Темирбеков М.Е.','Начислено за 2025','16,67 (10 мес. с 01.03.2025)','18,33 (11 мес. с 24.02.2025)','-1,67','E22; февраль 2025 в резерв не включён','по файлу'),
  @('Оразбаева М.К.','Начислено за 2025','16,5 (9,9 мес.)','16,67 (10 мес.)','-0,17','E23','по файлу'),
  @('Ажигалиева Д.Н.','Второй отпуск 2026, даты (7 раб. дн.)','13.05-21.05.2026','18.05-27.05.2026 (приказ 03-7/7, табель)','даты','лист «Движение за 2026», D5','в графике даты приказа; дни 14 совпадают'),
  @('Салыкова Д.О.','Отпуск 2026, окончание (приказ 03-7/3)','12.05-26.05.2026, 11 дн.','12.05-27.05.2026, 11 дн. по приказу, 12 по табелю','1 день','лист «Движение за 2026», D6','11 дн. как в файле и приказе'),
  @('Кононец А.В.','Отпуск 2026, окончание (приказ 03-7/18)','24.08-07.09.2026, 10 дн.','24.08-04.09.2026, 10 дн. (приказ, табель)','даты','лист «Резерв на 01.11.2026», R16','10 дн.'),
  @('Кононец А.В.','Начислено на конец октября 2026','14,47 (8,68 мес.)','16,67 (10 мес. на 10.11.2026)','-2,2','лист «Резерв на 01.11.2026», E16:F16','по датам графика, см. «Остатки»'),
  @('Альмаганбетов А.С.','Дата начала работы','08.01.2026','05.01.2026 (соглашение №8, приказ 03-6/3)','3 дня','лист «Движение за 2026», D7','расхождение не снято'),
  @('Ажигалиева Д.Н.','Использовано в 2026: приказы и табель','14 (приказы 03-7/1, 03-7/7)','15 по табелю','1 день','«Движение за 2026», E5','по приказам: 14'),
  @('Салыкова Д.О.','Использовано в 2026: приказ и табель','11 (приказ 03-7/3)','12 по табелю','1 день','«Движение за 2026», E6','по приказу: 11')
 )
 $r=6; $n=1
 foreach($e in $rows){
   $ws.Cells.Item($r,1).Formula=[string]$n
   for($k=0;$k -lt 7;$k++){ $c=$ws.Cells.Item($r,$k+2); Setp $c 'NumberFormat' '@'; $c.Formula=$e[$k] }
   $r++; $n++
 }
 # неиспользовано на 01.11.2026 / доступно на 10.11.2026 (формулы на листе «Остатки»)
 $names=@('Қаражан Е.Н.','Темирбеков М.Е.','Оразбаева М.К.','Кононец А.В.')
 $fileVals=@('4,33','6,33','7,17','4,47')
 for($i=0;$i -lt 4;$i++){
   $ws.Cells.Item($r,1).Formula=[string]$n
   $ws.Cells.Item($r,2).Formula=$names[$i]
   $ws.Cells.Item($r,3).Formula='Неиспользовано: файл на 01.11.2026, мой расчёт на 10.11.2026'
   $c=$ws.Cells.Item($r,4); Setp $c 'NumberFormat' '@'; $c.Formula=$fileVals[$i]
   $ws.Cells.Item($r,5).Formula=("='Остатки на 10.11.2026'!M"+(6+$i))
   $ws.Cells.Item($r,6).Formula=("=E$r-"+($fileVals[$i] -replace ',','.'))
   $ws.Cells.Item($r,7).Formula='лист «Резерв на 01.11.2026», H13:H16'
   $ws.Cells.Item($r,8).Formula='в файле накопление 20/12 помесячно (10 мес. на 01.11), у меня 20 дн. в год; на 31.12 итог одинаковый'
   $r++; $n++
 }
 $last=$r-1
 Setp $ws.Range("E$($last-3):F$last") 'NumberFormat' '0.00'
 $hr=$ws.Range('A5:H5'); Setp $hr.Font 'Bold' $true; Setp $hr.Interior 'Color' ([double]15921906); Setp $hr 'WrapText' $true; Setp $hr 'HorizontalAlignment' ([int]-4108); Setp $hr 'VerticalAlignment' ([int]-4108)
 Setp $ws.Range("A5:H$last").Borders 'LineStyle' ([int]1)
 Setp $ws.Range("A6:H$last") 'WrapText' $true; Setp $ws.Range("A6:H$last") 'VerticalAlignment' ([int]-4108)
 Setp $ws.Range("A6:A$last") 'HorizontalAlignment' ([int]-4108); Setp $ws.Range("D6:F$last") 'HorizontalAlignment' ([int]-4108)
 $w=@(5,20,38,26,30,10,44,40); for($i=0;$i -lt 8;$i++){ Setp $ws.Columns.Item($i+1) 'ColumnWidth' ([double]$w[$i]) }
 [void]$ws.Range("A6:H$last").Rows.AutoFit()
 Setp $ws.Rows.Item(5) 'RowHeight' ([double]32)
 Setp $ws.PageSetup 'Orientation' ([int]2); Setp $ws.PageSetup 'Zoom' $false; Setp $ws.PageSetup 'FitToPagesWide' ([int]1); Setp $ws.PageSetup 'FitToPagesTall' $false
 $xl.Calculate()
 L 'лист сверки создан'

 # контроль
 $out=@()
 foreach($sn in @('Остатки на 10.11.2026','Отпуска 2025','Сверка с файлом резерва')){
   $s=$wb.Worksheets.Item($sn); $rowsN=[Math]::Min($s.UsedRange.Rows.Count,30)
   for($rr=5;$rr -le $rowsN;$rr++){ $l=@(); for($cc=1;$cc -le 16;$cc++){ $l+=$s.Cells.Item($rr,$cc).Text }; if(($l -join '') -ne ''){ $out+=($sn.Substring(0,5)+' '+$rr+' | '+($l -join ' ; ')) } }
 }
 $out | Set-Content 'D:\Кадры\_исходники\apply_reserve_file_check.txt' -Encoding UTF8
 [void]$ws.ExportAsFixedFormat(0,$pdf)
 $wb.Save(); L 'сохранено'
 $wb.Close($false)
}catch{ L ('ОШИБКА: '+$_.Exception.Message+' @'+$_.InvocationInfo.ScriptLineNumber) }
finally{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
