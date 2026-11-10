$ErrorActionPreference='Stop'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
$f='D:\Шаблоны\Отпуска\Шаблон - график отпусков и остатки.xlsx'
$logo='D:\Шаблоны\Логотип\Логотип ЦАКФ (png).png'
$pdf='C:\Users\User\AppData\Local\Temp\claude\0fa2f3b7e932577a\0b7af0df-9350-47f3-8436-e9c3027c6b67\scratchpad\shablon_check.pdf'
$log='D:\Кадры\_исходники\template_headers.log'; Remove-Item $log -ErrorAction SilentlyContinue
function L($m){ Add-Content $log ((Get-Date -Format 'HH:mm:ss')+' '+$m) -Encoding UTF8 }
function Pic($ws,[double]$l,[double]$t,[double]$w,[double]$h){ [void]$ws.Shapes.AddPicture($logo,[int]0,[int]-1,[single]$l,[single]$t,[single]$w,[single]$h) }
function DT($c) { return ('TEXT(DAY(' + $c + '),"00")&"."&TEXT(MONTH(' + $c + '),"00")&"."&YEAR(' + $c + ')') }
$xl=New-Object -ComObject Excel.Application
Setp $xl 'Visible' $false; Setp $xl 'DisplayAlerts' $false
try{
 $wb=$xl.Workbooks.Open($f)
 $wp=$wb.Worksheets.Item('Параметры'); $wo=$wb.Worksheets.Item('Остатки'); $wk=$wb.Worksheets.Item('Календарь')
 foreach($ws in @($wp,$wo,$wk)){ if($ws.Shapes.Count -gt 0){ throw "на листе $($ws.Name) уже есть рисунки" } }

 # ---- Параметры: подписанты ----
 $lab=@('Утверждает: должность (строка 1)','Утверждает: должность (строка 2)','Утверждает: ФИО','Составил: должность','Составил: ФИО')
 $val=@('Главный исполнительный директор,','Руководитель Проектного офиса ЦАКФ','Е.Н. Қаражан','Заместитель Главного исполнительного директора по административным вопросам','М.Е. Темирбеков')
 for($i=0;$i -lt 5;$i++){ $wp.Cells.Item(8+$i,1).Formula=$lab[$i]; $c=$wp.Cells.Item(8+$i,2); Setp $c 'NumberFormat' '@'; $c.Formula=$val[$i] }
 Setp $wp.Columns.Item(2) 'ColumnWidth' ([double]50)
 Setp $wp.Range('A8:B12').Borders 'LineStyle' ([int]1)
 Setp $wp.Range('B8:B12').Interior 'Color' ([double]13431551)
 Setp $wp.Range('B2:B12') 'HorizontalAlignment' ([int]-4131)
 Setp $wp.Range('B8:B12') 'WrapText' $true
 L 'Параметры'

 # ---- Остатки и Календарь: шапка ----
 $subO = '="по состоянию на "&' + (DT 'Параметры!$B$3') + '&", с учётом переноса на следующий год"'
 $subK = '="по состоянию на "&' + (DT 'Параметры!$B$3')
 $cfg=@(
  @($wo,'C1:L2','C1','C3:L3','C3','="Остатки отпусков ЦАКФ на "&Параметры!$B$2&" год"',$subO,100,40),
  @($wk,'C1:L2','C1','C3:L3','C3','="Дни отпуска ЦАКФ по месяцам на "&Параметры!$B$2&" год"',$subK,100,40)
 )
 foreach($c in $cfg){
  $ws=$c[0]; [void]$ws.Activate(); Setp $xl.ActiveWindow 'FreezePanes' $false
  [void]$ws.Range('A1:A4').EntireRow.Insert()
  Setp $ws.Rows.Item(1) 'RowHeight' ([double]22); Setp $ws.Rows.Item(2) 'RowHeight' ([double]22); Setp $ws.Rows.Item(3) 'RowHeight' ([double]18); Setp $ws.Rows.Item(4) 'RowHeight' ([double]8)
  [void]$ws.Range($c[1]).Merge()
  $ws.Range($c[2]).Formula=$c[5]; $ws.Range($c[4]).Formula=$c[6]
  Setp $ws.Range($c[2]).Font 'Bold' $true; Setp $ws.Range($c[2]).Font 'Size' ([double]16)
  Setp $ws.Range($c[2]) 'HorizontalAlignment' ([int]-4108); Setp $ws.Range($c[2]) 'VerticalAlignment' ([int]-4108)
  Setp $ws.Range($c[3]) 'HorizontalAlignment' ([int]7)
  Setp $ws.Range($c[4]).Font 'Size' ([double]11)
  Pic $ws ($ws.Range('A1').Left+4) ($ws.Range('A1').Top+2) $c[7] $c[8]
  Setp $xl.ActiveWindow 'SplitRow' ([int]5); Setp $xl.ActiveWindow 'FreezePanes' $true
  L ('шапка: '+$ws.Name)
 }

 # ---- Новый лист: График для печати ----
 $wg=$wb.Worksheets.Add([Type]::Missing,$wb.Worksheets.Item($wb.Worksheets.Count)); $wg.Name='График для печати'
 $h=@('№','Сотрудник','Вид отпуска','С','По','Рабочих дней','Приказ №','Примечание')
 for($i=0;$i -lt 8;$i++){ $wg.Cells.Item(10,$i+1).Formula=$h[$i] }
 $wg.Range('A11:A40').Formula='=IF(Отпуска!B2="","",ROW()-10)'
 $wg.Range('B11:B40').Formula='=IF(Отпуска!B2="","",Отпуска!B2)'
 $wg.Range('C11:C40').Formula='=IF(Отпуска!B2="","",Отпуска!C2)'
 $wg.Range('D11:D40').Formula='=IF(Отпуска!B2="","",Отпуска!D2)'
 $wg.Range('E11:E40').Formula='=IF(Отпуска!B2="","",Отпуска!E2)'
 $wg.Range('F11:F40').Formula='=IF(Отпуска!B2="","",Отпуска!F2)'
 $wg.Range('G11:G40').Formula='=IF(Отпуска!B2="","",IF(Отпуска!G2="","",Отпуска!G2))'
 $wg.Range('H11:H40').Formula='=IF(Отпуска!B2="","",IF(Отпуска!J2="","",Отпуска!J2))'
 Setp $wg.Range('D11:E40') 'NumberFormat' 'dd.mm.yyyy'
 $w=@(5,24,24,12,12,12,12,48); for($i=0;$i -lt 8;$i++){ Setp $wg.Columns.Item($i+1) 'ColumnWidth' ([double]$w[$i]) }
 for($r=1;$r -le 5;$r++){ Setp $wg.Rows.Item($r) 'RowHeight' ([double]16) }
 Setp $wg.Rows.Item(6) 'RowHeight' ([double]8); Setp $wg.Rows.Item(7) 'RowHeight' ([double]24); Setp $wg.Rows.Item(8) 'RowHeight' ([double]18); Setp $wg.Rows.Item(9) 'RowHeight' ([double]8)
 $wg.Range('H1').Formula='УТВЕРЖДАЮ'
 $wg.Range('H2').Formula='=Параметры!B8'
 $wg.Range('H3').Formula='=Параметры!B9'
 $wg.Range('H4').Formula='="_________________ "&Параметры!B10'
 $wg.Range('H5').Formula='="«____» ______________ "&Параметры!B2&" г."'
 Setp $wg.Range('H1').Font 'Bold' $true
 Setp $wg.Range('H1:H5') 'HorizontalAlignment' ([int]-4131)
 [void]$wg.Range('A7:H7').Merge()
 $wg.Range('A7').Formula='="График отпусков ЦАКФ на "&Параметры!$B$2&" год"'
 $wg.Range('A8').Formula=('="по состоянию на "&' + (DT 'Параметры!$B$3'))
 Setp $wg.Range('A7').Font 'Bold' $true; Setp $wg.Range('A7').Font 'Size' ([double]16)
 Setp $wg.Range('A8').Font 'Size' ([double]11)
 Setp $wg.Range('A7') 'HorizontalAlignment' ([int]-4108); Setp $wg.Range('A8:H8') 'HorizontalAlignment' ([int]7)
 Setp $wg.Range('A7:A8') 'VerticalAlignment' ([int]-4108)
 $hr=$wg.Range('A10:H10'); Setp $hr.Font 'Bold' $true; Setp $hr.Interior 'Color' ([double]15921906); Setp $hr 'WrapText' $true; Setp $hr 'HorizontalAlignment' ([int]-4108); Setp $hr 'VerticalAlignment' ([int]-4108)
 Setp $wg.Range('A10:H40').Borders 'LineStyle' ([int]1)
 Setp $wg.Range('A11:A40') 'HorizontalAlignment' ([int]-4108); Setp $wg.Range('D11:G40') 'HorizontalAlignment' ([int]-4108)
 Setp $wg.Rows.Item(10) 'RowHeight' ([double]30)
 Pic $wg ($wg.Range('A1').Left+4) ($wg.Range('A1').Top+2) 150 60
 # подпись внизу: должность и место для подписи
 Setp $wg.Rows.Item(43) 'RowHeight' ([double]40)
 [void]$wg.Range('A43:E43').Merge()
 $wg.Range('A43').Formula='=Параметры!B11'
 Setp $wg.Range('A43') 'HorizontalAlignment' ([int]-4131); Setp $wg.Range('A43') 'VerticalAlignment' ([int]-4107); Setp $wg.Range('A43') 'WrapText' $true
 [void]$wg.Range('F43:G43').Merge()
 $wg.Range('F43').Formula='_______________'
 Setp $wg.Range('F43') 'HorizontalAlignment' ([int]-4108); Setp $wg.Range('F43') 'VerticalAlignment' ([int]-4107)
 $wg.Range('H43').Formula='=Параметры!B12'
 Setp $wg.Range('H43') 'HorizontalAlignment' ([int]-4131); Setp $wg.Range('H43') 'VerticalAlignment' ([int]-4107)
 Setp $wg.PageSetup 'PrintArea' '$A$1:$H$43'
 Setp $wg.PageSetup 'Orientation' ([int]2); Setp $wg.PageSetup 'Zoom' $false; Setp $wg.PageSetup 'FitToPagesWide' ([int]1); Setp $wg.PageSetup 'FitToPagesTall' $false
 L 'лист для печати'
 foreach($ws in @($wo,$wk)){ Setp $ws.PageSetup 'Orientation' ([int]2) }
 [void]$wp.Activate()
 $xl.Calculate()
 $wb.Save(); L 'шаблон сохранён'

 # ---- проверка на тестовых данных (не сохраняется) ----
 $we=$wb.Worksheets.Item('Сотрудники'); $wv=$wb.Worksheets.Item('Отпуска')
 $emp=@(@('Қаражан Е.Н.','=DATE(2025,2,18)',5),@('Темирбеков М.Е.','=DATE(2025,2,24)',5),@('Оразбаева М.К.','=DATE(2025,3,3)',5),@('Кононец А.В.','=DATE(2026,2,10)',0))
 $r=2; foreach($e in $emp){ $we.Cells.Item($r,1).Formula=$e[0]; $we.Cells.Item($r,3).Formula=$e[1]; $we.Cells.Item($r,5).Formula=[string]$e[2]; $r++ }
 $lv=@(@('Қаражан Е.Н.',2026,8,12,2026,8,19,'03-7/17'),@('Темирбеков М.Е.',2026,9,7,2026,9,18,'03-7/21'),@('Оразбаева М.К.',2026,9,16,2026,9,29,'03-7/22'),@('Кононец А.В.',2026,8,24,2026,9,4,'03-7/18'))
 $r=2; foreach($e in $lv){ $wv.Cells.Item($r,2).Formula=$e[0]; $wv.Cells.Item($r,3).Formula='Очередной (О)'; $wv.Cells.Item($r,4).Formula=("=DATE("+$e[1]+","+$e[2]+","+$e[3]+")"); $wv.Cells.Item($r,5).Formula=("=DATE("+$e[4]+","+$e[5]+","+$e[6]+")"); $c=$wv.Cells.Item($r,7); Setp $c 'NumberFormat' '@'; $c.Formula=('№'+$e[7]); $r++ }
 $wp.Range('B3').Formula='=DATE(2026,11,10)'
 $xl.Calculate()
 $out=@(); foreach($ws in @($wg,$wo)){ for($rr=1;$rr -le 16;$rr++){ $l=@(); for($cc=1;$cc -le 9;$cc++){$l+=$ws.Cells.Item($rr,$cc).Text}; $out+=($ws.Name.Substring(0,5)+' '+$rr+' | '+($l -join ' ; ')) } }
 $out | Set-Content 'D:\Кадры\_исходники\template_headers_check.txt' -Encoding UTF8
 [void]$wg.ExportAsFixedFormat(0,$pdf)
 $wb.Close($false); L 'проверка выполнена, тестовые данные не сохранены'
}catch{ L ('ОШИБКА: '+$_.Exception.Message+' @'+$_.InvocationInfo.ScriptLineNumber) }
finally{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
