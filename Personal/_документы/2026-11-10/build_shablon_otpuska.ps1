$ErrorActionPreference = 'Stop'
$dir = 'D:\Шаблоны\Отпуска'
$out = Join-Path $dir 'Шаблон - график отпусков и остатки.xlsx'
$log = 'D:\Кадры\_исходники\build_shablon_otpuska.log'
Remove-Item $log -ErrorAction SilentlyContinue
function L($m) { Add-Content -Path $log -Value ((Get-Date -Format 'HH:mm:ss') + ' ' + $m) -Encoding UTF8 }
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
if (Test-Path $out) { L 'файл уже есть, перезапись запрещена'; throw 'exists' }
$xl = New-Object -ComObject Excel.Application
Setp $xl 'Visible' $false
Setp $xl 'DisplayAlerts' $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 5) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $wp = $wb.Worksheets.Item(1); $wp.Name = 'Параметры'
  $we = $wb.Worksheets.Item(2); $we.Name = 'Сотрудники'
  $wv = $wb.Worksheets.Item(3); $wv.Name = 'Отпуска'
  $wo = $wb.Worksheets.Item(4); $wo.Name = 'Остатки'
  $wk = $wb.Worksheets.Item(5); $wk.Name = 'Календарь'

  # ---------- Параметры ----------
  $wp.Range('A1').Formula = 'Параметр'; $wp.Range('B1').Formula = 'Значение'
  $wp.Range('A2').Formula = 'Год расчёта';                          $wp.Range('B2').Formula = '2026'
  $wp.Range('A3').Formula = 'Дата расчёта';                         $wp.Range('B3').Formula = '=TODAY()'
  $wp.Range('A4').Formula = 'Рабочих дней отпуска в год';           $wp.Range('B4').Formula = '20'
  $wp.Range('A5').Formula = 'Максимальный остаток прошлого года';   $wp.Range('B5').Formula = '5'
  $wp.Range('A6').Formula = 'Максимальный перенос на следующий год'; $wp.Range('B6').Formula = '5'
  $wp.Range('D1').Formula = 'Нерабочие праздничные дни'
  $wp.Range('D2').Formula = '=DATE(2026,12,16)'
  $wp.Range('D3').Formula = '=DATE(2026,12,17)'
  Setp $wp.Range('B3') 'NumberFormat' 'dd.mm.yyyy'
  Setp $wp.Range('D2:D60') 'NumberFormat' 'dd.mm.yyyy'
  L 'Параметры'

  # ---------- Сотрудники ----------
  $h = @('ФИО', 'Должность', 'Дата начала работы', 'Дата увольнения', 'Остаток с прошлого года (дн.)', 'Компенсация по приказу об увольнении (дн.)', 'ИТД №')
  for ($i = 0; $i -lt $h.Count; $i++) { $we.Cells.Item(1, $i + 1).Formula = $h[$i] }
  Setp $we.Range('C2:D31') 'NumberFormat' 'dd.mm.yyyy'
  Setp $we.Range('A2:B31') 'NumberFormat' '@'
  Setp $we.Range('G2:G31') 'NumberFormat' '@'

  # ---------- Отпуска ----------
  $h = @('№', 'Сотрудник', 'Вид отпуска', 'С', 'По', 'Рабочих дней', 'Приказ №', 'Дата приказа', 'Заявление от', 'Основание')
  for ($i = 0; $i -lt $h.Count; $i++) { $wv.Cells.Item(1, $i + 1).Formula = $h[$i] }
  Setp $wv.Range('D2:E301') 'NumberFormat' 'dd.mm.yyyy'
  Setp $wv.Range('H2:I301') 'NumberFormat' 'dd.mm.yyyy'
  Setp $wv.Range('G2:G301') 'NumberFormat' '@'
  Setp $wv.Range('J2:J301') 'NumberFormat' '@'
  $wv.Range('F2:F301').Formula = '=IF(OR(D2="",E2=""),"",NETWORKDAYS(D2,E2,Параметры!$D$2:$D$60))'
  $wv.Range('A2:A301').Formula = '=IF(B2="","",ROW()-1)'
  try {
    $v = $wv.Range('B2:B301').Validation
    [void]$v.GetType().InvokeMember('Add', [Reflection.BindingFlags]::InvokeMethod, $null, $v, @(3, 1, 1, '=Сотрудники!$A$2:$A$31'))
    $v2 = $wv.Range('C2:C301').Validation
    [void]$v2.GetType().InvokeMember('Add', [Reflection.BindingFlags]::InvokeMethod, $null, $v2, @(3, 1, 1, 'Очередной (О),Больничный (Б),Командировка (К),Учебный (У)'))
    L 'списки выбора добавлены'
  } catch { L ('списки выбора: ' + $_.Exception.Message) }
  L 'Отпуска'

  # ---------- Остатки ----------
  $h = @('Сотрудник', 'Начало работы', 'Дата увольнения', 'Перенос с прошлого года', 'Мес. начисления на дату расчёта', 'Мес. начисления на 31.12', 'Использовано до даты расчёта', 'Запланировано после даты', 'Доступно на дату расчёта', 'Доступно на 31.12', 'Перенос на следующий год', 'Сгорит, если не использовать до 31.12', 'К выплате при увольнении (по расчёту, дн.)', 'Компенсация по приказу (дн.)', 'Рабочих дней до 31.12')
  for ($i = 0; $i -lt $h.Count; $i++) { $wo.Cells.Item(1, $i + 1).Formula = $h[$i] }
  $Y = 'Параметры!$B$2'; $X = 'Параметры!$B$3'; $N = 'Параметры!$B$4'
  $rng = 'Отпуска!$B$2:$B$301'; $kind = 'Отпуска!$C$2:$C$301'; $st = 'Отпуска!$D$2:$D$301'; $dd = 'Отпуска!$F$2:$F$301'
  $wo.Range('A2:A31').Formula = '=IF(Сотрудники!A2="","",Сотрудники!A2)'
  $wo.Range('B2:B31').Formula = '=IF(A2="","",Сотрудники!C2)'
  $wo.Range('C2:C31').Formula = '=IF(A2="","",IF(Сотрудники!D2="","",Сотрудники!D2))'
  $wo.Range('D2:D31').Formula = '=IF(A2="","",MIN(Параметры!$B$5,N(Сотрудники!E2)))'
  $mD = 'MIN(12,DATEDIF(B2,MIN(#END#,C2),"m")+1-IF(B2<DATE(' + $Y + ',1,1),DATEDIF(B2,DATE(' + $Y + '-1,12,31),"m")+1,0))'
  $mW = 'IF(B2<DATE(' + $Y + ',1,1),12,MIN(12,DATEDIF(B2,#END#,"m")+1))'
  $fE = '=IF(A2="","",IF(B2>' + $X + ',0,IF(C2<>"",IF(C2<DATE(' + $Y + ',1,1),0,' + ($mD -replace '#END#', $X) + '),' + ($mW -replace '#END#', $X) + ')))'
  $e31 = 'DATE(' + $Y + ',12,31)'
  $fF = '=IF(A2="","",IF(B2>' + $e31 + ',0,IF(C2<>"",IF(C2<DATE(' + $Y + ',1,1),0,' + ($mD -replace '#END#', $e31) + '),' + ($mW -replace '#END#', $e31) + ')))'
  $wo.Range('E2:E31').Formula = $fE
  $wo.Range('F2:F31').Formula = $fF
  $wo.Range('G2:G31').Formula = '=IF(A2="","",SUMIFS(' + $dd + ',' + $rng + ',A2,' + $kind + ',"Очередной (О)",' + $st + ',">="&DATE(' + $Y + ',1,1),' + $st + ',"<="&' + $X + '))'
  $wo.Range('H2:H31').Formula = '=IF(A2="","",SUMIFS(' + $dd + ',' + $rng + ',A2,' + $kind + ',"Очередной (О)",' + $st + ',">"&' + $X + ',' + $st + ',"<="&DATE(' + $Y + ',12,31)))'
  $wo.Range('I2:I31').Formula = '=IF(A2="","",D2+' + $N + '*E2/12-G2)'
  $wo.Range('J2:J31').Formula = '=IF(A2="","",D2+' + $N + '*F2/12-G2-H2)'
  $wo.Range('K2:K31').Formula = '=IF(A2="","",IF(C2<>"",0,MIN(Параметры!$B$6,MAX(0,J2))))'
  $wo.Range('L2:L31').Formula = '=IF(A2="","",IF(C2<>"",0,MAX(0,J2-K2)))'
  $wo.Range('M2:M31').Formula = '=IF(OR(A2="",C2=""),"",MAX(0,J2))'
  $wo.Range('N2:N31').Formula = '=IF(OR(A2="",Сотрудники!F2=""),"",Сотрудники!F2)'
  $wo.Range('O2:O31').Formula = '=IF(A2="","",IF(C2<>"",0,NETWORKDAYS(' + $X + ',' + $e31 + ',Параметры!$D$2:$D$60)))'
  Setp $wo.Range('B2:C31') 'NumberFormat' 'dd.mm.yyyy'
  Setp $wo.Range('I2:M31') 'NumberFormat' '0.0'
  L 'Остатки'

  # ---------- Календарь ----------
  $wk.Range('A1').Formula = 'Сотрудник'
  $mn = @('Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь')
  for ($m = 1; $m -le 12; $m++) { $wk.Cells.Item(1, $m + 1).Formula = $mn[$m - 1] }
  $wk.Range('N1').Formula = 'Всего за год'
  $wk.Range('A2:A31').Formula = '=IF(Сотрудники!A2="","",Сотрудники!A2)'
  for ($m = 1; $m -le 12; $m++) {
    $col = $wk.Cells.Item(2, $m + 1).Address($false, $false) -replace '\d', ''
    $f = '=IF($A2="","",SUMIFS(' + $dd + ',' + $rng + ',$A2,' + $kind + ',"Очередной (О)",' + $st + ',">="&DATE(' + $Y + ',' + $m + ',1),' + $st + ',"<="&EOMONTH(DATE(' + $Y + ',' + $m + ',1),0)))'
    $wk.Range($col + '2:' + $col + '31').Formula = $f
  }
  $wk.Range('N2:N31').Formula = '=IF($A2="","",SUM(B2:M2))'
  Setp $wk.Range('B2:N31') 'NumberFormat' '0;-0;;@'
  L 'Календарь'

  # ---------- оформление ----------
  $sheets = @(
    @($wp, 'A1:B6', 'A1:B1', @(38, 14), 'A1:B6'),
    @($we, 'A1:G31', 'A1:G1', @(30, 44, 16, 16, 18, 22, 12), 'A1:G31'),
    @($wv, 'A1:J301', 'A1:J1', @(5, 28, 24, 12, 12, 12, 12, 12, 12, 30), 'A1:J301'),
    @($wo, 'A1:O31', 'A1:O1', @(28, 14, 14, 13, 14, 14, 14, 14, 14, 14, 14, 16, 16, 16, 14), 'A1:O31'),
    @($wk, 'A1:N31', 'A1:N1', @(28, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 12), 'A1:N31')
  )
  foreach ($s in $sheets) {
    $ws = $s[0]
    $b = $ws.Range($s[1]).Borders; Setp $b 'LineStyle' ([int]1)
    $hr = $ws.Range($s[2])
    Setp $hr.Font 'Bold' $true
    Setp $hr.Interior 'Color' ([double]15921906)
    Setp $hr 'WrapText' $true
    Setp $hr 'HorizontalAlignment' ([int]-4108)
    Setp $hr 'VerticalAlignment' ([int]-4108)
    for ($c = 0; $c -lt $s[3].Count; $c++) { Setp $ws.Columns.Item($c + 1) 'ColumnWidth' ([double]$s[3][$c]) }
    Setp $ws.Rows.Item(1) 'RowHeight' ([double]48)
    Setp $ws.PageSetup 'Orientation' ([int]2)
    Setp $ws.PageSetup 'Zoom' $false
    Setp $ws.PageSetup 'FitToPagesWide' ([int]1)
    Setp $ws.PageSetup 'FitToPagesTall' $false
  }
  # поля ввода - светло-жёлтые
  $yellow = [double]13431551
  foreach ($a in @(@($wp, 'B2:B6'), @($wp, 'D2:D60'), @($we, 'A2:G31'), @($wv, 'B2:E301'), @($wv, 'G2:J301'))) { Setp $a[0].Range($a[1]).Interior 'Color' $yellow }
  Setp $wp.Range('D1') 'WrapText' $true
  Setp $wp.Columns.Item(4) 'ColumnWidth' ([double]24)
  Setp $wp.Range('D1').Font 'Bold' $true
  Setp $wp.Range('D1').Interior 'Color' ([double]15921906)
  Setp $wp.Range('D1') 'HorizontalAlignment' ([int]-4108)
  $bd = $wp.Range('D1:D60').Borders; Setp $bd 'LineStyle' ([int]1)
  foreach ($ws in @($wo, $wk, $wv, $we)) { [void]$ws.Activate(); Setp $xl.ActiveWindow 'SplitRow' ([int]1); Setp $xl.ActiveWindow 'FreezePanes' $true }
  Setp $wv.Range('D2:I301') 'HorizontalAlignment' ([int]-4108)
  Setp $wo.Range('B2:O31') 'HorizontalAlignment' ([int]-4108)
  Setp $wk.Range('B2:N31') 'HorizontalAlignment' ([int]-4108)
  Setp $we.Range('C2:G31') 'HorizontalAlignment' ([int]-4108)
  try {
    $dp = $wb.BuiltinDocumentProperties
    $it = [System.__ComObject].InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $dp, @('Company'))
    [System.__ComObject].InvokeMember('Value', [Reflection.BindingFlags]::SetProperty, $null, $it, @('CACF'))
  } catch { L ('Company: ' + $_.Exception.Message) }
  [void]$wp.Activate()
  $xl.Calculate()
  $wb.SaveAs($out, 51)
  L 'шаблон сохранён'

  # ---------- проверка на тестовых данных (в файл не сохраняется) ----------
  $emp = @(
    @('Қаражан Е.Н.', '=DATE(2025,2,18)', '', 5, ''), @('Темирбеков М.Е.', '=DATE(2025,2,24)', '', 5, ''),
    @('Оразбаева М.К.', '=DATE(2025,3,3)', '', 5, ''), @('Кононец А.В.', '=DATE(2026,2,10)', '', 0, ''),
    @('Джангарашев Р.М.', '=DATE(2024,7,1)', '=DATE(2026,1,16)', 0, 5), @('Альмаганбетов А.С.', '=DATE(2026,1,5)', '=DATE(2026,7,7)', 0, 10),
    @('Ажигалиева Д.Н.', '=DATE(2024,8,21)', '=DATE(2026,5,28)', 5, ''), @('Салыкова Д.О.', '=DATE(2024,9,16)', '=DATE(2026,5,28)', 5, '')
  )
  $r = 2
  foreach ($e in $emp) {
    $we.Cells.Item($r, 1).Formula = $e[0]
    $we.Cells.Item($r, 3).Formula = $e[1]
    if ($e[2] -ne '') { $we.Cells.Item($r, 4).Formula = $e[2] }
    $we.Cells.Item($r, 5).Formula = [string]$e[3]
    if ($e[4] -ne '') { $we.Cells.Item($r, 6).Formula = [string]$e[4] }
    $r++
  }
  $lv = @(
    @('Қаражан Е.Н.', 2026,8,12, 2026,8,19), @('Қаражан Е.Н.', 2026,8,27, 2026,9,2), @('Қаражан Е.Н.', 2026,10,7, 2026,10,14),
    @('Темирбеков М.Е.', 2026,6,3, 2026,6,9), @('Темирбеков М.Е.', 2026,9,7, 2026,9,18),
    @('Оразбаева М.К.', 2026,9,1, 2026,9,8), @('Оразбаева М.К.', 2026,9,16, 2026,9,29),
    @('Кононец А.В.', 2026,8,24, 2026,9,4),
    @('Ажигалиева Д.Н.', 2026,4,29, 2026,5,12), @('Ажигалиева Д.Н.', 2026,5,18, 2026,5,27), @('Салыкова Д.О.', 2026,5,12, 2026,5,27)
  )
  $r = 2
  foreach ($e in $lv) {
    $wv.Cells.Item($r, 2).Formula = $e[0]
    $wv.Cells.Item($r, 3).Formula = 'Очередной (О)'
    $wv.Cells.Item($r, 4).Formula = "=DATE($($e[1]),$($e[2]),$($e[3]))"
    $wv.Cells.Item($r, 5).Formula = "=DATE($($e[4]),$($e[5]),$($e[6]))"
    $r++
  }
  $wp.Range('B3').Formula = '=DATE(2026,11,10)'
  $xl.Calculate()
  $res = @()
  foreach ($ws in @($wv, $wo, $wk)) {
    $lastr = if ($ws.Name -eq 'Отпуска') { 13 } else { 9 }
    $lastc = if ($ws.Name -eq 'Отпуска') { 6 } else { 15 }
    for ($rr = 1; $rr -le $lastr; $rr++) {
      $l = @(); for ($c = 1; $c -le $lastc; $c++) { $l += $ws.Cells.Item($rr, $c).Text }
      $res += ($ws.Name + ' | ' + ($l -join ' ; '))
    }
  }
  $res | Set-Content -Path 'D:\Кадры\_исходники\build_shablon_otpuska_check.txt' -Encoding UTF8
  L 'проверка записана'
  $wb.Close($false)
}
catch { L ('ОШИБКА: ' + $_.Exception.Message + ' @' + $_.InvocationInfo.ScriptLineNumber) }
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
