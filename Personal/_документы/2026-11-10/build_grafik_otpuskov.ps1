$ErrorActionPreference = 'Stop'
$out = 'D:\Кадры\График отпусков и остатки на 2026 (2026-11-10).xlsx'
$log = 'D:\Кадры\_исходники\build_grafik_otpuskov.log'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
function L($m) { Add-Content -Path $log -Value ((Get-Date -Format 'HH:mm:ss') + ' ' + $m) -Encoding UTF8 }
Remove-Item $log -ErrorAction SilentlyContinue
$xl = New-Object -ComObject Excel.Application
Setp $xl 'Visible' $false
Setp $xl 'DisplayAlerts' $false
try {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $w1 = $wb.Worksheets.Item(1); $w1.Name = 'График отпусков 2026'
  $w2 = $wb.Worksheets.Item(2); $w2.Name = 'Остатки на 10.11.2026'
  $w3 = $wb.Worksheets.Item(3); $w3.Name = 'Отпуска 2025'
  L 'листы созданы'

  function D([int]$y, [int]$m, [int]$d) { return "=DATE($y,$m,$d)" }

  # ---------- Лист 1: график 2026 ----------
  $h1 = @('№', 'Сотрудник', 'Вид отпуска', 'С', 'По', 'Рабочих дней', 'Приказ', 'Табель (ТУРВ)')
  for ($i = 0; $i -lt $h1.Count; $i++) { $w1.Cells.Item(1, $i + 1).Formula = $h1[$i] }
  $r1 = @(
    @('Қаражан Е.Н.', 'Очередной (О)', @(2026,8,12), @(2026,8,19), 6, '03-7/17', 'август 2026: 12-14.08, 17-19.08'),
    @('Қаражан Е.Н.', 'Очередной (О)', @(2026,8,27), @(2026,9,2), 5, '03-7/19', 'август-сентябрь 2026: 27-28.08, 31.08-02.09'),
    @('Қаражан Е.Н.', 'Очередной (О)', @(2026,10,7), @(2026,10,14), 6, '03-7/24', 'октябрь 2026: 07-09.10, 12-14.10'),
    @('Темирбеков М.Е.', 'Очередной (О)', @(2026,6,3), @(2026,6,9), 5, '03-7/13', 'июнь 2026: 03-05.06, 08-09.06'),
    @('Темирбеков М.Е.', 'Очередной (О)', @(2026,9,7), @(2026,9,18), 10, '03-7/21', 'сентябрь 2026: 07-11.09, 14-18.09'),
    @('Оразбаева М.К.', 'Очередной (О)', @(2026,9,1), @(2026,9,8), 6, '03-7/20', 'сентябрь 2026: 01-04.09, 07-08.09'),
    @('Оразбаева М.К.', 'Очередной (О)', @(2026,9,16), @(2026,9,29), 10, '03-7/22', 'сентябрь 2026: 16-18.09, 21-25.09, 28-29.09'),
    @('Кононец А.В.', 'Очередной (О)', @(2026,8,24), @(2026,9,4), 10, '03-7/18', 'август-сентябрь 2026: 24-28.08, 31.08-04.09'),
    @('Қаражан Е.Н.', 'Без содержания (Об)', @(2026,5,18), @(2026,5,29), 9, '', 'май 2026: 18-22.05, 25-26.05, 28-29.05'),
    @('Темирбеков М.Е.', 'Без содержания (Об)', @(2026,5,18), @(2026,5,29), 9, '', 'май 2026: 18-22.05, 25-26.05, 28-29.05'),
    @('Оразбаева М.К.', 'Без содержания (Об)', @(2026,5,18), @(2026,5,29), 10, '', 'май 2026: 18-22.05, 25-29.05'),
    @('Кононец А.В.', 'Без содержания (Об)', @(2026,5,12), @(2026,6,5), 19, '', 'май-июнь 2026: 12-15.05, 18-22.05, 25-29.05, 01-05.06')
  )
  $row = 2
  foreach ($e in $r1) {
    $w1.Cells.Item($row, 1).Formula = [string]($row - 1)
    $w1.Cells.Item($row, 2).Formula = $e[0]
    $w1.Cells.Item($row, 3).Formula = $e[1]
    $w1.Cells.Item($row, 4).Formula = (D $e[2][0] $e[2][1] $e[2][2])
    $w1.Cells.Item($row, 5).Formula = (D $e[3][0] $e[3][1] $e[3][2])
    $w1.Cells.Item($row, 6).Formula = [string]$e[4]
    if ($e[5] -ne '') { Setp $w1.Cells.Item($row, 7) 'NumberFormat' '@'; $w1.Cells.Item($row, 7).Formula = '№' + $e[5] }
    $w1.Cells.Item($row, 8).Formula = $e[6]
    $row++
  }
  $last1 = $row - 1
  Setp $w1.Range("D2:E$last1") 'NumberFormat' 'dd.mm.yyyy'
  L 'лист 1 заполнен'

  # ---------- Лист 3: 2025 ----------
  $w3.Cells.Item(1,1).Formula = 'Сотрудник'
  $h3 = @('Сотрудник', 'Начало работы', 'Дней в год', 'Мес. начисления в 2025', 'Начислено за 2025', 'Использовано в 2025', 'Остаток на 31.12.2025', 'Перенос в 2026 (макс. 5)')
  for ($i = 0; $i -lt $h3.Count; $i++) { $w3.Cells.Item(1, $i + 1).Formula = $h3[$i] }
  $s3 = @(
    @('Қаражан Е.Н.', @(2025,2,18), 11),
    @('Темирбеков М.Е.', @(2025,2,24), 11),
    @('Оразбаева М.К.', @(2025,3,3), 10)
  )
  $row = 2
  foreach ($e in $s3) {
    $w3.Cells.Item($row, 1).Formula = $e[0]
    $w3.Cells.Item($row, 2).Formula = (D $e[1][0] $e[1][1] $e[1][2])
    $w3.Cells.Item($row, 3).Formula = '20'
    $w3.Cells.Item($row, 4).Formula = [string]$e[2]
    $w3.Cells.Item($row, 5).Formula = "=C$row*D$row/12"
    $w3.Cells.Item($row, 6).Formula = "=SUMIFS(`$F`$8:`$F`$12,`$A`$8:`$A`$12,A$row)"
    $w3.Cells.Item($row, 7).Formula = "=E$row-F$row"
    $w3.Cells.Item($row, 8).Formula = "=MIN(5,MAX(0,G$row))"
    $row++
  }
  Setp $w3.Range('B2:B4') 'NumberFormat' 'dd.mm.yyyy'
  $h3b = @('Сотрудник', 'Вид отпуска', 'С', 'По', 'Рабочих дней', 'Приказ')
  $w3.Cells.Item(7,1).Formula = 'Сотрудник'
  for ($i = 0; $i -lt $h3b.Count; $i++) { $w3.Cells.Item(7, $i + 1).Formula = $h3b[$i] }
  # колонка F на строках 8-12 = рабочие дни, поэтому порядок колонок: A сотрудник, B вид, C с, D по, E (пусто), F дней, G приказ
  $w3.Cells.Item(7,5).Formula = ''
  $w3.Cells.Item(7,6).Formula = 'Рабочих дней'
  $w3.Cells.Item(7,7).Formula = 'Приказ'
  $d25 = @(
    @('Қаражан Е.Н.', 'Очередной (О)', @(2025,8,25), @(2025,9,5), 10, '6-4 от 21.08.2025'),
    @('Қаражан Е.Н.', 'Очередной (О)', @(2025,12,25), @(2025,12,29), 3, '6-18 от 24.12.2025'),
    @('Темирбеков М.Е.', 'Очередной (О)', @(2025,11,17), @(2025,11,21), 5, '6-10'),
    @('Темирбеков М.Е.', 'Очередной (О)', @(2025,12,1), @(2025,12,9), 7, '6-13'),
    @('Оразбаева М.К.', 'Очередной (О)', @(2025,12,8), @(2025,12,22), 10, '6-14')
  )
  $row = 8
  foreach ($e in $d25) {
    $w3.Cells.Item($row, 1).Formula = $e[0]
    $w3.Cells.Item($row, 2).Formula = $e[1]
    $w3.Cells.Item($row, 3).Formula = (D $e[2][0] $e[2][1] $e[2][2])
    $w3.Cells.Item($row, 4).Formula = (D $e[3][0] $e[3][1] $e[3][2])
    $w3.Cells.Item($row, 6).Formula = [string]$e[4]
    Setp $w3.Cells.Item($row, 7) 'NumberFormat' '@'; $w3.Cells.Item($row, 7).Formula = '№' + $e[5]
    $row++
  }
  Setp $w3.Range('C8:D12') 'NumberFormat' 'dd.mm.yyyy'
  Setp $w3.Range('E2:E4') 'NumberFormat' '0.0'
  Setp $w3.Range('G2:H4') 'NumberFormat' '0.0'
  L 'лист 3 заполнен'

  # ---------- Лист 2: остатки ----------
  $h2 = @('Сотрудник', 'Начало работы', 'Дней в год', 'Мес. начисления на 10.11.2026', 'Мес. начисления на 31.12.2026', 'Перенос с 2025 (макс. 5)', 'Использовано в 2026', 'Доступно на 10.11.2026', 'Доступно на 31.12.2026', 'Перенос на 2027 (макс. 5)', 'Сгорит, если не использовать до 31.12', 'Рабочих дней с 10.11 по 31.12')
  for ($i = 0; $i -lt $h2.Count; $i++) { $w2.Cells.Item(1, $i + 1).Formula = $h2[$i] }
  $s2 = @(
    @('Қаражан Е.Н.', @(2025,2,18), 12, 12, '=''Отпуска 2025''!H2'),
    @('Темирбеков М.Е.', @(2025,2,24), 12, 12, '=''Отпуска 2025''!H3'),
    @('Оразбаева М.К.', @(2025,3,3), 12, 12, '=''Отпуска 2025''!H4'),
    @('Кононец А.В.', @(2026,2,10), 10, 11, '0')
  )
  $row = 2
  foreach ($e in $s2) {
    $w2.Cells.Item($row, 1).Formula = $e[0]
    $w2.Cells.Item($row, 2).Formula = (D $e[1][0] $e[1][1] $e[1][2])
    $w2.Cells.Item($row, 3).Formula = '20'
    $w2.Cells.Item($row, 4).Formula = [string]$e[2]
    $w2.Cells.Item($row, 5).Formula = [string]$e[3]
    $w2.Cells.Item($row, 6).Formula = $e[4]
    $w2.Cells.Item($row, 7).Formula = "=SUMIFS('График отпусков 2026'!`$F`$2:`$F`$$last1,'График отпусков 2026'!`$B`$2:`$B`$$last1,A$row,'График отпусков 2026'!`$C`$2:`$C`$$last1,""Очередной (О)"")"
    $w2.Cells.Item($row, 8).Formula = "=F$row+C$row*D$row/12-G$row"
    $w2.Cells.Item($row, 9).Formula = "=F$row+C$row*E$row/12-G$row"
    $w2.Cells.Item($row, 10).Formula = "=MIN(5,MAX(0,I$row))"
    $w2.Cells.Item($row, 11).Formula = "=MAX(0,I$row-J$row)"
    $w2.Cells.Item($row, 12).Formula = "=NETWORKDAYS(DATE(2026,11,10),DATE(2026,12,31),`$N`$2:`$N`$3)"
    $row++
  }
  $w2.Cells.Item(1,14).Formula = 'Нерабочие праздничные дни'
  $w2.Cells.Item(2,14).Formula = '=DATE(2026,12,16)'
  $w2.Cells.Item(3,14).Formula = '=DATE(2026,12,17)'
  Setp $w2.Range('B2:B5') 'NumberFormat' 'dd.mm.yyyy'
  Setp $w2.Range('N2:N3') 'NumberFormat' 'dd.mm.yyyy'
  Setp $w2.Range('H2:K5') 'NumberFormat' '0.0'
  L 'лист 2 заполнен'

  # ---------- оформление ----------
  foreach ($pair in @(@($w1, "A1:H$last1"), @($w2, 'A1:L5'), @($w3, 'A1:H4'), @($w3, 'A7:G12'))) {
    $ws = $pair[0]; $rg = $ws.Range($pair[1])
    Setp $rg.Borders 'LineStyle' ([int]1)
    Setp $rg 'VerticalAlignment' ([int]-4108)
  }
  foreach ($pair in @(@($w1,'A1:H1'), @($w2,'A1:L1'), @($w3,'A1:H1'), @($w3,'A7:G7'), @($w2,'N1'))) {
    $hr = $pair[0].Range($pair[1])
    Setp $hr.Font 'Bold' $true
    Setp $hr.Interior 'Color' ([double]15921906)
    Setp $hr 'WrapText' $true
    Setp $hr 'HorizontalAlignment' ([int]-4108)
  }
  Setp $w1.Columns.Item(1) 'ColumnWidth' ([double]5); Setp $w1.Columns.Item(2) 'ColumnWidth' ([double]20); Setp $w1.Columns.Item(3) 'ColumnWidth' ([double]22)
  Setp $w1.Columns.Item(4) 'ColumnWidth' ([double]12); Setp $w1.Columns.Item(5) 'ColumnWidth' ([double]12); Setp $w1.Columns.Item(6) 'ColumnWidth' ([double]12)
  Setp $w1.Columns.Item(7) 'ColumnWidth' ([double]10); Setp $w1.Columns.Item(8) 'ColumnWidth' ([double]48)
  Setp $w1.Range("A2:A$last1") 'HorizontalAlignment' ([int]-4108)
  Setp $w1.Range("D2:G$last1") 'HorizontalAlignment' ([int]-4108)
  Setp $w2.Columns.Item(1) 'ColumnWidth' ([double]20)
  for ($c = 2; $c -le 12; $c++) { $w2.Columns.Item($c).ColumnWidth = 15 }
  Setp $w2.Columns.Item(14) 'ColumnWidth' ([double]18)
  Setp $w2.Range('B2:L5') 'HorizontalAlignment' ([int]-4108)
  Setp $w2.Rows.Item(1) 'RowHeight' ([double]62)
  Setp $w3.Columns.Item(1) 'ColumnWidth' ([double]20)
  for ($c = 2; $c -le 8; $c++) { $w3.Columns.Item($c).ColumnWidth = 16 }
  Setp $w3.Rows.Item(1) 'RowHeight' ([double]48)
  Setp $w3.Range('B2:H4') 'HorizontalAlignment' ([int]-4108)
  Setp $w3.Range('B8:G12') 'HorizontalAlignment' ([int]-4108)
  $w1.Activate(); Setp $xl.ActiveWindow 'SplitRow' ([int]1); Setp $xl.ActiveWindow 'FreezePanes' $true
  foreach ($ws in @($w1, $w2, $w3)) {
    Setp $ws.PageSetup 'Orientation' ([int]2)
    Setp $ws.PageSetup 'Zoom' $false
    Setp $ws.PageSetup 'FitToPagesWide' ([int]1)
    Setp $ws.PageSetup 'FitToPagesTall' ([int]1)
  }
  $xl.Calculate()
  try { $dp = $wb.BuiltinDocumentProperties; $it = [System.__ComObject].InvokeMember('Item', [Reflection.BindingFlags]::GetProperty, $null, $dp, @('Company')); [System.__ComObject].InvokeMember('Value', [Reflection.BindingFlags]::SetProperty, $null, $it, @('CACF')) } catch { L ('Company: ' + $_.Exception.Message) }
  $w2.Activate()
  if (Test-Path $out) { L 'файл уже есть, перезапись запрещена'; throw 'exists' }
  $wb.SaveAs($out, 51)
  L 'сохранено'

  # перечитать значения
  $res = @()
  foreach ($ws in @($w2, $w3)) {
    $ur = $ws.UsedRange
    for ($r = 1; $r -le $ur.Rows.Count; $r++) {
      $line = @()
      for ($c = 1; $c -le $ur.Columns.Count; $c++) { $line += $ws.Cells.Item($r, $c).Text }
      $res += ($ws.Name + ' | ' + ($line -join ' ; '))
    }
  }
  $res | Set-Content -Path 'D:\Кадры\_исходники\build_grafik_otpuskov_check.txt' -Encoding UTF8
  L 'проверка записана'
  $wb.Close($false)
}
catch { L ('ОШИБКА: ' + $_.Exception.Message + ' @' + $_.InvocationInfo.ScriptLineNumber) }
finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
