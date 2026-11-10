$ErrorActionPreference='Stop'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
$f='D:\Кадры\График отпусков и остатки на 2026 (2026-11-10).xlsx'
$log='D:\Кадры\_исходники\add_former_employees.log'; Remove-Item $log -ErrorAction SilentlyContinue
function L($m){ Add-Content $log ((Get-Date -Format 'HH:mm:ss')+' '+$m) -Encoding UTF8 }
$xl=New-Object -ComObject Excel.Application
Setp $xl 'Visible' $false; Setp $xl 'DisplayAlerts' $false
try{
 $wb=$xl.Workbooks.Open($f)
 $w1=$wb.Worksheets.Item('График отпусков 2026'); $w2=$wb.Worksheets.Item('Остатки на 10.11.2026')
 # --- график: уволенные ---
 $rows=@(
  @('Ажигалиева Д.Н.','Очередной (О)',@(2026,4,29),@(2026,5,12),7,'03-7/1','апрель-май 2026: 29-30.04, 04-06.05, 08.05, 15.05'),
  @('Ажигалиева Д.Н.','Очередной (О)',@(2026,5,18),@(2026,5,27),7,'03-7/7','май 2026: 18-22.05, 25-27.05'),
  @('Салыкова Д.О.','Очередной (О)',@(2026,5,12),@(2026,5,27),11,'03-7/3','май 2026: 12-15.05, 18-22.05, 25-27.05'),
  @('Альмаганбетов А.С.','Без содержания (Об)',@(2026,5,12),@(2026,5,14),3,'','май 2026: 12-14.05')
 )
 $r=14
 foreach($e in $rows){
  $w1.Cells.Item($r,1).Formula=[string]($r-1)
  $w1.Cells.Item($r,2).Formula=$e[0]; $w1.Cells.Item($r,3).Formula=$e[1]
  $w1.Cells.Item($r,4).Formula="=DATE($($e[2][0]),$($e[2][1]),$($e[2][2]))"
  $w1.Cells.Item($r,5).Formula="=DATE($($e[3][0]),$($e[3][1]),$($e[3][2]))"
  $w1.Cells.Item($r,6).Formula=[string]$e[4]
  if($e[5] -ne ''){ $c=$w1.Cells.Item($r,7); Setp $c 'NumberFormat' '@'; $c.Formula='№'+$e[5] }
  $w1.Cells.Item($r,8).Formula=$e[6]
  $r++
 }
 Setp $w1.Range('D14:E17') 'NumberFormat' 'dd.mm.yyyy'
 Setp $w1.Range('A14:H17').Borders 'LineStyle' ([int]1)
 Setp $w1.Range('A14:A17') 'HorizontalAlignment' ([int]-4108)
 Setp $w1.Range('D14:G17') 'HorizontalAlignment' ([int]-4108)
 Setp $w1.Range('A14:H17') 'VerticalAlignment' ([int]-4108)
 L 'график дополнен'
 # --- остатки: блок уволенных ---
 $w2.Range('A8').Formula='Работники, уволенные в 2026 году'
 Setp $w2.Range('A8').Font 'Bold' $true
 $h=@('Сотрудник','Начало работы','Дата увольнения','Приказ об увольнении','Использовано в 2026 (очередной)','Компенсация неиспользованного отпуска по приказу (дн.)')
 for($i=0;$i -lt $h.Count;$i++){ $w2.Cells.Item(9,$i+1).Formula=$h[$i] }
 $fe=@(
  @('Джангарашев Р.М.',@(2024,7,1),@(2026,1,16),'3-6/2',5),
  @('Альмаганбетов А.С.',@(2026,1,5),@(2026,7,7),'03-6/7',10),
  @('Ажигалиева Д.Н.',@(2024,8,21),@(2026,5,28),'03-6/6',''),
  @('Салыкова Д.О.',@(2024,9,16),@(2026,5,28),'03-6/5','')
 )
 $r=10
 foreach($e in $fe){
  $w2.Cells.Item($r,1).Formula=$e[0]
  $w2.Cells.Item($r,2).Formula="=DATE($($e[1][0]),$($e[1][1]),$($e[1][2]))"
  $w2.Cells.Item($r,3).Formula="=DATE($($e[2][0]),$($e[2][1]),$($e[2][2]))"
  $c=$w2.Cells.Item($r,4); Setp $c 'NumberFormat' '@'; $c.Formula='№'+$e[3]
  $w2.Cells.Item($r,5).Formula="=SUMIFS('График отпусков 2026'!`$F`$2:`$F`$40,'График отпусков 2026'!`$B`$2:`$B`$40,A$r,'График отпусков 2026'!`$C`$2:`$C`$40,""Очередной (О)"")"
  if($e[4] -ne ''){ $w2.Cells.Item($r,6).Formula=[string]$e[4] } else { $w2.Cells.Item($r,6).Formula='-' }
  $r++
 }
 Setp $w2.Range('B10:C13') 'NumberFormat' 'dd.mm.yyyy'
 $hr=$w2.Range('A9:F9'); Setp $hr.Font 'Bold' $true; Setp $hr.Interior 'Color' ([double]15921906); Setp $hr 'WrapText' $true; Setp $hr 'HorizontalAlignment' ([int]-4108); Setp $hr 'VerticalAlignment' ([int]-4108)
 Setp $w2.Range('A9:F13').Borders 'LineStyle' ([int]1)
 Setp $w2.Range('B10:F13') 'HorizontalAlignment' ([int]-4108)
 Setp $w2.Rows.Item(9) 'RowHeight' ([double]62)
 Setp $w2.Columns.Item(13) 'ColumnWidth' ([double]3)
 $xl.Calculate()
 $out=@()
 foreach($ws in @($w1,$w2)){ $ur=$ws.UsedRange; for($rr=1;$rr -le $ur.Rows.Count;$rr++){ $l=@(); for($c=1;$c -le [Math]::Min(12,$ur.Columns.Count);$c++){$l+=$ws.Cells.Item($rr,$c).Text}; $out+=($ws.Name+' | '+($l -join ' ; ')) } }
 $out | Set-Content 'D:\Кадры\_исходники\add_former_employees_check.txt' -Encoding UTF8
 $wb.Save(); $wb.Close($false); L 'сохранено'
}catch{ L ('ОШИБКА: '+$_.Exception.Message+' @'+$_.InvocationInfo.ScriptLineNumber) }
finally{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
