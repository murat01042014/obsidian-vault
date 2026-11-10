$ErrorActionPreference='Stop'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
$tpl='D:\Шаблоны\Отпуска\Шаблон - график отпусков и остатки.xlsx'
$main='D:\Кадры\График отпусков и остатки на 2026 (2026-11-10).xlsx'
$log='D:\Кадры\_исходники\apply_holidays_2026.log'; Remove-Item $log -ErrorAction SilentlyContinue
function L($m){ Add-Content $log ((Get-Date -Format 'HH:mm:ss')+' '+$m) -Encoding UTF8 }
Add-Type -Namespace W -Name U -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$xl=New-Object -ComObject Excel.Application
$p=[uint32]0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$p); L "PID $p"
Setp $xl 'Visible' $false; Setp $xl 'DisplayAlerts' $false; Setp $xl 'AutomationSecurity' ([int]3)
$hol=@(
 @(2026,1,1,'Новый год'),@(2026,1,2,'Новый год'),@(2026,1,7,'Рождество Христово'),
 @(2026,3,9,'Перенос с 08.03 (Международный женский день)'),@(2026,3,23,'Наурыз мейрамы'),
 @(2026,3,24,'Перенос с 21.03 (Наурыз мейрамы)'),@(2026,3,25,'Перенос с 22.03 (Наурыз мейрамы)'),
 @(2026,5,1,'Праздник единства народа Казахстана'),@(2026,5,7,'День защитника Отечества'),
 @(2026,5,11,'Перенос с 09.05 (День Победы)'),@(2026,5,27,'Курбан айт'),@(2026,7,6,'День Столицы'),
 @(2026,10,26,'Перенос с 25.10 (День Республики)'),@(2026,12,16,'День Независимости')
)
try{
 # ---------- шаблон ----------
 $wb=$xl.Workbooks.Open($tpl,[int]0,$false)
 $wp=$wb.Worksheets.Item('Параметры')
 $wp.Range('D1').Formula='Нерабочие праздничные дни (будни)'
 $wp.Range('E1').Formula='Праздник / основание'
 $wp.Range('D2:E60').ClearContents() | Out-Null
 $r=2
 foreach($h in $hol){ $wp.Cells.Item($r,4).Formula=("=DATE("+$h[0]+","+$h[1]+","+$h[2]+")"); $c=$wp.Cells.Item($r,5); Setp $c 'NumberFormat' '@'; $c.Formula=$h[3]; $r++ }
 Setp $wp.Range('D2:D60') 'NumberFormat' 'dd.mm.yyyy'
 Setp $wp.Columns.Item(4) 'ColumnWidth' ([double]28); Setp $wp.Columns.Item(5) 'ColumnWidth' ([double]46)
 $hr=$wp.Range('E1'); Setp $hr.Font 'Bold' $true; Setp $hr.Interior 'Color' ([double]15921906); Setp $hr 'HorizontalAlignment' ([int]-4108); Setp $hr 'WrapText' $true
 Setp $wp.Range('D1:E60').Borders 'LineStyle' ([int]1)
 Setp $wp.Range('E2:E60').Interior 'Color' ([double]13431551)
 Setp $wp.Range('E2:E60') 'HorizontalAlignment' ([int]-4131)
 Setp $wp.Range('D2:D60') 'HorizontalAlignment' ([int]-4108)
 $xl.Calculate()
 $wb.Save(); L ('шаблон: праздников внесено '+$hol.Count)
 $wb.Close($false)

 # ---------- основной файл: только 16.12 ----------
 $wb=$xl.Workbooks.Open($main,[int]0,$false)
 $w2=$wb.Worksheets.Item('Остатки на 10.11.2026')
 $before=@($w2.Range('P6').Text,$w2.Range('R6').Text,$w2.Range('R7').Text)
 $w2.Range('R6').Formula='=DATE(2026,12,16)'
 $w2.Range('R7').ClearContents() | Out-Null
 $xl.Calculate()
 L ('основной: было P6='+$before[0]+' R6='+$before[1]+' R7='+$before[2]+'; стало P6='+$w2.Range('P6').Text+' R6='+$w2.Range('R6').Text+' R7='+$w2.Range('R7').Text)
 $wb.Save(); $wb.Close($false)
}catch{ L ('ОШИБКА: '+$_.Exception.Message+' @'+$_.InvocationInfo.ScriptLineNumber) }
finally{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
