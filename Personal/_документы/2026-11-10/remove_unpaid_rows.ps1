$ErrorActionPreference='Stop'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
$f='D:\Кадры\График отпусков и остатки на 2026 (2026-11-10).xlsx'
$xl=New-Object -ComObject Excel.Application
Setp $xl 'Visible' $false; Setp $xl 'DisplayAlerts' $false
try{
 $wb=$xl.Workbooks.Open($f)
 $w1=$wb.Worksheets.Item('График отпусков 2026'); $w2=$wb.Worksheets.Item('Остатки на 10.11.2026')
 # проверка: удаляем только строки Об
 $del=@(); for($r=17;$r -ge 2;$r--){ if($w1.Cells.Item($r,3).Text -like 'Без содержания*'){ $del+=$r } }
 foreach($r in $del){ [void]$w1.Rows.Item($r).Delete() }
 $last=1; while($w1.Cells.Item($last+1,2).Text -ne ''){ $last++ }
 for($r=2;$r -le $last;$r++){ $w1.Cells.Item($r,1).Formula=[string]($r-1) }
 $xl.Calculate()
 $out=@("удалено строк: $($del.Count) ($($del -join ','))","последняя строка графика: $last")
 for($r=1;$r -le $last;$r++){ $l=@(); for($c=1;$c -le 8;$c++){$l+=$w1.Cells.Item($r,$c).Text}; $out+=($l -join ' ; ') }
 $ur=$w2.UsedRange; for($rr=1;$rr -le $ur.Rows.Count;$rr++){ $l=@(); for($c=1;$c -le 12;$c++){$l+=$w2.Cells.Item($rr,$c).Text}; $out+=($l -join ' ; ') }
 $out+=("формула G2: "+$w2.Range('G2').Formula)
 $out+=("формула E10: "+$w2.Range('E10').Formula)
 $out | Set-Content 'D:\Кадры\_исходники\remove_unpaid_check.txt' -Encoding UTF8
 $wb.Save(); $wb.Close($false)
}finally{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
