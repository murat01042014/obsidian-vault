$ErrorActionPreference='Stop'
function Setp($o, $n, $v) { [void]$o.GetType().InvokeMember($n, [Reflection.BindingFlags]::SetProperty, $null, $o, [object[]]@($v)) }
$src='D:\ТАЗА\ШР ЦАКФ с 01.11.2026 (проект 2026-11-09).xlsx'
$dst='D:\ТАЗА\ШР ЦАКФ с 01.11.2026 (проект 2026-11-12).xlsx'
$log='D:\ТАЗА\_исходники\shr_tlencheva_madi_2026-11-12.log'; Remove-Item $log -ErrorAction SilentlyContinue
function L($m){ Add-Content $log ((Get-Date -Format 'HH:mm:ss')+' '+$m) -Encoding UTF8 }
if(Test-Path $dst){ L 'файл 12.11 уже есть, перезапись отменена'; throw 'exists' }
Copy-Item -LiteralPath $src -Destination $dst
Add-Type -Namespace W -Name U -MemberDefinition '[System.Runtime.InteropServices.DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr h, out uint p);'
$xl=New-Object -ComObject Excel.Application
$p=[uint32]0; [void][W.U]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd,[ref]$p); L "мой PID $p"
Setp $xl 'Visible' $false; Setp $xl 'DisplayAlerts' $false; Setp $xl 'AutomationSecurity' ([int]3); Setp $xl 'AskToUpdateLinks' $false
try{
 $wb=$xl.Workbooks.Open($dst,[int]0,$false)
 $ws=$wb.Worksheets.Item('ФОТ')
 # контроль исходных строк
 $b20=$ws.Range('B20').Text; $b33=$ws.Range('B33').Text
 L ("B20='"+$b20+"' B33='"+$b33+"' I20='"+$ws.Range('I20').Text+"' I33='"+$ws.Range('I33').Text+"'")
 if(($b20 -notlike '*Тленчиева*') -or ($b33 -notlike '*Нуржау*')){ throw 'строки не те' }
 $before=@($ws.Range('E50').Value2,$ws.Range('F50').Value2,$ws.Range('G50').Value2,$ws.Range('E24').Value2,$ws.Range('G24').Value2,$ws.Range('E18').Value2)
 L ('до: ИТОГО E,F,G = '+($before[0..2] -join ' / ')+'; ДПУ E,G = '+$before[3]+' / '+$before[4]+'; Деп.соц E = '+$before[5])

 # Тленчиева: приход в январе 2027, оформление как у приходящих в 2027 (форматы строки 27 - Арманов)
 [void]$ws.Range('B27:J27').Copy()
 [void]$ws.Range('B20:J20').PasteSpecial(-4122)     # xlPasteFormats
 try{ Setp $xl 'CutCopyMode' $false }catch{}
 $ws.Range('I20').Formula = '01.2027'
 Setp $ws.Range('I20') 'NumberFormat' '@'
 $ws.Range('I20').Value2 = '01.2027'

 # Мади: главный специалист с 09.11.2026, оклад 1 000 000 (форматы строки 26 - 2026 год)
 [void]$ws.Range('B26:J26').Copy()
 [void]$ws.Range('B33:J33').PasteSpecial(-4122)
 try{ Setp $xl 'CutCopyMode' $false }catch{}
 $ws.Range('B33').Value2 = 'Главный специалист (проектный менеджер) / Нуржау М.'
 $ws.Range('C33').Value2 = 'C'
 $ws.Range('E33').Value2 = 1000000
 $ws.Range('F33').Formula = '=IF(E33>0,E33-MIN(E33,4250000)*10%-MIN(E33,1700000)*2%-MAX(0,E33-MIN(E33,4250000)*10%-MIN(E33,1700000)*2%-129750)*10%,0)'
 $ws.Range('G33').Formula = '=D33*E33*12/1000'
 Setp $ws.Range('I33') 'NumberFormat' '@'
 $ws.Range('I33').Value2 = '09.11.2026'
 # примечание о приходе в 2027
 $ws.Range('B51').Value2 = '*приход в 2027 году выделен сиреневым (с 01.2027 и с 02.2027)'
 $xl.Calculate()
 $after=@($ws.Range('E50').Value2,$ws.Range('F50').Value2,$ws.Range('G50').Value2,$ws.Range('E24').Value2,$ws.Range('G24').Value2,$ws.Range('E18').Value2)
 L ('после: ИТОГО E,F,G = '+($after[0..2] -join ' / ')+'; ДПУ E,G = '+$after[3]+' / '+$after[4]+'; Деп.соц E = '+$after[5])
 # где ещё встречаются эти фамилии
 foreach($s in $wb.Worksheets){
   foreach($name in @('Тленчиева','Нуржау')){
     $f=$s.UsedRange.Find($name)
     if($f -ne $null){ $first=$f.Address($false,$false); $addr=@($first); $n=$s.UsedRange.FindNext($f); while($n -ne $null -and $n.Address($false,$false) -ne $first -and $addr.Count -lt 8){ $addr+=$n.Address($false,$false); $n=$s.UsedRange.FindNext($n) }; L ("лист '"+$s.Name+"' "+$name+": "+($addr -join ', ')) }
   }
 }
 $out=@()
 foreach($r in 18..33+50..54){ $l=@(); foreach($c in 'A','B','C','D','E','F','G','H','I'){ $l+=$ws.Range("$c$r").Text }; $out+=("$r | "+($l -join ' ; ')) }
 $out | Set-Content 'D:\ТАЗА\_исходники\shr_tlencheva_madi_check.txt' -Encoding UTF8
 $wb.Save(); L 'сохранено'; $wb.Close($false)
}catch{ L ('ОШИБКА: '+$_.Exception.Message+' @'+$_.InvocationInfo.ScriptLineNumber) }
finally{ $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
