# Сравнивает текущее содержимое папок с документами сотрудников ПО ТК со снимком прошлой проверки.
# Папки: D:\Кадры\Проекты ТД и Google Диск «Taza KZ HR документы». Снимок: D:\Кадры\_исходники\hr_snapshot.json
$ErrorActionPreference = 'Stop'
$roots = @(
  'D:\Кадры\Проекты ТД',
  'G:\.shortcut-targets-by-id\1Bn8_KyOWBRzVkKRjYtAPwSw5TZwwW4sZ\Taza KZ HR документы'
)
$snapFile = 'D:\Кадры\_исходники\hr_snapshot.json'

$now = [ordered]@{}
foreach ($r in $roots) {
  if (!(Test-Path -LiteralPath $r)) { Write-Output "ПАПКА НЕДОСТУПНА: $r"; continue }
  Get-ChildItem -LiteralPath $r -Recurse -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'desktop.ini' -and $_.Name -notlike '~$*' } | ForEach-Object {
    $now[$_.FullName] = '{0}|{1}' -f $_.Length, $_.LastWriteTimeUtc.ToString('yyyy-MM-ddTHH:mm:ss')
  }
}
$old = [ordered]@{}
if (Test-Path -LiteralPath $snapFile) {
  $j = Get-Content -LiteralPath $snapFile -Raw -Encoding UTF8 | ConvertFrom-Json
  foreach ($p in $j.PSObject.Properties) { $old[$p.Name] = [string]$p.Value }
}
$added = @($now.Keys | Where-Object { -not $old.Contains($_) })
$removed = @($old.Keys | Where-Object { -not $now.Contains($_) })
$changed = @($now.Keys | Where-Object { $old.Contains($_) -and $old[$_] -ne $now[$_] })
if (-not (Test-Path -LiteralPath $snapFile)) { Write-Output 'Снимок создан (первая проверка).' }
elseif ($added.Count + $removed.Count + $changed.Count -eq 0) { Write-Output 'БЕЗ ИЗМЕНЕНИЙ' }
else {
  Write-Output 'ЕСТЬ ИЗМЕНЕНИЯ:'
  $added | ForEach-Object { 'новый:     ' + $_ }
  $changed | ForEach-Object { 'изменён:   ' + $_ }
  $removed | ForEach-Object { 'удалён:    ' + $_ }
}
($now | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $snapFile -Encoding UTF8
Write-Output ('файлов в снимке: ' + $now.Count)
