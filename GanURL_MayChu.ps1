# TAPACO Nhan su - Gan URL may chu (Ung dung web Apps Script) tu tep webapp.txt vao index.html
# Dung khi mo app truc tiep tu o dia (file://): trinh duyet khong cho app tu doc webapp.txt.
# Chay: bam dup GanURL_MayChu.bat (cung thu muc). Tep index.html va webapp.txt nam trong ..\1_UngDung
param([string]$ThuMuc = '')
$ErrorActionPreference = 'Stop'
$goc = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($ThuMuc -eq '') { $ThuMuc = Join-Path (Split-Path -Parent $goc) '1_UngDung' }
$txt = Join-Path $ThuMuc 'webapp.txt'
$html = Join-Path $ThuMuc 'index.html'
if (-not (Test-Path $txt)) { Write-Host "Khong tim thay $txt"; exit 1 }
if (-not (Test-Path $html)) { Write-Host "Khong tim thay $html"; exit 1 }
$re = 'https://script\.google\.com/(?:a/macros/[^/\s]+|macros)/s/[\w-]+/exec'
$url = $null
foreach ($dong in [System.IO.File]::ReadAllLines($txt)) {
  $t = $dong.Trim([char]0xFEFF, ' ', "`t")
  if ($t -eq '' -or $t.StartsWith('#') -or $t.StartsWith('//')) { continue }
  $m = [regex]::Match($t, $re)
  if ($m.Success) { $url = $m.Value; break }
}
if (-not $url) { Write-Host 'webapp.txt chua co URL hop le (https://script.google.com/.../exec tren mot dong rieng, khong co dau # o dau).'; exit 1 }
$enc = New-Object System.Text.UTF8Encoding($false)
$s = [System.IO.File]::ReadAllText($html, $enc)
$mau = [regex]"const DEFAULT_API = '[^']*';"
if (-not $mau.IsMatch($s)) { Write-Host 'Khong tim thay dong "const DEFAULT_API" trong index.html (sai phien ban?).'; exit 1 }
$moi = $mau.Replace($s, "const DEFAULT_API = '" + $url + "';", 1)
if ($moi -eq $s) { Write-Host "index.html da gan san URL nay: $url"; exit 0 }
Copy-Item $html ($html + '.bak') -Force
[System.IO.File]::WriteAllText($html, $moi, $enc)
Write-Host "Da gan URL may chu vao index.html:"
Write-Host "  $url"
Write-Host "Ban sao cu: index.html.bak. Mo lai index.html (Ctrl+F5) - app tu ket noi may chu."
exit 0
