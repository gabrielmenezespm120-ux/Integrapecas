# Envia as mudancas deste projeto para o GitHub.
# Roda sozinho no fim de cada resposta do Claude (hook Stop em ~/.claude/settings.json).
# Sem mudancas, nao faz nada. Log em %TEMP%\orcamento-sync.log.
$git = 'C:\Program Files\Git\cmd\git.exe'
if (-not (Test-Path $git)) { $git = 'git' }
Set-Location $PSScriptRoot
if (-not (Test-Path .git)) { exit 0 }
$log = Join-Path $env:TEMP 'orcamento-sync.log'
Set-Content (Join-Path $env:TEMP 'orcamento-sync.ultimo') (Get-Date -Format s)  # prova de que o hook rodou

if (& $git status --porcelain) {
    & $git add -A
    & $git commit -q -m ("Atualizacao automatica " + (Get-Date -Format 'dd/MM/yyyy HH:mm'))
}
# envia tambem commits que ficaram para tras (ex.: sem internet na vez anterior)
$pendentes = & $git rev-list --count '@{u}..HEAD' 2>$null
if ($LASTEXITCODE -ne 0 -or [int]$pendentes -gt 0) {
    $saida = cmd /c "`"$git`" push -u origin HEAD 2>&1"
    Add-Content $log ("[" + (Get-Date -Format s) + "] push exit $LASTEXITCODE`n" + ($saida -join "`n"))
}
exit 0
