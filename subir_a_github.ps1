param(
    [Parameter(Mandatory=$false)]
    [string]$RepoUrl
)

$git = "$env:LOCALAPPDATA\Programs\MinGit\cmd\git.exe"

if (-not $RepoUrl) {
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host " CONECTAR REPOSITORIO LOCAL CON TU CUENTA DE GITHUB" -ForegroundColor Yellow
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "1. Entra a tu navegador a: https://github.com/new"
    Write-Host "2. Crea un repositorio (ejemplo: tesis-clinica-barquero)."
    Write-Host "3. Copia el enlace HTTPS que te da GitHub."
    Write-Host ""
    $RepoUrl = Read-Host "Pega aqui el enlace de tu repositorio de GitHub"
}

if ($RepoUrl) {
    & $git remote remove origin -ErrorAction SilentlyContinue
    & $git remote add origin $RepoUrl
    & $git branch -M main
    Write-Host "`nSubiendo archivos a GitHub..." -ForegroundColor Green
    & $git push -u origin main
    Write-Host "`n¡Repositorio publicado con éxito en GitHub!" -ForegroundColor Cyan
} else {
    Write-Host "No se proporciono URL. Operacion cancelada." -ForegroundColor Red
}
