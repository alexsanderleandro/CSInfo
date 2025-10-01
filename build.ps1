<#
build.ps1 - versão limpa
Gera assets/app.ico a partir de assets/ico.png (se necessário), executa PyInstaller e cria um ZIP com timestamp.
Uso: .\build.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $root

Write-Host "Build iniciado em: $root"

$png = Join-Path $root 'assets\ico.png'
$ico = Join-Path $root 'assets\app.ico'

if (Test-Path $png) {
    $regen = -not (Test-Path $ico) -or ((Get-Item $png).LastWriteTime -gt (Get-Item $ico).LastWriteTime)
    if ($regen) {
        Write-Host "Gerando $ico a partir de $png ..."
        $pyfile = Join-Path $env:TEMP 'csinfo_gen_icon.py'
        $lines = @(
            'from PIL import Image',
            "png = r'{0}'" -f $png,
            "ico = r'{0}'" -f $ico,
            'img = Image.open(png)',
            "if img.mode not in ('RGBA','RGB'):",
            "    img = img.convert('RGBA')",
            "sizes = [(16,16),(32,32),(48,48),(256,256)]",
            "img.save(ico, format='ICO', sizes=sizes)",
            "print('Ícone gerado:', ico)"
        )
        $lines | Out-File -FilePath $pyfile -Encoding utf8
        try {
            python $pyfile
        } catch {
            Write-Host "Aviso: falha ao gerar ícone com Python: $_"
        }
        Remove-Item $pyfile -ErrorAction SilentlyContinue
    } else {
        Write-Host "app.ico está atualizado. Pulando geração."
    }
} else {
    Write-Host "Aviso: PNG não encontrado em $png. Verifique assets/ico.png"
}

Write-Host "Limpando builds antigos..."
if (Test-Path build) { Remove-Item -Recurse -Force build }
if (Test-Path dist) { Remove-Item -Recurse -Force dist }
if (Test-Path __pycache__) { Remove-Item -Recurse -Force __pycache__ }

Write-Host "Executando PyInstaller..."
pyinstaller --noconfirm --onefile --windowed --name CSInfo --icon assets/app.ico --add-data "assets;assets" --paths . csinfo_gui.py

if (Test-Path (Join-Path $root 'dist\CSInfo.exe')) {
    Write-Host "Executável gerado: dist\CSInfo.exe"
    $zipRoot = Join-Path $root 'dist'
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $zipPath = Join-Path $root "dist\CSInfo_$ts.zip"
    try {
        # incluir tudo dentro de dist no ZIP
        Compress-Archive -Path (Join-Path $zipRoot '*') -DestinationPath $zipPath -Force
        Write-Host "ZIP criado: $zipPath"
    } catch {
        Write-Host "Erro ao criar ZIP via Compress-Archive: $_"
    }
} else {
    Write-Host "Erro: Executável não encontrado em dist\CSInfo.exe"
}

Write-Host "Build finalizado."