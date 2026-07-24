$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$repo = "mksmkss/Display"
$installDir = "$env:LOCALAPPDATA\Display"
$zipPath = "$env:TEMP\Display-Windows.zip"

Write-Host "最新版の情報を取得しています..."
$release = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest"
$asset = $release.assets | Where-Object { $_.name -eq "Display-Windows.zip" }
if (-not $asset) {
    throw "Display-Windows.zip が見つかりませんでした。まだリリースが公開されていない可能性があります。"
}

Write-Host "ダウンロード中... ($($release.tag_name))"
Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath

Write-Host "インストール中..."
if (Test-Path $installDir) {
    Remove-Item -Recurse -Force $installDir
}
Expand-Archive -Path $zipPath -DestinationPath $installDir -Force
Remove-Item $zipPath

$exePath = Join-Path $installDir "Display\Display.exe"

$WshShell = New-Object -ComObject WScript.Shell
$shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\Display.lnk")
$shortcut.TargetPath = $exePath
$shortcut.WorkingDirectory = Split-Path $exePath
$shortcut.Save()

Write-Host ""
Write-Host "インストール完了！デスクトップの「Display」アイコンから起動できます。"
