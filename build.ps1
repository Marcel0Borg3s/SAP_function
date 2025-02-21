$exclude = @("venv", "SAPproject2.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "SAPproject2.zip" -Force