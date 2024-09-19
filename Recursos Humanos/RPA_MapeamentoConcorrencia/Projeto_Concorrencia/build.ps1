$exclude = @("venv", "Projeto_Concorrencia.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "Projeto_Concorrencia.zip" -Force