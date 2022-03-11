<#
=================================================================================================
Script para mover diretórios inteiros e manter as permissões iguais
Autor: Wanderlei Hüttel
wanderlei@huttel.com.br
Versão 1.0 - 11/03/2021
=================================================================================================
#>

$folder_list = Get-Content -Path "C:\Users\administrator\Desktop\folders_to_move.txt"
$source_dir      = "K:\source"
$destination_dir = "W:\destination"

foreach ($folder in $folder_list){
    "Fazendo backup das permissões do diretório de origem '$folder'..."
    icacls "$source_dir\${folder}" /save "$destination_dir\${folder}_ntfs_perms.txt" /t /c /q
    
    Write-Host "Movendo o diretório de origem '$folder' para o destino..."
    Move-Item -Path "$source_dir\${folder}" -Destination "$destination_dir"
    
    Write-Host "Restaurando permissões do diretório de origem '$folder' para o diretório de destino..."
    icacls "$destination_dir" /restore "$destination_dir\${folder}_ntfs_perms.txt" /t /c /q
    
	Remove-Item -Path "$destination_dir\${folder}_ntfs_perms.txt"
}
