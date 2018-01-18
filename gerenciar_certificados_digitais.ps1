<#
=================================================================================================
Script para gerenciar importação e exportação de certificados digitais (NF-e, e-CNPJ, e-CPF, etc)
- Os certificados digitais devem conter a extensão .p12 ou .pfx
- Os certificados digitais devem estar em um mesmo diretório
- Para importação em lote, as senhas dos certificados devem ser iguais

Autor: Wanderlei Hüttel
wanderlei@huttel.com.br
Versão 1.1 - 18/01/2018
=================================================================================================
#>

#############################################################################################
# CONFIGURAÇÕES 
# Diretório padrão onde estão os certificados digitais
$dir = "\\servidor_r2d2\documentos\wanderlei\Certificados Digitais"
# Diretório Desktop do usuário logado para exportar os certificados
$desktop = "$Env:USERPROFILE\Desktop"



#############################################################################################
# Função para listar certificados
function ListarCertificados{
    Try{
        Get-ChildItem -Path Cert:\currentuser\My\ | Sort-Object NotAfter -descending | select  DnsNameList, NotAfter, Thumbprint | Format-Table @{Label="Nome";Expression={$nome = $_.DnsNameList -split ":" ; $nome[0];  }},  @{Label="CNPJ";Expression={$cnpj = $_.DnsNameList -split ":" ; $cnpj[1];  }},@{ Label="Validade"; Expression={$_.NotAfter} }, @{ Label="ID"; Expression={$_.ThumbPrint} }
        AguardarTeclado
    }
    Catch {}
}

#############################################################################################
# Função para importar certificados
function ImportarCertificados{
     Try{
         Get-ChildItem -Path $dir | select name | Format-Wide
         Write-Host "Digite o nome completo do certificado: " -foregroundcolor "green" -NoNewline
         $nome_certificado = Read-Host
         
         if( -not (Test-Path "$dir/$nome_certificado" )){
             Write-Host "Arquivo $nome_certificado não existe! Operação abortada!" -foregroundcolor "yellow"
             AguardarTeclado
         }
         
         Write-Host "Digite a senha de instalação do certificado: " -foregroundcolor "green" -NoNewline
         $password = Read-Host -asSecureString 
         Import-PfxCertificate –FilePath "$dir\$nome_certificado" -CertStoreLocation "Cert:\CurrentUser\My" -Password $password -Exportable  | Out-Null
         
         Write-Host "`nCertificado `"$nome_certificado`" importado com sucesso!`n"
         
         AguardarTeclado
     }
     Catch {
         Write-Host "`nOcorreu um erro na importação do certificado `"$nome_certificado`"!`n" -foregroundcolor "yellow"
         AguardarTeclado
     }
}

#############################################################################################
# Função para exportar certificados
function ExportarCertificados{
    Try{
        Get-ChildItem -Path Cert:\currentuser\My\ | select  DnsNameList, NotAfter, Thumbprint | Format-Table @{Label="Nome";Expression={$nome = $_.DnsNameList -split ":" ; $nome[0];  }},  @{Label="CNPJ";Expression={$cnpj = $_.DnsNameList -split ":" ; $cnpj[1];  }},@{ Label="Validade"; Expression={$_.NotAfter} }, @{ Label="ID"; Expression={$_.ThumbPrint} }
        Write-Host "Digite o 6 primeiros caracters do ID do certificado: " -foregroundcolor "green" -NoNewline
        $thumbprint = Read-Host
        $thumbprint = (Get-ChildItem -Path "Cert:\CurrentUser\My\" | where {$_.Thumbprint -match "$thumbprint" } ).Thumbprint
        
        if( -not (Test-Path "Cert:\CurrentUser\My\$thumbprint" )){
            Write-Host "Arquivo $thumbprint não existe! Operação abortada!" -foregroundcolor "yellow"
            AguardarTeclado
        }
        $nome_certificado = (((Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint").DnsNameList -split ":")[0] -replace ' ', '_' -replace ':','_' ).ToLower()
        $data_vcto = ((Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint").NotAfter).ToString("dd.MM.yyyy")
        (gci "Cert:\CurrentUser\My\$thumbprint").FriendlyName = $nome_certificado
        
        Write-Host "`nCertificado selecionado: ""$nome_certificado""`n"
        
        Write-Host "Digite uma senha para a exportação do certificado: " -foregroundcolor "green" -NoNewline
        $password = Read-Host -asSecureString 
        Export-PfxCertificate -Cert $( Get-ChildItem -path "Cert:\CurrentUser\My\$thumbprint") -FilePath "$Env:USERPROFILE\Desktop\${nome_certificado}_vcto_${data_vcto}.pfx" -ChainOption EndEntityCertOnly -NoProperties -Password $password #| Out-Null
        
        if ($?){
            Write-Host "`nCertificado `"$nome_certificado`".pfx exportadado com sucesso!`n"
        }
        AguardarTeclado
    }
    Catch{
        Write-Host "`nOcorreu um erro na exportação do certificado `"$nome_certificado`"!`nO certificado pode não ser exportável!`n" -foregroundcolor "yellow"
        AguardarTeclado
    }
}

#############################################################################################
# Função para remover certificados
function RemoverCertificados{
    Try{
        Get-ChildItem -Path Cert:\currentuser\My\ | select  DnsNameList, NotAfter, Thumbprint | Format-Table @{Label="Nome";Expression={$nome = $_.DnsNameList -split ":" ; $nome[0];  }},  @{Label="CNPJ";Expression={$cnpj = $_.DnsNameList -split ":" ; $cnpj[1];  }},@{ Label="Validade"; Expression={$_.NotAfter} }, @{ Label="ID"; Expression={$_.ThumbPrint} }
        Write-Host "Digite o 6 primeiros caracters do ID do certificado: " -foregroundcolor "green" -NoNewline
        $thumbprint = Read-Host
        $thumbprint = (Get-ChildItem -Path "Cert:\CurrentUser\My\" | where {$_.Thumbprint -match "$thumbprint" } ).Thumbprint
        
        if( -not (Test-Path "Cert:\CurrentUser\My\$thumbprint" )){
            Write-Host "Arquivo $thumbprint não existe! Operação abortada!`n" -foregroundcolor "yellow"
            AguardarTeclado
        }
        $nome_certificado = ((Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint").DnsNameList -split ":")[0] -replace ' ', '_' -replace ':','_'
        
        Write-Host "`nTem certeza que deseja remover o certificado `"$nome_certificado`" ?"
        $confirmar = Read-Host "S-Sim ou N-Não"
        while("s","n" -notcontains $confirmar){
        	   $confirmar = Read-Host "S-Sim ou N-Não"
        }
        
        if ($confirmar -eq "s" ){
            Remove-Item -Path "cert:\CurrentUser\My\$thumbprint"
            if ($?){
                Write-Host "`nO Certificado `"$nome_certificado`" foi removido com sucesso!`n"
            }
        }
        else {
            Write-Host "`nA remoção do Certificado `"$nome_certificado`" foi cancelada pelo usuário!`n"
        }
        AguardarTeclado
    }
    Catch{}
}


#############################################################################################
# Função para importar certificados
function ImportarCertificadosLote{
    Try{
     
        Write-Host "`nTem certeza que deseja importar todos os certificados?`nA senha deve ser a mesma para todos os arquivo!"
        $confirmar = Read-Host "S-Sim ou N-Não"
        while("s","n" -notcontains $confirmar){
        	   $confirmar = Read-Host "S-Sim ou N-Não"
        }
     
        if ($confirmar -eq "s" ){
            Write-Host "Digite a senha de instalação do certificado: " -foregroundcolor "green" -NoNewline
            $password = Read-Host -asSecureString
            $certificados = Get-ChildItem -Path $dir -filter {*.pfx}
            foreach ($cert in $certificados){
                $thumbprint = $cert.thumbprint
                $nome_certificado = $cert.Name
                Import-PfxCertificate –FilePath "$dir\$cert" -CertStoreLocation "Cert:\CurrentUser\My" -Password $password -Exportable  | Out-Null
                Write-Host "Certificado $nome_certificado importado com sucesso!"
            }
           
        }
        else {
            Write-Host "`nA importação dos certificados foi cancelada pelo usuário!`n" -foregroundcolor "yellow"
        }
        AguardarTeclado
     
    }
    Catch {
        Write-Host "`nOcorreu um erro na importação do certificado `"$nome_certificado`"!`n" -foregroundcolor "yellow"
        AguardarTeclado
    }
}



#############################################################################################
# Função para remover todos os certificados
function RemoverCertificadosLote{
    Try{
       
        Write-Host "`nTem certeza que deseja remover todos os certificados?"
        $confirmar = Read-Host "S-Sim ou N-Não"
        while("s","n" -notcontains $confirmar){
            $confirmar = Read-Host "S-Sim ou N-Não"
        }
        if ($confirmar -eq "s" ){
            $certificados = Get-ChildItem -Path "Cert:\currentuser\My\"
         
            if( $certificados.count -gt 0){
               
               foreach ($cert in $certificados){
                   $thumbprint = $cert.thumbprint
                   $nome_certificado = ((Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint").DnsNameList -split ":")[0] -replace ' ', '_' -replace ':','_'
                   Remove-Item -Path "cert:\CurrentUser\My\$thumbprint"
                   if ($?){
                       Write-Host "O certificado `"$nome_certificado`" foi removido com sucesso!"
                   }
                   else {
                       Write-Host "Ocorreu um erro ao remover o certificado `"$nome_certificado`"!"
                   }
               } #endforach
            } else {
                Write-Host "`nNão foram encontrados certificados para excluir!`n" -foregroundcolor "yellow"
            }
            
        }
        else {
            Write-Host "`nA remoção dos certificados foi cancelada pelo usuário!`n" -foregroundcolor "yellow"
        }
        AguardarTeclado
    }
    Catch{}
}

#############################################################################################
# Função Menu
function Menu() {
    Do {
      clear
      Write-Host "
    =========================================================================================
       Script para gerenciar a importação/exportação de certificados digitais 
       (NF-e, e-CNPJ, e-CPF, etc)
       - Os certificados digitais devem conter a extensão .p12 ou .pfx
       - Os certificados digitais devem estar em um mesmo diretório
       - Para importação em lote, as senhas dos certificados devem ser iguais
       
    
       Autor: Wanderlei Hüttel
       wanderlei@huttel.com.br
       Versão 1.1 - 18/01/2018
    =========================================================================================
      
      1 = Listar certificados instalados
      
      2 = Importar certificado
      
      3 = Exportar certificado
      
      4 = Remover certificado
      
      5 = Importar todos os certificados da pasta
      
      6 = Remover todos os certificados
      
      q = Sair
      --------------------------
      " -foregroundcolor "green"
      $opcao = Read-Host -prompt "  Selecione a opção do Menu"
    } until ($opcao -eq "1" -or $opcao -eq "2" -or $opcao -eq "3" -or $opcao -eq "4" -or $opcao -eq "5" -or $opcao -eq "6" -or $opcao -eq "q")
    
    Switch ($opcao){
      "1" { ListarCertificados; break; }
      "2" { ImportarCertificados;break; }
      "3" { ExportarCertificados; break; }
      "4" { RemoverCertificados; break; }
      "5" { ImportarCertificadosLote;break; }
      "6" { RemoverCertificadosLote; break; }
      "q" { exit 0; }
    }
}

#############################################################################################
# Função aguardar teclado
function AguardarTeclado(){
    Try{
        Write-Host "Pressione qualquer  tecla para continuar ..."
        $r = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        Menu
    }
    Catch{}
}



#############################################################################################
# Executa o menu
Menu
