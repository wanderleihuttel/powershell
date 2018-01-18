<#
=================================================================================================
Script para gerenciar importa��o e exporta��o de certificados digitais (NF-e, e-CNPJ, e-CPF, etc)
- Os certificados digitais devem conter a extens�o .p12 ou .pfx
- Os certificados digitais devem estar em um mesmo diret�rio
- Para importa��o em lote, as senhas dos certificados devem ser iguais

Autor: Wanderlei H�ttel
wanderlei@huttel.com.br
Vers�o 1.1 - 18/01/2018
=================================================================================================
#>

#############################################################################################
# CONFIGURA��ES 
# Diret�rio padr�o onde est�o os certificados digitais
$dir = "\\servidor_r2d2\documentos\wanderlei\Certificados Digitais"
# Diret�rio Desktop do usu�rio logado para exportar os certificados
$desktop = "$Env:USERPROFILE\Desktop"



#############################################################################################
# Fun��o para listar certificados
function ListarCertificados{
    Try{
        Get-ChildItem -Path Cert:\currentuser\My\ | Sort-Object NotAfter -descending | select  DnsNameList, NotAfter, Thumbprint | Format-Table @{Label="Nome";Expression={$nome = $_.DnsNameList -split ":" ; $nome[0];  }},  @{Label="CNPJ";Expression={$cnpj = $_.DnsNameList -split ":" ; $cnpj[1];  }},@{ Label="Validade"; Expression={$_.NotAfter} }, @{ Label="ID"; Expression={$_.ThumbPrint} }
        AguardarTeclado
    }
    Catch {}
}

#############################################################################################
# Fun��o para importar certificados
function ImportarCertificados{
     Try{
         Get-ChildItem -Path $dir | select name | Format-Wide
         Write-Host "Digite o nome completo do certificado: " -foregroundcolor "green" -NoNewline
         $nome_certificado = Read-Host
         
         if( -not (Test-Path "$dir/$nome_certificado" )){
             Write-Host "Arquivo $nome_certificado n�o existe! Opera��o abortada!" -foregroundcolor "yellow"
             AguardarTeclado
         }
         
         Write-Host "Digite a senha de instala��o do certificado: " -foregroundcolor "green" -NoNewline
         $password = Read-Host -asSecureString 
         Import-PfxCertificate �FilePath "$dir\$nome_certificado" -CertStoreLocation "Cert:\CurrentUser\My" -Password $password -Exportable  | Out-Null
         
         Write-Host "`nCertificado `"$nome_certificado`" importado com sucesso!`n"
         
         AguardarTeclado
     }
     Catch {
         Write-Host "`nOcorreu um erro na importa��o do certificado `"$nome_certificado`"!`n" -foregroundcolor "yellow"
         AguardarTeclado
     }
}

#############################################################################################
# Fun��o para exportar certificados
function ExportarCertificados{
    Try{
        Get-ChildItem -Path Cert:\currentuser\My\ | select  DnsNameList, NotAfter, Thumbprint | Format-Table @{Label="Nome";Expression={$nome = $_.DnsNameList -split ":" ; $nome[0];  }},  @{Label="CNPJ";Expression={$cnpj = $_.DnsNameList -split ":" ; $cnpj[1];  }},@{ Label="Validade"; Expression={$_.NotAfter} }, @{ Label="ID"; Expression={$_.ThumbPrint} }
        Write-Host "Digite o 6 primeiros caracters do ID do certificado: " -foregroundcolor "green" -NoNewline
        $thumbprint = Read-Host
        $thumbprint = (Get-ChildItem -Path "Cert:\CurrentUser\My\" | where {$_.Thumbprint -match "$thumbprint" } ).Thumbprint
        
        if( -not (Test-Path "Cert:\CurrentUser\My\$thumbprint" )){
            Write-Host "Arquivo $thumbprint n�o existe! Opera��o abortada!" -foregroundcolor "yellow"
            AguardarTeclado
        }
        $nome_certificado = (((Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint").DnsNameList -split ":")[0] -replace ' ', '_' -replace ':','_' ).ToLower()
        $data_vcto = ((Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint").NotAfter).ToString("dd.MM.yyyy")
        (gci "Cert:\CurrentUser\My\$thumbprint").FriendlyName = $nome_certificado
        
        Write-Host "`nCertificado selecionado: ""$nome_certificado""`n"
        
        Write-Host "Digite uma senha para a exporta��o do certificado: " -foregroundcolor "green" -NoNewline
        $password = Read-Host -asSecureString 
        Export-PfxCertificate -Cert $( Get-ChildItem -path "Cert:\CurrentUser\My\$thumbprint") -FilePath "$Env:USERPROFILE\Desktop\${nome_certificado}_vcto_${data_vcto}.pfx" -ChainOption EndEntityCertOnly -NoProperties -Password $password #| Out-Null
        
        if ($?){
            Write-Host "`nCertificado `"$nome_certificado`".pfx exportadado com sucesso!`n"
        }
        AguardarTeclado
    }
    Catch{
        Write-Host "`nOcorreu um erro na exporta��o do certificado `"$nome_certificado`"!`nO certificado pode n�o ser export�vel!`n" -foregroundcolor "yellow"
        AguardarTeclado
    }
}

#############################################################################################
# Fun��o para remover certificados
function RemoverCertificados{
    Try{
        Get-ChildItem -Path Cert:\currentuser\My\ | select  DnsNameList, NotAfter, Thumbprint | Format-Table @{Label="Nome";Expression={$nome = $_.DnsNameList -split ":" ; $nome[0];  }},  @{Label="CNPJ";Expression={$cnpj = $_.DnsNameList -split ":" ; $cnpj[1];  }},@{ Label="Validade"; Expression={$_.NotAfter} }, @{ Label="ID"; Expression={$_.ThumbPrint} }
        Write-Host "Digite o 6 primeiros caracters do ID do certificado: " -foregroundcolor "green" -NoNewline
        $thumbprint = Read-Host
        $thumbprint = (Get-ChildItem -Path "Cert:\CurrentUser\My\" | where {$_.Thumbprint -match "$thumbprint" } ).Thumbprint
        
        if( -not (Test-Path "Cert:\CurrentUser\My\$thumbprint" )){
            Write-Host "Arquivo $thumbprint n�o existe! Opera��o abortada!`n" -foregroundcolor "yellow"
            AguardarTeclado
        }
        $nome_certificado = ((Get-ChildItem -Path "Cert:\CurrentUser\My\$thumbprint").DnsNameList -split ":")[0] -replace ' ', '_' -replace ':','_'
        
        Write-Host "`nTem certeza que deseja remover o certificado `"$nome_certificado`" ?"
        $confirmar = Read-Host "S-Sim ou N-N�o"
        while("s","n" -notcontains $confirmar){
        	   $confirmar = Read-Host "S-Sim ou N-N�o"
        }
        
        if ($confirmar -eq "s" ){
            Remove-Item -Path "cert:\CurrentUser\My\$thumbprint"
            if ($?){
                Write-Host "`nO Certificado `"$nome_certificado`" foi removido com sucesso!`n"
            }
        }
        else {
            Write-Host "`nA remo��o do Certificado `"$nome_certificado`" foi cancelada pelo usu�rio!`n"
        }
        AguardarTeclado
    }
    Catch{}
}


#############################################################################################
# Fun��o para importar certificados
function ImportarCertificadosLote{
    Try{
     
        Write-Host "`nTem certeza que deseja importar todos os certificados?`nA senha deve ser a mesma para todos os arquivo!"
        $confirmar = Read-Host "S-Sim ou N-N�o"
        while("s","n" -notcontains $confirmar){
        	   $confirmar = Read-Host "S-Sim ou N-N�o"
        }
     
        if ($confirmar -eq "s" ){
            Write-Host "Digite a senha de instala��o do certificado: " -foregroundcolor "green" -NoNewline
            $password = Read-Host -asSecureString
            $certificados = Get-ChildItem -Path $dir -filter {*.pfx}
            foreach ($cert in $certificados){
                $thumbprint = $cert.thumbprint
                $nome_certificado = $cert.Name
                Import-PfxCertificate �FilePath "$dir\$cert" -CertStoreLocation "Cert:\CurrentUser\My" -Password $password -Exportable  | Out-Null
                Write-Host "Certificado $nome_certificado importado com sucesso!"
            }
           
        }
        else {
            Write-Host "`nA importa��o dos certificados foi cancelada pelo usu�rio!`n" -foregroundcolor "yellow"
        }
        AguardarTeclado
     
    }
    Catch {
        Write-Host "`nOcorreu um erro na importa��o do certificado `"$nome_certificado`"!`n" -foregroundcolor "yellow"
        AguardarTeclado
    }
}



#############################################################################################
# Fun��o para remover todos os certificados
function RemoverCertificadosLote{
    Try{
       
        Write-Host "`nTem certeza que deseja remover todos os certificados?"
        $confirmar = Read-Host "S-Sim ou N-N�o"
        while("s","n" -notcontains $confirmar){
            $confirmar = Read-Host "S-Sim ou N-N�o"
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
                Write-Host "`nN�o foram encontrados certificados para excluir!`n" -foregroundcolor "yellow"
            }
            
        }
        else {
            Write-Host "`nA remo��o dos certificados foi cancelada pelo usu�rio!`n" -foregroundcolor "yellow"
        }
        AguardarTeclado
    }
    Catch{}
}

#############################################################################################
# Fun��o Menu
function Menu() {
    Do {
      clear
      Write-Host "
    =========================================================================================
       Script para gerenciar a importa��o/exporta��o de certificados digitais 
       (NF-e, e-CNPJ, e-CPF, etc)
       - Os certificados digitais devem conter a extens�o .p12 ou .pfx
       - Os certificados digitais devem estar em um mesmo diret�rio
       - Para importa��o em lote, as senhas dos certificados devem ser iguais
       
    
       Autor: Wanderlei H�ttel
       wanderlei@huttel.com.br
       Vers�o 1.1 - 18/01/2018
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
      $opcao = Read-Host -prompt "  Selecione a op��o do Menu"
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
# Fun��o aguardar teclado
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
