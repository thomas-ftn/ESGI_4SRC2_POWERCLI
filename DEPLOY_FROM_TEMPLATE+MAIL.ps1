clear
$Report = @()

#IMPORT DES FONCTIONS
. $PSScriptRoot\functions.ps1

        Write-Host "`n     Script de création des VM`n`n"
        Start-Sleep -Seconds 1
        Write-Host "`n Saisissez votre mot de passe Vcenter `n"-BackgroundColor black
# PROMPT FOR PASSWORD
$pass_secure = Read-Host -AsSecureString  "Mot de passe" 
$pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pass_secure))

        Write-Host "`n Connexion au Vcenter`n" -ForegroundColor Yellow -BackgroundColor black
        Start-Sleep -Seconds 1
# CONNECT TO VCENTER
$ErrorActionPreference = "Stop"
try {
        Connect-VIServer -Server 172.180.0.50 -User Administrator@vsphere.local -Password $pass -ErrorAction stop | Out-Null
    }
catch {
    write-host "`n Echec de la connexion, vérifez votre mot de passe et relancez le script `n" -ForegroundColor red -BackgroundColor black
    break
}

write-host "`n Conexion au VCenter réussie `n"  -ForegroundColor green -BackgroundColor black
Start-Sleep -Seconds 1

Write-Host "`n Import de la liste des machines à créer depuis le fichier CSV `n`n"  -ForegroundColor Yellow -BackgroundColor black
Start-Sleep -Seconds 1


# CSV IMPORT
$Vms = Import-Csv C:\Users\admin\Documents\VMS.csv -Delimiter ';'
Start-Sleep -Seconds 1

Write-Output $Vms
Write-Host "`n ################################################"
Write-Host "`n"
# CONFIRM VIRTUAL MACHINE CREATION
Start-Sleep -Seconds 1
$confirmation = Read-Host "Confirmez vous la création des VM précedement citées ? (OUI pour confirmer)"
$confirmation = $confirmation.ToUpper()
Write-Host "`n"
if ( $confirmation -ne "OUI" ){
            Write-Host  "`n Abandon du déploiement `n" -ForegroundColor red -BackgroundColor black
}

else
{

    foreach ($Vm in $Vms) {
        
        $VmName = $Vm.Name
        $VmPool = $Vm.Pool
        $VmTemplate = $Vm.Template
        $VmDatastore = $Vm.Datastore
        $VmCustom = $Vm.Custom
        $Destinataire = $Vm.Destinataire
     
 Write-Host "`n TRAITEMENT DE LA DEMANDE DE VM  $VmName `n"  -ForegroundColor Yellow -BackgroundColor black

        if (!(get-vm $VmName -erroraction 0)){
                
                           Write-Host "`n Déploiement de la VM $VmName `n"  -ForegroundColor Yellow -BackgroundColor black

           
                               New-vm -ResourcePool $VmPool -Name $VmName -Template $VmTemplate -Datastore $VmDatastore -DrsAutomationLevel AsSpecifiedByCluster -OSCustomizationspec $VmCustom -erroraction 0 | Out-Null
                               
                               if ((get-vm $VmName -erroraction 0)){
                               

                                      Write-Host "`n Déploiement de $VmName terminé `n"  -ForegroundColor green -BackgroundColor black
                                      Start-Sleep -Seconds 1
        
                                  

                                                         Write-Host "Démmarage de la VM $VmName" -ForegroundColor yellow -BackgroundColor black
                                                           
                                                         Start-VM -VM $VmName -Confirm:$false -erroraction 0 | Out-Null
                                                         
                                                             if ((get-vm $VmName |Where-object {$_.powerstate -eq "poweredon"})){
                                                         
                                                              Write-Host "`n La VM $VmName est démarée `n" -ForegroundColor green -BackgroundColor black
                                                              Start-Sleep -Seconds 1
                                                              $Status = "REUSSI"
                                                              $Reason = "Déploiement et démarrage réussi"

                                                              Write-Host "`n Envoi du mail de notification de résussite du déploiement `n" -ForegroundColor Yellow -BackgroundColor black
                                                             }

                                                             else {
                                                          
                                                              Write-Host "`n La VM $VmName n'a pas démarré `n" -ForegroundColor red -BackgroundColor black
                                                              Start-Sleep -Seconds 1
                                                              $Status = "ECHEC"
                                                              $Reason = "Déploiement réussi, le démrrage n'a pas abouti. Une intervention est nécéssaire pour lancer la customisation de la machine"
                                                         
                                                              Write-Host "`n Envoi du mail de notification d'erreur de démmarage `n" -ForegroundColor Yellow -BackgroundColor black
                                                              send-mail-technicien -VM_Mail $VmName -Destinataire $Destinataire -Status $Status -Reason $Reason
                                                             }
            
                               }

                                 else {
                                  Write-Host "`n Echec de la création de la VM  $VmName `n" -ForegroundColor red -BackgroundColor black
                                  Start-Sleep -Seconds 1
                                  $Status = "ECHEC"
                                  $Reason = "Echec du déploiement de VM $VmName"
                                  Write-Host "`n Envoi du mail de notification d'échec du déploiement de VM $VmName`n" -ForegroundColor Yellow -BackgroundColor black
                                  send-mail-technicien -VM_Mail $VmName -Destinataire $Destinataire -Status $Status -Reason $Reason
                                 }
      
        } # END IF VM EXISTS

                   else {
             
                            Write-Host  "`n Le nom de VM $VmName est déja pris. Elle ne sera pas créée." -ForegroundColor red -BackgroundColor black
                            Start-Sleep -Seconds 1
                            $Status = "ECHEC"
                            $Reason = "Le nom de la VM est déja pris"
                            Write-Host "`n Envoi du mail de notification d'échec du déploiement de VM $VmName`n" -ForegroundColor Yellow -BackgroundColor black
                            send-mail-technicien -VM_Mail $VmName -Destinataire $Destinataire -Status $Status -Reason $Reason

                            Start-Sleep -Seconds 1 
                   
                   } # END ELSEIF VM EXISTS
                          
                          # SEND NOTIFICATION MAIL WITH STATUS AND RESON VALUES
                         send-mail-demandeur -VM_Mail $VmName -Destinataire $Destinataire -Status $Status -Reason $Reason


                        

                           # ADD LOGS TO REPORT TABLE 

                         
                           $object = New-Object -TypeName PSObject
                           $object | Add-Member -Name 'ServerName' -MemberType Noteproperty -Value $VmName
                           $object | Add-Member -Name 'Status' -MemberType Noteproperty -Value $Status
                           $object | Add-Member -Name 'Reason' -MemberType Noteproperty -Value $Reason
                           $Report += $object
                                   
       }

##### REPORT CREATION 
###HEADER CREATION

$header = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="fr" xml:lang="fr">
<head>
<title>System Status Report</title>
<style type="text/css">
<!--
body {
background-color: #E0E0E0;
font-family: sans-serif
}
table, th, td {
background-color: white;
border-collapse:collapse;
border: 1px solid black;
padding: 5px
}
-->
</style>
"@

####CONVERTING TO HTML AND FILTERING VALUES TO ADD COLOR

    $Report = $Report | ConvertTo-Html -Property ServerName,Status,Reason -Head $header | foreach {
    $PSItem -replace "<td>ECHEC</td>", "<td style='background-color:#FF8080'>ECHEC</td>" -replace "<td>REUSSI</td>", "<td style='background-color:#32CD32'>REUSSI</td>"
}



               Write-Host "`n Envoi du compte rendu des opérations au responsable `n"  -ForegroundColor Blue -BackgroundColor black  
           send-mail-responsable -VM_Mail $VmName -Destinataire $Destinataire -Report $Report

     Write-Host "`n ##### Fin du traitement ##### `n"  -ForegroundColor green -BackgroundColor black 
        }


Disconnect-VIServer * -Confirm:$false 