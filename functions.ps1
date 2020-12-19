function send-mail-responsable {

###### CRYPT PASSWORD ########
        Param([string]$VM_Mail, $Destinataire)
        $PassKey = [byte]95,13,58,45,22,11,88,82,11,34,67,91,19,20,96,82
        $Password = Get-Content C:\Users\admin\Documents\PassKey.txt | Convertto-SecureString -Key $PassKey
        $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))

###### VARIABLES #####
        $MailAccount = " t.f.bts.sio@gmail.com"
        $From = " Service IT<t.f.bts.sio@gmail.com>"
        $To = $Destinataire
        $SMTPServer = "smtp.gmail.com"
        $SMTPPort = "587"
    
        # UTF8
        $encodingMail = [System.Text.Encoding]::UTF8

###### INSIDE OF THE MAIL ######

        $Subject = "MANAGER : Rapport de creation des VM"

        $Body = "Bonjour<br/>"
        $Body += "<p>Vous trouverez ci-dessous le dernier rapport de création des VM</p>"
        $Body += "<h2>Rapport d'action</h2><br>"
        $Body += $Report 



###### SEND MANAGER MAIL ######
        $Password = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $MailAccount,$Password


        Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $EmailCredential -Encoding $encodingMail

}



function send-mail-demandeur {

###### CRYPT PASSWORD ########
        Param([string]$VM_Mail, $Destinataire)
        $PassKey = [byte]95,13,58,45,22,11,88,82,11,34,67,91,19,20,96,82
        $Password = Get-Content C:\Users\admin\Documents\PassKey.txt | Convertto-SecureString -Key $PassKey
        $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))

###### VARIABLES #####
        $MailAccount = " t.f.bts.sio@gmail.com"
        $From = " Service IT<t.f.bts.sio@gmail.com>"
        $To = $Destinataire
        $SMTPServer = "smtp.gmail.com"
        $SMTPPort = "587"
    
        # UTF8
        $encodingMail = [System.Text.Encoding]::UTF8

###### INSIDE OF THE MAIL ######

        $Subject = "Demandeur : Rapport de creation des VM"

            $Body = "Bonjour<br/>"
            $Body += "<p>Nous avons traité votre demande de création de machine virtuelle $VM_Mail </p>" 
            $Body += "<p>Le status de création de votre VM est actuellement en $Status<br>$Reason<br></p>"
            $Body += "<p>En cas de problème avec la création de votre VM vous allez sous peu être contacté par le support.</p>"



###### SEND MANAGER MAIL ######
        $Password = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $MailAccount,$Password


        Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $EmailCredential -Encoding $encodingMail

}



function send-mail-technicien {

###### CRYPT PASSWORD ########
        Param([string]$VM_Mail, $Destinataire)
        $PassKey = [byte]95,13,58,45,22,11,88,82,11,34,67,91,19,20,96,82
        $Password = Get-Content C:\Users\admin\Documents\PassKey.txt | Convertto-SecureString -Key $PassKey
        $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))

###### VARIABLES #####
        $MailAccount = " t.f.bts.sio@gmail.com"
        $From = " Service IT<t.f.bts.sio@gmail.com>"
        $To = $Destinataire
        $SMTPServer = "smtp.gmail.com"
        $SMTPPort = "587"
    
        # UTF8
        $encodingMail = [System.Text.Encoding]::UTF8

###### INSIDE OF THE MAIL ######

        $Subject = "Technicien : Erreur lors de la création de la VM $VM_MAIL"

        $Body = "Bonjour<br/>"
        $Body += "<p>Un problème a été rencontré lors de la création de la VM $VM_Mail </p><br>"
        $Body += "<p>Message d'erreur :</p>"
        $Body += "<p>$Status<br>$Reason<br></p>"



###### SEND MAIL ######
        $Password = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $EmailCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $MailAccount,$Password


        Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -Port $SMTPPort -UseSsl -Credential $EmailCredential -Encoding $encodingMail

}




