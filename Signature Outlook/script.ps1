$localpath= get-content env:APPDATA # definition du répertoire local dans lequel doivent être crées les signatures
$remotepath='.\Template' # path pour accéder au partage qui contient les sources html et les images

robocopy /e /r:3 /w:1 $remotepath $localpath\Microsoft\Signatures


$SysInfo = New-Object -ComObject "ADSystemInfo"
$ADpath2=$SysInfo.GetType().InvokeMember("username", "GetProperty", $Null, $SysInfo, $Null)

# recuperation des informations personnelles
$Info=[ADSI] "LDAP: / / $ADpath2"
$name=$info.displayname 
$mail=$info.mail
$mobile=$info.mobile 
$title=$info.title # récupération de l'intitulé de poste
$PhoneNumber=$info.telephonenumber # récupération du téléphone fixe
$adresse=$info.streetaddress # récupération de l'adresse postale
$postalCode=$info.postalcode # récupération du code postal
$City=$info.l # récupération de la ville
$FullAddress= $adresse+"-"+$postalCode+$City # mise en forme de type Adresse-Ville
$Service= $info.department

function PersonnalisationHTML ($file)
 { 
    $Perso=Get-Content $file
    $Perso = $Perso -replace "%name%","$name" # remplacement de %name% par le displayname d'AD
    $Perso = $Perso -replace "%mail%","$mail"
    $perso=$perso -replace '%pager%',"$PhoneNumber"
    $perso=$perso  -replace '%mobile%',"$mobile"
    $Perso = $Perso -replace "%title%","$title" #personnalisation des intitulés de poste
    $Perso = $Perso -replace "%FullAddress%","$FullAddress"
    $Perso = $Perso -replace "%Service%","$Service"

    set-content $file $perso #on ecrase le contenu du template par le contenu personnalisé
}

#Personnalisation TXT
# Même principe que la personnalisation html
 function PersonnalisationTXT ($file)
 { 
    $Perso=Get-Content $file
    $Perso = $Perso -replace "%name%","$name"
    $Perso = $Perso -replace "%mail%","$mail"
    $perso=$perso -replace "%pager%","$PhoneNumber"
    $perso=$perso  -replace "%mobile%","$mobile"
    $Perso = $Perso -replace "%title%","$title"
    $Perso = $Perso -replace "%FullAddress%","$FullAddress"
    $Perso = $Perso -replace "%Service%","$Service"

    set-content $file $perso
}

# Personnalisation des fichiers
PersonnalisationHTML ($localpath+'\Microsoft\Signatures\GIE.htm')
PersonnalisationTXT ($localpath+'\Microsoft\Signatures\GIE.txt')


# sortie de script
exit