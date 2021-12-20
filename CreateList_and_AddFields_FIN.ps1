Add-PSSnapin Microsoft.SharePoint.PowerShell
 
#Fonction pour creer une nouvelle liste
Function CreateList($SiteURL, $ListName)
{
     Try {
        $Web = Get-SPWeb -Identity $SiteURL
        $ListTemplate = [Microsoft.SharePoint.SPListTemplateType]::GenericList
        
        #Cherche si la liste existe déjà si elle n'existe pas je la créer
        if($Web.Lists.TryGetList($ListName) -eq $null)
        {
            $List = $Web.Lists.Add($ListName, $ListName, $ListTemplate) 
            write-host "Liste ${Listname} créée avec succès !" -ForegroundColor Green
        }
        #Si elle existe déjà j'affiche un message d'erreur
        else
        {
            write-host "Il existe déjà une liste ayant le même nom" -ForegroundColor Red
        }
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}


#Fonction pour créer une colonne
Function AddFieldToList($SiteURL,$ListName, $FieldName, $FieldType, $IsRequired)
{
    Try{
        #Récupere la liste
        $List = (Get-SPWeb $SiteURL).Lists.TryGetList($ListName)
         
        #Cherche si la liste existe si elle existe alors je cherche si la colonne n'existe pas
        if($List -ne $null)
        {
            if(!$List.Fields.ContainsField($FieldName))
            {     
                #Ajoute la colonne dans la liste
                $List.Fields.Add($FieldName,$FieldType,$IsRequired)
 
                #Met à jour la liste
                $List.Update()
 
                #Met à jour la vue par défaut pour afficher les nouvelles colonnes
                $View = $List.DefaultView
                $View.ViewFields.Add($FieldName)
                $View.Update()
 
                write-host "La nouvelle colonne ${FieldName} a été ajoutée à la liste ${ListName}" -ForegroundColor Green
            }
            else
            {
                write-host "la colonne ${FieldName} existe déjà dans la liste ${ListName}" -ForegroundColor Yellow
            }
        }
        else
        {
            write-host "La liste ${ListName} n'existe pas" -ForegroundColor Red
        }       
    }
     catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}
 
#Paramètres
$SiteURL="http://inside.lamotte.fr/doc-opm"

$TablListNantes=@('NANTES-33', 'NANTES-44-BasseGoulaine', 'NANTES-44-Blain', 'NANTES-44-Carquefou', 'NANTES-44-Chateaubriant', 'NANTES-44-LaBaule', 'NANTES-44-LaChapelle', 'NANTES-44-LaMontagne', 
'NANTES-44-Nantes-A_to_I', 'NANTES-44-Nantes-Le_C_to_Les_R', 'NANTES-44-Nantes-Les_T_to_V', 'NANTES-44-NortSurErdre', 'NANTES-44-Orvault', 'NANTES-44-Pornic', 'NANTES-44-Pornichet', 'NANTES-44-Reze', 
'NANTES-44-SainteLuce', 'NANTES-44-SaintHerblain', 'NANTES-44-SaintSebastien', 'NANTES-44-Sautron', 'NANTES-44-StEtienne', 'NANTES-49', 'NANTES-56-Baden', 'NANTES-56-Guidel', 'NANTES-56-Lorient', 'NANTES-56-Queven',
'NANTES-56-Riantec', 'NANTES-56-SaintPierre', 'NANTES-56-Vannes', 'NANTES-78', 'NANTES-94')

#$TablListRennes=@('RENNES-14', 'RENNES-17', 'RENNES-22', 'RENNES-29', 'RENNES-35-BainDeBretagne', 'RENNES-35-Betton', 'RENNES-35-BoisGervilly', 'RENNES-35-Bruz', 'RENNES-35-Cesson', 'RENNES-35-Chantepie',
#'RENNES-35-Domloup', 'RENNES-35-Etrelle', 'RENNES-35-Gosne', 'RENNES-35-Janze', 'RENNES-35-LaChapelle', 'RENNES-35-LeRheu', 'RENNES-35-Lhermitage', 'RENNES-35-Liffre', 'RENNES-35-MezieresSousC',
#'RENNES-35-MontaubanDeBr', 'RENNES-35-Noyal', 'RENNES-35-Parame', 'RENNES-35-Rennes-A_to_Le', 'RENNES-35-Rennes-Les_J_to_Les_O', 'RENNES-35-Rennes-Les_P_to_R', 'RENNES-35-Rennes-T_to_V', 'RENNES-35-Romagne',
#'RENNES-35-SaintBriac', 'RENNES-35-SaintGregoire-A_to_E', 'RENNES-35-SaintGregoire-L_to_P', 'RENNES-35-SaintJacques', 'RENNES-35-SaintMalo', 'RENNES-44Bur', 'RENNES-53', 'RENNES-56Bur', 'RENNES-72',
#'RENNES-74', 'RENNES-79', 'RENNES-85')

#$TablListSacib =@('SACIB-01-A_to_B', 'SACIB-01-C_to_Fl', 'SACIB-01-Fo_to_L', 'SACIB-01-N_to_R', 'SACIB-01-S_to_V', 'SACIB-02-A_to_Di', 'SACIB-02-Do_to_R', 'SACIB-02-S_to_V', 'SACIB-04', 'SACIB-05')
 
#Appel de la fonction afin de créer les listes pour Nantes
for($i=0; $i -le ($TablListNantes.length -1); $i++) {
CreateList $SiteURL $TablListNantes[$i]
}

#Appel de la fonction afin de créer les listes pour Rennes
#for($i=0; $i -le ($TablListRennes.length -1); $i++) {
#CreateList $SiteURL $TablListRennes[$i]
#}

#Appel de la fonction afin de créer les listes pour Sacib
#for($i=0; $i -le ($TablListSacib.length -1); $i++) {
#CreateList $SiteURL $TablListSacib[$i]
#}


### Nantes et Rennes ###
#Définition d'une première colonne
$FieldName = "DEPARTEMENT"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Rennes
#for ($i=0; $i -le ($TablListRennes.length -1); $i++){
#AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
#}


#Définition d'une deuxième colonne
$FieldName = "VILLE"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Rennes
#for ($i=0; $i -le ($TablListRennes.length -1); $i++){
#AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
#}



### Sacib ###
#Définition d'une première colonne
#$FieldName = "CLASSEMENT"
#$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
#$IsRequired = $False

#Ajoute la colonne pour les liste de Sacib
#for ($i=0; $i -le ($TablListSacib.length -1); $i++){
#AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
#}

#Définition d'une deuxieme colonne
#$FieldName = "1ER NIVEAU"
#$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
#$IsRequired = $False

#Ajoute la colonne pour les liste de Sacib
#for ($i=0; $i -le ($TablListSacib.length -1); $i++){
#AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
#}



### TOUTES LES LISTES ###
#Définition de la 3eme colonne
$FieldName = "OPERATION"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Rennes
#for ($i=0; $i -le ($TablListRennes.length -1); $i++){
#AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
#}

#Ajoute la colonne pour les liste de Sacib
#for ($i=0; $i -le ($TablListSacib.length -1); $i++){
#AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
#}


#Définition de la 4eme colonne
$FieldName = "CATEGORIE"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Rennes
#for ($i=0; $i -le ($TablListRennes.length -1); $i++){
#AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
#}

<#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}#>

#Définition de la 5eme colonne
$FieldName = "DOSSIER"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 6eme colonne
$FieldName = "Arborescence-EF-jusquà-dossier-de-rangement"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 7eme colonne
$FieldName = "Identifiant-unique-du-doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 8eme colonne
$FieldName = "Nom-du-Doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 9eme colonne
$FieldName = "Resume-Chemin-type-doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 10eme colonne
$FieldName = "Type-de-doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 11eme colonne
$FieldName = "Numero-de-version-du doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Number
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 12eme colonne
$FieldName = "Chemin-Doc-complet-EF"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 12eme colonne
$FieldName = "Chemin-Doc-complet-EF"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 13eme colonne
$FieldName = "Chemin-Doc-complet-sur-disque-sans-serveur"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 14eme colonne
$FieldName = "Chemin-Doc-complet-sur-disque-avec-serveur"
$FieldType = [Microsoft.SharePoint.SPFieldType]::URL
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 15eme colonne
$FieldName = "ETAT-DOC"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 16eme colonne
$FieldName = "REMARQUE-FICHIERS-CORROMPUS"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 17eme colonne
$FieldName = "INFO1"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>

#Définition de la 18eme colonne
$FieldName = "INFO2"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}
<#
#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}
#>