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
        #Si elle existe déjà je la supprime et je la créer
        else
        {
            write-host "Une liste ayant le même nom a été trouvée" -ForegroundColor Yellow

            $List = $Web.Lists[$ListName]
            $List.AllowDeletion = $true
            $List.Update()
            $List.Delete()
            $List = $Web.Lists.Add($ListName, $ListName, $ListTemplate) 
            write-host "Liste ${Listname} créée avec succès !" -ForegroundColor Green
        }
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}


#Fonction pour créer une colonne
Function AddFieldToList($SiteURL,$ListName, $FieldName, $FieldType, $IsRequired, $Visible)
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
                
                if($Visible)
                {
                #Met à jour la vue par défaut pour afficher les nouvelles colonnes
                $View = $List.DefaultView
                $View.ViewFields.Add($FieldName)
                $View.Update()
                }
 
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

$TablListNantes=@('NANTES-33-EYSINES-AUGUSTA-OBSOLETE', 'NANTES-33-FLOIRAC-LES-JARDINS-DE-FLORE', 'NANTES-33-SAINTE-EULALIE-LES-JARDINS-DE-GRACET', 'NANTES-44-BASSEGOULAINE-LES-COUPRIES', 'NANTES-44-BASSEGOULAINE-LES-VILLAS-CAUDALIE',
'NANTES-44-BLAIN-LA-CHAUSSEE-BLAIN', 'NANTES-44-CARQUEFOU-COTE-SCENE', 'NANTES-44-CHATEAUBRIANT', 'NANTES-44-LABAULE-CRYSTAL-PLAZZA', 'NANTES-44-LACHAPELLE-LES-PROMENADES-CHAPELAINES', 'NANTES-44-LAMONTAGNE',
'NANTES-44-NANTES-180-JULES-VERNE', 'NANTES-44-NANTES-BELLAGIO', 'NANTES-44-NANTES-CAP-A-LOUEST', 'NANTES-44-NANTES-CARRE-SAINT-ANDRE', 'NANTES-44-NANTES-COTE-PARC', 'NANTES-44-NANTES-EXALIS', 'NANTES-44-NANTES-ILIANA',
'NANTES-44-NANTES-LE-CARRE-SAINT-ANTOINE', 'NANTES-44-NANTES-LECHAPPEE-BELLE', 'NANTES-44-NANTES-LE-CLOS-SOLIS', 'NANTES-44-NANTES-LE-COURS-DES-ARTS', 'NANTES-44-NANTES-LES-HAUTS-SAINT-PASQUIER',
'NANTES-44-LES-PETITS-JARDINS', 'NANTES-44-NANTES-LES-RIVES-DE-SAINT-JOSEPH', 'NANTES-44-NANTES-LES-TERRASSES-DE-LA-HAUTE-MITRIE', 'NANTES-44-NANTES-MALAKOFF', 'NANTES-44-NANTES-NANTES-BD-MERSON',
'NANTES-44-NANTES-NINA-VERDE', 'NANTES-44-NANTES-PLAYTIME', 'NANTES-44-NANTES-QUAI-WEST', 'NANTES-44-NANTES-RESIDENCE-OPERA', 'NANTES-44-NANTES-RIVEA', 'NANTES-44-NANTES-TOUR-DAUVERGNE', 'NANTES-44-NANTES-VILLA-ALMERIA',
'NANTES-44-NANTES-VILLA-BAUSEJOUR', 'NANTES-44-NANTES-VILLA-DALBY', 'NANTES-44-NORTSURERDRE-RUE-FRANCOIS-DUPAS', 'NANTES-44-ORVAULT-JEUNESSE', 'NANTES-44-ORVAULT-LA-FREBAUDIERE',
'NANTES-44-ORVAULT-LES-RIVES-DE-CHANTILLY', 'NANTES-44-ORVAULT-ROUTE-BASSE-INDRE', 'NANTES-44-PORNICHET-LES-VILLAS-MARINAS', 'NANTES-44-PORNIC-LES-JARDINS-DE-LA-RIA', 'NANTES-44-REZE-BD-JEAN-MONNET', 
'NANTES-44-REZE-GRENADINE', 'NANTES-44-REZE-LES-GRENADINES-2EME-TRANCHE', 'NANTES-44-REZE-LES-PROMENADES-DE-SEVRE', 'NANTES-44-REZE-RUE-JOSEPH-CUGNOT', 'NANTES-44-REZE-VILLA-SEVRINA', 'NANTES-44-SAINTELUCE-NATUREA',
'NANTES-44-SAINTHERBLAIN-EXAPOLE', 'NANTES-44-SAINTHERBLAIN-EXAPOLEII', 'NANTES-44-SAINTHERBLAIN-LES-TERRASSES-DU-CHENE-VERT', 'NANTES-44-SAINTHERBLAIN-PARC-CASTELIA', 'NANTES-44-SAINTHERBLAIN-VILLA-FLORIA',
'NANTES-44-SAINTSEBASTIEN-PROMENADES-ENCHANTEES', 'NANTES-44-SAINTSEBASTIEN-SQUARE-PETIT-ANJOU', 'NANTES-44-SAUTRON-LES-ALLEES-ROSSINI', 'NANTES-44-STETIENNE-LES-JONQUILLES', 'NANTES-49-ANGERS-CARRE-GALLIENI',
'NANTES-49-ANGERS-DOMAINE-DE-LA-CERISAIE', 'NANTES-49-ANGERS-RESIDENCE-ELVIRA', 'NANTES-49-ANGERS-VILLAS-TANGO', 'NANTES-49-AVRILLE-LE-FLORIA', 'NANTES-56-BADEN-BADEN', 'NANTES-56-GUIDEL-ROZ-AVEL',
'NANTES-56-LORIENT-LES-JARDINS-DU-LEVANT', 'NANTES-56-QUEVEN-LE-DOMAINE-DE-VAL-QUEVEN', 'NANTES-56-RIANTEC-TY-RHU', 'NANTES-56-SAINTPIERRE-LE-HAMEAU-DES-TAMARIS', 'NANTES-56-VANNES-LE-DOMAINE-DU-BONDON',
'NANTES-56-VANNES-LES-JARDINS-DIRIS', 'NANTES-56-VANNES-VILLA-RAPHAELE', 'NANTES-78-JUZIERS-JUZIERS', 'NANTES-94-GENTILLY-L-ATRIODE-OBSOLETTE')

Write-Host $TablListNantes.Count

$TablListRennes=@('RENNES-14-DEMOUVILLE-THYSSEN-DEMOUVILLE', 'RENNES-17-AYTRE-LE-PATIO-DES-TILLEULS', 'RENNES-17-AYTRE-VILLAS-TILIA', 'RENNES-17-CLAVETTE-LE-TENNIS', 'RENNES-17-ESNANDES-ESNANDES',
'RENNES-17-LAROCHELLE-AZUREA', 'RENNES-17-LAROCHELLE-HORIZON-MER', 'RENNES-17-LAROCHELLE-LAROCHELLE-EINSTEIN', 'RENNES-17-LAROCHELLE-LE-GLOBE-TROTTER', 'RENNES-17-LAROCHELLE-VILLA-GARANCE',
'RENNES-17-NIEUL-SUR-MER-VILLA-ROSA' , 'RENNES-17-PERIGNY-LES-JARDINS-DES-ACANTHES', 'RENNES-22-LANVALLAY-LE-CLOS-DES-ORMEAUX', 'RENNES-22-LEHON-LE-CLOS-TRIARD', 'RENNES-22-PLENEUF-VALAND-COTE-MER',
'RENNES-22-PLENEUF-VALAND-LES-AIGUES-MARINES', 'RENNES-22-SAINT-CAST-LES-PIERRES-SONNANTES', 'RENNES-22-TREGUEUX-BREZILLET', 'RENNES-29-BREST-SAINT-PIERRE', 'RENNES-29-CARANTEC-LES-VILLAS-DE-KERGRIST',
'RENNES-29-CARANTEC-PARC-OCEAN', 'RENNES-29-GUILERS-LE-CLOS-VALENTIN', 'RENNES-29-GUIPAVAS-LES-RESIDENCES-SAINT-EXUPERY', 'RENNES-29-LANDIVISIAU-PARC-LANDIVISIAU', 'RENNES-29-LEDRENNEC-LES-JARDINS-DADRIEN',
'RENNES-29-LOCMARIA-PLOUZA-LES-JARDINS-DE-LOCMARIA', 'RENNES-29-PLOUARZEL-KERVEN', 'RENNES-29-PLOUGONVELIN-LES-JARDINS-DU-TREZ-HIR', 'RENNES-29-PONT-L-ABBE-TERRASSES-ETANG', 'RENNES-29-QUIMPER-DOMAINE-CENTRE', 
'RENNES-35-BAINDEBRETAGNE-LE-FORUM', 'RENNES-35-BAINDEBRETAGNE-RESIDENCE-DES-TANNEURS', 'RENNES-35-BETTON-LES-CAPUCINES', 'RENNES-35-BETTON-LES-CINQ-ILES', 'RENNES-35-BETTON-PRESQU-ILLE',
'RENNES-35-BETTON-VILLA-CANTALINA', 'RENNES-35-BETTON-VILLA-ODALIE', 'RENNES-35-BOISGERVILLY-LANCELOT-DU-LAC', 'RENNES-35-BRUZ-INEO', 'RENNES-35-BRUZ-LES-JARDINS-DE-BLOSSAC',
'RENNES-35-CESSON-SEVIGNE-CESSON-MAISON-MEDICALE', 'RENNES-35-CESSON-SEVIGNE-NET-PLUS', 'RENNES-35-CHANTEPIE-LE-DOMAINE-DU-CANAL', 'RENNES-35-CHANTEPIE-LES-COQUELICOTS', 'RENNES-35-CHANTEPIE-VILLA-ABELIA',
'RENNES-35-DOMLOUP-DOMLOUP', 'RENNES-35-ETRELLE-ETRELLE-VINCI-ENERGIES', 'RENNES-35-GOSNE-LES-PORTES-DOUEES', 'RENNES-35-JANZE-LES-AMBROISINES', 'RENNES-35-LACHAPELLE-RUE-DE-PACE', 'RENNES-35-LACHAPELLE-RUE-LECHLADE',
'RENNES-35-LACHAPELLE-VILLA-LAURENA', 'RENNES-35-LERHEU-JARDINS-DADELE', 'RENNES-35-LERHEU-LE-GRAND-JARDIN', 'RENNES-35-LERHEU-NEVENTI', 'RENNES-35-LERHEU-RUE-DE-RENNES', 'RENNES-35-LERHEU-THYSSEN-LE-RHEU', 
'RENNES-35-LERHEU-VILLAS-MELIES', 'RENNES-35-LHERMITAGE-LES-AQUARELLES', 'RENNES-35-LIFFRE-PARC-DES-ETANGS', 'RENNES-35-LIFFRE-ROSE-ARMOR', 'RENNES-35-MEZIERESSOUSC-LA-GRANDE-PREE',
'RENNES-35-MEZIERESSOUSC-LA-PREE-DU-PETIT-BOIS', 'RENNES-35-MONTAUBANDEBR-SAINT-ELOI', 'RENNES-35-NOYAL-LES-HORTENSIAS', 'RENNES-35-PARAME', 'RENNES-35-RENNES-ADIPH', 'RENNES-35-RENNES-AVENUE-MAGINOT',
'RENNES-35-RENNES-CAP-NORD', 'RENNES-35-RENNES-CARRE-DART', 'RENNES-35-RENNES-CASSINI', 'RENNES-35-RENNES-CASTEL-RIVIERA', 'RENNES-35-RENNES-COEUR-DE-VILLE', 'RENNES-35-RENNES-DEMAT', 'RENNES-35-RENNES-EOLYSII', 
'RENNES-35-RENNES-LA-VISITATION', 'RENNES-35-RENNES-LE-MURANO', 'RENNES-35-RENNES-LE-NOVEN', 'RENNES-35-RENNES-LE-SEXTANT', 'RENNES-35-RENNES-LES-JARDINS-DE-CHATILLON', 'RENNES-35-RENNES-LES-JARDINS-DE-NEROLI',
'RENNES-35-RENNES-LES-OPALINES', 'RENNES-35-RENNES-LES-PRAIRIES-DE-LILLE', 'RENNES-35-RENNES-LES-RIVES-DE-TASSIGNY', 'RENNES-35-RENNES-MADISON-PARC', 'RENNES-35-RENNES-OSIRIS', 'RENNES-35-RENNES-RESIDENCE-DE-VINCI',
'RENNES-35-RENNES-RUE-LEGRAVERAND', 'RENNES-35-RENNES-TERRANOVA', 'RENNES-35-RENNES-VILLA-CAMILLA', 'RENNES-35-RENNES-VILLA-DE-VINCI', 'RENNES-35-ROMAGNE-LE-CLOS-DES-SAULES', 'RENNES-35-SAINTBRIAC-LES-ROCHES-DOUVRES',
'RENNES-35-SAINTGREGOIRE-AXIS', 'RENNES-35-SAINTGREGOIRE-BPO', 'RENNES-35-SAINTGREGOIRE-EDONIA', 'RENNES-35-SAINTGREGOIRE-LA-BOUTIERE', 'RENNES-35-SAINTGREGOIRE-PARC-DE-BROCELIANDE', 'RENNES-35-SAINTGREGOIRE-PARC-ELLENA',
'RENNES-35-SAINTGREGOIRE-POLE-MEDICAL-LA-BOUTIERE', 'RENNES-35-SAINTJACQUES-ACTIVLAND-DARTY', 'RENNES-35-SAINTJACQUES-ACTIVLAND-MONDIAL-RELAY', 'RENNES-35-SAINTJACQUES-ADAGGIO-SOCIAL',
'RENNES-35-SAINTJACQUES-VILLA-GIULIA', 'RENNES-35-SAINTMALO-LES-BAYADERES', 'RENNES-35-SAINTMALO-LES-MARINES-DE-CHASLES', 'RENNES-35-SAINTMALO-PARC-DE-LHERMINE', 'RENNES-35-SAINTMALO-SQUARE-ACADIE',
'RENNES-44BUR-CARQUEFOU-ATALIS', 'RENNES-44BUR-COUERON-GLASSOLUTION', 'RENNES-44BUR-LACHAPELLE-INEO', 'RENNES-44BUR-NANTES-ESPACE-NEWTON', 'RENNES-44BUR-NANTES-EUROPA', 'RENNES-44BUR-NANTES-EXALIS',
'RENNES-44BUR-ORVAULT-BOIS-CESBRON', 'RENNES-44BUR-PORNIC-PORNIC', 'RENNES-44BUR-REZE-PSA-REZE', 'RENNES-44BUR-SAINTHERBLAIN-ASTURIA', 'RENNES-44BUR-SAINTHERBLAIN-EXAPOLE', 'RENNES-44BUR-SAINTHERBLAIN-SUNSET',
'RENNES-44BUR-TRIGNAC-GRAND-CHAMPS', 'RENNES-53-CHANGE-SPIE-CHANGE', 'RENNES-56BUR-RIEUX-VEOLIA-EAU', 'RENNES-72-LE-MANS-LE-GALILEE', 'RENNES-74-ANNECYLEVIEUX-PINSONS', 'RENNES-79-NIORT-LE-CLOS-DES-TILLEULS', 
'RENNES-85-FONTENAYLECOM-LES-COLLIBERTS', 'RENNES-85-LAROCHESURYON-RESIDENCE-ELLINE', 'RENNES-85-LESSABLESDOL', 'RENNES-85-SAINTHILAIRE-ST-HILAIRE-DU-RIEZ', 'RENNES-85-SAINTVINCENT-LE-SAINT-VINCENT',
'RENNES-85-ST-GILLES-CROIX-POEME')

Write-Host $TablListRennes.Count


$TablListSacib =@('SACIB-01-ALLEES-DU-HAVRE-SNC-BATIMALO', 'SACIB-01-ALLEES-DU-PORT', 'SACIB-01-ARGONAUTES-SARL-LOTIMALO', 'SACIB-01-BALUE-DOUTRELEAU', 'SACIB-01-BAYADERES', 'SACIB-01-BELLAVISTA', 'SACIB-01-BIGNON',
'SACIB-01-CENTRE-DIMAGERIE-MEDICALE-SNC-BATIMALO', 'SACIB-01-CHATEAU-MALO-NORD', 'SACIB-01-CLOS-BASTILLE', 'SACIB-01-CLOS-NEUF', 'SACIB-01-COMPTAGESMA', 'SACIB-01-COTE-DOCKS-SNC-CLERMONT', 'SACIB-01-DINAN', 
'SACIB-01-DIVISION-TERRAIN-LE-BIGNON', 'SACIB-01-DOMAINE-DU-MOULIN', 'SACIB-01-DOUTRELEAU-NOUVEL', 'SACIB-01-EUGENE-HERPIN', 'SACIB-01-FLERS-LES-JARDINS-DE-LORANGERIE', 'SACIB-01-FONTAINE-AU-VAIS',
'SACIB-01-FONTENELLES-IMPASSE-DU-MORSE', 'SACIB-01-FREHEL-LA-CARQUOIS', 'SACIB-01-GRANDE-BARONNIE', 'SACIB-01-HABERT', 'SACIB-01-HAMEAU-DE-LA-FONTAINE-SNC-BATIMALO', 'SACIB-01-HAVRE-FLEURY', 'SACIB-01-KER-LOUIS-QUEBEC',
'SACIB-01-LA-GOUESNIERE', 'SACIB-01-LE-PONT', 'SACIB-01-LILOE-2', 'SACIB-01-NAUTICA', 'SACIB-01-NESTOR', 'SACIB-01-NEWQUAY', 'SACIB-01-ODYSSEE', 'SACIB-01-OPALINES', 'SACIB-01-PATIOS-DU-HAVRE', 'SACIB-01-PLEURTUIT', 
'SACIB-01-PONTORSON-LE-SEQUOIA', 'SACIB-01-QUAI-SUD', 'SACIB-01-ROCH-PIERRE-LOUIS', 'SACIB-01-ROSSEL', 'SACIB-01-ROZ-SUR-COUESNON-LE-CLOS-DE-LA-GRANGE', 'SACIB-01-RUE-DU-VALLON', 'SACIB-01-SAINT-BRIAC',
'SACIB-01-SAINT-JOUAN-DES-GUERETS-LES-VOILES-ROUGES', 'SACIB-01-SAINT-MELOIR-DES-ONDES', 'SACIB-01-ST-BRIAC-ZAC-DES-TOURELLES', 'SACIB-01-TERRASSES-DE-RIVASSELOU', 'SACIB-01-VILLA-EDOUARD-VII', 'SACIB-01-VOILERIES',
'SACIB-02-ALLEES-DU-HAVRE', 'SACIB-02-ALLEES-DU-PORT', 'SACIB-02-CANCALE-BELLE-BRISE', 'SACIB-02-CLOS-BASTILLE', 'SACIB-02-COTE-DOCKS', 'SACIB-02-DINAN', 'SACIB-02-DINARD', 'SACIB-02-DIVISION-TERRAIN-LE-BIGNON',
'SACIB-02-DOL-DE-BRETAGNE', 'SACIB-02-FONTAINE-AUX-VAIS', 'SACIB-02-FREHEL-LA-CARQUOIS', 'SACIB-02-HAVRE-FLEURY', 'SACIB-02-HIREL', 'SACIB-02-KER-LOUIS-QUEBEC', 'SACIB-02-LA-GOUESNIERE', 'SACIB-02-NEWQUAY',
'SACIB-02-PATIOS-DU-HAVRE', 'SACIB-02-PLEURTUIT', 'SACIB-02-PONTORSON-LE-SEQUOIA', 'SACIB-02-ROCH-PIERRE-LOUIS', 'SACIB-02-ROZ-SUR-COUESNOU', 'SACIB-02-SAINT-BRIAC', 'SACIB-02-SAINT-JOUAN-DES-GUERETS-LES-VOILES-ROUGES',
'SACIB-02-SAINT-MELOIR-DES-ONDES', 'SACIB-02-ST-BRIAC-CHEMIN-DES-TOURELLES', 'SACIB-02-VILLE-ES-NONAIS', 'SACIB-04-DINAN-PARC-DU-COMTE-DE-LA-GARAYE', 'SACIB-04-FLERS-LES-JARDINS-DE-LORANGERIE', 'SACIB-04-FREHEL-LA-CARQUOIS',
'SACIB-04-HAVRE-FLEURY', 'SACIB-04-PLEUDIHEN-SUR-RANCE-VAL-DORIENT', 'SACIB-04-PONTORSON-LE-SEQUOIA', 'SACIB-04-SAINT-JOUAN-DES-GUERETS-LES-VOILES-ROUGES', 'SACIB-04-ST-BRIAC', 'SACIB-04-VILLA-FLORIANE-ANCIEUX',
'SACIB-05-CANCALE-BELLE-PRISE', 'SACIB-05-FLERS-LES-JARDINS-DE-LORANGERIE', 'SACIB-05-PONTORSON-LE-SEQUOIA')

Write-Host $TablListSacib.Count



#Appel de la fonction afin de créer les listes pour Nantes
for($i=0; $i -le ($TablListNantes.length -1); $i++) {
CreateList $SiteURL $TablListNantes[$i]
}

#Appel de la fonction afin de créer les listes pour Rennes
for($i=0; $i -le ($TablListRennes.length -1); $i++) {
CreateList $SiteURL $TablListRennes[$i]
}

#Appel de la fonction afin de créer les listes pour Sacib
for($i=0; $i -le ($TablListSacib.length -1); $i++) {
CreateList $SiteURL $TablListSacib[$i]
}


### Nantes et Rennes ###
#Définition d'une première colonne
$FieldName = "DEPARTEMENT"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $true

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition d'une deuxième colonne
$FieldName = "VILLE"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $true

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}



### Sacib ###
#Définition d'une première colonne
$FieldName = "CLASSEMENT"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}

#Définition d'une deuxieme colonne
$FieldName = "1ER NIVEAU"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}



### TOUTES LES LISTES ###
#Définition de la 3eme colonne
$FieldName = "OPERATION"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $true

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 4eme colonne
$FieldName = "CATEGORIE"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $true

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){ 
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}

#Définition de la 5eme colonne
$FieldName = "DOSSIER"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $true

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}


#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 6eme colonne
$FieldName = "Arborescence-EF-jusquà-dossier-de-rangement"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 7eme colonne
$FieldName = "Identifiant-unique-du-doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 8eme colonne
$FieldName = "Nom-du-Doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $true

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 9eme colonne
$FieldName = "Resume-Chemin-type-doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 10eme colonne
$FieldName = "Type-de-doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 11eme colonne
$FieldName = "Numero-de-version-du doc"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Number
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 12eme colonne
$FieldName = "Chemin-Doc-complet-EF"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}


#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}

#Définition de la 13eme colonne
$FieldName = "Chemin-Doc-complet-sur-disque-sans-serveur"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 14eme colonne
$FieldName = "Chemin-Doc-complet-sur-disque-avec-serveur"
$FieldType = [Microsoft.SharePoint.SPFieldType]::URL
$IsRequired = $False
$Visible = $true

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}


#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 15eme colonne
$FieldName = "ETAT-DOC"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false


#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 16eme colonne
$FieldName = "REMARQUE-FICHIERS-CORROMPUS"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired $Visible
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired $Visible
}


#Définition de la 17eme colonne
$FieldName = "INFO1"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}


#Définition de la 18eme colonne
$FieldName = "INFO2"
$FieldType = [Microsoft.SharePoint.SPFieldType]::Text
$IsRequired = $False
$Visible = $false

#Ajoute la colonne pour les liste de Nantes
for ($i=0; $i -le ($TablListNantes.length -1); $i++){
AddFieldToList $SiteURL $TablListNantes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Rennes
for ($i=0; $i -le ($TablListRennes.length -1); $i++){
AddFieldToList $SiteURL $TablListRennes[$i] $FieldName $FieldType $IsRequired
}

#Ajoute la colonne pour les liste de Sacib
for ($i=0; $i -le ($TablListSacib.length -1); $i++){
AddFieldToList $SiteURL $TablListSacib[$i] $FieldName $FieldType $IsRequired
}