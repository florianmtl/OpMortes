Add-PSSnapin Microsoft.SharePoint.PowerShell

#Paramètres
#Récupère le fichier CSV
Try {
    # NANTES
    $listNantes33Eysines = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-33-EYSINES-AUGUSTA-OBSOLETE.csv" -Delimiter ";"
    $listNantes33Floirac = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-33-FLOIRAC-LES-JARDINS-DE-FLORE.csv" -Delimiter ";"
    $listNantes33SainteEulalie = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-33-SAINTE-EULALIE-LES-JARDINS-DE-GRACET.csv" -Delimiter ";"
    $listNantes44BassegoulaineCoupries = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-BASSEGOULAINE-LES-COUPRIES.csv" -Delimiter ";"
    $listNantes44BassegoulaineVillas = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-BASSEGOULAINE-LES-VILLAS-CAUDALIE.csv" -Delimiter ";"
    $listNantes44Blain = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-BLAIN-LA-CHAUSSEE-BLAIN.csv" -Delimiter ";"
    $listNantes44Carquefou = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-CARQUEFOU-COTE-SCENE.csv" -Delimiter ";"
    $listNantes44Chateaubriant = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-CHATEAUBRIANT.csv" -Delimiter ";"
    $listNantes44LaBaule = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-LABAULE-CRYSTAL-PLAZZA.csv" -Delimiter ";"
    $listNantes44LaChapelle = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-LACHAPELLE-LES-PROMENADES-CHAPELAINES.csv" -Delimiter ";"
    $listNantes44LaMontagne = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-LAMONTAGNE.csv" -Delimiter ";"
    $listNantes44NantesJules = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-180-JULES-VERNE.csv" -Delimiter ";"
    $listNantes44NantesBellagio = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-BELLAGIO.csv" -Delimiter ";"
    $listNantes44NantesCapALouest = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-CAP-A-LOUEST.csv" -Delimiter ";"
    $listNantes44NantesCarre = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-CARRE-SAINT-ANDRE.csv" -Delimiter ";"
    $listNantes44NantesCoteParc = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-COTE-PARC.csv" -Delimiter ";"
    $listNantes44NantesExalis = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-EXALIS.csv" -Delimiter ";"
    $listNantes44NantesIliana = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-ILIANA.csv" -Delimiter ";"
    $listNantes44NantesLeCarre = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LE-CARRE-SAINT-ANTOINE.csv" -Delimiter ";"
    $listNantes44NantesLechappee = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LECHAPPEE-BELLE.csv" -Delimiter ";"
    $listNantes44NantesLeClos = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LE-CLOS-SOLIS.csv" -Delimiter ";"
    $listNantes44NantesLeCours = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LE-COURS-DES-ARTS.csv" -Delimiter ";"
    $listNantes44NantesLesHauts = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LES-HAUTS-SAINT-PASQUIER.csv" -Delimiter ";"
    $listNantes44NantesLesPetits = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LES-PETITS-JARDINS.csv" -Delimiter ";"
    $listNantes44NantesLesRives = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LES-RIVES-DE-SAINT-JOSEPH.csv" -Delimiter ";"
    $listNantes44NantesLesTerrasses = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-LES-TERRASSES-DE-LA-HAUTE-MITRIE.csv" -Delimiter ";"
    $listNantes44NantesMalakoff = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-MALAKOFF.csv" -Delimiter ";"
    $listNantes44NantesNantes = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-NANTES-BD-MERSON.csv" -Delimiter ";"
    $listNantes44NantesNina = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-NINA-VERDE.csv" -Delimiter ";"
    $listNantes44NantesPlaytime = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-PLAYTIME.csv" -Delimiter ";"
    $listNantes44NantesQuaiWest = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-QUAI-WEST.csv" -Delimiter ";"
    $listNantes44NantesResidence = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-RESIDENCE-OPERA.csv" -Delimiter ";"
    $listNantes44NantesRivea = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-RIVEA.csv" -Delimiter ";"
    $listNantes44NantesTour = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-TOUR-DAUVERGNE.csv" -Delimiter ";"
    $listNantes44NantesVillaAlmeria = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-VILLA-ALMERIA.csv" -Delimiter ";"
    $listNantes44NantesVillaBausejour = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-VILLA-BAUSEJOUR.csv" -Delimiter ";"
    $listNantes44NantesVillaDalby = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NANTES-VILLA-DALBY.csv" -Delimiter ";"
    $listNantes44NortSurErdre = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-NORTSURERDRE-RUE-FRANCOIS-DUPAS.csv" -Delimiter ";"
    $listNantes44OrvaultJeunesse = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-ORVAULT-JEUNESSE.csv" -Delimiter ";"
    $listNantes44OrvaultFrebaudiere = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-ORVAULT-LA-FREBAUDIERE.csv" -Delimiter ";"
    $listNantes44OrvaultLesRives = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-ORVAULT-LES-RIVES-DE-CHANTILLY.csv" -Delimiter ";"
    $listNantes44OrvaultRouteBasse = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-ORVAULT-ROUTE-BASSE-INDRE.csv" -Delimiter ";"
    $listNantes44Pornichet = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-PORNICHET-LES-VILLAS-MARINAS.csv" -Delimiter ";"
    $listNantes44Pornic = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-PORNIC-LES-JARDINS-DE-LA-RIA.csv" -Delimiter ";"
    $listNantes44RezeBd = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-REZE-BD-JEAN-MONNET.csv" -Delimiter ";"
    $listNantes44RezeGrenadine = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-REZE-GRENADINE.csv" -Delimiter ";"
    $listNantes44RezeLesGrenadines = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-REZE-LES-GRENADINES-2EME-TRANCHE.csv" -Delimiter ";"
    $listNantes44RezeLesPromenades = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-REZE-LES-PROMENADES-DE-SEVRE.csv" -Delimiter ";"
    $listNantes44RezeRueJoseph = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-REZE-RUE-JOSEPH-CUGNOT.csv" -Delimiter ";"
    $listNantes44RezeVillaSevrina = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-REZE-VILLA-SEVRINA.csv" -Delimiter ";"
    $listNantes44SainteLuce = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTELUCE-NATUREA.csv" -Delimiter ";"
    $listNantes44SaintHerblainExapole = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTHERBLAIN-EXAPOLE.csv" -Delimiter ";"
    $listNantes44SaintHerblainExapoleII = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTHERBLAIN-EXAPOLEII.csv" -Delimiter ";"
    $listNantes44SaintHerblainLesTerrasses = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTHERBLAIN-LES-TERRASSES-DU-CHENE-VERT.csv" -Delimiter ";"
    $listNantes44SaintHerblainParc = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTHERBLAIN-PARC-CASTELIA.csv" -Delimiter ";"
    $listNantes44SaintHerblainVilla = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTHERBLAIN-VILLA-FLORIA.csv" -Delimiter ";"
    $listNantes44SaintSebastienPromenades = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTSEBASTIEN-PROMENADES-ENCHANTEES.csv" -Delimiter ";"
    $listNantes44SaintSebastienSquare = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAINTSEBASTIEN-SQUARE-PETIT-ANJOU.csv" -Delimiter ";"
    $listNantes44Sautron = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-SAUTRON-LES-ALLEES-ROSSINI.csv" -Delimiter ";"
    $listNantes44StEtienne = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-44-STETIENNE-LES-JONQUILLES.csv" -Delimiter ";"
    $listNantes49AngersCarre = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-49-ANGERS-CARRE-GALLIENI.csv" -Delimiter ";"
    $listNantes49AngersDomaine = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-49-ANGERS-DOMAINE-DE-LA-CERISAIE.csv" -Delimiter ";"
    $listNantes49AngersResidence = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-49-ANGERS-RESIDENCE-ELVIRA.csv" -Delimiter ";"
    $listNantes49AngersVillas = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-49-ANGERS-VILLAS-TANGO.csv" -Delimiter ";"
    $listNantes49Avrille = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-49-AVRILLE-LE-FLORIA.csv" -Delimiter ";"
    $listNantes56Baden = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-BADEN-BADEN.csv" -Delimiter ";"
    $listNantes56Guidel = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-GUIDEL-ROZ-AVEL.csv" -Delimiter ";"
    $listNantes56Lorient = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-LORIENT-LES-JARDINS-DU-LEVANT.csv" -Delimiter ";"
    $listNantes56Queven = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-QUEVEN-LE-DOMAINE-DE-VAL-QUEVEN.csv" -Delimiter ";"
    $listNantes56Riantec = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-RIANTEC-TY-RHU.csv" -Delimiter ";"
    $listNantes56SaintPierre = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-SAINTPIERRE-LE-HAMEAU-DES-TAMARIS.csv" -Delimiter ";"
    $listNantes56VannesLeDomaine = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-VANNES-LE-DOMAINE-DU-BONDON.csv" -Delimiter ";"
    $listNantes56VannesLesJardins = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-VANNES-LES-JARDINS-DIRIS.csv" -Delimiter ";"
    $listNantes56VannesVilla = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-56-VANNES-VILLA-RAPHAELE.csv" -Delimiter ";"
    $listNantes78Juziers = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-78-JUZIERS-JUZIERS.csv" -Delimiter ";"
    $listNantes94Gentilly = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\Operations\OpMortes-NANTES-94-GENTILLY-L-ATRIODE-OBSOLETTE.csv" -Delimiter ";"

    Write-Host $tablListCsv.Count " fichiers récupérés avec succés !" -ForegroundColor Green;

}
catch {
    Write-Host "Un ou des fichiers CSV de Nantes est(sont) introuvable(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}
#>

#Récupère les listes de NANTES du site Sharepoint

Try {
    $lNantes33Eysines = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES33EYSINESAUGUSTAOBSOLETE")
    $lNantes33Floirac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES33FLOIRACLESJARDINSDEFLORE")
    $lNantes33SainteEulalie = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES33SAINTEEULALIELESJARDINSDEGRACET")
    $lNantes44BassegoulaineCoupries = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44BASSEGOULAINELESCOUPRIES")
    $lNantes44BassegoulaineVillas = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44BASSEGOULAINELESVILLASCAUDALIE")
    $lNantes44Blain = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44BLAINLACHAUSSEEBLAIN")
    $lNantes44Carquefou = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44CARQUEFOUCOTESCENE")
    $lNantes44Chateaubriant = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44CHATEAUBRIANT")
    $lNantes44LaBaule = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44LABAULECRYSTALPLAZZA")
    $lNantes44LaChapelle = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44LACHAPELLELESPROMENADESCHAPELAINES")
    $lNantes44LaMontagne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44LAMONTAGNE")
    $lNantes44NantesJules = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTES180JULESVERNE")
    $lNantes44NantesBellagio = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESBELLAGIO")
    $lNantes44NantesCapALouest = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESCAPALOUEST")
    $lNantes44NantesCarre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESCARRESAINTANDRE")
    $lNantes44NantesCoteParc = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESCOTEPARC")
    $lNantes44NantesExalis = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESEXALIS")
    $lNantes44NantesIliana = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESILIANA")
    $lNantes44NantesLeCarre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESLECARRESAINTANTOINE")
    $lNantes44NantesLechappee = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESLECHAPPEEBELLE")
    $lNantes44NantesLeClosSolis = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESLECLOSSOLIS")
    $lNantes44NantesLeCours = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESLECOURSDESARTS")
    $lNantes44NantesLesHauts = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESLESHAUTSSAINTPASQUIER")
    $lNantes44NantesLesPetits = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44LESPETITSJARDINS")
    $lNantes44NantesLesRives = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESLESRIVESDESAINTJOSEPH")
    $lNantes44NantesLesTerrasses = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESLESTERRASSESDELAHAUTEMITRIE")
    $lNantes44NantesMalakoff = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESMALAKOFF")
    $lNantes44NantesNantes = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESNANTESBDMERSON")
    $lNantes44NantesNinaVerde = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESNINAVERDE")
    $lNantes44NantesPlaytime = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESPLAYTIME")
    $lNantes44NantesQuaiWest = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESQUAIWEST")
    $lNantes44NantesResidenceOpera = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESRESIDENCEOPERA")
    $lNantes44NantesRivea = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESRIVEA")
    $lNantes44NantesTour = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESTOURDAUVERGNE")
    $lNantes44NantesVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESVILLAALMERIA")
    $lNantes44NantesVillaBauseJour = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESVILLABAUSEJOUR")
    $lNantes44NantesVillaDalby = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NANTESVILLADALBY")
    $lNantes44NordSurErdre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NORTSURERDRERUEFRANCOISDUPAS")
    $lNantes44OrvaultJeunesse = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44ORVAULTJEUNESSE")
    $lNantes44OrvaultLaFrebaudiere = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44ORVAULTLAFREBAUDIERE")
    $lNantes44OrvaultLesRives = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44ORVAULTLESRIVESDECHANTILLY")
    $lNantes44OrvaultRoute = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44ORVAULTROUTEBASSEINDRE")
    $lNantes44Pornichet = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44PORNICHETLESVILLASMARINAS")
    $lNantes44Pornic = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44PORNICLESJARDINSDELARIA")
    $lNantes44RezeBd = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44REZEBDJEANMONNET")
    $lNantes44RezeGrenadine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44REZEGRENADINE")
    $lNantes44RezeLesGrenadines = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44REZELESGRENADINES2EMETRANCHE")
    $lNantes44RezeLesPromenades = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44REZELESPROMENADESDESEVRE")
    $lNantes44RezeRueJoseph = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44REZERUEJOSEPHCUGNOT")
    $lNantes44RezeVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44REZEVILLASEVRINA")
    $lNantes44SainteLuce = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTELUCENATUREA")
    $lNantes44SaintHerblainExapole = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTHERBLAINEXAPOLE")
    $lNantes44SaintHerblainExapoleII = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTHERBLAINEXAPOLEII")
    $lNantes44SaintHerblainLesTerrasses = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTHERBLAINLESTERRASSESDUCHENEVERT")
    $lNantes44SaintHerblainParc = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTHERBLAINPARCCASTELIA")
    $lNantes44SaintHerblainVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTHERBLAINVILLAFLORIA")
    $lNantes44SaintSebastienPromenades = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTSEBASTIENPROMENADESENCHANTEES")
    $lNantes44SaintSebastienSquare = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAINTSEBASTIENSQUAREPETITANJOU")
    $lNantes44Sautron = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SAUTRONLESALLEESROSSINI")
    $lNantes44StEtienne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44STETIENNELESJONQUILLES")
    $lNantes49AngersCarre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES49ANGERSCARREGALLIENI")
    $lNantes49AngersDomaine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES49ANGERSDOMAINEDELACERISAIE")
    $lNantes49AngersResidence = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES49ANGERSRESIDENCEELVIRA")
    $lNantes49AngersVillas = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES49ANGERSVILLASTANGO")
    $lNantes49Avrille = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES49AVRILLELEFLORIA")
    $lNantes56Baden = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56BADENBADEN")
    $lNantes56Guidel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56GUIDELROZAVEL")
    $lNantes56Lorient = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56LORIENTLESJARDINSDULEVANT")
    $lNantes56Queven = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56QUEVENLEDOMAINEDEVALQUEVEN")
    $lNantes56Riantec = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56RIANTECTYRHU")
    $lNantes56SaintPierre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56SAINTPIERRELEHAMEAUDESTAMARIS")
    $lNantes56VannesLeDomaine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56VANNESLEDOMAINEDUBONDON")
    $lNantes56VannesLesJardins = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56VANNESLESJARDINSDIRIS")
    $lNantes56VannesVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56VANNESVILLARAPHAELE")
    $lNantes78Juziers = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES78JUZIERSJUZIERS")
    $lNantes94Gentilly = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES94GENTILLYLATRIODEOBSOLETTE")
    
    Write-Host $tablListSite.Count " listes récupérés avec succés !" -ForegroundColor Green;

}
Catch {
    Write-Host "Une ou des listes n'a(ont) pas été trouvée(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}


    $tablListCsv = @($listNantes33Eysines, $listNantes33Floirac, $listNantes33SainteEulalie, $listNantes44BassegoulaineCoupries, $listNantes44BassegoulaineVillas, $listNantes44Blain, $listNantes44Carquefou,
    $listNantes44Chateaubriant, $listNantes44LaBaule, $listNantes44LaChapelle, $listNantes44LaMontagne, $listNantes44NantesJules, $listNantes44NantesBellagio, $listNantes44NantesCapALouest, $listNantes44NantesCarre,
    $listNantes44NantesCoteParc, $listNantes44NantesExalis, $listNantes44NantesIliana, $listNantes44NantesLeCarre, $listNantes44NantesLechappee, $listNantes44NantesLeClos, $listNantes44NantesLeCours,
    $listNantes44NantesLesHauts, $listNantes44NantesLesPetits, $listNantes44NantesLesRives, $listNantes44NantesLesTerrasses, $listNantes44NantesMalakoff, $listNantes44NantesNantes, $listNantes44NantesNina,
    $listNantes44NantesPlaytime, $listNantes44NantesQuaiWest, $listNantes44NantesResidence, $listNantes44NantesRivea, $listNantes44NantesTour, $listNantes44NantesVillaAlmeria, $listNantes44NantesVillaBausejour, 
    $listNantes44NantesVillaDalby, $listNantes44NortSurErdre, $listNantes44OrvaultJeunesse, $listNantes44OrvaultFrebaudiere, $listNantes44OrvaultLesRives, $listNantes44OrvaultRouteBasse, $listNantes44Pornichet,
    $listNantes44Pornic, $listNantes44RezeBd, $listNantes44RezeGrenadine, $listNantes44RezeLesGrenadines, $listNantes44RezeLesPromenades, $listNantes44RezeRueJoseph, $listNantes44RezeVillaSevrina, $listNantes44SainteLuce, 
    $listNantes44SaintHerblainExapole, $listNantes44SaintHerblainExapoleII, $listNantes44SaintHerblainLesTerrasses, $listNantes44SaintHerblainParc, $listNantes44SaintHerblainVilla, $listNantes44SaintSebastienPromenades,
    $listNantes44SaintSebastienSquare, $listNantes44Sautron, $listNantes44StEtienne, $listNantes49AngersCarre, $listNantes49AngersDomaine, $listNantes49AngersResidence, $listNantes49AngersVillas, $listNantes49Avrille,
    $listNantes56Baden, $listNantes56Guidel, $listNantes56Lorient, $listNantes56Queven, $listNantes56Riantec, $listNantes56SaintPierre, $listNantes56VannesLeDomaine, $listNantes56VannesLesJardins,
    $listNantes56VannesVilla, $listNantes78Juziers, $listNantes94Gentilly)
    
    $tablListSite = @($lNantes33Eysines, $lNantes33Floirac, $lNantes33SainteEulalie, $lNantes44BassegoulaineCoupries, $lNantes44BassegoulaineVillas, $lNantes44Blain, $lNantes44Carquefou, $lNantes44Chateaubriant,
    $lNantes44LaBaule, $lNantes44LaChapelle, $lNantes44LaMontagne, $lNantes44NantesJules, $lNantes44NantesBellagio, $lNantes44NantesCapALouest, $lNantes44NantesCarre, $lNantes44NantesCoteParc, $lNantes44NantesExalis,
    $lNantes44NantesIliana, $lNantes44NantesLeCarre, $lNantes44NantesLechappee, $lNantes44NantesLeClosSolis, $lNantes44NantesLeCours, $lNantes44NantesLesHauts, $lNantes44NantesLesPetits, $lNantes44NantesLesRives,
    $lNantes44NantesLesTerrasses, $lNantes44NantesMalakoff, $lNantes44NantesNantes, $lNantes44NantesNinaVerde, $lNantes44NantesPlaytime, $lNantes44NantesQuaiWest, $lNantes44NantesResidenceOpera, $lNantes44NantesRivea,
    $lNantes44NantesTour, $lNantes44NantesVilla, $lNantes44NantesVillaBauseJour, $lNantes44NantesVillaDalby, $lNantes44NordSurErdre, $lNantes44OrvaultJeunesse, $lNantes44OrvaultLaFrebaudiere, $lNantes44OrvaultLesRives,
    $lNantes44OrvaultRoute, $lNantes44Pornichet, $lNantes44Pornic, $lNantes44RezeBd, $lNantes44RezeGrenadine, $lNantes44RezeLesGrenadines, $lNantes44RezeLesPromenades, $lNantes44RezeRueJoseph, $lNantes44RezeVilla,
    $lNantes44SainteLuce, $lNantes44SaintHerblainExapole, $lNantes44SaintHerblainExapoleII, $lNantes44SaintHerblainLesTerrasses, $lNantes44SaintHerblainParc, $lNantes44SaintHerblainVilla,
    $lNantes44SaintSebastienPromenades, $lNantes44SaintSebastienSquare, $lNantes44Sautron, $lNantes44StEtienne, $lNantes49AngersCarre, $lNantes49AngersDomaine, $lNantes49AngersResidence, $lNantes49AngersVillas,
    $lNantes49Avrille, $lNantes56Baden, $lNantes56Guidel, $lNantes56Lorient, $lNantes56Queven, $lNantes56Riantec, $lNantes56SaintPierre, $lNantes56VannesLeDomaine, $lNantes56VannesLesJardins, $lNantes56VannesVilla,
    $lNantes78Juziers, $lNantes94Gentilly)

#Ajout des données pour les 2 premières liste étant donné que la liste Sacib n'est pas construite de la même façon que les autres
Try
{

    for ($i=0; $i -le ($tablListSite.length -1); $i++) {
        #Parcours les éléments
        $r = 1;
        $counter = 0;

        foreach($item in $tablListCsv[$i])
        {
            $ni = $tablListSite[$i].items.Add();
            $ni["Titre"] = $r;
            $ni["DEPARTEMENT"] = $item.DEPARTEMENT;
            $ni["VILLE"] = $item.VILLE;
            $ni["OPERATION"] = $item.OPERATION;
            $ni["CATEGORIE"] = $item.CATEGORIE;
            $ni["DOSSIER"] = $item.DOSSIER;
            $ni["Arborescence-EF-jusquà-dossier-de-rangement"] = $item.'Arborescence-EF-jusqua-dossier-de-rangement';
            $ni["Identifiant-unique-du-doc"] = $item.'Identifiant-unique-du-doc';
            $ni["Nom-du-Doc"] = $item.'Nom-du-Doc';
            $ni["Resume-Chemin-type-doc"] = $item.'Resume-Chemin-type-doc';
            $ni["Type-de-doc"] = $item.'Type de doc';
            $ni["Numero-de-version-du doc"] = $item.'Numero-de-version-du doc';
            $ni["Chemin-Doc-complet-EF"] = $item.'Chemin-Doc-complet-EF';
            $ni["Chemin-Doc-complet-sur-disque-sans-serveur"] = $item.'Chemin-Doc-complet-sur-disque-sans-serveur';
            $ni["Chemin-Doc-complet-sur-disque-avec-serveur"] = $item.'Chemin-Doc-complet-sur-disque-avec-serveur';
            $ni["ETAT-DOC"] = $item.'ETAT-DOC';
            $ni["REMARQUE-FICHIERS-CORROMPUS"] = $item.'REMARQUE-FICHIERS-CORROMPUS';
            $ni["INFO1"] = $item.INFO1;
            $ni["INFO2"] = $item.INFO2;

            #Mise a jour de la liste
            $ni.Update()
            $r++;
            $counter++;

            Write-Progress -Id 1 -Activity "Importation des données" -Status 'Progress' -PercentComplete (($counter / ($tablListCsv[$i].Count+1)) * 100);
        }
        Write-Host ((${r} -1), " ligne(s) ajoutée(s) à la liste", $tablListSite[$i]) -ForegroundColor Green;
    }
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
}