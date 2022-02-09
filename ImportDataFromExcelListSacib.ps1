Add-PSSnapin Microsoft.SharePoint.PowerShell

#Paramètres
#Récupère le fichier CSV
Try {
    # SACIB
    $listSacib01AlleesHavre = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ALLEES-DU-HAVRE-SNC-BATIMALO.csv" -Delimiter ";"
    $listSacib01AlleesPort = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ALLEES-DU-PORT.csv" -Delimiter ";"
    $listSacib01Argonautes = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ARGONAUTES-SARL-LOTIMALO.csv" -Delimiter ";"
    $listSacib01Balue = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-BALUE-DOUTRELEAU.csv" -Delimiter ";"
    $listSacib01Bayaderes = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-BAYADERES.csv" -Delimiter ";"
    $listSacib01Bellavista = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-BELLAVISTA.csv" -Delimiter ";"
    $listSacib01Bignon = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-BIGNON.csv" -Delimiter ";"
    $listSacib01Centre = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-CENTRE-DIMAGERIE-MEDICALE-SNC-BATIMALO.csv" -Delimiter ";"
    $listSacib01Chateau = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-CHATEAU-MALO-NORD.csv" -Delimiter ";"
    $listSacib01ClosBastille = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-CLOS-BASTILLE.csv" -Delimiter ";"
    $listSacib01ClosNeuf = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-CLOS-NEUF.csv" -Delimiter ";"
    $listSacib01Compta = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-COMPTAGESMA.csv" -Delimiter ";"
    $listSacib01Cote = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-COTE-DOCKS-SNC-CLERMONT.csv" -Delimiter ";"
    $listSacib01Dinan = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-DINAN.csv" -Delimiter ";"
    $listSacib01Division = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-DIVISION-TERRAIN-LE-BIGNON.csv" -Delimiter ";"
    $listSacib01Domaine = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-DOMAINE-DU-MOULIN.csv" -Delimiter ";"
    $listSacib01Doutreleau = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-DOUTRELEAU-NOUVEL.csv" -Delimiter ";"
    $listSacib01EugeneHerpin = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-EUGENE-HERPIN.csv" -Delimiter ";"
    $listSacib01Flers = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-FLERS-LES-JARDINS-DE-LORANGERIE.csv" -Delimiter ";"
    $listSacib01Fontaine = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-FONTAINE-AU-VAIS.csv" -Delimiter ";"
    $listSacib01Fontenelles = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-FONTENELLES-IMPASSE-DU-MORSE.csv" -Delimiter ";"
    $listSacib01Frehel = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-FREHEL-LA-CARQUOIS.csv" -Delimiter ";"
    $listSacib01Grande = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-GRANDE-BARONNIE.csv" -Delimiter ";"
    $listSacib01Habert = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-HABERT.csv" -Delimiter ";"
    $listSacib01Hameau = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-HAMEAU-DE-LA-FONTAINE-SNC-BATIMALO.csv" -Delimiter ";"
    $listSacib01Havre = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-HAVRE-FLEURY.csv" -Delimiter ";"
    $listSacib01Ker = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-KER-LOUIS-QUEBEC.csv" -Delimiter ";"
    $listSacib01Gouesniere = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-LA-GOUESNIERE.csv" -Delimiter ";"
    $listSacib01LePont = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-LE-PONT.csv" -Delimiter ";"
    $listSacib01Liloe = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-LILOE-2.csv" -Delimiter ";"
    $listSacib01Nautica = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-NAUTICA.csv" -Delimiter ";"
    $listSacib01Nestor = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-NESTOR.csv" -Delimiter ";"
    $listSacib01Newquay = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-NEWQUAY.csv" -Delimiter ";"
    $listSacib01Odyssee = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ODYSSEE.csv" -Delimiter ";"
    $listSacib01Opalines = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-OPALINES.csv" -Delimiter ";"
    $listSacib01Patios = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-PATIOS-DU-HAVRE.csv" -Delimiter ";"
    $listSacib01Pleurtuit = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-PLEURTUIT.csv" -Delimiter ";"
    $listSacib01Pontorson = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-PONTORSON-LE-SEQUOIA.csv" -Delimiter ";"
    $listSacib01Quai = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-QUAI-SUD.csv" -Delimiter ";"
    $listSacib01Roch = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ROCH-PIERRE-LOUIS.csv" -Delimiter ";"
    $listSacib01Rossel = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ROSSEL.csv" -Delimiter ";"
    $listSacib01Roz = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ROZ-SUR-COUESNON-LE-CLOS-DE-LA-GRANGE.csv" -Delimiter ";"
    $listSacib01Rue = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-RUE-DU-VALLON.csv" -Delimiter ";"
    $listSacib01StBriac = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-SAINT-BRIAC.csv" -Delimiter ";"
    $listSacib01StJouan = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-SAINT-JOUAN-DES-GUERETS-LES-VOILES-ROUGES.csv" -Delimiter ";"
    $listSacib01StJMeloir = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-SAINT-MELOIR-DES-ONDES.csv" -Delimiter ";"
    $listSacib01StBriacZac = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-ST-BRIAC-ZAC-DES-TOURELLES.csv" -Delimiter ";"
    $listSacib01Terrasses = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-TERRASSES-DE-RIVASSELOU.csv" -Delimiter ";"
    $listSacib01Villa = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-VILLA-EDOUARD-VII.csv" -Delimiter ";"
    $listSacib01Voileries = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-01-VOILERIES.csv" -Delimiter ";"
    $listSacib02AlleesHavre = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-ALLEES-DU-HAVRE.csv" -Delimiter ";"
    $listSacib02AlleesPort= import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-ALLEES-DU-PORT.csv" -Delimiter ";"
    $listSacib02Cancale = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-CANCALE-BELLE-BRISE.csv" -Delimiter ";"
    $listSacib02Clos = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-CLOS-BASTILLE.csv" -Delimiter ";"
    $listSacib02Cote = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-COTE-DOCKS.csv" -Delimiter ";"
    $listSacib02Dinan = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-DINAN.csv" -Delimiter ";"
    $listSacib02Dinard = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-DINARD.csv" -Delimiter ";"
    $listSacib02Division = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-DIVISION-TERRAIN-LE-BIGNON.csv" -Delimiter ";"
    $listSacib02Dol = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-DOL-DE-BRETAGNE.csv" -Delimiter ";"
    $listSacib02Fontaine = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-FONTAINE-AUX-VAIS.csv" -Delimiter ";"
    $listSacib02Frehel = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-FREHEL-LA-CARQUOIS.csv" -Delimiter ";"
    $listSacib02Havre = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-HAVRE-FLEURY.csv" -Delimiter ";"
    $listSacib02Hirel = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-HIREL.csv" -Delimiter ";"
    $listSacib02KerLouis = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-KER-LOUIS-QUEBEC.csv" -Delimiter ";"
    $listSacib02LaGouesniere = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-LA-GOUESNIERE.csv" -Delimiter ";"
    $listSacib02Newquay = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-NEWQUAY.csv" -Delimiter ";"
    $listSacib02Patios = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-PATIOS-DU-HAVRE.csv" -Delimiter ";"
    $listSacib02Pleurtuit = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-PLEURTUIT.csv" -Delimiter ";"
    $listSacib02Pontorson = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-PONTORSON-LE-SEQUOIA.csv" -Delimiter ";"
    $listSacib02Roch = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-ROCH-PIERRE-LOUIS.csv" -Delimiter ";"
    $listSacib02Roz = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-ROZ-SUR-COUESNOU.csv" -Delimiter ";"
    $listSacib02StBriac = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-SAINT-BRIAC.csv" -Delimiter ";"
    $listSacib02StJouan = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-SAINT-JOUAN-DES-GUERETS-LES-VOILES-ROUGES.csv" -Delimiter ";"
    $listSacib02StMeloir = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-SAINT-MELOIR-DES-ONDES.csv" -Delimiter ";"
    $listSacib02StBriacChemin = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-ST-BRIAC-CHEMIN-DES-TOURELLES.csv" -Delimiter ";"
    $listSacib02Ville = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-02-VILLE-ES-NONAIS.csv" -Delimiter ";"
    $listSacib04Dinan = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-DINAN-PARC-DU-COMTE-DE-LA-GARAYE.csv" -Delimiter ";"
    $listSacib04Flers = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-FLERS-LES-JARDINS-DE-LORANGERIE.csv" -Delimiter ";"
    $listSacib04Frehel = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-FREHEL-LA-CARQUOIS.csv" -Delimiter ";"
    $listSacib04Havre = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-HAVRE-FLEURY.csv" -Delimiter ";"
    $listSacib04Pleudihen = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-PLEUDIHEN-SUR-RANCE-VAL-DORIENT.csv" -Delimiter ";"
    $listSacib04Pontorson = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-PONTORSON-LE-SEQUOIA.csv" -Delimiter ";"
    $listSacib04StJouan = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-SAINT-JOUAN-DES-GUERETS-LES-VOILES-ROUGES.csv" -Delimiter ";"
    $listSacib04StBriac = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-ST-BRIAC.csv" -Delimiter ";"
    $listSacib04Villa = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-04-VILLA-FLORIANE-ANCIEUX.csv" -Delimiter ";"
    $listSacib05Cancale = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\\Operations\OpMortes-SACIB-05-CANCALE-BELLE-PRISE.csv" -Delimiter ";"
    $listSacib05Flers = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-05-FLERS-LES-JARDINS-DE-LORANGERIE.csv" -Delimiter ";"
    $listSacib05Pontorson = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\Operations\OpMortes-SACIB-05-PONTORSON-LE-SEQUOIA.csv" -Delimiter ";"

    
    $tablListCsv = @($listSacib01AlleesHavre, $listSacib01AlleesPort, $listSacib01Argonautes, $listSacib01Balue, $listSacib01Bayaderes, $listSacib01Bellavista, $listSacib01Bignon, $listSacib01Centre, $listSacib01Chateau,
    $listSacib01ClosBastille, $listSacib01ClosNeuf, $listSacib01Compta, $listSacib01Cote, $listSacib01Dinan, $listSacib01Division, $listSacib01Domaine, $listSacib01Doutreleau, $listSacib01EugeneHerpin, $listSacib01Flers,
    $listSacib01Fontaine, $listSacib01Fontenelles, $listSacib01Frehel, $listSacib01Grande, $listSacib01Habert, $listSacib01Hameau, $listSacib01Havre, $listSacib01Ker, $listSacib01Gouesniere, $listSacib01LePont,
    $listSacib01Liloe, $listSacib01Nautica, $listSacib01Nestor, $listSacib01Newquay, $listSacib01Odyssee, $listSacib01Opalines, $listSacib01Patios, $listSacib01Pleurtuit, $listSacib01Pontorson, $listSacib01Quai,
    $listSacib01Roch, $listSacib01Rossel, $listSacib01Roz, $listSacib01Rue, $listSacib01StBriac,  $listSacib01StJouan, $listSacib01StJMeloir, $listSacib01StBriacZac, $listSacib01Terrasses, $listSacib01Villa,
    $listSacib01Voileries, $listSacib02AlleesHavre, $listSacib02AlleesPort, $listSacib02Cancale, $listSacib02Clos, $listSacib02Cote, $listSacib02Dinan, $listSacib02Dinard, $listSacib02Division, $listSacib02Dol,
    $listSacib02Fontaine, $listSacib02Frehel, $listSacib02Havre, $listSacib02Hirel, $listSacib02KerLouis, $listSacib02LaGouesniere, $listSacib02Newquay, $listSacib02Patios, $listSacib02Pleurtuit, $listSacib02Pontorson,
    $listSacib02Roch, $listSacib02Roz, $listSacib02StBriac, $listSacib02StJouan, $listSacib02StMeloir, $listSacib02StBriacChemin, $listSacib02Ville, $listSacib04Dinan, $listSacib04Flers, $listSacib04Frehel,
    $listSacib04Havre, $listSacib04Pleudihen, $listSacib04Pontorson, $listSacib04StJouan, $listSacib04StBriac, $listSacib04Villa, $listSacib05Cancale, $listSacib05Flers, $listSacib05Pontorson)
    
    Write-Host $tablListCsv.length " fichiers récupérés avec succés !" -ForegroundColor Green;

}
catch {
    Write-Host "Un ou des fichiers CSV de Sacib est(sont) introuvable(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}


#Récupère les listes de NANTES du site Sharepoint
Try {
    $lSacib01AlleesHavre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01ALLEESDUHAVRESNCBATIMALO")
    $lSacib01AlleesPort = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01ALLEESDUPORT")
    $lSacib01Argonautes = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01ARGONAUTESSARLLOTIMALO")
    $lSacib01Balue = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01BALUEDOUTRELEAU")
    $lSacib01Bayaderes = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01BAYADERES")
    $lSacib01Bellavista = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01BELLAVISTA")
    $lSacib01Bignon = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01BIGNON")
    $lSacib01Centre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01CENTREDIMAGERIEMEDICALESNCBATIMALO")
    $lSacib01Chateau = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01CHATEAUMALONORD")
    $lSacib01ClosBastille = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01CLOSBASTILLE")
    $lSacib01ClosNeuf = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01CLOSNEUF")
    $lSacib01Comptagesma = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01COMPTAGESMA")
    $lSacib01CoteDocks = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01COTEDOCKSSNCCLERMONT")
    $lSacib01Dinan = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01DINAN")
    $lSacib01Division = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01DIVISIONTERRAINLEBIGNON")
    $lSacib01Domaine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01DOMAINEDUMOULIN")
    $lSacib01Doutreleau = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01DOUTRELEAUNOUVEL")
    $lSacib01Eugene = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01EUGENEHERPIN")
    $lSacib01Flers = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01FLERSLESJARDINSDELORANGERIE")
    $lSacib01Fontaine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01FONTAINEAUVAIS")
    $lSacib01Fontenelles = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01FONTENELLESIMPASSEDUMORSE")
    $lSacib01Frehel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01FREHELLACARQUOIS")
    $lSacib01Grande = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01GRANDEBARONNIE")
    $lSacib01Habert = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01HABERT")
    $lSacib01Hameau = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01HAMEAUDELAFONTAINESNCBATIMALO")
    $lSacib01Havre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01HAVREFLEURY")
    $lSacib01Ker = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01KERLOUISQUEBEC")
    $lSacib01Gouesniere = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01LAGOUESNIERE")
    $lSacib01Pont = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01LEPONT")
    $lSacib01Liloe = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01LILOE2")
    $lSacib01Nautica = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01NAUTICA")
    $lSacib01Nestor = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01NESTOR")
    $lSacib01Newquay = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01NEWQUAY")
    $lSacib01Odyssee = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01ODYSSEE")
    $lSacib01Opalines = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01OPALINES")
    $lSacib01Patios = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01PATIOSDUHAVRE")
    $lSacib01Pleurtuit = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01PLEURTUIT")
    $lSacib01Pontorson = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01PONTORSONLESEQUOIA")
    $lSacib01Quai = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01QUAISUD")
    $lSacib01Roch = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01ROCHPIERRELOUIS")
    $lSacib01Rossel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01ROSSEL")
    $lSacib01Roz = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01ROZSURCOUESNONLECLOSDELAGRANGE")
    $lSacib01Rue = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01RUEDUVALLON")
    $lSacib01StBriac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01SAINTBRIAC")
    $lSacib01StJouan = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01SAINTJOUANDESGUERETSLESVOILESROUGES")
    $lSacib01StMeloir = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01SAINTMELOIRDESONDES")
    $lSacib01StBriacZac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01STBRIACZACDESTOURELLES")
    $lSacib01Terrasses = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01TERRASSESDERIVASSELOU")
    $lSacib01Villa = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01VILLAEDOUARDVII")
    $lSacib01Voileries = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01VOILERIES")
    $lSacib02AlleesHavre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02ALLEESDUHAVRE")
    $lSacib02AlleesPort = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02ALLEESDUPORT")
    $lSacib02Cancale = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02CANCALEBELLEBRISE")
    $lSacib02Clos = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02CLOSBASTILLE")
    $lSacib02Cote = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02COTEDOCKS")
    $lSacib02Dinan = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02DINAN")
    $lSacib02Dinard = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02DINARD")
    $lSacib02Division = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02DIVISIONTERRAINLEBIGNON")
    $lSacib02Dol = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02DOLDEBRETAGNE")
    $lSacib02Fontaine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02FONTAINEAUXVAIS")
    $lSacib02Frehel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02FREHELLACARQUOIS")
    $lSacib02Havre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02HAVREFLEURY")
    $lSacib02Hirel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02HIREL")
    $lSacib02KerLouis = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02KERLOUISQUEBEC")
    $lSacib02Gouesniere = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02LAGOUESNIERE")
    $lSacib02Newquay = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02NEWQUAY")
    $lSacib02Patios = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02PATIOSDUHAVRE")
    $lSacib02Pleurtuit = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02PLEURTUIT")
    $lSacib02Pontorson = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02PONTORSONLESEQUOIA")
    $lSacib02Roch = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02ROCHPIERRELOUIS")
    $lSacib02Roz = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02ROZSURCOUESNOU")
    $lSacib02StBriac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02SAINTBRIAC")
    $lSacib02StJouan = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02SAINTJOUANDESGUERETSLESVOILESROUGES")
    $lSacib02StMeloir = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02SAINTMELOIRDESONDES")
    $lSacib02StBriacChemin = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02STBRIACCHEMINDESTOURELLES")
    $lSacib02Ville = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02VILLEESNONAIS")
    $lSacib04Dinan = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04DINANPARCDUCOMTEDELAGARAYE")
    $lSacib04Flers = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04FLERSLESJARDINSDELORANGERIE")
    $lSacib04Frehel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04FREHELLACARQUOIS")
    $lSacib04Havre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04HAVREFLEURY")
    $lSacib04Pleudihen = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04PLEUDIHENSURRANCEVALDORIENT")
    $lSacib04Pontorson = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04PONTORSONLESEQUOIA")
    $lSacib04StJouan = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04SAINTJOUANDESGUERETSLESVOILESROUGES")
    $lSacib04StBriac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04STBRIAC")
    $lSacib04Villa = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04VILLAFLORIANEANCIEUX")
    $lSacib05Cancale = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB05CANCALEBELLEPRISE")
    $lSacib05Flers = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB05FLERSLESJARDINSDELORANGERIE")
    $lSacib05Pontorson = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB05PONTORSONLESEQUOIA")
  

    $tablListSite = @($lSacib01AlleesHavre, $lSacib01AlleesPort, $lSacib01Argonautes, $lSacib01Balue, $lSacib01Bayaderes, $lSacib01Bellavista, $lSacib01Bignon, $lSacib01Centre, $lSacib01Chateau, $lSacib01ClosBastille,
    $lSacib01ClosNeuf, $lSacib01Comptagesma, $lSacib01CoteDocks, $lSacib01Dinan, $lSacib01Division, $lSacib01Domaine, $lSacib01Doutreleau, $lSacib01Eugene, $lSacib01Flers, $lSacib01Fontaine, $lSacib01Fontenelles,
    $lSacib01Frehel, $lSacib01Grande, $lSacib01Habert, $lSacib01Hameau, $lSacib01Havre, $lSacib01Ker, $lSacib01Gouesniere, $lSacib01Pont, $lSacib01Liloe, $lSacib01Nautica, $lSacib01Nestor, $lSacib01Newquay,
    $lSacib01Odyssee, $lSacib01Opalines, $lSacib01Patios, $lSacib01Pleurtuit, $lSacib01Pontorson, $lSacib01Quai, $lSacib01Roch, $lSacib01Rossel, $lSacib01Roz, $lSacib01Rue, $lSacib01StBriac, $lSacib01StJouan,
    $lSacib01StMeloir, $lSacib01StBriacZac, $lSacib01Terrasses, $lSacib01Villa, $lSacib01Voileries, $lSacib02AlleesHavre, $lSacib02AlleesPort, $lSacib02Cancale, $lSacib02Clos, $lSacib02Cote, $lSacib02Dinan,
    $lSacib02Dinard, $lSacib02Division, $lSacib02Dol, $lSacib02Fontaine, $lSacib02Frehel, $lSacib02Havre, $lSacib02Hirel, $lSacib02KerLouis, $lSacib02Gouesniere, $lSacib02Newquay, $lSacib02Patios, $lSacib02Pleurtuit,
    $lSacib02Pontorson, $lSacib02Roch, $lSacib02Roz, $lSacib02StBriac, $lSacib02StJouan, $lSacib02StMeloir, $lSacib02StBriacChemin, $lSacib02Ville, $lSacib04Dinan, $lSacib04Flers, $lSacib04Frehel, $lSacib04Havre,
    $lSacib04Pleudihen, $lSacib04Pontorson, $lSacib04StJouan, $lSacib04StBriac, $lSacib04Villa, $lSacib05Cancale, $lSacib05Flers, $lSacib05Pontorson)
    
    #Write-Host $tablListSite.length " listes récupérés avec succés !" -ForegroundColor Green;

}
Catch {
    Write-Host "Une ou des listes n'a(ont) pas été trouvée(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}



#Ajout des données pour la liste SACIB
#Parcours les éléments
Try {
    for ($i=0; $i -le ($tablListSite.length -1); $i++) {
        #Parcours les éléments
        $r = 1;
        $counter = 0;
            foreach($item in $tablListCsv[$i])
            {
                $ni = $tablListSite[$i].items.Add();
                $ni["Titre"] = $r;
                $ni["CLASSEMENT"] = $item.CLASSEMENT;
                $ni["1ER NIVEAU"] = $item.'1ER NIVEAU';
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