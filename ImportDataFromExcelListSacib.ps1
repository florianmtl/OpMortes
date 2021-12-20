Add-PSSnapin Microsoft.SharePoint.PowerShell

#Paramètres
#Récupère le fichier CSV
Try {
    # SACIB
    $listSacib01AtoB = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-01-AtoB.csv" -Delimiter ";"
    $listSacib01CtoFl = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-01-CtoFL.csv" -Delimiter ";"
    $listSacib01FotoL = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-01-FOtoL.csv" -Delimiter ";"
    $listSacib01NtoR = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-01-NtoR.csv" -Delimiter ";"
    $listSacib01StoV = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-01-StoV.csv" -Delimiter ";"
    $listSacib02AtoDi = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-02-AtoDI.csv" -Delimiter ";"
    $listSacib02DotoR = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-02-DOtoR.csv" -Delimiter ";"
    $listSacib02StoV = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-02-StoV.csv" -Delimiter ";"
    $listSacib04 = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-04.csv" -Delimiter ";"
    $listSacib05 = import-csv -Path "\\srvares\EF-OpMortes\SACIB\CSV\OpMortes-SACIB-05.csv" -Delimiter ";"

    #Write-Host $tablListCsv.length " fichiers récupérés avec succés !" -ForegroundColor Green;

}
catch {
    Write-Host "Un ou des fichiers CSV de Sacib est(sont) introuvable(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}


#Récupère les listes de NANTES du site Sharepoint
Try {
    $lSacib01AtoB = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01A_to_B")
    $lSacib01CtoFl = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01C_to_Fl")
    $lSacib01FotoL = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01Fo_to_L")
    $lSacib01NtoR = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01N_to_R")
    $lSacib01StoV = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB01S_to_V")
    $lSacib02AtoDi = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02A_to_Di")
    $lSacib02DotoR = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02Do_to_R")
    $lSacib02StoV = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB02S_to_V")
    $lSacib04 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB04")
    $lSacib05 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/SACIB05")

    #Write-Host $tablListSite.length " listes récupérés avec succés !" -ForegroundColor Green;

}
Catch {
    Write-Host "Une ou des listes n'a(ont) pas été trouvée(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}

$tablListCsv = @($listSacib01AtoB, $listSacib01CtoFl, $listSacib01FotoL, $listSacib01NtoR, $listSacib01StoV, $listSacib02AtoDi, $listSacib02DotoR, $listSacib02StoV, $listSacib04, $listSacib05)
$tablListSite = @($lSacib01AtoB, $lSacib01CtoFl, $lSacib01FotoL, $lSacib01NtoR, $lSacib01StoV, $lSacib02AtoDi, $lSacib02DotoR, $lSacib02StoV, $lSacib04, $lSacib05)

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

                Write-Progress -Id 1 -Activity "Importation des données" -Status 'Progress' -PercentComplete (($counter / $tablListCsv[$i].Count) * 100);
            }
            Write-Host ((${r} -1), " ligne(s) ajoutée(s) à la liste", $tablListSite[$i]) -ForegroundColor Green;
        }
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
}
