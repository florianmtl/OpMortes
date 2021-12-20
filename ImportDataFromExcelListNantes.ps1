Add-PSSnapin Microsoft.SharePoint.PowerShell

#Paramètres
#Récupère le fichier CSV
Try {
    # NANTES
    $listNantes33 = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-33.csv" -Delimiter ";"
    $listNantes44Bassegoulaine = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-BASSEGOULAINE.csv" -Delimiter ";"
    $listNantes44Blain = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-BLAIN.csv" -Delimiter ";"
    $listNantes44Carquefou = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-CARQUEFOU.csv" -Delimiter ";"
    $listNantes44Chateaubriant = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-CHATEAUBRIANT.csv" -Delimiter ";"
    $listNantes44LaBaule = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-LABAULE.csv" -Delimiter ";"
    $listNantes44LaChapelle = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-LACHAPELLE.csv" -Delimiter ";"
    $listNantes44LaMontagne = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-LAMONTAGNE.csv" -Delimiter ";"
    $listNantes44NantesAtoI = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-NANTES-AtoI.csv" -Delimiter ";"
    $listNantes44NantesLeCtoLesR = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-NANTES-LE_CtoLES_R.csv" -Delimiter ";"
    $listNantes44NantesLesTtoV = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-NANTES-LES_TtoV.csv" -Delimiter ";"
    $listNantes44NortSurEdre = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-NORTSURERDRE.csv" -Delimiter ";"
    $listNantes44Orvault = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-ORVAULT.csv" -Delimiter ";"
    $listNantes44Pornic = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-PORNIC.csv" -Delimiter ";"
    $listNantes44Pornichet = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-PORNICHET.csv" -Delimiter ";"
    $listNantes44Reze = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-REZE.csv" -Delimiter ";"
    $listNantes44SainteLuce = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-SAINTELUCE.csv" -Delimiter ";"
    $listNantes44SaintHerblain = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-SAINTHERBLAIN.csv" -Delimiter ";"
    $listNantes44SaintSebastien = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-SAINTSEBASTIEN.csv" -Delimiter ";"
    $listNantes44Sautron = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-SAUTRON.csv" -Delimiter ";"
    $listNantes44StEtienne = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-44-STETIENNE.csv" -Delimiter ";"
    $listNantes49 = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-49.csv" -Delimiter ";"
    $listNantes56Baden = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-56-BADEN.csv" -Delimiter ";"
    $listNantes56Guidel = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-56-GUIDEL.csv" -Delimiter ";"
    $listNantes56Lorient = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-56-LORIENT.csv" -Delimiter ";"
    $listNantes56Queven = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-56-QUEVEN.csv" -Delimiter ";"
    $listNantes56Riantec = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-56-RIANTEC.csv" -Delimiter ";"
    $listNantes56SaintPierre = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-56-SAINTPIERRE.csv" -Delimiter ";"
    $listNantes56Vannes = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-56-VANNES.csv" -Delimiter ";"
    $listNantes78 = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-78.csv" -Delimiter ";"
    $listNantes94 = import-csv -Path "\\srvares\EF-OpMortes\NANTES\CSV\OpMortes-NANTES-94.csv" -Delimiter ";"



    Write-Host $tablListCsv.length " fichiers récupérés avec succés !" -ForegroundColor Green;

}
catch {
    Write-Host "Un ou des fichiers CSV de Nantes est(sont) introuvable(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}


#Récupère les listes de NANTES du site Sharepoint
Try {
    $lNantes33 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES33")
    $lNantes44BasseGoulaine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44BasseGoulaine")
    $lNantes44Blain = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Blain")
    $lNantes44Carquefou = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Carquefou")
    $lNantes44Chateaubriant = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Chateaubriant")
    $lNantes44LaBaule = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44LaBaule")
    $lNantes44LaChapelle = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44LaChapelle")
    $lNantes44LaMontagne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44LaMontagne")
    $lNantes44NantesAtoI = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NantesA_to_I")
    $lNantes44NantesLeCtoLesR = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NantesLe_C_to_Les_R")
    $lNantes44NantesLesTtoV = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NantesLes_T_to_V")
    $lNantes44NortSurErdre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44NortSurErdre")
    $lNantes44Orvault = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Orvault")
    $lNantes44Pornic = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Pornic")
    $lNantes44Pornichet = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Pornichet")
    $lNantes44Reze = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Reze")
    $lNantes44SainteLuce = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SainteLuce")
    $lNantes44SaintHerblain = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SaintHerblain")
    $lNantes44SaintSebastien = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44SaintSebastien")
    $lNantes44Sautron = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44Sautron")
    $lNantes44StEtienne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES44StEtienne")
    $lNantes49 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES49")
    $lNantes56Baden = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56Baden")
    $lNantes56Guidel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56Guidel")
    $lNantes56Lorient = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56Lorient")
    $lNantes56Queven = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56Queven")
    $lNantes56Riantec = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56Riantec")
    $lNantes56SaintPierre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56SaintPierre")
    $lNantes56Vannes = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES56Vannes")
    $lNantes78 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES78")
    $lNantes94 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/NANTES94")

    Write-Host $tablListSite.length " fichiers récupérés avec succés !" -ForegroundColor Green;

}
Catch {
    Write-Host "Une ou des listes n'a(ont) pas été trouvée(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}

    $tablListCsv = @($listNantes33, $listNantes44Bassegoulaine, $listNantes44Blain, $listNantes44Carquefou, $listNantes44Chateaubriant, $listNantes44LaBaule, $listNantes44LaChapelle, $listNantes44LaMontagne,
    $listNantes44NantesAtoI, $listNantes44NantesLeCtoLesR, $listNantes44NantesLesTtoV, $listNantes44NortSurEdre, $listNantes44Orvault, $listNantes44Pornic, $listNantes44Pornichet, $listNantes44Reze,
    $listNantes44SainteLuce, $listNantes44SaintHerblain, $listNantes44SaintSebastien, $listNantes44Sautron, $listNantes44StEtienne, $listNantes49, $listNantes56Baden, $listNantes56Guidel, $listNantes56Lorient,
    $listNantes56Queven, $listNantes56Riantec, $listNantes56SaintPierre, $listNantes56Vannes, $listNantes78, $listNantes94)

    $tablListSite = @($lNantes33, $lNantes44BasseGoulaine, $lNantes44Blain, $lNantes44Carquefou, $lNantes44Chateaubriant, $lNantes44LaBaule, $lNantes44LaChapelle, $lNantes44LaMontagne, $lNantes44NantesAtoI, 
    $lNantes44NantesLeCtoLesR, $lNantes44NantesLesTtoV, $lNantes44NortSurErdre, $lNantes44Orvault, $lNantes44Pornic, $lNantes44Pornichet, $lNantes44Reze, $lNantes44SainteLuce, $lNantes44SaintHerblain,
    $lNantes44SaintSebastien, $lNantes44Sautron, $lNantes44StEtienne, $lNantes49, $lNantes56Baden, $lNantes56Guidel, $lNantes56Lorient, $lNantes56Queven, $lNantes56Riantec, $lNantes56SaintPierre, $lNantes56Vannes,
    $lNantes78, $lNantes94)

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