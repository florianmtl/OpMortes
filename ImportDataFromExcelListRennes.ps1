Add-PSSnapin Microsoft.SharePoint.PowerShell

#Paramètres
#Récupère le fichier CSV
Try {
    # RENNES
    $listRennes14 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-14.csv" -Delimiter ";"
    $listRennes17 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-17.csv" -Delimiter ";"
    $listRennes22 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-22.csv" -Delimiter ";"
    $listRennes29 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-29.csv" -Delimiter ";"
    $listRennes35BainDeBretagne = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-BAINDEBRETAGNE.csv" -Delimiter ";"
    $listRennes35Betton = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-BETTON.csv" -Delimiter ";"
    $listRennes35Boisgervilly = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-BOISGERVILLY.csv" -Delimiter ";"
    $listRennes35Bruz = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-BRUZ.csv" -Delimiter ";"
    $listRennes35Cesson = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-CESSON.csv" -Delimiter ";"
    $listRennes35Chantepie = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-CHANTEPIE.csv" -Delimiter ";"
    $listRennes35Domloup = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-DOMLOUP.csv" -Delimiter ";"
    $listRennes35Etrelle = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-ETRELLE.csv" -Delimiter ";"
    $listRennes35Gosne = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-GOSNE.csv" -Delimiter ";"
    $listRennes35Janze = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-JANZE.csv" -Delimiter ";"
    $listRennes35LaChapelle = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-LACHAPELLE.csv" -Delimiter ";"
    $listRennes35LeRheu = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-LERHEU.csv" -Delimiter ";"
    $listRennes35Lhermitage = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-LHERMITAGE.csv" -Delimiter ";"
    $listRennes35Liffre = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-LIFFRE.csv" -Delimiter ";"
    $listRennes35MezieresSousC = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-MEZIERESSOUSC.csv" -Delimiter ";"
    $listRennes35MontaubanDeBr = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-MONTAUBANDEBR.csv" -Delimiter ";"
    $listRennes35Noyal = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-NOYAL.csv" -Delimiter ";"
    $listRennes35Parame = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-PARAME.csv" -Delimiter ";"
    $listRennes35RennesAtoLe = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-RENNES-AtoLE.csv" -Delimiter ";"
    $listRennes35RennesLesJtoLesO = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-RENNES-LESJtoLESO.csv" -Delimiter ";"
    $listRennes35RennesLesPtoR = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-RENNES-LESPtoR.csv" -Delimiter ";"
    $listRennes35RennesTtoV = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-RENNES-TtoV.csv" -Delimiter ";"
    $listRennes35Romagne = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-ROMAGNE.csv" -Delimiter ";"
    $listRennes35SaintBriac = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-SAINTBRIAC.csv" -Delimiter ";"
    $listRennes35SaintGregoireAtoE = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-SAINTGREGOIRE-AtoE.csv" -Delimiter ";"
    $listRennes35SaintGregoireLtoP = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-SAINTGREGOIRE-LtoP.csv" -Delimiter ";"
    $listRennes35SaintJacques = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-SAINTJACQUES.csv" -Delimiter ";"
    $listRennes35SaintMalo = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-35-SAINTMALO.csv" -Delimiter ";"
    $listRennes44Bur = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-44BUR.csv" -Delimiter ";"
    $listRennes53 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-53.csv" -Delimiter ";"
    $listRennes56Bur = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-56BUR.csv" -Delimiter ";"
    $listRennes72 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-72.csv" -Delimiter ";"
    $listRennes74 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-74.csv" -Delimiter ";"
    $listRennes79 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-79.csv" -Delimiter ";"
    $listRennes85 = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\OpMortes-RENNES-85.csv" -Delimiter ";"
     
    $tablListCsv = @($listRennes14, $listRennes17, $listRennes22, $listRennes29, $listRennes35BainDeBretagne, $listRennes35Betton, $listRennes35Boisgervilly, $listRennes35Bruz, $listRennes35Cesson,
    $listRennes35Chantepie, $listRennes35Domloup, $listRennes35Etrelle, $listRennes35Gosne, $listRennes35Janze, $listRennes35LaChapelle, $listRennes35LeRheu, $listRennes35Lhermitage, $listRennes35Liffre,
    $listRennes35MezieresSousC, $listRennes35MontaubanDeBr, $listRennes35Noyal, $listRennes35Parame, $listRennes35RennesAtoLe, $listRennes35RennesLesJtoLesO, $listRennes35RennesLesPtoR, $listRennes35RennesTtoV,
    $listRennes35Romagne, $listRennes35SaintBriac, $listRennes35SaintGregoireAtoE, $listRennes35SaintGregoireLtoP, $listRennes35SaintJacques, $listRennes35SaintMalo, $listRennes44Bur, $listRennes53, $listRennes56Bur,
    $listRennes72, $listRennes74, $listRennes79, $listRennes85)

    Write-Host $tablListCsv.length " fichiers récupérés avec succés !" -ForegroundColor Green;

}
catch {
    Write-Host "Un ou des fichiers CSV de Rennes est(sont) introuvable(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}


#Récupère les listes de NANTES du site Sharepoint
Try {
    $lRennes14 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES14")
    $lRennes17 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17")
    $lRennes22 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES22")
    $lRennes29 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29")
    $lRennes35BainDeBretagne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BainDeBretagne")
    $lRennes35Betton = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Betton")
    $lRennes35BoisGervilly = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BoisGervilly")
    $lRennes35Bruz = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Bruz")
    $lRennes35Cesson = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Cesson")
    $lRennes35Chantepie = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Chantepie")
    $lRennes35Domloup = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Domloup")
    $lRennes35Etrelle = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Etrelle")
    $lRennes35Gosne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Gosne")
    $lRennes35Janze = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Janze")
    $lRennes35LaChapelle = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LaChapelle")
    $lRennes35LeRheu = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LeRheu")
    $lRennes35Lhermitage = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Lhermitage")
    $lRennes35Liffre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Liffre")
    $lRennes35MezieresSousC = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35MezieresSousC")
    $lRennes35MontaubanDeBr = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35MontaubanDeBr")
    $lRennes35Noyal = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Noyal")
    $lRennes35Parame = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Parame")
    $lRennes35RennesAtoLe = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RennesA_to_Le")
    $lRennes35RennesLesJtoLesO = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RennesLes_J_to_Les_O")
    $lRennes35RennesLesPtoR = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RennesLes_P_to_R")
    $lRennes35RennesTtoV = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RennesT_to_V")
    $lRennes35Romagne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35Romagne")
    $lRennes35SaintBriac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SaintBriac")
    $lRennes35SaintGregroireAtoE = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SaintGregoireA_to_E")
    $lRennes35SaintGregoireLtoP = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SaintGregoireL_to_P")
    $lRennes35SaintJacques = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SaintJacques")
    $lRennes35SaintMalo = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SaintMalo")
    $lRennes44Bur = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44Bur")
    $lRennes53 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES53")
    $lRennes56Bur = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES56Bur")
    $lRennes72 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES72")
    $lRennes74 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES74")
    $lRennes79 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES79")
    $lRennes85 = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES85")


    $tablListSite = @($lRennes14, $lRennes17, $lRennes22, $lRennes29, $lRennes35BainDeBretagne, $lRennes35Betton, $lRennes35BoisGervilly, $lRennes35Bruz, $lRennes35Cesson, $lRennes35Chantepie, $lRennes35Domloup,
    $lRennes35Etrelle, $lRennes35Gosne, $lRennes35Janze, $lRennes35LaChapelle, $lRennes35LeRheu, $lRennes35Lhermitage, $lRennes35Liffre, $lRennes35MezieresSousC, $lRennes35MontaubanDeBr, $lRennes35Noyal,
    $lRennes35Parame, $lRennes35RennesAtoLe, $lRennes35RennesLesJtoLesO, $lRennes35RennesLesPtoR, $lRennes35RennesTtoV, $lRennes35Romagne, $lRennes35SaintBriac, $lRennes35SaintGregroireAtoE,
    $lRennes35SaintGregoireLtoP, $lRennes35SaintJacques, $lRennes35SaintMalo, $lRennes44Bur, $lRennes53, $lRennes56Bur, $lRennes72, $lRennes74, $lRennes79, $lRennes85)

    Write-Host $tablListSite.length " fichiers récupérés avec succés !" -ForegroundColor Green;

}
Catch {
    Write-Host "Une ou des listes n'a(ont) pas été trouvée(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}


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

            Write-Progress -Id 1 -Activity "Importation des données" -Status 'Progress' -PercentComplete (($counter / $tablListCsv[$i].Count) * 100);
        }
        Write-Host ((${r} -1), " ligne(s) ajoutée(s) à la liste", $tablListSite[$i]) -ForegroundColor Green;
    }
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
}