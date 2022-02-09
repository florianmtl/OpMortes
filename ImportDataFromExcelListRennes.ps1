Add-PSSnapin Microsoft.SharePoint.PowerShell

#Paramètres
#Récupère le fichier CSV
Try {
    # RENNES
    $listRennes14Demouville = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-14-DEMOUVILLE-THYSSEN-DEMOUVILLE.csv" -Delimiter ";"
    $listRennes17AytreLePatio = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-AYTRE-LE-PATIO-DES-TILLEULS.csv" -Delimiter ";"
    $listRennes17AytreVillas = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-AYTRE-VILLAS-TILIA.csv" -Delimiter ";"
    $listRennes17Clavette = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-CLAVETTE-LE-TENNIS.csv" -Delimiter ";"
    $listRennes17Esnandes = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-ESNANDES-ESNANDES.csv" -Delimiter ";"
    $listRennes17LaRochelleAzurea = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-LAROCHELLE-AZUREA.csv" -Delimiter ";"
    $listRennes17LaRochelleHorizon = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-LAROCHELLE-HORIZON-MER.csv" -Delimiter ";"
    $listRennes17LaRochelleLaRochelle = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-LAROCHELLE-LAROCHELLE-EINSTEIN.csv" -Delimiter ";"
    $listRennes17LaRochelleLeGlobe = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-LAROCHELLE-LE-GLOBE-TROTTER.csv" -Delimiter ";"
    $listRennes17LaRochelleVilla = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-LAROCHELLE-VILLA-GARANCE.csv" -Delimiter ";"
    $listRennes17Nieul = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-NIEUL-SUR-MER-VILLA-ROSA.csv" -Delimiter ";"
    $listRennes17Perigny = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-17-PERIGNY-LES-JARDINS-DES-ACANTHES.csv" -Delimiter ";"
    $listRennes22Lavanllay = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-22-LANVALLAY-LE-CLOS-DES-ORMEAUX.csv" -Delimiter ";"
    $listRennes22Lehon = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-22-LEHON-LE-CLOS-TRIARD.csv" -Delimiter ";"
    $listRennes22PleneufValand = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-22-PLENEUF-VALAND-COTE-MER.csv" -Delimiter ";"
    $listRennes22PleneufLesAigues = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-22-PLENEUF-VALAND-LES-AIGUES-MARINES.csv" -Delimiter ";"
    $listRennes22SaintCast = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-22-SAINT-CAST-LES-PIERRES-SONNANTES.csv" -Delimiter ";"
    $listRennes22Tregueux = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-22-TREGUEUX-BREZILLET.csv" -Delimiter ";"
    $listRennes29Brest = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-BREST-SAINT-PIERRE.csv" -Delimiter ";"
    $listRennes29CarantecVillas = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-CARANTEC-LES-VILLAS-DE-KERGRIST.csv" -Delimiter ";"
    $listRennes29CarantecParc = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-CARANTEC-PARC-OCEAN.csv" -Delimiter ";"
    $listRennes29Guilers = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-GUILERS-LE-CLOS-VALENTIN.csv" -Delimiter ";"
    $listRennes29Guipavas = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-GUIPAVAS-LES-RESIDENCES-SAINT-EXUPERY.csv" -Delimiter ";"
    $listRennes29Landivisiau = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-LANDIVISIAU-PARC-LANDIVISIAU.csv" -Delimiter ";"
    $listRennes29LeDrennec = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-LEDRENNEC-LES-JARDINS-DADRIEN.csv" -Delimiter ";"
    $listRennes29LocaMaria = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-LOCMARIA-PLOUZA-LES-JARDINS-DE-LOCMARIA.csv" -Delimiter ";"
    $listRennes29Plouarzel = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-PLOUARZEL-KERVEN.csv" -Delimiter ";"
    $listRennes29Plougonvelin = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-PLOUGONVELIN-LES-JARDINS-DU-TREZ-HIR.csv" -Delimiter ";"
    $listRennes29PontLabbe = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-PONT-L-ABBE-TERRASSES-ETANG.csv" -Delimiter ";"
    $listRennes29Quimper = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-29-QUIMPER-DOMAINE-CENTRE.csv" -Delimiter ";"
    $listRennes35BainDeBretagneForum = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BAINDEBRETAGNE-LE-FORUM.csv" -Delimiter ";"
    $listRennes35BaindeBretagneResidence = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BAINDEBRETAGNE-RESIDENCE-DES-TANNEURS.csv" -Delimiter ";"
    $listRennes35BettonLesCapucines = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BETTON-LES-CAPUCINES.csv" -Delimiter ";"
    $listRennes35BettonLesCinq = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BETTON-LES-CINQ-ILES.csv" -Delimiter ";"
    $listRennes35BettonPresquille = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BETTON-PRESQU-ILLE.csv" -Delimiter ";"
    $listRennes35BettonVillaCantalina = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BETTON-VILLA-CANTALINA.csv" -Delimiter ";"
    $listRennes35BettonVillaOdalie = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BETTON-VILLA-ODALIE.csv" -Delimiter ";"
    $listRennes35Boisgervilly = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BOISGERVILLY-LANCELOT-DU-LAC.csv" -Delimiter ";"
    $listRennes35BruzIneo = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BRUZ-INEO.csv" -Delimiter ";"
    $listRennes35BruzJardins = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-BRUZ-LES-JARDINS-DE-BLOSSAC.csv" -Delimiter ";"
    $listRennes35CessonSevigneCesson = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-CESSON-SEVIGNE-CESSON-MAISON-MEDICALE.csv" -Delimiter ";"
    $listRennes35CessonSevigneNet = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-CESSON-SEVIGNE-NET-PLUS.csv" -Delimiter ";"
    $listRennes35ChantepieLeDomaine = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-CHANTEPIE-LE-DOMAINE-DU-CANAL.csv" -Delimiter ";"
    $listRennes35ChantepieLesCoquelicots = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-CHANTEPIE-LES-COQUELICOTS.csv" -Delimiter ";"
    $listRennes35ChantepieVilla = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-CHANTEPIE-VILLA-ABELIA.csv" -Delimiter ";"
    $listRennes35Domloup = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-DOMLOUP-DOMLOUP.csv" -Delimiter ";"
    $listRennes35Etrelle = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-ETRELLE-ETRELLE-VINCI-ENERGIES.csv" -Delimiter ";"
    $listRennes35Gosne = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-GOSNE-LES-PORTES-DOUEES.csv" -Delimiter ";"
    $listRennes35Janze = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-JANZE-LES-AMBROISINES.csv" -Delimiter ";"
    $listRennes35LaChapelleRueDePace = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LACHAPELLE-RUE-DE-PACE.csv" -Delimiter ";"
    $listRennes35LaChapelleRueLechlade = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LACHAPELLE-RUE-LECHLADE.csv" -Delimiter ";"
    $listRennes35LaChapelleVilla = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LACHAPELLE-VILLA-LAURENA.csv" -Delimiter ";"
    $listRennes35LeRheuJardins = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LERHEU-JARDINS-DADELE.csv" -Delimiter ";"
    $listRennes35LeRheuLeGrand = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LERHEU-LE-GRAND-JARDIN.csv" -Delimiter ";"
    $listRennes35LeRheuNeventi = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LERHEU-NEVENTI.csv" -Delimiter ";"
    $listRennes35LeRheuRueDeRennes = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LERHEU-RUE-DE-RENNES.csv" -Delimiter ";"
    $listRennes35LeRheuThyssen = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LERHEU-THYSSEN-LE-RHEU.csv" -Delimiter ";"
    $listRennes35LeRheuVillas = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LERHEU-VILLAS-MELIES.csv" -Delimiter ";"
    $listRennes35Lhermitage = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LHERMITAGE-LES-AQUARELLES.csv" -Delimiter ";"
    $listRennes35LiffreParc = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LIFFRE-PARC-DES-ETANGS.csv" -Delimiter ";"
    $listRennes35LiffreRose = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-LIFFRE-ROSE-ARMOR.csv" -Delimiter ";"
    $listRennes35MeziereSousCLaGrande = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-MEZIERESSOUSC-LA-GRANDE-PREE.csv" -Delimiter ";"
    $listRennes35MeziereSousCLaPree = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-MEZIERESSOUSC-LA-PREE-DU-PETIT-BOIS.csv" -Delimiter ";"
    $listRennes35MontaubanDeBr = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-MONTAUBANDEBR-SAINT-ELOI.csv" -Delimiter ";"
    $listRennes35Noyal = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-NOYAL-LES-HORTENSIAS.csv" -Delimiter ";"
    $listRennes35Parame = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-PARAME.csv" -Delimiter ";"
    $listRennes35RennesAdiph = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-ADIPH.csv" -Delimiter ";"
    $listRennes35RennesAvenue = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-AVENUE-MAGINOT.csv" -Delimiter ";"
    $listRennes35RennesCap = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-CAP-NORD.csv" -Delimiter ";"
    $listRennes35RennesCarre = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-CARRE-DART.csv" -Delimiter ";"
    $listRennes35RennesCassini = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-CASSINI.csv" -Delimiter ";"
    $listRennes35RennesCastel = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-CASTEL-RIVIERA.csv" -Delimiter ";"
    $listRennes35RennesCoeur = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-COEUR-DE-VILLE.csv" -Delimiter ";"
    $listRennes35RennesDemat = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-DEMAT.csv" -Delimiter ";"
    $listRennes35RennesEolysII = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-EOLYSII.csv" -Delimiter ";"
    $listRennes35RennesLaVisitation = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LA-VISITATION.csv" -Delimiter ";"
    $listRennes35RennesLeMurano = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LE-MURANO.csv" -Delimiter ";"
    $listRennes35RennesLeNoven = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LE-NOVEN.csv" -Delimiter ";"
    $listRennes35RennesLeSextant = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LE-SEXTANT.csv" -Delimiter ";"
    $listRennes35RennesLesJardinsChat = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LES-JARDINS-DE-CHATILLON.csv" -Delimiter ";"
    $listRennes35RennesLesJardinsNero = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LES-JARDINS-DE-NEROLI.csv" -Delimiter ";"
    $listRennes35RennesLesOpalines = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LES-OPALINES.csv" -Delimiter ";"
    $listRennes35RennesLesPrairies = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LES-PRAIRIES-DE-LILLE.csv" -Delimiter ";"
    $listRennes35RennesLesRives = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-LES-RIVES-DE-TASSIGNY.csv" -Delimiter ";"
    $listRennes35RennesMadison = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-MADISON-PARC.csv" -Delimiter ";"
    $listRennes35RennesOsiris = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-OSIRIS.csv" -Delimiter ";"
    $listRennes35RennesResidence = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-RESIDENCE-DE-VINCI.csv" -Delimiter ";"
    $listRennes35RennesRue = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-RUE-LEGRAVERAND.csv" -Delimiter ";"
    $listRennes35RennesTerranova = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-TERRANOVA.csv" -Delimiter ";"
    $listRennes35RennesVillaCamilla = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-VILLA-CAMILLA.csv" -Delimiter ";"
    $listRennes35RennesVillaDeVinci = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-RENNES-VILLA-DE-VINCI.csv" -Delimiter ";"
    $listRennes35Romagne = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-ROMAGNE-LE-CLOS-DES-SAULES.csv" -Delimiter ";"
    $listRennes35StBriac = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTBRIAC-LES-ROCHES-DOUVRES.csv" -Delimiter ";"
    $listRennes35StGregoireAxis = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTGREGOIRE-AXIS.csv" -Delimiter ";"
    $listRennes35StGregoireBpo = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTGREGOIRE-BPO.csv" -Delimiter ";"
    $listRennes35StGregoireEdonia = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTGREGOIRE-EDONIA.csv" -Delimiter ";"
    $listRennes35StGregoireBoutiere = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTGREGOIRE-LA-BOUTIERE.csv" -Delimiter ";"
    $listRennes35StGregoireParcBroce = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTGREGOIRE-PARC-DE-BROCELIANDE.csv" -Delimiter ";"
    $listRennes35StGregoireParcEllena = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTGREGOIRE-PARC-ELLENA.csv" -Delimiter ";"
    $listRennes35StGregoirePoleMedical = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTGREGOIRE-POLE-MEDICAL-LA-BOUTIERE.csv" -Delimiter ";"
    $listRennes35StJacquesDarty = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTJACQUES-ACTIVLAND-DARTY.csv" -Delimiter ";"
    $listRennes35StJacquesMondialRelay = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTJACQUES-ACTIVLAND-MONDIAL-RELAY.csv" -Delimiter ";"
    $listRennes35StJacquesAdaggio = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTJACQUES-ADAGGIO-SOCIAL.csv" -Delimiter ";"
    $listRennes35StJacquesVilla = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTJACQUES-VILLA-GIULIA.csv" -Delimiter ";"
    $listRennes35StMaloBayaderes = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTMALO-LES-BAYADERES.csv" -Delimiter ";"
    $listRennes35StMaloMarines = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTMALO-LES-MARINES-DE-CHASLES.csv" -Delimiter ";"
    $listRennes35StMaloParc = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTMALO-PARC-DE-LHERMINE.csv" -Delimiter ";"
    $listRennes35StMaloSquare = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-35-SAINTMALO-SQUARE-ACADIE.csv" -Delimiter ";"
    $listRennes44BURCarquefou = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-CARQUEFOU-ATALIS.csv" -Delimiter ";"
    $listRennes44BURCoueron = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-COUERON-GLASSOLUTION.csv" -Delimiter ";"
    $listRennes44BURLaChapelle = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-LACHAPELLE-INEO.csv" -Delimiter ";"
    $listRennes44BURNantesEspace = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-NANTES-ESPACE-NEWTON.csv" -Delimiter ";"
    $listRennes44BURNantesEuropa = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-NANTES-EUROPA.csv" -Delimiter ";"
    $listRennes44BURNantesExalis = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-NANTES-EXALIS.csv" -Delimiter ";"
    $listRennes44BUROrvault = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-ORVAULT-BOIS-CESBRON.csv" -Delimiter ";"
    $listRennes44BURPornic = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-PORNIC-PORNIC.csv" -Delimiter ";"
    $listRennes44BURReze = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-REZE-PSA-REZE.csv" -Delimiter ";"
    $listRennes44BURStHerblainAsturia = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-SAINTHERBLAIN-ASTURIA.csv" -Delimiter ";"
    $listRennes44BURStHerblainExapole = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-SAINTHERBLAIN-EXAPOLE.csv" -Delimiter ";"
    $listRennes44BURStHerblainSunset = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-SAINTHERBLAIN-SUNSET.csv" -Delimiter ";"
    $listRennes44BURTrignac = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-44BUR-TRIGNAC-GRAND-CHAMPS.csv" -Delimiter ";"
    $listRennes53Change = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-53-CHANGE-SPIE-CHANGE.csv" -Delimiter ";"
    $listRennes56BURRieux = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-56BUR-RIEUX-VEOLIA-EAU.csv" -Delimiter ";"
    $listRennes72LeMans = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-72-LE-MANS-LE-GALILEE.csv" -Delimiter ";"
    $listRennes74Annecy = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-74-ANNECYLEVIEUX-PINSONS.csv" -Delimiter ";"
    $listRennes79Niort = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-79-NIORT-LE-CLOS-DES-TILLEULS.csv" -Delimiter ";"
    $listRennes85Fontenay = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-85-FONTENAYLECOM-LES-COLLIBERTS.csv" -Delimiter ";"
    $listRennes85LaRoche = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-85-LAROCHESURYON-RESIDENCE-ELLINE.csv" -Delimiter ";"
    $listRennes85LesSablesdol = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-85-LESSABLESDOL.csv" -Delimiter ";"
    $listRennes85SaintHilaire = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-85-SAINTHILAIRE-ST-HILAIRE-DU-RIEZ.csv" -Delimiter ";"
    $listRennes85StVincent = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-85-SAINTVINCENT-LE-SAINT-VINCENT.csv" -Delimiter ";"
    $listRennes85StGilles = import-csv -Path "\\srvares\EF-OpMortes\RENNES\CSV\Operations\OpMortes-RENNES-85-ST-GILLES-CROIX-POEME.csv" -Delimiter ";"
   
    $tablListCsv = @($listRennes14Demouville, $listRennes17AytreLePatio, $listRennes17AytreVillas, $listRennes17Clavette, $listRennes17Esnandes, $listRennes17LaRochelleAzurea, $listRennes17LaRochelleHorizon,
    $listRennes17LaRochelleLaRochelle, $listRennes17LaRochelleLeGlobe, $listRennes17LaRochelleVilla, $listRennes17Nieul, $listRennes17Perigny, $listRennes22Lavanllay, $listRennes22Lehon, $listRennes22PleneufValand, 
    $listRennes22PleneufLesAigues, $listRennes22SaintCast, $listRennes22Tregueux, $listRennes29Brest, $listRennes29CarantecVillas, $listRennes29CarantecParc, $listRennes29Guilers, $listRennes29Guipavas,
    $listRennes29Landivisiau, $listRennes29LeDrennec, $listRennes29LocaMaria, $listRennes29Plouarzel, $listRennes29Plougonvelin, $listRennes29PontLabbe, $listRennes29Quimper, $listRennes35BainDeBretagneForum,
    $listRennes35BaindeBretagneResidence, $listRennes35BettonLesCapucines, $listRennes35BettonLesCinq, $listRennes35BettonPresquille, $listRennes35BettonVillaCantalina, $listRennes35BettonVillaOdalie,
    $listRennes35Boisgervilly, $listRennes35BruzIneo, $listRennes35BruzJardins, $listRennes35CessonSevigneCesson, $listRennes35CessonSevigneNet, $listRennes35ChantepieLeDomaine, $listRennes35ChantepieLesCoquelicots,
    $listRennes35ChantepieVilla, $listRennes35Domloup, $listRennes35Etrelle, $listRennes35Gosne, $listRennes35Janze, $listRennes35LaChapelleRueDePace, $listRennes35LaChapelleRueLechlade, $listRennes35LaChapelleVilla,
    $listRennes35LeRheuJardins, $listRennes35LeRheuLeGrand, $listRennes35LeRheuNeventi, $listRennes35LeRheuRueDeRennes, $listRennes35LeRheuThyssen, $listRennes35LeRheuVillas, $listRennes35Lhermitage,
    $listRennes35LiffreParc, $listRennes35LiffreRose, $listRennes35MeziereSousCLaGrande, $listRennes35MeziereSousCLaPree, $listRennes35MontaubanDeBr, $listRennes35Noyal, $listRennes35Parame, $listRennes35RennesAdiph,
    $listRennes35RennesAvenue, $listRennes35RennesCap, $listRennes35RennesCarre, $listRennes35RennesCassini, $listRennes35RennesCastel, $listRennes35RennesCoeur, $listRennes35RennesDemat, $listRennes35RennesEolysII,
    $listRennes35RennesLaVisitation, $listRennes35RennesLeMurano, $listRennes35RennesLeNoven, $listRennes35RennesLeSextant, $listRennes35RennesLesJardinsChat, $listRennes35RennesLesJardinsNero,
    $listRennes35RennesLesOpalines, $listRennes35RennesLesPrairies, $listRennes35RennesLesRives, $listRennes35RennesMadison, $listRennes35RennesOsiris, $listRennes35RennesResidence, $listRennes35RennesRue,
    $listRennes35RennesTerranova, $listRennes35RennesVillaCamilla, $listRennes35RennesVillaDeVinci, $listRennes35Romagne, $listRennes35StBriac, $listRennes35StGregoireAxis, $listRennes35StGregoireBpo,
    $listRennes35StGregoireEdonia, $listRennes35StGregoireBoutiere, $listRennes35StGregoireParcBroce, $listRennes35StGregoireParcEllena, $listRennes35StGregoirePoleMedical, $listRennes35StJacquesDarty,
    $listRennes35StJacquesMondialRelay, $listRennes35StJacquesAdaggio, $listRennes35StJacquesVilla, $listRennes35StMaloBayaderes, $listRennes35StMaloMarines, $listRennes35StMaloParc, $listRennes35StMaloSquare,
    $listRennes44BURCarquefou, $listRennes44BURCoueron, $listRennes44BURLaChapelle, $listRennes44BURNantesEspace, $listRennes44BURNantesEuropa, $listRennes44BURNantesExalis, $listRennes44BUROrvault,
    $listRennes44BURPornic, $listRennes44BURReze, $listRennes44BURStHerblainAsturia, $listRennes44BURStHerblainExapole, $listRennes44BURStHerblainSunset, $listRennes44BURTrignac, $listRennes53Change,
    $listRennes56BURRieux, $listRennes72LeMans, $listRennes74Annecy, $listRennes79Niort, $listRennes85Fontenay, $listRennes85LaRoche, $listRennes85LesSablesdol, $listRennes85SaintHilaire, $listRennes85StVincent,
    $listRennes85StGilles)
    

    Write-Host $tablListCsv.Count " fichiers récupérés avec succés !" -ForegroundColor Green;

}
catch {
    Write-Host "Un ou des fichiers CSV de Rennes est(sont) introuvable(s)" -ForegroundColor Red;
    Write-Host $_.Exception.Message -ForegroundColor Yellow;
    Break;
}


#Récupère les listes de NANTES du site Sharepoint
Try {
    $lRennes14Demouville = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES14DEMOUVILLETHYSSENDEMOUVILLE")
    $lRennes17AytreLePatio = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17AYTRELEPATIODESTILLEULS")
    $lRennes17AytreVillas = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17AYTREVILLASTILIA")
    $lRennes17Clavette = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17CLAVETTELETENNIS")
    $lRennes17Esnandes = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17ESNANDESESNANDES")
    $lRennes17LaRochelleAzurea = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17LAROCHELLEAZUREA")
    $lRennes17LaRochelleHorizon = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17LAROCHELLEHORIZONMER")
    $lRennes17LaRochelleLaRochelle = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17LAROCHELLELAROCHELLEEINSTEIN")
    $lRennes17LaRochelleGlobe = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17LAROCHELLELEGLOBETROTTER")
    $lRennes17LaRochelleVilla= (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17LAROCHELLEVILLAGARANCE")
    $lRennes17Nieul = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17NIEULSURMERVILLAROSA")
    $lRennes17Perigny = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES17PERIGNYLESJARDINSDESACANTHES")
    $lRennes22Lanvallay = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES22LANVALLAYLECLOSDESORMEAUX")
    $lRennes22Lehon = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES22LEHONLECLOSTRIARD")
    $lRennes22PleneufCote = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES22PLENEUFVALANDCOTEMER")
    $lRennes22PleneufValand = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES22PLENEUFVALANDLESAIGUESMARINES")
    $lRennes22StCast = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES22SAINTCASTLESPIERRESSONNANTES")
    $lRennes22Tregueux = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES22TREGUEUXBREZILLET")
    $lRennes29Brest = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29BRESTSAINTPIERRE")
    $lRennes29CarantecVillas = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29CARANTECLESVILLASDEKERGRIST")
    $lRennes29CarantecParc = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29CARANTECPARCOCEAN")
    $lRennes29Guilers = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29GUILERSLECLOSVALENTIN")
    $lRennes29Guipavas = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29GUIPAVASLESRESIDENCESSAINTEXUPERY")
    $lRennes29Landivisiau = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29LANDIVISIAUPARCLANDIVISIAU")
    $lRennes29LeDrennec = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29LEDRENNECLESJARDINSDADRIEN")
    $lRennes29Locmaria = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29LOCMARIAPLOUZALESJARDINSDELOCMARIA")
    $lRennes29Plouarzel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29PLOUARZELKERVEN")
    $lRennes29Plougonvelin = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29PLOUGONVELINLESJARDINSDUTREZHIR")
    $lRennes29PontLabbe = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29PONTLABBETERRASSESETANG")
    $lRennes29Quimper = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES29QUIMPERDOMAINECENTRE")
    $lRennes35BaindeBrForum = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BAINDEBRETAGNELEFORUM")
    $lRennes35BaindeBrResidence = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BAINDEBRETAGNERESIDENCEDESTANNEURS")
    $lRennes35BettonLesCapucines = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BETTONLESCAPUCINES")
    $lRennes35BettonCinq = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BETTONLESCINQILES")
    $lRennes35BettonPresquille = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BETTONPRESQUILLE")
    $lRennes35BettonVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BETTONVILLACANTALINA")
    $lRennes35BettonVillaOdalie = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BETTONVILLAODALIE")
    $lRennes35BoisGervilly = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BOISGERVILLYLANCELOTDULAC")
    $lRennes35BettonVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BRUZINEO")
    $lRennes35BruzIneo = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BETTONVILLACANTALINA")
    $lRennes35BruzJardins = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35BRUZLESJARDINSDEBLOSSAC")
    $lRennes35CessonMaison = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35CESSONSEVIGNECESSONMAISONMEDICALE")
    $lRennes35CessonNet = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35CESSONSEVIGNENETPLUS")
    $lRennes35ChantepieLeDomaine = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35CHANTEPIELEDOMAINEDUCANAL")
    $lRennes35ChantepieLesCoquelicots = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35CHANTEPIELESCOQUELICOTS")
    $lRennes35ChantepieVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35CHANTEPIEVILLAABELIA")
    $lRennes35Domloup = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35DOMLOUPDOMLOUP")
    $lRennes35Etrelle = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35ETRELLEETRELLEVINCIENERGIES")
    $lRennes35Gosne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35GOSNELESPORTESDOUEES")
    $lRennes35Janze = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35JANZELESAMBROISINES")
    $lRennes35LaChapelleRuePace = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LACHAPELLERUEDEPACE")
    $lRennes35LaChapelleRueLechlade = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LACHAPELLERUELECHLADE")
    $lRennes35LaChapelleVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LACHAPELLEVILLALAURENA")
    $lRennes35LeRheuJardins = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LERHEUJARDINSDADELE")
    $lRennes35LeRheuGrand = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LERHEULEGRANDJARDIN")
    $lRennes35LeRheuNeventi = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LERHEUNEVENTI")
    $lRennes35LeRheuRue = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LERHEURUEDERENNES")
    $lRennes35LeRheuThyssen = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LERHEUTHYSSENLERHEU")
    $lRennes35LeRheuVillas = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LERHEUVILLASMELIES")
    $lRennes35Lhermitage = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LHERMITAGELESAQUARELLES")
    $lRennes35LiffreParc = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LIFFREPARCDESETANGS1")
    $lRennes35LiffreRose = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35LIFFREROSEARMOR")
    $lRennes35MezieresSousCGrande = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35MEZIERESSOUSCLAGRANDEPREE")
    $lRennes35MezieresSousCLaPree = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35MEZIERESSOUSCLAPREEDUPETITBOIS")
    $lRennes35MontaubanDeBr = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35MONTAUBANDEBRSAINTELOI")
    $lRennes35Noyal = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35NOYALLESHORTENSIAS")
    $lRennes35Parame = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35PARAME")
    $lRennes35RennesAdiph = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESADIPH")
    $lRennes35RennesAvenue = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESAVENUEMAGINOT")
    $lRennes35RennesCapNord = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESCAPNORD")
    $lRennes35RennesCarre = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESCARREDART")
    $lRennes35RennesCassini = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESCASSINI")
    $lRennes35RennesCastel = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESCASTELRIVIERA")
    $lRennes35RennesCoeur = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESCOEURDEVILLE")
    $lRennes35RennesDemat = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESDEMAT")
    $lRennes35RennesEolysII = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESEOLYSII")
    $lRennes35RennesVisitation = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLAVISITATION")
    $lRennes35RennesMurano = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLEMURANO")
    $lRennes35RennesNoven = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLENOVEN")
    $lRennes35RennesLeSextant = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLESEXTANT")
    $lRennes35RennesLesJardinsChatillon = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLESJARDINSDECHATILLON")
    $lRennes35RennesLesJardinsNeroli = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLESJARDINSDENEROLI")
    $lRennes35RennesLesOpalines = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLESOPALINES")
    $lRennes35RennesLesPrairies = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLESPRAIRIESDELILLE")
    $lRennes35RennesLesRives = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESLESRIVESDETASSIGNY")
    $lRennes35RennesMadisonParc = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESMADISONPARC")
    $lRennes35RennesOsiris = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESOSIRIS")
    $lRennes35RennesResidence = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESRESIDENCEDEVINCI")
    $lRennes35RennesRue = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESRUELEGRAVERAND")
    $lRennes35RennesTerranova = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESTERRANOVA")
    $lRennes35RennesVillaCamilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESVILLACAMILLA")
    $lRennes35RennesVillaVinci = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35RENNESVILLADEVINCI")
    $lRennes35Romagne = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35ROMAGNELECLOSDESSAULES")
    $lRennes35StBriac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTBRIACLESROCHESDOUVRES")
    $lRennes35StGregoireAxis = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTGREGOIREAXIS")
    $lRennes35StGregoireBpo = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTGREGOIREBPO")
    $lRennes35StGregoireEdonia = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTGREGOIREEDONIA")
    $lRennes35StGregoireLaBoutiere = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTGREGOIRELABOUTIERE")
    $lRennes35StGregoireParcBroceliande = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTGREGOIREPARCDEBROCELIANDE")
    $lRennes35StGregoireParcEllena = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTGREGOIREPARCELLENA")
    $lRennes35StGregoirePole = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTGREGOIREPOLEMEDICALLABOUTIERE")
    $lRennes35StJacquesParty = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTJACQUESACTIVLANDDARTY")
    $lRennes35StJacquesMondial = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTJACQUESACTIVLANDMONDIALRELAY")
    $lRennes35StJacquesAdaggio = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTJACQUESADAGGIOSOCIAL")
    $lRennes35StJacquesVilla = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTJACQUESVILLAGIULIA")
    $lRennes35StMaloBayaderes = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTMALOLESBAYADERES")
    $lRennes35StMaloMarines = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTMALOLESMARINESDECHASLES")
    $lRennes35StMaloParc = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTMALOPARCDELHERMINE")
    $lRennes35StMaloSquare = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES35SAINTMALOSQUAREACADIE")
    $lRennes44BURCarquefou = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURCARQUEFOUATALIS")
    $lRennes44BURCoueron = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURCOUERONGLASSOLUTION")
    $lRennes44BURLaChapelle = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURLACHAPELLEINEO")
    $lRennes44BURNantesEspace = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURNANTESESPACENEWTON")
    $lRennes44BURNantesEuropa = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURNANTESEUROPA")
    $lRennes44BURNantesExalis = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURNANTESEXALIS")
    $lRennes44BUROrvault = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURORVAULTBOISCESBRON")
    $lRennes44BURPornic = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURPORNICPORNIC")
    $lRennes44BURReze = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURREZEPSAREZE")
    $lRennes44BURStHerblainAsturia = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURSAINTHERBLAINASTURIA")
    $lRennes44BURStHerblainExapole = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURSAINTHERBLAINEXAPOLE")
    $lRennes44BURStHerblainSunset = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURSAINTHERBLAINSUNSET")
    $lRennes44BURTrignac = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES44BURTRIGNACGRANDCHAMPS")
    $lRennes53Change = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES53CHANGESPIECHANGE")
    $lRennes56BURRieux = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES56BURRIEUXVEOLIAEAU")
    $lRennes72LeMans = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES72LEMANSLEGALILEE")
    $lRennes74Annecy = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES74ANNECYLEVIEUXPINSONS")
    $lRennes79Niort = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES79NIORTLECLOSDESTILLEULS")
    $lRennes85Fontenay = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES85FONTENAYLECOMLESCOLLIBERTS")
    $lRennes85LaRoche = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES85LAROCHESURYONRESIDENCEELLINE")
    $lRennes85LesSables = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES85LESSABLESDOL")
    $lRennes85StHilaire = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES85SAINTHILAIRESTHILAIREDURIEZ")
    $lRennes85StVincent = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES85SAINTVINCENTLESAINTVINCENT")
    $lRennes85StGilles = (Get-Spweb "http://inside.lamotte.fr/doc-opm").GetList("http://inside.lamotte.fr/doc-opm/Lists/RENNES85STGILLESCROIXPOEME")
   
    $tablListSite = @($lRennes14Demouville, $lRennes17AytreLePatio, $lRennes17AytreVillas, $lRennes17Clavette, $lRennes17Esnandes, $lRennes17LaRochelleAzurea, $lRennes17LaRochelleHorizon, $lRennes17LaRochelleLaRochelle,
    $lRennes17LaRochelleGlobe, $lRennes17LaRochelleVilla, $lRennes17Nieul, $lRennes17Perigny, $lRennes22Lanvallay, $lRennes22Lehon, $lRennes22PleneufCote, $lRennes22PleneufValand, $lRennes22StCast, $lRennes22Tregueux,
    $lRennes29Brest, $lRennes29CarantecVillas, $lRennes29CarantecParc, $lRennes29Guilers, $lRennes29Guipavas, $lRennes29Landivisiau, $lRennes29LeDrennec, $lRennes29Locmaria, $lRennes29Plouarzel, $lRennes29Plougonvelin,
    $lRennes29PontLabbe, $lRennes29Quimper, $lRennes35BaindeBrForum, $lRennes35BaindeBrResidence, $lRennes35BettonLesCapucines, $lRennes35BettonCinq, $lRennes35BettonPresquille, $lRennes35BettonVilla,
    $lRennes35BettonVillaOdalie, $lRennes35BoisGervilly, $lRennes35BruzIneo, $lRennes35BruzJardins, $lRennes35CessonMaison, $lRennes35CessonNet, $lRennes35ChantepieLeDomaine,
    $lRennes35ChantepieLesCoquelicots, $lRennes35ChantepieVilla, $lRennes35Domloup, $lRennes35Etrelle, $lRennes35Gosne, $lRennes35Janze, $lRennes35LaChapelleRuePace, $lRennes35LaChapelleRueLechlade,
    $lRennes35LaChapelleVilla, $lRennes35LeRheuJardins, $lRennes35LeRheuGrand, $lRennes35LeRheuNeventi, $lRennes35LeRheuRue, $lRennes35LeRheuThyssen, $lRennes35LeRheuVillas, $lRennes35Lhermitage, $lRennes35LiffreParc,
    $lRennes35LiffreRose, $lRennes35MezieresSousCGrande, $lRennes35MezieresSousCLaPree, $lRennes35MontaubanDeBr, $lRennes35Noyal, $lRennes35Parame, $lRennes35RennesAdiph, $lRennes35RennesAvenue, $lRennes35RennesCapNord,
    $lRennes35RennesCarre, $lRennes35RennesCassini, $lRennes35RennesCastel, $lRennes35RennesCoeur, $lRennes35RennesDemat, $lRennes35RennesEolysII, $lRennes35RennesVisitation, $lRennes35RennesMurano,
    $lRennes35RennesNoven, $lRennes35RennesLeSextant, $lRennes35RennesLesJardinsChatillon, $lRennes35RennesLesJardinsNeroli, $lRennes35RennesLesOpalines, $lRennes35RennesLesPrairies, $lRennes35RennesLesRives,
    $lRennes35RennesMadisonParc, $lRennes35RennesOsiris, $lRennes35RennesResidence, $lRennes35RennesRue, $lRennes35RennesTerranova, $lRennes35RennesVillaCamilla, $lRennes35RennesVillaVinci, $lRennes35Romagne,
    $lRennes35StBriac, $lRennes35StGregoireAxis, $lRennes35StGregoireBpo, $lRennes35StGregoireEdonia, $lRennes35StGregoireLaBoutiere, $lRennes35StGregoireParcBroceliande, $lRennes35StGregoireParcEllena,
    $lRennes35StGregoirePole, $lRennes35StJacquesParty, $lRennes35StJacquesMondial, $lRennes35StJacquesAdaggio, $lRennes35StJacquesVilla, $lRennes35StMaloBayaderes, $lRennes35StMaloMarines, $lRennes35StMaloParc,
    $lRennes35StMaloSquare, $lRennes44BURCarquefou, $lRennes44BURCoueron, $lRennes44BURLaChapelle, $lRennes44BURNantesEspace, $lRennes44BURNantesEuropa, $lRennes44BURNantesExalis, $lRennes44BUROrvault,
    $lRennes44BURPornic, $lRennes44BURReze, $lRennes44BURStHerblainAsturia, $lRennes44BURStHerblainExapole, $lRennes44BURStHerblainSunset, $lRennes44BURTrignac, $lRennes53Change, $lRennes56BURRieux, $lRennes72LeMans,
    $lRennes74Annecy, $lRennes79Niort, $lRennes85Fontenay, $lRennes85LaRoche, $lRennes85LesSables, $lRennes85StHilaire, $lRennes85StVincent, $lRennes85StGilles)
    

    Write-Host $tablListSite.length " listes récupérés avec succés !" -ForegroundColor Green;

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

            Write-Progress -Id 1 -Activity "Importation des données" -Status 'Progress' -PercentComplete (($counter / ($tablListCsv[$i].Count+1)) * 100);
        }
        Write-Host ((${r} -1), " ligne(s) ajoutée(s) à la liste", $tablListSite[$i]) -ForegroundColor Green;
    }
}
catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
}