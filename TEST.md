# Plan de test VBA (Actions du Ruban)

## Environnement de test

- Classeur avec macros activées.
- Feuilles requises présentes :
  - `Suivi_CR`
  - `Suivi_Livrables`
  - `PowQ_Extract`
  - `PowQ_Suivi_UVR`
  - `PowQ_EDU_CE_VHST`
  - `BN_Suivi dossier Safety`
  - `config` (masquée)
- Fichiers de test préparés :
  - Fichier Excel avec `Suivi_Livrables`
  - Fichier Excel sans `Suivi_Livrables`
  - Fichier EDU valide pour `PowQ_EDU_CE_VHST`
  - Fichier UVR valide pour `PowQ_Suivi_UVR`
  - Fichier Extract valide pour `PowQ_Extract`
  - Variantes invalides ou incomplètes de ces fichiers PowQ (tests d'erreur)
  - Fichier verrouillé/indisponible optionnel pour le cas d'erreur
- Dossier d'archive/sortie en écriture.

Légende des colonnes de résultat :

- `PASS` / `FAIL`
- Ajouter une note courte en cas d'échec.

---

## Matrice d'exécution


| ID  | Bouton / Macro                                             | Entrée / Préparation                                                        | Résultat attendu                                                                             | PASS/FAIL | Notes                      |
| --- | ---------------------------------------------------------- | --------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------- | --------- | -------------------------- |
| T01 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Données valides dans toutes les feuilles sources                            | `Suivi_Livrables` reconstruit, popup de succès avec compteurs, aucun verrou bloqué           | PASS      |                            |
| T02 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Une feuille requise manquante (renommer temporairement)                     | Popup d'erreur claire, état final non corrompu, paramètres applicatifs restaurés             | PASS      |                            |
| T03 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Exécuter deux fois de suite                                                 | Les deux exécutions se terminent proprement                                                  | PASS      |                            |
| T04 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Ajouter un nouveau STR dans `PowQ_EDU_CE_VHST` avec sprint max valide       | Un nouveau bloc STR est créé dans `Suivi_Livrables` après mise à jour                        | PASS      |                            |
| T05 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Supprimer un STR existant dans `PowQ_EDU_CE_VHST`                           | Le bloc STR supprimé n'apparaît plus dans `Suivi_Livrables` après reconstruction             | PASS      |                            |
| T06 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Modifier une valeur STR (renommer code/nom STR) dans `PowQ_EDU_CE_VHST`     | L'ancien bloc STR disparaît et le nouveau bloc renommé apparaît dans `Suivi_Livrables`       | PASS      |                            |
| T07 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Dans `config`, ajouter une nouvelle valeur dans `Fonctions`                 | Pour chaque STR/sprint/type, des lignes sont générées pour la fonction ajoutée               | PASS      |                            |
| T08 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Dans `config`, supprimer une valeur existante de `Fonctions`                | Les lignes utilisant la fonction supprimée ne sont plus générées dans `Suivi_Livrables`      | PASS      |                            |
| T09 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Dans `config`, renommer une valeur de `Fonctions`                           | Les lignes passent de l'ancien libellé de fonction au nouveau après reconstruction           | PASS      |                            |
| T10 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Dans `config`, ajouter une nouvelle valeur `Type de livrable`               | De nouveaux blocs de type apparaissent pour chaque combinaison STR/sprint/fonction           | PASS      |                            |
| T11 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Dans `config`, supprimer une valeur `Type de livrable`                      | Les blocs du type supprimé n'apparaissent plus dans `Suivi_Livrables`                        | PASS      |                            |
| T12 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Dans `config`, renommer une valeur `Type de livrable`                       | Les lignes passent de l'ancien libellé de type au nouveau après reconstruction               | PASS      |                            |
| T13 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Dans `config`, vider toutes les valeurs `Type de livrable`                  | Le prompt de secours apparaît (par défaut ADL1/SwDS) et la mise à jour suit le choix         | PASS      |                            |
| T14 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Prompt de secours `Type de livrable` : répondre **Oui**                     | `ADL1`/`SwDS` sont réinjectés dans `config`, puis la reconstruction continue                 | PASS      |                            |
| T15 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Prompt de secours `Type de livrable` : répondre **Non**                     | La mise à jour s'arrête avec erreur explicite, sans corruption des données                   | PASS      |                            |
| T16 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Annuler le choix du dossier partagé au démarrage                            | Message d'annulation clair, aucun traitement destructif lancé                                | PASS      |                            |
| T17 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | `PowQ_EDU_CE_VHST` sans STR exploitable                                     | Message "Aucune STR disponible", sortie propre, paramètres Excel restaurés                   | PASS      |                            |
| T18 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | `config` sans valeur dans `Fonctions`                                       | Erreur bloquante claire sur `Fonctions`, rollback propre (lock/libération état Excel)        | PASS      |                            |
| T19 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | STR présente dans CR mais absente de VHST, prompt synchro = **Oui**         | STR ajoutée dans VHST avec sprint max CR, puis reconstruction avec cette STR                 | PASS      |                            |
| T20 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Sprint max CR > sprint max VHST, prompt synchro = **Oui**                   | Max sprint VHST est mis à jour avec la valeur CR, puis reconstruction cohérente              | PASS      |                            |
| T21 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Prompt synchro VHST (ajout/mise à jour) : répondre **Cancel**               | Sortie de la boucle de synchro, mise à jour continue sans plantage                           | PASS      |                            |
| T22 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Préremplir `L,N,P,Q,R,S,Y`, puis lancer reconstruction                      | Les valeurs manuelles de `L,N,P,Q,R,S,Y` sont restaurées après rebuild                       | PASS      |                            |
| T23 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Prompt synchro VHST (écart sprint), répondre Non                            | Aucune mise à jour VHST, reconstruction continue sans crash                                  | PASS      |                            |
| T24 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Feuille `PowQ_Suivi_UVR` absente                                            | Erreur claire, lock libéré, état Excel restauré                                              |           |                            |
| T25 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Colonne `fin ref` absente dans `PowQ_Extract`                               | Colonne I de `Suivi_Livrables` reste vide, traitement terminé sans erreur                    | PASS      |                            |
| T26 | `Mise à jour` / `Ribbon_UpdateSuivi`                       | Verrou présent avec le même utilisateur                                     | Message "mise à jour déjà en cours", arrêt propre sans modification                          |           |                            |
| T27 | `Archiver` / `Ribbon_ArchiveSuivi`                         | Cliquer et confirmer Oui                                                    | Fichier d'archive créé dans le dossier daté, lignes de données livrables vidées              | PASS      |                            |
| T28 | `Archiver` / `Ribbon_ArchiveSuivi`                         | Cliquer et confirmer Non                                                    | Aucun fichier d'archive, aucun reset de feuille                                              | PASS      |                            |
| T29 | `Archiver` / `Ribbon_ArchiveSuivi`                         | Après archivage, choisir ouvrir fichier = Oui                               | Le fichier d'archive s'ouvre                                                                 | PASS      |                            |
| T30 | `Archiver` / `Ribbon_ArchiveSuivi`                         | Feuille `Suivi_Livrables` absente                                           | Message feuille introuvable, sortie propre                                                   | PASS      |                            |
| T31 | `Archiver` / `Ribbon_ArchiveSuivi`                         | `SaveAs` impossible (dossier inaccessible/lecture seule)                    | Erreur explicite, état Excel restauré, aucun fichier partiel                                 | PASS      |                            |
| T32 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | Fournir un fichier EDU valide                                               | `PowQ_EDU_CE_VHST` mis à jour, popup de succès                                               | PASS      |                            |
| T33 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Fournir un fichier UVR valide                                               | `PowQ_Suivi_UVR` mis à jour, formats de date attendus                                        | PASS      |                            |
| T34 | `Extract` / `Ribbon_PowQExtract`                           | Fournir un fichier Extract valide                                           | `PowQ_Extract` mis à jour                                                                    | PASS      |                            |
| T35 | `PowQ Tout` / `Ribbon_PowQAll`                             | Fournir les 3 fichiers et confirmer                                         | Les trois mises à jour sont exécutées, résumé d'état final affiché                           | PASS      |                            |
| T36 | `PowQ Tout` / `Ribbon_PowQAll`                             | Annuler un sélecteur                                                        | Le processus se termine proprement sans casse partielle                                      | PASS      |                            |
| T37 | `PowQ Tout` / `Ribbon_PowQAll`                             | Après sélection des 3 fichiers, confirmer = **Non**                         | Aucun traitement PowQ n'est lancé, sortie propre                                             | PASS      |                            |
| T38 | `Extract` / `Ribbon_PowQExtract`                           | Feuille `PowQ_Extract` contient déjà un/des tableau(x), confirmer **Oui**   | Les anciens tableaux sont supprimés puis `Extract_MSP` est recréé correctement               | PASS      |                            |
| T39 | `Extract` / `Ribbon_PowQExtract`                           | Feuille `PowQ_Extract` contient déjà un/des tableau(x), confirmer **Non**   | Le traitement s'arrête avec message clair, aucune reconstruction forcée                      | PASS      |                            |
| T40 | `Extract` / `Ribbon_PowQExtract`                           | Source vide / sans lignes exploitables                                      | Message d'avertissement, aucune donnée incohérente écrite dans `PowQ_Extract`                | PASS      |                            |
| T41 | `Extract` / `Ribbon_PowQExtract`                           | Toutes les lignes filtrées invalides (`A` vide ou sprint non numérique)     | Message "Aucune ligne valide trouvée.", sortie propre                                        | PASS      |                            |
| T42 | `Extract` / `Ribbon_PowQExtract`                           | Fichier source déjà ouvert                                                  | Mise à jour réussie sans fermeture non voulue du classeur utilisateur                        | PASS      | le fichier source se ferme |
| T43 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | Une colonne obligatoire absente dans la source                              | Erreur claire "Colonne obligatoire introuvable", pas d'écriture partielle                    | PASS      |                            |
| T44 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | En-têtes requis présents mais pas sur la même ligne                         | Erreur claire sur l'alignement des en-têtes, traitement interrompu proprement                | PASS      |                            |
| T45 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | Colonnes source de longueurs différentes                                    | Table redimensionnée à la longueur max, colonnes plus courtes complétées par des blancs      | PASS      |                            |
| T46 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | Feuille source `Références_CE_VHST` absente dans le fichier d'entrée        | Erreur claire "feuille introuvable", arrêt propre sans écriture partielle                    | PASS      |                            |
| T47 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | Feuille source présente mais vide                                           | Message d'avertissement, aucune mise à jour incohérente sur `PowQ_EDU_CE_VHST`               | PASS      |                            |
| T48 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | Après update réussi, vérifier les validations de `Suivi_CR` (B/I/K/M)       | Listes de validation alimentées depuis `PowQ_EDU_CE_VHST` et dropdowns fonctionnels          | PASS      |                            |
| T49 | `EDU_CE_VHST` / `Ribbon_PowQEDUCEVHST`                     | Aucun tableau dans `PowQ_EDU_CE_VHST`                                       | Tableau `EDU_CE_VHST` créé avec en-têtes requis puis alimenté                                | PASS      |                            |
| T50 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | `Suivi_UVR` sans `Suivi_UVR` mais avec autres tableaux, confirmer **Oui**   | Tableaux existants supprimés puis table `Suivi_UVR` recréée                                  | PASS      |                            |
| T51 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | `Suivi_UVR` sans `Suivi_UVR` mais avec autres tableaux, confirmer **Non**   | Processus arrêté avec message explicite, aucune suppression non voulue                       | PASS      |                            |
| T52 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | En-tête attendu absent dans la source `Global`                              | Erreur de correspondance colonnes (colonne manquante), arrêt propre                          | PASS      |                            |
| T53 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Fichier UVR déjà ouvert avant lancement                                     | Le traitement réussit sans fermer le classeur déjà ouvert par l'utilisateur                  |           |                            |
| T54 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Feuille `Global` absente dans le fichier source                             | Erreur explicite "feuille introuvable", arrêt propre                                         | PASS      |                            |
| T55 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Feuille `Global` présente mais vide                                         | Message d'avertissement, aucune écriture incohérente sur `PowQ_Suivi_UVR`                    |           |                            |
| T56 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Aucun en-tête exploitable en ligne 12 (A:W)                                 | Erreur claire sur en-têtes non exploitables, arrêt propre                                    |           |                            |
| T57 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Aucun tableau dans `PowQ_Suivi_UVR` (création depuis en-têtes source)       | Tableau `Suivi_UVR` créé avec les en-têtes source et données chargées                        |           |                            |
| T58 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Vérifier normalisation dates sur colonnes H/I/K/L/O/Q/S                     | Dates bien converties/formattées en `dd/mm/yyyy`, valeurs invalides nettoyées                |           |                            |
| T59 | `Suivi_UVR` / `Ribbon_PowQUVR`                             | Source avec `#N/A`/erreurs Excel dans colonnes importées                    | Valeurs invalides converties en vide, aucun crash                                            |           |                            |
| T60 | `PowQ Tout` / `Ribbon_PowQAll`                             | Un update échoue et les autres réussissent                                  | Résumé final affiche les statuts mixés `OK`/`ERROR` par processus                            |           |                            |
| T61 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | Config + CR + VHST valides                                                  | Combinaisons manquantes ajoutées, col E mise à jour, tri + bordures appliqués                | PASS      |                            |
| T62 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | BN contient des lignes obsolètes                                            | Le prompt des lignes obsolètes apparaît ; Oui supprime les lignes obsolètes                  |           |                            |
| T63 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | Prompt lignes obsolètes, répondre Non                                       | Lignes obsolètes conservées, aucune suppression                                              |           |                            |
| T64 | `Archiver BN_Suivi` / `Ribbon_ArchiveBNSuivi`              | Cliquer et confirmer Oui                                                    | Archive BN créée et zone de données BN réinitialisée                                         | PASS      |                            |
| T65 | `Archiver BN_Suivi` / `Ribbon_ArchiveBNSuivi`              | Confirmer Non                                                               | Aucun changement                                                                             | PASS      |                            |
| T66 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | `PowQ_EDU_CE_VHST` sans en-tête `Nom_STR`                                   | Erreur claire sur en-tête manquant (`Nom_STR`), arrêt propre                                 |           |                            |
| T67 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | `PowQ_EDU_CE_VHST` sans en-tête `Sprint`                                    | Erreur claire sur en-tête manquant (`Sprint`), arrêt propre                                  |           |                            |
| T68 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | `config` sans en-tête `Fonctions`                                           | Erreur claire sur en-tête manquant (`Fonctions`), aucun changement partiel                   |           |                            |
| T69 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | Relancer avec mêmes données déjà synchronisées                              | Message "Aucun changement", aucune ligne BN ajoutée                                          | PASS      |                            |
| T70 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | Après ajout massif de lignes                                                | Tri final correct sur B puis C puis D                                                        |           |                            |
| T71 | `Compléter BN_Suivi` / `Ribbon_AddBNSuivi`                 | Après ajout/mise à jour                                                     | Bordures B:G présentes sur toutes les lignes impactées                                       | PASS      |                            |
| T72 | `Archiver BN_Suivi` / `Ribbon_ArchiveBNSuivi`              | Feuille `BN_Suivi dossier Safety` absente                                   | Message feuille introuvable, sortie propre                                                   |           |                            |
| T73 | `Archiver BN_Suivi` / `Ribbon_ArchiveBNSuivi`              | Annuler la sélection du dossier partagé                                     | Message d'annulation, aucun archivage, aucune suppression de lignes                          | PASS      |                            |
| T74 | `Archiver BN_Suivi` / `Ribbon_ArchiveBNSuivi`              | Dossier archive inaccessible / lecture seule                                | Échec explicite, état Excel restauré, log d'erreur écrit si chemin dispo                     |           |                            |
| T75 | `Archiver BN_Suivi` / `Ribbon_ArchiveBNSuivi`              | Archivage réussi                                                            | Vérifier reset exact : seules les lignes 1-2 restent, données à partir de la ligne 3 vidées  | PASS      |                            |
| T76 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Sélectionner 2 fichiers avec `Suivi_Livrables`                              | Fichier généré créé (`Collect_HHMMSS_DDMMYYYY.xlsx`), statut avec entrées OK                 | PASS      |                            |
| T77 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Sélectionner un fichier sans `Suivi_Livrables`                              | Le fichier manquant est marqué `IGNORE`, le processus continue                               | PASS      |                            |
| T78 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Sélectionner le classeur courant (classeur macros)                          | Aucun classeur fermé, aucun crash de copie, statut conforme au résultat réel                 | PASS      |                            |
| T79 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Annuler le sélecteur du dossier de sauvegarde                               | Aucun fichier sauvegardé, classeur de sortie masqué fermé, statut affiché                    | PASS      |                            |
| T80 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Après sauvegarde, choisir ouvrir = Oui                                      | La fenêtre du fichier généré s'ouvre/s'active                                                | pass      |                            |
| T81 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Après sauvegarde, choisir ouvrir = Non                                      | Le classeur généré se ferme proprement                                                       | PASS      |                            |
| T82 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Inclure un fichier verrouillé/indisponible                                  | Fichier marqué `ECHEC`, les autres fichiers sont quand même traités                          |           |                            |
| T83 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Annuler le sélecteur de fichiers initial                                    | Sortie immédiate sans création de classeur de sortie                                         | PASS      |                            |
| T84 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Tous les fichiers sans `Suivi_Livrables` (ou tous en échec copie/ouverture) | Message "Aucune feuille ... trouvée", statut détaillé, aucun fichier enregistré              | PASS      |                            |
| T85 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Deux fichiers avec même nom de base                                         | Noms d'onglets générés uniques (`nom`, `nom_1`, etc.), aucune collision                      | PASS      |                            |
| T86 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Nom fichier avec caractères interdits (`\ / : * ? [ ]`)                     | Nom d'onglet nettoyé automatiquement et collecte réussie                                     | PASS      |                            |
| T87 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Fichier source déjà ouvert (hors classeur macros)                           | Le fichier déjà ouvert n'est pas fermé par la collecte                                       |           |                            |
| T88 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Échec de sauvegarde (`SaveAs`) : dossier en lecture seule/inaccessible      | Erreur explicite, fermeture propre du classeur de sortie, aucun fichier partiel              |           |                            |
| T89 | `Collecter Suivi_Livrable` / `Ribbon_CollectSuiviLivrable` | Source avec colonnes masquées                                               | Les colonnes masquées restent masquées dans la feuille copiée                                |           |                            |
| T90 | Générique - Annulation sûre                                | Annuler à la première confirmation/sélecteur                                | Aucun changement destructif, aucun crash                                                     | PASS      |                            |
| T91 | Générique - Restauration état Excel en erreur              | Déclencher une erreur contrôlée (ex. feuille/chemin manquant)               | `ScreenUpdating`, `EnableEvents`, `Calculation` restaurés ; erreur claire pour l'utilisateur | PASS      |                            |
| T92 | `Ruban` / `RunMacroSafe`                                   | Renommer temporairement une macro callback cible                            | Message "Macro introuvable" avec liste des macros testées                                    | PASS      |                            |


---

## Cas limites étendus


| ID  | Zone                    | Scénario                                                                                                  | Résultat attendu                                                                     | PASS/FAIL | Notes |
| --- | ----------------------- | --------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------ | --------- | ----- |
| E01 | Performance             | Exécuter `Mise à jour` avec un gros volume de données (grandes feuilles `Suivi_CR` / PowQ)                | Se termine sans crash ; la sortie reste cohérente                                    |           |       |
| E02 | Performance             | Exécuter `Collecter Suivi_Livrable` sur de nombreux fichiers (ex. 20+)                                    | Le processus se termine ; la liste d'état finale reste correcte                      |           |       |
| E03 | Concurrence/Verrou      | Déclencher `Mise à jour` alors que la cellule de verrou est déjà remplie par un autre utilisateur/session | Message de verrou explicite ; aucune corruption des données                          |           |       |
| E04 | Concurrence/Verrou      | Simuler un verrou périmé (ancien horodatage), choisir **Oui** pour effacer                                | Le verrou périmé est supprimé et la mise à jour continue                             |           |       |
| E05 | Concurrence/Verrou      | Simuler un verrou périmé, choisir **Non**                                                                 | La mise à jour s'arrête proprement sans changement                                   |           |       |
| E06 | Système de fichiers     | Dossier cible Archive/Collect en lecture seule                                                            | Erreur de sauvegarde claire ; état applicatif restauré                               |           |       |
| E07 | Système de fichiers     | Dossier réseau/partagé indisponible pendant l'exécution                                                   | Erreur claire et sortie sûre ; aucun classeur masqué bloqué                          |           |       |
| E08 | Système de fichiers     | Chemin très long / caractères spéciaux dans les noms de fichiers source                                   | Collect enregistre quand même ; noms de feuilles nettoyés en sécurité                |           |       |
| E09 | Robustesse des entrées  | Le fichier sélectionné pour collect est protégé par mot de passe / impossible à ouvrir                    | Fichier marqué `ECHEC` ; les autres fichiers continuent                              |           |       |
| E10 | Robustesse des entrées  | Le fichier sélectionné pour collect a un `Suivi_Livrables` vide                                           | Feuille copiée (ou signalée clairement si inaccessible), aucun crash                 |           |       |
| E11 | Régression de formatage | Après `Mise à jour`, vérifier couleurs/bordures/formats numériques dans `Suivi_Livrables`                 | Les règles de formatage sont toujours appliquées correctement                        |           |       |
| E12 | Régression de formatage | Après cycles d'ajout/archivage BN, vérifier cohérence bordures/ordre/contenu BN                           | La mise en page BN reste cohérente après exécutions répétées                         |           |       |
| E13 | Locale/Date             | Exécuter sur un système avec des paramètres régionaux de date différents                                  | Parsing/formatage des dates reste valide dans les sorties                            |           |       |
| E14 | Stabilité               | Exécuter tous les boutons dans une même session de façon répétée (boucle smoke)                           | Aucun drapeau bloqué, aucun classeur masqué non sauvegardé, aucune erreur cumulative |           |       |
| E15 | Compatibilité           | Valider sur une autre version/build Excel (version/bitness différente)                                    | Les callbacks Ruban et macros s'exécutent normalement                                |           |       |


---

## Checklist optionnelle après exécution


| Vérification                         | Attendu                                                     | PASS/FAIL | Notes |
| ------------------------------------ | ----------------------------------------------------------- | --------- | ----- |
| Boutons du Ruban toujours cliquables | Aucun état désactivé/bloqué                                 | PASS      |       |
| Aucun classeur masqué non sauvegardé | Aucun prompt de sauvegarde inattendu à la fermeture d'Excel | PASS      |       |
| Journaux d'erreurs (si applicable)   | Une ligne de log écrite pour les erreurs forcées            | PASS      |       |
| Intégrité des données                | Les feuilles principales restent lisibles et cohérentes     | PASS      |       |


