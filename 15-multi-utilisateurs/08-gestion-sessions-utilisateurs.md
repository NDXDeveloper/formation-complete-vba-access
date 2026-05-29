🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.8. Gestion des sessions utilisateurs (LDB et table de sessions)

Savoir **qui est connecté** à une application partagée répond à des besoins très concrets : afficher la liste des utilisateurs actifs, demander à chacun de se déconnecter avant une opération de maintenance, faire respecter une limite de licences, ou simplement diagnostiquer un verrou. Deux approches complémentaires existent : exploiter le **fichier de verrouillage** créé par le moteur, ou maintenir sa propre **table de sessions**. La première est automatique mais limitée ; la seconde demande du travail mais offre une information riche. Cette section détaille les deux.

## Le fichier de verrouillage (.laccdb / .ldb)

Comme vu à la [section 15.2](02-modes-partage.md), l'ouverture d'une base partagée crée un fichier de verrouillage portant le même nom — `.laccdb` pour une base `.accdb`, `.ldb` pour une base `.mdb`. Son rôle premier est de déterminer quels enregistrements sont verrouillés dans la base partagée, et par qui.

Sa structure est simple et documentée. Pour chaque session qui ouvre la base, le moteur écrit une entrée de **64 octets** : les **32 premiers octets contiennent le nom de l'ordinateur**, les **32 suivants la « security name » de l'utilisateur** (par exemple « Admin »). Comme le nombre maximal d'utilisateurs simultanés supporté par le moteur est de **255**, la taille du fichier de verrouillage ne dépasse jamais environ **16 kilo-octets**.

Le fichier n'est pas toujours supprimé à la fermeture : le moteur l'efface lorsque le dernier utilisateur quitte la base, sauf dans deux cas — lorsque la base est marquée comme endommagée, et lorsque le dernier sortant ne dispose pas des **droits de suppression** sur le dossier (rappel : le partage exige des droits complets sur le dossier, cf. 15.2).

## Lire le fichier directement — et pourquoi c'est insuffisant

On pourrait être tenté de lire ce fichier en le découpant en blocs de 64 octets pour en extraire les noms de poste et d'utilisateur. C'est techniquement possible, mais **trompeur**, pour deux raisons.

D'abord, lorsqu'un utilisateur ferme la base, **son entrée n'est pas retirée** du fichier : elle peut seulement être écrasée plus tard, quand un autre utilisateur ouvre la base. Le fichier contient donc des entrées « fantômes » de sessions terminées. **On ne peut pas, à partir du fichier seul, savoir qui est réellement connecté à un instant donné.**

Ensuite, la « security name » stockée est le nom de connexion au sens de la sécurité du moteur. En l'absence de sécurité par groupe de travail — situation par défaut des bases `.accdb` —, cette valeur vaut « Admin » pour tout le monde. Le fichier identifie donc le **poste**, mais pas la **personne**.

## La bonne approche : le User Roster

Plutôt que de lire et d'analyser le fichier, la méthode recommandée consiste à utiliser la fonctionnalité **« User Roster »** du fournisseur OLE DB d'Access. Sa supériorité tient à un point essentiel : là où la lecture brute du fichier ne distingue pas une entrée active d'une entrée fantôme, le User Roster **vérifie les verrous réels** et indique donc qui est *réellement* présent.

On l'obtient en ADO via la méthode `OpenSchema` avec un schéma spécifique au fournisseur. Le jeu d'enregistrements retourné contient les colonnes `COMPUTER_NAME`, `LOGIN_NAME`, `CONNECTED` et `SUSPECTED_STATE`.

```vba
Public Sub AfficherUtilisateursConnectes()
    Const SCHEMA_USERROSTER As String = "{947bb102-5d43-11d1-bdbf-00c04fb92675}"
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset

    ' Ouvrir une connexion vers la base dont on veut le roster
    ' (typiquement le back-end). CurrentProject.Connection vise le fichier courant.
    Set cn = CurrentProject.Connection
    Set rs = cn.OpenSchema(adSchemaProviderSpecific, , SCHEMA_USERROSTER)

    Do Until rs.EOF
        Debug.Print rs!COMPUTER_NAME, _
                    "Connecté=" & rs!CONNECTED, _
                    "Suspect=" & rs!SUSPECTED_STATE
        rs.MoveNext
    Loop
    rs.Close
End Sub
```

Pour interroger le **back-end** dans une architecture front-end / back-end, on ouvre une connexion ADO vers le fichier du back-end plutôt que d'utiliser `CurrentProject.Connection`, qui pointe vers le front-end. Le User Roster reste néanmoins soumis à la même limite d'identification : `LOGIN_NAME` vaut souvent « Admin », et c'est `COMPUTER_NAME` qui distingue réellement les sessions.

Il existe par ailleurs des **outils tout faits** issus de la communauté Access (utilitaires « Who's Connected ») qui exploitent ce mécanisme pour lister les postes connectés — pratiques pour un diagnostic ponctuel.

## Les limites communes du fichier et du roster

Que l'on passe par le fichier ou par le roster, l'information reste **pauvre** du point de vue applicatif : on obtient un nom de poste et un nom de connexion technique, sans savoir **qui** s'est authentifié dans *votre* application, avec **quel rôle**, **depuis quand**, ni **ce qu'il fait**. Pour cela, il faut une information de niveau applicatif : une table de sessions.

## La table de sessions

L'approche applicative consiste à maintenir, dans le back-end, une table recensant les sessions de l'application elle-même. Contrairement aux TempVars ([section 15.7](07-tempvars.md)), qui sont propres à chaque session et **non partagées**, cette table est commune à tous les utilisateurs : c'est précisément ce qui permet à chacun de savoir qui d'autre est connecté.

Une table de sessions contient typiquement un identifiant de session, l'**utilisateur applicatif** (celui qui s'est authentifié, cf. [section 20.6](/20-securite-protection/06-gestion-droits-utilisateurs.md)), le **poste** (`Environ("COMPUTERNAME")`), l'**utilisateur Windows** (`Environ("USERNAME")`, ou l'API `GetUserName`, cf. [section 22.2](/22-api-windows-integration-office/02-api-courantes.md)), l'**heure d'ouverture**, et une **heure de dernière activité** servant de « battement de cœur ».

Son cycle de vie comprend trois temps : créer la ligne à la connexion, la rafraîchir périodiquement pendant la session, et la supprimer à la fermeture.

```vba
' Table : tblSessions(IdSession NuméroAuto, Utilisateur, Poste,
'                     UtilisateurWindows, OuvertureLe, DerniereActiviteLe)

' --- À la connexion : créer la ligne et mémoriser sa clé dans une TempVar ---
Public Sub OuvrirSession()
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("tblSessions", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
        rs!Utilisateur = Nz(TempVars!UtilisateurID, "")
        rs!Poste = Environ("COMPUTERNAME")
        rs!UtilisateurWindows = Environ("USERNAME")
        rs!OuvertureLe = Now()
        rs!DerniereActiviteLe = Now()
        TempVars!IdSession = rs!IdSession      ' la clé NuméroAuto est lisible après AddNew
    rs.Update
    rs.Close
    Set rs = Nothing
End Sub

' --- Battement de cœur : à appeler régulièrement (ex. Form_Timer, toutes les 60 s) ---
Public Sub BattreCoeur()
    If Not IsNull(TempVars!IdSession) Then
        CurrentDb.Execute "UPDATE tblSessions SET DerniereActiviteLe = Now() " & _
            "WHERE IdSession = " & TempVars!IdSession & ";", dbFailOnError
    End If
End Sub

' --- À la fermeture propre ---
Public Sub FermerSession()
    If Not IsNull(TempVars!IdSession) Then
        CurrentDb.Execute "DELETE FROM tblSessions WHERE IdSession = " & _
            TempVars!IdSession & ";", dbFailOnError
    End If
End Sub
```

Le point délicat est la **fermeture anormale** : un plantage ou une coupure laisse une ligne orpheline dans la table. C'est le rôle du battement de cœur (`DerniereActiviteLe`) : on considère comme défuntes les sessions dont la dernière activité remonte au-delà d'un seuil (par exemple plus du double de l'intervalle de battement), et on les purge périodiquement. Le couple « battement régulier + nettoyage des sessions périmées » est ce qui rend la table fiable malgré les sorties brutales.

## Usages de la table de sessions

Une table de sessions ouvre plusieurs possibilités :

- afficher aux administrateurs **qui utilise l'application** en temps quasi réel ;
- **demander la déconnexion** avant une maintenance : on inscrit un drapeau dans une table de contrôle que les front-ends interrogent à intervalle régulier, ce qui permet d'inviter les utilisateurs à fermer — indispensable avant une opération exigeant un accès **exclusif** au back-end (compactage, modification de structure ; cf. 15.2) ;
- **limiter les sessions** : empêcher une double connexion d'un même utilisateur, ou respecter un nombre maximal d'utilisateurs ;
- alimenter l'**audit** et la traçabilité ([section 20.7](/20-securite-protection/07-audit-acces-modifications.md)) ;
- identifier, au niveau applicatif, **qui détient** une ressource — information plus parlante que le simple nom de poste du fichier de verrouillage (cf. [section 15.4](04-detection-conflits-verrouillage.md)).

## Quelle approche choisir ?

Le **User Roster** ne demande aucune infrastructure et convient pour un besoin ponctuel : compter les sessions, lister les postes connectés. Mais il reste limité (poste et nom technique, identification de la personne incertaine) et donne une vue de bas niveau.

La **table de sessions** demande de la construire et de l'entretenir (création à la connexion, battement de cœur, suppression à la sortie, nettoyage des sessions périmées), mais elle fournit une information riche, fiable et extensible, au niveau de l'application. C'est l'approche à retenir dès qu'on a besoin de savoir *qui*, dans le vocabulaire de l'application, fait *quoi* — la grande majorité des cas réels.

À titre indicatif sur le dimensionnement : si une solution fichier-serveur peut techniquement atteindre 255 utilisateurs simultanés, il est recommandé, pour une application où les utilisateurs ajoutent et mettent à jour fréquemment des données, de ne pas dépasser **25 à 50 utilisateurs**. Ces limites sont approfondies à la [section 15.10](10-limites-moteur-ace.md).

## Points de vigilance

- **Le fichier de verrouillage ne dit pas qui est connecté maintenant** : les entrées des sessions fermées subsistent jusqu'à être écrasées.
- **Le nom stocké est « Admin »** par défaut en `.accdb` : le fichier identifie le poste, pas la personne.
- **Préférer le User Roster à la lecture brute** : lui seul vérifie les verrous réels et distingue les sessions actives.
- **Pour interroger le back-end**, ouvrir une connexion vers son fichier, pas `CurrentProject.Connection`.
- **La table de sessions doit gérer les sorties brutales** : battement de cœur + purge des sessions périmées.
- **Ne pas confondre table de sessions (partagée) et TempVars (par session)** : l'une fait communiquer les utilisateurs, l'autre non.

## En résumé

Pour savoir qui utilise une application partagée, deux moyens se complètent. Le **fichier de verrouillage** (`.laccdb`/`.ldb`) stocke, par entrées de 64 octets (32 pour le poste, 32 pour le nom de sécurité), les sessions ayant ouvert la base, dans la limite de 255 utilisateurs (soit ≤ 16 Ko) ; mais il conserve les entrées des sessions fermées et n'identifie que le poste, ce qui le rend impropre à déterminer seul qui est connecté. Le **User Roster** du fournisseur OLE DB, interrogé via `OpenSchema`, est préférable car il vérifie les verrous réels et renvoie `COMPUTER_NAME`, `LOGIN_NAME`, `CONNECTED` et `SUSPECTED_STATE`. Pour une information de niveau applicatif — utilisateur authentifié, rôle, heure d'ouverture, activité —, on maintient une **table de sessions** dans le back-end, alimentée à la connexion, rafraîchie par un battement de cœur et purgée des sessions périmées. C'est cette table qui permet d'afficher les utilisateurs actifs, de demander des déconnexions avant maintenance et d'asseoir l'audit.

La dernière section technique du chapitre traite de la liaison dynamique au back-end : la [connexion à une base distante via table liée par code](09-tables-liees-par-code.md).

⏭️ [15.9. Connexion à une base distante via table liée par code](/15-multi-utilisateurs/09-tables-liees-par-code.md)
