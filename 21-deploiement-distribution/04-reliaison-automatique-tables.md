🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.4. Reliaison automatique des tables back-end (re-link)

Dans une architecture front-end / back-end, le front-end ne contient pas les tables de données : il contient des **tables liées**, c'est-à-dire des pointeurs vers les tables réelles stockées dans le fichier back-end. Chaque table liée mémorise le chemin d'accès au back-end. Tant que ce chemin reste valide, tout fonctionne ; mais dès que le back-end change d'emplacement — autre serveur, autre dossier, déploiement sur un nouveau site — les liens deviennent caducs et le front-end ne trouve plus ses données.

La **reliaison** (re-link) est l'opération qui consiste à mettre à jour le chemin mémorisé par les tables liées pour qu'il pointe vers le bon emplacement du back-end. Automatiser cette reliaison est indispensable à toute application Access déployée, car le chemin du back-end en production diffère presque toujours de celui utilisé pendant le développement.

## Le problème : des liens qui pointent vers le mauvais emplacement

Plusieurs situations rendent les liens invalides, et toutes se produisent au moment du déploiement plutôt que pendant le développement.

Le cas le plus courant tient à la **différence entre l'environnement de développement et la production**. Le développeur construit le front-end avec des liens pointant vers le back-end situé sur son propre poste (par exemple `C:\Dev\MonApp_be.accdb`). Une fois déployé, le back-end réside ailleurs (par exemple `\\Serveur\Partage\MonApp_be.accdb`). Les liens hérités du développement sont alors tous erronés.

Ce problème est amplifié par l'**auto-update** (section 21.3) : lorsqu'une nouvelle version du front-end est copiée sur le poste de l'utilisateur, cette copie fraîche transporte les liens *du développeur*, qui ne correspondent pas à l'environnement de l'utilisateur. Une reliaison doit donc systématiquement suivre la copie d'un front-end neuf.

Enfin, l'usage de **lecteurs réseau mappés** fragilise les liens : un lien pointant vers `S:\data\be.accdb` se casse si un poste ne dispose pas du lecteur `S:` ou le mappe vers un autre partage. Un chemin **UNC** (`\\Serveur\Partage\be.accdb`) est nettement plus stable d'un poste à l'autre, comme le détaille la section 21.5.

## Comprendre les tables liées et leur chaîne de connexion

Du point de vue de DAO, une table liée est un objet `TableDef` dont la propriété `Connect` est renseignée. Une table *locale* (stockée dans le fichier courant) a une propriété `Connect` vide ; une table *liée* à un back-end Access a une `Connect` de la forme `;DATABASE=chemin_du_back-end`. Une table liée à une source externe via ODBC a, elle, une `Connect` débutant par `ODBC;`. C'est précisément cette propriété qui permet de reconnaître et de traiter les liens par code.

La gestion programmatique des tables liées au sens large (création, suppression, inspection) est traitée à la section 12.7, et l'architecture DAO sous-jacente au chapitre 9. La présente section se concentre sur l'angle du déploiement : détecter automatiquement des liens invalides et les rétablir vers le bon back-end.

## Lire et extraire le chemin du back-end actuel

Avant de reconstruire un lien, il est utile de savoir vers où il pointe. Le chemin se trouve dans la propriété `Connect`, après le marqueur `DATABASE=`. Une chaîne de connexion peut toutefois comporter d'autres éléments à la suite (notamment un mot de passe), c'est pourquoi l'extraction doit s'arrêter au point-virgule suivant :

```vba
' Extrait le chemin du back-end à partir d'une chaîne Connect de type ";DATABASE=..."
Public Function CheminBackEnd(ByVal chaineConnect As String) As String
    Const MARQUEUR As String = "DATABASE="
    Dim p As Long, q As Long

    CheminBackEnd = ""
    p = InStr(1, chaineConnect, MARQUEUR, vbTextCompare)
    If p = 0 Then Exit Function

    p = p + Len(MARQUEUR)
    q = InStr(p, chaineConnect, ";")        ' éventuel ;PWD= qui suit le chemin
    If q = 0 Then
        CheminBackEnd = Mid$(chaineConnect, p)
    Else
        CheminBackEnd = Mid$(chaineConnect, p, q - p)
    End If
End Function
```

## Détecter qu'une reliaison est nécessaire

Plutôt que de relier systématiquement à chaque démarrage, on commence par vérifier si les liens existants sont valides : il suffit qu'au moins un lien vers un back-end Access pointe vers un fichier introuvable pour conclure qu'une reliaison s'impose. Cette vérification est peu coûteuse et peut être exécutée à chaque lancement : si tout est en ordre, elle ne fait rien.

```vba
' Renvoie True si au moins un lien Access pointe vers un back-end introuvable
Public Function ReliaisonNecessaire() As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim chemin As String

    ReliaisonNecessaire = False
    Set db = CurrentDb

    For Each tdf In db.TableDefs
        ' Ne considérer que les liens vers un back-end Access (exclure ODBC)
        If InStr(tdf.Connect, "DATABASE=") > 0 And InStr(tdf.Connect, "ODBC;") = 0 Then
            chemin = CheminBackEnd(tdf.Connect)
            If Len(chemin) = 0 Or Dir(chemin) = "" Then
                ReliaisonNecessaire = True
                Exit For
            End If
        End If
    Next tdf

    Set tdf = Nothing
    Set db = Nothing
End Function
```

La fonction `Dir` renvoyant une chaîne vide lorsque le fichier n'existe pas, et fonctionnant sur les chemins UNC, elle constitue un test de validité simple et fiable.

## Effectuer la reliaison par code

La reliaison proprement dite parcourt l'ensemble des tables liées, remplace leur chaîne de connexion par le chemin correct, puis applique la méthode `RefreshLink` qui rétablit effectivement la liaison. L'opération doit impérativement être encadrée par une gestion d'erreurs (chapitre 13), car `RefreshLink` peut échouer si le fichier est introuvable, verrouillé, ou si les droits d'accès sont insuffisants :

```vba
' Relie toutes les tables liées Access vers le nouveau back-end
Public Function RelierBackEnd(ByVal nouveauChemin As String, _
                              Optional ByVal motDePasse As String = "") As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim nouvelleConnexion As String

    On Error GoTo Gestion

    nouvelleConnexion = ";DATABASE=" & nouveauChemin
    If Len(motDePasse) > 0 Then
        nouvelleConnexion = nouvelleConnexion & ";PWD=" & motDePasse
    End If

    Set db = CurrentDb
    For Each tdf In db.TableDefs
        ' Ne traiter que les liens vers un back-end Access (exclure ODBC)
        If InStr(tdf.Connect, "DATABASE=") > 0 And InStr(tdf.Connect, "ODBC;") = 0 Then
            tdf.Connect = nouvelleConnexion
            tdf.RefreshLink
        End If
    Next tdf

    RelierBackEnd = True

Sortie:
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

Gestion:
    RelierBackEnd = False
    MsgBox "Échec de la reliaison : " & Err.Description, vbCritical
    Resume Sortie
End Function
```

Le filtrage sur la chaîne de connexion est essentiel : il garantit que seuls les liens vers le back-end Access sont modifiés, et que d'éventuels liens ODBC vers des sources externes (SQL Server, par exemple) ne sont pas écrasés par mégarde.

## Trouver le bon chemin du back-end

La question centrale de la reliaison automatique est : *d'où le front-end tire-t-il le chemin correct ?* Coder ce chemin en dur dans l'application irait à l'encontre du but recherché, puisqu'il faudrait recompiler à chaque changement d'emplacement. Les approches viables consistent à externaliser cette information :

- Un **fichier de configuration** (`.ini` ou `.txt`) déposé à un emplacement connu, contenant le chemin du back-end.
- Le **registre Windows**, dont la lecture et l'écriture sont traitées à la section 22.9.
- Un **emplacement réseau par défaut** convenu pour l'organisation.

Dans tous les cas, il est prudent de prévoir un **repli** : si le chemin configuré est invalide ou absent, on demande à l'utilisateur — ou à l'administrateur lors de l'installation — de localiser lui-même le back-end.

### Le repli par boîte de dialogue

La boîte de dialogue de sélection de fichier d'Office permet à l'utilisateur de parcourir le système et de désigner le fichier back-end. Elle s'appuie sur la bibliothèque Office (référence évoquée à la section 2.5) :

```vba
' Laisse l'utilisateur localiser le fichier back-end
Public Function ChoisirBackEnd() As String
    Dim fd As Office.FileDialog

    ChoisirBackEnd = ""
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Localiser le fichier de données de l'application"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Base de données Access", "*.accdb;*.mdb"
        If .Show = -1 Then                  ' -1 : l'utilisateur a validé
            ChoisirBackEnd = .SelectedItems(1)
        End If
    End With

    Set fd = Nothing
End Function
```

## Mémoriser le chemin résolu

Une fois la reliaison réussie à partir d'un chemin choisi manuellement, il serait pénible de redemander cet emplacement à chaque démarrage. La bonne pratique consiste à **mémoriser le chemin qui a fonctionné** — dans un fichier de configuration, dans le registre (section 22.9) ou dans une table locale — afin que les lancements suivants l'utilisent directement, sans solliciter l'utilisateur. La reliaison ne redevient interactive que si ce chemin mémorisé cesse à son tour d'être valide.

## Reliaison au démarrage : orchestration complète

L'ensemble s'assemble en une séquence de démarrage qui s'exécute avant tout accès aux données. Elle enchaîne logiquement la détection, la tentative automatique à partir du chemin connu, et le repli interactif :

```vba
' À appeler au démarrage, avant tout accès aux données
Public Function InitialiserLiaisons() As Boolean
    Dim chemin As String

    InitialiserLiaisons = True
    If Not ReliaisonNecessaire() Then Exit Function     ' liens déjà valides

    ' 1) Tenter le chemin mémorisé ou configuré par défaut
    chemin = CheminBackEndConfigure()                   ' fichier config, registre...
    If Len(chemin) > 0 And Dir(chemin) <> "" Then
        If RelierBackEnd(chemin) Then Exit Function
    End If

    ' 2) Repli : demander à l'utilisateur de localiser le back-end
    chemin = ChoisirBackEnd()
    If Len(chemin) = 0 Then
        MsgBox "Le fichier de données est introuvable. " & _
               "L'application va se fermer.", vbCritical
        InitialiserLiaisons = False
        Exit Function
    End If

    ' 3) Relier, puis mémoriser le chemin pour les prochains démarrages
    If RelierBackEnd(chemin) Then
        MemoriserCheminBackEnd chemin
    Else
        InitialiserLiaisons = False
    End If
End Function
```

Les procédures `CheminBackEndConfigure` et `MemoriserCheminBackEnd` évoquées ici encapsulent respectivement la lecture et l'écriture du chemin (fichier de configuration ou registre). Cette séquence est particulièrement importante **juste après une mise à jour automatique** du front-end (section 21.3), où la copie fraîche arrive nécessairement avec les liens du développeur et doit donc être reliée avant la première utilisation.

## Cas particuliers

### Back-end protégé par mot de passe

Si le back-end est chiffré par un mot de passe de base de données (section 20.2), la chaîne de connexion des liens doit inclure ce mot de passe, sous la forme `;DATABASE=chemin;PWD=motdepasse`. C'est la raison pour laquelle la fonction `RelierBackEnd` ci-dessus accepte un paramètre de mot de passe optionnel. Il convient de gérer la confidentialité de ce mot de passe avec prudence plutôt que de le coder en clair dans l'application.

### Liens ODBC vers des serveurs externes

Les tables liées à une source externe (SQL Server, par exemple, dans une architecture hybride) utilisent une chaîne de connexion d'un tout autre format, débutant par `ODBC;`. Leur reliaison obéit à des règles spécifiques et ne doit jamais être confondue avec celle des liens Access — d'où le filtrage systématique excluant `ODBC;` dans les exemples ci-dessus. La reliaison de sources ODBC est abordée dans le cadre des connexions externes (chapitre 10) et de l'architecture hybride Access + SQL Server (section 23.4).

### Liens vers plusieurs back-ends

Certaines applications répartissent leurs données sur plusieurs fichiers back-end. Dans ce cas, le filtrage sur le seul `DATABASE=` ne suffit plus : il faut distinguer les liens par leur chemin d'origine afin de ne rediriger que ceux concernés par le déplacement, et conserver intacts les liens vers les autres back-ends.

## Articulation avec le reste du déploiement

La reliaison automatique est étroitement liée aux autres mécanismes du chapitre :

- Elle fait suite à la copie d'un front-end neuf par l'**auto-update** (section 21.3), qui apporte des liens à corriger.
- Elle privilégie les **chemins UNC** plutôt que les lecteurs mappés pour la stabilité des liens (section 21.5).
- Elle s'intègre à l'**installation automatisée** (section 21.8), qui établit la première liaison sur un poste vierge.
- Elle repose sur la **gestion des tables liées** en DAO (section 12.7) et exige une **gestion d'erreurs** robuste (chapitre 13).

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans la mise en place d'une reliaison automatique :

- **Coder le chemin du back-end en dur**, ce qui oblige à recompiler à chaque déplacement et rend la reliaison automatique illusoire.
- **Oublier de relier après une mise à jour** du front-end : la copie fraîche conserve les liens du développeur et échoue à atteindre les données tant qu'elle n'a pas été reliée.
- **Écraser des liens ODBC** en ne filtrant pas correctement les chaînes de connexion lors de la reliaison.
- **Privilégier un lecteur mappé** (`S:\`) plutôt qu'un chemin UNC, fragilisant les liens d'un poste à l'autre.
- **Négliger la gestion d'erreurs** autour de `RefreshLink`, et laisser l'application planter lorsque le back-end est momentanément inaccessible.
- **Ne pas mémoriser le chemin résolu**, et imposer à l'utilisateur de relocaliser le back-end à chaque démarrage.

⏭️ [21.5. Déploiement en réseau local et gestion des chemins UNC](/21-deploiement-distribution/05-deploiement-reseau-chemins-unc.md)
