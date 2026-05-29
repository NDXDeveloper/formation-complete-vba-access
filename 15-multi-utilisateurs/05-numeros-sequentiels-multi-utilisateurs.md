🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 15.5. Génération de numéros séquentiels fiables en multi-utilisateur (DMax+1, table de compteurs)

Attribuer un numéro séquentiel — numéro de facture, de commande, de dossier — paraît trivial en mono-utilisateur, et devient un piège classique dès que plusieurs personnes saisissent en même temps. C'est l'illustration parfaite des problèmes de concurrence évoqués depuis le début du chapitre : deux utilisateurs qui calculent « le prochain numéro » au même instant obtiennent le même résultat. Cette section explique pourquoi, et présente les techniques qui garantissent l'unicité et, si nécessaire, l'absence de trous.

## Pourquoi pas le NuméroAuto ?

La tentation première est d'utiliser un champ **NuméroAuto** comme numéro métier. C'est une erreur. Le NuméroAuto est conçu pour fournir une clé primaire de substitution unique, pas une numérotation séquentielle continue. Il présente plusieurs caractéristiques rédhibitoires pour cet usage :

- il peut comporter des **trous** : un enregistrement supprimé ne libère pas son numéro, une insertion échouée le consomme, et — comme vu à la [section 14.1](/14-transactions/01-concept-transaction.md) — un `Rollback` ne récupère pas le numéro déjà attribué ;
- il n'est pas garanti **strictement croissant ni contigu**, et peut même être configuré en mode aléatoire ;
- sa valeur de départ et son pas ne se maîtrisent pas comme on l'attendrait d'une numérotation métier.

Le NuméroAuto reste excellent comme **clé technique interne**. Mais pour un numéro visible et signifiant (« FACT-2024-0001 »), surtout s'il doit être continu pour des raisons comptables ou légales, il faut générer le numéro soi-même.

## Le piège du DMax+1 naïf

L'approche la plus répandue consiste à lire le plus grand numéro existant et à lui ajouter 1 :

```vba
NouveauNum = Nz(DMax("NumCommande", "Commandes"), 0) + 1
```

En mono-utilisateur, cela fonctionne. En multi-utilisateur, cette ligne cache une **situation de compétition** (*race condition*). Considérons deux utilisateurs A et B :

1. A évalue `DMax` et obtient 1000 ; il calcule 1001.
2. Avant que A n'ait enregistré, B évalue `DMax` et obtient toujours 1000 ; il calcule 1001 lui aussi.
3. A enregistre la commande 1001.
4. B enregistre la commande 1001.

Résultat : **deux commandes portent le numéro 1001**. La fenêtre entre la lecture et l'écriture est courte, mais sous charge concurrente, la collision finit immanquablement par se produire. `DMax+1` employé seul n'est donc fiable qu'en mono-utilisateur. (`DMax` est une fonction de domaine, abordée à la [section 11.10](/11-sql-access-vba/10-fonctions-domaine.md).)

## Le filet de sécurité indispensable : l'index unique

Avant toute technique, une règle absolue : le champ de numérotation doit porter un **index unique** (voir [section 12.5](/12-querydefs-tabledefs/05-creer-modifier-tables.md)). Cet index est la garantie *structurelle* qu'aucun doublon ne pourra jamais être enregistré, quelle que soit la qualité du code. Si une collision survient malgré tout, l'index la rejette en levant l'**erreur 3022** (création de valeurs en double dans un index ou une clé). C'est sur ce rejet que s'appuient les stratégies robustes : l'index empêche le doublon, et le code réagit à l'erreur 3022.

Sans index unique, aucune technique de numérotation n'est réellement sûre en multi-utilisateur.

## Solution 1 : DMax+1 protégé par transaction et réessai

La première solution conserve la simplicité du `DMax+1`, mais l'entoure de deux protections : un **index unique** (le filet) et une **boucle de réessai** sur l'erreur 3022. Le principe : on calcule le numéro et on tente l'insertion ; si l'index rejette un doublon, on recalcule et on retente.

```vba
Public Function ObtenirNumCommande() As Long
    ' Prérequis : index UNIQUE sur Commandes.NumCommande
    Const MAX_TENTATIVES As Integer = 5
    Dim db As DAO.Database
    Dim tentative As Integer
    Dim num As Long
    Dim n As Long, d As String

    Set db = CurrentDb
    For tentative = 1 To MAX_TENTATIVES
        num = Nz(DMax("NumCommande", "Commandes"), 0) + 1

        On Error Resume Next
        Err.Clear
        db.Execute "INSERT INTO Commandes (NumCommande, DateCommande) " & _
                   "VALUES (" & num & ", Date());", dbFailOnError

        Select Case Err.Number
            Case 0                              ' succès
                On Error GoTo 0
                ObtenirNumCommande = num
                Exit Function
            Case 3022                           ' doublon : un autre a pris ce numéro
                ' on boucle et on recalcule DMax+1
            Case Else
                n = Err.Number: d = Err.Description
                On Error GoTo 0
                Err.Raise n, , d
        End Select
    Next tentative

    On Error GoTo 0
    Err.Raise vbObjectError + 1, , _
              "Impossible d'obtenir un numéro après plusieurs tentatives."
End Function
```

Deux principes accompagnent cette approche : **calculer le numéro le plus tard possible** (au moment de l'écriture, pas à l'ouverture du formulaire) et **enregistrer immédiatement**, pour réduire au minimum la fenêtre de compétition. L'index unique fait le reste.

Cette solution est simple et efficace, mais elle n'empêche pas les **trous** : si une insertion échoue après coup, le numéro calculé n'est pas réutilisé. Elle convient donc lorsque l'unicité est exigée mais que de petits trous sont tolérables.

## Solution 2 : la table de compteurs

La seconde solution est la plus robuste, et la seule réellement adaptée aux numérotations qui doivent être **contiguës**. Elle repose sur une table dédiée qui mémorise, pour chaque séquence, la dernière valeur utilisée — par exemple `Compteurs(NomSequence, DerniereValeur)`.

Pour obtenir un numéro, on ouvre la ligne du compteur en **verrouillage pessimiste** (et dans une transaction), on lit la valeur, on l'incrémente, on la réécrit, puis on libère le verrou. Le verrou **sérialise** les accès : si deux utilisateurs sollicitent le compteur en même temps, le second attend que le premier ait terminé, et obtient donc nécessairement la valeur suivante.

```vba
' Table : Compteurs(NomSequence Texte [indexé sans doublon], DerniereValeur Entier long)
Public Function ProchainNumero(nomSequence As String) As Long
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim num As Long
    Dim transOuverte As Boolean

    On Error GoTo Gestion_Erreur
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb

    ws.BeginTrans
    transOuverte = True

    Set rs = db.OpenRecordset( _
        "SELECT * FROM Compteurs WHERE NomSequence = '" & nomSequence & "';", _
        dbOpenDynaset)
    rs.LockEdits = True                   ' verrou pessimiste sur la ligne du compteur

    If rs.EOF Then
        rs.AddNew                         ' première utilisation de cette séquence
        rs!NomSequence = nomSequence
        num = 1
    Else
        rs.Edit
        num = rs!DerniereValeur + 1
    End If
    rs!DerniereValeur = num
    rs.Update                             ' libère le verrou

    ws.CommitTrans
    transOuverte = False

    rs.Close
    Set rs = Nothing: Set db = Nothing: Set ws = Nothing
    ProchainNumero = num
    Exit Function

Gestion_Erreur:
    If transOuverte Then ws.Rollback
    ' Un verrou transitoire sur le compteur peut être retenté (cf. section 15.4)
    Err.Raise Err.Number, , Err.Description
End Function
```

Le verrou n'étant tenu que le temps d'un lire-incrémenter-écrire, la contention reste minime malgré la sérialisation.

Point crucial pour une numérotation **sans trou** : l'incrément du compteur et la création de l'enregistrement qui l'utilise doivent appartenir à la **même transaction**. Ainsi, si l'enregistrement final échoue, le `Rollback` annule aussi l'incrément, et le numéro n'est pas perdu. À l'inverse, allouer le numéro dans une transaction séparée est plus simple mais peut laisser un trou en cas d'échec ultérieur. Le choix dépend de l'exigence : pour des numéros de facture soumis à des contraintes comptables, la contiguïté justifie d'englober numéro et enregistrement dans une transaction unique.

## Numéros composés et réinitialisés

Les numéros métier sont rarement de simples entiers : « FACT-2024-0001 », séquences réinitialisées chaque année, ou propres à un client. La table de compteurs s'y prête naturellement : il suffit d'utiliser une **clé de séquence par périmètre**, par exemple `NomSequence = "Facture-2024"`. Le passage à l'année suivante crée automatiquement une nouvelle ligne `"Facture-2025"` repartant de 1, sans logique supplémentaire. Le numéro affiché est ensuite composé par mise en forme (`"FACT-" & annee & "-" & Format(num, "0000")`).

Avec l'approche `DMax+1`, l'équivalent consiste à filtrer le domaine — `DMax("NumSeq", "Factures", "Annee = " & Year(Date()))` — et à poser un **index unique composé** sur `(Annee, NumSeq)`, qui joue alors le rôle de filet de sécurité par année.

## Où attribuer le numéro dans un formulaire

Le principe directeur découle de tout ce qui précède : **attribuer le numéro le plus tard possible**, juste avant l'enregistrement, jamais à l'ouverture du formulaire ni dès le début de la saisie — sans quoi la valeur calculée serait déjà périmée au moment du `Update`.

En pratique, sur un formulaire lié, on génère le numéro dans l'événement `BeforeUpdate` (en le réservant aux nouveaux enregistrements via `Me.NewRecord`), ou dans une routine d'enregistrement explicite. L'approche la plus sûre reste d'effectuer la génération dans le code d'écriture lui-même, protégé par les mécanismes des solutions 1 ou 2.

## Quelle approche choisir ?

`DMax+1` protégé (index unique + réessai) convient lorsque l'on veut surtout garantir l'**unicité** et que de petits trous sont acceptables : c'est simple et suffisant pour beaucoup de cas. La **table de compteurs** s'impose lorsque la numérotation doit être **contiguë et maîtrisée** (factures, séquences légales), ou lorsque le périmètre est composite (par année, par site). Dans tous les cas, l'**index unique** est non négociable.

## Points de vigilance

- **Jamais de NuméroAuto comme numéro métier continu** : trous garantis, valeur de départ non maîtrisée.
- **Index unique obligatoire** sur le champ (ou la combinaison) de numérotation : c'est le seul garde-fou réellement sûr.
- **`DMax+1` seul est dangereux en multi-utilisateur** : à protéger systématiquement par index unique + réessai sur 3022.
- **Calculer tard, enregistrer vite** : minimiser la fenêtre entre lecture et écriture.
- **Contiguïté = même transaction** : pour des numéros sans trou, allouer le numéro et créer l'enregistrement dans une seule transaction.
- **Compteur verrouillé brièvement** : le verrou pessimiste sur la ligne du compteur doit être tenu le minimum de temps.

## En résumé

Générer un numéro séquentiel fiable en multi-utilisateur exige de contourner la situation de compétition du `DMax+1` naïf, qui produit des doublons lorsque deux utilisateurs lisent la même valeur maximale simultanément. Le NuméroAuto ne convient pas comme numéro métier continu (trous, absence de contrôle). Deux solutions s'offrent, reposant toutes deux sur un **index unique** indispensable : `DMax+1` entouré d'une **transaction et d'un réessai** sur l'erreur 3022 (simple, mais tolère des trous), ou une **table de compteurs** lue en verrouillage pessimiste (robuste, et contiguë si l'incrément partage la transaction de l'enregistrement). Pour les numéros composés ou réinitialisés par année, une clé de séquence par périmètre dans la table de compteurs est la solution la plus propre. Et dans un formulaire, le numéro s'attribue le plus tard possible, juste avant l'écriture.

La section suivante revient sur le contrôle du verrouillage au niveau de l'interface : la [propriété RecordLocks des formulaires](06-recordlocks-formulaires.md).

⏭️ [15.6. Propriété RecordLocks des formulaires](/15-multi-utilisateurs/06-recordlocks-formulaires.md)
