🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.6. Simulation de données et environnements de test

Les sections précédentes ont établi que les tests exigent un état de départ connu et un lieu d'exécution distinct de la production. Reste à savoir comment s'en doter. Cette section traite de deux moyens complémentaires : **simuler des données** — en générer pour peupler une base d'essai — et **organiser des environnements de test** isolés, que l'on sait mettre en place, réinitialiser et sélectionner. Elle prolonge directement les techniques des sections [19.3](03-tests-modules-donnees.md) et [19.4](04-tests-unitaires-framework.md).

## Deux besoins distincts

Il faut bien séparer les deux. La **simulation de données** consiste à produire des enregistrements — en volume et en variété — pour alimenter une base. L'**environnement de test** est l'espace contrôlé où l'on exécute, séparé de la production, avec les mécanismes pour le provisionner et le remettre à zéro. L'un sans l'autre ne suffit pas : des données réalistes dans la base de production ne valent rien, et un environnement isolé mais vide ne teste pas grand-chose.

## Pourquoi ne jamais tester sur la production

Développer ou tester sur les données réelles est à proscrire. Le risque de **corruption** ou de modification accidentelle est réel ; les données contiennent des **informations sensibles** qu'on ne doit pas exposer en développement ; et la production n'est pas **reproductible** — son état change en permanence, ce qui rend tout test non répétable. D'où la double nécessité de données simulées et d'un environnement à part.

## Simuler des données

### Jeux de test manuels

Pour vérifier un comportement précis, rien ne vaut un petit ensemble d'enregistrements **conçus à la main** : le cas typique, le cas limite, le cas invalide. Précis et lisibles, ils s'insèrent à la mise en place du test (voir les sections [19.3](03-tests-modules-donnees.md) et [19.4](04-tests-unitaires-framework.md)) et conviennent aux tests unitaires ou d'intégration ciblés.

### Génération programmatique

Pour le volume, on génère par code, en piochant des valeurs plausibles dans des tableaux. Envelopper l'insertion en masse dans une transaction l'accélère considérablement, et l'ouverture en `dbAppendOnly` allège encore l'opération.

```vba
Public Sub GenererClients(ByVal nb As Long)
    If EnvironnementCourant() = "PROD" Then Exit Sub   ' garde-fou (voir plus bas)

    Dim ws As DAO.Workspace, db As DAO.Database, rs As DAO.Recordset
    Dim prenoms As Variant, noms As Variant, villes As Variant, i As Long
    prenoms = Array("Marie", "Jean", "Sophie", "Luc", "Claire", "Paul")
    noms = Array("Martin", "Bernard", "Dubois", "Petit", "Durand")
    villes = Array("Paris", "Lyon", "Marseille", "Toulouse", "Nantes")

    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Clients", dbOpenDynaset, dbAppendOnly)

    ws.BeginTrans                                  ' accélère fortement l'insertion en masse
    For i = 1 To nb
        rs.AddNew
        rs!Prenom = prenoms(Int(Rnd * (UBound(prenoms) + 1)))
        rs!Nom = noms(Int(Rnd * (UBound(noms) + 1)))
        rs!Ville = villes(Int(Rnd * (UBound(villes) + 1)))
        rs!DateInscription = DateSerial(2020, 1, 1) + Int(Rnd * 2000)
        rs!Solde = Round(Rnd * 10000, 2)
        rs.Update
    Next i
    ws.CommitTrans
    rs.Close
End Sub
```

> ℹ️ L'accélération par transaction est traitée au [chapitre 14.2](../14-transactions/02-transactions-dao.md) ; l'option `dbAppendOnly` à la [section 18.2](../18-optimisation-performance/02-optimisation-recordsets.md). Pour de très gros volumes, une requête `INSERT … SELECT` peut surpasser la boucle VBA.

### Couvrir les cas limites

Un bon jeu de données ne se contente pas du cas nominal : il inclut les **valeurs nulles**, les chaînes vides ou de longueur maximale, les **bornes** (dates extrêmes, zéro, valeurs négatives), les **doublons** (pour éprouver les contraintes d'unicité) et les **caractères spéciaux** — l'apostrophe, notamment, qui révèle les failles de construction SQL (voir le [chapitre 11.5](../11-sql-access-vba/05-sql-parametree.md)). Les formats locaux (dates, décimales) méritent aussi d'être testés (voir le [chapitre 11.7](../11-sql-access-vba/07-localisation-formatage-sql.md)).

### Données de production anonymisées

Pour un réalisme et un volume représentatifs, on peut partir d'une **copie de la production dont on masque les champs sensibles** — noms, adresses, courriels remplacés par des valeurs factices. On obtient ainsi des distributions réalistes sans exposer de données personnelles. Cette technique est puissante, mais l'anonymisation doit être **complète et irréversible** : c'est une exigence de confidentialité, non une option.

> ℹ️ Le traitement des données sensibles est abordé au [chapitre 20.8](../20-securite-protection/08-chiffrement-donnees-sensibles.md).

## Reproductibilité

Un test doit être déterministe. Or `Rnd` produit, en l'absence de `Randomize`, **la même séquence à chaque exécution** — ce qui, paradoxalement, est utile : on obtient des données identiques d'un essai à l'autre. Si l'on appelle `Randomize` pour varier les données, alors on s'abstient d'**affirmer des valeurs générées précises** et l'on vérifie plutôt des invariants ou des comptes. Dans tous les cas, l'annulation par transaction (voir la [section 19.3](03-tests-modules-donnees.md)) garantit que la base retrouve son état, quelle que soit la donnée injectée.

## Mettre en place un environnement de test

### Séparer développement, test et production

Chaque environnement dispose de son propre back-end (et parfois de son front-end) : développement, test, production. La règle est absolue — on ne teste jamais contre le back-end de production.

### Sélectionner l'environnement par configuration

Le front-end ne doit pas coder en dur le chemin du back-end : il le détermine par **configuration** (table de paramètres, fichier, clé de registre, `TempVar`, ou paramètre de ligne de commande). On peut alors diriger le même front-end vers l'environnement voulu sans toucher au code.

```vba
Function CheminBackEnd() As String
    Select Case EnvironnementCourant()          ' lu depuis la configuration
        Case "PROD": CheminBackEnd = "\\serveur\prod\Donnees_BE.accdb"
        Case "TEST": CheminBackEnd = "C:\Test\Donnees_BE.accdb"
        Case "DEV":  CheminBackEnd = "C:\Dev\Donnees_BE.accdb"
    End Select
End Function
```

> ℹ️ La configuration peut s'appuyer sur les `TempVars` ([15.7](../15-multi-utilisateurs/07-tempvars.md)), le registre ([22.9](../22-api-windows-integration-office/09-registre-windows.md)) ou la ligne de commande ([21.8](../21-deploiement-distribution/08-installation-automatisee.md)). La reliaison des tables vers le bon back-end est traitée aux chapitres [15.9](../15-multi-utilisateurs/09-tables-liees-par-code.md) et [21.4](../21-deploiement-distribution/04-reliaison-automatique-tables.md).

### Réinitialiser depuis un modèle

Pour repartir d'un état vierge et reproductible à chaque campagne de tests, on conserve un **back-end modèle** (un fichier de référence, vide ou pré-rempli) que l'on copie sur la base de travail avant exécution.

```vba
Public Sub ReinitialiserBaseTest()
    If EnvironnementCourant() = "PROD" Then Exit Sub   ' garde-fou
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Const MODELE As String = "C:\Test\Modele_BE.accdb"
    Const TEST As String = "C:\Test\Donnees_BE.accdb"
    If fso.FileExists(TEST) Then fso.DeleteFile TEST
    fso.CopyFile MODELE, TEST              ' état vierge reproductible
    ' (on relie ensuite les tables du front-end à TEST)
End Sub
```

Selon le besoin, on peut aussi recréer le schéma par code (DDL ou `TableDef`) puis l'alimenter, ou — pour une isolation plus fine au sein d'une même exécution — s'en remettre à l'annulation par transaction plutôt qu'à la recréation du fichier.

> ℹ️ La copie de fichiers via `FileSystemObject` est traitée au [chapitre 22.10](../22-api-windows-integration-office/10-filesystemobject.md), la création de tables par code au [chapitre 12.5](../12-querydefs-tabledefs/05-creer-modifier-tables.md).

## Garde-fous

Manipuler des environnements multiples appelle des protections. On rend **visible** l'environnement courant — un bandeau ou un titre d'application affichant « TEST » — pour ne jamais confondre une base d'essai avec la production. Et l'on **protège les opérations destructrices** (génération, réinitialisation) en les faisant refuser de s'exécuter contre la production, comme le montre le test `EnvironnementCourant() = "PROD"` placé en tête des routines précédentes.

```vba
If EnvironnementCourant() = "PROD" Then
    MsgBox "Opération interdite en production.", vbCritical
    Exit Sub
End If
```

## Intégration aux tests

Ces mécanismes alimentent le cycle de test : la mise en place réinitialise l'environnement (copie du modèle ou `BeginTrans`) et génère les jeux nécessaires ; le test s'exécute et vérifie ; le nettoyage annule ou réinitialise. Disposer d'un environnement de **volume réaliste** est par ailleurs indispensable aux tests de performance : une base de production simulée révèle les requêtes lentes et les index manquants qu'une base de développement de quelques lignes dissimule.

> ℹ️ Les tests de performance s'appuient sur le profilage ([18.1](../18-optimisation-performance/01-profilage-mesure-performances.md)) et l'indexation ([18.6](../18-optimisation-performance/06-indexation-performances-sql.md)).

## Points clés à retenir

- Tests et performance exigent des données contrôlées et un environnement distinct ; on ne teste jamais sur la production (corruption, données sensibles, non-reproductibilité).
- On simule les données par des jeux manuels (cas précis), par génération programmatique (volume, accélérée par transaction et `dbAppendOnly`), et en couvrant systématiquement les cas limites.
- Une copie de production anonymisée donne du réalisme — à condition d'un masquage complet et irréversible des données personnelles.
- Pour la reproductibilité : exploiter la séquence fixe de `Rnd`, ou ne pas affirmer les valeurs générées ; l'annulation par transaction garde la base propre.
- L'environnement se sélectionne par configuration (jamais de chemin codé en dur) et se réinitialise depuis un back-end modèle copié avant chaque campagne.
- Des garde-fous s'imposent : rendre l'environnement visible et empêcher toute opération destructrice de s'exécuter contre la production.

---


⏭️ [19.7. Débogage à distance et techniques sur poste utilisateur](/19-debogage-tests/07-debogage-distance.md)
