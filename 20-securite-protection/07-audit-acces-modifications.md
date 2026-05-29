🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 20.7. Audit des accès et des modifications par code

La gestion des droits vue en [section 20.6](06-gestion-droits-utilisateurs.md) répond à la question de l'autorisation : qui a le droit d'agir. L'audit répond à une question complémentaire et tout aussi importante : qui a effectivement agi, quand, et sur quoi. Tenir un journal des accès et des modifications apporte la traçabilité indispensable pour investiguer un incident, reconstituer l'historique d'un enregistrement, satisfaire à des obligations de conformité — par exemple sur les données personnelles — ou simplement responsabiliser les utilisateurs.

Il faut d'emblée comprendre la nature de l'audit : c'est un dispositif de détection, non de prévention. Il n'empêche aucune action ; il en conserve la trace après coup. À ce titre, il ne remplace pas la gestion des droits, il la prolonge.

On distingue deux grandes formes d'audit, que cette section traite successivement. L'audit des accès enregistre les connexions à l'application : qui l'a ouverte, quand, depuis quelle machine. L'audit des modifications enregistre les changements de données : quel enregistrement a été créé, modifié ou supprimé, par qui, et idéalement avec les valeurs avant et après.

## Concevoir les tables de journalisation

Deux tables suffisent, l'une pour les connexions, l'autre pour les modifications. Elles ont vocation à résider dans le back-end partagé, afin de centraliser les traces de tous les utilisateurs. Par convention, ces tables sont alimentées uniquement en insertion : l'application ne propose jamais d'interface pour les modifier ou les supprimer.

```sql
CREATE TABLE tblJournalConnexions (
    ConnexionID          AUTOINCREMENT PRIMARY KEY,
    IdentifiantWindows   TEXT(100),
    NomMachine           TEXT(100),
    DateHeureConnexion   DATETIME,
    DateHeureDeconnexion DATETIME
);

CREATE TABLE tblJournalModifications (
    ModifID           AUTOINCREMENT PRIMARY KEY,
    DateHeure         DATETIME,
    IdentifiantWindows TEXT(100),
    NomTable          TEXT(64),
    CleEnregistrement TEXT(100),
    Action            TEXT(20),
    NomChamp          TEXT(64),
    AncienneValeur    LONGTEXT,
    NouvelleValeur    LONGTEXT
);
```

Pour le journal des modifications, l'approche retenue ici enregistre une ligne par champ modifié, ce qui offre une granularité fine et permet de savoir exactement quelle donnée a changé et comment. Une alternative plus grossière consisterait à enregistrer une seule ligne par enregistrement, avec une représentation textuelle de l'avant et de l'après. La version champ par champ est généralement préférable pour des données sensibles. Les colonnes de valeurs sont de type texte long afin d'accueillir des contenus de toute taille.

## Journaliser les connexions

L'audit des accès se branche au démarrage et à la fermeture de l'application. À l'ouverture, on insère une ligne consignant l'identité et la machine, et l'on mémorise l'identifiant de cette ligne pour pouvoir, à la fermeture, y renseigner l'heure de déconnexion. Le nom de la machine est obtenu via la variable d'environnement `COMPUTERNAME`, ou par l'API `GetComputerName` présentée en [section 22.2](../22-api-windows-integration-office/02-api-courantes.md).

```vba
' === Module : modAudit ===
Option Compare Database
Option Explicit

Public Sub JournaliserConnexion()
    ' À appeler une fois au démarrage (AutoExec ou formulaire de démarrage).
    On Error GoTo Echec
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblJournalConnexions", dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    rs!IdentifiantWindows = Environ$("USERNAME")
    rs!NomMachine = Environ$("COMPUTERNAME")
    rs!DateHeureConnexion = Now()
    rs.Update

    ' Récupérer l'identifiant auto-généré (voir 9.6) pour la déconnexion
    rs.Bookmark = rs.LastModified
    TempVars!ConnexionID = rs!ConnexionID

Nettoyage:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
Echec:
    ' Un échec d'audit ne doit jamais empêcher l'application de démarrer.
    Resume Nettoyage
End Sub

Public Sub JournaliserDeconnexion()
    ' À appeler à la fermeture. Au mieux : non garanti en cas de plantage.
    On Error Resume Next
    If IsNull(TempVars!ConnexionID) Then Exit Sub
    CurrentDb.Execute "UPDATE tblJournalConnexions " & _
                      "SET DateHeureDeconnexion = Now() " & _
                      "WHERE ConnexionID = " & TempVars!ConnexionID, dbFailOnError
End Sub
```

Il faut être conscient que l'heure de déconnexion ne peut être capturée que si l'application se ferme proprement. Un plantage, une coupure de courant ou une fermeture forcée laisseront la déconnexion non renseignée. C'est une limite inhérente, à accepter : la connexion, elle, est toujours fiablement enregistrée.

## Journaliser les modifications : le bon endroit, les événements de formulaire

L'endroit naturel pour tracer les changements de données est l'événement de formulaire, car c'est là que l'on dispose à la fois des valeurs et du contexte. Deux événements jouent un rôle clé.

L'événement `BeforeUpdate` se déclenche juste avant l'enregistrement d'une création ou d'une modification. Access y met à disposition une fonctionnalité précieuse : chaque contrôle lié expose, en plus de sa valeur courante via `Value`, la valeur antérieure via `OldValue`. En comparant les deux, on identifie précisément ce qui a changé. La propriété `NewRecord` du formulaire indique par ailleurs s'il s'agit d'une création, cas où `OldValue` est vide et où tout est considéré comme nouveau.

L'événement `Delete` se déclenche pour chaque enregistrement sur le point d'être supprimé. Comme la suppression n'a pas encore eu lieu, on peut y lire les valeurs de l'enregistrement avant qu'elles ne disparaissent.

## Un module d'audit réutilisable

On factorise l'écriture dans le journal en une procédure unique. Elle utilise un recordset DAO en mode ajout plutôt qu'une requête SQL, ce qui évite tout problème de guillemets ou d'apostrophes avec des valeurs arbitraires, et s'apparente à du SQL paramétré (voir [section 11.5](../11-sql-access-vba/05-sql-parametree.md)).

```vba
Public Sub JournaliserModification(ByVal nomTable As String, _
                                   ByVal cleEnreg As String, _
                                   ByVal action As String, _
                                   ByVal nomChamp As String, _
                                   ByVal ancienne As Variant, _
                                   ByVal nouvelle As Variant)
    On Error GoTo Echec
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblJournalModifications", dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    rs!DateHeure = Now()
    rs!IdentifiantWindows = Environ$("USERNAME")
    rs!NomTable = nomTable
    rs!CleEnregistrement = cleEnreg
    rs!Action = action
    rs!NomChamp = nomChamp
    rs!AncienneValeur = ValeurEnTexte(ancienne)
    rs!NouvelleValeur = ValeurEnTexte(nouvelle)
    rs.Update

Nettoyage:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
Echec:
    Resume Nettoyage    ' la défaillance de l'audit ne bloque pas l'action métier
End Sub

Private Function ValeurEnTexte(ByVal v As Variant) As Variant
    ' Convertit une valeur en texte pour le journal, en conservant Null.
    If IsNull(v) Then
        ValeurEnTexte = Null
    Else
        ValeurEnTexte = CStr(v)
    End If
End Function
```

On ajoute ensuite une procédure qui, appelée depuis `BeforeUpdate`, parcourt les contrôles liés du formulaire et journalise chaque champ modifié.

```vba
Public Sub AuditerFormulaire(ByRef frm As Form, _
                             ByVal nomTable As String, _
                             ByVal nomChampCle As String)
    On Error Resume Next

    Dim estCreation As Boolean
    estCreation = frm.NewRecord

    Dim strCle As String
    strCle = Nz(frm.Controls(nomChampCle).Value, "(nouveau)")

    Dim strAction As String
    strAction = IIf(estCreation, "Création", "Modification")

    Dim ctl As Control
    For Each ctl In frm.Controls
        If EstControleLie(ctl) Then
            If estCreation Then
                If Not IsNull(ctl.Value) Then
                    JournaliserModification nomTable, strCle, strAction, _
                                            ctl.ControlSource, Null, ctl.Value
                End If
            Else
                ' Chr$(0) sert de sentinelle pour comparer aussi les Null
                If Nz(ctl.Value, Chr$(0)) <> Nz(ctl.OldValue, Chr$(0)) Then
                    JournaliserModification nomTable, strCle, strAction, _
                                            ctl.ControlSource, ctl.OldValue, ctl.Value
                End If
            End If
        End If
    Next ctl
End Sub

Private Function EstControleLie(ByRef ctl As Control) As Boolean
    ' Vrai si le contrôle est lié à un champ (et non à une expression "=...").
    Dim src As String
    On Error Resume Next
    src = ctl.ControlSource
    On Error GoTo 0
    EstControleLie = (Len(src) > 0) And (Left$(src, 1) <> "=")
End Function
```

Pour des journaux plus riches, on remplacera avantageusement `Environ$("USERNAME")` par l'identité chargée dans `modSecurite` (voir [section 20.6](06-gestion-droits-utilisateurs.md)), par exemple l'identifiant et le nom complet de l'utilisateur de la session.

## Brancher l'audit sur un formulaire

L'intégration dans un formulaire donné se réduit alors à quelques lignes dans son module de classe.

```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Journalise les changements juste avant l'enregistrement
    AuditerFormulaire Me, "tblClients", "ClientID"
End Sub

Private Sub Form_Delete(Cancel As Integer)
    ' Journalise la suppression à partir des valeurs encore lisibles
    JournaliserModification "tblClients", Nz(Me!ClientID, ""), _
                            "Suppression", "(enregistrement)", _
                            Me!Nom & "", Null
End Sub
```

Un point mérite attention pour les créations : la clé primaire d'un nouvel enregistrement n'est pas toujours connue au moment de `BeforeUpdate`. Lorsqu'il s'agit d'un NuméroAuto, sa valeur définitive est en pratique souvent déjà attribuée, mais en cas de doute on peut compléter la trace dans l'événement `AfterInsert`, où la clé est garantie.

## Que faut-il auditer ?

La tentation de tout journaliser doit être combattue : elle engendre un volume considérable, dégrade les performances par les écritures supplémentaires et noie l'information utile dans le bruit. Mieux vaut cibler les tables et les opérations sensibles. Les suppressions, irréversibles, sont les premières à auditer, de même que les données financières, les données personnelles et les modifications de droits. À l'inverse, des tables de référence rarement touchées ne justifient pas nécessairement un audit fin.

Une précaution de confidentialité s'impose également : n'inscrivez jamais en clair, dans les colonnes de valeurs du journal, des données elles-mêmes sensibles comme des mots de passe. Cela rejoint les considérations de la [section 20.8](08-chiffrement-donnees-sensibles.md).

## Intégrité, croissance et limites

L'intégrité du journal appelle la même lucidité que la gestion des droits. Le journal s'exécute dans l'application et n'est protégé que par les couches qui l'entourent. Sans `.accde`, le code d'audit peut être désactivé ; sans chiffrement du back-end ni verrouillage de l'interface, un utilisateur technique pourrait atteindre les tables de journal et les altérer. En Access natif, on ne peut pas empêcher absolument cette falsification par quelqu'un qui dispose d'un accès au fichier. Pour une garantie d'inviolabilité plus forte, on journalise hors d'Access : dans une base SQL Server au moyen de déclencheurs côté serveur (voir [chapitre 23](../23-migration-interoperabilite/README.md)), ou dans le journal d'événements de Windows. L'audit Access reste néanmoins parfaitement adapté à la traçabilité courante d'une application déployée dans un environnement raisonnablement maîtrisé.

La croissance des tables de journal doit enfin être anticipée. Un audit actif génère beaucoup de lignes, et une base Access est plafonnée à 2 Go (voir [section 18.9](../18-optimisation-performance/09-limites-taille-2go.md)). Prévoyez une procédure d'archivage ou de purge périodique des traces anciennes, afin de maîtriser la taille du fichier dans la durée.

## Ce qu'il faut retenir

L'audit complète la gestion des droits en consignant les actions réellement effectuées. Deux tables, alimentées uniquement en insertion et hébergées dans le back-end, suffisent à journaliser les connexions et les modifications. Les changements de données se capturent dans l'événement `BeforeUpdate` grâce à la comparaison de `Value` et `OldValue`, et dans l'événement `Delete` pour les suppressions, le tout factorisé dans un module réutilisable robuste aux valeurs arbitraires. Outil de détection et non de prévention, l'audit n'est pleinement fiable que combiné aux autres protections du chapitre, et atteint sa forme la plus inviolable lorsqu'il est délégué à un serveur. Ciblé sur les opérations sensibles et accompagné d'une politique d'archivage, il offre à l'application une traçabilité solide.

---


⏭️ [20.8. Chiffrement des données sensibles dans les tables](/20-securite-protection/08-chiffrement-donnees-sensibles.md)
