🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 16.6. Patron Repository — abstraction de l'accès aux données

Plusieurs sections précédentes y ont renvoyé : une entité comme `clsClient` ne devrait pas accéder elle-même à la base ([section 16.3](03-methodes-encapsulation.md)), et une interface permet de dépendre d'une abstraction plutôt que d'une implémentation ([section 16.5](05-collections-interfaces.md)). Le **patron Repository** réunit ces idées pour résoudre un problème concret et récurrent : où et comment ranger le code d'accès aux données. Il propose de le **centraliser** derrière une **abstraction**, ce qui apporte séparation des responsabilités, testabilité, et — point capital pour Access — une grande facilité de migration.

## Le problème : l'accès aux données dispersé

Dans une application Access typique, le code d'accès aux données est éparpillé : des chaînes SQL réécrites dans plusieurs formulaires, du code DAO inséré dans des gestionnaires de clic, de la logique métier mêlée à des `OpenRecordset`. Les conséquences sont lourdes : une requête à corriger l'est à plusieurs endroits, une modification du schéma oblige à traquer le SQL partout, et **changer de moteur de données** — migrer de DAO vers SQL Server — impose de réécrire du code disséminé dans toute l'application. Sans compter qu'il devient quasiment impossible de tester la logique sans une vraie base.

## L'idée du Repository

Un **repository** est un objet qui s'interpose entre le domaine (les entités métier) et la source de données. Il expose une interface proche d'une collection — `ObtenirParId`, `ObtenirTous`, `Ajouter`, `Modifier`, `Supprimer` — opérant sur des **objets métier**, et masque entièrement la mécanique sous-jacente (SQL, DAO, ADO). Tout l'accès aux données d'un type d'entité est ainsi rassemblé dans **une seule classe**.

Du point de vue du code qui l'utilise, on ne manipule plus des recordsets et des requêtes, mais des objets : « donne-moi le client 42 », « enregistre ce client ». Le *comment* est l'affaire du repository.

## La structure : entité, interface, implémentations

Le patron met en jeu trois rôles, articulés autour des interfaces vues en 16.5 :

- l'**entité** — l'objet métier, `clsClient` ;
- l'**interface du repository** — `IClientRepository`, qui définit le contrat d'accès aux données ;
- une ou plusieurs **implémentations concrètes** — `clsClientRepositoryDAO`, et éventuellement `clsClientRepositoryADO` ou `clsClientRepositorySqlServer`, chacune réalisant le contrat.

Le code consommateur (logique métier, formulaires) dépend de l'**interface**, jamais d'une implémentation précise. C'est ce qui rend les implémentations interchangeables.

## L'interface du repository

L'interface définit le contrat, sans aucune implémentation (cf. 16.5). On peut y faire retourner une collection personnalisée typée (`clsClients`, section 16.5) plutôt qu'une `Collection` brute.

```vba
' --- Module de classe : IClientRepository (interface) ---
Option Explicit

Public Function ObtenirParId(ByVal id As Long) As clsClient
End Function

Public Function ObtenirTous() As clsClients      ' collection personnalisée
End Function

Public Sub Ajouter(ByVal client As clsClient)
End Sub

Public Sub Modifier(ByVal client As clsClient)
End Sub

Public Sub Supprimer(ByVal id As Long)
End Sub
```

## Une implémentation concrète (DAO)

L'implémentation DAO réalise le contrat et concentre **tout** le code d'accès aux données. C'est aussi là que se fait le **mapping** entre les lignes du recordset et les objets métier.

```vba
' --- Module de classe : clsClientRepositoryDAO ---
Option Explicit
Implements IClientRepository

Private Function IClientRepository_ObtenirParId(ByVal id As Long) As clsClient
    Dim rs As DAO.Recordset
    Dim c As clsClient
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Id, Nom, Email FROM Clients WHERE Id = " & id & ";", dbOpenSnapshot)
    If Not rs.EOF Then
        Set c = New clsClient
        c.Init rs!Id, Nz(rs!Nom, ""), Nz(rs!Email, "")   ' lignes -> objet
    End If
    rs.Close: Set rs = Nothing
    Set IClientRepository_ObtenirParId = c
End Function

Private Function IClientRepository_ObtenirTous() As clsClients
    Dim rs As DAO.Recordset, liste As clsClients, c As clsClient
    Set liste = New clsClients
    Set rs = CurrentDb.OpenRecordset("SELECT Id, Nom, Email FROM Clients;", dbOpenSnapshot)
    Do Until rs.EOF
        Set c = New clsClient
        c.Init rs!Id, Nz(rs!Nom, ""), Nz(rs!Email, "")
        liste.Add c
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing
    Set IClientRepository_ObtenirTous = liste
End Function

Private Sub IClientRepository_Ajouter(ByVal client As clsClient)
    ' En production : requête PARAMÉTRÉE (cf. 11.5), pas une concaténation.
    CurrentDb.Execute "INSERT INTO Clients (Nom, Email) VALUES ('" & _
        client.Nom & "', '" & client.Email & "');", dbFailOnError
End Sub

' IClientRepository_Modifier et IClientRepository_Supprimer suivent le même principe.
```

L'entité expose ici une méthode `Init` qui permet au repository de renseigner ses champs, y compris l'`Id` en lecture seule (cf. [section 16.2](02-proprietes-get-let-set.md)). Une remarque importante : pour la lisibilité, ces exemples concatènent le SQL ; **en production, on utilise des requêtes paramétrées** afin de prévenir les injections et les problèmes de format ([section 11.5](/11-sql-access-vba/05-sql-parametree.md)).

## Utiliser le repository

Le code consommateur dépend de l'**interface**. Il ne contient aucune ligne de SQL ni de DAO :

```vba
Dim repo As IClientRepository
Set repo = New clsClientRepositoryDAO        ' (idéalement fourni par une fabrique)

Dim c As clsClient
Set c = repo.ObtenirParId(42)
If Not c Is Nothing Then
    c.Email = "nouveau@exemple.fr"
    repo.Modifier c
End If
```

Le formulaire ou la couche métier orchestre des objets et des appels au repository — rien de plus. La mécanique d'accès est invisible à ce niveau.

## Changer d'implémentation : le bénéfice migration

C'est le grand intérêt du patron, et le fil rouge évoqué depuis le début de la formation. Puisque le code consommateur ne dépend que de `IClientRepository`, **changer de moteur de données ne touche qu'une chose : l'implémentation injectée**.

```vba
' Pour passer à SQL Server, seule cette ligne change :
Set repo = New clsClientRepositorySqlServer
' ... tout le code consommateur reste INCHANGÉ
```

Mieux : si la création de l'implémentation est centralisée dans une **fabrique** ([section 16.7](07-singleton-factory-observer.md)), même cette ligne ne figure qu'à un seul endroit. La migration d'une application Access vers un serveur ([chapitre 23](/23-migration-interoperabilite/README.md), notamment l'[architecture hybride](/23-migration-interoperabilite/04-architecture-hybride-access-sql-server.md)) s'en trouve considérablement simplifiée : on écrit une nouvelle implémentation du repository, et la logique métier n'est pas affectée.

## Injection de dépendance

Dans les exemples ci-dessus, le consommateur crée lui-même son repository (`Set repo = New ...`). Une approche plus découplée consiste à le lui **fournir** de l'extérieur — par exemple en le passant à une méthode `Init` d'un objet métier, ou via `OpenArgs` pour un formulaire. C'est l'**injection de dépendance** : l'objet ne décide plus quelle implémentation utiliser, il la reçoit. Cela réduit encore le couplage et prépare la testabilité. Cette idée structure l'architecture en couches de la [section 16.8](08-architecture-en-couches.md).

## Testabilité : un repository factice

L'abstraction ouvre une possibilité précieuse : écrire une implémentation **factice** du repository, en mémoire, qui réalise le même contrat sans base de données.

```vba
' --- clsClientRepositoryFake : Implements IClientRepository, stockage en Collection ---
' ObtenirParId, Ajouter, etc. travaillent sur une Collection interne.
```

En injectant ce `clsClientRepositoryFake` à la place de l'implémentation DAO, on peut **tester la logique métier sans toucher à la base** — des tests rapides, reproductibles et isolés. C'est l'un des grands apports du patron, mis à profit dans les tests unitaires ([section 19.4](/19-debogage-tests/04-tests-unitaires-framework.md)).

## Portée et granularité

On définit en général **un repository par type d'entité** : `ClientRepository`, `CommandeRepository`, etc. VBA n'ayant pas de génériques, un repository « générique » unique serait peu commode ; le découpage par entité reste plus clair. Un repository manipule des **entités complètes** (des objets `clsClient`), et non des fragments de données : c'est la couche métier qui demande « le client 42 », le repository qui sait le reconstituer.

## Où s'inscrit le Repository

Le repository constitue, à lui seul, l'abstraction de la **couche d'accès aux données** (DAL) : il s'intercale entre la logique métier (BLL) et la base. Il est donc une pièce maîtresse de l'architecture en couches, sujet de la [section 16.8](08-architecture-en-couches.md), qui montrera comment l'ensemble s'agence.

## Points de vigilance

- **Dépendre de l'interface, pas de l'implémentation** : le code consommateur manipule `IClientRepository`, jamais `clsClientRepositoryDAO` directement.
- **Tout le SQL/DAO dans le repository** : aucune requête ne doit subsister dans les formulaires ou la logique métier.
- **Le repository manipule des objets**, pas des recordsets : c'est lui qui fait le mapping lignes ↔ entités.
- **Requêtes paramétrées en production** : les concaténations des exemples sont à remplacer par des paramètres (11.5).
- **Un repository par entité** : éviter le repository « fourre-tout ».
- **Centraliser le choix d'implémentation** (fabrique, 16.7) et envisager l'injection pour la testabilité.

## En résumé

Le **patron Repository** centralise tout l'accès aux données d'une entité dans une classe dédiée, exposant une interface proche d'une collection (`ObtenirParId`, `ObtenirTous`, `Ajouter`, `Modifier`, `Supprimer`) qui opère sur des **objets métier** et masque le SQL, DAO ou ADO sous-jacent. Sa structure repose sur une **interface** (`IClientRepository`) dont dépend le code consommateur, et une ou plusieurs **implémentations concrètes** (DAO, ADO, SQL Server) qui réalisent le mapping entre lignes et entités. Les bénéfices sont la centralisation, la séparation des responsabilités, la **testabilité** (via un repository factice) et surtout une **migration facilitée** : changer de moteur ne touche que l'implémentation injectée, jamais la logique métier. C'est l'application directe de l'encapsulation et des interfaces, et la pièce d'accès aux données de l'architecture en couches.

La section suivante outille la création et la gestion de ces objets avec trois patrons classiques : [Singleton, Factory et Observer en VBA](07-singleton-factory-observer.md).

⏭️ [16.7. Patrons Singleton, Factory et Observer en VBA](/16-poo-vba-access/07-singleton-factory-observer.md)
