🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.3. Techniques de test des modules de données (DAO/ADO)

Le code d'accès aux données est le plus difficile à tester, parce qu'il dépend d'un état **mutable et partagé** : la base. Un test qui lit dépend de ce que contient la base à l'instant ; un test qui écrit peut endommager des données réelles ou laisser des résidus. Pour qu'un test ait de la valeur, il doit être **répétable**, **isolé** et **non destructeur**. Cette section présente les techniques qui rendent ces propriétés atteignables en VBA, que l'on travaille en DAO ou en ADO.

## Pourquoi le code de données résiste au test

Trois caractéristiques compliquent la tâche. D'abord, le résultat dépend de l'état de la base, qui varie dans le temps et d'un utilisateur à l'autre — « lancer et regarder » ne prouve rien de fiable. Ensuite, les opérations d'écriture ont des effets de bord persistants : elles modifient des données qu'il faudra remettre en état. Enfin, le comportement dépend de l'environnement : base connectée, tables liées, réseau. Tester proprement suppose donc de **maîtriser ces conditions**.

## Rendre le code testable : la conception d'abord

Le plus grand levier n'est pas une astuce de test mais une décision de conception : du code mal structuré est intestable, quelle que soit la technique.

### Séparer accès aux données, logique métier et interface

Du code d'accès aux données enfoui dans des événements de formulaire ne peut pas être éprouvé isolément. Regroupé dans une couche d'accès aux données — des fonctions ciblées comme `ClientParId` ou `EnregistrerCommande` —, il devient testable indépendamment de l'interface. C'est l'un des bénéfices de l'architecture en couches et du patron Repository.

> ℹ️ Voir le patron Repository ([16.6](../16-poo-vba-access/06-patron-repository.md)) et l'architecture en couches DAL/BLL/UI ([16.8](../16-poo-vba-access/08-architecture-en-couches.md)).

### Injecter la base plutôt que dépendre de `CurrentDb`

Une fonction qui appelle systématiquement `CurrentDb` est liée à la base courante. En lui passant la base (ou la connexion) **en paramètre**, on peut diriger le même code vers une base de test.

```vba
' Difficile à tester : dépendance implicite
Function ClientParId(ByVal id As Long) As String
    ClientParId = Nz(DLookup("Nom", "Clients", "ClientID = " & id), "")
End Function

' Testable : la base est injectée
Function ClientParId(ByVal db As DAO.Database, ByVal id As Long) As String
    ClientParId = Nz(DLookup("Nom", "Clients", "ClientID = " & id), "")
End Function
```

Cette injection de dépendance est la clé qui permettra, plus loin, d'isoler les transactions et de substituer un faux dépôt.

## Maîtriser les données : le jeu de test

Un test a besoin d'un **état de départ connu**. On suit le schéma *Arranger–Agir–Vérifier* (Arrange-Act-Assert) : préparer des données connues, exécuter le code, comparer le résultat à l'attendu. La préparation insère les enregistrements nécessaires avant le test, et le nettoyage les supprime après — de sorte qu'aucun test ne dépende d'un autre. La mise en place de ces jeux de données et d'environnements isolés fait l'objet de la [section 19.6](06-simulation-donnees-test.md).

## Isoler avec une transaction : tester sans laisser de trace

La technique la plus élégante pour les tests d'écriture consiste à encadrer le test par une **transaction** que l'on annule à la fin : l'écriture a bien lieu, on en vérifie l'effet, puis le `Rollback` ramène la base à son état initial — sans aucun résidu.

```vba
Sub TestAjoutClient()
    Dim ws As DAO.Workspace, db As DAO.Database
    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb                       ' même Workspace que la transaction

    ws.BeginTrans
    On Error GoTo Nettoyage

    ' Agir : exécuter le code de données sous test
    db.Execute "INSERT INTO Clients (Nom) VALUES ('TEST_UNIT')", dbFailOnError

    ' Vérifier : la transaction voit ses propres écritures non validées
    Debug.Assert DCount("*", "Clients", "Nom = 'TEST_UNIT'") = 1

Nettoyage:
    ws.Rollback                              ' annule tout : état initial restauré
End Sub
```

Pour que cela fonctionne, le code testé doit opérer dans le **même Workspace** que la transaction — d'où l'intérêt d'injecter une base issue de ce Workspace (ici `CurrentDb`, qui appartient au Workspace par défaut).

> ℹ️ Les transactions DAO sont traitées au [chapitre 14.2](../14-transactions/02-transactions-dao.md), `DBEngine` et `Workspace` au [chapitre 4.5](../04-modele-objet-access/05-dbengine-workspace-database.md). En ADO, l'isolation s'obtient de même via `Connection.BeginTrans` ([14.3](../14-transactions/03-transactions-ado.md)).

### Limites de l'isolation par transaction

Cette approche a ses bornes. Les instructions **DDL** (`CREATE`, `ALTER`, `DROP`) ne sont pas transactionnelles sous le moteur ACE : elles ne sont pas annulées par un `Rollback`. L'isolation reste donc cantonnée aux données (DML), pas au schéma. Et tout doit se dérouler dans le même Workspace ou la même connexion ; vers une source ODBC, le comportement transactionnel obéit à ses propres règles.

## Vérifier les résultats

La vérification consiste à **relire** la base après l'action et à comparer à l'attendu : un `DCount` pour confirmer un nombre d'enregistrements, un `DLookup` ou un Recordset pour contrôler des valeurs précises. On compare les comptes avant et après pour valider une insertion ou une suppression. Le verdict s'exprime par des assertions (`Debug.Assert`, voir la [section 19.2](02-debug-print-assert.md)) ou, mieux, à l'aide du framework de test de la [section 19.4](04-tests-unitaires-framework.md).

### Capturer les valeurs générées

Un écueil fréquent : vouloir vérifier un identifiant **NuméroAuto** précis. Ces valeurs changent à chaque exécution (et le compactage les réinitialise, voir la [section 18.7](../18-optimisation-performance/07-compactage-automatique.md)) ; on ne doit donc jamais affirmer une valeur fixe. On **capture** plutôt l'identifiant généré pour s'en servir ensuite. En DAO, après `Update`, le signet `LastModified` permet de se repositionner sur l'enregistrement créé.

```vba
rs.AddNew
rs!Nom = "TEST_UNIT"
rs.Update
rs.Bookmark = rs.LastModified      ' se placer sur l'enregistrement ajouté
nouveauId = rs!ClientID            ' récupérer le NuméroAuto réellement généré
```

(Une requête `SELECT @@IDENTITY` exécutée sur la même connexion juste après une insertion rend le même service.)

> ℹ️ L'ajout d'enregistrements est détaillé au [chapitre 9.6](../09-dao-data-access-objects/06-ajout-enregistrements.md) ; la subtilité de `RecordCount` (et de `MoveLast`) à la [section 18.2](../18-optimisation-performance/02-optimisation-recordsets.md).

## Vers le vrai test unitaire : s'affranchir de la base

Tester avec une transaction reste un **test d'intégration** : on touche une vraie base. Pour des **tests unitaires** rapides et totalement isolés de la logique métier, on découple le code de la base au moyen d'une interface. On définit un contrat d'accès aux données (par exemple `IClientRepository`, via `Implements`), avec deux implémentations : l'une réelle, fondée sur DAO ou ADO, l'autre **factice**, en mémoire (un `Dictionary`). Le code métier dépend de l'interface, jamais du moteur. En test, on injecte le faux dépôt — aucun accès à la base n'a lieu.

```vba
' Le code métier dépend du contrat, pas de DAO :
'   Function CalculerRemise(repo As IClientRepository, id As Long) As Double
' En production : on lui passe ClientRepositoryDAO (lit la base).
' En test       : on lui passe ClientRepositoryFake (Dictionary en mémoire).
```

On sépare ainsi clairement les tests d'intégration (vraie base de test) des tests unitaires (faux dépôt, sans base).

> ℹ️ Les interfaces `Implements` sont traitées au [chapitre 16.5](../16-poo-vba-access/05-collections-interfaces.md), le patron Repository au [chapitre 16.6](../16-poo-vba-access/06-patron-repository.md).

## Pièges propres aux tests de données

Plusieurs travers guettent ces tests. La **fuite d'état** d'un test à l'autre (des résidus) fausse les exécutions suivantes : transactions ou nettoyage systématique l'évitent. La **dépendance d'ordre** (un test qui suppose qu'un autre a tourné avant) rend les résultats imprévisibles : chaque test doit être autonome. Les **NuméroAuto** imposent de capturer les identifiants plutôt que de les coder en dur. Enfin, exécuter les tests contre des **tables liées** passe par le réseau et touche une base partagée : on leur préfère une base de test locale, plus rapide et sans effet sur les autres. Les mêmes principes valent en DAO comme en ADO ; seuls les objets diffèrent.

> ℹ️ Le choix DAO / ADO est discuté au [chapitre 10.1](../10-ado-access/01-dao-vs-ado.md), les opérations CRUD en ADO au [chapitre 10.5](../10-ado-access/05-crud-ado.md). Le formatage du SQL (dates, décimales) à tester est traité au [chapitre 11.7](../11-sql-access-vba/07-localisation-formatage-sql.md).

## Points clés à retenir

- Le code de données est dur à tester car il dépend d'un état mutable et partagé ; un bon test est répétable, isolé et non destructeur.
- La conception prime : séparer l'accès aux données de la logique et de l'interface, et **injecter** la base en paramètre plutôt que dépendre de `CurrentDb`.
- Préparer un état connu (Arranger–Agir–Vérifier) avec mise en place et nettoyage, chaque test restant autonome.
- Encadrer les tests d'écriture par `BeginTrans`/`Rollback` les rend non destructeurs — en gardant à l'esprit que le DDL n'est pas annulé et que tout doit rester dans le même Workspace.
- Vérifier en relisant la base (`DCount`, `DLookup`, Recordset) ; ne jamais affirmer un NuméroAuto fixe, mais capturer l'identifiant généré (`LastModified`).
- Pour de vrais tests unitaires, isoler la logique de la base via une interface et un faux dépôt en mémoire, en distinguant tests d'intégration et tests unitaires.

---


⏭️ [19.4. Tests unitaires en VBA — framework léger maison](/19-debogage-tests/04-tests-unitaires-framework.md)
