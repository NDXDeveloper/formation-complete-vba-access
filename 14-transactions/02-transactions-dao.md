🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.2. Transactions DAO (BeginTrans, CommitTrans, Rollback)

La [section 14.1](01-concept-transaction.md) a posé le principe « tout ou rien ». Il s'agit maintenant de le mettre en œuvre concrètement avec DAO, la technologie d'accès aux données native d'Access. DAO expose trois méthodes — `BeginTrans`, `CommitTrans` et `Rollback` — qui matérialisent exactement le cycle de vie décrit précédemment. Encore faut-il savoir **où** elles se situent dans le modèle objet, **quelle est leur portée** et **comment les associer correctement** aux écritures pour obtenir une atomicité réelle.

## Où vivent les transactions DAO : l'objet Workspace

Contrairement à ce que l'on pourrait attendre, les méthodes transactionnelles ne sont pas portées par l'objet `Database`, mais par l'objet **`Workspace`** (espace de travail). C'est un point fondamental, car il détermine la portée de la transaction.

Un `Workspace` représente une session de travail avec le moteur de base de données. Il contient une collection `Databases` regroupant toutes les bases ouvertes dans cette session. Les trois méthodes transactionnelles s'appliquent au niveau du workspace :

- `Workspace.BeginTrans` — ouvre une transaction ;
- `Workspace.CommitTrans` — valide la transaction ;
- `Workspace.Rollback` — annule la transaction.

Il n'existe pas de méthode `BeginTrans` au niveau de `DBEngine` ni de `Database` : tout passe obligatoirement par un objet `Workspace`.

## L'espace de travail par défaut

Lorsqu'une application Access s'exécute, le moteur ouvre automatiquement un espace de travail par défaut, accessible par :

```vba
Dim ws As DAO.Workspace
Set ws = DBEngine.Workspaces(0)
```

C'est cet espace de travail par défaut qui héberge la base de données courante. Dans l'immense majorité des cas, c'est sur lui que l'on ouvre les transactions. Il n'est pas nécessaire — ni recommandé en usage courant — de créer ses propres workspaces ; on travaille avec celui qui existe déjà.

## La portée d'une transaction DAO

Voici le point décisif à intégrer : **une transaction DAO s'applique au workspace tout entier, et donc à toutes les bases de données ouvertes dans ce workspace**. Toute écriture effectuée, par n'importe quel moyen, sur une base appartenant à ce workspace pendant que la transaction est ouverte fait partie de cette transaction.

Cette portée a deux conséquences pratiques :

- toutes les écritures réalisées via la base courante après un `BeginTrans` sur l'espace par défaut sont rattachées à la transaction, qu'elles passent par `Execute` ou par un `Recordset` ;
- on ne peut pas isoler simplement deux transactions indépendantes simultanées dans le même workspace ; pour cela, il faudrait travailler dans des workspaces distincts, ce qui relève d'un usage avancé peu fréquent.

## CurrentDb et les transactions

Puisque la transaction est portée par le workspace et non par la base, comment articuler tout cela avec `CurrentDb`, la fonction habituellement utilisée pour obtenir une référence à la base courante ?

`CurrentDb()` retourne une référence à la base de données ouverte dans l'espace de travail par défaut. Une transaction ouverte sur `DBEngine.Workspaces(0)` couvre donc bien les opérations effectuées sur cette base. En revanche, il faut connaître une particularité : **chaque appel à `CurrentDb()` crée un nouvel objet `Database`**. Multiplier les appels est à la fois coûteux et source de confusion.

La bonne pratique consiste donc à **capturer une seule référence et à la réutiliser** pour toutes les opérations de la transaction :

```vba
Dim ws As DAO.Workspace
Dim db As DAO.Database

Set ws = DBEngine.Workspaces(0)
Set db = CurrentDb              ' une seule fois, réutilisée ensuite
```

On retient la règle simple : la transaction s'ouvre sur le **workspace**, les écritures se font sur une **référence de base unique** issue de ce même workspace.

## Le rôle décisif de l'option dbFailOnError

C'est sans doute le piège le plus important — et le plus méconnu — des transactions DAO. Lorsque l'on exécute une requête action avec la méthode `Execute` d'un objet `Database`, le comportement par défaut est **silencieux face aux erreurs** : si certains enregistrements ne peuvent être traités (violation d'une règle de validation, conflit de clé, enregistrement verrouillé…), `Execute` traite ce qu'il peut, **ignore le reste sans déclencher d'erreur**, et poursuit.

La conséquence est dramatique pour une transaction : si aucune erreur n'est levée, le gestionnaire d'erreur ne se déclenche pas, aucun `Rollback` n'est appelé, et la transaction est validée — avec des données partielles ou incohérentes. L'atomicité est perdue, en silence.

La solution tient en un mot-clé : **`dbFailOnError`**. Passée en option à `Execute`, cette constante demande au moteur de **lever une erreur d'exécution** dès qu'un enregistrement ne peut être traité, et d'annuler les modifications de cette requête. C'est cette erreur qui, interceptée, déclenchera le `Rollback` de la transaction.

```vba
db.Execute "UPDATE Stock SET Quantite = Quantite - 1 " & _
           "WHERE RefProduit = 'ABC';", dbFailOnError
```

**Toute requête action exécutée dans une transaction doit impérativement comporter l'option `dbFailOnError`.** Sans elle, la transaction ne protège de rien.

Pour cette raison, on privilégie `db.Execute ..., dbFailOnError` plutôt que `DoCmd.RunSQL` dans le code transactionnel : `RunSQL` n'offre pas la même remontée fiable des erreurs et déclenche des boîtes de dialogue d'avertissement. La comparaison détaillée des deux approches est traitée à la [section 18.3](/18-optimisation-performance/03-execute-vs-runsql.md).

## Transactions et Recordsets

Les écritures effectuées via un `Recordset` (`AddNew`/`Update`, `Edit`/`Update`, `Delete`) participent elles aussi à la transaction, dès lors que le recordset est ouvert à partir d'une base appartenant au workspace sur lequel la transaction a été ouverte :

```vba
Dim rs As DAO.Recordset
Set rs = db.OpenRecordset("LignesCommande", dbOpenDynaset)
With rs
    .AddNew
    !NumCommande = lngNumCommande
    !RefProduit = "ABC"
    !Quantite = 3
    .Update                 ' écriture rattachée à la transaction du workspace
End With
rs.Close
```

Si un `Rollback` intervient ensuite, l'enregistrement ajouté par ce recordset sera annulé au même titre que les écritures réalisées par `Execute`.

## Structure robuste type

En réunissant ces éléments — workspace, référence de base unique, `dbFailOnError`, et gestion d'erreur déclenchant le `Rollback` — on obtient le patron standard d'une procédure transactionnelle DAO. C'est cette structure qu'il convient de reproduire systématiquement.

```vba
Public Sub EnregistrerCommande()
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim transOuverte As Boolean

    On Error GoTo Gestion_Erreur

    Set ws = DBEngine.Workspaces(0)
    Set db = CurrentDb

    ws.BeginTrans                       ' --- Début de la transaction ---
    transOuverte = True

    ' Écriture de l'en-tête
    db.Execute "INSERT INTO Commandes (DateCommande, IdClient) " & _
               "VALUES (Date(), 42);", dbFailOnError

    ' Écriture des lignes (chacune avec dbFailOnError)
    db.Execute "INSERT INTO LignesCommande (NumCommande, RefProduit, Quantite) " & _
               "VALUES (1001, 'ABC', 3);", dbFailOnError
    db.Execute "INSERT INTO LignesCommande (NumCommande, RefProduit, Quantite) " & _
               "VALUES (1001, 'XYZ', 1);", dbFailOnError

    ws.CommitTrans                      ' --- Tout a réussi : validation ---
    transOuverte = False

Nettoyage:
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

Gestion_Erreur:
    If transOuverte Then ws.Rollback    ' --- Une erreur : annulation totale ---
    MsgBox "Enregistrement annulé : " & Err.Description, vbExclamation
    Resume Nettoyage
End Sub
```

Plusieurs détails de cette structure méritent l'attention :

- l'indicateur `transOuverte` permet de savoir, dans le gestionnaire d'erreur, s'il existe réellement une transaction à annuler. C'est indispensable, car une erreur peut survenir **avant** le `BeginTrans` (par exemple à l'ouverture de la base) ; appeler `Rollback` alors qu'aucune transaction n'est ouverte provoquerait une erreur ;
- le `CommitTrans` n'intervient **qu'à la toute fin**, une fois toutes les écritures réussies ;
- l'étiquette `Nettoyage` centralise la libération des objets et constitue le point de sortie unique, atteint aussi bien après un succès qu'après une annulation.

## Conclure proprement et libérer les objets

Une transaction doit toujours se conclure par un `CommitTrans` ou un `Rollback`. Si une transaction reste ouverte et que le workspace est libéré ou que la session se termine, **les écritures non validées sont automatiquement annulées**. Ce filet de sécurité ne dispense pas d'écrire un code propre : une transaction laissée ouverte par négligence maintient des verrous et peut bloquer d'autres utilisateurs.

La libération des objets (`Set rs = Nothing`, `Set db = Nothing`, `Set ws = Nothing`) suit les règles générales vues au [chapitre 9](/09-dao-data-access-objects/11-cloture-liberation-dao.md). On veillera notamment à fermer et libérer les recordsets avant de valider ou d'annuler.

## Transactions imbriquées

DAO autorise l'**imbrication** de transactions : un `BeginTrans` peut être appelé alors qu'une transaction est déjà ouverte, créant un niveau interne. Chaque niveau possède alors son propre dénouement, et l'annulation d'un niveau externe annule également les niveaux internes qu'il contient. Ce mécanisme, ainsi que ses limites, fait l'objet de la [section 14.4](04-transactions-imbriquees.md).

## Erreurs et pièges fréquents

- **Oublier `dbFailOnError`** — l'erreur la plus grave. Les requêtes action échouent en silence, la transaction valide des données incohérentes. À considérer comme non négociable.
- **Appeler `CommitTrans` ou `Rollback` sans `BeginTrans` préalable** — provoque une erreur d'exécution. D'où l'usage de l'indicateur `transOuverte` pour protéger le gestionnaire d'erreur.
- **Multiplier les appels à `CurrentDb`** — chaque appel crée un nouvel objet. Capturer une référence unique et la réutiliser.
- **Valider trop tôt** — un `CommitTrans` placé au milieu de la séquence fige les premières écritures et fait perdre l'atomicité du reste.
- **Laisser une transaction ouverte** — verrous maintenus, blocages potentiels en multi-utilisateur. Toute transaction doit se conclure.
- **Tables liées via ODBC** — une transaction ouverte sur le workspace local Jet/ACE n'encadre pas de la même manière les écritures vers un serveur distant : le comportement transactionnel relève alors en grande partie du serveur. Ce cas est abordé sous l'angle de l'isolation à la [section 14.7](07-niveaux-isolation.md).

## En résumé

En DAO, les transactions sont portées par l'objet `Workspace`, le plus souvent l'espace par défaut `DBEngine.Workspaces(0)`. On ouvre la transaction avec `BeginTrans`, on la conclut par `CommitTrans` (validation) ou `Rollback` (annulation). La portée étant celle du workspace, toutes les écritures sur la base courante — par `Execute` comme par `Recordset` — y sont rattachées. Deux exigences conditionnent la fiabilité de l'ensemble : exécuter toute requête action avec l'option **`dbFailOnError`** afin que les erreurs soient réellement levées, et **coupler la transaction à une gestion d'erreur** qui déclenche le `Rollback` au moment voulu, protégée par un indicateur d'état.

La section suivante traite du même mécanisme avec l'autre technologie d'accès aux données : les [transactions en ADO](03-transactions-ado.md).

⏭️ [14.3. Transactions ADO](/14-transactions/03-transactions-ado.md)
