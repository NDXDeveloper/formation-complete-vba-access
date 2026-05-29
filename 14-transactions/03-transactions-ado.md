🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.3. Transactions ADO

ADO (ActiveX Data Objects) propose son propre mécanisme transactionnel, conceptuellement identique à celui de DAO mais articulé autour d'un objet différent et avec quelques particularités qu'il faut connaître pour ne pas commettre d'erreurs. Le principe « tout ou rien » exposé à la [section 14.1](01-concept-transaction.md) reste évidemment le même ; seules changent la mécanique et la terminologie. Cette section suit le même plan que la précédente afin de faciliter la comparaison.

## Où vivent les transactions ADO : l'objet Connection

Là où DAO porte les transactions au niveau du `Workspace`, **ADO les porte au niveau de l'objet `Connection`**. C'est la connexion qui ouvre, valide ou annule la transaction, et c'est par elle que doivent transiter toutes les écritures concernées.

Cette différence d'objet hôte entraîne une différence de portée importante, sur laquelle nous reviendrons : une transaction ADO ne couvre que les opérations effectuées **à travers la connexion qui l'a ouverte**.

## Les trois méthodes : BeginTrans, CommitTrans, RollbackTrans

L'objet `Connection` expose trois méthodes transactionnelles :

- `Connection.BeginTrans` — ouvre une transaction ;
- `Connection.CommitTrans` — valide la transaction ;
- `Connection.RollbackTrans` — annule la transaction.

Deux points de vigilance immédiats par rapport à DAO :

- la méthode d'annulation s'appelle **`RollbackTrans`** en ADO, et non `Rollback` comme en DAO. C'est une source d'erreur classique lorsqu'on passe d'une technologie à l'autre : `cn.Rollback` n'existe pas et provoquera une erreur ;
- `BeginTrans` **retourne une valeur** : un entier `Long` indiquant le niveau de la nouvelle transaction (1 pour le premier niveau, 2 pour une transaction imbriquée, etc.). DAO, lui, ne retourne rien. Cette valeur de retour est facultative à exploiter, mais elle existe.

```vba
Dim niveau As Long
niveau = cn.BeginTrans      ' niveau = 1 pour une transaction de premier niveau
```

## La portée d'une transaction ADO

C'est ici que se situe la principale différence pratique avec DAO. Une transaction ADO est **liée à une connexion précise** : seules les écritures réalisées à travers **cette même connexion** font partie de la transaction. Une opération exécutée via une autre connexion — même ouverte sur la même base — n'y participe pas et sera validée indépendamment.

La règle pratique est donc stricte : **toutes les écritures de l'unité de travail doivent passer par le même objet `Connection`** que celui sur lequel `BeginTrans` a été appelé. Cela vaut pour les appels à `Execute` comme pour les `Recordset`, dont la propriété `ActiveConnection` doit pointer vers cette connexion.

## Obtenir la connexion

Deux situations se présentent en pratique.

Pour travailler sur la base courante, le plus simple est d'utiliser la connexion ADO de l'application, exposée par `CurrentProject.Connection` :

```vba
Dim cn As ADODB.Connection
Set cn = CurrentProject.Connection
```

Cette connexion appartient à Access : on **ne la ferme jamais** (`cn.Close` est à proscrire ici, car cela fermerait la connexion propre de l'application). On se contente de libérer la variable locale (`Set cn = Nothing`) en fin de procédure.

Pour travailler sur une base externe ou dans une architecture ADO dédiée, on ouvre au contraire une connexion que l'on a créée soi-même, avec une chaîne de connexion appropriée (voir [section 10.3](/10-ado-access/03-objet-connection.md)). Dans ce cas, c'est à nous de la fermer et de la libérer une fois le travail terminé.

## Erreurs en ADO : pas de dbFailOnError

La section 14.2 a longuement insisté sur la nécessité de l'option `dbFailOnError` avec `Execute` en DAO, faute de quoi les erreurs passaient inaperçues. **Cette précaution n'a pas d'équivalent à prévoir en ADO** : par défaut, le fournisseur OLE DB **lève une erreur d'exécution** lorsqu'une opération échoue (violation de contrainte, conflit, etc.). L'erreur est interceptable normalement par `On Error`, et le détail est également disponible via la collection `Errors` de la connexion (voir [section 10.8](/10-ado-access/08-gestion-erreurs-ado.md)).

Autrement dit, le comportement par défaut d'ADO est, sur ce point, plus sûr que celui de DAO : il n'y a pas de risque d'échec silencieux du même ordre. Cela ne dispense évidemment **pas** d'une gestion d'erreur rigoureuse — c'est elle qui déclenchera le `RollbackTrans`.

## Transactions et Recordsets ADO

Un `Recordset` ADO participe à la transaction dès lors qu'il est ouvert sur la **même connexion** et avec un type de curseur et un mode de verrouillage permettant la mise à jour (par exemple `adOpenKeyset` associé à `adLockOptimistic`). Les notions de curseurs et de verrouillage sont détaillées à la [section 10.4](/10-ado-access/04-recordset-ado-curseurs.md).

```vba
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "LignesCommande", cn, adOpenKeyset, adLockOptimistic, adCmdTableDirect
With rs
    .AddNew
    !NumCommande = 1001
    !RefProduit = "ABC"
    !Quantite = 3
    .Update                 ' rattaché à la transaction de la connexion cn
End With
rs.Close
Set rs = Nothing
```

Le point essentiel est le deuxième argument de `Open` : c'est la connexion `cn` — celle qui porte la transaction — qui est passée comme connexion active.

## Structure robuste type

En réunissant ces éléments, on obtient le patron standard d'une procédure transactionnelle ADO. Il est volontairement calqué sur celui de la section 14.2, à deux différences près : l'objet hôte est la connexion, et la méthode d'annulation est `RollbackTrans`.

```vba
Public Sub EnregistrerCommande_ADO()
    Dim cn As ADODB.Connection
    Dim transOuverte As Boolean

    On Error GoTo Gestion_Erreur

    Set cn = CurrentProject.Connection

    cn.BeginTrans                       ' --- Début de la transaction ---
    transOuverte = True

    cn.Execute "INSERT INTO Commandes (NumCommande, IdClient) " & _
               "VALUES (1001, 42);"
    cn.Execute "INSERT INTO LignesCommande (NumCommande, RefProduit, Quantite) " & _
               "VALUES (1001, 'ABC', 3);"
    cn.Execute "INSERT INTO LignesCommande (NumCommande, RefProduit, Quantite) " & _
               "VALUES (1001, 'XYZ', 1);"

    cn.CommitTrans                      ' --- Tout a réussi : validation ---
    transOuverte = False

Nettoyage:
    Set cn = Nothing                    ' on libère, on ne ferme PAS CurrentProject.Connection
    Exit Sub

Gestion_Erreur:
    If transOuverte Then cn.RollbackTrans   ' --- Une erreur : annulation totale ---
    MsgBox "Enregistrement annulé : " & Err.Description, vbExclamation
    Resume Nettoyage
End Sub
```

La logique est identique à celle vue en DAO : l'indicateur `transOuverte` protège le gestionnaire d'erreur contre un `RollbackTrans` appelé sans transaction ouverte, la validation n'intervient qu'en toute fin, et l'étiquette `Nettoyage` offre un point de sortie unique.

## Validation/annulation avec maintien automatique

ADO offre une option absente de DAO : grâce à la propriété `Attributes` de la connexion, on peut activer le **maintien automatique de transaction**. En positionnant les drapeaux `adXactCommitRetaining` et/ou `adXactAbortRetaining`, une nouvelle transaction est automatiquement ouverte juste après chaque `CommitTrans` ou `RollbackTrans`. Par défaut, ce maintien est désactivé : après un dénouement, plus aucune transaction n'est ouverte tant qu'on ne rappelle pas `BeginTrans`. Ce comportement est utile dans certains traitements par lots, mais surprenant si on l'active sans le savoir ; mieux vaut donc le laisser inactif en l'absence de besoin précis.

## Dépendance au fournisseur OLE DB

Le support transactionnel — et en particulier celui des transactions imbriquées — **dépend du fournisseur OLE DB** utilisé. Le fournisseur ACE (`Microsoft.ACE.OLEDB`) prend en charge les transactions sur les bases Access. Lorsqu'on se connecte à une base externe (SQL Server, Oracle…), c'est le fournisseur correspondant qui gouverne le comportement transactionnel, conformément aux règles du serveur. La connexion à des bases externes est traitée à la [section 10.9](/10-ado-access/09-connexion-bases-externes.md), et la question de l'isolation à la [section 14.7](07-niveaux-isolation.md).

## Transactions imbriquées

ADO autorise l'imbrication des transactions, la valeur retournée par `BeginTrans` indiquant le niveau atteint. Le comportement (un `CommitTrans` ou un `RollbackTrans` agit sur le niveau courant) et les limites associées sont traités à la [section 14.4](04-transactions-imbriquees.md).

## DAO ou ADO pour les transactions ?

Sur le plan transactionnel, les deux technologies offrent le même service : encadrer un ensemble d'écritures sous une unité atomique. Le choix se fait donc en fonction de la technologie déjà retenue pour l'accès aux données, plutôt que sur les capacités transactionnelles elles-mêmes. Pour une application travaillant sur des tables Access natives, DAO reste souvent le choix le plus direct et le plus performant ; ADO s'impose davantage dans les architectures connectées à des sources externes ou déjà bâties autour de ce modèle. La comparaison générale des deux approches fait l'objet de la [section 10.1](/10-ado-access/01-dao-vs-ado.md).

## Erreurs et pièges fréquents

- **Écrire `Rollback` au lieu de `RollbackTrans`** — la méthode n'existe pas sur l'objet `Connection`. C'est l'erreur la plus fréquente au passage DAO → ADO.
- **Mélanger les connexions** — toute écriture doit passer par la connexion qui porte la transaction. Une opération via une autre connexion échappe à la transaction et sera validée indépendamment.
- **Oublier d'aligner `ActiveConnection`** — un recordset ouvert sur une autre connexion que celle de la transaction n'y participe pas.
- **Appeler `CommitTrans`/`RollbackTrans` sans transaction ouverte** — provoque une erreur. D'où l'indicateur `transOuverte`.
- **Fermer `CurrentProject.Connection`** — à ne jamais faire ; on ne ferme que les connexions que l'on a soi-même créées.
- **Activer involontairement le maintien automatique** (`adXactCommitRetaining`) — peut laisser une transaction ouverte de façon inattendue après un commit.
- **Laisser une transaction ouverte** — verrous maintenus et blocages possibles en multi-utilisateur (voir [chapitre 15](/15-multi-utilisateurs/README.md)).

## En résumé

En ADO, les transactions sont portées par l'objet `Connection` et manipulées avec `BeginTrans`, `CommitTrans` et **`RollbackTrans`** (attention à ce nom, différent du `Rollback` de DAO). Leur portée est celle de la connexion : toutes les écritures de l'unité de travail doivent passer par le même objet `Connection`, recordsets compris via leur `ActiveConnection`. Contrairement à DAO, aucune option de type `dbFailOnError` n'est nécessaire : ADO lève les erreurs par défaut. Le patron de mise en œuvre est, à ces nuances près, identique à celui de DAO : `BeginTrans`, écritures, `CommitTrans` final, et `RollbackTrans` déclenché par une gestion d'erreur protégée par un indicateur d'état. Le support exact, notamment de l'imbrication, dépend du fournisseur OLE DB employé.

La section suivante approfondit le cas particulier des [transactions imbriquées](04-transactions-imbriquees.md), commun à DAO et à ADO.

⏭️ [14.4. Transactions imbriquées](/14-transactions/04-transactions-imbriquees.md)
