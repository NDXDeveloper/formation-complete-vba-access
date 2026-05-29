🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.8 Suppression d'enregistrements (`Delete`)

## Introduction

Après l'ajout (section [9.6](/09-dao-data-access-objects/06-ajout-enregistrements.md)) et la modification (section [9.7](/09-dao-data-access-objects/07-modification-enregistrements.md)), la suppression d'enregistrements complète les opérations d'écriture sur un `Recordset`. Elle s'appuie sur une seule méthode, `Delete`, dont le fonctionnement diffère nettement des deux précédentes sur deux points : elle agit en une seule étape, sans `Update`, et elle laisse le curseur dans un état particulier qu'il faut impérativement gérer. Ces deux spécificités sont la source des erreurs les plus fréquentes, et constituent le cœur de cette section.

## Une opération en une seule étape

À la différence de `AddNew` et d'`Edit`, qui passent par un tampon validé par `Update`, la méthode `Delete` agit **immédiatement** : l'appel supprime l'enregistrement courant sur-le-champ, sans qu'aucun `Update` ne soit nécessaire — ni même possible.

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb()
Set rs = db.OpenRecordset("tblClients", dbOpenDynaset)

' rs est positionné sur l'enregistrement à supprimer
rs.Delete                     ' suppression immédiate, sans Update

rs.Close
Set rs = Nothing
Set db = Nothing
```

## Le prérequis : être positionné sur l'enregistrement

Comme `Edit`, la méthode `Delete` opère sur l'**enregistrement courant**. Il faut donc s'y être positionné au préalable — à l'ouverture, par déplacement (section [9.4](/09-dao-data-access-objects/04-navigation-recordset.md)) ou par recherche (section [9.9](/09-dao-data-access-objects/09-recherche-recordset.md)). Appeler `Delete` alors qu'aucun enregistrement n'est courant — jeu vide, position `EOF`/`BOF` — déclenche l'erreur **3021** (« Aucun enregistrement courant »).

## Le piège majeur : l'enregistrement courant après `Delete`

C'est le point le plus important de cette section. Après un `Delete`, l'enregistrement est bien supprimé, mais **le curseur reste positionné dessus** — sur un enregistrement désormais invalide. Le curseur ne se déplace pas automatiquement vers l'enregistrement suivant.

La conséquence est directe : toute tentative de **lire les champs** de cet enregistrement courant supprimé déclenche l'erreur **3167** (« L'enregistrement est supprimé »). Il faut donc, immédiatement après `Delete`, déplacer le curseur — typiquement par `MoveNext` — avant toute autre opération sur l'enregistrement courant.

```vba
rs.Delete
' rs!NomClient                ' ERREUR 3167 : l'enregistrement est supprimé
rs.MoveNext                   ' on quitte l'enregistrement supprimé
```

## Supprimer plusieurs enregistrements : la boucle correcte

Le cas le plus fréquent est la suppression d'enregistrements au fil d'un parcours. Le motif correct combine la boucle de navigation et la méthode `Delete`, en plaçant un `MoveNext` qui s'exécute **dans tous les cas** :

```vba
Do While Not rs.EOF
    If rs!Obsolete = True Then
        rs.Delete
    End If
    rs.MoveNext               ' avance, qu'il y ait eu suppression ou non
Loop
```

Ce motif fonctionne car le `MoveNext` final fait avancer le curseur après chaque tour, qu'une suppression ait eu lieu ou non ; après un `Delete`, ce `MoveNext` amène simplement à l'enregistrement suivant. L'erreur classique consiste à supprimer sans déplacer ensuite, ou à tenter de lire l'enregistrement courant juste après l'avoir supprimé — deux pièges que ce schéma évite.

## Lire avant de supprimer

Il découle de ce qui précède une bonne pratique : si l'on a besoin des **données de l'enregistrement** que l'on s'apprête à supprimer — par exemple pour les journaliser dans une table d'audit (section [20.7](/20-securite-protection/07-audit-acces-modifications.md)) — on les lit **avant** d'appeler `Delete`, car elles deviennent ensuite inaccessibles.

```vba
Dim ancienNom As String
ancienNom = rs!NomClient & ""   ' lecture AVANT la suppression
rs.Delete
rs.MoveNext
```

## Intégrité référentielle

Une suppression peut échouer pour des raisons d'**intégrité référentielle**. Si l'enregistrement à supprimer possède des enregistrements liés dans des tables enfants, et que l'intégrité référentielle est appliquée **sans** suppression en cascade, `Delete` échoue avec l'erreur **3200** (l'enregistrement ne peut être supprimé car des enregistrements liés existent). Si, en revanche, la **suppression en cascade** est activée sur la relation, les enregistrements liés sont supprimés en même temps. Ce comportement dépend donc de la définition des relations, traitée à la section [12.6](/12-querydefs-tabledefs/06-collection-relations.md), et des règles d'intégrité abordées à la section [14.6](/14-transactions/06-integrite-referentielle.md).

## Annuler une suppression : la transaction

`Delete` étant immédiate, il n'existe pas de `CancelUpdate` pour la suppression : il n'y a pas de tampon à annuler. La seule manière de rendre une suppression réversible est de l'inscrire dans une **transaction**, qu'un `Rollback` permettra d'annuler en cas de besoin. Cette approche est aussi la garantie qu'un ensemble de suppressions réussisse ou échoue d'un bloc. Les transactions DAO sont traitées au chapitre [14](/14-transactions/README.md).

## Concurrence et verrouillage

En environnement multi-utilisateur, l'enregistrement visé peut être verrouillé par un autre utilisateur, ce qui fait échouer la suppression (erreurs des familles **3260** / **3197**). La gestion de ces situations relève des stratégies de verrouillage exposées au chapitre [15](/15-multi-utilisateurs/README.md).

## Type de `Recordset` requis

Comme les autres opérations d'écriture, `Delete` exige un `Recordset` **modifiable** : types **Table** ou **Dynaset** uniquement. Les types **Snapshot** et **Forward-only**, en lecture seule, ne le permettent pas (section [9.3](/09-dao-data-access-objects/03-recordset-types.md)).

## `Delete` ou `DELETE` SQL ?

Comme pour l'ajout et la modification, le choix dépend du contexte. Pour supprimer **un grand nombre d'enregistrements** répondant à un même critère, une requête `DELETE … WHERE` exécutée via `db.Execute` est nettement plus rapide qu'une boucle de `Delete`, car elle opère en une seule passe côté moteur. La boucle de `Delete` se justifie lorsque la décision de supprimer dépend d'une **logique fine, propre à chaque enregistrement**, difficile à exprimer en SQL. Les requêtes action `DELETE` sont traitées à la section [11.3](/11-sql-access-vba/03-requetes-action-insert-update-delete.md), et la comparaison `Execute`/`RunSQL` à la section [18.3](/18-optimisation-performance/03-execute-vs-runsql.md).

## Et pour un formulaire ?

Pour supprimer l'enregistrement affiché dans un formulaire ouvert, on agit en général sur sa copie de travail, le `RecordsetClone`, en pensant à resynchroniser l'affichage ; ce mécanisme est détaillé à la section [9.12](/09-dao-data-access-objects/12-recordsetclone.md).

## Pièges courants

Les erreurs récurrentes sont les suivantes. Lire l'enregistrement courant juste après l'avoir supprimé, ce qui déclenche l'erreur 3167. Oublier de déplacer le curseur après `Delete` dans une boucle. Appeler `Delete` sans enregistrement courant (erreur 3021). S'attendre à pouvoir annuler une suppression sans transaction. Tenter une suppression sur un `Recordset` non modifiable. Et négliger les contraintes d'intégrité référentielle, qui font échouer la suppression d'un enregistrement encore référencé.

## Points clés à retenir

`Delete` supprime l'enregistrement courant **immédiatement**, sans `Update`. Il faut s'être positionné sur cet enregistrement au préalable, sous peine d'erreur 3021. Surtout, après la suppression, le curseur reste sur l'enregistrement supprimé : il faut le déplacer (`MoveNext`) avant toute lecture, sans quoi survient l'erreur 3167. Dans une boucle, le motif sûr consiste à supprimer puis à avancer systématiquement. Si l'on a besoin des données supprimées, on les lit avant l'appel. Pour rendre une suppression réversible, on l'encadre dans une transaction. Enfin, pour de gros volumes, une requête `DELETE` est préférable à une boucle, et l'opération exige toujours un `Recordset` modifiable.

⏭️ [9.9. Recherche dans un Recordset (FindFirst, FindNext, Seek)](/09-dao-data-access-objects/09-recherche-recordset.md)
