🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.7 Modification d'enregistrements (`Edit`/`Update`)

## Introduction

Après l'ajout d'enregistrements (section [9.6](/09-dao-data-access-objects/06-ajout-enregistrements.md)), voici son pendant : la modification d'enregistrements existants, à l'aide du couple `Edit`/`Update`. Les deux mécanismes se ressemblent — tous deux passent par un tampon d'édition validé par `Update` — mais une différence fondamentale les sépare : `AddNew` crée un enregistrement neuf, tandis qu'`Edit` agit sur l'enregistrement **déjà existant** sur lequel le curseur est positionné. Cette nuance commande tout le reste.

## Le prérequis : être positionné sur l'enregistrement

`Edit` modifie l'**enregistrement courant**. Avant d'éditer, il faut donc avoir amené le curseur sur l'enregistrement voulu : à l'ouverture (le `Recordset` est positionné sur le premier enregistrement), par déplacement (section [9.4](/09-dao-data-access-objects/04-navigation-recordset.md)), ou par recherche (`Find`, `Seek`, section [9.9](/09-dao-data-access-objects/09-recherche-recordset.md)). Appeler `Edit` alors qu'aucun enregistrement n'est courant — sur un jeu vide, ou en position `EOF`/`BOF` — déclenche l'erreur **3021** (« Aucun enregistrement courant »). C'est la principale différence à garder en tête par rapport à `AddNew`, qui n'a, lui, besoin d'aucun enregistrement courant.

## Le déroulé en trois temps

La modification suit trois étapes : entrer en mode édition par `Edit`, affecter les nouvelles valeurs, valider par `Update`.

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb()
Set rs = db.OpenRecordset("tblClients", dbOpenDynaset)

' rs est positionné sur le premier enregistrement à l'ouverture
rs.Edit                       ' 1. entre en mode modification
rs!Telephone = "0235123456"   ' 2. affecte les nouvelles valeurs
rs!DateMAJ = Now()
rs.Update                     ' 3. valide les changements

rs.Close
Set rs = Nothing
Set db = Nothing
```

## Ce que font `Edit` et `Update`

`Edit` copie l'enregistrement courant dans le tampon d'édition. Les affectations qui suivent modifient ce tampon, et non encore la base : tant que `Update` n'est pas appelé, les changements ne sont pas persistés. `Update` écrit alors le tampon dans la base et vérifie les contraintes (champs obligatoires, unicité, règles de validation).

Un comportement distingue ici `Edit` d'`AddNew` : après un `Edit`/`Update`, le curseur **reste positionné sur l'enregistrement modifié**. Contrairement à l'ajout, il n'est donc pas nécessaire de recourir à `LastModified` pour retrouver l'enregistrement : on est déjà dessus.

## Modifier en fonction de la valeur courante

Comme `Edit` travaille sur un enregistrement existant, on peut calculer la nouvelle valeur d'un champ à partir de son ancienne valeur — un schéma lecture-modification-écriture très courant :

```vba
rs.Edit
rs!Solde = rs!Solde + montant    ' la nouvelle valeur dérive de l'ancienne
rs.Update
```

On veillera, dans ce type d'opération, à la gestion des `Null` vue à la section [9.5](/09-dao-data-access-objects/05-lecture-modification-champs.md) : si le champ peut être `Null`, on encadre sa lecture par `Nz` (`rs!Solde = Nz(rs!Solde, 0) + montant`) pour éviter une erreur ou un résultat `Null`.

## Annuler une modification

Pour renoncer à une modification avant validation, on appelle `CancelUpdate`, qui abandonne le tampon : l'enregistrement reste inchangé. Le fonctionnement est identique à celui décrit pour l'ajout (section 9.6), de même que la règle d'hygiène associée : on ne laisse jamais un `Recordset` en attente entre `Edit` et `Update`. En particulier, **déplacer le curseur** (`MoveNext`, etc.) alors qu'une édition est en cours et non validée **perd silencieusement** les changements.

## Modifier plusieurs enregistrements : la boucle

Le cas le plus fréquent est la modification d'un ensemble d'enregistrements au fil d'un parcours. Le motif combine la boucle de navigation (section 9.4) et le triplet `Edit`/affectation/`Update`, en n'oubliant pas le `MoveNext` final :

```vba
Do While Not rs.EOF
    rs.Edit
    rs!Statut = "Traité"
    rs.Update
    rs.MoveNext
Loop
```

Chaque enregistrement est ici édité et validé individuellement. Lorsqu'il faut appliquer une **même modification** à de nombreux enregistrements, cette boucle n'est cependant pas la solution la plus performante.

## `Edit` ou `UPDATE` SQL ?

Comme pour l'ajout, le choix dépend du besoin. Pour appliquer une **modification uniforme** à un grand nombre d'enregistrements, une requête `UPDATE` exécutée via `db.Execute` est nettement plus rapide qu'une boucle d'`Edit`, car elle traite l'ensemble en une seule opération côté moteur. La boucle `Edit` s'impose en revanche lorsque la modification dépend d'une **logique propre à chaque enregistrement** — calcul conditionnel, transformation variable d'une ligne à l'autre, valeur dérivée de l'état courant de l'enregistrement. Les requêtes action `UPDATE` sont traitées à la section [11.3](/11-sql-access-vba/03-requetes-action-insert-update-delete.md), et la comparaison `Execute`/`RunSQL` à la section [18.3](/18-optimisation-performance/03-execute-vs-runsql.md).

## Contraintes et erreurs

Les contraintes vérifiées à l'`Update` sont les mêmes que pour l'ajout (section 9.6) : champ obligatoire non renseigné, violation d'unicité (erreur **3022**), valeur incompatible avec le type ou dépassant la taille, règle de validation non respectée. À cela s'ajoutent deux erreurs propres au mode édition : l'erreur **3020** si l'on appelle `Update` sans `Edit` ni `AddNew` préalable, et l'erreur **3021** si l'on appelle `Edit` sans enregistrement courant. Ces opérations gagnent à être encadrées par une gestion d'erreurs (chapitre [13](/13-gestion-erreurs/README.md)) et, lorsqu'un ensemble de modifications doit être atomique, par une transaction (chapitre [14](/14-transactions/README.md)).

## Modification concurrente et conflits

En environnement multi-utilisateur, un risque spécifique apparaît : entre le moment où l'on appelle `Edit` et celui où l'on appelle `Update`, un autre utilisateur peut avoir modifié — ou verrouillé — le même enregistrement. Par défaut, DAO applique un verrouillage *optimiste* sur les Dynaset, et le conflit se révèle au moment de l'`Update`, sous forme d'erreurs telles que **3197** (les données ont été modifiées par un autre utilisateur) ou **3260** (l'enregistrement est actuellement verrouillé). La détection et la résolution de ces conflits, qui s'appuient notamment sur la propriété `OldValue` (section 9.5), sont traitées à la section [14.5](/14-transactions/05-conflits-mise-a-jour.md), et les stratégies de verrouillage à la section [15.3](/15-multi-utilisateurs/03-strategies-verrouillage.md).

## Type de `Recordset` requis

Comme `AddNew`, la méthode `Edit` exige un `Recordset` **modifiable** : types **Table** ou **Dynaset** uniquement. Les types **Snapshot** et **Forward-only** étant en lecture seule, toute tentative d'édition sur ces types échoue (section [9.3](/09-dao-data-access-objects/03-recordset-types.md)).

## Et pour un formulaire ?

Pour modifier l'enregistrement affiché dans un formulaire ouvert, on ne pilote généralement pas directement la source du formulaire, mais sa copie de travail, le `RecordsetClone`, dont l'usage est exposé à la section [9.12](/09-dao-data-access-objects/12-recordsetclone.md).

## Pièges courants

Les erreurs récurrentes sont les suivantes. Appeler `Edit` sans s'être assuré qu'un enregistrement est bien courant, ce qui provoque l'erreur 3021. Croire que la modification est enregistrée dès l'affectation des champs : elle ne l'est qu'après `Update`. Quitter le mode édition par un déplacement plutôt que par `Update` ou `CancelUpdate`, ce qui perd les changements. Oublier le traitement des `Null` dans un calcul dérivé de la valeur courante. Et, en multi-utilisateur, ne pas anticiper les conflits qui se manifestent à l'`Update`.

## Points clés à retenir

La modification d'un enregistrement en DAO suit le triplet `Edit`, affectation des champs, `Update`. Contrairement à l'ajout, `Edit` opère sur l'enregistrement courant : il faut donc s'y être positionné au préalable, faute de quoi l'erreur 3021 survient. Après validation, le curseur reste sur l'enregistrement modifié. On annule une édition en cours par `CancelUpdate`, et l'on ne déplace jamais le curseur avant d'avoir validé ou annulé. Pour des modifications uniformes en masse, une requête `UPDATE` est préférable à une boucle d'`Edit`. Enfin, l'opération exige un `Recordset` modifiable et impose, en environnement partagé, d'anticiper les conflits de mise à jour révélés au moment de l'`Update`.

⏭️ [9.8. Suppression d'enregistrements (Delete)](/09-dao-data-access-objects/08-suppression-enregistrements.md)
