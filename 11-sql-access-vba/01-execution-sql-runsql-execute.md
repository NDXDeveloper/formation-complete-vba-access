🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 11.1. Exécution de SQL par `DoCmd.RunSQL` et `CurrentDb.Execute`

Avant d'écrire la moindre requête sophistiquée, il faut savoir l'**exécuter** correctement. Or, comme le laissait entendre l'introduction du chapitre, les deux méthodes principales pour lancer du SQL action depuis VBA — **`DoCmd.RunSQL`** et **`CurrentDb.Execute`** — ne sont **pas équivalentes**. Leurs différences de comportement, notamment en matière de messages d'avertissement, de détection d'erreurs et de performance, ont des conséquences bien réelles sur la fiabilité de votre code.

Cette section compare ces deux voies, met en lumière leurs pièges respectifs, et dégage une recommandation claire. Un point commun, d'emblée : aucune des deux ne sert à exécuter un `SELECT` qui renvoie des lignes — cela relève de l'ouverture d'un recordset, traitée à la section suivante.

---

## Deux méthodes pour les requêtes action

Les deux méthodes étudiées ici exécutent des **requêtes action** — `INSERT`, `UPDATE`, `DELETE` — ainsi que des instructions **DDL** (`CREATE`, `ALTER`, `DROP`, section 11.4). En revanche, **ni l'une ni l'autre ne peut exécuter un `SELECT`** destiné à renvoyer des enregistrements.

C'est une erreur de débutant fréquente : tenter de lancer un `SELECT` via `RunSQL` ou `Execute` déclenche une erreur (avec `CurrentDb.Execute`, l'erreur 3065 *« Impossible d'exécuter une requête sélection »* est explicite). Pour lire des données, on ouvre un **recordset** sur le résultat d'un `SELECT` — c'est l'objet de la section 11.2. Gardez cette distinction présente à l'esprit : *exécuter une action* et *lire des données* sont deux opérations distinctes, aux outils différents.

---

## `DoCmd.RunSQL` : la voie de l'automatisation Access

`DoCmd.RunSQL` appartient à l'objet `DoCmd`, le grand objet d'automatisation d'Access (chapitre 5). Sa syntaxe est simple :

```vba
DoCmd.RunSQL "UPDATE Clients SET Actif = True WHERE Region = 'Nord';"
```

Un second argument facultatif, `UseTransaction` (à `True` par défaut), enveloppe l'action dans une transaction.

La caractéristique déterminante de `RunSQL` est qu'elle **passe par la couche d'interface** d'Access. Cela a une conséquence immédiate et souvent gênante : par défaut, Access **affiche des boîtes de dialogue de confirmation** du type *« Vous êtes sur le point de mettre à jour N ligne(s)… »*. Dans un traitement automatisé, ces interruptions sont indésirables.

### Le piège de `SetWarnings`

Pour supprimer ces messages, on encadre l'appel par `DoCmd.SetWarnings False` … `True` :

```vba
DoCmd.SetWarnings False
DoCmd.RunSQL "UPDATE Clients SET Actif = True WHERE Region = 'Nord';"
DoCmd.SetWarnings True
```

Mais cette technique recèle un **piège redoutable** : si une erreur survient **entre** la désactivation et la réactivation, les avertissements **restent désactivés pour toute la session Access**. L'utilisateur ne verra alors plus aucune confirmation, y compris lors d'opérations manuelles ultérieures — un comportement dangereux et déroutant.

La parade impérative consiste à **réactiver les avertissements dans une structure de gestion d'erreurs**, garantissant leur rétablissement quoi qu'il arrive :

```vba
Public Sub MettreAJourClients()
    On Error GoTo Gestion

    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE Clients SET Actif = True WHERE Region = 'Nord';"

Nettoyage:
    DoCmd.SetWarnings True      ' rétabli dans TOUS les cas
    Exit Sub

Gestion:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
    Resume Nettoyage
End Sub
```

La gestion de `SetWarnings` est également abordée au chapitre 5.7. Retenez ici que toute désactivation doit s'accompagner d'une réactivation garantie.

---

## `CurrentDb.Execute` : la voie DAO

`CurrentDb.Execute` est une méthode de l'objet `Database` de DAO (chapitre 9). Sa syntaxe ressemble à celle de `RunSQL`, mais son comportement diffère sur des points essentiels :

```vba
CurrentDb.Execute "UPDATE Clients SET Actif = True WHERE Region = 'Nord';", dbFailOnError
```

À la différence de `RunSQL`, `Execute` **ne passe pas par l'interface** : **aucune boîte de dialogue**, aucune confirmation, donc nul besoin de `SetWarnings`. Elle est aussi **plus rapide**, car elle s'adresse directement au moteur.

### `dbFailOnError` : l'option indispensable

L'option **`dbFailOnError`** est le point le plus important de cette section. Sans elle, `Execute` adopte un comportement périlleux : si certaines lignes ne peuvent être traitées (violation de clé, règle de validation, conflit de verrou), elles sont **ignorées silencieusement**, et **aucune erreur n'est levée**. Votre code croit l'opération réussie alors qu'une partie des données n'a pas été modifiée.

Avec `dbFailOnError`, le comportement devient fiable : en cas d'échec, les modifications de l'instruction sont **annulées** et une **erreur interceptable** est déclenchée, que votre gestionnaire `On Error` capte. C'est la condition d'une détection d'erreurs digne de ce nom — **utilisez `dbFailOnError` systématiquement**.

### Récupérer le nombre de lignes affectées — et le piège de `CurrentDb`

`Execute` permet de connaître le **nombre de lignes affectées**, via la propriété `RecordsAffected` de l'objet `Database`. Mais un piège subtil guette ici, lié à un comportement de `CurrentDb` : **chaque appel à `CurrentDb` renvoie un *nouvel* objet `Database`**. Écrire `CurrentDb.Execute "…"` puis `CurrentDb.RecordsAffected` interroge donc **deux objets différents**, et le second renvoie un résultat erroné.

La solution consiste à **conserver une référence** à un unique objet `Database` :

```vba
Public Sub MettreAJourEtCompter()
    Dim db As DAO.Database
    Set db = CurrentDb                  ' une seule instance, réutilisée

    db.Execute "UPDATE Clients SET Actif = True WHERE Region = 'Nord';", dbFailOnError
    Debug.Print db.RecordsAffected & " ligne(s) modifiée(s)."

    Set db = Nothing
End Sub
```

Ce comportement de `CurrentDb`, et sa comparaison avec `DBEngine(0)(0)`, sont détaillés au chapitre 4.6. Retenez la règle : **dès que l'on a besoin de `RecordsAffected` (ou de plusieurs opérations cohérentes), on conserve la référence `Database`** plutôt que de rappeler `CurrentDb`.

---

## Les autres options d'`Execute`

Au-delà de `dbFailOnError`, `Execute` accepte d'autres constantes combinables. La plus utile en pratique est **`dbSeeChanges`**, **obligatoire** lorsqu'on modifie par code des **tables liées SQL Server** comportant une colonne d'identité : son absence provoque l'erreur 3197. On l'emploie alors en combinaison : `dbFailOnError + dbSeeChanges`. D'autres options existent (`dbInconsistent`, `dbDenyWrite`…), plus rarement nécessaires. La connexion aux bases externes ayant été traitée au chapitre 10.9, retenez simplement que `dbSeeChanges` est le compagnon habituel des tables liées au serveur.

---

## Tableau comparatif

| Critère | `DoCmd.RunSQL` | `CurrentDb.Execute` |
|---|---|---|
| Type de requête | Action / DDL (**pas** de `SELECT`) | Action / DDL (**pas** de `SELECT`) |
| Boîtes d'avertissement | **Oui** (à supprimer via `SetWarnings`) | **Non** |
| Détection des erreurs | Via `On Error` (comportement UI) | **Fiable** avec `dbFailOnError` |
| Lignes affectées | Difficile à obtenir | Via `RecordsAffected` (référence à conserver) |
| Passe par la couche interface | Oui | Non |
| Performance | Moindre | Meilleure |
| Usage recommandé | Compatibilité, comportement d'interface | **Défaut recommandé** |

---

## Et ADO ?

Pour mémoire, ADO offre une troisième voie, étudiée au chapitre 10 : **`Connection.Execute`** (sections 10.3 et 10.5). Une différence notable la distingue toutefois des deux méthodes ci-dessus : l'`Execute` d'ADO **peut renvoyer un recordset** pour un `SELECT`, là où `DoCmd.RunSQL` et `DAO.Database.Execute` en sont incapables. Dans un contexte ADO, c'est donc cette méthode que l'on emploiera ; dans un contexte DAO/ACE, `CurrentDb.Execute` reste la référence.

---

## Transactions

Les deux méthodes s'inscrivent dans une logique transactionnelle. `RunSQL` enveloppe par défaut son action dans une transaction (paramètre `UseTransaction`), et `Execute` participe aux transactions DAO si on l'encadre par `BeginTrans` / `CommitTrans`. Le regroupement de plusieurs opérations en une unité atomique — pour qu'elles réussissent ou échouent ensemble — fait l'objet du chapitre 14.2.

---

## Recommandation

Pour exécuter du SQL action dans un contexte Access/ACE, la recommandation est nette : **`CurrentDb.Execute` avec l'option `dbFailOnError`**, en conservant la référence `Database` lorsqu'on a besoin du nombre de lignes affectées. Cette approche combine absence de dialogues parasites, détection fiable des erreurs et meilleures performances. C'est aussi la conclusion du chapitre 18.3, qui compare les deux méthodes sous l'angle de la performance.

On réserve `DoCmd.RunSQL` aux cas où l'on souhaite spécifiquement son comportement d'interface, ou pour des raisons de compatibilité avec du code existant (par exemple issu de la conversion de macros) — en veillant alors toujours à encadrer `SetWarnings` par une gestion d'erreurs. Enfin, pour exécuter un `SELECT` et lire des données, on n'emploie aucune de ces deux méthodes : on ouvre un recordset.

---

## En résumé

`DoCmd.RunSQL` et `CurrentDb.Execute` exécutent toutes deux des requêtes **action** et du **DDL**, mais **jamais de `SELECT`**. `RunSQL` passe par l'interface d'Access, d'où des boîtes de confirmation à neutraliser via `SetWarnings` — avec l'impératif de les rétablir dans un gestionnaire d'erreurs. `CurrentDb.Execute`, plus directe et plus rapide, n'affiche aucun dialogue ; son option **`dbFailOnError`** est indispensable à une détection fiable des erreurs, faute de quoi les lignes en échec sont ignorées en silence. Pour obtenir `RecordsAffected`, il faut conserver la référence `Database`, car chaque appel à `CurrentDb` crée un nouvel objet. La recommandation par défaut est donc `CurrentDb.Execute` + `dbFailOnError`, `RunSQL` étant réservé aux besoins d'interface ou de compatibilité.

Maintenant que nous savons exécuter une action, l'autre grande famille d'opérations attend : la **lecture** de données. Pour cela, on n'exécute pas la requête « à blanc » — on ouvre un recordset sur le résultat d'un `SELECT`, afin d'en parcourir les enregistrements.


⏭️ [11.2. Requêtes SELECT et ouverture de Recordset](/11-sql-access-vba/02-select-recordset.md)
