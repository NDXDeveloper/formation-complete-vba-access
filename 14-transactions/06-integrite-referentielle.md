🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 14.6. Règles d'intégrité référentielle par code

Après la concurrence, ce chapitre revient à l'intégrité **structurelle** des données. L'**intégrité référentielle** est la règle qui garantit la cohérence des liens entre tables : elle empêche qu'un enregistrement « enfant » fasse référence à un parent inexistant. C'est l'un des piliers de la fiabilité d'une base relationnelle, et Access permet de la définir, de l'inspecter et de la supprimer entièrement par code.

> La gestion générale de la collection `Relations` est traitée à la [section 12.6](/12-querydefs-tabledefs/06-collection-relations.md). La présente section se concentre sur l'aspect **intégrité** : ce que ces règles garantissent, comment les mettre en place par VBA, et comment elles se manifestent à l'exécution.

## Qu'est-ce que l'intégrité référentielle ?

Concrètement, appliquer l'intégrité référentielle entre une table primaire et une table liée interdit trois opérations dangereuses : ajouter dans la table liée un enregistrement sans correspondance dans la clé primaire de la table primaire, modifier dans la table primaire une valeur qui rendrait orphelins des enregistrements liés, et supprimer dans la table primaire des enregistrements ayant des enregistrements liés correspondants.

Autrement dit, l'intégrité référentielle garantit qu'à tout instant, chaque clé étrangère pointe vers une clé primaire réellement existante. Pas d'enregistrement orphelin, pas de lien rompu. Sans cette règle, une suppression maladroite d'un client laisserait derrière elle des commandes pointant vers un client disparu — incohérence classique et difficile à rattraper.

## Comment Access applique l'intégrité référentielle

Dans Access, l'intégrité référentielle est portée par les **relations** entre tables, représentées en DAO par l'objet `Relation` et regroupées dans la collection `Relations` de la base. Créer une relation avec intégrité référentielle revient à inscrire dans la base une règle que le moteur ACE fera respecter automatiquement à chaque écriture, quelle que soit la voie utilisée (formulaire, requête, code).

C'est un point important : l'intégrité référentielle définie au niveau des relations s'applique **globalement**, indépendamment du code. Une fois la règle posée, on n'a plus à la vérifier manuellement dans chaque procédure : le moteur s'en charge.

## Les attributs d'une relation

Le comportement d'une relation est gouverné par sa propriété `Attributes`, une combinaison de constantes de l'énumération `RelationAttributeEnum`. Les principales sont : `dbRelationUnique` (valeur 1, relation un-à-un), `dbRelationDontEnforce` (valeur 2, relation non appliquée — pas d'intégrité référentielle), `dbRelationUpdateCascade` (valeur 256, mises à jour en cascade) et `dbRelationDeleteCascade` (valeur 4096, suppressions en cascade).

Le point essentiel à retenir : **l'intégrité référentielle est appliquée par défaut**. C'est l'attribut `dbRelationDontEnforce` qui, lorsqu'on le positionne, désactive l'intégrité. Une relation créée sans cet attribut fait donc respecter l'intégrité référentielle. Pour créer un simple lien d'affichage sans contrôle (par exemple vers une table liée Excel qui ne supporte pas l'intégrité), on ajoute explicitement `dbRelationDontEnforce`.

## Créer une relation avec intégrité référentielle

La création se fait avec la méthode `CreateRelation`, dont la signature est `CreateRelation(NomRelation, TablePrimaire, TableÉtrangère, Attributs)`, où `NomRelation` est le nom unique de la relation, `TablePrimaire` la table référencée contenant la clé, et `TableÉtrangère` la table référençante contenant la clé étrangère.

Trois prérequis doivent être satisfaits, faute de quoi l'ajout échoue :

- le champ référencé dans la table primaire doit être **clé primaire ou indexé sans doublon** (voir [section 12.5](/12-querydefs-tabledefs/05-creer-modifier-tables.md)) ;
- les champs appariés doivent avoir des **types compatibles** ;
- les **données existantes** doivent déjà respecter l'intégrité : on ne peut pas imposer une règle référentielle à des données qui la violent déjà (présence d'orphelins).

La séquence de création suit un ordre strict. Avant de pouvoir utiliser la méthode `Append` sur l'objet `Relation`, il faut y avoir ajouté les objets `Field` définissant la correspondance entre clés primaire et étrangère. Une fois l'objet ajouté à la collection, ses propriétés ne peuvent plus être modifiées.

```vba
Public Sub CreerRelationCommandesLignes()
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field

    On Error GoTo Gestion_Erreur
    Set db = CurrentDb

    ' 1. Créer la relation : nom unique, table primaire (clé), table étrangère (liée)
    Set rel = db.CreateRelation("FK_LignesCommande_Commandes", _
                                "Commandes", "LignesCommande")

    ' 2. Définir les attributs : intégrité (par défaut) + cascade MAJ et suppression
    '    ATTENTION : on combine les constantes avec Or, JAMAIS avec And
    rel.Attributes = dbRelationUpdateCascade Or dbRelationDeleteCascade

    ' 3. Apparier les champs : champ de la table primaire, puis nom du champ étranger
    Set fld = rel.CreateField("NumCommande")   ' champ clé dans Commandes
    fld.ForeignName = "NumCommande"            ' champ étranger dans LignesCommande
    rel.Fields.Append fld

    ' 4. Rendre la relation permanente
    db.Relations.Append rel
    db.Relations.Refresh

    MsgBox "Relation créée avec intégrité référentielle.", vbInformation

Nettoyage:
    Set fld = Nothing
    Set rel = Nothing
    Set db = Nothing
    Exit Sub

Gestion_Erreur:
    MsgBox "Création impossible : " & Err.Number & " — " & Err.Description, vbExclamation
    Resume Nettoyage
End Sub
```

L'objet `Field` créé représente le côté gauche de la relation (la table primaire), le côté droit étant désigné par la propriété `ForeignName`. Pour une clé composée de plusieurs champs, on répète l'étape 3 pour chaque paire de champs.

Notons aussi que chaque relation doit porter un **nom unique** dans la base : ce nom est la clé d'identification de la relation, et tenter d'en créer deux avec le même nom provoque une erreur.

## Le piège And / Or

Une erreur extrêmement répandue — au point d'apparaître dans d'anciens exemples de documentation — consiste à combiner les attributs avec l'opérateur `And` au lieu de `Or` : écrire `dbRelationUpdateCascade And dbRelationDeleteCascade` ne produit pas les deux cascades, mais la valeur **0**. En effet, ces constantes sont des bits distincts (256 et 4096) ; leur `And` binaire ne partage aucun bit commun et vaut donc zéro, ce qui désactive toute cascade sans la moindre erreur visible.

La règle est simple et vaut pour tous les indicateurs binaires : **on utilise `Or` pour combiner (poser) des attributs, et `And` pour tester la présence d'un attribut**. Confondre les deux est l'une des causes les plus fréquentes de relations « qui ne fonctionnent pas comme prévu ».

## Les cascades : mise à jour et suppression

Les deux attributs de cascade assouplissent l'intégrité référentielle sans la rompre : avec `dbRelationDeleteCascade` ou `dbRelationUpdateCascade`, le moteur autorise les modifications et suppressions, mais répercute ces changements sur les enregistrements liés afin que les règles restent respectées.

Concrètement :

- **`dbRelationUpdateCascade`** — si la clé primaire d'un parent change, les clés étrangères des enregistrements liés sont automatiquement mises à jour pour suivre.
- **`dbRelationDeleteCascade`** — si un enregistrement parent est supprimé, tous ses enregistrements liés sont automatiquement supprimés.

La suppression en cascade est puissante mais doit être employée avec prudence : supprimer un client effacera silencieusement toutes ses commandes et, si la cascade se propage, toutes les lignes de ces commandes. C'est un comportement voulu dans certains modèles, dangereux dans d'autres.

## Inspecter les relations existantes

On peut parcourir la collection `Relations` pour documenter ou auditer le modèle. Le test de chaque attribut se fait par un `And` binaire — illustration concrète de la règle énoncée plus haut.

```vba
Public Sub InspecterRelations()
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim integrite As Boolean

    Set db = CurrentDb
    For Each rel In db.Relations
        ' Intégrité appliquée SAUF si l'attribut dbRelationDontEnforce est présent
        integrite = ((rel.Attributes And dbRelationDontEnforce) = 0)

        Debug.Print rel.Name & " : " & rel.Table & " -> " & rel.ForeignTable
        Debug.Print "  Intégrité : " & IIf(integrite, "Oui", "Non") & _
            " | Cascade MAJ : " & _
            IIf((rel.Attributes And dbRelationUpdateCascade) <> 0, "Oui", "Non") & _
            " | Cascade Suppr. : " & _
            IIf((rel.Attributes And dbRelationDeleteCascade) <> 0, "Oui", "Non")

        For Each fld In rel.Fields
            Debug.Print "    " & rel.Table & "." & fld.Name & _
                        "  =  " & rel.ForeignTable & "." & fld.ForeignName
        Next fld
    Next rel

    Set db = Nothing
End Sub
```

## Supprimer une relation

La suppression d'une relation se fait par son nom :

```vba
db.Relations.Delete "FK_LignesCommande_Commandes"
db.Relations.Refresh
```

Puisqu'une relation ne peut pas être modifiée après son ajout, **modifier une règle d'intégrité revient à supprimer la relation puis à la recréer** avec les nouveaux attributs.

## Le comportement à l'exécution : les erreurs d'intégrité

Lorsque l'intégrité référentielle est appliquée, une tentative d'écriture qui la violerait déclenche une erreur d'exécution interceptable. Les deux cas typiques sont :

- l'**ajout ou la modification d'un enfant orphelin** (clé étrangère sans parent correspondant) : erreur **3201**, indiquant qu'un enregistrement associé est requis dans la table primaire ;
- la **suppression ou la modification d'un parent ayant des enfants**, sans cascade : erreur **3200**, indiquant que l'enregistrement ne peut être supprimé ou modifié car la table contient des enregistrements associés.

Lorsque la cascade correspondante est activée, ces opérations ne lèvent plus d'erreur : le moteur les autorise et propage automatiquement le changement. Côté code, on traitera donc ces numéros d'erreur dans le gestionnaire pour afficher un message clair à l'utilisateur (« Impossible de supprimer ce client : des commandes lui sont rattachées »).

## Intégrité référentielle et transactions

L'intégrité référentielle et les transactions se complètent naturellement. Une violation d'intégrité se traduit par une erreur ; si l'écriture fautive se produit à l'intérieur d'une transaction, le gestionnaire d'erreur déclenche un `Rollback` qui annule l'ensemble du bloc. On obtient ainsi la garantie qu'un traitement multi-tables aboutit **soit à un état complet et cohérent, soit à rien** : l'atomicité de la transaction et les règles d'intégrité travaillent de concert.

De même, une suppression en cascade exécutée dans une transaction est atomique : le parent et l'ensemble des enfants supprimés en cascade sont validés ou annulés d'un seul tenant.

## Une alternative : la déclaration en DDL

L'intégrité référentielle peut aussi être déclarée en SQL DDL, au moyen d'une contrainte `FOREIGN KEY ... REFERENCES` lors de la création ou de la modification d'une table. Cette approche, complémentaire de l'objet `Relation`, est abordée à la [section 11.4](/11-sql-access-vba/04-requetes-ddl.md). Le résultat est le même type de règle, posée par une autre voie.

## Points de vigilance

- **L'intégrité est appliquée par défaut.** C'est `dbRelationDontEnforce` qui la désactive, et non l'inverse.
- **`Or` pour combiner, `And` pour tester.** Combiner des attributs avec `And` donne 0 et désactive silencieusement les cascades.
- **Données propres exigées.** Impossible d'imposer l'intégrité à des données contenant déjà des orphelins ; nettoyer d'abord.
- **Clé primaire ou index unique requis** sur le champ référencé de la table primaire.
- **Pas de modification après `Append`.** Pour changer une relation, la supprimer et la recréer.
- **Cascade de suppression = effacements en chaîne.** Vérifier que ce comportement est réellement souhaité avant de l'activer.
- **Tables liées ODBC.** L'intégrité d'un serveur externe est gérée par ce serveur ; une relation locale non appliquée (`dbRelationDontEnforce`) sert alors uniquement à l'affichage.

## En résumé

L'intégrité référentielle garantit que chaque clé étrangère pointe vers une clé primaire existante, en interdisant les enfants orphelins, les modifications de clé primaire qui rompraient un lien, et les suppressions de parents ayant des enfants. En DAO, elle se définit via l'objet `Relation` et `CreateRelation`, en appariant les champs (`CreateField` + `ForeignName`) puis en ajoutant la relation à la collection ; **l'intégrité est appliquée par défaut**, sauf attribut `dbRelationDontEnforce`. Les cascades (`dbRelationUpdateCascade`, `dbRelationDeleteCascade`) propagent les changements aux enregistrements liés — en se combinant avec `Or`, jamais avec `And`. À l'exécution, les violations se traduisent par les erreurs 3201 (orphelin) et 3200 (parent avec enfants), interceptables et idéalement encadrées par une transaction pour préserver l'atomicité de l'ensemble.

La dernière section du chapitre aborde les limites du moteur natif face à un serveur : les [niveaux d'isolation](07-niveaux-isolation.md).

⏭️ [14.7. Niveaux d'isolation (ODBC / serveur lié) — limites du moteur ACE natif](/14-transactions/07-niveaux-isolation.md)
