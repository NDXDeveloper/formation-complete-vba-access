🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.6 Ajout d'enregistrements (`AddNew`/`Update`)

## Introduction

La section précédente a montré comment lire et écrire la valeur d'un champ, et a renvoyé ici le déroulé complet de l'ajout d'un enregistrement. C'est l'objet de cette section : créer de nouveaux enregistrements dans une table à l'aide du couple `AddNew`/`Update`. Ce mécanisme procédural, ligne à ligne, complète l'approche ensembliste du SQL (`INSERT`) : il s'avère particulièrement pratique lorsqu'il faut appliquer une logique métier à chaque insertion ou récupérer immédiatement la valeur d'une clé générée automatiquement.

## Le déroulé en trois temps

L'ajout d'un enregistrement suit toujours trois étapes : on ouvre un nouvel enregistrement vierge par `AddNew`, on renseigne ses champs, puis on valide l'insertion par `Update`.

```vba
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb()
Set rs = db.OpenRecordset("tblClients", dbOpenDynaset)

rs.AddNew                     ' 1. ouvre un nouvel enregistrement vierge
rs!NomClient = "Dupont"       ' 2. renseigne les champs
rs!Ville = "Rouen"
rs.Update                     ' 3. valide l'insertion

rs.Close
Set rs = Nothing
Set db = Nothing
```

Tant que `Update` n'est pas appelé, aucun enregistrement n'est réellement ajouté à la table.

## Ce que fait `AddNew`

L'appel à `AddNew` crée un nouvel enregistrement dans une zone tampon (le *buffer* d'édition), sans encore l'écrire dans la base. Les champs de cet enregistrement sont initialisés à leur **valeur par défaut** (la propriété `DefaultValue` de la colonne, lorsqu'elle est définie) ou à `Null`. On n'est donc tenu de renseigner que les champs que l'on souhaite remplir ; les autres conservent leur valeur par défaut ou restent `Null`. L'enregistrement n'existe, à ce stade, que dans le tampon.

## Ce que fait `Update`

`Update` écrit le contenu du tampon dans la base : l'enregistrement est alors persisté. C'est aussi à ce moment que le moteur vérifie les contraintes (champs obligatoires, règles de validation, unicité des index) et déclenche, le cas échéant, les erreurs correspondantes.

## Le piège de la position après `Update`

Un point surprend systématiquement : après `Update`, le **curseur ne se positionne pas sur l'enregistrement nouvellement ajouté**. Il revient à l'endroit où il se trouvait avant l'appel à `AddNew`. L'enregistrement a bien été créé, mais l'enregistrement courant n'est pas lui.

Pour se positionner sur l'enregistrement que l'on vient d'ajouter, on utilise la propriété `LastModified`, qui renvoie le signet du dernier enregistrement ajouté ou modifié, et on l'affecte au signet courant :

```vba
rs.AddNew
rs!NomClient = "Martin"
rs.Update
rs.Bookmark = rs.LastModified   ' se positionne sur l'enregistrement ajouté
```

Cette technique est notamment indispensable pour relire une valeur générée par le moteur.

## Récupérer la clé auto-générée

Lorsqu'une table possède une clé primaire de type **NuméroAuto**, sa valeur est attribuée par le moteur, et non par le code. Pour la récupérer après l'ajout, la méthode fiable consiste à se repositionner sur le nouvel enregistrement via `LastModified`, puis à lire le champ :

```vba
rs.AddNew
rs!NomClient = "Martin"
rs.Update
rs.Bookmark = rs.LastModified
Debug.Print rs!IDClient          ' valeur du NuméroAuto généré
```

À titre indicatif, avec DAO sur le moteur ACE, la valeur d'un champ NuméroAuto est en réalité disponible dès l'appel à `AddNew`, dans le tampon, et reste lisible avant `Update`. Le passage par `LastModified` après `Update` demeure néanmoins l'approche la plus robuste et la plus universelle. Notons par ailleurs que le NuméroAuto n'est pas adapté à la génération de numéros séquentiels métier (numéros de facture, par exemple), car il peut comporter des trous ; pour ce besoin, on se reportera à la section [15.5](/15-multi-utilisateurs/05-numeros-sequentiels-multi-utilisateurs.md).

## Annuler un ajout : `CancelUpdate`

Si, après avoir appelé `AddNew`, on souhaite renoncer à l'insertion avant de l'avoir validée, on appelle `CancelUpdate`. Le tampon est alors abandonné et aucun enregistrement n'est ajouté.

```vba
rs.AddNew
rs!NomClient = "Provisoire"
' un contrôle métier échoue...
rs.CancelUpdate                  ' abandonne l'ajout
```

Il faut retenir une règle d'hygiène : on ne laisse jamais un `Recordset` en attente entre `AddNew` et `Update`. On le clôt toujours par l'un ou l'autre — `Update` pour valider, `CancelUpdate` pour annuler. À défaut, un déplacement du curseur (`MoveNext`, etc.) **avant** `Update` provoque la perte silencieuse de l'ajout en cours.

## Champs obligatoires, contraintes et erreurs

Plusieurs contraintes sont vérifiées au moment de l'`Update` et peuvent déclencher des erreurs. Un **champ obligatoire** (propriété `Required`, équivalent d'un `NOT NULL`) non renseigné fait échouer la validation. Une **violation d'unicité** — tenter d'insérer une clé primaire ou une valeur d'index unique déjà existante — déclenche l'erreur **3022**. Une valeur **incompatible avec le type** du champ, ou une chaîne dépassant la taille autorisée, provoquent également des erreurs. Enfin, comme vu à la section [9.5](/09-dao-data-access-objects/05-lecture-modification-champs.md), appeler `Update` sans `AddNew` ni `Edit` préalable lève l'erreur **3020**.

Ces opérations gagnent donc à être encadrées par une gestion d'erreurs (chapitre [13](/13-gestion-erreurs/README.md)). Et lorsque plusieurs ajouts doivent réussir ou échouer ensemble, on les place dans une transaction, de sorte qu'un échec annule l'ensemble (chapitre [14](/14-transactions/README.md)).

## Type de `Recordset` requis

`AddNew` exige un `Recordset` **modifiable** : seuls les types **Table** et **Dynaset** le permettent. Les types **Snapshot** et **Forward-only** étant en lecture seule (section [9.3](/09-dao-data-access-objects/03-recordset-types.md)), toute tentative d'ajout sur ces types échoue. C'est une vérification à faire en amont lorsque l'ajout n'aboutit pas comme attendu.

## `AddNew` ou `INSERT` SQL ?

Deux approches permettent d'insérer des données, et le choix dépend du contexte. Pour ajouter **un grand nombre d'enregistrements** d'un coup, en particulier à partir d'une autre source, une requête `INSERT INTO … SELECT` exécutée via `db.Execute` est généralement bien plus rapide qu'une boucle d'`AddNew`, car elle évite les allers-retours répétés avec le moteur. À l'inverse, `AddNew` s'impose lorsqu'il faut appliquer une **logique propre à chaque enregistrement**, traiter les lignes une à une, ou récupérer aisément la clé générée. Pour un enregistrement isolé, `AddNew` est souvent la solution la plus lisible.

Les requêtes action `INSERT` sont traitées à la section [11.3](/11-sql-access-vba/03-requetes-action-insert-update-delete.md), et la comparaison `Execute`/`RunSQL` à la section [18.3](/18-optimisation-performance/03-execute-vs-runsql.md).

## Et pour un formulaire ?

Lorsqu'on souhaite ajouter un enregistrement à un formulaire ouvert, on n'agit pas directement sur la source du formulaire mais sur sa copie de travail, le `RecordsetClone`, selon des principes spécifiques exposés à la section [9.12](/09-dao-data-access-objects/12-recordsetclone.md).

## Pièges courants

Les erreurs les plus fréquentes sont les suivantes. Croire que l'enregistrement est ajouté dès `AddNew` : il ne l'est qu'après `Update`. S'attendre à ce que le curseur pointe sur le nouvel enregistrement après `Update` : il revient à sa position antérieure, et il faut `LastModified` pour atteindre le nouvel enregistrement. Quitter le mode ajout par un déplacement plutôt que par `Update` ou `CancelUpdate`, ce qui perd l'ajout. Oublier de renseigner un champ obligatoire, ou tenter d'insérer une clé en doublon, ce qui échoue à l'`Update`. Enfin, tenter un `AddNew` sur un `Recordset` non modifiable (Snapshot, Forward-only).

## Points clés à retenir

L'ajout d'un enregistrement en DAO suit trois temps : `AddNew`, affectation des champs, puis `Update` qui valide réellement l'insertion. Après `Update`, le curseur revient à sa position d'origine ; pour atteindre l'enregistrement créé — et lire par exemple sa clé NuméroAuto — on emploie `rs.Bookmark = rs.LastModified`. On annule un ajout en cours par `CancelUpdate`, et l'on ne laisse jamais un `Recordset` en attente entre `AddNew` et `Update`. L'opération exige un `Recordset` modifiable (Table ou Dynaset) et vérifie les contraintes à la validation, d'où l'intérêt d'une gestion d'erreurs, voire d'une transaction pour des ajouts multiples. Pour de gros volumes, une requête `INSERT` reste préférable à une boucle d'`AddNew`.

⏭️ [9.7. Modification d'enregistrements (Edit/Update)](/09-dao-data-access-objects/07-modification-enregistrements.md)
