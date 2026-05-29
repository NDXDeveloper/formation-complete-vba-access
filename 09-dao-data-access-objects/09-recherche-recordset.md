🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 9.9 Recherche dans un Recordset (`FindFirst`, `FindNext`, `Seek`)

## Introduction

Naviguer séquentiellement dans un `Recordset` (section [9.4](/09-dao-data-access-objects/04-navigation-recordset.md)) permet de parcourir tous les enregistrements, mais pas de localiser directement celui que l'on cherche. DAO propose pour cela deux familles de méthodes de recherche, aux caractéristiques très différentes : les méthodes **`Find`**, souples et applicables à la plupart des jeux d'enregistrements, et la méthode **`Seek`**, beaucoup plus rapide mais réservée à un usage précis. Cette section détaille les deux, le moyen de tester leur résultat, et — point important — la question de savoir s'il vaut mieux rechercher dans un `Recordset` ou filtrer dès la source.

## Deux familles de recherche

Le choix de la méthode dépend d'abord du **type** de `Recordset` (section [9.3](/09-dao-data-access-objects/03-recordset-types.md)). Les méthodes `Find` s'appliquent aux types **Dynaset** et **Snapshot**, mais pas au type Table ni au type Forward-only. La méthode `Seek`, à l'inverse, ne fonctionne **que** sur le type **Table** et requiert qu'un index ait été désigné au préalable. Cette répartition n'est pas un détail : tenter un `Find` sur un `Recordset` de type Table, ou un `Seek` sur un Dynaset, échoue.

## Tester le résultat : `NoMatch`

Avant toute chose, une règle commune aux deux familles : après une recherche, on **teste systématiquement la propriété `NoMatch`** avant d'utiliser l'enregistrement courant. `NoMatch` vaut `False` si un enregistrement a été trouvé (le curseur est alors positionné dessus) et `True` dans le cas contraire. Lorsque `NoMatch` est `True`, la position du curseur est indéfinie : il ne faut surtout pas se fier à l'enregistrement courant.

```vba
rs.FindFirst "NomClient = 'Dupont'"
If Not rs.NoMatch Then
    Debug.Print "Trouvé : " & rs!Ville   ' positionné sur l'enregistrement
Else
    Debug.Print "Aucun enregistrement correspondant"
End If
```

Omettre ce test est l'une des erreurs les plus courantes : on risque alors de lire ou de modifier un enregistrement qui n'est pas celui attendu.

## Les méthodes `Find`

Les quatre méthodes `Find` se distinguent par leur point de départ et leur sens : `FindFirst` cherche depuis le début du jeu, `FindLast` depuis la fin, `FindNext` à partir de la position courante vers l'avant, et `FindPrevious` vers l'arrière. Toutes prennent en argument une **chaîne de critère** qui s'écrit comme une clause `WHERE` sans le mot-clé `WHERE` : `"Ville = 'Rouen'"`.

Pour parcourir **tous** les enregistrements correspondant à un critère, on combine `FindFirst` et `FindNext` dans une boucle, en répétant le critère :

```vba
Dim critere As String
critere = "Ville = 'Rouen'"

rs.FindFirst critere
Do While Not rs.NoMatch
    Debug.Print rs!NomClient
    rs.FindNext critere
Loop
```

La boucle s'arrête naturellement lorsque `FindNext` ne trouve plus de correspondance et passe `NoMatch` à `True`.

## Le format du critère

La chaîne de critère obéit aux règles d'expression du moteur, qui réservent quelques pièges. Les valeurs **texte** sont délimitées par des apostrophes ou des guillemets : `"Ville = 'Rouen'"`. Les **dates** doivent être encadrées de dièses **et exprimées au format américain** (mm/jj/aaaa), indépendamment des réglages régionaux — un point de localisation traité plus largement à la section [11.7](/11-sql-access-vba/07-localisation-formatage-sql.md). On construit donc un critère de date ainsi :

```vba
rs.FindFirst "DateCommande = #" & Format(maDate, "mm/dd/yyyy") & "#"
```

Un autre piège concerne les chaînes contenant une **apostrophe** (par exemple un nom comme « O'Brien ») : avec des apostrophes comme délimiteurs, le critère devient invalide. On utilise alors le guillemet double comme délimiteur, ou l'on double les apostrophes internes :

```vba
' Délimiteur guillemet double, robuste face aux apostrophes
rs.FindFirst "Nom = " & Chr(34) & nom & Chr(34)
```

D'une manière générale, construire un critère par concaténation demande de la rigueur sur les délimiteurs ; les bonnes pratiques de construction dynamique sont détaillées à la section [11.6](/11-sql-access-vba/06-construction-dynamique-sql.md).

## La méthode `Seek`

`Seek` exploite directement un **index** de la table, ce qui en fait une recherche extrêmement rapide — sans commune mesure avec le parcours séquentiel d'un `Find` sur de gros volumes. Son usage impose deux conditions : un `Recordset` de type Table, et la désignation de l'index à utiliser via la propriété `Index` avant l'appel.

```vba
Dim rs As DAO.Recordset
Set rs = db.OpenRecordset("tblClients", dbOpenTable)

rs.Index = "PrimaryKey"           ' index sur lequel porte la recherche
rs.Seek "=", 42                   ' cherche la clé égale à 42

If Not rs.NoMatch Then
    Debug.Print rs!NomClient
End If
```

Le premier argument de `Seek` est un **opérateur de comparaison** passé sous forme de chaîne : `"="`, `">"`, `">="`, `"<"` ou `"<="`. Les arguments suivants sont les **valeurs de clé** recherchées ; pour un index portant sur plusieurs colonnes, on fournit une valeur par colonne, dans l'ordre de l'index. À la différence de `FindNext`, `Seek` ne part pas de la position courante : il localise toujours, via l'index, le premier enregistrement satisfaisant le critère. L'index utilisé doit évidemment exister sur la table ; la création d'index par code est traitée à la section [12.5](/12-querydefs-tabledefs/05-creer-modifier-tables.md).

## `Find` ou `Seek` ?

| Critère | Méthodes `Find` | `Seek` |
|---|---|---|
| Types de `Recordset` | Dynaset, Snapshot | Table uniquement |
| Index requis | Non | Oui (propriété `Index`) |
| Forme du critère | Chaîne type clause `WHERE` | Opérateur + valeur(s) de clé |
| Performance | Séquentielle, plus lente | Indexée, très rapide |
| Sources acceptées | Tables, requêtes, tables liées | Table locale uniquement |
| Indicateur de résultat | `NoMatch` | `NoMatch` |

En résumé : lorsqu'on recherche par une clé indexée sur une table locale et que la performance compte, `Seek` est imbattable. Dans tous les autres cas — table liée, requête, jointure, critère portant sur un champ non indexé — on emploie `Find`.

## Rechercher ou filtrer à la source ?

Il faut prendre du recul sur ces méthodes. Bien souvent, la meilleure « recherche » consiste à **ne pas rechercher du tout** dans le `Recordset`, mais à ouvrir d'emblée un jeu déjà restreint par une clause `WHERE` :

```vba
Set rs = db.OpenRecordset( _
    "SELECT * FROM tblClients WHERE Ville = 'Rouen'", dbOpenSnapshot)
```

Cette approche est généralement plus efficace que d'ouvrir l'ensemble de la table puis de la parcourir avec `Find`, car le moteur exploite ses index et ne renvoie que les enregistrements pertinents. Les méthodes `Find` et `Seek` trouvent donc surtout leur intérêt lorsqu'on dispose **déjà** d'un `Recordset` ouvert et qu'on veut y localiser un enregistrement sans relancer de requête — c'est précisément le cas classique du `RecordsetClone` d'un formulaire, traité à la section [9.12](/09-dao-data-access-objects/12-recordsetclone.md). Le filtrage à la source via `SELECT` est abordé à la section [11.2](/11-sql-access-vba/02-select-recordset.md), et le filtrage côté `Recordset` (propriétés `Filter` et `Sort`) à la section suivante, [9.10](/09-dao-data-access-objects/10-tri-filtrage-recordset.md).

## Performance

Sur de gros volumes, l'écart de performance est considérable : `Seek`, indexé, est d'un ordre de grandeur plus rapide que `Find`, qui parcourt les enregistrements un à un. Mais l'optimisation la plus efficace reste, là encore, de filtrer dès la source : moins de données rapatriées, index exploités par le moteur. L'impact des index sur les performances est approfondi à la section [18.6](/18-optimisation-performance/06-indexation-performances-sql.md).

## Pièges courants

Les erreurs récurrentes sont les suivantes. Ne pas tester `NoMatch` après une recherche, et utiliser un enregistrement courant indéfini. Employer `Find` sur un `Recordset` de type Table, ou `Seek` sur un Dynaset, ce qui échoue. Oublier de désigner l'`Index` avant un `Seek`. Mal délimiter un critère `Find` — apostrophes dans une chaîne, dates non exprimées au format américain. Et, plus largement, parcourir toute une table avec `Find` là où une requête filtrée à la source aurait été nettement plus efficace.

## Points clés à retenir

DAO offre deux moyens de recherche dans un `Recordset` : les méthodes `Find` (sur Dynaset et Snapshot), souples mais séquentielles, et `Seek` (sur Table uniquement, avec un index désigné), beaucoup plus rapide. Dans les deux cas, on teste impérativement `NoMatch` avant d'exploiter l'enregistrement courant. La construction du critère `Find` exige de la vigilance sur les délimiteurs de texte et le format américain des dates. Enfin, lorsque c'est possible, filtrer les données dès la source par une clause `WHERE` est préférable à une recherche dans un jeu complet : `Find` et `Seek` brillent surtout pour localiser un enregistrement dans un `Recordset` déjà ouvert, comme celui d'un formulaire.

⏭️ [9.10. Tri et filtrage côté Recordset](/09-dao-data-access-objects/10-tri-filtrage-recordset.md)
