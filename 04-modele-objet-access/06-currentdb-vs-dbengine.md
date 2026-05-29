🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.6. CurrentDb() vs DBEngine(0)(0) — différences et cas d'usage

Nous y voici : la distinction que tout ce chapitre préparait. La section 4.5 a montré que **deux écritures** mènent à la base courante — `CurrentDb()` et `DBEngine(0)(0)` —, mais qu'elles **ne sont pas équivalentes**. Comprendre pourquoi, et savoir laquelle employer, évite certains des bugs les plus déroutants du développement Access. Cette section est, sans exagérer, l'une des plus importantes du chapitre.

## La différence fondamentale

Tout tient à la nature des deux expressions :

- **`CurrentDb()`** est une **méthode** de l'objet `Application`. À **chaque appel**, elle renvoie un **nouvel objet `Database`** indépendant, qui pointe vers la base courante. Comme cet objet est neuf, ses collections (`TableDefs`, `QueryDefs`…) reflètent l'**état courant** de la base.
- **`DBEngine(0)(0)`** renvoie une référence vers **le même et unique objet `Database`** qu'Access utilise en interne — celui que l'interface elle-même manipule. Ses collections ne se rafraîchissent pas automatiquement et peuvent être **périmées**.

En une phrase : `CurrentDb()` fabrique une **vue neuve et à jour**, tandis que `DBEngine(0)(0)` donne accès à l'**objet partagé d'Access**, susceptible d'être en retard sur la réalité.

## Conséquence 1 : la fraîcheur des collections

Premier effet concret. Si l'on **crée un objet par code** (une requête, une table — chapitre 12) ou que la structure change, l'objet `DBEngine(0)(0)`, avec ses collections en cache, peut **ne pas le voir** immédiatement ; il faudrait appeler `.Refresh` sur la collection. `CurrentDb()`, objet neuf, reflète d'emblée l'état courant.

```vba
' On crée une requête par code (chapitre 12)
CurrentDb.CreateQueryDef "qryTest", "SELECT * FROM tblClients"

' DBEngine(0)(0) peut ne PAS la voir tout de suite (collection en cache)
'   -> nécessiterait DBEngine(0)(0).QueryDefs.Refresh

' CurrentDb, objet neuf, reflète l'état courant
Debug.Print CurrentDb.QueryDefs("qryTest").Name   ' OK
```

C'est la cause de ce symptôme classique : « j'ai créé une requête par code, mais mon énumération ne la voit pas ».

## Conséquence 2 : le piège des appels répétés à CurrentDb

Voici **le** piège le plus célèbre — et le revers de la qualité précédente. Puisque **chaque appel** à `CurrentDb` renvoie un **objet différent**, deux appels consécutifs ne partagent **aucun état**. L'exemple canonique concerne `RecordsAffected`, qui indique combien d'enregistrements une requête action a touchés :

```vba
' ❌ INCORRECT : deux appels à CurrentDb = deux objets différents
CurrentDb.Execute "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02", dbFailOnError
Debug.Print CurrentDb.RecordsAffected   ' renvoie 0 : ce n'est PAS le même objet !

' ✅ CORRECT : on capture l'objet une fois, puis on le réutilise
Dim db As DAO.Database
Set db = CurrentDb
db.Execute "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02", dbFailOnError
Debug.Print db.RecordsAffected          ' renvoie le bon nombre
```

Le `RecordsAffected` de la version incorrecte porte sur un **objet neuf** qui n'a rien exécuté : il vaut donc 0. C'est une erreur extrêmement fréquente, qui se règle en **capturant `CurrentDb` dans une variable**.

## Conséquence 3 : la performance

Troisième effet. Comme `CurrentDb()` **crée un objet** à chaque appel, l'invoquer en boucle est un gaspillage. La règle est la même : **capturer une fois**, réutiliser ensuite.

```vba
' ❌ À éviter : CurrentDb recréé à chaque tour de boucle
Dim i As Long
For i = 1 To 1000
    CurrentDb.Execute "...", dbFailOnError   ' coûteux
Next i

' ✅ Préférable : un seul objet, réutilisé
Dim db As DAO.Database
Set db = CurrentDb
For i = 1 To 1000
    db.Execute "...", dbFailOnError
Next i
```

## La bonne pratique

Les trois conséquences convergent vers une recommandation simple et **quasi universelle** :

> **Capturez `CurrentDb` dans une variable une seule fois, puis réutilisez cette variable.**

```vba
Dim db As DAO.Database
Set db = CurrentDb
' ... toutes les opérations sur db ...
Set db = Nothing
```

Cette pratique cumule les avantages : une **vue à jour** (pas de collection périmée), un **objet stable** (donc `RecordsAffected` fiable et état cohérent), et **aucun coût** d'appels répétés. C'est l'idiome standard de l'accès aux données sous Access, que l'on retrouvera dans tous les exemples DAO des chapitres 9 et suivants.

## Quand utiliser quoi

- **`CurrentDb()`** (stocké dans une variable) : le choix **par défaut** pour la quasi-totalité des accès aux données et des manipulations de structure. C'est ce que l'on recommande dans le code moderne.
- **`DBEngine(0)(0)`** : à réserver aux **cas particuliers** où l'on a précisément besoin du **même objet** qu'utilise Access (situations rares), ou à du code hérité. Pour du code neuf, on lui préfère `CurrentDb`.

## Un cas voisin : CodeDb()

Mentionnons enfin une fonction proche, utile à connaître : **`CodeDb()`**. Là où `CurrentDb` renvoie la base **de l'utilisateur** (le front-end courant), `CodeDb` renvoie la base **qui contient le code en cours d'exécution**. La distinction n'a d'importance que dans un scénario de **bibliothèque** ou de **complément** (add-in) : le code d'un complément, exécuté pendant que l'utilisateur travaille dans sa propre base, doit parfois référencer **sa** base à lui, et non celle de l'utilisateur.

```vba
' CurrentDb : la base de l'utilisateur (front-end)
' CodeDb    : la base contenant le code en cours (bibliothèque / complément)
Dim dbCode As DAO.Database
Set dbCode = CodeDb
```

## Tableau comparatif

| Critère | `CurrentDb()` | `DBEngine(0)(0)` |
|---|---|---|
| Nature | méthode d'`Application` | objet existant de l'espace par défaut |
| Objet renvoyé | un **nouvel** objet à chaque appel | **le même** objet (celui d'Access) |
| Collections | rafraîchies (vue à jour) | potentiellement périmées (`.Refresh` requis) |
| Partage avec l'interface Access | non (objet indépendant) | oui |
| Coût | création d'un objet à chaque appel | accès immédiat |
| Recommandation | **à privilégier**, dans une variable | legacy / cas particuliers |

## À retenir

- **`CurrentDb()`** est une **méthode** qui renvoie un **nouvel objet `Database` à jour** à chaque appel ; **`DBEngine(0)(0)`** renvoie **le même objet partagé** avec Access, aux collections potentiellement **périmées**.
- **Piège n°1** : créer un objet par code peut ne pas apparaître dans `DBEngine(0)(0)` (cache) — `CurrentDb` reflète l'état courant.
- **Piège n°2 (le plus fréquent)** : `CurrentDb.Execute` suivi de `CurrentDb.RecordsAffected` **échoue** (deux objets différents) — il faut **capturer `CurrentDb` dans une variable**.
- **Bonne pratique** : `Set db = CurrentDb` **une seule fois**, puis réutiliser `db` — vue à jour, objet stable, sans surcoût. C'est l'idiome standard (chapitres 9 et suivants).
- **`CodeDb()`** renvoie la base **contenant le code** (utile pour les bibliothèques/compléments), à distinguer de `CurrentDb` (base de l'utilisateur).

---


⏭️ [4.7. Screen et SysCmd — objets utilitaires souvent oubliés](/04-modele-objet-access/07-screen-syscmd.md)
