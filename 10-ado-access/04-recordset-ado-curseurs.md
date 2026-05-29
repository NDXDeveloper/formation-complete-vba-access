🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.4. L'objet Recordset ADO — curseurs et modes de verrouillage

Si l'objet `Connection` établit le lien avec la source de données, c'est l'objet **`Recordset`** qui transporte réellement les enregistrements. C'est le cheval de bataille d'ADO, l'objet avec lequel on lit, parcourt et modifie les données. Mais contrairement au recordset DAO, le recordset ADO se définit par **deux paramètres déterminants** : le *type de curseur* et le *mode de verrouillage*.

Ces deux notions — propres à ADO — gouvernent presque tout : la façon dont on navigue, la possibilité ou non de modifier les données, la visibilité des changements effectués par d'autres utilisateurs, les performances et la consommation mémoire. Bien les comprendre, c'est éviter la double erreur classique : choisir un curseur trop coûteux pour une simple lecture, ou un curseur trop limité pour une édition. Cette section pose ces fondations ; les opérations concrètes de lecture et d'écriture suivront en section 10.5.

---

## Le Recordset, cheval de bataille d'ADO

Un `Recordset` représente un **ensemble d'enregistrements** issu d'une table, d'une requête `SELECT` ou d'une commande. Il joue le même rôle fonctionnel que le recordset DAO (chapitre 9.3), mais avec son propre modèle conceptuel. Là où DAO distingue des *types* de recordset (Table, Dynaset, Snapshot, Forward-only), ADO sépare deux dimensions indépendantes : **comment on parcourt** les données (le curseur) et **comment on les verrouille** pour les modifier (le verrouillage). C'est la combinaison de ces deux choix qui façonne le comportement final.

---

## La signature de la méthode `Open`

Toute l'expressivité du recordset se concentre dans les paramètres de sa méthode `Open` :

```vba
recordset.Open Source, ActiveConnection, CursorType, LockType, Options
```

Le **`Source`** est la requête SQL, le nom de table ou l'objet `Command` à exploiter. **`ActiveConnection`** est la connexion sur laquelle s'appuyer (un objet `Connection`, ou à défaut une chaîne de connexion). **`CursorType`** et **`LockType`** sont les deux paramètres centraux que cette section détaille. Enfin, **`Options`** précise la nature du `Source` (SQL, table, procédure…) afin d'éviter à ADO d'avoir à la deviner.

```vba
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM Clients", cn, adOpenForwardOnly, adLockReadOnly, adCmdText
```

Ces paramètres peuvent aussi être positionnés via les propriétés correspondantes (`rs.CursorType`, `rs.LockType`, `rs.CursorLocation`) **avant** l'appel à `Open` — utile, notamment, pour fixer l'emplacement du curseur, comme nous le verrons.

---

## Les types de curseurs (`CursorType`)

ADO propose quatre types de curseurs, du plus économe au plus puissant.

Le curseur **`adOpenForwardOnly`** (valeur 0, le défaut) n'autorise qu'un parcours **vers l'avant** : on avance d'enregistrement en enregistrement, sans pouvoir revenir en arrière. C'est le plus rapide et le moins gourmand en mémoire, mais le plus limité. Il ne reflète aucun changement effectué par d'autres et ne donne pas de comptage fiable.

Le curseur **`adOpenStatic`** (valeur 3) fournit un **instantané figé** des données au moment de l'ouverture. Il permet une navigation **bidirectionnelle** complète (avant, arrière, premier, dernier) et un comptage fiable, mais ne montre aucune modification ultérieure faite par d'autres utilisateurs. C'est le curseur des recordsets côté client et déconnectés.

Le curseur **`adOpenKeyset`** (valeur 1) fixe à l'ouverture l'**ensemble des clés** (la liste des enregistrements concernés), tout en restant sensible aux modifications : on voit les **changements de valeurs** et les **suppressions** effectués par d'autres sur ces enregistrements, mais **pas les ajouts** survenus après l'ouverture. Il offre navigation bidirectionnelle et comptage fiable.

Le curseur **`adOpenDynamic`** (valeur 2) est le plus « vivant » : il reflète **tous les changements** des autres utilisateurs — modifications, suppressions *et* ajouts. C'est aussi le plus coûteux, et son comptage est souvent indisponible.

| Type de curseur | Navigation | Voit modifs/suppressions d'autrui | Voit les ajouts d'autrui | `RecordCount` | Coût |
|---|---|---|---|---|---|
| `adOpenForwardOnly` | Avant uniquement | Non | Non | `-1` (indisponible) | Minimal |
| `adOpenStatic` | Bidirectionnelle | Non (instantané figé) | Non | Fiable | Modéré |
| `adOpenKeyset` | Bidirectionnelle | Oui | Non | Fiable | Modéré à élevé |
| `adOpenDynamic` | Bidirectionnelle | Oui | Oui | Souvent `-1` | Élevé |

> ⚠️ **Spécificité du moteur ACE** : le fournisseur ACE/Jet ne prend pas réellement en charge les curseurs **dynamiques** côté serveur. Une demande de `adOpenDynamic` est généralement *rétrogradée* silencieusement en keyset. De même, un curseur côté client est **toujours** statique, quel que soit le type demandé. Autrement dit, le curseur que vous *obtenez* peut différer de celui que vous *demandez* — un point sur lequel nous reviendrons.

---

## Les modes de verrouillage (`LockType`)

Le second paramètre détermine si — et comment — les données peuvent être modifiées. À la différence de DAO, **ADO ne possède pas de méthode `Edit`** : on modifie directement les champs puis on appelle `Update`. C'est le mode de verrouillage choisi à l'ouverture qui régit le comportement de ces modifications.

Le mode **`adLockReadOnly`** (valeur 1, le défaut) interdit toute modification : le recordset est en **lecture seule**. C'est le mode le plus rapide, à privilégier dès qu'aucune écriture n'est prévue.

Le mode **`adLockPessimistic`** (valeur 2) applique un verrouillage **pessimiste** : l'enregistrement est verrouillé **dès le début de son édition** et le reste jusqu'à l'appel d'`Update` (ou son annulation). Il garantit qu'aucun autre utilisateur ne modifiera l'enregistrement pendant la saisie, au prix d'une plus forte contention.

Le mode **`adLockOptimistic`** (valeur 3) applique un verrouillage **optimiste** : l'enregistrement n'est verrouillé que brièvement, **au moment de l'`Update`**. La contention est moindre, mais un conflit peut survenir si deux utilisateurs ont modifié le même enregistrement entre-temps (gestion des conflits, chapitre 14.5). C'est le mode habituel pour un recordset modifiable.

Le mode **`adLockBatchOptimistic`** (valeur 4) met en cache les modifications pour les appliquer **par lot** via `UpdateBatch`. Il est indissociable des recordsets déconnectés et des mises à jour groupées (section 10.7).

| Mode de verrouillage | Édition | Quand le verrou est posé | Usage typique |
|---|---|---|---|
| `adLockReadOnly` | Impossible | — | Lecture seule (le plus rapide) |
| `adLockPessimistic` | Oui | Dès le début de l'édition | Forte concurrence, données critiques |
| `adLockOptimistic` | Oui | Au moment de l'`Update` | Édition courante, faible concurrence |
| `adLockBatchOptimistic` | Oui | À l'`UpdateBatch` | Recordsets déconnectés, mises à jour par lot |

---

## L'emplacement du curseur : `adUseServer` ou `adUseClient`

Une troisième dimension, souvent négligée, complète le tableau : l'**emplacement** du curseur, fixé par la propriété `CursorLocation` (sur la connexion ou sur le recordset, **avant** ouverture).

Avec **`adUseServer`** (le défaut), le curseur est géré **côté source de données**. C'est l'option qui, en théorie, donne accès à toute la palette des types de curseurs — mais qui, avec ACE, reste soumise aux limitations évoquées plus haut.

Avec **`adUseClient`**, le curseur est géré **côté client**, en mémoire. Ce mode présente deux caractéristiques majeures. D'une part, il produit **toujours un curseur statique**, quel que soit le `CursorType` demandé. D'autre part, il est **indispensable aux recordsets déconnectés** (section 10.7) et débloque certaines fonctionnalités absentes des curseurs serveur — un `RecordCount` toujours fiable, la propriété `Sort`, certains filtres côté client.

```vba
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient            ' AVANT Open
rs.Open "SELECT * FROM Commandes", cn, adOpenStatic, adLockBatchOptimistic, adCmdText
```

---

## Le « firehose cursor » : la combinaison la plus rapide

La combinaison **`adOpenForwardOnly` + `adLockReadOnly`** porte un surnom dans la communauté : le *firehose cursor* (« curseur lance à incendie »). C'est la configuration **la plus rapide possible** pour lire des données : un seul passage vers l'avant, sans verrouillage, sans surcharge.

C'est aussi la **combinaison par défaut** d'un `Open` dont on ne précise ni le curseur ni le verrouillage. Elle est idéale pour tous les traitements en un seul passage : alimenter une liste déroulante, produire un état, agréger des valeurs, parcourir des enregistrements sans avoir besoin de revenir en arrière ni de les modifier.

```vba
' Lecture la plus rapide : un seul parcours, lecture seule
rs.Open "SELECT NomClient FROM Clients ORDER BY NomClient", cn, _
        adOpenForwardOnly, adLockReadOnly, adCmdText
```

Le réflexe à adopter : **ne payez que ce dont vous avez besoin**. Un curseur bidirectionnel modifiable consomme des ressources inutiles si vous vous contentez de lire une fois vers l'avant.

---

## Vérifier ce que l'on a vraiment obtenu

Compte tenu des rétrogradations silencieuses du fournisseur ACE, il est prudent — surtout en phase de mise au point — de **vérifier le curseur réellement accordé** plutôt que de supposer qu'il correspond à la demande. Les propriétés `CursorType`, `LockType` et `CursorLocation`, lues *après* l'ouverture, renvoient l'état effectif :

```vba
rs.Open "SELECT * FROM Clients", cn, adOpenDynamic, adLockOptimistic, adCmdText
Debug.Print "Curseur demandé : dynamique"
Debug.Print "Curseur obtenu  : " & rs.CursorType   ' souvent 1 (keyset) avec ACE
```

La méthode **`Supports`** complète cette vérification en testant si le recordset prend en charge une fonctionnalité donnée, avant de s'y fier :

```vba
If rs.Supports(adMovePrevious) Then rs.MovePrevious   ' navigation arrière possible ?
If rs.Supports(adUpdate) Then ' ... modification possible ?
```

Cette précaution évite des erreurs d'exécution sur des opérations que le curseur effectivement obtenu ne permet pas.

---

## Le paramètre `Options` : aider ADO et gagner en performance

Le dernier paramètre d'`Open` indique la **nature du `Source`**. Le préciser n'est pas obligatoire, mais c'est recommandé : sans lui, ADO doit *deviner* le type de commande, ce qui occasionne des allers-retours superflus.

Les valeurs les plus courantes sont **`adCmdText`** (le `Source` est une instruction SQL), **`adCmdTable`** (le `Source` est un nom de table, ADO génère un `SELECT *`), **`adCmdStoredProc`** (procédure stockée) et **`adCmdTableDirect`** (ouverture directe d'une table, nécessaire pour exploiter un index). Indiquer la bonne valeur est une optimisation simple et systématique.

```vba
rs.Open "Clients", cn, adOpenStatic, adLockReadOnly, adCmdTable   ' nom de table explicite
```

---

## Choisir la bonne combinaison

Plutôt que de mémoriser toutes les combinaisons possibles, raisonnez par besoin. Le tableau suivant propose des configurations de référence.

| Besoin | `CursorType` | `LockType` | `CursorLocation` |
|---|---|---|---|
| Lecture rapide, un seul passage | `adOpenForwardOnly` | `adLockReadOnly` | `adUseServer` |
| Lecture avec navigation / comptage | `adOpenStatic` | `adLockReadOnly` | au choix |
| Édition courante (faible concurrence) | `adOpenKeyset` | `adLockOptimistic` | `adUseServer` |
| Édition à concurrence stricte | `adOpenKeyset` | `adLockPessimistic` | `adUseServer` |
| Déconnecté / mise à jour par lot | `adOpenStatic` | `adLockBatchOptimistic` | `adUseClient` |

Deux principes guident ces choix. D'abord, **commencez par le moins coûteux** : le firehose pour lire, et n'enrichissez le curseur que si un besoin réel l'exige (revenir en arrière, compter, modifier). Ensuite, **alignez le verrouillage sur l'intention** : `adLockReadOnly` si vous ne modifiez rien, `adLockOptimistic` pour l'édition usuelle, et réservez le pessimiste aux situations où la concurrence sur un même enregistrement doit être strictement empêchée. Les implications de ces choix en environnement multi-utilisateur sont approfondies au chapitre 15.

---

## En résumé

Le recordset ADO se définit par deux paramètres indissociables : le **type de curseur** (forward-only, statique, keyset, dynamique), qui régit la navigation, la visibilité des changements et le coût, et le **mode de verrouillage** (lecture seule, pessimiste, optimiste, par lot), qui régit l'édition. Une troisième dimension, l'**emplacement du curseur** (serveur ou client), conditionne notamment l'accès aux fonctionnalités déconnectées. La combinaison forward-only + lecture seule — le *firehose* — est la plus rapide et constitue le défaut pour la simple lecture. Enfin, le moteur ACE imposant ses propres limites (pas de curseur dynamique réel, curseur client toujours statique), il est prudent de vérifier après ouverture le curseur réellement accordé via `CursorType`, `LockType` et la méthode `Supports`.

Maintenant que nous savons *ouvrir* un recordset dans la bonne configuration, il reste à l'exploiter : lire ses champs, modifier des enregistrements, en ajouter et en supprimer. C'est l'objet de la section suivante.


⏭️ [10.5. Lecture, modification, ajout et suppression avec ADO](/10-ado-access/05-crud-ado.md)
