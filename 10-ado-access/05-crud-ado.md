🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 10.5. Lecture, modification, ajout et suppression avec ADO

Nous savons désormais ouvrir un recordset dans la bonne configuration (section 10.4). Il reste à l'exploiter concrètement : c'est le cœur du travail sur les données, et l'objet de cette section. Nous couvrons ici les quatre opérations fondamentales, regroupées sous l'acronyme **CRUD** — *Create, Read, Update, Delete* : ajouter, lire, modifier et supprimer des enregistrements.

Une mise en garde s'impose d'emblée pour qui vient de DAO : **ADO ne possède pas de méthode `Edit`**. Le flux de modification diffère donc, et cette différence est la principale source de confusion lors du passage de DAO à ADO. Nous y reviendrons explicitement.

---

## Lire et parcourir un recordset

### Tester si le recordset est vide

Avant de parcourir un recordset, il faut s'assurer qu'il contient des données. Le test le plus fiable, valable quel que soit le curseur, repose sur les propriétés **`EOF`** (*End Of File*) et **`BOF`** (*Begin Of File*) : si les deux sont vraies simultanément, le recordset est vide.

```vba
If rs.BOF And rs.EOF Then
    MsgBox "Aucun enregistrement trouvé."
Else
    ' ... traitement ...
End If
```

La propriété `RecordCount` peut aussi servir (`If rs.RecordCount = 0`), mais rappelons qu'elle n'est fiable qu'avec un curseur statique, keyset ou côté client — pas avec un firehose, où elle renvoie `-1` (section 10.4). Le test `BOF And EOF` est donc plus universel.

### La boucle de parcours standard

Le parcours classique avance du premier au dernier enregistrement tant que `EOF` n'est pas atteint :

```vba
Do While Not rs.EOF
    Debug.Print rs.Fields("NomClient").Value
    rs.MoveNext
Loop
```

L'oubli du **`rs.MoveNext`** est l'erreur de débutant par excellence : sans lui, la boucle reste figée sur le premier enregistrement et tourne indéfiniment. Pour les curseurs qui le permettent (tous sauf forward-only), la navigation s'enrichit de `MoveFirst`, `MovePrevious`, `MoveLast` et `Move n`.

### Accéder aux valeurs des champs

ADO offre plusieurs notations pour lire un champ, de la plus explicite à la plus concise :

```vba
rs.Fields("NomClient").Value   ' notation complète et explicite
rs.Fields("NomClient")         ' .Value est la propriété par défaut
rs("NomClient")                ' Fields est la collection par défaut
rs!NomClient                   ' notation « bang », la plus brève
```

Toutes désignent le même champ. En code de production, la **notation complète** (`rs.Fields("…").Value`) est la plus lisible et la moins ambiguë ; la **notation bang** (`rs!…`) est appréciée pour sa concision dans des traitements courts. L'essentiel est de rester cohérent au sein d'un projet.

### Gérer les valeurs `Null`

Un champ vide renvoie la valeur spéciale **`Null`**, et toute opération (concaténation, calcul) impliquant `Null` propage ce `Null` — source d'erreurs fréquentes. En contexte Access, la fonction **`Nz`** est l'outil de prédilection pour neutraliser ce risque en fournissant une valeur de remplacement :

```vba
Dim sVille As String
sVille = Nz(rs!Ville, "(non renseignée)")   ' "" ou tout autre défaut si Null
```

À défaut, un test explicite par `IsNull(rs.Fields("Ville").Value)` permet de traiter le cas séparément. Ne jamais présumer qu'un champ contient une valeur est une discipline qui épargne bien des plantages.

### La collection `Fields`

Le recordset expose une collection **`Fields`** que l'on peut parcourir, par exemple pour traiter dynamiquement toutes les colonnes :

```vba
Dim fld As ADODB.Field
For Each fld In rs.Fields
    Debug.Print fld.Name & " = " & Nz(fld.Value, "Null")
Next fld
```

Chaque objet `Field` porte aussi des métadonnées utiles (`Name`, `Type`, `DefinedSize`, etc.), précieuses pour du code générique.

---

## Modifier un enregistrement (`Update`) — sans `Edit` !

Voici le point qui déroute les habitués de DAO. En DAO, modifier un enregistrement impose d'appeler d'abord `Edit`. **En ADO, cette étape n'existe pas** : on positionne directement les nouvelles valeurs, puis on valide avec `Update`.

```vba
' Le recordset doit avoir été ouvert avec un verrouillage modifiable
' (adLockOptimistic ou adLockPessimistic) et un curseur le permettant.
rs.Fields("Ville").Value = "Lyon"
rs.Fields("Actif").Value = True
rs.Update   ' valide les modifications
```

Tenter de modifier un champ sur un recordset ouvert en `adLockReadOnly` déclenche une erreur : le mode de verrouillage choisi à l'ouverture conditionne donc directement cette opération (section 10.4).

ADO propose une variante compacte de `Update`, qui accepte un **tableau de champs** et un **tableau de valeurs** :

```vba
rs.Update Array("Ville", "Actif"), Array("Lyon", True)
```

Deux comportements méritent d'être connus. D'une part, **se déplacer** vers un autre enregistrement (`MoveNext`, etc.) **valide implicitement** les modifications en attente — comme en DAO. D'autre part, la méthode **`CancelUpdate`** permet d'annuler les changements non encore validés. La propriété **`EditMode`** renseigne enfin sur l'état d'édition courant (`adEditNone`, `adEditInProgress`, `adEditAdd`, `adEditDelete`), utile pour savoir si une validation est en attente.

---

## Ajouter un enregistrement (`AddNew`)

L'ajout suit une logique en trois temps : `AddNew` crée un nouvel enregistrement vierge, on renseigne ses champs, puis `Update` l'enregistre définitivement.

```vba
rs.AddNew
rs.Fields("NomClient").Value = "Dupont"
rs.Fields("Ville").Value = "Paris"
rs.Update
```

Comme pour la modification, la forme à tableaux est disponible :

```vba
rs.AddNew Array("NomClient", "Ville"), Array("Dupont", "Paris")
```

Après l'`Update`, le nouvel enregistrement devient l'enregistrement courant (avec un curseur keyset ou dynamique). On peut alors **récupérer la valeur d'un champ NuméroAuto** généré automatiquement, simplement en le relisant :

```vba
rs.AddNew
rs.Fields("NomClient").Value = "Dupont"
rs.Update
Debug.Print "Nouvel identifiant : " & rs.Fields("IDClient").Value
```

Pour des scénarios serveur (SQL Server), on privilégie souvent la requête `SELECT @@IDENTITY` afin de récupérer l'identité générée — une approche abordée dans le contexte des bases externes (section 10.9).

---

## Supprimer un enregistrement (`Delete`)

La suppression retire l'enregistrement **courant** d'un seul appel :

```vba
rs.Delete
```

Un piège classique guette ici : **après `Delete`, l'enregistrement courant est invalide**. Tenter de lire ses champs déclenche une erreur. Il faut donc systématiquement se déplacer vers un enregistrement valide après la suppression :

```vba
rs.Delete
rs.MoveNext   ' quitter l'enregistrement supprimé, désormais inaccessible
```

Cette contrainte rend les **suppressions en boucle** délicates. Le schéma sûr consiste à supprimer puis à avancer immédiatement, sans tenter d'accéder à l'enregistrement supprimé :

```vba
Do While Not rs.EOF
    If rs.Fields("Inactif").Value = True Then
        rs.Delete       ' marque l'enregistrement courant comme supprimé
    End If
    rs.MoveNext         ' avance dans tous les cas
Loop
```

Notez toutefois que, pour supprimer un grand nombre d'enregistrements selon un critère, une **requête action SQL** (`DELETE … WHERE …`) est nettement plus efficace que ce parcours — point sur lequel nous revenons plus bas.

---

## Récapitulatif du flux d'écriture : DAO vs ADO

Pour ancrer la différence essentielle, voici les trois opérations d'écriture côte à côte dans les deux technologies.

| Opération | Flux DAO | Flux ADO |
|---|---|---|
| **Modifier** | `rs.Edit` → champs → `rs.Update` | champs → `rs.Update` (**pas d'`Edit`**) |
| **Ajouter** | `rs.AddNew` → champs → `rs.Update` | `rs.AddNew` → champs → `rs.Update` |
| **Supprimer** | `rs.Delete` | `rs.Delete` (puis se déplacer) |

L'ajout et la suppression se ressemblent d'une technologie à l'autre. C'est bien la **modification** qui diffère : l'absence de `Edit` en ADO est la seule chose réellement nouvelle à intégrer pour qui maîtrise déjà DAO.

---

## Vérifier les capacités avant d'écrire

Compte tenu des rétrogradations de curseur propres au moteur ACE (section 10.4), il est prudent de **vérifier qu'une opération est supportée** avant de la tenter, via la méthode `Supports` :

```vba
If rs.Supports(adUpdate) Then
    rs.Fields("Ville").Value = "Lyon"
    rs.Update
End If
```

Les constantes utiles sont notamment `adUpdate` (modification), `adAddNew` (ajout) et `adDelete` (suppression). Ce test évite de buter sur une erreur d'exécution lorsque le curseur effectivement obtenu ne permet pas l'écriture.

---

## Un avantage des recordsets : des valeurs typées sans formatage

Travailler par recordset présente un atout souvent sous-estimé face à la construction de SQL dynamique. Lorsqu'on affecte une valeur à un champ, on lui transmet une **valeur VBA native** — un nombre, une date, un booléen — **sans aucun formatage** :

```vba
rs.Fields("DateCommande").Value = Date          ' date native, aucun souci de format
rs.Fields("Montant").Value = 1234.56            ' décimale native, pas de séparateur à gérer
```

On échappe ainsi aux pièges classiques du SQL dynamique : dates à formater en US/ISO, séparateur décimal, échappement des apostrophes (problèmes traités au chapitre 11.7). Pour des mises à jour ponctuelles d'enregistrements précis, le recordset est donc souvent plus sûr qu'une requête construite à la main.

---

## Recordset ou SQL ? La question de la performance

Le confort du recordset a une contrepartie : il opère **enregistrement par enregistrement**. Pour modifier ou supprimer des milliers de lignes selon un critère, parcourir un recordset est lent — chaque tour de boucle implique un aller-retour avec le moteur.

Dans ces cas, une **requête action ensembliste** exécutée directement est radicalement plus rapide :

```vba
' Bien plus efficace qu'une boucle pour une mise à jour de masse
cn.Execute "UPDATE Clients SET Actif = False WHERE DerniereCommande < #2024-01-01#;", , adCmdText
```

La règle pratique : **réservez le recordset au traitement ligne à ligne** (lecture séquentielle, modification d'enregistrements ciblés, logique conditionnelle complexe) et **privilégiez le SQL ensembliste** pour les opérations de masse. L'exécution de SQL est détaillée au chapitre 11, l'objet `Command` paramétré en section 10.6, et les considérations de performance au chapitre 18.

---

## Refermer le recordset

Comme pour la connexion (section 10.3), un recordset se referme et se libère explicitement une fois le travail terminé :

```vba
rs.Close
Set rs = Nothing
```

On y associe les mêmes précautions : tester l'état avant de fermer (`If rs.State = adStateOpen Then rs.Close`) au sein d'une structure de gestion d'erreurs garantissant la libération en toute circonstance.

---

## En résumé

Les quatre opérations CRUD d'ADO reposent sur un petit ensemble de gestes. La **lecture** s'appuie sur le test `BOF And EOF` puis sur une boucle `Do While Not rs.EOF` ponctuée de `MoveNext`, avec un soin particulier pour les valeurs `Null` (via `Nz`). La **modification** se fait *sans `Edit`* : on affecte les champs puis on appelle `Update` — c'est la seule réelle nouveauté par rapport à DAO. L'**ajout** suit le trio `AddNew` → champs → `Update`, et la **suppression** un `Delete` impérativement suivi d'un déplacement. Vérifier les capacités via `Supports` prémunit contre les limites du curseur ACE. Enfin, le recordset excelle pour le traitement ligne à ligne et l'écriture de valeurs typées sans formatage, tandis que le SQL ensembliste reste le bon choix pour les opérations de masse.

Jusqu'ici, nos commandes étaient passées sous forme de simples chaînes SQL. Mais cette approche montre vite ses limites dès que des valeurs variables entrent en jeu — risques d'injection, problèmes de formatage. La section suivante introduit l'objet conçu pour y répondre proprement : l'objet `Command` et ses paramètres.


⏭️ [10.6. Objet Command et paramètres de requête](/10-ado-access/06-objet-command-parametres.md)
