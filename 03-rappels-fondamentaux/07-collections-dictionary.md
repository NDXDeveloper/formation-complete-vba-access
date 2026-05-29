🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.7. Collections et Dictionary (Scripting.Dictionary)

Dernier rappel du chapitre : les **collections** et le **`Dictionary`**, deux conteneurs plus souples que les tableaux pour regrouper des données. Là où un tableau impose un type unique et un accès par indice, ces structures acceptent des éléments variés, se redimensionnent d'elles-mêmes et — pour le `Dictionary` — permettent un accès par **clé**. Bonne nouvelle pour la suite : le modèle objet d'Access est lui-même truffé de collections (chapitre 4), si bien que les maîtriser sert bien au-delà de ce chapitre.

## Au-delà des tableaux

Les tableaux (section 3.5) conviennent aux ensembles de valeurs homogènes, à indice numérique. Mais dès qu'on veut **ajouter et retirer librement** des éléments, mélanger des types, ou retrouver une valeur par une **clé** plutôt que par sa position, les **collections** et le **`Dictionary`** sont plus adaptés. Ils dispensent de gérer soi-même la taille et les décalages d'indices.

## L'objet Collection

L'objet `Collection` est **intégré à VBA** : aucune référence n'est nécessaire. On le crée avec `New`, puis on l'alimente.

```vba
Dim col As New Collection
col.Add "Pomme"                    ' ajout simple (prend l'indice 1)
col.Add "Banane", "B"              ' ajout avec une clé "B"
col.Add "Cerise"

Debug.Print col(1)                 ' "Pomme" (indice à base 1)
Debug.Print col("B")               ' "Banane" (accès par clé)
Debug.Print col.Count              ' 3
col.Remove 3                       ' supprime "Cerise"
```

Ses membres principaux sont `Add` (avec une **clé** optionnelle de type chaîne), `Item` (accès par indice **à base 1** ou par clé), `Remove` et `Count`. On le parcourt avec `For Each` ou par indice :

```vba
Dim element As Variant
For Each element In col
    Debug.Print element
Next element
```

### Les limites de Collection

L'objet `Collection` montre vite ses limites :

- **pas de méthode pour tester l'existence d'une clé** : il faut intercepter l'erreur déclenchée par un accès infructueux ;
- **impossible de lister** les clés ou les éléments ;
- **impossible de modifier** une valeur en place (il faut la retirer puis la rajouter).

La première limite, en particulier, oblige à un contournement peu élégant :

```vba
' Collection n'a pas de méthode Exists : on doit passer par la gestion d'erreurs
Function CleExiste(col As Collection, cle As String) As Boolean
    On Error Resume Next
    col.Item cle
    CleExiste = (Err.Number = 0)
    On Error GoTo 0
End Function
```

## Le Scripting.Dictionary

Le `Dictionary` provient de la bibliothèque **Microsoft Scripting Runtime** (section 2.5). On l'utilise en **liaison précoce** (avec la référence) ou en **liaison tardive** (sans référence, plus robuste au déploiement — section 2.6).

```vba
' Liaison précoce (nécessite la référence Microsoft Scripting Runtime) :
Dim dict As New Scripting.Dictionary

' — ou liaison tardive (sans référence) :
' Dim dict As Object
' Set dict = CreateObject("Scripting.Dictionary")

dict.Add "FR", "France"
dict.Add "BE", "Belgique"
dict("CH") = "Suisse"              ' affecter une clé inexistante l'ajoute

Debug.Print dict("FR")            ' "France"
Debug.Print dict.Exists("BE")     ' True  -> l'atout majeur
Debug.Print dict.Count            ' 3
```

C'est un magasin **clé / valeur** : chaque élément possède une **clé unique** et une valeur associée. Ses membres clés :

- `Add cle, valeur` ajoute une paire ;
- `Item(cle)` lit ou écrit la valeur (et **affecter une clé inexistante la crée**) ;
- **`Exists(cle)`** indique si une clé est présente — c'est son grand avantage sur `Collection` ;
- `Remove(cle)` et `RemoveAll` suppriment ;
- `Keys` et `Items` renvoient des **tableaux** de toutes les clés / valeurs ;
- `CompareMode` règle la **sensibilité à la casse** des clés (à fixer **avant** d'ajouter des éléments).

Le parcours se fait naturellement via les clés :

```vba
Dim cle As Variant
For Each cle In dict.Keys
    Debug.Print cle & " = " & dict(cle)
Next cle
```

> ⚠️ **Piège à connaître.** Lire une clé **inexistante** avec `dict(cle)` ne renvoie pas seulement une valeur vide : cela **crée silencieusement** la clé (avec une valeur `Empty`). Pour vérifier sans rien créer, utilisez toujours **`Exists`** au préalable.

## Collection ou Dictionary ?

| Critère | Collection | Dictionary |
|---|---|---|
| Disponibilité | **intégrée** (aucune référence) | bibliothèque **Scripting** (ou liaison tardive) |
| Accès | indice (base 1) **ou** clé | par **clé** |
| Tester l'existence d'une clé | non (gestion d'erreurs requise) | **`Exists`** |
| Lister les clés / valeurs | non | `Keys` / `Items` |
| Modifier une valeur en place | non (retirer puis rajouter) | `dict(cle) = valeur` |
| Sensibilité à la casse des clés | insensible | réglable (`CompareMode`) |

En pratique : la **`Collection`** suffit pour une **liste simple** que l'on alimente et parcourt, sans dépendance externe. Le **`Dictionary`** est préférable dès qu'on a besoin de **recherches par clé**, de **tests d'existence**, de **mises à jour** ou de **lister** les clés — c'est-à-dire la plupart des cas où l'on raisonne en termes de clé/valeur.

## Côté Access : usages typiques

Dans une application Access, on manipule les **données de la base** avec des recordsets et du SQL ; collections et dictionnaires servent plutôt aux **structures en mémoire**. Quelques usages fréquents :

- **mettre en cache** des correspondances (par exemple un code → un libellé) pour éviter des appels répétés à `DLookup`, au bénéfice des performances (section 18.4) ;
- **dédoublonner** des valeurs (les clés d'un `Dictionary` étant uniques) ;
- **compter des occurrences** ou bâtir un index en mémoire.

Par ailleurs, le **modèle objet d'Access** est constitué de très nombreuses **collections** intégrées — `Forms`, `Controls`, `Fields`, `TableDefs`, `QueryDefs`… — que l'on parcourt précisément avec `For Each` (section 3.2). Comprendre la notion de collection est donc un prérequis direct du **chapitre 4**. Enfin, pour un `Dictionary` destiné à être **déployé** sur des postes variés, la **liaison tardive** évite toute dépendance à la référence Scripting (section 2.6).

## À retenir

- Au-delà des tableaux, deux conteneurs souples : l'objet **`Collection`** (intégré) et le **`Scripting.Dictionary`** (bibliothèque Scripting, section 2.5).
- **`Collection`** : accès par **indice (base 1)** ou clé, mais **pas de test d'existence** (contournement par gestion d'erreurs), ni listing des clés, ni mise à jour en place.
- **`Dictionary`** : magasin **clé/valeur** avec **`Exists`**, `Keys`/`Items`, mise à jour directe et `CompareMode` — généralement préférable pour les données indexées par clé.
- **Piège du `Dictionary`** : lire une clé inexistante avec `dict(cle)` la **crée** ; testez d'abord avec **`Exists`**.
- Sous Access, ces structures servent au **cache** et aux traitements en mémoire ; le **modèle objet** lui-même regorge de **collections** parcourues avec `For Each` (chapitre 4). Pour le déploiement, la **liaison tardive** du `Dictionary` est plus robuste (section 2.6).

---


> ✅ **Fin du chapitre 3.** La suite se poursuit avec le chapitre 4 — Modèle objet Access

⏭️ [4. Modèle objet Access](/04-modele-objet-access/README.md)
