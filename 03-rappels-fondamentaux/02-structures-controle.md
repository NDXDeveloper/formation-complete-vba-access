🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.2. Structures de contrôle (If, Select Case, boucles)

Deuxième rappel : les **structures de contrôle**, qui pilotent le déroulement du code — exécuter (ou non) un bloc selon une condition, répéter un traitement un certain nombre de fois, ou tant qu'une condition tient. Ce sont exactement les mêmes qu'en VBA Excel ; cette section les passe en revue rapidement, en mettant en lumière les usages typiques sous Access.

## Les structures conditionnelles

### If...Then...Else

La structure de base. Elle existe en version **sur une ligne** (pour un cas simple) et en version **bloc** (pour plusieurs instructions ou plusieurs cas) :

```vba
' Sur une ligne
If montant > 1000 Then remise = 0.1

' En bloc, avec ElseIf et Else
If montant > 1000 Then
    remise = 0.1
ElseIf montant > 500 Then
    remise = 0.05
Else
    remise = 0
End If
```

Notez que **`ElseIf`** s'écrit en un seul mot, et que la version bloc se ferme par `End If`.

### La fonction IIf

`IIf` (*Immediate If*) est une **fonction** qui renvoie une valeur selon une condition — pratique en une ligne, et très utilisée sous Access dans les requêtes et les sources de contrôle :

```vba
remise = IIf(montant > 1000, 0.1, 0)
```

Mais elle recèle un **piège majeur** : `IIf` **évalue ses deux branches**, y compris celle qui n'est pas retenue. Une protection apparente peut donc échouer :

```vba
' ⚠️ total / diviseur est calculé MÊME quand diviseur vaut 0 -> erreur possible
resultat = IIf(diviseur = 0, 0, total / diviseur)
```

Quand l'une des branches risque de provoquer une erreur (division par zéro, accès à `Null`…), préférez un `If` bloc, qui, lui, n'exécute que la branche concernée.

### Select Case

Lorsqu'une même valeur doit être comparée à de nombreux cas, `Select Case` est plus lisible qu'une longue cascade de `ElseIf` :

```vba
Select Case codeStatut
    Case 0
        libelle = "Brouillon"
    Case 1, 2                 ' plusieurs valeurs
        libelle = "En cours"
    Case 3 To 5               ' une plage
        libelle = "Terminé"
    Case Is > 5               ' une comparaison
        libelle = "Archivé"
    Case Else                 ' tous les autres cas
        libelle = "Inconnu"
End Select
```

Il accepte des **valeurs multiples** (séparées par des virgules), des **plages** (`To`) et des **comparaisons** (`Is`). Pour un test inline à plusieurs branches dans une expression, VBA propose aussi la fonction `Switch` — mais, comme `IIf`, elle évalue l'ensemble de ses arguments.

## Les boucles

### For...Next

La boucle à **compteur**, quand le nombre d'itérations est connu :

```vba
For i = 1 To 10
    Debug.Print i
Next i

For i = 10 To 1 Step -1       ' pas négatif : décompte
    Debug.Print i
Next i
```

Le mot-clé `Step` fixe l'incrément (positif, négatif, ou différent de 1).

### For Each...Next

Pour parcourir tous les éléments d'une **collection** ou d'un **tableau**, sans gérer d'indice. C'est un outil précieux sous Access, où l'on itère souvent sur des collections du modèle objet : contrôles d'un formulaire, champs d'un recordset, formulaires ouverts, `TableDefs`…

```vba
' Lister les contrôles d'un formulaire
Dim ctl As Control
For Each ctl In Me.Controls
    Debug.Print ctl.Name
Next ctl
```

### Do...Loop

La boucle la plus souple : elle répète **tant qu'une condition est vraie** (`While`) ou **jusqu'à ce qu'elle le devienne** (`Until`), le test pouvant se placer **en début** (la boucle peut ne jamais s'exécuter) ou **en fin** (elle s'exécute au moins une fois).

```vba
' Test en début
Do While i < 10
    i = i + 1
Loop

Do Until i >= 10
    i = i + 1
Loop

' Test en fin (au moins une itération)
Do
    i = i + 1
Loop While i < 10
```

C'est avec `Do...Loop` que s'écrit le **motif de boucle le plus courant sous Access** : le parcours d'un **recordset**, du premier au dernier enregistrement.

```vba
' Parcours d'un recordset (DAO) — voir chapitre 9
Do While Not rs.EOF
    Debug.Print rs!Nom
    rs.MoveNext          ' ⚠️ indispensable : sans MoveNext, boucle infinie !
Loop
```

L'appel à `rs.MoveNext` est **vital** : c'est lui qui fait avancer le curseur. L'oublier fige le programme dans une boucle sans fin. La navigation dans les recordsets est détaillée en section 9.4.

### While...Wend

Une forme plus ancienne, encore valide mais **moins souple** (elle ne gère pas de sortie anticipée propre) :

```vba
While i < 10
    i = i + 1
Wend
```

À niveau de fonctionnalité égal, `Do...Loop` lui est aujourd'hui préféré.

## Sortir prématurément

On peut interrompre une boucle ou une procédure avant son terme :

- **`Exit For`** et **`Exit Do`** quittent la boucle en cours ;
- **`Exit Sub`** et **`Exit Function`** quittent la procédure.

```vba
For i = 1 To 1000
    If trouve Then Exit For    ' inutile de continuer
Next i
```

## Le cas de GoTo

L'instruction `GoTo`, qui transfère l'exécution vers une étiquette, existe en VBA mais doit être **évitée** pour le contrôle du flux : elle conduit vite à un code illisible. Son **seul usage recommandé** est la **gestion des erreurs** (`On Error GoTo …`), traitée au chapitre 13.

## Attention aux boucles infinies

Une boucle dont la condition de sortie n'est jamais atteinte — compteur non incrémenté, `rs.MoveNext` oublié, condition mal écrite — bloque l'application. Pour **interrompre** un code parti en boucle, le raccourci **<kbd>Ctrl</kbd>+<kbd>Pause</kbd>** (*Ctrl+Break*) stoppe l'exécution ; il fait partie des touches spéciales d'Access évoquées en section 1.4 (et ne fonctionne donc que si celles-ci sont actives).

## À retenir

- `If…Then…Else` (avec `ElseIf`) pour les conditions ; `Select Case` quand une même valeur est confrontée à de nombreux cas (valeurs multiples, plages `To`, comparaisons `Is`).
- **`IIf` évalue ses deux branches** : ne l'utilisez pas pour « protéger » d'une erreur ; un `If` bloc est alors plus sûr.
- `For…Next` pour un nombre connu d'itérations (avec `Step`), `For Each` pour parcourir une **collection** — fréquent sous Access (contrôles, champs, formulaires…).
- `Do…Loop` est la boucle la plus souple ; le motif **`Do While Not rs.EOF … rs.MoveNext … Loop`** est le parcours de recordset typique (chapitre 9) — **ne jamais oublier `MoveNext`**.
- Sortez d'une boucle ou d'une procédure avec `Exit For` / `Exit Do` / `Exit Sub` / `Exit Function` ; réservez **`GoTo`** à la gestion d'erreurs (chapitre 13).
- En cas de boucle infinie, **<kbd>Ctrl</kbd>+<kbd>Pause</kbd>** interrompt l'exécution (touches spéciales, section 1.4).

---


⏭️ [3.3. Procédures Sub et fonctions Function](/03-rappels-fondamentaux/03-procedures-fonctions.md)
