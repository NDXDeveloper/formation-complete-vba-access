🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 3.3. Procédures Sub et fonctions Function

Troisième rappel : les **procédures**, qui découpent le code en unités nommées et réutilisables. VBA en distingue deux sortes — les **`Sub`**, qui réalisent des actions, et les **`Function`**, qui réalisent des actions *et* renvoient une valeur. Bien les utiliser, et bien leur passer des arguments, est au cœur de tout code structuré.

## Sub et Function : la différence essentielle

La distinction tient en une phrase : une **`Sub`** exécute des instructions sans rien renvoyer, tandis qu'une **`Function`** exécute des instructions **et retourne une valeur**. Conséquence directe : une `Function` peut s'employer dans une **expression** (on en exploite le résultat), alors qu'une `Sub` ne le peut pas.

Ce point a une portée toute particulière sous Access, comme on l'a vu en section 2.3 : seules les **`Function` `Public` placées dans un module standard** sont appelables depuis une **requête**, une **source de contrôle** ou une macro (action `ExécuterCode`). C'est donc sous forme de fonction que l'on expose une logique réutilisable au moteur d'Access.

## Définir et appeler une procédure Sub

```vba
Sub AfficherBienvenue(nom As String)
    MsgBox "Bonjour " & nom
End Sub
```

L'appel peut se faire de deux manières :

```vba
AfficherBienvenue "Marie"            ' sans Call : pas de parenthèses englobantes
Call AfficherBienvenue("Marie")      ' avec Call : parenthèses obligatoires
```

Le mot-clé `Call` est facultatif. La règle à retenir : **sans `Call`**, on n'entoure pas les arguments de parenthèses ; **avec `Call`**, on les met.

## Définir et appeler une fonction Function

Une fonction se déclare avec un **type de retour** (`As Type`), et l'on affecte son résultat… à **son propre nom** :

```vba
Function CalculerTTC(ht As Currency, taux As Double) As Currency
    CalculerTTC = ht * (1 + taux)    ' le résultat est affecté au nom de la fonction
End Function
```

Pour **récupérer** la valeur renvoyée, on appelle la fonction avec des **parenthèses** :

```vba
Dim prix As Currency
prix = CalculerTTC(100, 0.2)         ' prix vaut 120
```

Une fonction peut aussi renvoyer un **objet**. Dans ce cas, on utilise `Set` des deux côtés — à l'intérieur (affectation du résultat) comme à l'appel :

```vba
Function NouveauRecordset() As DAO.Recordset
    Set NouveauRecordset = CurrentDb.OpenRecordset("tblClients")
End Function

Dim rs As DAO.Recordset
Set rs = NouveauRecordset()          ' Set côté appelant (c'est un objet)
```

## Les arguments

Les arguments (ou paramètres) transmettent des données à une procédure. On les déclare avec leur type, séparés par des virgules. Plusieurs subtilités méritent un rappel.

### Passage par référence (ByRef) ou par valeur (ByVal)

C'est l'un des points les plus mal connus de VBA. Par **défaut**, les arguments sont passés **par référence (`ByRef`)** : la procédure reçoit un accès direct à la variable d'origine et peut donc la **modifier**.

```vba
Sub Doubler(ByRef n As Long)
    n = n * 2
End Sub

Dim valeur As Long
valeur = 5
Doubler valeur          ' ByRef : valeur vaut désormais 10
```

À l'inverse, **`ByVal`** transmet une **copie** : la variable d'origine est protégée.

```vba
Sub Doubler(ByVal n As Long)   ' la copie est modifiée, pas l'original
    n = n * 2
End Sub
' valeur resterait 5
```

Le fait que `ByRef` soit le comportement par défaut surprend souvent et provoque des effets de bord involontaires. La bonne pratique : **être explicite**, et préférer **`ByVal`** lorsqu'on n'a pas l'intention de modifier l'argument.

### Arguments optionnels

Un argument peut être rendu **facultatif** avec `Optional`, en lui donnant une valeur par défaut. Les arguments optionnels doivent figurer **après** les obligatoires.

```vba
Function Saluer(nom As String, Optional politesse As String = "Bonjour") As String
    Saluer = politesse & " " & nom
End Function

Saluer "Marie"                  ' "Bonjour Marie"
Saluer "Marie", "Bonsoir"       ' "Bonsoir Marie"
```

Pour un argument optionnel de type `Variant` sans valeur par défaut, la fonction `IsMissing` permet de tester s'il a été fourni.

### Arguments nommés

Plutôt que de respecter l'ordre des paramètres, on peut les **nommer** à l'appel, avec la syntaxe `nom:=valeur`. C'est particulièrement utile sous Access pour les méthodes comportant de **nombreux arguments optionnels**, comme celles de `DoCmd` : on ne renseigne que ceux qui comptent, sans aligner des virgules vides.

```vba
' Sans arguments nommés, il faudrait des virgules vides pour atteindre WhereCondition :
DoCmd.OpenForm "frmClients", , , "ClientID = 1"

' Avec arguments nommés, l'intention est limpide :
DoCmd.OpenForm FormName:="frmClients", WhereCondition:="ClientID = 1"
```

### Nombre variable d'arguments (ParamArray)

Enfin, `ParamArray` autorise un nombre **variable** d'arguments, reçus sous forme de tableau de `Variant`. Il doit être le **dernier** paramètre de la procédure. Usage plus rare, mais commode pour des fonctions du type « somme d'un nombre quelconque de valeurs ».

## Renvoyer un résultat

Pour produire une valeur en sortie, deux approches :

- une **`Function`** renvoie sa valeur via son nom (le cas standard) ;
- une **`Sub`** peut « renvoyer » des résultats par l'intermédiaire de ses arguments **`ByRef`** (on modifie la variable de l'appelant) — pratique quand il y a plusieurs valeurs à retourner.

## Portée : Public ou Private

Comme les variables, les procédures ont une **portée**. Dans un module standard, une procédure est **`Public`** par défaut (appelable de partout) ; la déclarer **`Private`** la restreint à son module. C'est précisément ce qui détermine qu'une fonction soit, ou non, accessible depuis une requête (section 2.3). La portée est approfondie à la section suivante (3.4).

## Sortie anticipée

On quitte une procédure avant la fin avec **`Exit Sub`** ou **`Exit Function`** (rappel de la section 3.2). Dans une fonction, on prendra soin d'avoir affecté la valeur de retour avant de sortir.

## La récursivité

Une fonction peut **s'appeler elle-même** : c'est la récursivité, utile pour les traitements naturellement récursifs (arborescences, calculs comme la factorielle).

```vba
Function Factorielle(n As Long) As Double
    If n <= 1 Then
        Factorielle = 1
    Else
        Factorielle = n * Factorielle(n - 1)   ' appel récursif
    End If
End Function
```

Elle suppose toujours une **condition d'arrêt** ; sans elle, l'enchaînement d'appels finit par saturer la pile et provoque une erreur.

## Et les propriétés (Property)

Il existe enfin des procédures spéciales — **`Property Get`**, **`Property Let`** et **`Property Set`** — qui définissent les propriétés des objets dans les **modules de classe**. Elles relèvent de la programmation orientée objet et sont traitées en section 16.2.

## À retenir

- Une **`Sub`** exécute des actions sans renvoyer de valeur ; une **`Function`** en renvoie une, affectée à **son propre nom**, et utilisable dans une expression.
- **Rappel Access** : seules les **`Function` `Public` d'un module standard** sont appelables depuis une **requête**, une source de contrôle ou une macro (section 2.3).
- Les arguments sont passés **`ByRef` par défaut** (la procédure peut modifier l'original) ; soyez explicite et préférez **`ByVal`** quand vous ne modifiez pas l'argument.
- Utilisez **`Optional`** (avec valeur par défaut) et surtout les **arguments nommés** (`nom:=valeur`) pour les méthodes à nombreux paramètres comme celles de `DoCmd`.
- Une fonction peut renvoyer un **objet** (`Set` côté appel) ; une `Sub` peut produire des sorties via ses arguments **`ByRef`**.
- La **portée** (`Public`/`Private`) conditionne l'accessibilité (section 3.4) ; les **propriétés** (`Property Get/Let/Set`) appartiennent aux modules de classe (section 16.2).

---


⏭️ [3.4. Portée des variables et durée de vie](/03-rappels-fondamentaux/04-portee-variables.md)
