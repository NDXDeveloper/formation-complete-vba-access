🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.1. Rôle central de DoCmd dans l'automatisation Access

Avant de détailler les méthodes de `DoCmd` dans les sections suivantes, posons-en la **logique d'ensemble**. Cette première section répond à trois questions : qu'est-ce que `DoCmd`, comment il fonctionne, et quelle est sa place parmi les outils d'automatisation d'Access. C'est la grille de lecture qui rendra le reste du chapitre limpide.

## Qu'est-ce que DoCmd ?

`DoCmd` est une propriété de l'objet `Application` (chapitre 4) qui renvoie l'objet **`DoCmd`**. Cet objet expose des **méthodes** correspondant aux **actions** que l'on effectue dans Access : ouvrir un formulaire, lancer une requête, imprimer un état, importer des données… En somme, `DoCmd` est l'outil par lequel le code **fait faire** quelque chose à Access.

Comme `Application` est implicite (section 4.2), on l'utilise directement :

```vba
DoCmd.OpenForm "frmClients"
DoCmd.Close acForm, "frmClients"
DoCmd.OpenReport "rptFactures", acViewPreview
```

## La logique : des méthodes qui sont des actions

Le point clé pour comprendre `DoCmd` : ses méthodes **reprennent les actions** que l'on retrouve aussi dans les macros (sections 1.3 et 2.7). À l'action de macro `OuvrirFormulaire` correspond la méthode `DoCmd.OpenForm` ; à `ExécuterSQL`, la méthode `DoCmd.RunSQL` ; et ainsi de suite.

Cette correspondance n'est pas un hasard : `DoCmd` est, historiquement, la **passerelle VBA vers le répertoire d'actions** d'Access. C'est précisément pour cela que **convertir une macro en code** (section 2.7) produit en abondance des appels à `DoCmd`. Qui connaît les actions de macro retrouvera donc, presque à l'identique, les méthodes de `DoCmd`.

## Syntaxe et arguments

L'appel suit la forme `DoCmd.Méthode argument1, argument2, …`. La particularité des méthodes de `DoCmd` est qu'elles comportent souvent **de nombreux arguments**, dont la plupart sont **optionnels** et que l'on laisse à leur valeur par défaut.

C'est typiquement le cas de `OpenForm`, qui possède **sept** arguments (`FormName`, `View`, `FilterName`, `WhereCondition`, `DataMode`, `WindowMode`, `OpenArgs`) alors qu'on n'en renseigne souvent qu'un ou deux. D'où l'intérêt des **arguments nommés** (section 3.3), qui évitent d'aligner des virgules vides pour atteindre celui qui compte :

```vba
' Sans arguments nommés : des virgules vides jusqu'à WhereCondition
DoCmd.OpenForm "frmClients", , , "ClientID = 1"

' Avec arguments nommés : limpide
DoCmd.OpenForm FormName:="frmClients", WhereCondition:="ClientID = 1"
```

Les méthodes de `DoCmd` s'appuient par ailleurs sur des **constantes** Access (`acNormal`, `acViewPreview`, `acForm`, `acFormEdit`…), fournies par la bibliothèque objet d'Access (référencée par défaut, section 2.5).

## La place de DoCmd parmi les outils

Pour situer `DoCmd`, il est utile de distinguer **trois niveaux** pour « faire quelque chose » dans Access :

| Niveau | Exemple | Nature |
|---|---|---|
| **Macros** | action *OuvrirFormulaire* | sans code (section 1.3) |
| **`DoCmd`** | `DoCmd.OpenForm "frmClients"` | objet d'action en VBA |
| **Méthodes objet directes** | `Me.Requery`, `Me.Filter = "…"` | modèle objet (section 5.8) |

`DoCmd` se situe ainsi **entre** le monde des macros (dont il reprend les actions) et le modèle objet « pur » (où l'on manipule directement les propriétés et méthodes des objets). Pour une grande partie des opérations de « commande », il reste l'outil le plus commode. Mais, comme l'annonce ce chapitre, certaines tâches gagnent à passer par des **méthodes objet directes**, plus claires ou plus puissantes — c'est l'objet de la section 5.8.

## Les grandes familles de méthodes

Les méthodes de `DoCmd` se regroupent en quelques familles, chacune traitée dans une section dédiée.

| Famille | Méthodes principales | Section |
|---|---|---|
| Ouverture / fermeture | `OpenForm`, `OpenReport`, `OpenQuery`, `OpenTable`, `Close` | 5.2 |
| Navigation | `GoToRecord`, `GoToControl`, `FindRecord` | 5.3 |
| Requêtes | `RunSQL`, `OpenQuery` | 5.4 |
| Import / export | `TransferDatabase`, `TransferSpreadsheet`, `TransferText` | 5.5 |
| Impression / export | `PrintOut`, `OutputTo` | 5.6 |
| Utilitaires | `SetWarnings`, `Hourglass`, `Beep`, `Quit`, `RunCommand` | 5.7 |

Au-delà de ces familles, citons la méthode **`RunCommand`**, qui exécute une **commande intégrée** d'Access (constantes `acCmd…`) — l'équivalent d'un clic dans les menus. Elle illustre l'ampleur du répertoire couvert par `DoCmd` et est présentée en section 5.7.

## DoCmd et la gestion des erreurs

Dernier point à connaître dès maintenant : les actions de `DoCmd` peuvent **échouer** (objet inexistant, opération impossible) ou être **annulées**. Elles déclenchent alors des erreurs interceptables (chapitre 13).

L'erreur la plus emblématique est la **n° 2501**, « L'action a été annulée ». Elle survient, par exemple, lorsqu'une ouverture est **annulée** dans l'événement `Open` du formulaire (via `Cancel`), ou lorsqu'un export/une impression est interrompu par l'utilisateur.

```vba
' Une action DoCmd peut être annulée (erreur 2501)
On Error Resume Next
DoCmd.OpenForm "frmClients"
If Err.Number = 2501 Then
    ' l'ouverture a été annulée (ex. Cancel = True dans l'événement Open)
End If
On Error GoTo 0
```

Prévoir ce cas évite des messages d'erreur intempestifs lorsqu'une annulation est, en réalité, un comportement normal. La gestion des erreurs fait l'objet du chapitre 13, et l'annulation d'événements avec `Cancel` est traitée en section 8.12.

## À retenir

- **`DoCmd`** est l'objet de l'**action** dans Access : ses **méthodes** exécutent les opérations courantes (ouvrir, fermer, exécuter, importer, imprimer…).
- Ses méthodes **reprennent les actions de macro** (sections 1.3 et 2.7), d'où la production de `DoCmd` lors de la conversion d'une macro.
- Elles comportent souvent **beaucoup d'arguments optionnels** : les **arguments nommés** (section 3.3) en simplifient l'appel ; elles utilisent les **constantes** Access (`acViewPreview`…).
- `DoCmd` se situe **entre** les macros et les méthodes objet directes ; pour certaines tâches, ces dernières sont préférables (section 5.8).
- Les actions de `DoCmd` peuvent être **annulées** (erreur **2501**) : à anticiper dans la gestion des erreurs (chapitres 13 et 8.12).

---


⏭️ [5.2. Ouverture et fermeture d'objets (OpenForm, OpenReport, OpenQuery, OpenTable)](/05-objet-docmd/02-ouverture-fermeture-objets.md)
