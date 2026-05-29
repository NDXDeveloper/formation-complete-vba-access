🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.3. Navigation et déplacement entre enregistrements

Après l'ouverture des objets, voici les méthodes de `DoCmd` qui permettent de **se déplacer entre les enregistrements** — passer au suivant, aller au premier, créer un nouvel enregistrement, rechercher une valeur. Mais avant tout, une distinction essentielle doit être posée, car elle est source de confusions fréquentes.

## Une distinction essentielle d'emblée

Il existe **deux choses très différentes** que l'on appelle « naviguer entre les enregistrements » :

1. **Déplacer l'enregistrement courant d'un formulaire** (ou d'une feuille de données) — celui que **l'utilisateur voit**. C'est ce que font les méthodes de `DoCmd` présentées ici. On agit sur l'**interface**.
2. **Parcourir un recordset par code** — un curseur de données en mémoire, avec `MoveNext`, `MoveFirst`, `EOF`, `BOF`… C'est l'objet du **chapitre 9** (section 9.4). On agit sur les **données**, indépendamment de l'affichage.

Cette section ne traite **que la première**. Confondre les deux est un classique : `DoCmd.GoToRecord` déplace l'enregistrement **affiché** d'un formulaire ; il ne « parcourt » pas des données en arrière-plan. Le pont entre les deux mondes existe — c'est le **`RecordsetClone`** d'un formulaire (section 9.12) —, mais il s'agit là d'un mécanisme distinct.

## GoToRecord — se déplacer entre les enregistrements

```
DoCmd.GoToRecord [ObjectType], [ObjectName], [Record], [Offset]
```

L'argument **`Record`** indique où aller ; en l'absence d'`ObjectType`/`ObjectName`, l'action porte sur l'objet **actif**.

| Constante | Déplacement |
|---|---|
| `acFirst` | premier enregistrement |
| `acLast` | dernier enregistrement |
| `acNext` | suivant (de *Offset* enregistrements) |
| `acPrevious` | précédent |
| `acGoTo` | l'enregistrement n° *Offset* |
| `acNewRec` | le **nouvel** enregistrement (vierge) |

```vba
DoCmd.GoToRecord , , acNewRec        ' aller au nouvel enregistrement (bouton « Ajouter »)
DoCmd.GoToRecord , , acNext          ' enregistrement suivant
DoCmd.GoToRecord , , acPrevious      ' précédent
DoCmd.GoToRecord , , acFirst         ' premier
DoCmd.GoToRecord , , acLast          ' dernier
DoCmd.GoToRecord , , acGoTo, 5       ' aller à l'enregistrement n° 5
```

L'usage le plus fréquent est **`acNewRec`**, derrière un bouton « Nouveau » ou « Ajouter » (il suppose que le formulaire **autorise les ajouts**).

**Piège à connaître** : tenter d'aller **au-delà du dernier** enregistrement (ou avant le premier) déclenche l'**erreur 2105** (« Impossible d'atteindre l'enregistrement spécifié »). Il faut donc souvent l'anticiper :

```vba
On Error Resume Next
DoCmd.GoToRecord , , acNext
If Err.Number = 2105 Then
    MsgBox "Vous êtes déjà au dernier enregistrement."
End If
On Error GoTo 0
```

## GoToControl — donner le focus à un contrôle

```
DoCmd.GoToControl ControlName
```

`GoToControl` déplace le **focus** vers un contrôle de l'objet actif. Le contrôle doit être **visible** et **activé**, faute de quoi une erreur survient.

```vba
DoCmd.GoToControl "txtNom"           ' donne le focus au contrôle txtNom
```

Notons d'ores et déjà qu'il existe une alternative plus **directe**, via le modèle objet — la méthode `SetFocus` du contrôle —, souvent préférée dans le code du formulaire lui-même (section 5.8) :

```vba
Me.txtNom.SetFocus                   ' équivalent plus direct depuis le formulaire
```

## GoToPage — pages d'un formulaire multi-pages

```
DoCmd.GoToPage PageNumber, [Right], [Down]
```

`GoToPage` fait défiler un formulaire comportant des **sauts de page** jusqu'à la page indiquée. Cet usage est devenu rare : on lui préfère aujourd'hui le **contrôle onglet** (`TabControl`, section 8.4) pour organiser un formulaire en plusieurs pages.

## FindRecord et FindNext — rechercher une valeur

```
DoCmd.FindRecord FindWhat, [Match], [MatchCase], [Search], [SearchAsFormatted], [OnlyCurrentField], [FindFirst]
```

`FindRecord` est l'**équivalent programmatique de Ctrl+F** : il recherche une valeur **dans les enregistrements affichés** et positionne le formulaire sur le premier qui correspond. Il opère sur le **contrôle actif** (ou l'ensemble du formulaire selon `OnlyCurrentField`).

```vba
' Rechercher "Dupont" dans le contrôle actif
Me.txtNom.SetFocus
DoCmd.FindRecord "Dupont", acEntire, , acSearchAll, , acCurrent
DoCmd.FindNext                       ' occurrence suivante
```

Les arguments les plus utiles : `Match` (`acEntire`, `acStart`, `acAnywhere`), `Search` (`acSearchAll`, `acUp`, `acDown`) et `OnlyCurrentField` (`acCurrent`, `acAll`). `FindNext` répète la dernière recherche.

**À ne pas confondre**, là encore, avec la recherche **dans les données** : pour localiser un enregistrement par code dans un recordset (sans passer par l'affichage), on utilise les méthodes `FindFirst`, `FindNext` ou `Seek` **d'un recordset** (section 9.9), qui n'ont rien à voir avec le `FindRecord` de `DoCmd`.

## Naviguer dans les données ou dans le formulaire : récapitulatif

| Besoin | Outil | Référence |
|---|---|---|
| Déplacer l'enregistrement **affiché** d'un formulaire | `DoCmd.GoToRecord` | cette section |
| Rechercher dans les enregistrements **affichés** | `DoCmd.FindRecord` / `FindNext` | cette section |
| Parcourir un **recordset** par code | `MoveNext`, `MoveFirst`, `EOF`… | section 9.4 |
| Rechercher dans un **recordset** par code | `FindFirst`, `Seek`… | section 9.9 |
| Synchroniser le formulaire et un recordset | `RecordsetClone` | section 9.12 |

## À retenir

- Distinguez la navigation **dans le formulaire** (enregistrement affiché, via `DoCmd` — cette section) de la navigation **dans un recordset** (curseur de données, via `MoveNext`… — section 9.4).
- **`GoToRecord`** déplace l'enregistrement courant (`acFirst`, `acLast`, `acNext`, `acPrevious`, `acGoTo`, **`acNewRec`**) ; aller au-delà des bornes lève l'**erreur 2105**.
- **`acNewRec`** est l'usage typique du bouton « Nouveau / Ajouter » (sous réserve que le formulaire autorise les ajouts).
- **`GoToControl`** donne le focus à un contrôle ; l'alternative directe `Me.contrôle.SetFocus` lui est souvent préférée (section 5.8).
- **`FindRecord`** / **`FindNext`** = équivalent de Ctrl+F (recherche dans l'affichage), **distinct** des méthodes de recherche d'un recordset (`FindFirst`, `Seek`, section 9.9).

---


⏭️ [5.4. Exécution de requêtes (RunSQL, OpenQuery)](/05-objet-docmd/04-execution-requetes.md)
