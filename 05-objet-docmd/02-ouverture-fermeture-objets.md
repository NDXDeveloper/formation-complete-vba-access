🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.2. Ouverture et fermeture d'objets (OpenForm, OpenReport, OpenQuery, OpenTable)

Voici la famille de méthodes la plus utilisée de `DoCmd` : celles qui **ouvrent** et **ferment** les objets de l'application. Cette section détaille `OpenForm`, `OpenReport`, `OpenQuery`, `OpenTable` et `Close`, avec leurs arguments essentiels et quelques pièges à éviter — dont un, sur les états, particulièrement redoutable.

## OpenForm — ouvrir un formulaire

C'est sans doute la méthode la plus employée de tout `DoCmd`. Sa syntaxe :

```
DoCmd.OpenForm FormName, [View], [FilterName], [WhereCondition], [DataMode], [WindowMode], [OpenArgs]
```

| Argument | Rôle | Valeurs courantes |
|---|---|---|
| `FormName` | nom du formulaire (**obligatoire**) | — |
| `View` | mode d'affichage | `acNormal`, `acDesign`, `acFormDS` |
| `FilterName` | filtre nommé | — |
| `WhereCondition` | clause WHERE (sans le mot « WHERE ») | `"ClientID = 5"` |
| `DataMode` | mode de saisie | `acFormEdit`, `acFormAdd`, `acFormReadOnly` |
| `WindowMode` | mode de fenêtre | `acWindowNormal`, `acHidden`, `acDialog` |
| `OpenArgs` | paramètre transmis au formulaire | une chaîne / valeur |

Trois arguments méritent une attention particulière.

```vba
DoCmd.OpenForm "frmClients"                                   ' ouverture simple
DoCmd.OpenForm "frmClients", WhereCondition:="ClientID = 5"   ' filtré sur un client
DoCmd.OpenForm "frmClients", DataMode:=acFormReadOnly         ' en lecture seule
DoCmd.OpenForm "frmClients", WindowMode:=acHidden             ' ouvert masqué
```

**`WhereCondition`** est l'argument vedette : il filtre le formulaire sur les enregistrements voulus, au moyen d'une clause WHERE fournie **sans** le mot-clé `WHERE`. Comme il s'agit d'une chaîne, gare aux **apostrophes** et aux **dates** (sections 3.6 et 11.7).

**`acDialog`** ouvre le formulaire en **boîte de dialogue modale** : non seulement il bloque l'accès au reste de l'application, mais surtout il **met le code en pause** — l'instruction suivante ne s'exécute qu'une fois le formulaire **fermé ou masqué**. C'est le mécanisme des boîtes de dialogue personnalisées (section 6.7).

```vba
' Ouverture en dialogue modal : le code attend ici
DoCmd.OpenForm "frmParametres", WindowMode:=acDialog
' ... la suite ne s'exécute qu'après fermeture/masquage de frmParametres
```

**`OpenArgs`** transmet un paramètre au formulaire à son ouverture ; côté formulaire, on le lit via `Me.OpenArgs` (section 6.6).

```vba
DoCmd.OpenForm "frmDetail", OpenArgs:="123"
```

## OpenReport — ouvrir un état

```
DoCmd.OpenReport ReportName, [View], [FilterName], [WhereCondition], [WindowMode], [OpenArgs]
```

> ⚠️ **Le piège à connaître absolument.** Si l'on omet l'argument `View` (ou qu'on passe `acViewNormal`), l'état part **directement à l'imprimante** ! Pour l'afficher à l'écran, il **faut** utiliser **`acViewPreview`**.

```vba
DoCmd.OpenReport "rptFactures"                    ' ⚠️ IMPRIME immédiatement !
DoCmd.OpenReport "rptFactures", acViewPreview     ' aperçu à l'écran (le cas courant)
```

L'argument **`WhereCondition`** filtre les données de l'état, ce qui est la base des **états paramétrés** (section 7.10) :

```vba
DoCmd.OpenReport "rptFactures", acViewPreview, , "Annee = 2025"
```

À noter : si le `WhereCondition` ne ramène **aucun** enregistrement, l'événement **`NoData`** de l'état se déclenche, ce qui permet d'éviter d'afficher un état vide (section 7.9).

## OpenQuery — ouvrir (ou exécuter) une requête

```
DoCmd.OpenQuery QueryName, [View], [DataMode]
```

Le comportement dépend du **type** de requête :

- pour une requête **SELECT**, `OpenQuery` ouvre la **feuille de données** affichant les résultats ;
- pour une requête **action** (INSERT, UPDATE, DELETE, création de table), `OpenQuery` **l'exécute** — en affichant, par défaut, des messages de confirmation.

```vba
DoCmd.OpenQuery "qryClientsActifs"     ' SELECT : affiche les résultats
DoCmd.OpenQuery "qryArchiver"          ' requête action : l'ouvrir l'EXÉCUTE
```

Pour **exécuter** une requête action par code, on lui préfère souvent la méthode **`db.Execute`** de DAO (chapitre 11) : elle n'affiche pas de messages de confirmation et signale mieux les erreurs. Ce choix est approfondi en section 5.4, et comparé en détail en section 18.3.

## OpenTable — ouvrir une table

```
DoCmd.OpenTable TableName, [View], [DataMode]
```

`OpenTable` ouvre une table en **feuille de données**. En pratique, on l'utilise **rarement en production** — on ne montre généralement pas les tables brutes à l'utilisateur, qui passe par des formulaires (chapitre 6). La méthode reste utile pour des outils de **développement** ou d'**administration**.

```vba
DoCmd.OpenTable "tblClients"           ' ouvre la table (usage surtout en développement)
```

## Close — fermer un objet

```
DoCmd.Close [ObjectType], [ObjectName], [Save]
```

| Argument | Rôle | Valeurs |
|---|---|---|
| `ObjectType` | type d'objet | `acForm`, `acReport`, `acQuery`, `acTable`… |
| `ObjectName` | nom de l'objet | — |
| `Save` | enregistrement des modifications | `acSaveYes`, `acSaveNo`, `acSavePrompt` (défaut) |

Deux écritures principales :

```vba
DoCmd.Close                            ' ferme l'objet ACTIF
DoCmd.Close acForm, "frmClients"       ' ferme un formulaire précis
```

**Attention à `DoCmd.Close` sans argument** : il ferme l'objet **actif**, c'est-à-dire celui qui a le focus — ce qui n'est pas toujours celui que l'on croit, surtout si le focus a changé. Pour fermer un formulaire **depuis son propre code**, mieux vaut être explicite :

```vba
DoCmd.Close acForm, Me.Name            ' un formulaire se ferme lui-même, sans ambiguïté
```

Enfin, l'argument **`Save`** concerne les **modifications de conception** de l'objet (et non les **données**, qui sont, elles, déjà enregistrées automatiquement — rappel de la section 1.2). Passer `acSaveNo` évite l'invite « Enregistrer les modifications de conception ? » :

```vba
DoCmd.Close acForm, "frmClients", acSaveNo   ' ferme sans demander pour la conception
```

## Points communs et pièges

- Toutes ces méthodes peuvent déclencher l'**erreur 2501** si l'action est **annulée** (ouverture annulée dans l'événement `Open`, état sans données avec `NoData`, etc.) — voir section 5.1 et chapitre 13.
- Les chaînes `WhereCondition` exigent les mêmes précautions que tout SQL dynamique : **apostrophes** et **dates** au bon format (sections 3.6 et 11.7).

## À retenir

- **`OpenForm`** : argument vedette **`WhereCondition`** pour filtrer ; **`acDialog`** ouvre en modal et **met le code en pause** ; **`OpenArgs`** passe un paramètre (lu via `Me.OpenArgs`, section 6.6).
- **`OpenReport`** : **piège majeur** — sans **`acViewPreview`**, l'état **s'imprime** directement. `WhereCondition` permet les états paramétrés (section 7.10).
- **`OpenQuery`** : une requête **SELECT** affiche ses résultats, une requête **action** s'**exécute** ; pour exécuter par code, **`db.Execute`** est souvent préférable (sections 5.4 et 18.3).
- **`OpenTable`** ouvre une table en feuille de données — surtout utile en **développement**.
- **`Close`** sans argument ferme l'objet **actif** (prudence) ; soyez explicite (`Me.Name`). L'argument **`Save`** vise la **conception**, pas les données.

---


⏭️ [5.3. Navigation et déplacement entre enregistrements](/05-objet-docmd/03-navigation-enregistrements.md)
