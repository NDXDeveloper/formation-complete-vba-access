🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.2. Contrôles texte, liste déroulante et liste (TextBox, ComboBox, ListBox)

Trois contrôles concentrent l'essentiel de la saisie et de la sélection dans une application Access : la **zone de texte** (saisie libre), la **liste déroulante** (choix dans une liste, avec saisie optionnelle) et la **liste** (choix dans une liste toujours visible, simple ou multiple). Cette section détaille leurs propriétés clés, leur manipulation par code et les pièges récurrents — en particulier l'indexation des colonnes et la lecture des sélections multiples.

L'accès à ces contrôles repose sur la hiérarchie vue au chapitre 8.1 ; leur liaison aux données est traitée au chapitre 8.6.

---

## 8.2.1. La zone de texte (TextBox)

La zone de texte est le contrôle le plus courant : saisie et affichage de texte, de nombres ou de dates. Ses propriétés principales :

| Propriété | Rôle |
|---|---|
| `Value` | Contenu validé du contrôle |
| `Text` | Texte en cours d'édition (uniquement si le contrôle a le focus) |
| `ControlSource` | Champ lié, ou expression ; vide si indépendant (chapitre 8.6) |
| `Format`, `InputMask`, `DecimalPlaces` | Mise en forme et masque de saisie (chapitre 8.10) |
| `Enabled`, `Locked` | Activation et verrouillage |
| `DefaultValue` | Valeur par défaut à la création d'un enregistrement |

```vba
Me.txtNom.Value = "Dupont"
Me.txtMontant.Enabled = False
```

### Value et Text : une distinction importante

- **`Value`** est la valeur **validée**, accessible à tout moment.
- **`Text`** est le texte du **tampon d'édition**, accessible **uniquement** lorsque le contrôle a le focus. Y accéder hors focus déclenche une erreur.

Cette distinction est cruciale dans l'événement `Change` (chapitre 8.9), où l'on veut lire la **saisie en cours** — donc `Text`, et non `Value` (qui ne sera mise à jour qu'à la validation) :

```vba
Private Sub txtRecherche_Change()
    Debug.Print Me.txtRecherche.Text   ' saisie en cours, avant validation
End Sub
```

---

## 8.2.2. La liste déroulante (ComboBox)

La liste déroulante propose un choix dans une liste, avec possibilité de **saisir** une valeur. Outre les propriétés de liste communes (section 8.2.4), elle expose :

| Propriété | Rôle |
|---|---|
| `LimitToList` | Restreint la saisie aux éléments de la liste |
| `Text` | Texte affiché |
| `Value` | Valeur de la **colonne liée** de la ligne sélectionnée |

```vba
' Alimenter une liste de valeurs fixes
Me.cboStatut.RowSourceType = "Value List"
Me.cboStatut.RowSource = "Brouillon;Validé;Annulé"
```

L'événement `AfterUpdate` permet de réagir à une sélection, et l'événement `NotInList` de gérer une valeur saisie absente de la liste (section 8.2.9).

---

## 8.2.3. La liste (ListBox)

La liste affiche en permanence ses éléments, sans saisie possible. Elle partage les propriétés de liste de la combo, et ajoute la **sélection multiple** :

| Propriété | Rôle |
|---|---|
| `MultiSelect` | Mode de sélection : `0` Aucune, `1` Simple, `2` Étendue |
| `Value` | Valeur liée de l'élément sélectionné (sélection **simple** uniquement) |
| `ItemsSelected` | Collection des éléments sélectionnés (sélection **multiple**) |
| `Selected(i)` | État de sélection de la ligne `i` (lecture/écriture) |

En sélection **simple**, `Value` donne la valeur liée choisie. En sélection **multiple**, `Value` n'a plus de sens : on parcourt `ItemsSelected` (section 8.2.10).

---

## 8.2.4. Propriétés communes aux deux contrôles de liste

Liste déroulante et liste partagent un socle commun de propriétés :

| Propriété | Rôle |
|---|---|
| `RowSource` / `RowSourceType` | Origine et type de la liste (section 8.2.5) |
| `BoundColumn` | Colonne (indexée à partir de **1**) fournissant la `Value` |
| `ColumnCount` | Nombre de colonnes affichées |
| `ColumnWidths` | Largeurs des colonnes (0 pour masquer une colonne) |
| `ColumnHeads` | Affiche une ligne d'en-têtes |
| `ListCount` / `ListIndex` | Nombre de lignes / indice de la ligne sélectionnée (`-1` si aucune) |
| `Column(col, row)` | Valeur d'une cellule (colonne indexée à partir de **0**) |
| `ItemData(i)` | Valeur liée de la ligne `i` |

---

## 8.2.5. RowSource et RowSourceType

La propriété `RowSourceType` détermine l'interprétation de `RowSource` :

| `RowSourceType` | `RowSource` contient… |
|---|---|
| `"Table/Query"` | Un nom de table ou une instruction SQL |
| `"Value List"` | Une liste de valeurs séparées par des points-virgules |
| `"Field List"` | Un nom de table (la liste affiche ses **champs**) |

On peut définir `RowSource` dynamiquement par code, dans le même esprit que le `RecordSource` d'un formulaire (chapitre 6.5) :

```vba
Me.cboClient.RowSource = "SELECT ClientID, Nom FROM T_Clients ORDER BY Nom"
Me.cboClient.Requery
```

---

## 8.2.6. Afficher une valeur, en stocker une autre

C'est l'usage le plus puissant des contrôles de liste : **afficher** un libellé compréhensible (le nom d'un client) tout en **stockant** un identifiant. On l'obtient en combinant `ColumnCount`, `BoundColumn` et `ColumnWidths` :

```
' En conception :
'   RowSource    : SELECT ClientID, Nom FROM T_Clients ORDER BY Nom
'   ColumnCount  : 2
'   BoundColumn  : 1          (la Value sera l'ID)
'   ColumnWidths : 0cm;4cm    (la 1re colonne, l'ID, est masquée)
```

```vba
Debug.Print Me.cboClient.Value       ' l'identifiant (colonne liée)
Debug.Print Me.cboClient.Column(1)   ' le nom (2e colonne, indice 1)
```

L'utilisateur voit le nom ; le code et la base manipulent l'identifiant.

---

## 8.2.7. Column et l'indexation : un piège classique

Attention à la **double convention** d'indexation, source d'erreurs fréquentes :

- **`BoundColumn`** est indexée à partir de **1** (la première colonne est `1`).
- **`Column(index)`** est indexée à partir de **0** (la première colonne est `0`).

Ainsi, pour une liste à deux colonnes `ID; Nom` avec `BoundColumn = 1` :

```vba
Me.cboClient.Value         ' = la colonne 1 au sens BoundColumn (l'ID)
Me.cboClient.Column(0)     ' = la même colonne, l'ID (indice 0)
Me.cboClient.Column(1)     ' = le Nom (indice 1)
```

`Column` accepte aussi un second argument pour viser une ligne précise : `Column(colonne, ligne)`.

---

## 8.2.8. Listes déroulantes en cascade

Une liste peut **dépendre** d'une autre : choisir un pays restreint la liste des villes. On redéfinit le `RowSource` de la seconde liste dans l'événement `AfterUpdate` de la première, puis on la rafraîchit :

```vba
Private Sub cboPays_AfterUpdate()
    Me.cboVille.RowSource = "SELECT VilleID, Nom FROM T_Villes " & _
                            "WHERE PaysID = " & Me.cboPays.Value & _
                            " ORDER BY Nom"
    Me.cboVille.Requery
End Sub
```

On veillera à réinitialiser la seconde liste si le premier choix change, pour éviter une valeur incohérente.

---

## 8.2.9. L'événement NotInList — ajouter à la volée

Lorsque `LimitToList = True` et que l'utilisateur saisit une valeur **absente** de la liste, l'événement `NotInList` se déclenche. Il permet, par exemple, de proposer l'ajout du nouvel élément :

```vba
Private Sub cboCategorie_NotInList(NewData As String, Response As Integer)
    If MsgBox("Ajouter la catégorie « " & NewData & " » ?", _
              vbYesNo + vbQuestion) = vbYes Then
        CurrentDb.Execute "INSERT INTO T_Categories (Nom) VALUES ('" & _
                          Replace(NewData, "'", "''") & "')"
        Response = acDataErrAdded      ' Access ré-interroge la liste
    Else
        Response = acDataErrContinue   ' géré : pas de message d'erreur
    End If
End Sub
```

L'argument `Response` indique à Access la suite : `acDataErrAdded` (l'élément a été ajouté), `acDataErrContinue` (géré, sans message) ou `acDataErrDisplay` (afficher l'erreur standard). La construction sécurisée de l'`INSERT` relève du chapitre 11.5.

---

## 8.2.10. Lire une sélection multiple (ItemsSelected)

Dans une liste en mode `MultiSelect`, on parcourt la collection `ItemsSelected`, dont chaque élément est l'**indice** d'une ligne sélectionnée. La valeur liée s'obtient via `ItemData` :

```vba
Dim varItem As Variant

For Each varItem In Me.lstProduits.ItemsSelected
    Debug.Print Me.lstProduits.ItemData(varItem)        ' valeur liée
    Debug.Print Me.lstProduits.Column(1, varItem)       ' autre colonne de la ligne
Next varItem
```

On peut aussi sélectionner ou désélectionner par code via `Selected` :

```vba
Me.lstProduits.Selected(0) = True    ' sélectionner la première ligne
```

---

## 8.2.11. Rafraîchir une liste (Requery)

Lorsque les données sous-jacentes changent (ajout d'un client, d'une catégorie), la liste **ne se met pas à jour** automatiquement : il faut la **ré-interroger** :

```vba
Me.cboClient.Requery
```

C'est notamment indispensable après un ajout réalisé ailleurs dans l'application, ou après avoir modifié `RowSource` par code.

---

## 8.2.12. Pièges et bonnes pratiques

- **Distinguer `Value` et `Text`** : `Text` n'est lisible qu'avec le focus (utile dans `Change`) ; `Value` est la valeur validée.
- **Vérifier `BoundColumn`** : c'est elle qui détermine la valeur stockée ; une erreur ici stocke la mauvaise colonne.
- **Indexation `Column` à partir de 0**, mais `BoundColumn` à partir de 1 : ne pas confondre.
- **Masquer la colonne d'identifiant** via `ColumnWidths` (largeur 0) pour afficher un libellé tout en stockant l'ID.
- **Rafraîchir avec `Requery`** après toute modification des données sous-jacentes ou du `RowSource`.
- **Sélection multiple** : ne pas lire `Value` (sans objet) ; parcourir `ItemsSelected` avec `ItemData`.
- **Gérer `Null`** : aucune sélection donne `ListIndex = -1` et une `Value` nulle ; protéger les lectures avec `Nz`.
- **`LimitToList` sans `NotInList`** : une valeur inconnue déclenche un message d'erreur standard ; gérer l'événement pour un comportement maîtrisé.

---

## 8.2.13. Récapitulatif

- La **zone de texte** gère saisie et affichage ; sa propriété `Value` est la valeur validée, tandis que `Text` (le tampon d'édition) n'est accessible qu'avec le focus — distinction clé pour l'événement `Change`.
- La **liste déroulante** et la **liste** partagent un socle : `RowSource`/`RowSourceType`, `BoundColumn`, `ColumnCount`, `ColumnWidths`, `Column`, `ListIndex`/`ListCount`.
- Le pattern central consiste à **afficher un libellé en stockant un identifiant** (`ColumnCount`, `BoundColumn`, colonne masquée par `ColumnWidths`), en se méfiant de la **double indexation** (`Column` à partir de 0, `BoundColumn` à partir de 1).
- Les **listes en cascade** se construisent en redéfinissant le `RowSource` dans `AfterUpdate` puis en appelant `Requery` ; l'événement **`NotInList`** permet d'ajouter un élément à la volée.
- La **sélection multiple** d'une liste se lit via la collection **`ItemsSelected`** et `ItemData`, jamais via `Value`.
- Toute modification des données impose un **`Requery`** pour rafraîchir la liste ; les événements de ces contrôles sont détaillés au chapitre 8.9 et leur liaison aux données au chapitre 8.6.

⏭️ [8.3. Contrôles booléens (CheckBox, OptionGroup, ToggleButton)](/08-controles-evenements/03-controles-booleens.md)
