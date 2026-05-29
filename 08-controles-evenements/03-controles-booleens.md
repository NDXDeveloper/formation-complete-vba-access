🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.3. Contrôles booléens (CheckBox, OptionGroup, ToggleButton)

Trois contrôles servent à représenter des choix de type « oui/non » ou « l'un parmi plusieurs » : la **case à cocher**, le **bouton bascule** et le **groupe d'options**. Leur point commun apparent — exprimer un choix simple — masque une distinction essentielle : la case à cocher et le bouton bascule renvoient un **booléen**, tandis que le groupe d'options renvoie un **nombre**. Bien comprendre cette différence évite la confusion la plus fréquente avec ces contrôles.

L'accès à ces contrôles suit la hiérarchie du chapitre 8.1 (le groupe d'options est un **conteneur**), et leur liaison aux données relève du chapitre 8.6.

---

## 8.3.1. Trois contrôles, deux usages

| Besoin | Contrôle | Valeur renvoyée |
|---|---|---|
| Un oui/non **isolé** | Case à cocher (`CheckBox`) ou bouton bascule (`ToggleButton`) | Booléen (`True`/`False`) |
| Un choix **exclusif** parmi plusieurs | Groupe d'options (`OptionGroup`) | Un **nombre** (l'`OptionValue`) |

La case à cocher et le bouton bascule sont deux apparences d'une même logique booléenne. Le groupe d'options, lui, relève d'une logique différente, détaillée à partir de la section 8.3.6.

---

## 8.3.2. La représentation des booléens dans Access

En VBA et dans Access, un booléen prend deux valeurs numériques :

- **`True` vaut -1** ;
- **`False` vaut 0**.

Ce point a des conséquences pratiques, notamment dans les critères SQL : pour filtrer un champ Oui/Non, on écrit `WHERE Actif = True` (ou `= -1`, ou `<> 0`), jamais `= 1`. Les champs de type Oui/Non stockent précisément ces valeurs.

---

## 8.3.3. La case à cocher (CheckBox)

La case à cocher exprime un état booléen : cochée (`True`) ou décochée (`False`). Ses propriétés clés :

| Propriété | Rôle |
|---|---|
| `Value` | `True` / `False` (ou `Null` si `TripleState`) |
| `TripleState` | Autorise un troisième état indéterminé (`Null`) — section 8.3.5 |
| `ControlSource` | Champ Oui/Non lié, ou vide si indépendant (chapitre 8.6) |

```vba
If Me.chkActif.Value Then
    ' la case est cochée (True)
End If

Me.chkActif.Value = True
```

Un usage très courant : activer ou désactiver d'autres contrôles selon l'état de la case, dans l'événement `AfterUpdate` (chapitre 8.9) :

```vba
Private Sub chkLivraison_AfterUpdate()
    Me.txtAdresseLivraison.Enabled = Me.chkLivraison.Value
End Sub
```

---

## 8.3.4. Le bouton bascule (ToggleButton)

Le bouton bascule possède exactement la **même sémantique** que la case à cocher (`Value` à `True`/`False`), mais sous l'apparence d'un bouton qui reste **enfoncé** lorsqu'il est actif. On l'utilise pour un rendu plus visuel, ou comme élément d'un groupe d'options.

| Propriété | Rôle |
|---|---|
| `Value` | `True` / `False` |
| `Caption` | Texte affiché sur le bouton |
| `Picture` | Image affichée sur le bouton |

```vba
Me.tglMode.Value = True   ' le bouton apparaît enfoncé
```

---

## 8.3.5. La case à cocher à trois états (TripleState)

Avec `TripleState = True`, la case à cocher accepte un **troisième état**, `Null`, en plus de `True` et `False`. Elle cycle alors entre coché, décoché et indéterminé (grisé). C'est utile pour distinguer « non renseigné » de « non ».

La lecture doit alors traiter explicitement le cas `Null` — et **non** via `Select Case ... Case Null`, qui ne fonctionne pas de façon fiable avec `Null` :

```vba
If IsNull(Me.chkValide.Value) Then
    Debug.Print "Indéterminé"
ElseIf Me.chkValide.Value Then
    Debug.Print "Validé"
Else
    Debug.Print "Refusé"
End If
```

---

## 8.3.6. Le groupe d'options (OptionGroup) : un choix exclusif numérique

Le groupe d'options est un **cadre conteneur** qui regroupe plusieurs options (boutons d'option, cases à cocher ou boutons bascule) et les rend **mutuellement exclusives** : une seule peut être sélectionnée à la fois.

**Point capital** : un groupe d'options ne renvoie **pas** un booléen, mais un **nombre** — la valeur (`OptionValue`) de l'option sélectionnée. Le groupe se lie donc à un champ **numérique**, et chaque option correspond à un entier.

```
' grpStatut : groupe d'options lié à un champ numérique
'   ○ Brouillon   (OptionValue = 1)
'   ○ Validé      (OptionValue = 2)
'   ○ Annulé      (OptionValue = 3)
```

`Me.grpStatut.Value` renvoie alors `1`, `2` ou `3` selon l'option choisie (ou `Null` si aucune).

---

## 8.3.7. La propriété OptionValue

Chaque option placée dans un groupe possède une propriété **`OptionValue`** : l'entier que prend le groupe lorsque cette option est sélectionnée. Access l'attribue automatiquement (1, 2, 3…) au fur et à mesure de l'ajout des options, et on peut la modifier.

Les options elles-mêmes ne sont **pas** liées individuellement à un champ : elles sont identifiées par leur `OptionValue`. Seul le **cadre** (le groupe) porte la liaison via `ControlSource`.

---

## 8.3.8. Lire et réagir à un groupe d'options

On lit la valeur du groupe et on l'interprète, typiquement avec un `Select Case` :

```vba
Select Case Me.grpStatut.Value
    Case 1: Debug.Print "Brouillon"
    Case 2: Debug.Print "Validé"
    Case 3: Debug.Print "Annulé"
    Case Else: Debug.Print "Aucune sélection"   ' Value = Null
End Select
```

On sélectionne une option par code en affectant l'`OptionValue` voulu :

```vba
Me.grpStatut.Value = 2   ' sélectionne l'option « Validé »
```

L'événement `AfterUpdate` du groupe se déclenche à chaque changement de sélection :

```vba
Private Sub grpStatut_AfterUpdate()
    ' Afficher le motif uniquement si « Annulé » est choisi
    Me.txtMotifAnnulation.Visible = (Me.grpStatut.Value = 3)
End Sub
```

---

## 8.3.9. Accéder aux contrôles d'un groupe d'options

Le groupe d'options étant un **conteneur** (chapitre 8.1), ses options ont pour `Parent` le groupe, et non le formulaire. On les parcourt via la collection `Controls` du groupe :

```vba
Dim ctl As Control
For Each ctl In Me.grpStatut.Controls
    Debug.Print ctl.Name, ctl.OptionValue
Next ctl
```

C'est ce qui explique que ces options n'apparaissent pas comme enfants directs du formulaire lors d'une énumération non récursive (section 8.1.8).

---

## 8.3.10. Liaison aux données : Oui/Non ou numérique

Le choix du type de champ lié dépend du contrôle (chapitre 8.6) :

- une **case à cocher** ou un **bouton bascule** se lie à un champ **Oui/Non** ;
- un **groupe d'options** se lie à un champ **numérique** (qui stocke l'`OptionValue`).

```vba
' Filtrer un formulaire sur un champ Oui/Non
Me.Filter = "Actif = True"
Me.FilterOn = True

' Filtrer sur la valeur d'un groupe d'options (numérique)
Me.Filter = "Statut = " & Me.grpStatut.Value
Me.FilterOn = True
```

En cas d'usage indépendant (non lié), la valeur est simplement lue et traitée par code.

---

## 8.3.11. Pièges et bonnes pratiques

- **Un groupe d'options renvoie un nombre, pas un booléen** : c'est la confusion la plus fréquente. Lier le groupe à un champ **numérique** et raisonner sur l'`OptionValue`.
- **`True` vaut -1**, jamais `1` : en SQL, écrire `= True`, `= -1` ou `<> 0` pour un champ Oui/Non.
- **Définir l'`OptionValue` de chaque option** ; les options ne sont pas liées individuellement, seul le cadre l'est.
- **Gérer le cas `Null`** d'un groupe sans sélection, et d'une case à cocher en mode `TripleState` (via `IsNull`, pas `Case Null`).
- **Le groupe d'options est un conteneur** : ses options ont pour `Parent` le groupe (chapitre 8.1).
- **Réagir dans `AfterUpdate`** pour activer/masquer d'autres contrôles selon le choix (chapitres 8.9 et 8.11).
- **Choisir le bon type de champ lié** : Oui/Non pour les cases et bascules, numérique pour les groupes (chapitre 8.6).

---

## 8.3.12. Récapitulatif

- La **case à cocher** et le **bouton bascule** expriment un **booléen** (`True`/`False`) ; ils se lient à un champ **Oui/Non**. Le bouton bascule n'est qu'une apparence différente de la même logique.
- Dans Access, **`True` vaut -1** et `False` vaut 0 — à retenir pour les critères SQL.
- La case à cocher peut, en mode **`TripleState`**, prendre une troisième valeur `Null` à traiter explicitement (via `IsNull`).
- Le **groupe d'options** est un **conteneur** qui rend des options exclusives et renvoie un **nombre** (l'`OptionValue` de l'option choisie) ; il se lie à un champ **numérique**.
- Chaque option porte un `OptionValue` ; seul le **cadre** est lié aux données, les options étant identifiées par cette valeur.
- On lit le groupe avec un `Select Case` sur sa `Value`, on sélectionne par code en affectant l'`OptionValue`, et on réagit dans `AfterUpdate` (chapitre 8.9) ; ses options se parcourent via sa collection `Controls` (chapitre 8.1).

⏭️ [8.4. Contrôle onglet (TabControl) — navigation multi-pages](/08-controles-evenements/04-controle-onglet.md)
