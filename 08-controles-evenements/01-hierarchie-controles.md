🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.1. Hiérarchie des contrôles dans Access

Avant de manipuler les contrôles par code, il faut comprendre **comment Access les organise**. Les contrôles ne sont pas tous des « enfants » directs du formulaire : certains sont imbriqués dans des conteneurs (onglets, groupes d'options). Connaître cette hiérarchie — et les différentes façons de référencer un contrôle — évite quantité d'erreurs lors de l'accès par code, notamment quand on parcourt l'ensemble des contrôles.

Cette section pose ces fondations, communes aux formulaires et aux états (dont le modèle a été introduit aux chapitres 6.1 et 7.1).

---

## 8.1.1. La collection Controls

Tout formulaire et tout état possède une collection **`Controls`** qui regroupe ses contrôles. On y accède par le nom du contrôle ou par sa position :

```vba
Me.Controls("txtNom")     ' par le nom
Me.Controls(0)            ' par l'indice (fragile : dépend de l'ordre)
```

Cette collection est le point d'entrée pour énumérer les contrôles (section 8.1.8) ou y accéder dynamiquement quand le nom n'est connu qu'à l'exécution.

---

## 8.1.2. Les quatre façons de référencer un contrôle

Un même contrôle peut être désigné de quatre manières :

```vba
Me.txtNom                 ' notation point
Me!txtNom                 ' notation point d'exclamation (bang)
Me.Controls("txtNom")     ' explicite, par le nom
Me.Controls(0)            ' explicite, par l'indice
```

Les trois premières sont équivalentes pour un nom connu ; la quatrième, par indice, est à éviter (l'ordre des contrôles peut changer). Le choix entre notation point et point d'exclamation mérite une explication.

---

## 8.1.3. Notation point (.) ou point d'exclamation (!) : laquelle choisir ?

| Notation | Résolution | Atouts | Limites |
|---|---|---|---|
| `Me.txtNom` | À la **compilation** | Vérifiée à la compilation, IntelliSense, légèrement plus rapide | Le nom doit être un identifiant valide |
| `Me!txtNom` | À l'**exécution** | Accepte les noms avec espaces/caractères spéciaux | Aucune vérification à la compilation |
| `Me.Controls(variable)` | À l'**exécution** | Indispensable quand le nom est dans une **variable** | — |

**Recommandation** : pour un nom de contrôle **fixe**, privilégier la **notation point** (`Me.txtNom`), qui bénéficie de la complétion automatique et signale une faute de frappe dès la compilation. Réserver `Me.Controls(variable)` aux cas où le nom est **calculé** :

```vba
Dim i As Integer
For i = 1 To 5
    Me.Controls("txtLigne" & i).Value = 0   ' nom dynamique
Next i
```

> Techniquement, `Me.txtNom` fonctionne parce qu'Access crée une propriété cachée pour chaque contrôle dans le module du formulaire. La notation point d'exclamation, elle, interroge la collection `Controls` par défaut.

---

## 8.1.4. Les contrôles conteneurs : une hiérarchie imbriquée

Tous les contrôles ne sont pas placés **directement** sur le formulaire. Certains sont des **conteneurs** qui en hébergent d'autres :

- le **contrôle onglet** (`TabControl`, chapitre 8.4) contient des **pages**, qui contiennent elles-mêmes des contrôles ;
- le **groupe d'options** (`OptionGroup`, chapitre 8.3) contient ses boutons d'option, cases à cocher ou boutons bascule.

Il en résulte une **hiérarchie imbriquée** :

```
Formulaire
 ├── txtNom                    (directement sur le formulaire)
 ├── Contrôle onglet
 │    ├── Page 1
 │    │    └── txtAdresse      (sur la page 1)
 │    └── Page 2
 │         └── txtTelephone    (sur la page 2)
 └── Groupe d'options
      ├── optHomme
      └── optFemme
```

**Point capital** : un contrôle imbriqué (comme `txtAdresse`) n'est pas un enfant direct du formulaire, mais de sa **page**. Cela a deux conséquences pratiques, traitées ci-dessous :

- l'**accès par le nom** ignore l'imbrication (`Me.txtAdresse` fonctionne) ;
- l'**énumération** des contrôles, elle, doit en tenir compte.

---

## 8.1.5. La propriété Parent

Chaque contrôle possède une propriété **`Parent`** désignant son conteneur immédiat :

- pour un contrôle posé directement sur le formulaire, `Parent` est le **formulaire** ;
- pour un contrôle dans un groupe d'options, `Parent` est le **groupe d'options** ;
- pour un contrôle sur une page d'onglet, `Parent` est la **page**.

```vba
Debug.Print Me.txtNom.Parent.Name        ' nom du formulaire
Debug.Print Me.optHomme.Parent.Name      ' nom du groupe d'options
```

On peut « remonter » la hiérarchie jusqu'au formulaire : `ctl.Parent.Parent…`. C'est utile, par exemple, pour qu'un contrôle imbriqué accède à une propriété du formulaire sans présumer de sa position.

---

## 8.1.6. La propriété Section

À ne pas confondre avec `Parent` : la propriété **`Section`** d'un contrôle indique dans **quelle section** il se trouve, sous forme d'une constante (chapitres 7.1 et 8.7) :

```vba
Select Case Me.txtNom.Section
    Case acDetail:    Debug.Print "Section Détail"
    Case acHeader:    Debug.Print "En-tête de formulaire"
    Case acFooter:    Debug.Print "Pied de formulaire"
End Select
```

`Parent` répond à « *dans quel conteneur ?* », `Section` à « *dans quelle bande ?* ».

---

## 8.1.7. La propriété ControlType

La propriété **`ControlType`** renvoie une constante identifiant le **type** du contrôle. Elle est essentielle pour filtrer les contrôles lors d'un parcours :

| Constante | Contrôle |
|---|---|
| `acTextBox` | Zone de texte |
| `acComboBox` | Liste déroulante |
| `acListBox` | Liste |
| `acCheckBox` | Case à cocher |
| `acCommandButton` | Bouton de commande |
| `acLabel` | Étiquette |
| `acOptionGroup` | Groupe d'options |
| `acTabCtl` | Contrôle onglet |
| `acPage` | Page d'onglet |
| `acSubform` | Sous-formulaire / sous-état |

```vba
' Vider toutes les zones de texte du formulaire
Dim ctl As Control
For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then ctl.Value = Null
Next ctl
```

---

## 8.1.8. Parcourir tous les contrôles (énumération récursive)

Parcourir `Me.Controls` ne suffit pas toujours à atteindre **tous** les contrôles : ceux qui sont imbriqués dans un onglet ou un groupe d'options se trouvent dans la collection `Controls` de leur **conteneur**. Pour les énumérer de façon fiable, on **recurse** dans les conteneurs :

```vba
Sub ParcourirControles(ByVal conteneur As Object, Optional ByVal niveau As Integer = 0)
    Dim ctl As Control
    For Each ctl In conteneur.Controls
        Debug.Print String(niveau * 2, " ") & ctl.Name & _
                    " (" & ctl.ControlType & ")"

        ' Descendre dans les conteneurs
        Select Case ctl.ControlType
            Case acTabCtl, acPage, acOptionGroup
                ParcourirControles ctl, niveau + 1
        End Select
    Next ctl
End Sub
```

Appel depuis le module du formulaire :

```vba
ParcourirControles Me
```

Cette procédure descend dans le contrôle onglet (qui livre ses pages), dans chaque page (qui livre ses contrôles) et dans les groupes d'options. C'est le **modèle robuste** pour toute opération de masse sur les contrôles (chapitre 8.11).

---

## 8.1.9. Accéder aux contrôles depuis l'extérieur du formulaire

Depuis un autre module, on accède aux contrôles d'un formulaire **ouvert** via la collection `Forms` (chapitre 6.2) :

```vba
Forms!F_Client!txtNom.Value
Forms("F_Client").Controls("txtNom").Value
```

Pour un contrôle situé dans un **sous-formulaire**, on traverse le contrôle conteneur puis sa propriété `Form` (chapitre 6.4) :

```vba
Forms!F_Commande!ctlSousForm.Form!txtPrix.Value
```

---

## 8.1.10. Les étiquettes attachées

L'étiquette associée à une zone de texte (ou à un autre contrôle) est un **contrôle distinct**, hébergé dans la collection `Controls` du contrôle lui-même. On y accède par son indice 0 :

```vba
' Modifier le libellé de l'étiquette attachée à txtNom
Me.txtNom.Controls(0).Caption = "Nom complet :"
```

C'est pourquoi un formulaire comporte souvent **deux fois plus** de contrôles qu'il n'y paraît : chaque champ saisissable s'accompagne de son étiquette.

---

## 8.1.11. Pièges et bonnes pratiques

- **Ne pas présumer que tout contrôle est enfant du formulaire** : les contrôles imbriqués (onglets, groupes d'options) ont un autre `Parent`.
- **Pour l'accès par nom, l'imbrication est transparente** : `Me.txtAdresse` fonctionne même sur une page d'onglet. Seule l'**énumération** doit gérer la hiérarchie (recursion, section 8.1.8).
- **Privilégier la notation point** pour les noms fixes (vérification à la compilation, IntelliSense) ; réserver `Me.Controls(variable)` aux noms calculés.
- **Distinguer `Parent` et `Section`** : le conteneur immédiat d'une part, la bande d'autre part.
- **Filtrer par `ControlType`** lors des parcours pour n'agir que sur les contrôles voulus.
- **Ne pas oublier les étiquettes attachées** (`ctl.Controls(0)`) lors des manipulations de masse.
- **Éviter l'accès par indice** (`Controls(0)`) pour un contrôle précis : l'ordre n'est pas garanti.

---

## 8.1.12. Récapitulatif

- Chaque formulaire et état expose une collection **`Controls`** ; un contrôle se référence par `Me.Nom`, `Me!Nom`, `Me.Controls("Nom")` ou (à éviter) par indice.
- La **notation point** est vérifiée à la compilation et recommandée pour les noms fixes ; `Me.Controls(variable)` sert aux noms **calculés**.
- Les **contrôles conteneurs** (onglet → pages, groupe d'options) créent une **hiérarchie imbriquée** : un contrôle imbriqué n'est pas enfant direct du formulaire.
- La propriété **`Parent`** désigne le conteneur immédiat, la propriété **`Section`** la bande du formulaire, et **`ControlType`** le type du contrôle.
- L'**énumération fiable** de tous les contrôles passe par une **recursion** dans les conteneurs (section 8.1.8) — base des opérations de masse (chapitre 8.11).
- L'accès depuis l'extérieur se fait via `Forms!…` (chapitre 6.2) et, pour les sous-formulaires, via `.Form` (chapitre 6.4) ; les **étiquettes attachées** sont des contrôles distincts (`ctl.Controls(0)`).

⏭️ [8.2. Contrôles texte, liste déroulante et liste (TextBox, ComboBox, ListBox)](/08-controles-evenements/02-controles-texte-listes.md)
