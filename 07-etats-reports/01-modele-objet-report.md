🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.1. Modèle objet Report — sections et propriétés

Avant de piloter un état par code, il faut comprendre **comment Access le représente** : sa place dans la hiérarchie des objets, sa structure en sections, et les propriétés qui en gouvernent la source de données et le rendu. Cette section pose ces fondations, sur lesquelles s'appuient toutes les suivantes du chapitre.

Le modèle objet d'un état est très proche de celui d'un formulaire (chapitre 6.1) : mêmes notions de sections, de contrôles, de propriété `RecordSource`. Les différences tiennent à la finalité — un état est en **lecture seule** et produit une **sortie mise en forme** — et à quelques propriétés et collections spécifiques.

---

## 7.1.1. L'objet Report dans la hiérarchie Access

Un état est représenté par un objet **`Report`**, situé sous l'objet `Application` dans la hiérarchie décrite au chapitre 4.1 :

```
Application
 └── Reports (collection des états OUVERTS)
      └── Report
           ├── Section (collection des sections)
           │    └── Controls (collection des contrôles)
           └── GroupLevel (collection des regroupements/tris)
```

Comme pour les formulaires, un état tire ses données d'une source définie par sa propriété **`RecordSource`** (une table, une requête ou une instruction SQL). À la différence d'un formulaire, cette source est exploitée en **lecture seule** : un état présente des données, il n'en modifie jamais.

---

## 7.1.2. Accéder à un état : collection Reports et AllReports

Deux collections donnent accès aux états, avec une distinction importante :

- **`Reports`** ne contient que les états **actuellement ouverts**.
- **`CurrentProject.AllReports`** liste **tous les états définis** dans l'application, ouverts ou non (chapitre 4.4).

```vba
' États actuellement ouverts
Debug.Print Reports.Count
Debug.Print Reports!E_Ventes.Caption          ' notation point d'exclamation
Debug.Print Reports("E_Ventes").Caption        ' notation explicite

' Tous les états définis (ouverts ou non)
Dim obj As AccessObject
For Each obj In CurrentProject.AllReports
    Debug.Print obj.Name, obj.IsLoaded
Next obj
```

La propriété `IsLoaded` de `AllReports` permet de tester si un état est ouvert avant de le solliciter — précaution indispensable, car accéder à un état fermé via la collection `Reports` déclenche une erreur.

---

## 7.1.3. La structure en sections d'un état

Un état est organisé en **sections** empilées verticalement, chacune jouant un rôle précis lors de la mise en forme. Le détail de leur comportement fait l'objet de la section 7.3 ; voici la vue d'ensemble et les **constantes** servant à les désigner par code :

| Constante | Valeur | Section |
|---|---|---|
| `acHeader` | 1 | En-tête d'état (une fois, en tête du rapport) |
| `acPageHeader` | 3 | En-tête de page (en haut de chaque page) |
| `acGroupLevel1Header` | 5 | En-tête de groupe — niveau 1 |
| `acDetail` | 0 | Détail (une fois par enregistrement) |
| `acGroupLevel1Footer` | 6 | Pied de groupe — niveau 1 |
| `acPageFooter` | 4 | Pied de page (en bas de chaque page) |
| `acFooter` | 2 | Pied d'état (une fois, en fin de rapport) |

Les niveaux de regroupement supplémentaires suivent la même logique : `acGroupLevel2Header` (7), `acGroupLevel2Footer` (8), etc.

L'accès à une section se fait via la propriété `Section`, indexée par ces constantes :

```vba
' Masquer la section détail
Me.Section(acDetail).Visible = False

' Couleur de fond de l'en-tête d'état
Me.Section(acHeader).BackColor = vbYellow

' Hauteur de la section détail (en twips : ~567 twips = 1 cm)
Me.Section(acDetail).Height = 567
```

> Les unités de dimension (`Height`, positions) sont exprimées en **twips** : 1 440 twips par pouce, soit environ 567 twips par centimètre.

---

## 7.1.4. Le mot-clé Me et l'accès aux éléments

Dans le module de classe d'un état, le mot-clé **`Me`** désigne l'instance de l'état courant, exactement comme pour un formulaire. Les contrôles s'y réfèrent de la même manière :

```vba
Me.txtTotal.Value = 1000              ' notation point
Me!txtTotal.Value = 1000              ' notation point d'exclamation
Me.Controls("txtTotal").Value = 1000  ' notation explicite par nom
```

L'accès aux contrôles d'un état et leur manipulation reposent sur les mêmes principes que pour les formulaires (chapitre 8).

---

## 7.1.5. Propriétés essentielles de l'objet Report

Les principales propriétés d'un état, à connaître pour le piloter par code :

| Propriété | Rôle |
|---|---|
| `RecordSource` | Source de données : table, requête ou instruction SQL |
| `Filter` / `FilterOn` | Expression de filtre et activation (voir section 7.4) |
| `OrderBy` / `OrderByOn` | Tri d'affichage et activation |
| `Caption` | Texte de la barre de titre de l'aperçu |
| `HasData` | Indique la présence d'enregistrements (central pour l'événement `NoData`, section 7.9) |
| `Page` / `Pages` | Numéro de la page courante / nombre total de pages |
| `MoveLayout`, `NextRecord`, `PrintSection` | Pilotent le flux de mise en forme dans l'événement `Format` |
| `FormatCount` / `PrintCount` | Nombre de mises en forme / d'impressions d'une section |
| `GroupLevel` | Collection des niveaux de regroupement et de tri (section 7.5) |

Quelques exemples d'usage courant :

```vba
' Adapter le titre de l'aperçu
Me.Caption = "Ventes du trimestre"

' Tester la présence de données (détaillé en 7.9)
If Me.HasData Then
    Debug.Print "L'état contient des enregistrements."
End If

' Afficher le numéro de page (typiquement dans un contrôle du pied de page)
Me.txtPage.Value = "Page " & Me.Page & " sur " & Me.Pages
```

> La propriété `Pages` (nombre total de pages) n'est fiable qu'une fois l'état entièrement mis en forme ; `Page` (page courante) est disponible tout au long de l'impression.

---

## 7.1.6. Propriétés des sections

Chaque section possède ses propres propriétés, qui gouvernent son rendu :

| Propriété | Rôle |
|---|---|
| `Visible` | Affiche ou masque la section |
| `CanGrow` / `CanShrink` | Autorise la section à s'agrandir / se réduire selon son contenu |
| `Height` | Hauteur de la section (en twips) |
| `ForceNewPage` | Saut de page : aucun (0), avant (1), après (2), avant et après (3) |
| `KeepTogether` | Maintient autant que possible la section sur une même page |
| `RepeatSection` | Répète un en-tête de groupe en haut de chaque page |
| `BackColor` / `BackStyle` | Couleur et style de fond de la section |
| `NewRowOrCol` | Nouvelle rangée ou colonne (états multi-colonnes) |

```vba
' Permettre à la section détail de s'adapter à un texte long
Me.Section(acDetail).CanGrow = True
Me.Section(acDetail).CanShrink = True

' Forcer un saut de page après l'en-tête de groupe de niveau 1
Me.Section(acGroupLevel1Header).ForceNewPage = 2   ' après la section
```

---

## 7.1.7. Distinguer les niveaux : état, section, contrôle

Une source fréquente de confusion consiste à mélanger les propriétés selon leur **niveau** d'application :

- Les propriétés de l'**état** (`RecordSource`, `Filter`, `Caption`, `Page`…) concernent le rapport dans son ensemble.
- Les propriétés de **section** (`Visible`, `CanGrow`, `Height`, `ForceNewPage`…) s'appliquent à une bande précise et se règlent via `Me.Section(…)`.
- Les propriétés de **contrôle** (`Value`, `ForeColor`, `Visible` d'un contrôle…) concernent un élément individuel et se règlent via `Me.NomControle`.

Un même nom de propriété peut exister à plusieurs niveaux (`Visible`, `BackColor`) : il faut donc toujours préciser **sur quel objet** on agit.

---

## 7.1.8. Quand les propriétés sont-elles modifiables ?

Toutes les propriétés ne se règlent pas au même moment du cycle de vie de l'état (détaillé en 7.2) :

- `RecordSource`, `Filter`, `OrderBy`, `Caption` : à régler de préférence **avant l'ouverture** ou dans l'événement **`Open`**.
- Les propriétés de mise en forme des sections (`Visible`, `CanGrow`) et les indicateurs de flux (`MoveLayout`, `NextRecord`, `PrintSection`) : à manipuler dans les événements **`Format`** et **`Print`**, qui se déclenchent au fil de la mise en forme.
- `Page` : disponible pendant la mise en forme et l'impression ; `Pages` : fiable seulement en fin de traitement.

Tenter de modifier une propriété hors de la phase appropriée produit soit une erreur, soit un effet ignoré.

---

## 7.1.9. Pièges et bonnes pratiques

- **Un état est en lecture seule** : il présente des données, il n'en écrit jamais. Toute logique d'écriture relève des formulaires (chapitre 6) ou de l'accès aux données par code (chapitre 9).
- **Vérifier l'ouverture avant l'accès** : utiliser `CurrentProject.AllReports("…").IsLoaded` (chapitre 4.4) ; la collection `Reports` ne contient que les états ouverts.
- **Préciser le niveau d'une propriété** : distinguer clairement état, section et contrôle, surtout pour les propriétés homonymes (`Visible`, `BackColor`).
- **Désigner les sections par leurs constantes** (`acDetail`, `acHeader`…) plutôt que par des indices numériques en dur : le code reste lisible.
- **Respecter le moment d'affectation** : régler `RecordSource` et `Filter` à l'ouverture, manipuler les propriétés de section dans `Format`/`Print`.
- **Se méfier de `Pages`** : ce total n'est connu qu'en fin de mise en forme ; ne pas l'exploiter prématurément.

---

## 7.1.10. Récapitulatif

- L'objet `Report` occupe, sous `Application`, une place analogue à celle du formulaire, avec une source définie par `RecordSource` exploitée en **lecture seule**.
- La collection `Reports` ne liste que les états **ouverts** ; `CurrentProject.AllReports` liste **tous** les états et expose `IsLoaded` (chapitre 4.4).
- Un état s'organise en **sections** (en-tête/pied d'état et de page, en-têtes/pieds de groupe, détail), désignées par des constantes `ac…` et accessibles via `Me.Section(…)` ; leur rôle est détaillé en 7.3.
- Les **propriétés de l'état** (`RecordSource`, `Filter`, `Caption`, `HasData`, `Page`/`Pages`, `GroupLevel`…) et les **propriétés de section** (`Visible`, `CanGrow`, `Height`, `ForceNewPage`…) se règlent à des **moments précis** du cycle de vie (chapitre 7.2).
- Il est essentiel de distinguer les niveaux **état / section / contrôle**, en particulier pour les propriétés portant le même nom.

⏭️ [7.2. Événements d'un état (Open, Activate, NoData, Format, Print, Close)](/07-etats-reports/02-evenements-etat.md)
