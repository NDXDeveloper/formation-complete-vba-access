🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.3. Sections : En-tête, Détail, Pied de page, En-têtes/pieds de groupe

Les **sections** sont les briques structurelles d'un état. Chacune occupe une bande horizontale et possède un **rythme d'impression** précis : certaines s'impriment une seule fois, d'autres à chaque page, d'autres encore une fois par enregistrement ou par groupe. Comprendre ce rythme — et savoir dans quelle section placer chaque élément — est indispensable pour concevoir un état correct.

Cette section détaille les sept types de sections, leur disposition sur la page, et leur manipulation par code. Elle prolonge la vue d'ensemble de la section 7.1 et s'appuie sur les événements `Format`/`Print` présentés en 7.2.

---

## 7.3.1. Les sept types de sections et leur rythme d'impression

| Section | Constante | Fréquence d'impression |
|---|---|---|
| En-tête d'état | `acHeader` | **Une fois**, tout au début de l'état |
| En-tête de page | `acPageHeader` | En haut de **chaque page** |
| En-tête de groupe | `acGroupLevel1Header`… | Une fois au **début de chaque groupe** |
| Détail | `acDetail` | **Une fois par enregistrement** |
| Pied de groupe | `acGroupLevel1Footer`… | Une fois à la **fin de chaque groupe** |
| Pied de page | `acPageFooter` | En bas de **chaque page** |
| Pied d'état | `acFooter` | **Une fois**, tout à la fin de l'état |

Toutes ces sections ne sont pas obligatoires : l'en-tête/pied d'état et de page peuvent être absents, et les sections de groupe n'existent que si l'on a défini des regroupements (voir 7.3.6).

---

## 7.3.2. Disposition des sections sur la page

Le schéma suivant illustre l'ordre des sections au fil de l'impression :

```
┌──────────────────────────────────────────┐  ◄── Première page
│  EN-TÊTE D'ÉTAT            (1 fois)      │
├──────────────────────────────────────────┤
│  EN-TÊTE DE PAGE          (chaque page)  │
├──────────────────────────────────────────┤
│    En-tête de groupe      (par groupe)   │
│      Détail               (par enreg.)   │
│      Détail                              │
│      Détail                              │
│    Pied de groupe         (par groupe)   │
├──────────────────────────────────────────┤
│  PIED DE PAGE             (chaque page)  │
└──────────────────────────────────────────┘

   ... pages intermédiaires :
       EN-TÊTE DE PAGE → corps → PIED DE PAGE ...

┌──────────────────────────────────────────┐  ◄── Dernière page
│  EN-TÊTE DE PAGE                         │
│      ... derniers détails / pieds ...    │
│  PIED D'ÉTAT              (1 fois)       │
│  PIED DE PAGE                            │
└──────────────────────────────────────────┘
```

> **Détail subtil** : le pied d'état s'imprime une seule fois, **juste au-dessus** du dernier pied de page, sur la dernière page. Il ne crée pas de page supplémentaire à lui seul.

---

## 7.3.3. La section Détail

La section **Détail** constitue le corps de l'état : elle se répète **une fois par enregistrement** de la source. C'est elle qui affiche les données ligne par ligne.

C'est aussi la section dont les événements `Format` et `Print` (chapitre 7.2) sont les plus sollicités, notamment pour la mise en forme conditionnelle d'une ligne :

```vba
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    ' Alterner discrètement la couleur de fond une ligne sur deux
    If (Me.txtNumLigne.Value Mod 2) = 0 Then
        Me.Section(acDetail).BackColor = RGB(240, 240, 240)
    Else
        Me.Section(acDetail).BackColor = vbWhite
    End If
End Sub
```

---

## 7.3.4. En-tête et pied d'état

- L'**en-tête d'état** (`acHeader`) s'imprime **une fois**, au tout début : page de titre, logo, libellé général, période couverte.
- Le **pied d'état** (`acFooter`) s'imprime **une fois**, à la fin : c'est l'emplacement naturel des **totaux généraux**.

Une fonction d'agrégat placée dans le pied d'état porte sur **l'ensemble** des enregistrements de l'état :

```
' Source d'un contrôle du pied d'état (total général)
=Somme([Montant])
```

> Le comportement de l'en-tête/pied de page sur la page qui contient l'en-tête ou le pied d'état se règle via les propriétés **`PageHeader`** et **`PageFooter`** de l'état (valeurs : toutes les pages, sauf avec en-tête d'état, sauf avec pied d'état, sauf avec les deux).

---

## 7.3.5. En-tête et pied de page

- L'**en-tête de page** (`acPageHeader`) s'imprime en **haut de chaque page** : titres de colonnes, intitulé répété.
- Le **pied de page** (`acPageFooter`) s'imprime en **bas de chaque page** : numéro de page, date d'impression.

Le numéro de page s'obtient avec les propriétés `Page` et `Pages` (chapitre 7.1), généralement via la source d'un contrôle :

```
' Source d'un contrôle du pied de page
=[Page] & " / " & [Pages]
```

> **Restriction importante** : les sections de page **n'acceptent pas** les fonctions d'agrégat portant sur les enregistrements (`Somme`, `Moyenne`…). Y placer `=Somme([Montant])` produit `#Erreur`. Ces fonctions ne sont valides que dans les sections de **groupe** et d'**état** (voir 7.3.8).

---

## 7.3.6. En-têtes et pieds de groupe

Les sections de groupe n'apparaissent que lorsqu'on définit un ou plusieurs **regroupements** (volet « Regrouper, trier et total » en mode création, ou collection `GroupLevel` par code — voir chapitre 7.5). Access autorise jusqu'à **dix niveaux** de regroupement, chacun pouvant disposer d'un en-tête, d'un pied, ou des deux.

- L'**en-tête de groupe** s'imprime au début de chaque groupe : il porte généralement le **libellé du groupe** (par exemple le nom de la catégorie).
- Le **pied de groupe** s'imprime à la fin de chaque groupe : il accueille les **sous-totaux**.

```
' Source d'un contrôle du pied de groupe (sous-total du groupe courant)
=Somme([Montant])
```

> **Piège de numérotation** : les constantes de section utilisent l'indice **1** pour le premier groupe (`acGroupLevel1Header`), alors que la collection `GroupLevel` et le nom de code de la section sont **basés sur 0**. Ainsi, le premier groupe correspond à : constante `acGroupLevel1Header` (5), nom de code `GroupHeader0`, et élément `GroupLevel(0)`. Cette incohérence d'indexation est une source d'erreurs fréquente.

---

## 7.3.7. Manipuler les sections par code

L'accès aux sections se fait via `Me.Section(constante)`, comme vu en 7.1. On agit le plus souvent sur leur visibilité ou leurs sauts de page, dans l'événement `Format` de la section concernée.

```vba
' Masquer dynamiquement un en-tête de groupe selon une condition
Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
    Me.Section(acGroupLevel1Header).Visible = (Nz(Me.txtNbLignes.Value, 0) > 0)
End Sub
```

> Rappel (chapitre 7.2) : ces manipulations dans `Format`/`Print` n'opèrent qu'en aperçu et à l'impression, pas en mode État. Pour un rendu cohérent en mode État, recourir à la mise en forme conditionnelle (chapitre 17.8).

---

## 7.3.8. Fonctions d'agrégat selon la section : un piège majeur

L'endroit où l'on place une fonction d'agrégat détermine ce qu'elle calcule — et **si** elle fonctionne :

| Section | `=Somme([Montant])` | Calcule… |
|---|---|---|
| Détail | Sans objet (valeur de la ligne) | — |
| En-tête / pied de groupe | ✔ Valide | Total du **groupe** courant |
| En-tête / pied d'état | ✔ Valide | Total **général** |
| En-tête / pied de **page** | �’✗ `#Erreur` | **Non autorisé** |

Pour des cumuls progressifs (numérotation de lignes, totaux courants), on utilise la propriété **`RunningSum`** d'un contrôle plutôt qu'une fonction d'agrégat. Ces techniques de calcul sont développées au chapitre 7.5.

---

## 7.3.9. Propriétés de regroupement et de pagination utiles

Plusieurs propriétés gouvernent le comportement des sections face aux sauts de page (rappelées de 7.1, ici dans leur contexte d'usage) :

- **`ForceNewPage`** : impose un saut de page avant et/ou après une section. Typiquement, `ForceNewPage = 1` (avant la section) sur un en-tête de groupe place **chaque groupe sur une nouvelle page**.
- **`KeepTogether`** (propriété de section) : tente de maintenir la section sur une seule page.
- **`RepeatSection`** : répète un **en-tête de groupe** en haut de chaque page lorsque le groupe s'étale sur plusieurs pages — précieux pour conserver le contexte.
- La propriété **`KeepTogether` du groupe** (volet de regroupement) propose en outre de garder ensemble *tout le groupe*, ou *le groupe avec sa première ligne de détail*.

```vba
' Chaque groupe de niveau 1 commence sur une nouvelle page
Me.Section(acGroupLevel1Header).ForceNewPage = 1

' Répéter l'en-tête de groupe en haut de chaque page
Me.Section(acGroupLevel1Header).RepeatSection = True
```

---

## 7.3.10. Pièges et bonnes pratiques

- **Ne pas confondre en-tête d'état et en-tête de page** : le premier s'imprime **une fois**, le second **à chaque page**. Les titres de colonnes vont dans l'en-tête de **page**.
- **Placer les agrégats dans la bonne section** : `Somme`, `Moyenne`… fonctionnent dans les sections de **groupe** et d'**état**, jamais dans les sections de **page**.
- **Utiliser `RunningSum`** pour les cumuls progressifs et la numérotation, plutôt que des fonctions d'agrégat (chapitre 7.5).
- **Attention à l'indexation des groupes** : indice 1 dans les constantes (`acGroupLevel1Header`), mais 0 dans `GroupLevel` et les noms de code (`GroupHeader0`).
- **`RepeatSection`** maintient le contexte des longs groupes sur plusieurs pages ; à activer dès qu'un groupe risque de déborder.
- **Manipuler les sections dans `Format`** pour l'aperçu/impression, et prévoir une mise en forme conditionnelle (chapitre 17.8) pour le mode État.

---

## 7.3.11. Récapitulatif

- Un état comporte jusqu'à sept types de sections, chacune avec un **rythme d'impression** propre : en-tête/pied d'état (une fois), en-tête/pied de page (chaque page), en-têtes/pieds de groupe (par groupe), détail (par enregistrement).
- Le **pied d'état** s'imprime une seule fois, juste au-dessus du dernier pied de page.
- La section **Détail** se répète par enregistrement et concentre la mise en forme conditionnelle via `Format`.
- Les **sections de groupe** n'existent qu'avec des regroupements (jusqu'à 10 niveaux) et accueillent libellés et **sous-totaux** ; attention à l'indexation mixte 1/0.
- Les **fonctions d'agrégat** sont valides dans les sections de groupe et d'état, mais **interdites** dans les sections de page (`#Erreur`).
- Les propriétés `ForceNewPage`, `KeepTogether` et `RepeatSection` gouvernent la pagination des sections ; les calculs progressifs reposent sur `RunningSum` (chapitre 7.5).

⏭️ [7.4. Manipulation dynamique du RecordSource et des filtres](/07-etats-reports/04-recordsource-filtres-dynamiques.md)
