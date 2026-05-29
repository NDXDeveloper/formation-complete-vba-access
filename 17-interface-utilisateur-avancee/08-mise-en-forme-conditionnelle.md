🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.8. Mise en forme conditionnelle par code

Faire ressortir visuellement une donnée selon sa valeur — un solde négatif en rouge, un stock sous le seuil sur fond orange, une échéance dépassée en gras — est l'un des moyens les plus efficaces de rendre une interface lisible. L'œil repère instantanément l'exception sans avoir à lire chaque chiffre. C'est l'objet de la **mise en forme conditionnelle** : adapter l'apparence des contrôles (couleur de fond, couleur de texte, gras) en fonction des données affichées ou de conditions calculées.

Access propose une fonctionnalité de mise en forme conditionnelle accessible dans le ruban, qui définit des règles de manière déclarative. Mais ces règles sont **statiques** : elles sont figées à la conception. Le pilotage **par code**, sujet de cette section, ouvre des possibilités que l'interface ne permet pas — des règles dont les seuils proviennent d'une table de paramètres, des règles qui dépendent du contexte d'exécution ou des préférences de l'utilisateur, ou l'application d'un même jeu de règles à de nombreux contrôles en boucle. Comprendre comment manipuler ces règles par VBA suppose d'abord de saisir un piège propre à Access, celui des formulaires continus, qui détermine tout le reste.

## Deux mécanismes, deux logiques

Il existe en Access deux façons radicalement différentes de changer l'apparence d'un contrôle selon une condition, et les confondre est la principale source d'erreurs sur ce sujet.

Le premier mécanisme est la **collection `FormatConditions`** : un ensemble de règles, natif à Access, qu'on attache à un contrôle et qu'Access **réévalue ligne par ligne**. C'est le mécanisme de la fonctionnalité du ruban, mais il est tout aussi accessible par code.

Le second est la **manipulation directe des propriétés** du contrôle (`BackColor`, `ForeColor`, `FontBold`…) dans le code d'un événement. On modifie alors l'apparence du contrôle lui-même, à un instant donné.

La différence cruciale tient à la manière dont chacun se comporte sur un formulaire affichant **plusieurs enregistrements à la fois** — et c'est ce qu'il faut examiner avant tout.

## Le piège des formulaires continus

Sur un **formulaire en mode continu** ou en **feuille de données** (modes d'affichage de la section 6.3), Access n'affiche pas un jeu de contrôles par enregistrement : il affiche **un seul jeu de contrôles, répété visuellement** pour chaque ligne. Cette réalité a une conséquence décisive : modifier une propriété d'apparence d'un contrôle par code — par exemple `Me!txtSolde.ForeColor = vbRed` — affecte **toutes les lignes à la fois**, puisqu'il n'y a, en réalité, qu'un seul contrôle `txtSolde`. Colorer le solde de l'enregistrement courant en rouge colore donc le solde de *tous* les enregistrements visibles. C'est le piège classique, et il prend de court tous ceux qui découvrent le sujet.

La manipulation directe des propriétés ne convient donc qu'aux contextes où **un seul enregistrement** est affiché. Pour une mise en forme par ligne sur un formulaire continu, **seule la collection `FormatConditions` fonctionne**, parce qu'Access l'évalue pour chaque ligne lors du rendu. Sur un état, un troisième cas s'ouvre, exploitant les événements de section. Le tableau suivant résume cette cartographie, qui gouverne tout choix de méthode :

| Contexte d'affichage | Manipulation directe des propriétés | Collection `FormatConditions` |
|----------------------|-------------------------------------|-------------------------------|
| Formulaire en mode **simple** (un enregistrement) | ✔ fonctionne par enregistrement (dans `Form_Current`) | ✔ fonctionne |
| Formulaire **continu** / **feuille de données** | ✘ colore **toutes** les lignes identiquement | ✔ seule voie correcte (évaluation par ligne) |
| **État** (Report) | ✔ fonctionne par ligne (événement `Format`/`Print`, avec réinitialisation) | ✔ fonctionne par ligne |

Retenir cette grille évite l'essentiel des déconvenues : on ne tente jamais de colorer ligne à ligne un formulaire continu par manipulation directe, et l'on réserve cette dernière aux formulaires simples et aux états.

## La collection FormatConditions

Chaque contrôle de saisie (zone de texte, liste déroulante) expose une collection `FormatConditions` qui contient ses règles de mise en forme. On y ajoute des règles avec la méthode `Add`, qui renvoie un objet `FormatCondition` dont on règle ensuite l'apparence.

La méthode `Add` prend d'abord un **type de condition** : `acFieldValue` (selon la valeur du champ), `acExpression` (selon une expression libre) ou `acFieldHasFocus` (lorsque le contrôle a le focus). Pour le type `acFieldValue`, on précise ensuite un **opérateur** — `acLessThan`, `acGreaterThan`, `acEqual`, `acBetween`, etc. — et une ou deux valeurs de comparaison. Pour `acExpression`, l'opérateur est omis et l'on fournit directement l'expression, qui peut référencer les champs par leur nom.

L'objet `FormatCondition` retourné expose les propriétés d'apparence : `BackColor`, `ForeColor`, `FontBold`, `FontItalic`, `FontUnderline` et `Enabled`. À noter — c'est une limite à connaître — qu'il **n'expose ni le nom ni la taille de police**, ni la visibilité : ces aspects ne sont modifiables que par manipulation directe, donc hors formulaires continus.

Un réflexe est indispensable : **vider les règles existantes avant d'en ajouter**, via `FormatConditions.Delete`, faute de quoi les règles s'accumulent à chaque passage. Par ailleurs, le nombre de règles par contrôle, autrefois très limité (trois dans les versions anciennes), a été porté à plusieurs dizaines dans les versions récentes — suffisant pour la plupart des besoins, mais non illimité.

## Mise en forme par valeur et par expression

Voici l'application typique : configurer, à l'ouverture du formulaire, des règles colorant le solde selon sa valeur. On centralise cette configuration dans une procédure appelée depuis `Form_Load` :

```vba
' --- Dans le module du formulaire ---
Option Compare Database
Option Explicit

Private Sub Form_Load()
    AppliquerMiseEnFormeSolde
End Sub

Private Sub AppliquerMiseEnFormeSolde()
    Dim fc As FormatCondition

    With Me!txtSolde
        .FormatConditions.Delete                 ' on repart d'une base propre

        ' Solde strictement négatif → texte rouge gras.
        Set fc = .FormatConditions.Add(acFieldValue, acLessThan, "0")
        fc.ForeColor = vbRed
        fc.FontBold = True

        ' Solde nul → texte orangé.
        Set fc = .FormatConditions.Add(acFieldValue, acEqual, "0")
        fc.ForeColor = RGB(200, 120, 0)
    End With
End Sub
```

Ces règles, posées sur le contrôle `txtSolde`, s'appliqueront **par ligne** y compris sur un formulaire continu, puisque c'est Access qui les réévalue à chaque rendu. Pour une condition portant sur un **autre** champ que celui affiché, on emploie le type `acExpression` :

```vba
' Griser le fond du nom lorsque le client est marqué inactif.
Dim fc As FormatCondition
With Me!txtNom
    .FormatConditions.Delete
    Set fc = .FormatConditions.Add(acExpression, , "[Statut]='Inactif'")
    fc.BackColor = RGB(235, 235, 235)
    fc.FontItalic = True
End With
```

## Mise en forme pilotée par les données

C'est ici que le pilotage par code prend tout son sens, là où la fonctionnalité statique du ruban est impuissante. Lorsque les **seuils** d'une règle ne sont pas fixés une fois pour toutes mais proviennent d'une table de paramètres (modifiable par un administrateur sans rouvrir l'application en conception), seul le code peut les lire à l'exécution et construire les règles en conséquence. On combine alors `FormatConditions` et les fonctions de domaine (section 11.10) :

```vba
' Colorer le stock selon un seuil bas lu dynamiquement dans les paramètres.
Private Sub AppliquerSeuilStock()
    Dim fc As FormatCondition
    Dim seuilBas As Long

    ' Le seuil provient de la configuration, pas d'une valeur figée.
    seuilBas = Nz(DLookup("Valeur", "tblParametres", _
                          "Cle = 'SeuilStockBas'"), 10)

    With Me!txtStock
        .FormatConditions.Delete
        Set fc = .FormatConditions.Add(acFieldValue, acLessThan, CStr(seuilBas))
        fc.BackColor = RGB(255, 205, 205)
        fc.FontBold = True
    End With
End Sub
```

Le même principe permet d'adapter les règles aux **préférences de l'utilisateur** (un daltonien pourrait choisir un autre code couleur), au **profil** connecté, ou à tout autre paramètre déterminé seulement à l'exécution. Cette adaptabilité dynamique est précisément ce que la définition statique ne sait pas offrir, et la raison d'être de la mise en forme conditionnelle par code.

## Manipulation directe dans les événements

Pour les contextes mono-enregistrement, la manipulation directe des propriétés reste pertinente, notamment lorsqu'on doit agir sur des aspects que `FormatConditions` ne couvre pas (taille de police, visibilité).

Sur un **formulaire en mode simple**, on agit dans l'événement `Form_Current` (section 8.8), déclenché à chaque changement d'enregistrement :

```vba
' Formulaire en mode SIMPLE uniquement (un enregistrement affiché).
Private Sub Form_Current()
    If Nz(Me!txtSolde, 0) < 0 Then
        Me!txtSolde.ForeColor = vbRed
        Me!txtSolde.FontBold = True
    Else
        Me!txtSolde.ForeColor = vbBlack          ' réinitialisation indispensable
        Me!txtSolde.FontBold = False
    End If
End Sub
```

Sur un **état**, la mise en forme par ligne s'obtient dans l'événement `Format` (ou `Print`) de la section concernée (événements d'état de la section 7.2), qui se déclenche pour chaque enregistrement lors de la mise en page :

```vba
' Sur un état : mise en forme par ligne dans l'événement Format du Détail.
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    If Nz(Me!txtSolde, 0) < 0 Then
        Me!txtSolde.ForeColor = vbRed
    Else
        Me!txtSolde.ForeColor = vbBlack
    End If
End Sub
```

Dans ces deux cas, un point est capital : **toujours prévoir la branche `Else` de réinitialisation**. Les propriétés persistent en effet d'un enregistrement (ou d'une ligne) au suivant ; sans remise à l'état normal, une couleur appliquée à une ligne « déteindrait » sur les lignes suivantes ne remplissant pourtant pas la condition. C'est l'oubli le plus fréquent de cette approche.

## Quel mécanisme choisir

La décision découle directement de la grille présentée plus haut, affinée par la nature de la mise en forme souhaitée. Pour un **formulaire continu ou une feuille de données**, le choix est imposé : la collection `FormatConditions` est l'unique voie correcte. Pour un **formulaire simple** ou un **état**, les deux mécanismes fonctionnent ; on privilégiera `FormatConditions` pour sa simplicité et son évaluation automatique, en réservant la manipulation directe aux cas où l'on doit agir sur la taille de police, le nom de police ou la visibilité — propriétés que `FormatConditions` ne gère pas. Enfin, dès que les **règles dépendent de données ou de conditions d'exécution**, le pilotage par code s'impose, quel que soit le mécanisme, face à la définition statique du ruban.

## Quelques applications notables

Plusieurs usages courants méritent d'être signalés. La **mise en évidence du contrôle qui a le focus** s'obtient simplement avec le type `acFieldHasFocus`, qui colore le champ actif pour guider l'œil pendant la saisie — un complément naturel à la navigation clavier de la section 17.6.

La **mise en évidence de l'enregistrement courant** sur un formulaire continu est un classique : on pose, sur les contrôles, une règle `acExpression` comparant la clé de chaque ligne à celle de l'enregistrement actif, ce qui surligne la ligne sélectionnée. C'est un motif reconnu, qui demande un rafraîchissement à chaque changement d'enregistrement.

Enfin, pour l'**alternance de couleur entre les lignes** (le « zébrage » qui facilite la lecture des listes), on n'écrira aucun code : les formulaires continus et feuilles de données disposent d'une propriété intégrée, **`AlternateBackColor`** (couleur de fond de remplacement), qui réalise cet effet nativement. Réinventer le zébrage par mise en forme conditionnelle serait une complication inutile.

## Pièges courants

Le piège dominant, longuement exposé, est la **manipulation directe sur formulaire continu**, qui colore toutes les lignes au lieu d'une seule : on bascule alors obligatoirement sur `FormatConditions`.

Vient ensuite le **fond qui ne s'affiche pas** : pour qu'une `BackColor` soit visible sur une zone de texte, sa propriété `BackStyle` doit valoir *Normal* et non *Transparent* — un réglage souvent oublié qui fait croire à une couleur inopérante.

L'**oubli de `FormatConditions.Delete`** avant d'ajouter des règles entraîne leur accumulation au fil des ouvertures ou des appels, avec des comportements erratiques.

L'**absence de réinitialisation** (`Else`) en manipulation directe laisse une couleur déteindre d'un enregistrement sur le suivant, sur les formulaires simples comme dans les événements d'état.

Deux limites enfin : `FormatCondition` ne gère **ni le nom ni la taille de police** ni la visibilité, qui exigent la manipulation directe ; et le **nombre de règles par contrôle**, bien que généreux dans les versions récentes, reste borné.

## Synthèse

La mise en forme conditionnelle adapte l'apparence des contrôles aux données, rendant l'information lisible d'un coup d'œil. Deux mécanismes coexistent, dont le choix dépend du contexte d'affichage : la **collection `FormatConditions`**, native et évaluée **par ligne** par Access, et la **manipulation directe des propriétés** dans les événements. La règle d'or tient au comportement des formulaires continus, où un jeu unique de contrôles est répété : la manipulation directe y colore toutes les lignes, si bien que seule `FormatConditions` permet une mise en forme par ligne. La manipulation directe reste réservée aux formulaires simples (`Form_Current`) et aux états (événement `Format`/`Print`, avec une réinitialisation impérative).

Le pilotage **par code** dépasse la fonctionnalité statique du ruban dès que les règles doivent être dynamiques — seuils lus dans une table de paramètres, conditions d'exécution, préférences utilisateur. On y ajoute les règles avec `FormatConditions.Add` (types `acFieldValue`, `acExpression`, `acFieldHasFocus`), après avoir vidé les règles existantes, en veillant aux limites de l'objet `FormatCondition` (ni police ni visibilité) et au réglage `BackStyle`. Pour rester cohérent, les couleurs employées gagneront à être puisées dans la palette centralisée de l'application, sujet de la section suivante (17.9), plutôt qu'éparpillées en valeurs littérales d'un contrôle à l'autre.

---

> 🔗 **Pour aller plus loin :** le comportement par ligne dépend des modes d'affichage des formulaires (section 6.3) ; la mise en forme des états s'appuie sur leurs événements `Format`/`Print` (section 7.2) et celle des formulaires simples sur l'événement `Current` (section 8.8) ; la manipulation des propriétés à l'exécution relève de la section 17.11 ; les règles pilotées par les données mobilisent les fonctions de domaine (section 11.10) et une table de paramètres ; la cohérence des couleurs renvoie aux thèmes visuels (section 17.9) ; et la mise en évidence du champ actif complète la navigation clavier (section 17.6).

⏭️ [17.9. Thèmes visuels et personnalisation cohérente de l'application](/17-interface-utilisateur-avancee/09-themes-visuels.md)
