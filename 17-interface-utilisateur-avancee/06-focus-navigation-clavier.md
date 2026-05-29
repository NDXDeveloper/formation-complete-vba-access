🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.6. Gestion du focus et navigation clavier avancée

Pour l'utilisateur occasionnel, la souris suffit. Pour l'opérateur de saisie qui enchaîne des centaines de fiches par jour, en revanche, chaque aller-retour vers la souris est une perte de temps et de fluidité. Une application de saisie réellement efficace se pilote **entièrement au clavier** : on passe d'un champ au suivant sans quitter la frappe, on valide d'une touche, on déclenche les fonctions par des raccourcis, et le curseur se positionne toujours là où on l'attend. Cette ergonomie discrète distingue un outil agréable d'un outil pénible, et conditionne la productivité des utilisateurs intensifs.

Cette maîtrise repose sur trois piliers : le **contrôle du focus** — savoir quel contrôle est actif et le déplacer à volonté —, la **navigation au clavier** — l'ordre de tabulation et le comportement des touches de déplacement —, et la **capture des touches** — l'interception du clavier au niveau du formulaire pour implémenter des raccourcis. Les événements de focus eux-mêmes (`Enter`, `Exit`, `GotFocus`, `LostFocus`) et la validation des données ayant été traités au chapitre 8 (sections 8.9 et 8.10), cette section se concentre sur leur exploitation au service d'une navigation clavier avancée.

## Le focus et le contrôle actif

Le *focus* désigne le contrôle qui reçoit les frappes au clavier à un instant donné — celui dans lequel le curseur clignote. À tout moment, un seul contrôle détient le focus sur le formulaire actif. Toute navigation clavier consiste, au fond, à faire passer ce focus d'un contrôle à l'autre de la manière la plus fluide possible.

Deux outils fondamentaux permettent de manipuler le focus par code. La méthode **`SetFocus`** déplace le focus vers un contrôle précis, et la propriété **`ActiveControl`** identifie le contrôle qui le détient actuellement :

```vba
' Donner le focus à un contrôle précis.
Me!txtNom.SetFocus

' Connaître le contrôle actuellement actif.
Dim ctl As Access.Control
Set ctl = Screen.ActiveControl        ' ou Me.ActiveControl
MsgBox "Champ actif : " & ctl.Name
```

`SetFocus` est l'outil de base de toute navigation programmée : positionner le curseur sur le premier champ à l'ouverture d'un formulaire, renvoyer l'utilisateur vers un champ invalide, ou enchaîner automatiquement les étapes d'une saisie guidée.

## SetFocus : règles et pièges

`SetFocus` obéit à des conditions strictes dont la méconnaissance provoque des erreurs déroutantes. Un contrôle ne peut recevoir le focus que s'il est à la fois **visible** (`Visible = True`) et **activé** (`Enabled = True`). Tenter de donner le focus à un contrôle masqué ou désactivé déclenche l'erreur d'exécution 2110 (« Microsoft Access ne peut pas déplacer le focus vers le contrôle »). Avant un `SetFocus` conditionnel, il faut donc s'assurer que la cible est bien affichée et active — ou, si on vient de la rendre visible par code, appliquer le changement avant de lui donner le focus.

Un second piège tient à la **chronologie des événements** (section 8.7). On ne peut pas toujours déplacer le focus dans les tout premiers événements d'un formulaire : un `SetFocus` placé dans `Form_Load` échoue fréquemment, en particulier vers un sous-formulaire, car l'interface n'est pas encore pleinement réalisée. Le positionnement initial du curseur se fait plus sûrement dans l'événement `Form_Current` ou après l'ouverture complète. De même, certaines opérations sur un contrôle — la lecture de sa propriété `.Text`, la sélection de son contenu — **exigent que le contrôle détienne d'abord le focus**, faute de quoi elles échouent.

## L'ordre de tabulation

La touche Tabulation est l'épine dorsale de la navigation clavier : elle fait passer le focus d'un contrôle au suivant. L'ordre dans lequel s'effectue ce passage est gouverné par deux propriétés. La propriété **`TabIndex`** (numérotée à partir de zéro) fixe la position de chaque contrôle dans la séquence au sein de sa section. La propriété **`TabStop`** (Oui/Non), lorsqu'elle vaut *Non*, **retire le contrôle de la séquence** : la tabulation le saute, bien qu'il reste accessible à la souris — utile pour les champs en lecture seule ou rarement modifiés.

L'ordre de tabulation se règle le plus souvent en conception (menu *Ordre de tabulation*), mais il est modifiable par code lorsque le contexte l'exige :

```vba
' Réorganiser dynamiquement la séquence de tabulation.
Me!txtNom.TabIndex = 0
Me!txtPrenom.TabIndex = 1

' Exclure un champ calculé de la navigation clavier.
Me!txtTotal.TabStop = False
```

Un ordre de tabulation cohérent avec la logique de lecture du formulaire (de haut en bas, de gauche à droite, ou selon le flux de saisie réel) est l'une des optimisations les plus rentables : un ordre désordonné, où le curseur saute d'un coin à l'autre, ruine à lui seul l'ergonomie clavier.

## Le comportement en fin de séquence : la propriété Cycle

Que se passe-t-il lorsqu'on appuie sur Tabulation depuis le **dernier** contrôle ? Le comportement dépend de la propriété **`Cycle`** du formulaire, qui propose trois réglages. Avec *Tous les enregistrements* (le comportement par défaut), la tabulation passe au premier champ de l'enregistrement **suivant** — adapté à une saisie en chaîne. Avec *Enregistrement actif*, le focus **boucle** sur le premier champ du **même** enregistrement, ce qui évite de quitter accidentellement la fiche en cours — souvent préférable pour une saisie maîtrisée. Avec *Page active*, le cycle se limite à la page courante.

```vba
' Confiner la tabulation à l'enregistrement courant.
Me.Cycle = 1        ' Enregistrement actif
```

Le choix de ce réglage selon le mode de travail (saisie continue ou édition fiche par fiche) fait partie intégrante de la conception de la navigation clavier.

## La touche Entrée comme Tabulation

Beaucoup d'utilisateurs, par habitude, attendent que la touche **Entrée** fasse passer au champ suivant, comme la Tabulation. Access offre d'abord un réglage global pour cela : *Fichier → Options → Paramètres du client → Modification → « Touche Entrée »*, avec les choix *Ne pas déplacer*, *Champ suivant* et *Enregistrement suivant*. Ce réglage est cependant **propre au poste utilisateur**, et non à l'application.

Pour un comportement maîtrisé au niveau du formulaire, la technique courante consiste à intercepter la touche Entrée et à la transformer en Tabulation dans l'événement `KeyDown`, le formulaire devant pour cela capter les touches (propriété `KeyPreview`, détaillée plus bas) :

```vba
' Sur le formulaire (KeyPreview = True) : Entrée se comporte comme Tab.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 0 Then
        KeyCode = vbKeyTab          ' on substitue Tab à Entrée
    End If
End Sub
```

Modifier la valeur de `KeyCode` dans `KeyDown` change la touche effectivement traitée par Access : réaffecter `vbKeyTab` à la place de `vbKeyReturn` produit donc une tabulation. Une réserve s'impose toutefois : sur un **champ de saisie multiligne**, la touche Entrée doit conserver son rôle d'insertion de ligne ; il faut alors exempter ces champs de la substitution (en testant le contrôle actif). On évitera par ailleurs le recours à `SendKeys "{TAB}"`, technique fragile et sensible au contexte, au profit de la réaffectation de `KeyCode`.

Dans le même registre, la propriété **`AutoTab`** d'une zone de texte fait avancer automatiquement le focus dès que le champ est rempli au maximum de son masque de saisie — pratique pour les codes de longueur fixe, où l'on enchaîne les segments sans toucher à la Tabulation.

## Sélectionner le texte à l'entrée d'un champ

Un raffinement d'ergonomie très apprécié consiste à **sélectionner tout le contenu d'un champ** lorsqu'il reçoit le focus : l'utilisateur peut alors écraser la valeur existante d'une simple frappe, sans avoir à l'effacer. La sélection s'effectue via les propriétés `SelStart` et `SelLength`, dans l'événement `GotFocus` (ou `Enter`) :

```vba
Private Sub txtNom_GotFocus()
    Me!txtNom.SelStart = 0
    Me!txtNom.SelLength = Len(Me!txtNom.Text & "")
End Sub
```

On retrouve ici la règle énoncée plus haut : la propriété `.Text` et les propriétés de sélection ne sont accessibles que parce que le contrôle détient le focus au moment du `GotFocus`. Pour appliquer ce comportement à tous les champs sans dupliquer le code, on peut centraliser la sélection dans une procédure appelée depuis l'événement `Enter` de chaque contrôle, ou exploiter le contrôle actif (`Screen.ActiveControl`) depuis un point unique.

## KeyPreview : capter le clavier au niveau du formulaire

Voici le mécanisme central pour les raccourcis clavier. Par défaut, une frappe est reçue par le contrôle qui détient le focus, et lui seul. La propriété **`KeyPreview`** du formulaire, mise à `True`, change cette règle : le **formulaire reçoit alors les événements clavier *avant* le contrôle actif**. Cela permet d'implémenter des raccourcis valables sur l'ensemble du formulaire, quel que soit le champ en cours d'édition.

```vba
' Activer la capture des touches au niveau du formulaire.
Private Sub Form_Load()
    Me.KeyPreview = True
End Sub
```

Un point important : même avec `KeyPreview`, la touche est ensuite transmise au contrôle actif — à moins qu'on ne la **neutralise** explicitement. C'est ce qui permet de « consommer » une touche destinée à un raccourci, pour qu'elle ne produise pas en plus son effet normal dans le champ.

## Les événements clavier : KeyDown, KeyPress, KeyUp

Trois événements rendent compte de l'activité du clavier, et le choix entre eux n'est pas indifférent.

L'événement **`KeyDown(KeyCode, Shift)`** se déclenche pour **toute** touche enfoncée, y compris les touches non imprimables (fonctions F1–F12, flèches, Échap, Tabulation). Il fournit le **code de touche virtuel** (`vbKeyF1`, `vbKeyReturn`, `vbKeyLeft`…) et l'**état des touches de modification** dans son paramètre `Shift`. C'est l'événement de choix pour les raccourcis à base de touches de fonction ou de combinaisons Ctrl/Alt, et pour neutraliser une touche (en posant `KeyCode = 0`).

L'événement **`KeyPress(KeyAscii)`** ne se déclenche que pour les touches produisant un **caractère**, et fournit le code ANSI de ce caractère. Il est dédié au **filtrage de la saisie** — n'autoriser que les chiffres, convertir en majuscules, refuser certains caractères — la neutralisation se faisant par `KeyAscii = 0`.

L'événement **`KeyUp`**, déclenché au relâchement, est plus rarement utilisé.

Le paramètre `Shift` de `KeyDown` est un **masque de bits** combinant les touches de modification, qu'on teste par un ET logique avec les constantes `acShiftMask` (1), `acCtrlMask` (2) et `acAltMask` (4).

## Raccourcis clavier personnalisés

En combinant `KeyPreview` et `KeyDown`, on dote l'application de raccourcis professionnels. L'exemple ci-dessous implémente l'enregistrement par Ctrl+S, l'aide par F1 et l'actualisation par F5 :

```vba
' Sur le formulaire, avec KeyPreview = True.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Ctrl+S : enregistrer l'enregistrement courant.
    If KeyCode = vbKeyS And (Shift And acCtrlMask) > 0 Then
        KeyCode = 0                     ' on neutralise la frappe
        If Me.Dirty Then Me.Dirty = False
        Exit Sub
    End If

    ' Touches de fonction.
    Select Case KeyCode
        Case vbKeyF1                    ' F1 : aide
            KeyCode = 0
            AfficherAide
        Case vbKeyF5                    ' F5 : actualiser
            KeyCode = 0
            Me.Requery
    End Select
End Sub
```

Poser `KeyCode = 0` après avoir traité un raccourci est essentiel : cela **consomme** la touche pour qu'elle ne produise pas, en plus, son comportement habituel dans le champ actif. À l'inverse, on prendra garde à **ne pas intercepter des combinaisons standard** attendues par les utilisateurs ou réservées par Windows et Access, sous peine de désorienter plus que d'aider.

## Touches d'accès rapide et boutons par défaut

Deux mécanismes complètent la navigation clavier des boutons et des dialogues. Les **touches d'accès rapide** (mnémoniques) s'obtiennent en plaçant le caractère `&` devant une lettre dans la légende d'un bouton : une légende `&Enregistrer` crée le raccourci Alt+E. Appliqué à l'étiquette attachée à un contrôle, ce procédé permet à Alt+lettre de **donner directement le focus** au champ associé.

Sur les boîtes de dialogue (section 6.7), les propriétés des boutons de commande **`Default`** et **`Cancel`** rendent l'interaction clavier naturelle : le bouton dont `Default` vaut *Oui* est déclenché par la touche **Entrée**, et celui dont `Cancel` vaut *Oui* l'est par **Échap**. Un dialogue ainsi configuré se valide ou s'annule sans la souris — un standard d'ergonomie que les utilisateurs attendent.

## Focus et validation

La gestion du focus est l'instrument privilégié de la validation des données (section 8.10). Lorsqu'une saisie est invalide, on **retient le focus sur le champ fautif** tant qu'il n'est pas corrigé, en annulant le départ du contrôle. Cela s'obtient en posant `Cancel = True` dans l'événement `Exit` du contrôle ou `BeforeUpdate` du formulaire (motif d'annulation détaillé en section 8.12) :

```vba
Private Sub txtEmail_Exit(Cancel As Integer)
    If Len(Me!txtEmail & "") > 0 And InStr(Me!txtEmail, "@") = 0 Then
        MsgBox "Adresse email invalide.", vbExclamation
        Cancel = True               ' le focus reste sur le champ
    End If
End Sub
```

Cette technique guide l'utilisateur sans rupture : le curseur ne quitte pas un champ tant que sa valeur n'est pas acceptable, ce qui rend la correction immédiate et évite les messages d'erreur tardifs en fin de saisie.

## Rappel : événements de focus, deux familles

Pour exploiter correctement le focus, il faut distinguer deux familles d'événements (détaillées aux sections 8.7 et 8.9). Les événements **`Enter`** et **`Exit`** se déclenchent lors des déplacements de focus **entre contrôles d'un même formulaire**, mais **ne se déclenchent pas** lorsque le formulaire entier perd ou regagne le focus au profit d'une autre fenêtre. Les événements **`GotFocus`** et **`LostFocus`**, eux, accompagnent tout gain ou perte effectif du focus, y compris lors d'un changement de fenêtre.

Cette nuance oriente le choix : `Enter`/`Exit` conviennent à la logique de navigation interne (préparer un champ à la saisie, valider à sa sortie), tandis que `GotFocus`/`LostFocus` captent l'activation réelle du contrôle. La sélection automatique du texte vue plus haut, par exemple, fonctionne dans l'une ou l'autre famille selon que l'on souhaite ou non qu'elle se reproduise après un retour depuis une autre fenêtre.

## Pièges courants

Le piège le plus fréquent est le `SetFocus` **vers un contrôle masqué ou désactivé**, qui déclenche l'erreur 2110 : il faut vérifier `Visible` et `Enabled` avant de déplacer le focus. Suit le `SetFocus` **trop précoce**, dans `Form_Load`, qui échoue faute d'interface réalisée — on positionne le curseur initial dans `Form_Current` ou plus tard.

L'oubli d'activer **`KeyPreview`** est la cause numéro un de raccourcis clavier inopérants : sans cette propriété, le formulaire ne voit jamais les touches, captées par le seul contrôle actif. À l'inverse, **neutraliser trop de touches** (ou des combinaisons standard) perturbe les utilisateurs : on ne consomme une touche, par `KeyCode = 0` ou `KeyAscii = 0`, que lorsqu'on a effectivement traité un raccourci.

L'accès à la propriété **`.Text` ou aux propriétés de sélection sans focus préalable** provoque une erreur : ces opérations exigent que le contrôle soit actif. Enfin, un **ordre de tabulation incohérent** — souvent hérité d'ajouts successifs de contrôles — ruine la navigation clavier malgré tout le reste : il mérite une vérification systématique en conception.

## Synthèse

Une navigation clavier soignée repose sur trois leviers. Le **contrôle du focus**, via `SetFocus` et `ActiveControl`, permet de positionner et de retenir le curseur, sous réserve que la cible soit visible, activée, et l'événement assez tardif. La **navigation au clavier** se règle par l'ordre de tabulation (`TabIndex`, `TabStop`), le comportement de fin de séquence (`Cycle`), la transformation d'Entrée en Tabulation et la sélection automatique du contenu à l'entrée d'un champ. La **capture du clavier**, enfin, s'appuie sur la propriété `KeyPreview` et les événements `KeyDown`/`KeyPress` pour implémenter des raccourcis — en neutralisant les touches consommées par `KeyCode = 0` ou `KeyAscii = 0`.

Complétée par les touches d'accès rapide mnémoniques, les boutons `Default`/`Cancel` des dialogues et la rétention du focus au service de la validation, cette maîtrise transforme une application utilisable à la souris en un outil de saisie rapide, fluide et entièrement pilotable au clavier — un gain de productivité décisif pour les utilisateurs intensifs, et un atout d'accessibilité. Comme toute finition, elle s'applique avec discernement : précieuse pour les formulaires de saisie dense, superflue pour les écrans de consultation occasionnelle.

---

> 🔗 **Pour aller plus loin :** les événements de focus (`Enter`, `Exit`, `GotFocus`, `LostFocus`) et leur enchaînement sont détaillés aux sections 8.9 et 8.7 ; la validation des données et l'annulation par `Cancel` aux sections 8.10 et 8.12 ; les propriétés des contrôles texte (masques, `AutoTab`) à la section 8.2 ; les propriétés `Cycle` et `KeyPreview` relèvent du modèle objet Form (section 6.1) ; les boutons `Default`/`Cancel` accompagnent les dialogues de la section 6.7 ; et le réglage global de la touche Entrée figure parmi les options abordées avec le démarrage à la section 17.7.

⏭️ [17.7. Options de démarrage et verrouillage de l'interface](/17-interface-utilisateur-avancee/07-options-demarrage.md)
