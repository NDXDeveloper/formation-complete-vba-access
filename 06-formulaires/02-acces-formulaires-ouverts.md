🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 6.2. Accès aux formulaires ouverts via Forms!NomFormulaire

La section précédente a montré comment un formulaire se référence **lui-même** avec `Me`. Voyons maintenant comment atteindre, depuis le code, un **autre** formulaire — et ses contrôles —, à condition qu'il soit **ouvert**. C'est le rôle de la collection `Forms` et de la notation `Forms!NomFormulaire`.

## La collection Forms : les formulaires ouverts

Rappel des sections 4.1 et 4.4 : la collection **`Forms`** ne contient que les formulaires **actuellement ouverts** (à distinguer de `CurrentProject.AllForms`, qui liste **tous** les formulaires). On accède à un membre de trois manières :

```vba
Forms!frmClients          ' notation « ! » (bang)
Forms("frmClients")       ' par nom (chaîne) — pratique si le nom est dans une variable
Forms(0)                  ' par indice (rarement utile)
```

Comme vu en section 4.1, le **`!`** accède à un élément **nommé** de la collection : `Forms!frmClients` équivaut à `Forms("frmClients")`. On préfère la forme `Forms("…")` lorsque le nom est **dynamique** (contenu dans une variable).

Pour un nom comportant des **espaces** ou des caractères spéciaux, on emploie des **crochets** ou la forme chaîne :

```vba
Forms![Fiche Client].Caption = "..."     ' crochets
Forms("Fiche Client").Caption = "..."    ' chaîne
```

## Accéder à un formulaire ouvert

`Forms!frmClients` renvoie l'objet **`Form`** correspondant ; on peut donc lire et modifier ses propriétés :

```vba
Forms!frmClients.Caption = "Clients"
Forms!frmClients.AllowAdditions = False
```

## Accéder aux contrôles d'un autre formulaire

C'est l'usage principal : depuis un formulaire (ou un module standard), atteindre un **contrôle** situé sur un **autre** formulaire ouvert. La chaîne suit la forme `Forms!NomFormulaire!NomContrôle` :

```vba
Debug.Print Forms!frmClients!txtNom         ' lecture (.Value est implicite)
Forms!frmClients!txtNom = "Dupont"          ' écriture
Forms!frmClients!txtNom.Enabled = False      ' une propriété du contrôle
```

La structure complète est donc `Forms!Formulaire!Contrôle.Propriété`. Comme `Value` est la propriété par défaut d'un contrôle, on peut l'omettre en lecture/écriture de valeur.

## Me ou Forms!NomFormulaire ?

Le choix est simple et important :

- pour accéder à **son propre** formulaire, depuis son code, utilisez **`Me`** (section 6.1) : c'est plus court, plus rapide, et insensible à un renommage du formulaire ;
- pour accéder à un **autre** formulaire, utilisez **`Forms!…`**.

```vba
' Dans son PROPRE code : Me (recommandé)
Me.txtNom = "Dupont"

' Pour un AUTRE formulaire ouvert : Forms!
Forms!frmCommandes!txtTotal = 0
```

N'utilisez pas `Forms!frmMien!…` à l'intérieur de `frmMien` : `Me` y est toujours préférable.

## Le piège : le formulaire doit être ouvert

Puisque `Forms` ne contient que les formulaires **ouverts**, référencer un formulaire **fermé** déclenche une **erreur** (typiquement la n° 2450, « Access ne trouve pas le formulaire référencé »). C'est l'un des pièges les plus courants.

La parade : **vérifier que le formulaire est ouvert** avant d'y accéder, grâce à la propriété **`IsLoaded`** (section 4.4) :

```vba
If CurrentProject.AllForms("frmClients").IsLoaded Then
    Forms!frmClients!txtNom = "Dupont"
Else
    MsgBox "Le formulaire des clients n'est pas ouvert."
End If
```

## Le cas des sous-formulaires

Référencer un contrôle situé **dans un sous-formulaire** demande une étape supplémentaire, source de confusion fréquente. Un sous-formulaire est hébergé dans un **contrôle sous-formulaire** posé sur le formulaire parent ; pour atteindre le formulaire imbriqué, il faut passer par la propriété **`.Form`** de ce contrôle :

```vba
Forms!frmCommande!sfrmLignes.Form!txtQuantite = 5
'        parent      contrôle    .Form  contrôle du sous-formulaire
```

Ici, `sfrmLignes` est le **nom du contrôle sous-formulaire** (qui peut différer du nom du formulaire qu'il contient), et `.Form` fait le pont vers l'objet `Form` imbriqué. Ce mécanisme — ainsi que la liaison parent/enfant — est détaillé en section 6.4.

## Optimisation : capturer le formulaire dans une variable

Lorsqu'on accède **plusieurs fois** au même formulaire, écrire `Forms!frmClients!…` à chaque ligne oblige Access à **re-résoudre** la référence à chaque fois. Il est plus efficace (et plus lisible) de la **capturer une fois** dans une variable :

```vba
Dim frm As Form
Set frm = Forms!frmClients
frm.Caption = "Clients"
frm!txtNom = "Dupont"
frm!txtVille = "Rouen"
```

Cette habitude rejoint les principes d'optimisation du chapitre 18.

## À retenir

- La collection **`Forms`** ne contient que les formulaires **ouverts** ; on y accède par `Forms!Nom`, `Forms("Nom")` (noms dynamiques) ou `Forms(0)`.
- Les contrôles d'un autre formulaire se référencent par la chaîne **`Forms!Formulaire!Contrôle.Propriété`** (`.Value` implicite).
- Utilisez **`Me`** pour son propre formulaire, **`Forms!…`** pour un **autre**.
- **Piège** : référencer un formulaire **fermé** lève une erreur (2450) — vérifiez d'abord avec **`IsLoaded`** (section 4.4).
- Pour un contrôle de **sous-formulaire**, passez par le **contrôle sous-formulaire** puis **`.Form`** (`Forms!parent!ctrlSF.Form!champ`, section 6.4).
- Pour des accès répétés, **capturez** le formulaire dans une variable (`Set frm = Forms!…`).

---


⏭️ [6.3. Formulaires en mode simple, continu et feuille de données](/06-formulaires/03-modes-affichage-formulaire.md)
