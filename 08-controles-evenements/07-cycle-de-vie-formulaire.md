🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 8.7. Cycle de vie complet d'un formulaire (ordre des événements)

Un formulaire ne s'exécute pas linéairement : il **réagit** à une succession d'événements, dont l'**ordre est précis et déterminant**. Placer un traitement dans `Open` plutôt que dans `Load`, dans `Load` plutôt que dans `Current`, change radicalement son comportement. La grande majorité des bugs en VBA Access — une valeur lue trop tôt, une validation qui ne se déclenche pas, un code exécuté une fois au lieu d'une fois par enregistrement — provient d'une mauvaise compréhension de cette séquence.

Cette section est la **clé de voûte** de la partie consacrée aux événements : elle décrit l'ordre des événements dans les quatre phases de la vie d'un formulaire. Les détails de chaque famille d'événements suivront aux chapitres 8.8 (formulaire) et 8.9 (contrôle).

---

## 8.7.1. Les quatre phases de la vie d'un formulaire

| Phase | Description |
|---|---|
| **Ouverture** | Le formulaire se charge et affiche son premier enregistrement |
| **Navigation** | L'utilisateur passe d'un enregistrement à un autre |
| **Modification / enregistrement** | Une donnée est saisie, puis enregistrée |
| **Fermeture** | Le formulaire est déchargé de la mémoire |

Chaque phase déclenche une suite d'événements dans un ordre fixe, détaillé ci-après.

---

## 8.7.2. L'ouverture : Open → Load → Resize → Activate → Current

À l'ouverture, les événements du formulaire se succèdent ainsi :

```
Open (annulable)  →  Load  →  Resize  →  Activate  →  Current
                                                          │
                                       (premier contrôle)  Enter → GotFocus
```

- **`Open(Cancel)`** : tout premier événement, **avant** que la source de données ne soit interrogée. Annulable.
- **`Load`** : la source est chargée, les contrôles affichent le premier enregistrement. Non annulable.
- **`Resize`** : le formulaire est dimensionné et affiché.
- **`Activate`** : la fenêtre du formulaire devient active.
- **`Current`** : le premier enregistrement devient l'enregistrement courant.

Puis le premier contrôle reçoit le focus (`Enter`, `GotFocus` — chapitre 8.9).

### Open et Load : la distinction décisive

C'est la distinction la plus importante de toute la séquence :

- À l'événement **`Open`**, la source **n'est pas encore interrogée** : on ne peut pas lire les valeurs des contrôles liés. C'est l'endroit pour **définir `RecordSource` / `Filter`**, demander des paramètres, ou **annuler** l'ouverture (chapitres 6.5 et 8.12).
- À l'événement **`Load`**, la source **est chargée** : les contrôles contiennent des données. C'est l'endroit pour **initialiser les contrôles**, fixer des valeurs par défaut d'un contrôle indépendant, ou **positionner** le formulaire sur un enregistrement précis (chapitres 6.6 et 6.8).

```vba
Private Sub Form_Open(Cancel As Integer)
    Me.RecordSource = "SELECT * FROM T_Clients WHERE Actif = True"  ' avant chargement
End Sub

Private Sub Form_Load()
    Me.txtDateDuJour.Value = Date    ' contrôle indépendant, données disponibles
End Sub
```

---

## 8.7.3. La navigation entre enregistrements

Lorsque l'utilisateur passe à un autre enregistrement (après enregistrement d'une éventuelle modification en cours), la séquence est :

```
[Contrôle quitté]      Exit → LostFocus
[Enregistrement]       (si modifié) BeforeUpdate → AfterUpdate
[Formulaire]           Current        ◄── nouvel enregistrement courant
[Nouveau contrôle]     Enter → GotFocus
```

L'événement **`Current`** est ici central : il se déclenche **à chaque** changement d'enregistrement (et à l'ouverture, pour le premier). C'est l'endroit pour adapter l'interface **à l'enregistrement courant** : activer/désactiver des contrôles, ajuster un affichage selon les données de la ligne.

```vba
Private Sub Form_Current()
    ' Verrouiller un champ selon le statut de l'enregistrement courant
    Me.txtMontant.Locked = (Me.Statut.Value = "Validé")
End Sub
```

---

## 8.7.4. La modification et l'enregistrement

Lorsqu'une donnée est saisie puis enregistrée, plusieurs événements interviennent à des niveaux différents :

```
1re modification du record :  … → Dirty (formulaire, une seule fois) → Change → …
Sortie d'un contrôle édité :  BeforeUpdate → AfterUpdate   (niveau CONTRÔLE)
Enregistrement du record :    BeforeUpdate (annulable) → AfterUpdate   (niveau FORMULAIRE)
```

- **`Dirty`** se déclenche **une fois**, à la **première** modification de l'enregistrement (chapitre 8.8).
- À chaque frappe, le contrôle déclenche `Change` (et `KeyDown`/`KeyPress`/`KeyUp` — chapitre 8.9).

### Deux niveaux : contrôle et formulaire

Distinction essentielle, approfondie au chapitre 8.8 :

- Les événements **`BeforeUpdate`/`AfterUpdate` de niveau contrôle** se déclenchent quand un contrôle **modifié perd le focus** : sa valeur est validée dans l'enregistrement en cours.
- Les événements **`BeforeUpdate`/`AfterUpdate` de niveau formulaire** se déclenchent quand l'**enregistrement entier** est sauvegardé (par navigation, fermeture, ou enregistrement explicite).

Le **`BeforeUpdate` du formulaire** est **annulable** : c'est l'emplacement privilégié de la **validation** avant écriture (chapitres 8.10 et 8.12).

```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Len(Nz(Me.txtNom.Value, "")) = 0 Then
        MsgBox "Le nom est obligatoire.", vbExclamation
        Cancel = True          ' empêche l'enregistrement
    End If
End Sub
```

---

## 8.7.5. La fermeture : Unload → Deactivate → Close

À la fermeture, après enregistrement d'une éventuelle modification, la séquence est :

```
[si modifié]   BeforeUpdate → AfterUpdate   (enregistrement)
[Contrôle]     Exit → LostFocus
Unload (annulable)  →  Deactivate  →  Close
```

- **`Unload(Cancel)`** : **annulable** ; c'est l'endroit pour confirmer la fermeture ou vérifier des données non enregistrées (chapitre 6.8).
- **`Deactivate`** : la fenêtre perd son statut actif.
- **`Close`** : **non annulable** ; dernier événement, dédié au **nettoyage** (libération d'objets, réinitialisation de variables).

```vba
Private Sub Form_Unload(Cancel As Integer)
    If Me.Dirty Then
        If MsgBox("Enregistrer les modifications avant de fermer ?", _
                  vbYesNoCancel) = vbCancel Then
            Cancel = True       ' annule la fermeture
        End If
    End If
End Sub
```

> À retenir : `Unload` peut **empêcher** la fermeture, `Close` non.

---

## 8.7.6. Vue d'ensemble

```
OUVERTURE
   Open → Load → Resize → Activate → Current → [Enter/GotFocus du 1er contrôle]

NAVIGATION (vers un autre enregistrement)
   [Exit/LostFocus] → [BeforeUpdate/AfterUpdate si modifié] → Current → [Enter/GotFocus]

MODIFICATION / ENREGISTREMENT
   [Dirty (1 fois) + Change par frappe]
   Sortie d'un contrôle :  BeforeUpdate → AfterUpdate          (contrôle)
   Sauvegarde du record :  BeforeUpdate (annulable) → AfterUpdate   (formulaire)

FERMETURE
   [BeforeUpdate/AfterUpdate si modifié] → Unload (annulable) → Deactivate → Close
```

---

## 8.7.7. Où placer quoi ?

| Objectif | Événement |
|---|---|
| Définir `RecordSource`/`Filter`, annuler l'ouverture | **`Open`** |
| Initialiser les contrôles, positionner sur un enregistrement | **`Load`** |
| Réagir à l'enregistrement courant (interface par enregistrement) | **`Current`** |
| Valider avant l'enregistrement (annulable) | **`BeforeUpdate`** (formulaire) |
| Réagir après un enregistrement réussi | **`AfterUpdate`** (formulaire) |
| Détecter la première modification | **`Dirty`** |
| Confirmer avant la fermeture (annulable) | **`Unload`** |
| Nettoyage final des ressources | **`Close`** |

---

## 8.7.8. Pièges liés à l'ordre des événements

- **Lire des données dans `Open`** : la source n'est pas encore chargée ; manipuler les valeurs dans **`Load`**.
- **Confondre `Load` et `Current`** : `Load` ne se déclenche **qu'une fois** à l'ouverture, `Current` à **chaque** enregistrement. Pour réagir à chaque ligne, utiliser `Current`, jamais `Load` — c'est une erreur très fréquente.
- **Confondre les deux niveaux de `BeforeUpdate`/`AfterUpdate`** : niveau contrôle (sortie d'un contrôle modifié) et niveau formulaire (sauvegarde du record) — chapitre 8.8.
- **Valider dans le mauvais événement** : la validation annulable se place dans le **`BeforeUpdate` du formulaire** (chapitres 8.10 et 8.12).
- **Croire qu'`Activate` ne se déclenche qu'à l'ouverture** : il se déclenche **chaque fois** que le formulaire redevient actif.
- **Tenter d'annuler dans `Close`** : `Close` n'est pas annulable ; c'est **`Unload`** qui l'est.

---

## 8.7.9. Récapitulatif

- L'ordre des événements est **déterminant** ; il s'organise en quatre phases : ouverture, navigation, modification/enregistrement, fermeture.
- **Ouverture** : `Open` (annulable, **avant** chargement des données) → `Load` (données disponibles) → `Resize` → `Activate` → `Current`. La distinction **`Open` / `Load`** est décisive.
- **Navigation** : sortie du contrôle, éventuel enregistrement, puis **`Current`** (déclenché à **chaque** enregistrement) et entrée dans le nouveau contrôle.
- **Modification/enregistrement** : `Dirty` (une fois) puis `Change` par frappe ; **deux niveaux** de `BeforeUpdate`/`AfterUpdate` — contrôle (sortie) et formulaire (sauvegarde). Le **`BeforeUpdate` du formulaire**, annulable, porte la validation.
- **Fermeture** : éventuel enregistrement, puis **`Unload`** (annulable) → `Deactivate` → `Close` (non annulable, nettoyage).
- Le tableau « où placer quoi » (section 8.7.7) résume l'emplacement de chaque traitement ; les détails suivent aux chapitres 8.8 (événements de formulaire), 8.9 (événements de contrôle), 8.10 (validation) et 8.12 (annulation).

⏭️ [8.8. Événements de formulaire (Current, BeforeUpdate, AfterUpdate, Dirty)](/08-controles-evenements/08-evenements-formulaire.md)
