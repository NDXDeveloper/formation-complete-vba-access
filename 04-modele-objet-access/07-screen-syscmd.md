🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 4.7. Screen et SysCmd — objets utilitaires souvent oubliés

Pour clore le chapitre, deux objets utilitaires que l'on oublie facilement, mais qui rendent de réels services : **`Screen`**, qui donne accès à l'objet **actif**, et **`SysCmd`**, une fonction « couteau suisse » aux usages variés — barre de progression, messages d'état, informations système. Tous deux complètent la panoplie du modèle objet Access.

## L'objet Screen : l'objet actif

`Screen` est une propriété de l'objet `Application`. Il donne accès à l'objet **qui a le focus** dans l'interface, à l'instant T :

- **`ActiveForm`** — le formulaire actif ;
- **`ActiveReport`** — l'état actif (par exemple en aperçu) ;
- **`ActiveControl`** — le contrôle qui a le focus ;
- **`ActiveDatasheet`** — la feuille de données active ;
- **`PreviousControl`** — le contrôle qui avait le focus juste avant ;
- **`MousePointer`** — la forme du pointeur de souris (lecture/écriture).

```vba
Debug.Print Screen.ActiveForm.Name
Debug.Print Screen.ActiveControl.Name
```

### À quoi sert Screen ?

Son intérêt principal est d'écrire du **code générique**, qui agit sur « ce qui est actif » **sans coder en dur** le nom d'un formulaire ou d'un contrôle. C'est typiquement ce qu'on attend d'une procédure appelée depuis un bouton de ruban ou une barre d'outils (chapitre 17.1) : elle opère sur le formulaire ou le contrôle courant, quel qu'il soit.

```vba
' Une procédure générique : actualiser le formulaire actif, sans connaître son nom
Sub ActualiserFormulaireActif()
    On Error Resume Next              ' au cas où aucun formulaire n'est actif
    Screen.ActiveForm.Requery
End Sub
```

### Attention quand rien n'est actif

Une précaution s'impose : si **aucun** formulaire ou contrôle n'est actif, accéder à `Screen.ActiveForm` ou `Screen.ActiveControl` **déclenche une erreur** (« Aucun formulaire actif » / « Aucun contrôle actif »). Il faut donc soit **intercepter l'erreur** (comme ci-dessus), soit s'assurer du contexte avant d'y accéder. À noter aussi : pour un formulaire principal contenant des sous-formulaires, `Screen.ActiveForm` désigne le formulaire **principal**.

## La fonction SysCmd : un couteau suisse

`SysCmd` est une fonction (méthode d'`Application`) dont le comportement dépend de son **premier argument**, une constante d'action. Elle regroupe plusieurs services sans rapport apparent entre eux ; passons en revue les plus utiles.

### La barre de progression (barre d'état)

L'usage le plus connu : afficher une **barre de progression** dans la barre d'état d'Access, pour un traitement long. Trois étapes — initialiser, mettre à jour, retirer :

```vba
Dim i As Long
SysCmd acSysCmdInitMeter, "Traitement en cours...", 100   ' initialiser (max = 100)
For i = 1 To 100
    ' ... traitement ...
    SysCmd acSysCmdUpdateMeter, i                          ' mettre à jour
Next i
SysCmd acSysCmdRemoveMeter                                 ' retirer
```

C'est la barre de progression **intégrée** à Access. Pour une barre de progression **personnalisée**, intégrée à un formulaire, on construira plutôt son propre indicateur (section 17.5).

### Le texte de la barre d'état

On peut aussi y afficher un simple **message** :

```vba
SysCmd acSysCmdSetStatus, "Enregistrement..."
' ...
SysCmd acSysCmdClearStatus
```

### L'état d'un objet

`SysCmd` offre une façon — plus **ancienne** — de connaître l'état d'un objet (ouvert, modifié, nouveau) :

```vba
' Méthode historique pour tester si un objet est ouvert
If SysCmd(acSysCmdGetObjectState, acForm, "frmClients") <> 0 Then
    Debug.Print "frmClients est ouvert"
End If
```

Comme annoncé en section 4.4, on lui préfère aujourd'hui la propriété **`IsLoaded`** d'un `AccessObject` (`CurrentProject.AllForms("frmClients").IsLoaded`), plus lisible. `SysCmd(acSysCmdGetObjectState, …)` reste néanmoins présent dans beaucoup de code existant.

### Informations système

Enfin, `SysCmd` renseigne sur l'**environnement** d'exécution :

```vba
Debug.Print SysCmd(acSysCmdAccessVer)    ' version d'Access
Debug.Print SysCmd(acSysCmdAccessDir)    ' dossier d'installation d'Access

' Détecter l'exécution sous Access Runtime (utile au déploiement, section 21.2)
If SysCmd(acSysCmdRuntime) Then
    Debug.Print "Exécution sous Access Runtime"
End If
```

La détection du **Runtime** (`acSysCmdRuntime`) est particulièrement précieuse au déploiement : elle permet d'adapter le comportement de l'application selon qu'elle tourne sous la version complète ou sous le Runtime (chapitre 21). On notera que `SysCmd(acSysCmdAccessVer)` fait double emploi avec la propriété `Application.Version` vue en section 4.2 — deux façons d'obtenir la version.

## À retenir

- **`Screen`** donne accès à l'objet **actif** (`ActiveForm`, `ActiveControl`, `ActiveReport`…) ; il sert surtout à écrire du **code générique** agissant sur l'objet courant (par ex. depuis un ruban, chapitre 17.1).
- Accéder à `Screen.ActiveForm`/`ActiveControl` quand **rien n'est actif** lève une **erreur** : prévoir une interception ou une vérification.
- **`SysCmd`** est une fonction polyvalente : **barre de progression** (`acSysCmdInitMeter`/`UpdateMeter`/`RemoveMeter`), **texte d'état** (`acSysCmdSetStatus`), **état d'un objet** (`acSysCmdGetObjectState`) et **infos système**.
- Pour tester si un objet est ouvert, `SysCmd(acSysCmdGetObjectState, …)` existe mais **`IsLoaded`** (section 4.4) est préférable.
- `SysCmd(acSysCmdRuntime)` **détecte le Runtime** (utile au déploiement, chapitre 21) ; `acSysCmdAccessVer` donne la version, comme `Application.Version` (section 4.2).

---


> ✅ **Fin du chapitre 4.** La suite se poursuit avec le chapitre 5 — L'objet DoCmd

⏭️ [5. L'objet DoCmd](/05-objet-docmd/README.md)
