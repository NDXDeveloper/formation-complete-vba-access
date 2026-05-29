🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.7. Autres méthodes utiles (SetWarnings, Hourglass, Beep, Quit)

Au-delà des grandes familles d'actions, `DoCmd` propose une série de méthodes **utilitaires** que l'on utilise au quotidien. Cette section présente les quatre du titre — `SetWarnings`, `Hourglass`, `Beep`, `Quit` — puis quelques autres. Un fil rouge les relie : plusieurs modifient un **état global** d'Access qu'il faut penser à **restaurer**.

## SetWarnings — activer/désactiver les messages système

```vba
DoCmd.SetWarnings False   ' désactive les messages de confirmation
DoCmd.SetWarnings True    ' les rétablit
```

`SetWarnings` contrôle l'affichage des **messages de confirmation** d'Access (« Vous allez mettre à jour N ligne(s)… », confirmations de suppression, d'exécution de requête action…). On la rencontre surtout autour des requêtes action lancées par `RunSQL` ou `OpenQuery` (section 5.4).

Comme déjà souligné en section 5.4, c'est un **état global** et son usage est **risqué** : si une erreur survient entre la désactivation et la réactivation, le code peut sortir avant le `SetWarnings True`, et les messages restent **désactivés pour toute la session Access**. Deux parades :

- **idéalement, s'en passer** : la méthode **`db.Execute`** (DAO, section 5.4) n'affiche aucun message — donc rien à supprimer ;
- **sinon, toujours restaurer** dans un gestionnaire d'erreurs, comme illustré plus bas.

## Hourglass — le curseur d'attente

```vba
DoCmd.Hourglass True      ' affiche le curseur d'attente (sablier)
' ... traitement long ...
DoCmd.Hourglass False     ' rétablit le curseur normal
```

`Hourglass` affiche ou masque le **curseur d'attente** pendant une opération longue, pour signaler à l'utilisateur que l'application travaille. Là encore, c'est un **état à restaurer** : oublier `Hourglass False` (par exemple à cause d'une erreur) laisse le sablier affiché indéfiniment. Pour un contrôle plus fin du pointeur (autres formes), on dispose de `Screen.MousePointer` (section 4.7).

## Beep — un signal sonore

```vba
DoCmd.Beep                ' joue le son système par défaut
```

`Beep` émet le **signal sonore** par défaut de Windows. Simple, il sert de retour audio discret — par exemple pour souligner une erreur de saisie ou la fin d'un traitement. À utiliser avec parcimonie.

## Quit — quitter Access

```vba
DoCmd.Quit acQuitSaveNone     ' quitter sans enregistrer les modifications de conception
```

`DoCmd.Quit` **ferme Access**. Son argument `Options` détermine le sort des modifications :

- **`acQuitPrompt`** — demander à l'utilisateur ;
- **`acQuitSaveAll`** — tout enregistrer (valeur par défaut) ;
- **`acQuitSaveNone`** — ne rien enregistrer.

Comme pour la méthode `Close` (section 5.2), ces options concernent les modifications de **conception** des objets, **pas les données** (déjà persistées automatiquement, rappel de la section 1.2). `DoCmd.Quit` fait par ailleurs **double emploi** avec `Application.Quit` (section 4.2).

## Le réflexe commun : restaurer l'état global

`SetWarnings`, `Hourglass` (et `Echo`, vu plus bas) ont un point commun fondamental : elles modifient un **état global** d'Access. Si une erreur interrompt le code **avant** la restauration, cet état reste figé — messages désactivés, sablier bloqué, écran gelé. Le **réflexe** à adopter est donc de **toujours restaurer** ces états, idéalement via un gestionnaire d'erreurs avec une étiquette de sortie commune (chapitre 13).

```vba
Sub TraitementSensible()
    On Error GoTo Gestion
    DoCmd.SetWarnings False
    DoCmd.Hourglass True

    DoCmd.RunSQL "DELETE FROM tblTemp"     ' opération sans message
    ' ... autres traitements ...

Sortie:
    DoCmd.Hourglass False                  ' restauré quoi qu'il arrive
    DoCmd.SetWarnings True
    Exit Sub

Gestion:
    MsgBox Err.Description
    Resume Sortie                          ' passe par Sortie -> états rétablis
End Sub
```

Le point clé est le `Resume Sortie` : en cas d'erreur, le code **repasse par l'étiquette `Sortie`**, garantissant la restauration de l'état global. C'est l'un des bénéfices concrets d'une gestion d'erreurs structurée (chapitre 13).

## Quelques autres méthodes utiles

`DoCmd` regorge d'autres méthodes ; en voici quelques-unes fréquemment employées.

| Méthode | Rôle | Détail |
|---|---|---|
| `Echo` | activer / suspendre le rafraîchissement de l'écran | performance (chapitre 18) ; état à restaurer |
| `RunCommand` | exécuter une **commande intégrée** d'Access | constantes `acCmd…` |
| `RunMacro` | exécuter une macro par son nom | — |
| `CancelEvent` | annuler l'événement en cours | en VBA, préférer l'argument `Cancel` (section 8.12) |
| `Maximize` / `Minimize` / `Restore` | dimensionner la fenêtre | usage hérité |

La méthode **`RunCommand`** mérite une mention : elle déclenche n'importe quelle **commande intégrée** d'Access (l'équivalent d'un clic dans un menu), via des constantes `acCmd…`. Par exemple, enregistrer l'enregistrement courant :

```vba
DoCmd.RunCommand acCmdSaveRecord     ' enregistrer l'enregistrement en cours
```

Quant à **`Echo`**, elle fonctionne comme `Application.Echo` (section 4.2) pour figer l'affichage le temps d'un traitement lourd — et relève du même réflexe de **restauration** que `SetWarnings` et `Hourglass`.

## À retenir

- **`SetWarnings`** désactive/active les **messages système** ; c'est un **état global risqué** — préférez **`db.Execute`** (section 5.4) ou restaurez impérativement dans un gestionnaire d'erreurs.
- **`Hourglass`** affiche le **curseur d'attente** (à rétablir après le traitement) ; `Screen.MousePointer` offre un contrôle plus fin (section 4.7).
- **`Beep`** émet un **signal sonore** simple ; **`Quit`** ferme Access (options de **conception**, pas de données ; équivaut à `Application.Quit`, section 4.2).
- **Réflexe commun** : `SetWarnings`, `Hourglass` et `Echo` modifient un **état global** — restaurez-le **quoi qu'il arrive**, via une étiquette de sortie et `Resume` (chapitre 13).
- Autres méthodes utiles : **`Echo`** (performance, chapitre 18), **`RunCommand`** (commandes intégrées `acCmd…`), `RunMacro`, `CancelEvent` (préférer `Cancel`, section 8.12).

---


⏭️ [5.8. Alternatives modernes à DoCmd (méthodes objet directes)](/05-objet-docmd/08-alternatives-modernes-docmd.md)
