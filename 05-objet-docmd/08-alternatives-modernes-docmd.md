🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.8. Alternatives modernes à DoCmd (méthodes objet directes)

Pour clore le chapitre, prenons du recul. `DoCmd` est l'objet de l'action, hérité du répertoire des macros (section 5.1). Mais, comme ce chapitre l'a régulièrement signalé, **toutes les opérations ne gagnent pas à passer par lui**. Pour beaucoup, manipuler **directement** le modèle objet — une propriété, une méthode de l'objet concerné — est plus clair, plus puissant et plus sûr.

Précisons d'emblée le ton : il ne s'agit **pas de bannir `DoCmd`**, qui reste indispensable, mais de savoir **quand** une voie directe est préférable. Cette section met les deux approches en regard et propose une règle simple pour choisir.

## Pourquoi des alternatives ?

`DoCmd` correspond au monde des macros : ses méthodes sont des **actions** génériques. Les méthodes objet directes, elles, agissent sur l'**objet réel** que l'on tient en main (`Me`, un contrôle, une base). Elles présentent souvent quatre avantages :

- **plus explicites** : le code dit exactement sur quel objet il agit, plutôt que sur « l'objet actif » ;
- **plus puissantes** : on peut lire, modifier, combiner les propriétés finement ;
- **plus sûres** : moins d'effets de bord globaux (pas de `SetWarnings`, par exemple) et un meilleur signalement des erreurs ;
- **plus cohérentes** avec une approche orientée objet (chapitre 16).

## Les principales correspondances

### Exécuter du SQL : db.Execute plutôt que RunSQL

C'est l'exemple phare, déjà développé en section 5.4. `db.Execute` n'affiche aucun message, signale les erreurs (`dbFailOnError`) et renseigne `RecordsAffected`.

```vba
' Voie DoCmd
DoCmd.RunSQL "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02"

' Alternative directe (recommandée)
CurrentDb.Execute "UPDATE tblTarifs SET PrixHT = PrixHT * 1.02", dbFailOnError
```

### Filtrer : Me.Filter plutôt que ApplyFilter

Plutôt que `DoCmd.ApplyFilter`, on peut régler **directement** le filtre du formulaire — ce qui permet aussi de le **lire** et de le **modifier** (section 6.5).

```vba
' Voie DoCmd
DoCmd.ApplyFilter , "ClientID = 5"

' Alternative directe
Me.Filter = "ClientID = 5"
Me.FilterOn = True
```

### Actualiser : Me.Requery plutôt que DoCmd.Requery

La méthode `Requery` de l'objet (formulaire ou contrôle) est plus ciblée et plus lisible.

```vba
' Voie DoCmd
DoCmd.Requery "lstCommandes"

' Alternative directe
Me.lstCommandes.Requery
```

### Donner le focus : SetFocus plutôt que GoToControl

```vba
' Voie DoCmd
DoCmd.GoToControl "txtNom"

' Alternative directe
Me.txtNom.SetFocus
```

### Enregistrer l'enregistrement courant : Me.Dirty = False

Pour forcer l'enregistrement de l'enregistrement en cours, l'idiome `Me.Dirty = False` est plus clair que `DoCmd.RunCommand acCmdSaveRecord` (section 8.8).

```vba
' Voie DoCmd
DoCmd.RunCommand acCmdSaveRecord

' Alternative directe
If Me.Dirty Then Me.Dirty = False
```

### Naviguer : RecordsetClone selon le cas

Pour des déplacements simples, `DoCmd.GoToRecord` convient. Mais pour rechercher puis se positionner sur un enregistrement précis, manipuler le **`RecordsetClone`** du formulaire et son **`Bookmark`** est bien plus puissant (section 9.12).

## Quand DoCmd reste le bon outil

L'équilibre est essentiel : `DoCmd` n'est **pas obsolète**. Pour plusieurs familles d'opérations, il demeure la voie naturelle — et la meilleure :

- **ouvrir et fermer** des objets (`OpenForm`, `OpenReport`, `Close`) : il n'existe pas d'alternative simple ;
- **importer et exporter** des données (`TransferDatabase`, `TransferSpreadsheet`, `TransferText`, `OutputTo`) ;
- **imprimer** (`PrintOut`, `OutputTo`) ;
- exécuter une **commande intégrée** (`RunCommand`).

Pour ces tâches, chercher à « éviter » `DoCmd` serait artificiel et contre-productif.

## L'heuristique

On peut résumer le choix en une règle de bon sens, selon la **nature** de l'opération :

- **opérations sur les données** (exécuter du SQL, manipuler des enregistrements) → privilégier **DAO / méthodes directes** (`db.Execute`, `RecordsetClone`) ;
- **modification de l'état d'un objet** (filtre, actualisation, focus, enregistrement) → privilégier la **propriété / méthode de l'objet** lui-même (`Me.Filter`, `Me.Requery`, `SetFocus`, `Me.Dirty`) ;
- **ouverture / fermeture / impression / import-export** → **`DoCmd`** reste le choix naturel.

## Tableau de synthèse

| Opération | Voie `DoCmd` | Alternative directe | Préférence |
|---|---|---|---|
| Exécuter du SQL action | `DoCmd.RunSQL` | `db.Execute … dbFailOnError` | **directe** (5.4) |
| Filtrer un formulaire | `DoCmd.ApplyFilter` | `Me.Filter` + `Me.FilterOn` | **directe** (6.5) |
| Actualiser | `DoCmd.Requery` | `Me.Requery` / `ctl.Requery` | **directe** |
| Donner le focus | `DoCmd.GoToControl` | `Me.txtNom.SetFocus` | **directe** |
| Enregistrer l'enreg. courant | `DoCmd.RunCommand acCmdSaveRecord` | `Me.Dirty = False` | **directe** (8.8) |
| Naviguer dans les enreg. | `DoCmd.GoToRecord` | `RecordsetClone` + `Bookmark` | selon le cas (9.12) |
| Ouvrir / fermer un objet | `DoCmd.OpenForm` / `Close` | *(pas d'équivalent simple)* | **`DoCmd`** |
| Importer / exporter | `DoCmd.Transfer…` / `OutputTo` | *(DAO/ADO possible mais lourd)* | **`DoCmd`** |

## À retenir

- `DoCmd` reprend les **actions** des macros ; pour de nombreuses opérations, une **méthode objet directe** est plus **explicite**, **puissante** et **sûre**.
- Correspondances clés : **`db.Execute`** (vs `RunSQL`, section 5.4), **`Me.Filter`/`FilterOn`** (vs `ApplyFilter`), **`Me.Requery`**, **`SetFocus`** (vs `GoToControl`), **`Me.Dirty = False`** (vs `acCmdSaveRecord`, section 8.8).
- `DoCmd` **reste le bon outil** pour **ouvrir/fermer**, **importer/exporter**, **imprimer** et exécuter des **commandes intégrées** : ne pas chercher à l'éviter là où il excelle.
- **Heuristique** : données → DAO/direct ; état d'un objet → méthode de l'objet ; ouverture/fermeture/impression/transfert → `DoCmd`.

---


> ✅ **Fin du chapitre 5.** La suite se poursuit avec le chapitre 6 — Formulaires (Forms)

⏭️ [6. Formulaires (Forms)](/06-formulaires/README.md)
