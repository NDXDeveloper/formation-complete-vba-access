🔝 Retour au [Sommaire](/SOMMAIRE.md)

# A. Référence des propriétés et méthodes DoCmd

Cette annexe récapitule l'objet `DoCmd`, point d'entrée central de l'automatisation dans Access. Elle complète le **chapitre 5** : ce dernier explique le rôle et le bon usage de `DoCmd`, tandis que cette annexe sert d'aide-mémoire pour retrouver rapidement une méthode, sa syntaxe et ses arguments.

## DoCmd : un objet de méthodes, sans propriétés

Contrairement à la plupart des objets du modèle Access (`Form`, `Report`, `Control`…), `DoCmd` **n'expose aucune propriété manipulable** par le développeur. Ce n'est pas un objet de données : c'est un objet utilitaire qui regroupe, sous forme de méthodes, l'équivalent des actions de macro d'Access. On y accède via `Application.DoCmd` (le préfixe `Application` étant implicite), puis on appelle directement une de ses méthodes :

```vba
DoCmd.OpenForm "frmClients"
DoCmd.Close acForm, "frmClients", acSaveNo
```

L'intitulé « propriétés et méthodes » de cette annexe ne couvre donc, en pratique, que des méthodes.

## Comment lire cette référence

- Les **arguments entre crochets** `[...]` sont facultatifs. Lorsqu'ils sont omis, Access applique une valeur par défaut (indiquée dans le texte ou les tableaux de constantes).
- Les valeurs reposent presque toujours sur des **constantes intrinsèques** (`acForm`, `acNormal`, `acDialog`…). Il faut toujours préférer ces constantes aux valeurs numériques : le code reste lisible et indépendant des numéros internes. Les valeurs chiffrées ne sont indiquées ci-dessous qu'à titre informatif.
- En VBA, les arguments d'une méthode appelée **sans parenthèses** (instruction) sont séparés par des virgules. Pour conserver un argument intermédiaire à sa valeur par défaut, on laisse simplement sa position vide entre deux virgules.

```vba
' On saute FilterName et on ne fournit que WhereCondition
DoCmd.OpenForm "frmClients", , , "Ville = 'Rouen'"
```

## Constantes intrinsèques fréquentes

### Type d'objet — `AcObjectType`

Utilisé par `Close`, `DeleteObject`, `CopyObject`, `OutputTo`, `SelectObject`, etc.

| Constante | Valeur | Objet |
|---|---|---|
| `acTable` | 0 | Table |
| `acQuery` | 1 | Requête |
| `acForm` | 2 | Formulaire |
| `acReport` | 3 | État |
| `acMacro` | 4 | Macro |
| `acModule` | 5 | Module |
| `acStoredProcedure` | 9 | Procédure stockée (projet ADP, obsolète) |
| `acFunction` | 10 | Fonction (projet ADP, obsolète) |
| `acDefault` | -1 | Objet actif / par défaut |

### Vue d'ouverture — `AcView` / `AcFormView`

`OpenReport`, `OpenQuery`, `OpenTable` utilisent `AcView` ; `OpenForm` utilise `AcFormView`. Les valeurs principales :

| Constante | Valeur | Vue |
|---|---|---|
| `acNormal` / `acViewNormal` | 0 | Mode normal (ouverture pour l'utilisateur) |
| `acDesign` / `acViewDesign` | 1 | Mode Création |
| `acPreview` / `acViewPreview` | 2 | Aperçu avant impression |
| `acFormDS` | 3 | Feuille de données (formulaire) |
| `acViewReport` | 5 | Mode État (report view) |
| `acLayout` / `acViewLayout` | 6 | Mode Page |

> ⚠️ Les vues PivotTable et PivotChart (`acFormPivotTable`, `acViewPivotChart`…) sont **obsolètes** depuis Access 2013 et ne doivent plus être utilisées.

### Mode de données — `AcFormOpenDataMode` (OpenForm)

| Constante | Valeur | Effet |
|---|---|---|
| `acFormPropertySettings` | -1 | Reprend les propriétés du formulaire (défaut) |
| `acFormAdd` | 0 | Ouverture en saisie (nouvel enregistrement) |
| `acFormEdit` | 1 | Lecture **et** modification |
| `acFormReadOnly` | 2 | Lecture seule |

### Mode fenêtre — `AcWindowMode`

| Constante | Valeur | Effet |
|---|---|---|
| `acWindowNormal` | 0 | Affichage normal (défaut) |
| `acHidden` | 1 | Ouvert mais masqué |
| `acIcon` | 2 | Réduit en icône |
| `acDialog` | 3 | Modal — **suspend l'exécution du code** jusqu'à fermeture/masquage |

### Déplacement d'enregistrement — `AcRecord` (GoToRecord)

| Constante | Valeur | Déplacement |
|---|---|---|
| `acPrevious` | 0 | Précédent |
| `acNext` | 1 | Suivant (défaut) |
| `acFirst` | 2 | Premier |
| `acLast` | 3 | Dernier |
| `acGoTo` | 4 | Numéro précis (via `Offset`) |
| `acNewRec` | 5 | Nouvel enregistrement |

### Enregistrement à la fermeture — `AcCloseSave` (Close)

| Constante | Valeur | Effet |
|---|---|---|
| `acSavePrompt` | 0 | Demande à l'utilisateur (défaut) |
| `acSaveYes` | 1 | Enregistre sans demander |
| `acSaveNo` | 2 | Ferme sans enregistrer les modifications de conception |

### Sens du transfert — `AcDataTransferType`

Utilisé par `TransferDatabase` et `TransferSpreadsheet`.

| Constante | Valeur | Sens |
|---|---|---|
| `acImport` | 0 | Importation vers Access |
| `acExport` | 1 | Exportation depuis Access |
| `acLink` | 2 | Liaison (table liée) |

### Formats de sortie — `OutputTo` / `SendObject`

Constantes principales encore prises en charge : `acFormatPDF`, `acFormatXLSX`, `acFormatXLS`, `acFormatRTF` (Word), `acFormatTXT` (texte), `acFormatHTML`, `acFormatXPS`. Le format `acFormatSNP` (Snapshot) est **obsolète**.

### Options de fermeture de l'application — `AcQuitOption` (Quit)

| Constante | Valeur | Effet |
|---|---|---|
| `acQuitPrompt` | 0 | Demande pour chaque objet non enregistré |
| `acQuitSaveAll` | 1 | Enregistre tout puis quitte (défaut) |
| `acQuitSaveNone` | 2 | Quitte sans enregistrer |

---

## Ouverture et fermeture d'objets

- **`OpenForm`** `(FormName, [View], [FilterName], [WhereCondition], [DataMode], [WindowMode], [OpenArgs])`
  Ouvre un formulaire. `WhereCondition` filtre les enregistrements (clause SQL sans le mot `WHERE`) ; `DataMode` impose le mode de saisie ; `WindowMode = acDialog` rend le formulaire modal ; `OpenArgs` transmet une chaîne récupérable via la propriété `OpenArgs` du formulaire.
- **`OpenReport`** `(ReportName, [View], [FilterName], [WhereCondition], [WindowMode], [OpenArgs])`
  Ouvre un état. Préférer `View = acViewPreview` pour l'aperçu et `acViewNormal` pour l'impression directe.
- **`OpenQuery`** `(QueryName, [View], [DataMode])`
  Ouvre une requête sauvegardée. Pour une requête SELECT, affiche le résultat en feuille de données ; pour une requête action, l'exécute.
- **`OpenTable`** `(TableName, [View], [DataMode])`
  Ouvre une table en feuille de données (à éviter dans une application livrée à l'utilisateur final).
- **`OpenModule`** `([ModuleName], [ProcedureName])`
  Ouvre un module dans l'éditeur VBA, éventuellement positionné sur une procédure.
- **`Close`** `([ObjectType], [ObjectName], [Save])`
  Ferme l'objet indiqué. Sans argument, ferme l'objet actif. Toujours préciser `acSaveNo` pour éviter d'enregistrer involontairement des modifications de conception.

## Navigation, recherche et filtres

- **`GoToRecord`** `([ObjectType], [ObjectName], [Record], [Offset])`
  Se déplace dans les enregistrements de l'objet (formulaire, table…). `acNewRec` place sur un nouvel enregistrement.
- **`GoToControl`** `(ControlName)`
  Donne le focus à un contrôle du formulaire actif.
- **`GoToPage`** `([PageNumber], [Right], [Down])`
  Affiche une page donnée d'un formulaire multi-pages.
- **`FindRecord`** `(FindWhat, [Match], [MatchCase], [Search], [SearchAsFormatted], [OnlyCurrentField], [FindFirst])`
  Recherche une valeur dans le formulaire actif (équivalent de la boîte de dialogue Rechercher).
- **`FindNext`**
  Recherche l'occurrence suivante du dernier `FindRecord`.
- **`ApplyFilter`** `([FilterName], [WhereCondition], [ControlName])`
  Applique un filtre au formulaire/état actif. `WhereCondition` est une clause SQL sans `WHERE`.
- **`SearchForRecord`** `([ObjectType], [ObjectName], [Record], [WhereCondition])`
  (Access 2010+) Recherche par condition, plus robuste que `FindRecord` pour le code.
- **`ShowAllRecords`**
  Supprime tous les filtres appliqués à l'objet actif.

## Exécution d'actions et de requêtes

- **`RunSQL`** `(SQLStatement, [UseTransaction])`
  Exécute une **requête action** (INSERT, UPDATE, DELETE, requête DDL). N'accepte pas de requête SELECT. Affiche des messages de confirmation par défaut (voir *Points de vigilance*).
- **`RunQuery`** *(n'existe pas)* — pour exécuter une requête sauvegardée par code, utiliser `OpenQuery`, ou `CurrentDb.Execute` / `QueryDef.Execute` (chapitres 11 et 12).
- **`RunMacro`** `(MacroName, [RepeatCount], [RepeatExpression])`
  Lance une macro Access nommée.
- **`RunCommand`** `(Command)`
  Exécute une commande intégrée d'Access via une constante `AcCommand` (ex. `acCmdSaveRecord`, `acCmdDeleteRecord`, `acCmdFind`, `acCmdRefresh`, `acCmdUndo`, `acCmdPrint`). La liste complète des constantes `acCmd…` est consultable dans l'Explorateur d'objets (F2).
- **`RunDataMacro`** `(MacroName)`
  (Access 2010+) Exécute une macro de données nommée.
- **`CancelEvent`**
  Annule l'événement en cours (équivalent de `Cancel = True` dans certaines macros). En code VBA, on privilégie l'argument `Cancel` de l'événement (chapitre 8.12).

## Import / export de données

- **`TransferDatabase`** `([TransferType], [DatabaseType], [DatabaseName], [ObjectType], [Source], [Destination], [StructureOnly], [StoreLogin])`
  Importe, exporte ou **lie** des objets depuis/vers une autre base (Access, ODBC…). Sert notamment à créer des tables liées par code.
- **`TransferSpreadsheet`** `([TransferType], [SpreadsheetType], [TableName], [FileName], [HasFieldNames], [Range], [UseOA])`
  Échange des données avec un classeur Excel. `HasFieldNames = True` indique que la première ligne contient les en-têtes ; `Range` cible une feuille ou une plage nommée.
- **`TransferText`** `([TransferType], [SpecificationName], [TableName], [FileName], [HasFieldNames], [HTMLTableName], [CodePage])`
  Échange des données avec un fichier texte (délimité, à largeur fixe, HTML). `SpecificationName` réutilise une spécification d'import/export enregistrée.
- **`RunSavedImportExport`** `(SavedImportExportName)`
  (Access 2010+) Rejoue une opération d'import/export sauvegardée.
- **`OutputTo`** `([ObjectType], [ObjectName], [OutputFormat], [OutputFile], [AutoStart], [TemplateFile], [Encoding], [OutputQuality])`
  Exporte un objet Access vers un fichier (PDF, XLSX, RTF, TXT, HTML…). `AutoStart = True` ouvre le fichier généré dans l'application associée. Voie privilégiée pour générer un PDF d'état (chapitre 7.8).

## Impression et génération de fichiers

- **`PrintOut`** `([PrintRange], [PageFrom], [PageTo], [PrintQuality], [Copies], [CollateCopies])`
  Imprime l'objet actif sur l'imprimante par défaut. `PrintRange = acPages` permet de cibler une plage de pages via `PageFrom`/`PageTo`.
- **`OutputTo`** — voir ci-dessus : pour produire un fichier plutôt qu'une impression papier.
- **`SendObject`** `([ObjectType], [ObjectName], [OutputFormat], [To], [Cc], [Bcc], [Subject], [MessageText], [EditMessage], [TemplateFile])`
  Envoie un objet Access en pièce jointe d'un e-mail via le client de messagerie par défaut. Pour un contrôle fin de l'envoi, préférer l'automation d'Outlook (chapitre 22.5).

## Gestion des objets (copie, suppression, renommage)

- **`CopyObject`** `([DestinationDatabase], [NewName], [SourceObjectType], [SourceObjectName])`
  Copie un objet (dans la base courante ou vers une autre base).
- **`DeleteObject`** `([ObjectType], [ObjectName])`
  Supprime un objet de la base. Sans argument, supprime l'objet sélectionné.
- **`Rename`** `(NewName, [ObjectType], [OldName])`
  Renomme un objet.
- **`Save`** `([ObjectType], [ObjectName])`
  Enregistre l'objet (sa conception). Sans argument, enregistre l'objet actif.
- **`SelectObject`** `(ObjectType, [ObjectName], [InDatabaseWindow])`
  Sélectionne un objet, ouvert ou dans le volet de navigation.
- **`RepaintObject`** `([ObjectType], [ObjectName])`
  Force le rafraîchissement de l'affichage d'un objet (redessine les contrôles).
- **`Requery`** `([ControlName])`
  Met à jour les données de l'objet actif ou d'un contrôle, en réexécutant sa source.

## Contrôle de l'interface et de l'affichage

- **`Maximize`** / **`Minimize`** / **`Restore`**
  Agrandit, réduit ou restaure la fenêtre de l'objet actif (pertinent surtout en mode fenêtres superposées).
- **`MoveSize`** `([Right], [Down], [Width], [Height])`
  Déplace et dimensionne la fenêtre active. **Valeurs exprimées en twips** (1 cm ≈ 567 twips).
- **`Hourglass`** `(HourglassOn)`
  Affiche (`True`) ou masque (`False`) le sablier pendant un traitement long.
- **`Echo`** `(EchoOn, [StatusBarText])`
  Active/désactive le rafraîchissement de l'écran. `Echo False` accélère les traitements lourds en masquant les redessins intermédiaires.
- **`SetWarnings`** `(WarningsOn)`
  Active/désactive les messages système de confirmation (requêtes action, suppressions…).
- **`Beep`**
  Émet le son système par défaut.
- **`LockNavigationPane`** `(Lock)`
  Verrouille le volet de navigation pour empêcher la suppression d'objets.
- **`SetProperty`** `(ControlName, Property, Value)`
  Modifie une propriété d'un contrôle du formulaire actif via une constante `AcProperty` (`acPropertyEnabled`, `acPropertyVisible`, `acPropertyLocked`, `acPropertyForeColor`, `acPropertyValue`…). En VBA pur, on accède plus directement à la propriété (`Me.MonControle.Visible = False`).

## Application et navigation

- **`Quit`** `([Options])`
  Ferme l'application Access selon l'option de sauvegarde choisie.
- **`SetParameter`** `(Name, Expression)`
  (Access 2010+) Définit un paramètre transmis à l'objet ouvert ensuite (état, requête).
- **`NavigateTo`** `([Category], [Group])`
  Pilote l'affichage du volet de navigation (catégorie, groupe).
- **`BrowseTo`** `(ObjectType, ObjectName, [PathToSubformControl], [WhereCondition], [DataMode], [OpenArgs])`
  Charge un objet dans un contrôle de sous-formulaire, notamment dans les **formulaires de navigation** (chapitre 17.3).
- **`RefreshRecord`**
  (Access 2010+) Réactualise l'enregistrement courant du formulaire.
- **`ClearMacroError`**
  Efface les informations d'erreur de macro (objet `MacroError`).

---

## Points de vigilance

> ⚠️ **`SetWarnings False` doit toujours être réactivé.** Si une erreur interrompt le code entre `DoCmd.SetWarnings False` et `DoCmd.SetWarnings True`, les avertissements restent désactivés pour toute la session. Réactiver impérativement dans le gestionnaire d'erreurs (label de nettoyage). Mieux : éviter `RunSQL` + `SetWarnings` et préférer `CurrentDb.Execute …, dbFailOnError`, qui n'affiche aucun message et signale proprement les erreurs (chapitres 11 et 18.3).

> ⚠️ **`Echo False` et `Hourglass True` doivent eux aussi être réinitialisés.** Un écran figé ou un sablier persistant après une erreur est un symptôme classique d'un `Echo True` / `Hourglass False` jamais atteint. Les replacer dans la routine de nettoyage.

> ⚠️ **`OpenForm … acDialog` suspend le code appelant.** Tant que le formulaire modal n'est pas fermé **ou masqué** (`Visible = False`), l'instruction suivante n'est pas exécutée. C'est le mécanisme de base des boîtes de dialogue personnalisées (chapitre 6.7) : le formulaire se masque au lieu de se fermer pour que le code appelant puisse lire ses contrôles.

> ⚠️ **`RunSQL` n'exécute que des requêtes action.** Pour une requête SELECT, il faut ouvrir un `Recordset` (chapitre 9) et non utiliser `RunSQL`.

## Méthodes héritées et obsolètes

Présentes pour la compatibilité, mais à ne plus employer dans du code neuf :

- **`DoMenuItem`** — remplacée par `RunCommand`.
- **`ShowToolbar`**, **`AddMenu`**, **`SetMenuItem`** — liées aux anciennes barres de commandes/menus ; le ruban personnalisé (chapitre 17.1) est la voie actuelle.
- **`OpenDataAccessPage`** — les Data Access Pages ne sont plus prises en charge.
- **`OpenView`**, **`OpenStoredProcedure`**, **`OpenFunction`** — spécifiques aux projets ADP (`.adp`), abandonnés depuis Access 2013.

---

> 💡 **À retenir.** `DoCmd` reste l'outil le plus direct pour les actions globales (ouvrir/fermer un objet, importer/exporter, imprimer). Pour manipuler les **propriétés** d'un formulaire, d'un état ou d'un contrôle, on passe directement par les objets concernés plutôt que par `DoCmd` (voir les alternatives modernes, chapitre 5.8). Pour la liste exhaustive des constantes (`AcCommand`, formats, vues…), l'**Explorateur d'objets** de l'éditeur VBA (`F2`) fait foi.

⏭️ [B. Propriétés essentielles des formulaires et états](/annexes/b-proprietes-formulaires-etats.md)
