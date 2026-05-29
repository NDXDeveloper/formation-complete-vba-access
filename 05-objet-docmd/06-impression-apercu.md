🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.6. Impression et aperçu (PrintOut, OutputTo)

Dernière grande famille « métier » de `DoCmd` : l'**impression** et l'**export** d'objets. Deux méthodes s'en chargent — **`PrintOut`**, qui envoie à l'**imprimante**, et **`OutputTo`**, qui exporte vers un **fichier** (PDF, Excel, RTF…). Cette section les détaille, en prolongeant l'aperçu déjà vu avec `OpenReport` (section 5.2) et la distinction posée en fin de section 5.5.

## Imprimer : PrintOut

```
DoCmd.PrintOut [PrintRange], [PageFrom], [PageTo], [PrintQuality], [Copies], [CollateCopies]
```

`PrintOut` envoie à l'imprimante l'**objet actif** — celui qui a le focus. Elle ne prend **pas** de nom d'objet en argument : on imprime ce qui est ouvert et actif. En revanche, elle permet de préciser la **plage de pages**, la **qualité** et le **nombre de copies**.

```vba
DoCmd.PrintOut                          ' imprime tout, 1 copie
DoCmd.PrintOut acPages, 1, 3, , 2       ' pages 1 à 3, 2 copies
```

Pour **imprimer un état précis**, le plus simple n'est d'ailleurs pas `PrintOut`, mais l'ouverture directe en impression — en se souvenant du « piège » de la section 5.2 (sans aperçu, l'état part à l'imprimante) :

```vba
DoCmd.OpenReport "rptFactures", acViewNormal   ' acViewNormal = impression directe
```

`PrintOut` reste utile lorsqu'on veut contrôler les copies ou les pages, ou imprimer l'objet **déjà actif** (un formulaire, une feuille de données, un aperçu en cours).

## Aperçu avant impression

L'**aperçu** ne relève pas de `PrintOut` mais de l'ouverture de l'état en mode aperçu, déjà rencontrée en section 5.2 :

```vba
DoCmd.OpenReport "rptFactures", acViewPreview   ' aperçu à l'écran
```

## Exporter vers un fichier : OutputTo

```
DoCmd.OutputTo ObjectType, [ObjectName], [OutputFormat], [OutputFile], [AutoStart], [TemplateFile], [Encoding]
```

`OutputTo` exporte un **objet** (état, formulaire, requête, table) vers un **fichier**, dans un format au choix. Ses arguments clés : **`ObjectType`** (`acOutputReport`, `acOutputForm`, `acOutputQuery`, `acOutputTable`…), **`OutputFormat`** (le format de sortie), **`OutputFile`** (le chemin du fichier) et **`AutoStart`** (ouvrir le fichier après création).

```vba
' Exporter une requête vers Excel
DoCmd.OutputTo acOutputQuery, "qryVentes", acFormatXLSX, "C:\Exports\ventes.xlsx"
```

Les formats les plus utilisés :

| Format | Constante |
|---|---|
| PDF | `acFormatPDF` |
| Excel (.xlsx) | `acFormatXLSX` |
| RTF (Word) | `acFormatRTF` |
| Texte | `acFormatTXT` |
| HTML | `acFormatHTML` |

Si l'on omet `ObjectName`, c'est l'objet **actif** qui est exporté ; si l'on omet `OutputFile`, Access **demande** l'emplacement.

## Le cas vedette : générer un PDF

L'usage de loin le plus fréquent d'`OutputTo` est la **génération de PDF** à partir d'un état. Depuis Access 2007, l'export PDF est **natif** : nul besoin de bibliothèque externe.

```vba
' Exporter un état en PDF
DoCmd.OutputTo acOutputReport, "rptFactures", acFormatPDF, "C:\Exports\facture.pdf"

' …et l'ouvrir aussitôt (AutoStart = True)
DoCmd.OutputTo acOutputReport, "rptFactures", acFormatPDF, _
    "C:\Exports\facture.pdf", True
```

C'est la pierre angulaire de nombreux scénarios : produire une facture en PDF, l'archiver, ou l'**envoyer par e-mail**. L'export des états (PDF, RTF, HTML) est détaillé en section 7.8, et la combinaison « état paramétré → PDF → envoi automatisé » en section 7.10 (en s'appuyant sur Outlook, section 22.5).

## PrintOut ou OutputTo ?

| Critère | `PrintOut` | `OutputTo` |
|---|---|---|
| Destination | l'**imprimante** | un **fichier** |
| Objet visé | l'objet **actif** | un objet **désigné par son nom** (ou actif) |
| Formats | — | PDF, Excel, RTF, texte, HTML |
| Usage typique | impression papier | génération de PDF, export de fichiers |

## Sélection de l'imprimante

Par défaut, `PrintOut` et l'impression d'un état utilisent l'**imprimante par défaut**. Pour cibler une imprimante précise, Access expose un objet **`Printer`** (via `Application.Printer` ou la propriété `Printer` d'un état), qui permet de choisir le périphérique et de régler ses paramètres avant impression. Ce contrôle plus fin dépasse le cadre de `DoCmd` ; on le rencontrera avec les états (chapitre 7).

## Précautions

- Comme les autres actions de `DoCmd`, une impression ou un export **annulé** (l'utilisateur ferme la boîte de dialogue) déclenche l'**erreur 2501** (section 5.1) — à prévoir dans la gestion des erreurs (chapitre 13).
- Pour les chemins de fichiers, préférez un chemin relatif construit à partir de `CurrentProject.Path` (section 4.3) plutôt qu'un chemin codé en dur.

## À retenir

- **`PrintOut`** envoie l'objet **actif** à l'**imprimante** (avec contrôle des pages et des copies) ; pour imprimer un état précis, **`OpenReport … acViewNormal`** est la voie directe (section 5.2).
- L'**aperçu** passe par **`OpenReport … acViewPreview`**, pas par `PrintOut`.
- **`OutputTo`** exporte un **objet vers un fichier** (PDF, Excel, RTF, texte, HTML) ; `AutoStart` ouvre le fichier produit.
- **Cas vedette** : la **génération de PDF** (`acFormatPDF`), native depuis 2007, sans outil externe — base de l'archivage et de l'envoi par e-mail (sections 7.8 et 7.10).
- Anticipez l'**erreur 2501** (annulation) et utilisez `CurrentProject.Path` pour les chemins (section 4.3).

---


⏭️ [5.7. Autres méthodes utiles (SetWarnings, Hourglass, Beep, Quit)](/05-objet-docmd/07-autres-methodes-utiles.md)
