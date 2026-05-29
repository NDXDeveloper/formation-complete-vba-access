🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 7.8. Export d'états (PDF, RTF, HTML) par DoCmd.OutputTo

Un état n'est pas seulement destiné à l'écran ou à l'imprimante : il faut souvent le **transformer en fichier** pour l'archiver ou le transmettre. La méthode `DoCmd.OutputTo` produit, en une instruction, un fichier PDF, RTF, HTML ou autre à partir d'un état. C'est l'outil central de l'export et la brique de base de la diffusion automatisée (chapitre 7.10).

Cette section détaille `OutputTo`, ses formats, le cas particulier de l'export d'un état **filtré**, ainsi que la gestion des noms de fichiers et des erreurs.

---

## 7.8.1. La méthode DoCmd.OutputTo

La signature comporte plusieurs arguments, dont la plupart sont facultatifs :

```vba
DoCmd.OutputTo ObjectType, [ObjectName], [OutputFormat], [OutputFile], _
               [AutoStart], [TemplateFile], [Encoding], [OutputQuality]
```

| Argument | Rôle |
|---|---|
| `ObjectType` | Type d'objet : `acOutputReport` (aussi `acOutputForm`, `acOutputQuery`, `acOutputTable`…) |
| `ObjectName` | Nom de l'état à exporter |
| `OutputFormat` | Format de sortie : `acFormatPDF`, `acFormatRTF`, `acFormatHTML`… |
| `OutputFile` | Chemin complet du fichier ; si omis, Access **demande** l'emplacement |
| `AutoStart` | `True` pour **ouvrir** le fichier après l'export |
| `TemplateFile` | Modèle (pour l'export HTML) |
| `Encoding` | Encodage (ex. `acUTF8`) |
| `OutputQuality` | Qualité PDF : `acExportQualityPrint` ou `acExportQualityScreen` |

> **Attention** : `OutputTo` **écrase** un fichier existant **sans avertissement**. Il faut donc gérer soi-même toute confirmation si nécessaire.

---

## 7.8.2. Les formats d'export

| Constante | Format | Remarque |
|---|---|---|
| `acFormatPDF` | PDF | **Recommandé** : haute fidélité, mise en page figée |
| `acFormatRTF` | RTF (Word) | Éditable, mais fidélité de mise en page limitée |
| `acFormatHTML` | HTML | Web ; mise en page très approximative |
| `acFormatXLSX` | Excel | Médiocre depuis un **état** (préférer l'export d'une requête) |
| `acFormatTXT` | Texte | Données brutes |
| `acFormatSNP` | Snapshot | Obsolète, à éviter |

Pour un rendu fidèle, le **PDF** est de loin le meilleur choix. Les autres formats déforment d'autant plus la mise en page que l'état est complexe (voir 7.8.9).

---

## 7.8.3. Export PDF : le cas le plus courant

L'export PDF se réduit à une ligne :

```vba
DoCmd.OutputTo acOutputReport, "E_Factures", acFormatPDF, _
               "C:\Exports\Factures.pdf"
```

Si l'on omet le chemin, Access ouvre une boîte de dialogue d'enregistrement :

```vba
DoCmd.OutputTo acOutputReport, "E_Factures", acFormatPDF
```

---

## 7.8.4. Exporter un état filtré : ouvrir d'abord, exporter ensuite

Point **essentiel** : `OutputTo` n'accepte **pas** d'argument `WhereCondition`. On ne peut donc pas lui passer directement un filtre. Pour exporter un état restreint à un périmètre (une facture, un client), on exploite le fait qu'`OutputTo` exporte l'**instance déjà ouverte** d'un état : on ouvre d'abord l'état filtré (de préférence **caché**), puis on l'exporte, puis on le referme.

```vba
Dim strChemin As String
strChemin = "C:\Exports\Facture_" & lngFactureID & ".pdf"

' 1. Ouvrir l'état filtré, en fenêtre cachée
DoCmd.OpenReport "E_Factures", acViewPreview, , "FactureID=" & lngFactureID, acHidden

' 2. Exporter l'instance ouverte (le filtre est pris en compte)
DoCmd.OutputTo acOutputReport, "E_Factures", acFormatPDF, strChemin

' 3. Refermer
DoCmd.Close acReport, "E_Factures"
```

C'est le pattern de référence pour produire des PDF **paramétrés**, à la base de l'envoi automatisé (chapitre 7.10). Le filtrage par `WhereCondition` à l'ouverture est détaillé au chapitre 7.4.

---

## 7.8.5. Nom de fichier dynamique et sécurisé

On compose généralement le nom de fichier à partir de données (date, identifiant, nom de client). Deux précautions s'imposent : construire un **chemin valide** et **purger** les caractères interdits dans un nom de fichier.

```vba
Function NomFichierValide(ByVal s As String) As String
    Dim caracteres As Variant, c As Variant
    caracteres = Array("/", "\", ":", "*", "?", """", "<", ">", "|")
    For Each c In caracteres
        s = Replace(s, c, "_")
    Next c
    NomFichierValide = s
End Function
```

```vba
Dim strDossier As String, strChemin As String
strDossier = Environ("USERPROFILE") & "\Documents\"
strChemin = strDossier & NomFichierValide("Facture_" & strClient & "_" & _
            Format(Date, "yyyy-mm-dd")) & ".pdf"
```

Le dossier `CurrentProject.Path` permet aussi de générer les fichiers à côté de l'application.

---

## 7.8.6. Ouvrir le fichier après export (AutoStart)

L'argument `AutoStart` à `True` ouvre le fichier produit dans l'application associée (le lecteur PDF par défaut, par exemple) :

```vba
DoCmd.OutputTo acOutputReport, "E_Factures", acFormatPDF, strChemin, True
```

Cela suppose qu'une application soit bien associée au format sur le poste.

---

## 7.8.7. Gestion des erreurs à l'export

Plusieurs incidents peuvent survenir : fichier déjà ouvert (verrouillé), chemin invalide, ou état sans données. Si l'état est vide et que son événement `NoData` annule l'affichage, `OutputTo` déclenche l'**erreur 2501** (chapitre 7.9). Un traitement robuste intercepte ces cas et garantit la fermeture de l'état caché :

```vba
Sub ExporterFacturePDF(ByVal lngID As Long)
    Dim strChemin As String
    On Error GoTo Gestion

    strChemin = Environ("USERPROFILE") & "\Documents\Facture_" & lngID & ".pdf"

    DoCmd.OpenReport "E_Factures", acViewPreview, , "FactureID=" & lngID, acHidden
    DoCmd.OutputTo acOutputReport, "E_Factures", acFormatPDF, strChemin, True
    DoCmd.Close acReport, "E_Factures"
    Exit Sub

Gestion:
    Select Case Err.Number
        Case 2501
            MsgBox "Export annulé (aucune donnée pour cette facture ?).", vbExclamation
        Case Else
            MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
    End Select

    ' S'assurer que l'état caché est bien refermé
    If CurrentProject.AllReports("E_Factures").IsLoaded Then
        DoCmd.Close acReport, "E_Factures"
    End If
End Sub
```

La gestion structurée des erreurs est traitée au chapitre 13.

---

## 7.8.8. Qualité du PDF et autres options

L'argument `OutputQuality` règle le compromis taille/qualité du PDF :

- `acExportQualityPrint` (par défaut) : qualité d'impression ;
- `acExportQualityScreen` : fichier plus léger, adapté à l'écran.

```vba
DoCmd.OutputTo acOutputReport, "E_Catalogue", acFormatPDF, strChemin, _
               , , , acExportQualityScreen
```

Pour l'export HTML, l'argument `TemplateFile` permet d'appliquer un modèle, et `Encoding` (par exemple `acUTF8`) fixe l'encodage.

---

## 7.8.9. RTF, HTML et Excel : fidélité et limites

- **RTF** produit un document éditable dans Word, mais la mise en page d'un état complexe (colonnes, sous-états, regroupements) est **mal restituée**. Convient aux états simples destinés à être retouchés.
- **HTML** offre une fidélité encore plus approximative ; à réserver à des sorties web élémentaires.
- **Excel** : exporter un **état** vers Excel donne de piètres résultats, car la structure visuelle ne correspond pas à un tableau. Pour alimenter Excel en **données**, il vaut mieux exporter une **requête ou une table** avec `DoCmd.TransferSpreadsheet` (chapitre 5.5).

En résumé : le PDF pour la fidélité, le RTF pour l'édition d'états simples, et `TransferSpreadsheet` (et non `OutputTo`) pour les données vers Excel.

---

## 7.8.10. Pièges et bonnes pratiques

- **Pas de `WhereCondition` dans `OutputTo`** : pour exporter un état filtré, l'ouvrir d'abord (caché) puis l'exporter (section 7.8.4).
- **Le PDF pour la fidélité** : RTF et HTML déforment les mises en page complexes.
- **`OutputTo` écrase sans demander** : gérer soi-même toute confirmation d'écrasement.
- **Purger les caractères interdits** du nom de fichier (`\ / : * ? " < > |`).
- **Intercepter l'erreur 2501** (état vide annulé dans `NoData`) et fermer l'état caché en cas d'incident (chapitres 7.9 et 13).
- **Fichier verrouillé** : un PDF déjà ouvert dans un lecteur empêche l'export ; prévoir le message d'erreur.
- **Données vers Excel** : préférer `TransferSpreadsheet` sur une requête (chapitre 5.5) plutôt que l'export d'un état.
- **Pour envoyer un état par e-mail**, `DoCmd.SendObject` produit et joint le fichier en une étape ; ce point est traité au chapitre 7.10 (et l'automation Outlook au chapitre 22.5).

---

## 7.8.11. Récapitulatif

- `DoCmd.OutputTo` exporte un état vers un fichier ; ses arguments clés sont le **type** (`acOutputReport`), le **format** (`acFormatPDF`, `acFormatRTF`, `acFormatHTML`…), le **chemin** et `AutoStart`.
- Le **PDF** est le format recommandé pour sa fidélité ; RTF/HTML déforment les mises en page complexes, et les données destinées à Excel passent plutôt par `TransferSpreadsheet` (chapitre 5.5).
- `OutputTo` n'accepte **pas** de filtre : pour un export **paramétré**, ouvrir l'état filtré (caché) avec `WhereCondition`, puis l'exporter, puis le refermer (chapitre 7.4).
- Construire un **nom de fichier valide** (caractères interdits purgés) et gérer le fait qu'`OutputTo` **écrase sans avertissement**.
- Encadrer l'export d'une **gestion d'erreurs** (erreur 2501 sur état vide, fichier verrouillé) et garantir la fermeture de l'état caché (chapitres 7.9 et 13).
- Pour la **diffusion par e-mail**, `DoCmd.SendObject` joint directement le fichier — détaillé au chapitre 7.10.

⏭️ [7.9. Gestion de l'événement NoData — éviter les états vides](/07-etats-reports/09-gestion-nodata.md)
