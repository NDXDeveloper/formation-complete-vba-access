🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 5.5. Import et export de données (TransferDatabase, TransferSpreadsheet, TransferText)

Une autre grande famille de `DoCmd` : l'**échange de données** avec l'extérieur. Trois méthodes couvrent trois sources : **`TransferDatabase`** (autres bases et ODBC), **`TransferSpreadsheet`** (Excel) et **`TransferText`** (fichiers texte / CSV). Chacune sait **importer**, **exporter** et **lier**. Cette section les présente, avec une attention particulière à un usage clé : la liaison des tables d'un back-end.

## Trois méthodes, trois sources

Le point commun des trois méthodes est leur premier argument, le **type de transfert** :

- **`acImport`** — importer dans la base courante ;
- **`acExport`** — exporter vers l'extérieur ;
- **`acLink`** — créer une **table liée** vers la source (les données restent à l'extérieur).

Chaque méthode décline ces trois modes pour sa source respective : autre base de données, classeur Excel, ou fichier texte.

## TransferSpreadsheet — échanger avec Excel

```
DoCmd.TransferSpreadsheet [TransferType], [SpreadsheetType], TableName, FileName, [HasFieldNames], [Range]
```

L'argument **`SpreadsheetType`** précise le format Excel : `acSpreadsheetTypeExcel12Xml` pour le `.xlsx` actuel, `acSpreadsheetTypeExcel9` pour l'ancien `.xls`, etc. **`HasFieldNames`** indique si la première ligne contient les en-têtes de colonnes.

```vba
' Importer une feuille Excel dans une table
DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, _
    "tblClients", "C:\Imports\clients.xlsx", True

' Exporter une table (ou une requête) vers Excel
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, _
    "tblClients", "C:\Exports\clients.xlsx", True

' Lier une feuille Excel sous forme de table liée
DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel12Xml, _
    "lnkClients", "C:\Data\clients.xlsx", True
```

C'est la voie rapide pour les échanges Excel ↔ Access. Pour un contrôle plus fin (mise en forme, plusieurs feuilles, formules…), on passera plutôt par l'**automation d'Excel** (section 22.3).

## TransferText — échanger avec des fichiers texte

```
DoCmd.TransferText [TransferType], [SpecificationName], TableName, FileName, [HasFieldNames], [HTMLTableName], [CodePage]
```

Le `TransferType` se décline ici selon le format texte : délimité (`acImportDelim`, `acExportDelim`), à largeur fixe (`acImportFixed`…) ou HTML (`acImportHTML`…). Deux arguments méritent une attention particulière.

**`SpecificationName`** désigne une **spécification** d'import/export sauvegardée — un modèle définissant la façon d'interpréter le fichier (séparateurs, types de champs, format des dates…). Pour un CSV non trivial, c'est indispensable : sans spécification, Access **devine** la structure, avec des résultats parfois inexacts. On crée ces spécifications via l'assistant d'import/export.

**`CodePage`** fixe l'**encodage**, ce qui est crucial pour les **accents** : `65001` correspond à l'UTF-8.

```vba
' Importer un CSV en s'appuyant sur une spécification sauvegardée
DoCmd.TransferText acImportDelim, "SpecImportClients", _
    "tblClients", "C:\Imports\clients.csv", True

' Exporter une table vers un CSV en UTF-8 (pour préserver les accents)
DoCmd.TransferText acExportDelim, , "tblClients", _
    "C:\Exports\clients.csv", True, , 65001   ' 65001 = UTF-8
```

## TransferDatabase — autres bases et sources ODBC

```
DoCmd.TransferDatabase [TransferType], DatabaseType, DatabaseName, [ObjectType], Source, Destination, [StructureOnly]
```

C'est la méthode la plus riche. **`DatabaseType`** indique le type de source — `"Microsoft Access"` pour une autre base `.accdb`/`.mdb`, `"ODBC Database"` pour une source ODBC (SQL Server, etc.) —, et **`DatabaseName`** son emplacement (chemin ou chaîne de connexion ODBC).

Particularité notable : son argument **`ObjectType`** permet de transférer **n'importe quel type d'objet** (pas seulement des tables) — `acTable`, `acQuery`, `acForm`, `acReport`, `acModule`, `acMacro`.

L'usage le plus important est la **liaison des tables d'un back-end**, au cœur de l'architecture **front-end / back-end** (sections 1.5, 15 et 21.4) :

```vba
' Lier une table d'une base back-end (front-end / back-end)
DoCmd.TransferDatabase acLink, "Microsoft Access", _
    "C:\Data\Backend.accdb", acTable, "tblClients", "tblClients"
```

Mais on peut aussi **importer un objet** depuis une autre base Access — par exemple récupérer un état d'une base modèle :

```vba
DoCmd.TransferDatabase acImport, "Microsoft Access", _
    "C:\Modeles\Modeles.accdb", acReport, "rptFactures", "rptFactures"
```

…ou **lier une table d'un serveur** via ODBC (chapitre 23) :

```vba
DoCmd.TransferDatabase acLink, "ODBC Database", _
    "ODBC;DRIVER=SQL Server;SERVER=monServeur;DATABASE=maBase;Trusted_Connection=Yes", _
    acTable, "dbo.Clients", "Clients"
```

L'argument **`StructureOnly`** (True/False) permet de transférer la **structure seule** (sans les données) ou structure + données.

## Transfer… ou OutputTo ?

Une distinction utile pour ne pas confondre les outils :

- les méthodes **`Transfer…`** échangent des **données** : `TableName` désigne une **table ou une requête** ;
- la méthode **`OutputTo`** (section 5.6) exporte un **objet rendu** vers un fichier dans un format donné (PDF, Excel, RTF, texte, HTML) — par exemple un **état mis en forme**.

Ainsi, pour exporter les résultats d'une requête vers Excel, `TransferSpreadsheet` exporte les **données** brutes, tandis qu'`OutputTo` exporte la **présentation** de l'objet. `OutputTo` est traité en section 5.6.

## Précautions

- Ces méthodes peuvent **échouer** (fichier verrouillé, format incompatible, chemin introuvable) : elles s'inscrivent dans une gestion d'erreurs (chapitre 13).
- Pour des **chemins relatifs** à la base, appuyez-vous sur `CurrentProject.Path` (section 4.3) plutôt que sur des chemins codés en dur.
- Pour le texte, **spécification** et **encodage** (`CodePage`) sont les clés d'un import/export fiable, en particulier avec des caractères accentués.

## À retenir

- Trois méthodes pour trois sources : **`TransferDatabase`** (autres bases, ODBC), **`TransferSpreadsheet`** (Excel), **`TransferText`** (texte/CSV) — chacune en mode **`acImport`**, **`acExport`** ou **`acLink`**.
- **`TransferText`** : pensez à la **spécification** (`SpecificationName`) pour les formats précis et au **`CodePage`** (65001 = UTF-8) pour les **accents**.
- **`TransferDatabase`** est la plus puissante : elle transfère **tout type d'objet**, et sert notamment à **lier les tables d'un back-end** (`acLink` + `"Microsoft Access"`, sections 15 et 21.4) ou des sources **ODBC** (chapitre 23).
- Ne pas confondre **`Transfer…`** (échange de **données**) et **`OutputTo`** (export d'un **objet rendu** : PDF, etc., section 5.6).
- Anticipez les **erreurs** (chapitre 13) et utilisez `CurrentProject.Path` pour les chemins relatifs (section 4.3).

---


⏭️ [5.6. Impression et aperçu (PrintOut, OutputTo)](/05-objet-docmd/06-impression-apercu.md)
