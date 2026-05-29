🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.10. Manipulation du système de fichiers (FileSystemObject)

Cette dernière section du chapitre traite de la manipulation du système de fichiers depuis VBA au moyen du **FileSystemObject** (FSO), l'interface objet de référence pour travailler avec les fichiers, les dossiers, les lecteurs et les fichiers texte. Le FSO a déjà été sollicité à plusieurs reprises — pour la copie du front-end lors de l'auto-update et de l'installation (sections 21.3, 21.5 et 21.8) — et constitue l'alternative lisible et structurée tant aux fonctions de fichiers natives de VBA qu'à l'appel direct de l'API.

## Qu'est-ce que le FileSystemObject ?

Le FileSystemObject fait partie du **Microsoft Scripting Runtime**, présent en standard sur toutes les versions modernes de Windows. On l'instancie en liaison tardive par `CreateObject("Scripting.FileSystemObject")`, ou en liaison précoce en référençant la bibliothèque « Microsoft Scripting Runtime » (section 2.6). Contrairement au ScriptControl évoqué à la section 22.8, le FSO **fonctionne indifféremment en 32 et en 64 bits** : aucune contrainte de bitness ne le concerne.

Il offre un **modèle objet unifié** couvrant l'ensemble des besoins de fichiers : un objet racine `FileSystemObject`, et des objets `File`, `Folder`, `Drive` et `TextStream` pour manipuler respectivement les fichiers, les dossiers, les lecteurs et le contenu texte.

## FSO ou fonctions natives de VBA ?

VBA dispose de fonctions de fichiers intégrées — `Dir`, `FileCopy`, `Kill`, `MkDir`, `Name`, `FileDateTime`… — parfaitement valables pour des opérations simples. Le FSO se distingue par sa **lisibilité**, son approche **objet** et sa **richesse** : analyse des chemins, parcours récursif d'arborescences, informations sur les lecteurs, lecture et écriture de texte avec gestion de l'encodage.

Le choix entre les deux relève surtout du style et du besoin : on réserve volontiers les fonctions natives aux opérations ponctuelles et simples, et l'on préfère le FSO pour tout travail structuré sur le système de fichiers. En liaison tardive, il faut définir soi-même les quelques constantes du FSO (modes d'ouverture, encodage) ; la liaison précoce les fournit, avec l'IntelliSense.

## Opérations sur les fichiers et les dossiers

Les méthodes de l'objet racine couvrent l'existence, la copie, le déplacement et la suppression :

```vba
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists("C:\donnees\rapport.pdf") Then
    ' Le troisième argument à True autorise l'écrasement
    fso.CopyFile "C:\donnees\rapport.pdf", "D:\archives\rapport.pdf", True
End If

fso.DeleteFile "C:\temp\ancien.txt", True   ' True = forcer, même en lecture seule
```

Quelques particularités méritent d'être notées : `CopyFile` écrase par défaut (paramètre facultatif), tandis que `MoveFile` **échoue si la destination existe** ; `DeleteFile` n'efface un fichier en lecture seule que si l'on force la suppression ; et ces méthodes acceptent des **caractères génériques** pour traiter plusieurs fichiers à la fois.

Pour les dossiers, une limite importante du FSO est que **`CreateFolder` ne crée qu'un seul niveau** et échoue si le dossier existe déjà ou si son parent est absent. Créer toute une arborescence suppose donc de procéder niveau par niveau, ce qu'une petite fonction récursive encapsule commodément :

```vba
' Crée toute l'arborescence d'un chemin de dossier, niveau par niveau
Public Sub CreerArborescence(ByVal dossier As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(dossier) Then Exit Sub

    Dim parent As String
    parent = fso.GetParentFolderName(dossier)
    If Len(parent) > 0 Then CreerArborescence parent   ' créer d'abord le parent
    fso.CreateFolder dossier
End Sub
```

## Manipuler les chemins

L'une des forces les plus utiles du FSO réside dans ses fonctions d'**analyse de chemins**, qui évitent les manipulations de chaînes fragiles et propres aux erreurs :

```vba
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Dim chemin As String
chemin = "C:\donnees\factures\facture_2026.pdf"

Debug.Print fso.GetParentFolderName(chemin)   ' C:\donnees\factures
Debug.Print fso.GetFileName(chemin)           ' facture_2026.pdf
Debug.Print fso.GetBaseName(chemin)           ' facture_2026
Debug.Print fso.GetExtensionName(chemin)      ' pdf

' Construire un chemin en gérant correctement le séparateur
Debug.Print fso.BuildPath("C:\donnees", "export.csv")   ' C:\donnees\export.csv
```

`BuildPath` est particulièrement appréciable : il assemble un dossier et un nom sans se soucier de la présence ou de l'absence de la barre oblique finale, source classique de chemins malformés lorsqu'on les concatène à la main.

## Lire et écrire des fichiers texte

Le FSO lit et écrit le texte au moyen de l'objet `TextStream`, obtenu par `CreateTextFile` (création) ou `OpenTextFile` (ouverture). L'écriture est directe :

```vba
Const ForReading As Long = 1
Const ForWriting As Long = 2
Const ForAppending As Long = 8

Dim fso As Object, ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")

' Écriture (True = écraser le fichier existant)
Set ts = fso.CreateTextFile("C:\export\journal.txt", True)
ts.WriteLine "Début du traitement : " & Now
ts.WriteLine "Ligne de données"
ts.Close
```

La lecture s'effectue ligne par ligne — en s'arrêtant à `AtEndOfStream` — ou intégralement par `ReadAll` :

```vba
' Lecture ligne par ligne
Set ts = fso.OpenTextFile("C:\export\journal.txt", ForReading)
Do Until ts.AtEndOfStream
    Debug.Print ts.ReadLine
Loop
ts.Close
```

Une **limite essentielle** doit être connue, d'autant plus pertinente après les sections consacrées au web : le `TextStream` du FSO gère le texte **ANSI et UTF-16 (Unicode)**, mais **pas l'UTF-8**. Or l'UTF-8 est l'encodage dominant des données web et des fichiers JSON (sections 22.7 et 22.8). Pour lire ou écrire correctement un fichier UTF-8, il faut donc recourir à l'objet `ADODB.Stream`, qui permet de spécifier le jeu de caractères, plutôt qu'au FSO.

## Parcourir une arborescence

Le parcours récursif d'une arborescence de dossiers est un usage où le FSO brille, là où la fonction native `Dir`, par nature séquentielle et fragile à réutiliser de manière imbriquée, s'avère malcommode. Les collections `Files` et `SubFolders` d'un objet `Folder` se parcourent simplement, et la récursivité découle naturellement de la structure :

```vba
' Parcourt récursivement un dossier et traite chaque fichier
Public Sub ParcourirDossier(ByVal cheminDossier As String)
    Dim fso As Object, dossier As Object, fichier As Object, sousDossier As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(cheminDossier) Then Exit Sub

    Set dossier = fso.GetFolder(cheminDossier)

    ' Fichiers du dossier courant
    For Each fichier In dossier.Files
        Debug.Print fichier.Path & " (" & fichier.Size & " octets)"
    Next fichier

    ' Descente récursive dans les sous-dossiers
    For Each sousDossier In dossier.SubFolders
        ParcourirDossier sousDossier.Path
    Next sousDossier
End Sub
```

## Propriétés des fichiers et information sur les lecteurs

L'objet `File`, obtenu par `GetFile`, expose les propriétés d'un fichier — notamment sa **date de dernière modification**, sur laquelle reposent les comparaisons de version de l'auto-update (sections 21.3 et 21.5) :

```vba
Dim fichier As Object
Set fichier = fso.GetFile("C:\app\MonApp.accde")
Debug.Print fichier.DateLastModified
```

L'objet `Drive`, obtenu par `GetDrive`, renseigne quant à lui sur les lecteurs — utile, par exemple, pour vérifier l'**espace disponible** avant un export volumineux :

```vba
Dim lecteur As Object
Set lecteur = fso.GetDrive("C")
Debug.Print lecteur.FreeSpace & " octets libres"
```

## Articulation avec le reste du chapitre et de la formation

Cette section répond aux usages du système de fichiers disséminés dans le chapitre 21 et s'appuie sur d'autres parties de la formation :

- La **copie de fichiers** sous-tend l'auto-update et l'installation (sections 21.3, 21.5 et 21.8).
- La **date de dernière modification** alimente la comparaison de version de l'auto-update (sections 21.3 et 21.5).
- La gestion de l'**UTF-8 via `ADODB.Stream`** concerne les fichiers JSON et les données web des sections 22.7 et 22.8.
- L'**absence de contrainte de bitness** contraste avec le ScriptControl de la section 22.8, dans le prolongement de la section 21.7.
- Le choix de **liaison** relève de la section 2.6, et la **gestion d'erreurs** des opérations de fichiers du chapitre 13.

Cette section clôt le chapitre 22 et son volet d'interaction avec le système, après l'appel d'API, l'automation des applications Office et l'ouverture sur le web.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'usage du FileSystemObject :

- **Attendre de `CreateFolder` qu'il crée toute une arborescence**, alors qu'il ne crée qu'un niveau ; procéder niveau par niveau.
- **Lire ou écrire un fichier UTF-8 avec le `TextStream`**, qui ne gère que l'ANSI et l'UTF-16 ; recourir à `ADODB.Stream` pour l'UTF-8.
- **Supposer que `MoveFile` écrase la destination**, alors qu'il échoue si elle existe déjà.
- **Concaténer des chemins à la main** plutôt que d'employer `BuildPath`, au risque d'erreurs de séparateur.
- **Tenter de supprimer un fichier en lecture seule sans forcer**, ou un dossier non vide sans le paramètre adéquat.
- **Réutiliser `Dir` de manière imbriquée** pour un parcours d'arborescence, là où la récursivité du FSO est plus sûre.

⏭️ [23. Migration et interopérabilité](/23-migration-interoperabilite/README.md)
