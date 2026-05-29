🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.3. Automation avec Excel — export et import de données

Excel est, de loin, la cible d'automation la plus fréquente depuis Access. Le besoin est récurrent : exporter des données vers un classeur pour les analyser, les mettre en forme et les diffuser ; ou, dans l'autre sens, importer dans Access des données saisies dans une feuille de calcul. Cette section, première des sections consacrées à l'automation Office, **établit d'abord les fondations communes** — liaison, création d'instance, libération des objets — qui seront réutilisées pour Word, Outlook et PowerPoint (sections 22.4 à 22.6), avant de les appliquer à l'export et à l'import de données avec Excel.

## Avant l'automation : les solutions natives

Avant de recourir à l'automation, il faut connaître la voie la plus simple. Pour un export ou un import **direct** d'une table ou d'une requête vers une feuille de calcul, Access propose des méthodes natives qui ne nécessitent aucune automation : `DoCmd.TransferSpreadsheet` et `DoCmd.OutputTo`, présentées à la section 5.5. Une seule ligne suffit alors à exporter une requête entière vers un fichier `.xlsx`.

L'automation ne se justifie que lorsqu'on a besoin de ce que ces méthodes ne savent pas faire : mise en forme élaborée, écriture sur plusieurs feuilles, insertion de formules, placement des données à des emplacements précis, transformation au fil de l'import, ou pilotage fin du classeur. Dans ce cas seulement, on prend la main sur Excel par son modèle objet.

## Les fondations de l'automation

Piloter Excel par automation, c'est manipuler son **modèle objet** depuis Access, exactement comme on manipule les objets d'Access lui-même. Trois notions en constituent le socle, valables pour toutes les applications Office.

### Liaison précoce ou liaison tardive

Deux manières s'offrent pour référencer le modèle objet d'Excel.

La **liaison précoce** (early binding) suppose d'activer une référence à la bibliothèque « Microsoft Excel XX.X Object Library » (section 2.5), puis de déclarer des objets typés : `Dim xlApp As Excel.Application`. Elle apporte l'IntelliSense et la vérification à la compilation, mais lie l'application à une **version précise** de la bibliothèque, ce qui pose les problèmes de compatibilité entre versions décrits à la section 21.6.

La **liaison tardive** (late binding) déclare des objets génériques (`Dim xlApp As Object`) et crée l'instance par `CreateObject`. Elle ne requiert aucune référence, s'avère **indépendante de la version**, mais prive de l'IntelliSense et oblige à définir soi-même les constantes d'Excel (telles `xlUp` ou le code de format `.xlsx`), absentes en l'absence de référence. Le choix entre les deux est traité à la section 2.6 ; **pour un déploiement sur un parc hétérogène, la liaison tardive est généralement préférable**, et c'est elle qui est employée dans les exemples ci-dessous.

### Créer ou récupérer une instance

`CreateObject("Excel.Application")` démarre une **nouvelle** instance d'Excel. À l'inverse, `GetObject` permet de récupérer une instance déjà en cours, ou d'ouvrir directement un classeur. Deux propriétés de l'application méritent d'être réglées dès le départ :

- **`Visible`** : maintenue à `False`, elle masque Excel pendant le traitement, ce qui améliore nettement la performance et évite que l'utilisateur ne voie un classeur s'agiter à l'écran.
- **`DisplayAlerts`** : mise à `False`, elle supprime les invites d'Excel (confirmation d'écrasement, par exemple) qui, autrement, bloqueraient un traitement automatisé.

### Libérer les objets : éviter les processus fantômes

C'est **le** point critique de l'automation, et le piège le plus courant. Une instance d'Excel créée par code, mais non refermée et non libérée, subsiste en mémoire sous la forme d'un processus `EXCEL.EXE` fantôme, invisible mais consommateur de ressources, qui s'accumule à chaque exécution.

La discipline à respecter est stricte : refermer le classeur, appeler `Quit` sur l'application, puis libérer **chaque** variable objet en la réinitialisant à `Nothing`, dans l'ordre **inverse** de leur création. Le tout doit être placé dans une routine de nettoyage atteinte même en cas d'erreur (chapitre 13). Le squelette suivant sert de modèle à toute automation Office :

```vba
Public Sub StructureAutomation()
    Dim xlApp As Object
    Dim xlClasseur As Object
    Dim xlFeuille As Object

    On Error GoTo Nettoyage

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlClasseur = xlApp.Workbooks.Add
    Set xlFeuille = xlClasseur.Worksheets(1)

    ' ... traitement ...

Nettoyage:
    ' Libération en ordre inverse de création, à l'épreuve des erreurs
    On Error Resume Next
    If Not xlClasseur Is Nothing Then xlClasseur.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlFeuille = Nothing
    Set xlClasseur = Nothing
    Set xlApp = Nothing
End Sub
```

Un corollaire important : il faut **éviter les références « à deux points »** qui créent des objets intermédiaires non affectés à une variable, comme `xlApp.Workbooks.Add.Worksheets(1)`. Ces objets temporaires échappent à la libération et favorisent les processus fantômes. Mieux vaut affecter chaque objet à sa propre variable, comme ci-dessus, afin de pouvoir le libérer explicitement.

## Exporter des données d'Access vers Excel

### La méthode rapide : CopyFromRecordset

La technique recommandée pour exporter un volume de données consiste à ouvrir un jeu d'enregistrements (chapitre 9) puis à le **déverser en une seule opération** dans une plage, au moyen de la méthode `CopyFromRecordset`. Celle-ci copie l'intégralité du recordset à partir de la cellule indiquée, sans inclure les noms de champs — que l'on écrit donc séparément :

```vba
Public Sub ExporterVersExcel(ByVal sql As String, ByVal cheminFichier As String)
    Const xlOpenXMLWorkbook As Long = 51   ' code du format .xlsx (à définir en liaison tardive)

    Dim xlApp As Object, xlClasseur As Object, xlFeuille As Object
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim i As Long

    On Error GoTo Nettoyage

    ' Jeu d'enregistrements à exporter (snapshot : lecture seule, rapide)
    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    ' Démarrer Excel, masqué
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlClasseur = xlApp.Workbooks.Add
    Set xlFeuille = xlClasseur.Worksheets(1)

    ' 1) En-têtes : noms des champs sur la première ligne
    For i = 0 To rs.Fields.Count - 1
        xlFeuille.Cells(1, i + 1).Value = rs.Fields(i).Name
    Next i

    ' 2) Déverser tout le jeu d'enregistrements en une seule opération
    If Not rs.EOF Then
        rs.MoveFirst
        xlFeuille.Range("A2").CopyFromRecordset rs
    End If

    ' 3) Mise en forme légère
    xlFeuille.Rows(1).Font.Bold = True
    xlFeuille.Columns.AutoFit

    ' 4) Enregistrement au format .xlsx
    xlClasseur.SaveAs cheminFichier, xlOpenXMLWorkbook

Nettoyage:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not xlClasseur Is Nothing Then xlClasseur.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlFeuille = Nothing
    Set xlClasseur = Nothing
    Set xlApp = Nothing
    Set rs = Nothing
    Set db = Nothing
End Sub
```

### L'écriture par tableau

Lorsque les données à exporter sont déjà calculées en mémoire plutôt qu'issues d'un recordset, on peut écrire un **tableau à deux dimensions dans une plage en une seule affectation**, la taille de la plage devant correspondre aux dimensions du tableau :

```vba
Dim donnees(1 To 3, 1 To 2) As Variant
donnees(1, 1) = "Dupont": donnees(1, 2) = 1200
donnees(2, 1) = "Martin": donnees(2, 2) = 980
donnees(3, 1) = "Durand": donnees(3, 2) = 1450
xlFeuille.Range("A1:B3").Value = donnees
```

### Éviter l'écriture cellule par cellule

Ces deux méthodes ont une raison d'être commune : **chaque appel au modèle objet d'Excel franchit la frontière entre les deux processus (Access et Excel), opération coûteuse.** Écrire les données cellule par cellule, dans une boucle qui touche chaque case individuellement, multiplie ces franchissements et se révèle plus lent de plusieurs ordres de grandeur. `CopyFromRecordset` et l'écriture par tableau, qui réalisent l'opération en un seul échange, sont donc à privilégier systématiquement.

### Mise en forme et enregistrement

Une fois les données en place, la mise en forme s'applique également de préférence à des plages entières plutôt que cellule par cellule : largeur des colonnes (`AutoFit`), en-têtes en gras, formats de nombre. L'enregistrement se fait par `SaveAs` en précisant le **code du format** ; les plus utiles sont 51 pour `.xlsx` (`xlOpenXMLWorkbook`) et 56 pour l'ancien `.xls` (`xlExcel8`). En liaison tardive, ces codes doivent être déclarés comme constantes, leurs noms symboliques n'étant pas disponibles.

## Importer des données d'Excel vers Access

### Lire efficacement : la plage en un tableau

Le même principe de minimisation des échanges vaut pour la lecture. Plutôt que de lire chaque cellule, on **charge toute une plage dans un tableau** en une seule opération, puis on parcourt ce tableau en mémoire. Encore faut-il délimiter la zone utile : la dernière ligne renseignée d'une colonne se trouve commodément avec `End(xlUp)` à partir du bas de la feuille :

```vba
Public Sub ImporterDepuisExcel(ByVal cheminFichier As String)
    Const xlUp As Long = -4162   ' constante à définir en liaison tardive

    Dim xlApp As Object, xlClasseur As Object, xlFeuille As Object
    Dim donnees As Variant, derniereLigne As Long, i As Long
    Dim db As DAO.Database, rs As DAO.Recordset

    On Error GoTo Nettoyage

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlClasseur = xlApp.Workbooks.Open(cheminFichier)
    Set xlFeuille = xlClasseur.Worksheets(1)

    ' Dernière ligne utilisée de la colonne A
    derniereLigne = xlFeuille.Cells(xlFeuille.Rows.Count, 1).End(xlUp).Row

    ' Lire la plage utile (de la ligne 2 à la fin) en une seule opération
    If derniereLigne >= 2 Then
        donnees = xlFeuille.Range("A2:C" & derniereLigne).Value

        ' Insérer dans la table via un Recordset DAO ouvert en ajout
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tblImport", dbOpenDynaset, dbAppendOnly)
        For i = LBound(donnees, 1) To UBound(donnees, 1)
            rs.AddNew
            rs!Nom = Nz(donnees(i, 1), "")
            rs!Quantite = Nz(donnees(i, 2), 0)
            rs!Prix = Nz(donnees(i, 3), 0)
            rs.Update
        Next i
    End If

Nettoyage:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not xlClasseur Is Nothing Then xlClasseur.Close SaveChanges:=False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlFeuille = Nothing
    Set xlClasseur = Nothing
    Set xlApp = Nothing
    Set rs = Nothing
    Set db = Nothing
End Sub
```

Un tableau lu depuis une plage est **indexé à partir de 1** sur deux dimensions — `donnees(ligne, colonne)`. Un cas particulier mérite attention : lorsqu'une plage ne contient qu'une **seule cellule**, sa propriété `Value` renvoie une valeur scalaire et non un tableau ; en lisant une plage de plusieurs colonnes comme ici, on obtient en revanche toujours un tableau.

### Insérer dans Access

L'insertion proprement dite s'appuie sur les techniques d'écriture d'Access : un **Recordset DAO** ouvert en ajout (`AddNew` / `Update`), comme ci-dessus, ou des requêtes `INSERT` (chapitre 11). C'est à cette étape que l'on tire parti de l'automation par rapport à un import natif : validation des valeurs, transformation, gestion des cas particuliers, lecture partielle de la feuille.

## Performance : minimiser les échanges entre processus

La leçon transversale de cette section est que **le coût de l'automation tient au franchissement de la frontière entre processus**. Pour le réduire, plusieurs réglages se combinent : maintenir `Visible` à `False` pendant le traitement, désactiver le rafraîchissement de l'écran (`Application.ScreenUpdating = False`) et le recalcul automatique (`Application.Calculation` en mode manuel) lors des écritures massives, puis les rétablir ; et surtout privilégier les **opérations en bloc** — `CopyFromRecordset`, lecture et écriture par tableau — plutôt que les accès cellule par cellule. Ces principes rejoignent les considérations de performance du chapitre 18.

## Articulation avec le reste du chapitre et de la formation

Cette section sert de socle aux autres sections d'automation et s'appuie sur plusieurs parties de la formation :

- Les **fondations de l'automation** (liaison, instance, libération) sont reprises pour Word, Outlook et PowerPoint aux sections 22.4 à 22.6.
- Les **solutions natives** d'export et d'import relèvent de la section 5.5.
- Le choix **liaison précoce / tardive** est traité à la section 2.6, et la compatibilité entre versions à la section 21.6.
- Les **jeux d'enregistrements** mobilisés à l'export comme à l'import relèvent du chapitre 9, et les requêtes `INSERT` du chapitre 11.
- La **gestion d'erreurs** garantissant le nettoyage relève du chapitre 13, et les considérations de **performance** du chapitre 18.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'automation d'Excel :

- **Ne pas libérer les objets**, laissant s'accumuler des processus `EXCEL.EXE` fantômes en mémoire.
- **Utiliser des références « à deux points »** qui créent des objets intermédiaires impossibles à libérer.
- **Écrire ou lire cellule par cellule** au lieu d'opérations en bloc, avec un effondrement des performances.
- **Oublier de définir les constantes Excel** (codes de format, `xlUp`…) en liaison tardive, où leurs noms symboliques sont indisponibles.
- **Omettre la routine de nettoyage en cas d'erreur**, de sorte qu'une interruption laisse Excel ouvert et des objets non libérés.
- **Recourir à l'automation pour un besoin simple** que `DoCmd.TransferSpreadsheet` traiterait en une ligne (section 5.5).

⏭️ [22.4. Automation avec Word — génération de documents à partir d'Access](/22-api-windows-integration-office/04-automation-word.md)
