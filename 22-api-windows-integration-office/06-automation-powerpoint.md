🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.6. Automation avec PowerPoint — génération de présentations

PowerPoint clôt la série des sections d'automation Office. Le besoin est ici de **produire des présentations** à partir des données d'Access : rapports périodiques sous forme de diapositives, tableaux de bord à diffuser, ou présentations construites une diapositive par enregistrement (un produit, une région, une période par diapositive). Comme pour Word, on remplit le plus souvent un modèle, ou l'on bâtit les diapositives par code.

Cette section s'appuie sur les fondations de la section 22.3. Comme Outlook avait ses particularités, PowerPoint en présente une qu'il faut connaître d'emblée : il **se masque mal** pendant l'automation.

## Les fondations, et une spécificité PowerPoint

Les principes de la section 22.3 s'appliquent : liaison tardive privilégiée (`CreateObject("PowerPoint.Application")`), et libération soigneuse des objets. Le modèle objet de PowerPoint s'organise autour de l'objet `Application`, de la collection `Presentations`, de chaque `Presentation` contenant ses `Slides`, et de chaque `Slide` exposant ses `Shapes` — formes parmi lesquelles les **espaces réservés** (placeholders) accueillent titre et contenu. En liaison tardive, les constantes de PowerPoint (dispositions, codes de format) doivent être définies manuellement.

### PowerPoint se masque mal : minimiser plutôt que cacher

À la différence d'Excel et de Word, qui se laissent volontiers masquer par `Visible = False` pendant le traitement, **PowerPoint refuse fréquemment d'être caché** : tenter de mettre `Application.Visible` à faux déclenche souvent une erreur ou reste sans effet, l'application n'autorisant pas le masquage de sa fenêtre. La pratique consiste donc à **laisser PowerPoint visible**, quitte à **minimiser sa fenêtre** une fois une présentation ouverte (via `Application.WindowState`), pour limiter la gêne visuelle. Il faut s'attendre, avec PowerPoint, à voir la présentation se construire à l'écran — comportement normal, et non un défaut.

## Partir d'un modèle ou bâtir de zéro

Comme pour Word, deux voies s'offrent. On peut **partir d'un modèle** (`Presentations.Open` sur un fichier `.pptx` ou `.potx` préparé) afin de bénéficier d'une charte graphique cohérente, puis enregistrer sous le nom de sortie. On peut aussi **créer une présentation vierge** (`Presentations.Add`) et tout construire par code. Le modèle est préférable dès lors que l'apparence doit être maîtrisée et homogène.

## Ajouter une diapositive et remplir ses espaces réservés

Une diapositive s'ajoute par `Slides.Add`, en précisant sa position et sa **disposition** (une constante `PpSlideLayout`). On accède ensuite à ses espaces réservés par la collection `Placeholders` — le premier étant le titre, le second le contenu — dont on renseigne le texte via `TextFrame.TextRange.Text`.

L'exemple suivant génère une présentation complète : une diapositive de titre, puis une diapositive par enregistrement d'une requête :

```vba
Public Sub GenererPresentation(ByVal sql As String, _
                               ByVal titre As String, _
                               ByVal cheminSortie As String)
    Const ppLayoutTitle As Long = 1
    Const ppLayoutText As Long = 2
    Const ppSaveAsOpenXMLPresentation As Long = 24   ' .pptx (liaison tardive)

    Dim pptApp As Object, pres As Object, diapo As Object
    Dim db As DAO.Database, rs As DAO.Recordset
    Dim n As Long

    On Error GoTo Nettoyage

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True                 ' PowerPoint se masque mal : on le laisse visible

    Set pres = pptApp.Presentations.Add

    ' Diapositive de titre
    Set diapo = pres.Slides.Add(1, ppLayoutTitle)
    diapo.Shapes.Placeholders(1).TextFrame.TextRange.Text = titre
    diapo.Shapes.Placeholders(2).TextFrame.TextRange.Text = Format$(Date, "dd/mm/yyyy")

    ' Une diapositive par enregistrement
    n = 1
    Do Until rs.EOF
        n = n + 1
        Set diapo = pres.Slides.Add(n, ppLayoutText)
        diapo.Shapes.Placeholders(1).TextFrame.TextRange.Text = Nz(rs!Titre, "")
        diapo.Shapes.Placeholders(2).TextFrame.TextRange.Text = Nz(rs!Contenu, "")
        rs.MoveNext
    Loop

    ' Enregistrement au format .pptx
    pres.SaveAs cheminSortie, ppSaveAsOpenXMLPresentation

Nettoyage:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not pres Is Nothing Then pres.Close
    ' Ne quitter que si aucune autre présentation n'est ouverte
    If Not pptApp Is Nothing Then
        If pptApp.Presentations.Count = 0 Then pptApp.Quit
    End If
    Set diapo = Nothing
    Set pres = Nothing
    Set pptApp = Nothing
    Set rs = Nothing
    Set db = Nothing
End Sub
```

La méthode `Slides.Add` employée ici est la voie classique, simple à mobiliser en liaison tardive ; les versions récentes proposent aussi `Slides.AddSlide`, qui s'appuie sur les dispositions personnalisées du masque de diapositives.

## Insérer un tableau de données

Pour présenter un jeu de données sur une diapositive, on ajoute un **tableau** par `Shapes.AddTable`, en précisant le nombre de lignes et de colonnes ainsi que sa position et sa taille (en points). On remplit ensuite chaque cellule par `Table.Cell(ligne, colonne).Shape.TextFrame.TextRange.Text` :

```vba
' Ajouter un tableau et le remplir depuis un recordset
Dim formeTableau As Object, t As Object
Dim lignes As Long, c As Long, ligne As Long

If rs.EOF Then Exit Sub
rs.MoveLast: rs.MoveFirst          ' nécessaire pour que RecordCount soit exact
lignes = rs.RecordCount + 1        ' + la ligne d'en-tête

Set formeTableau = diapo.Shapes.AddTable(lignes, rs.Fields.Count, 30, 80, 660, 300)
Set t = formeTableau.Table

' En-têtes : noms des champs
For c = 1 To rs.Fields.Count
    t.Cell(1, c).Shape.TextFrame.TextRange.Text = rs.Fields(c - 1).Name
Next c

' Données
ligne = 1
Do Until rs.EOF
    ligne = ligne + 1
    For c = 1 To rs.Fields.Count
        t.Cell(ligne, c).Shape.TextFrame.TextRange.Text = Nz(rs.Fields(c - 1).Value, "")
    Next c
    rs.MoveNext
Loop
```

Deux remarques. D'une part, contrairement à Word, PowerPoint n'offre pas de conversion de texte en tableau : le remplissage est **nécessairement cellule par cellule**, ce qui, conjugué au coût du franchissement entre processus rappelé à la section 22.3 et à l'impossibilité de masquer PowerPoint, rend l'opération visible et relativement lente pour de gros tableaux. D'autre part, l'appel à `MoveLast` puis `MoveFirst` est indispensable pour obtenir un `RecordCount` exact : sur un jeu d'enregistrements fraîchement ouvert, ce compteur n'est fiable qu'après avoir parcouru tous les enregistrements (chapitre 9).

## Insérer une image ou un graphique

On peut enrichir une diapositive d'une image — un logo, ou un graphique pré-généré et enregistré en fichier — au moyen de `Shapes.AddPicture` :

```vba
Const msoFalse As Long = 0
Const msoTrue As Long = -1
diapo.Shapes.AddPicture FileName:="C:\logo.png", _
    LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
    Left:=30, Top:=30, Width:=120, Height:=60
```

Le paramètre `LinkToFile` à faux et `SaveWithDocument` à vrai intègrent l'image dans la présentation plutôt que de la lier à un fichier externe, ce qui garantit que la présentation reste autonome une fois diffusée.

## Enregistrer, exporter en PDF, ou présenter

Comme pour Word, la présentation générée connaît plusieurs destins.

On peut l'**enregistrer** au format `.pptx` (code 24) comme dans l'exemple ci-dessus, ou l'**exporter en PDF** via le code 32, utile pour une diffusion figée — capacité qui rejoint les techniques d'export d'états de la section 7.8 :

```vba
Const ppSaveAsPDF As Long = 32
pres.SaveAs "C:\sortie\presentation.pdf", ppSaveAsPDF
```

On peut enfin **présenter le résultat à l'utilisateur** pour qu'il le complète ou le projette : on laisse alors PowerPoint visible et la présentation ouverte, en libérant ses propres variables sans refermer la présentation — scénario d'autant plus naturel que PowerPoint est de toute façon visible.

### Quitter prudemment

PowerPoint hébergeant plusieurs présentations au sein d'une même instance, et pouvant être déjà ouvert par l'utilisateur, la prudence s'impose au moment de quitter. La routine de nettoyage ci-dessus illustre la bonne approche : refermer la présentation que l'on a créée, puis **ne quitter l'application que si plus aucune présentation n'est ouverte** (`Presentations.Count = 0`). On évite ainsi de fermer par mégarde les présentations sur lesquelles l'utilisateur travaillait.

## Articulation avec le reste du chapitre et de la formation

Cette section conclut l'automation Office et s'appuie sur plusieurs parties de la formation :

- Les **fondations de l'automation** sont établies à la section 22.3, et le choix de liaison à la section 2.6.
- Elle fait partie d'un ensemble avec l'automation de **Word** (section 22.4) et d'**Outlook** (section 22.5), avec lequel elle partage structure et précautions.
- L'**export en PDF** rejoint la section 7.8.
- Les **jeux d'enregistrements** alimentant les diapositives relèvent du chapitre 9 (avec la précision sur `RecordCount`), la **gestion d'erreurs** du chapitre 13, et les considérations de **performance** du chapitre 18.
- La compatibilité entre versions, comme pour toute automation, relève de la section 21.6.

Les sections suivantes du chapitre quittent l'automation Office pour ouvrir l'application sur le web (services REST, JSON) puis sur le système (registre, système de fichiers).

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'automation de PowerPoint :

- **Tenter de masquer PowerPoint** par `Visible = False`, ce qui échoue ou lève une erreur ; le laisser visible, ou minimiser sa fenêtre.
- **Quitter l'application sans précaution**, et fermer les présentations que l'utilisateur avait ouvertes ; ne quitter que si `Presentations.Count = 0`.
- **Lire `RecordCount` sans avoir appelé `MoveLast`**, et dimensionner un tableau sur un compteur erroné.
- **Remplir un gros tableau cellule par cellule** sans mesurer le coût, accentué par l'impossibilité de masquer l'application.
- **Lier les images au lieu de les intégrer**, rendant la présentation dépendante de fichiers externes une fois diffusée.
- **Ne pas libérer les objets**, laissant subsister des instances de PowerPoint en mémoire.

⏭️ [22.7. Appel de Web Services REST depuis VBA Access (MSXML2, WinHttp)](/22-api-windows-integration-office/07-web-services-rest.md)
