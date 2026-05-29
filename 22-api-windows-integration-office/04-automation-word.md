🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.4. Automation avec Word — génération de documents à partir d'Access

Après Excel, Word est la deuxième cible d'automation par ordre de fréquence. Le besoin est différent : il ne s'agit plus d'échanger des tableaux de données, mais de **produire des documents** à partir des informations stockées dans Access — courriers, contrats, attestations, certificats, factures, devis. Le schéma récurrent consiste à **remplir un modèle** avec les données d'un ou plusieurs enregistrements, puis à enregistrer, imprimer ou présenter le document obtenu.

Cette section s'appuie sur les fondations établies à la section 22.3 et se concentre sur les techniques propres à Word.

## S'appuyer sur les fondations de l'automation

Les principes posés à la section 22.3 s'appliquent à l'identique : on privilégie la **liaison tardive** (`Dim wdApp As Object`, `CreateObject("Word.Application")`) pour l'indépendance vis-à-vis des versions (sections 2.6 et 21.6), on règle `Visible` et `DisplayAlerts`, et l'on veille à **libérer tous les objets** dans une routine de nettoyage à l'épreuve des erreurs, sous peine de laisser des processus `WINWORD.EXE` fantômes en mémoire. On évite de même les références « à deux points » non affectées à une variable.

Le modèle objet de Word s'organise autour de quelques objets clés : `Application` contient la collection `Documents`, chaque `Document` exposant ses `Range`, `Bookmarks`, `Tables`, `Paragraphs` et son `Content`. En liaison tardive, les constantes de Word (codes de format, options de remplacement) doivent être définies manuellement.

## Partir d'un modèle plutôt que d'un document

La bonne pratique consiste à fonder la génération sur un **modèle** (un fichier `.dotx` ou `.docx` préparé) plutôt que de modifier un document existant. La méthode `Documents.Add` avec le paramètre `Template` crée un **nouveau** document à partir du modèle, laissant ce dernier intact :

```vba
Set wdDoc = wdApp.Documents.Add(Template:=cheminModele)
```

On remplit ensuite ce nouveau document, puis on l'enregistre sous le nom de sortie voulu. Le modèle original n'est jamais altéré et reste disponible pour les générations suivantes.

## Méthode 1 : les signets — la voie robuste

La technique la plus fiable repose sur les **signets** (bookmarks). On place, dans le modèle, des signets nommés aux endroits où les données doivent s'insérer (`Nom`, `Adresse`, `DateDuJour`…). Depuis Access, on insère ensuite la valeur de chaque champ à l'emplacement du signet correspondant. L'emplacement est ainsi maîtrisé par le concepteur du modèle, indépendamment du code.

Une subtilité doit être connue : **affecter du texte à la plage d'un signet supprime le signet** une fois l'insertion faite. Pour une génération en une passe, c'est sans conséquence ; mais pour préserver le signet — par exemple afin de le réutiliser —, il faut le recréer sur le texte inséré. La fonction d'aide suivante encapsule ce comportement robuste :

```vba
' Insère une valeur dans un signet et préserve celui-ci sur le texte inséré
Private Sub RemplirSignet(ByVal wdDoc As Object, ByVal nom As String, _
                          ByVal valeur As String)
    Dim plage As Object
    If Not wdDoc.Bookmarks.Exists(nom) Then Exit Sub
    Set plage = wdDoc.Bookmarks(nom).Range
    plage.Text = valeur
    wdDoc.Bookmarks.Add nom, plage      ' recrée le signet sur le texte inséré
End Sub
```

L'ensemble s'assemble en une procédure de génération qui lit les données d'un client, remplit le modèle et enregistre le document :

```vba
Public Sub GenererCourrier(ByVal idClient As Long, _
                           ByVal cheminModele As String, _
                           ByVal cheminSortie As String)
    Const wdFormatXMLDocument As Long = 12   ' code du format .docx (liaison tardive)

    Dim wdApp As Object, wdDoc As Object
    Dim db As DAO.Database, rs As DAO.Recordset

    On Error GoTo Nettoyage

    ' Lire les données du client
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT Nom, Adresse, Ville FROM tblClients WHERE ID=" & idClient, _
        dbOpenSnapshot)
    If rs.EOF Then GoTo Nettoyage

    ' Démarrer Word, masqué, et créer un document à partir du modèle
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = False
    Set wdDoc = wdApp.Documents.Add(Template:=cheminModele)

    ' Remplir les signets
    RemplirSignet wdDoc, "Nom", Nz(rs!Nom, "")
    RemplirSignet wdDoc, "Adresse", Nz(rs!Adresse, "")
    RemplirSignet wdDoc, "Ville", Nz(rs!Ville, "")
    RemplirSignet wdDoc, "DateDuJour", Format$(Date, "dd/mm/yyyy")

    ' Enregistrer au format .docx
    wdDoc.SaveAs2 cheminSortie, wdFormatXMLDocument

Nettoyage:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
    Set rs = Nothing
    Set db = Nothing
End Sub
```

## Méthode 2 : rechercher-remplacer des marqueurs

Une approche plus légère consiste à parsemer le modèle de **marqueurs textuels** — par exemple `«Nom»` ou `{Ville}` — puis à les remplacer par les valeurs au moyen du moteur de recherche de Word :

```vba
Const wdReplaceAll As Long = 2
With wdDoc.Content.Find
    .Text = "«Nom»"
    .Replacement.Text = "Dupont"
    .Execute Replace:=wdReplaceAll
End With
```

Cette méthode est simple à mettre en œuvre, mais plus fragile que les signets : elle suppose que les marqueurs soient uniques et qu'ils ne soient pas fractionnés par la mise en forme. On la réserve aux cas simples ou aux modèles que l'on ne maîtrise pas suffisamment pour y poser des signets.

## Insérer un tableau de données

Un document généré comporte souvent un **tableau** alimenté par plusieurs enregistrements — les lignes d'une facture, par exemple. La leçon de performance de la section 22.3 s'applique pleinement ici : remplir un tableau Word cellule par cellule, en franchissant la frontière entre processus à chaque case, est lent. La méthode efficace consiste à **construire un texte délimité, à l'insérer en une fois, puis à le convertir en tableau** :

```vba
' Construire le contenu tabulé à partir d'un recordset
Dim texte As String
texte = "Article" & vbTab & "Quantité" & vbTab & "Prix" & vbCr   ' en-tête
Do Until rs.EOF
    texte = texte & rs!Article & vbTab & rs!Quantite & vbTab & rs!Prix & vbCr
    rs.MoveNext
Loop

' Insérer à l'emplacement d'un signet, puis convertir en tableau
Dim plage As Object
Set plage = wdDoc.Bookmarks("Tableau").Range
plage.Text = texte
plage.ConvertToTable Separator:=vbTab
```

La conversion d'un bloc de texte en tableau, en une opération, est très supérieure en performance au remplissage case par case.

## Construire un document de toutes pièces

Lorsqu'aucun modèle ne convient — parce que la structure du document est entièrement dynamique —, on peut **bâtir le document par code** : écrire du texte avec `Content.InsertAfter` ou par l'objet `Selection`, ajouter des paragraphes, appliquer des styles, insérer des tableaux via `Tables.Add`. Cette approche offre une liberté totale, au prix d'un code plus volumineux et d'une mise en forme entièrement à la charge du développeur. Elle se justifie quand la disposition varie d'un document à l'autre au point de ne pouvoir être figée dans un modèle.

## Le publipostage

Pour produire en série des documents répétitifs — un courrier identique par destinataire —, Word dispose de son propre mécanisme de **publipostage** (mail merge), que l'on peut piloter par automation. Le principe est de connecter le document de fusion à une **table ou requête Access** comme source de données, puis de déclencher la fusion, qui génère un document par enregistrement.

Cette voie est puissante pour les envois en nombre, mais sa mise en place est plus complexe que le remplissage d'un modèle unique, la connexion à la source de données ayant ses propres subtilités. On la réserve donc aux cas de production massive de documents similaires, là où remplir un modèle enregistrement par enregistrement serait moins efficace.

## Enregistrer, exporter en PDF, ou présenter à l'utilisateur

Le document généré peut connaître trois destins, selon le besoin.

On peut l'**enregistrer silencieusement** au format `.docx` (code 12) comme dans l'exemple ci-dessus, puis refermer Word. On peut aussi l'**exporter en PDF**, Word sachant produire directement ce format via le code 17 :

```vba
Const wdFormatPDF As Long = 17
wdDoc.SaveAs2 "C:\sortie\courrier.pdf", wdFormatPDF
```

Cette capacité rejoint les techniques d'export d'états en PDF abordées à la section 7.8, et ouvre la voie à des chaînes de traitement où le document généré est ensuite envoyé par courriel via Outlook (section 22.5).

Enfin, on peut souhaiter **présenter le document à l'utilisateur** pour qu'il le relise, l'imprime ou le signe. Dans ce cas, on règle `wdApp.Visible = True` et l'on **n'appelle pas** `Quit` : on libère ses propres variables objets, mais on laisse Word ouvert avec le document à l'écran. C'est une différence notable avec l'export Excel, qui se fait le plus souvent en silence : la génération de courriers se conclut fréquemment par la remise du document à l'utilisateur.

## Articulation avec le reste du chapitre et de la formation

Cette section prolonge l'automation Office et s'appuie sur plusieurs parties de la formation :

- Les **fondations de l'automation** (liaison, instance, libération) sont établies à la section 22.3.
- Le choix **liaison précoce / tardive** est traité à la section 2.6, et la compatibilité entre versions à la section 21.6.
- Les **jeux d'enregistrements** alimentant les documents relèvent du chapitre 9.
- L'**export en PDF** rejoint la section 7.8, et l'**envoi par courriel** du document généré, la section 22.5.
- La **gestion d'erreurs** garantissant le nettoyage relève du chapitre 13, et les considérations de **performance** (insertion en bloc plutôt que case par case) du chapitre 18.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans l'automation de Word :

- **Ne pas libérer les objets**, laissant s'accumuler des processus `WINWORD.EXE` fantômes en mémoire.
- **Modifier le modèle au lieu d'en créer un nouveau document** : utiliser `Documents.Add(Template:=…)` pour préserver le modèle.
- **Oublier la consommation du signet** lors de l'insertion, et perdre un signet que l'on comptait réutiliser.
- **Remplir un tableau case par case** plutôt que de convertir un bloc de texte, avec un effondrement des performances.
- **Définir des marqueurs de rechercher-remplacer ambigus** ou fractionnés par la mise en forme, qui ne seront pas correctement substitués.
- **Quitter Word alors que l'utilisateur devait relire le document**, ou à l'inverse laisser Word ouvert en mémoire alors qu'une génération silencieuse était attendue.

⏭️ [22.5. Automation avec Outlook — envoi d'emails et gestion des contacts](/22-api-windows-integration-office/05-automation-outlook.md)
