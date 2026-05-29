🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 22.8. Parsing JSON en VBA (sans bibliothèque externe)

Le JSON est la langue commune des services REST : c'est sous ce format que la section 22.7 reçoit les réponses des API qu'elle interroge. Or VBA ne dispose d'**aucun support natif du JSON** — ni pour l'analyser, ni pour le produire. Cette section montre comment combler ce manque **sans recourir à aucune bibliothèque externe**, au moyen de techniques qui fonctionnent aussi bien en 32 qu'en 64 bits.

L'analyse (parsing) en est le cœur ; la construction du JSON, utile pour les corps de requête de la section 22.7, est abordée en fin de section.

## VBA et JSON : aucun support natif

Faute de fonction intégrée, trois voies « sans bibliothèque » se présentent :

- Le **ScriptControl**, qui délègue l'analyse au moteur JavaScript de Windows — séduisant, mais **inutilisable en 64 bits**, comme expliqué ci-dessous.
- L'**extraction pragmatique** par expression régulière, simple mais fragile, réservée aux réponses plates et connues.
- L'**analyseur maison**, un petit code récursif robuste et portable — la solution de référence dès que la structure dépasse le trivial.

## Pourquoi éviter le ScriptControl

De nombreux tutoriels anciens proposent d'analyser le JSON via `MSScriptControl.ScriptControl`, en évaluant le texte par le moteur JScript. Cette approche a un défaut rédhibitoire : le ScriptControl est un composant **32 bits uniquement**, qui **ne fonctionne pas dans Office 64 bits**. Le 64 bits étant désormais l'architecture installée par défaut (section 21.7), cette voie n'est plus viable pour un déploiement moderne. On l'écarte donc, malgré sa fréquence dans la documentation ancienne.

## Extraction pragmatique pour les cas simples

Lorsqu'on ne cherche qu'**une valeur précise dans une réponse plate et connue**, une expression régulière suffit. Le composant `VBScript.RegExp`, intégré à Windows et compatible avec les deux architectures, permet d'extraire la valeur associée à une clé :

```vba
' Extrait la valeur de chaîne associée à une clé (réponse plate connue — fragile)
Public Function ValeurJSONSimple(ByVal texte As String, ByVal cle As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = """" & cle & """\s*:\s*""([^""]*)"""
    re.Global = False
    If re.Test(texte) Then
        ValeurJSONSimple = re.Execute(texte)(0).SubMatches(0)
    End If
End Function
```

Cette méthode est expéditive, mais **fragile** : elle ne gère que les valeurs de chaîne, échoue sur les structures imbriquées, ignore les caractères échappés et se brise au moindre changement de format. On la réserve aux cas où l'on connaît parfaitement une réponse simple et où l'on n'a besoin que d'un champ.

## Un analyseur JSON maison

Pour traiter du JSON quelconque — objets imbriqués, tableaux, types variés —, la solution robuste consiste à écrire un petit **analyseur récursif** qui transforme le texte en structures VBA : un `Scripting.Dictionary` pour chaque objet (chapitre 3.7), une `Collection` pour chaque tableau, et des valeurs scalaires pour le reste. Le module suivant, placé dans un module standard, en fournit une implémentation complète :

```vba
Option Explicit

Private mJSON As String
Private mPos As Long

' === Point d'entrée : renvoie un Dictionary (objet), une Collection
'     (tableau) ou une valeur scalaire ===
Public Function JSONParse(ByVal texte As String) As Variant
    Dim resultat As Variant
    mJSON = texte
    mPos = 1
    PasserEspaces
    AffecterVariant resultat, AnalyserValeur()
    If IsObject(resultat) Then
        Set JSONParse = resultat
    Else
        JSONParse = resultat
    End If
End Function

' Affectation uniforme, que la source soit un objet ou un scalaire
Private Sub AffecterVariant(ByRef cible As Variant, ByRef source As Variant)
    If IsObject(source) Then
        Set cible = source
    Else
        cible = source
    End If
End Sub

' Analyse d'une valeur selon son premier caractère
Private Function AnalyserValeur() As Variant
    PasserEspaces
    Select Case Mid$(mJSON, mPos, 1)
        Case "{": Set AnalyserValeur = AnalyserObjet()
        Case "[": Set AnalyserValeur = AnalyserTableau()
        Case """": AnalyserValeur = AnalyserChaine()
        Case "t", "f": AnalyserValeur = AnalyserBooleen()
        Case "n": AnalyserValeur = AnalyserNull()
        Case Else: AnalyserValeur = AnalyserNombre()
    End Select
End Function

' Objet JSON -> Scripting.Dictionary
Private Function AnalyserObjet() As Object
    Dim dico As Object, cle As String, valeur As Variant, suivant As String
    Set dico = CreateObject("Scripting.Dictionary")

    mPos = mPos + 1                         ' passer "{"
    PasserEspaces
    If Mid$(mJSON, mPos, 1) = "}" Then      ' objet vide
        mPos = mPos + 1
        Set AnalyserObjet = dico
        Exit Function
    End If

    Do
        PasserEspaces
        cle = AnalyserChaine()              ' la clé est une chaîne
        PasserEspaces
        mPos = mPos + 1                     ' passer ":"
        AffecterVariant valeur, AnalyserValeur()
        dico.Add cle, valeur
        PasserEspaces
        suivant = Mid$(mJSON, mPos, 1)
        mPos = mPos + 1                     ' passer "," ou "}"
    Loop While suivant = ","

    Set AnalyserObjet = dico
End Function

' Tableau JSON -> Collection
Private Function AnalyserTableau() As Object
    Dim col As Collection, valeur As Variant, suivant As String
    Set col = New Collection

    mPos = mPos + 1                         ' passer "["
    PasserEspaces
    If Mid$(mJSON, mPos, 1) = "]" Then      ' tableau vide
        mPos = mPos + 1
        Set AnalyserTableau = col
        Exit Function
    End If

    Do
        AffecterVariant valeur, AnalyserValeur()
        col.Add valeur
        PasserEspaces
        suivant = Mid$(mJSON, mPos, 1)
        mPos = mPos + 1                     ' passer "," ou "]"
    Loop While suivant = ","

    Set AnalyserTableau = col
End Function

' Chaîne JSON -> String (échappements courants gérés)
Private Function AnalyserChaine() As String
    Dim resultat As String, c As String, e As String, hex As String

    mPos = mPos + 1                         ' passer le guillemet ouvrant
    Do While mPos <= Len(mJSON)
        c = Mid$(mJSON, mPos, 1)
        mPos = mPos + 1
        Select Case c
            Case """"                       ' guillemet fermant
                AnalyserChaine = resultat
                Exit Function
            Case "\"                        ' séquence d'échappement
                e = Mid$(mJSON, mPos, 1)
                mPos = mPos + 1
                Select Case e
                    Case """": resultat = resultat & """"
                    Case "\": resultat = resultat & "\"
                    Case "/": resultat = resultat & "/"
                    Case "n": resultat = resultat & vbLf
                    Case "r": resultat = resultat & vbCr
                    Case "t": resultat = resultat & vbTab
                    Case "u"                ' \uXXXX
                        hex = Mid$(mJSON, mPos, 4)
                        mPos = mPos + 4
                        resultat = resultat & ChrW$(CLng("&H" & hex))
                    Case Else: resultat = resultat & e
                End Select
            Case Else
                resultat = resultat & c
        End Select
    Loop
    AnalyserChaine = resultat
End Function

' Booléen
Private Function AnalyserBooleen() As Boolean
    If Mid$(mJSON, mPos, 4) = "true" Then
        mPos = mPos + 4
        AnalyserBooleen = True
    Else
        mPos = mPos + 5                     ' "false"
        AnalyserBooleen = False
    End If
End Function

' Null
Private Function AnalyserNull() As Variant
    mPos = mPos + 4                         ' "null"
    AnalyserNull = Null
End Function

' Nombre -> Double. Val est indépendant de la locale, contrairement à CDbl
Private Function AnalyserNombre() As Variant
    Dim debut As Long
    debut = mPos
    Do While mPos <= Len(mJSON)
        Select Case Mid$(mJSON, mPos, 1)
            Case "0" To "9", "-", "+", ".", "e", "E"
                mPos = mPos + 1
            Case Else
                Exit Do
        End Select
    Loop
    AnalyserNombre = Val(Mid$(mJSON, debut, mPos - debut))
End Function

' Passer les caractères d'espacement
Private Sub PasserEspaces()
    Do While mPos <= Len(mJSON)
        Select Case Mid$(mJSON, mPos, 1)
            Case " ", vbTab, vbCr, vbLf
                mPos = mPos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub
```

Un point de cet analyseur mérite d'être souligné, car il touche un problème de **localisation** récurrent pour un public francophone : la conversion des nombres emploie `Val` et non `CDbl`. En effet, le JSON utilise toujours le **point** comme séparateur décimal, alors que `CDbl` suit la locale du poste et attendrait une virgule en configuration française — `CDbl("3.14")` y échouerait. `Val`, lui, interprète invariablement le point comme séparateur décimal, indépendamment de la locale. Cette distinction rejoint les questions de séparateur décimal abordées à la section 11.7.

Cet analyseur couvre les cas courants des réponses d'API. Il constitue une **base solide** plutôt qu'un validateur exhaustif : on pourra le durcir au besoin (échappements rares, paires de substitution Unicode, validation stricte, optimisation de la concaténation pour de très longues chaînes).

## Naviguer dans le résultat

L'analyseur restituant des `Dictionary` et des `Collection`, on navigue dans le résultat par clé (accès au dictionnaire) et par indice (parcours de la collection, **indexée à partir de 1**). Un principe guide les affectations : on emploie `Set` pour un nœud qui est un objet (objet ou tableau imbriqué) et `=` pour une valeur scalaire. En reprenant une réponse obtenue via la section 22.7 :

```vba
Dim reponse As String
reponse = AppelGET("https://api.exemple.fr/clients/42")   ' cf. section 22.7

Dim client As Object
Set client = JSONParse(reponse)

Debug.Print client("nom")                  ' valeur scalaire
Debug.Print client("adresse")("ville")     ' objet imbriqué

' Parcourir un tableau d'objets
Dim commandes As Object, i As Long
Set commandes = client("commandes")        ' une Collection
For i = 1 To commandes.Count
    Debug.Print commandes(i)("reference"), commandes(i)("montant")
Next i
```

## Insérer le résultat dans Access

Une fois la réponse analysée et parcourue, son intégration dans la base relève des techniques habituelles d'Access : écriture par **Recordset DAO** (`AddNew` / `Update`, chapitre 9) ou par requêtes `INSERT` (chapitre 11). On parcourt par exemple la collection des objets, et l'on insère un enregistrement par élément, en validant et en transformant les valeurs au passage.

## Construire du JSON pour l'envoi

L'opération inverse — produire un corps JSON à transmettre à une requête `POST` (section 22.7) — ne dispose pas davantage de fonction native. Pour un objet simple, on construit la chaîne à la main, en veillant à **échapper** les valeurs textuelles et à formater les nombres avec un point décimal :

```vba
' Échappe une chaîne pour l'insérer dans du JSON
Private Function EchapperJSON(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EchapperJSON = s
End Function

' Construit un objet JSON simple
Dim corps As String
corps = "{""nom"":""" & EchapperJSON(nom) & """," & _
        """montant"":" & Trim$(Str$(montant)) & "}"
```

Ici encore, la localisation impose une précaution : `Str` produit toujours un séparateur décimal **point**, indépendamment de la locale, là où `CStr` ou `Format` suivraient la configuration française et inséreraient une virgule, invalide en JSON. Pour des structures complexes, on étoffe ce principe en un véritable sérialiseur, ou l'on génère le JSON à partir d'un `Dictionary`.

## Articulation avec le reste du chapitre et de la formation

Cette section prolonge directement la précédente et s'appuie sur plusieurs parties de la formation :

- Le **JSON analysé** provient des services REST de la section 22.7, à laquelle elle fournit aussi la construction des corps de requête.
- Les structures produites — **`Scripting.Dictionary`** et **`Collection`** — relèvent de la section 3.7.
- L'**insertion dans Access** mobilise les jeux d'enregistrements (chapitre 9) et les requêtes `INSERT` (chapitre 11).
- La précaution sur le **séparateur décimal** rejoint la section 11.7.
- L'inadéquation du **ScriptControl en 64 bits** renvoie à la section 21.7, et la gestion d'erreurs au chapitre 13.

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans le traitement du JSON en VBA :

- **Recourir au ScriptControl**, inutilisable en Office 64 bits (section 21.7).
- **Convertir les nombres avec `CDbl`** (ou les formater avec `CStr`/`Format`), qui suivent la locale et trahissent le séparateur décimal du JSON ; employer `Val` à l'analyse et `Str` à la construction.
- **S'appuyer sur l'extraction par expression régulière** au-delà des réponses plates et connues, là où elle se révèle fragile.
- **Confondre `Set` et `=`** lors de la navigation, selon que le nœud est un objet ou une valeur scalaire.
- **Oublier d'échapper les valeurs textuelles** lors de la construction d'un corps JSON.
- **Indexer une collection à partir de 0**, alors qu'une `Collection` VBA commence à 1.

⏭️ [22.9. Lecture/écriture du registre Windows](/22-api-windows-integration-office/09-registre-windows.md)
