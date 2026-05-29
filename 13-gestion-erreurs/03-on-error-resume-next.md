🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.3. On Error Resume Next — utilisation contrôlée

## Introduction

`On Error Resume Next` est l'instruction de gestion d'erreurs la plus simple à écrire et, paradoxalement, la plus difficile à employer correctement. Là où `On Error GoTo` détourne l'exécution vers un traitement explicite, `On Error Resume Next` se contente de **poursuivre** : lorsqu'une erreur survient, l'instruction fautive est ignorée et l'exécution reprend à l'instruction suivante, comme si de rien n'était.

Cette apparente facilité en fait l'instruction la plus mal utilisée du langage. Placée sans discernement, elle transforme une application en boîte noire où les erreurs disparaissent sans laisser de trace, où les bugs survivent silencieusement et où les données peuvent se corrompre à l'insu de tous. Pourtant, employée avec rigueur, elle est irremplaçable pour une catégorie précise de situations. Tout l'enjeu de cette section tient dans son sous-titre : l'**utilisation contrôlée**. Il ne s'agit pas de bannir `On Error Resume Next`, mais d'en circonscrire l'usage à un cadre strict, fondé sur deux règles indissociables : une portée minimale et une vérification immédiate.

## Le comportement de On Error Resume Next

`On Error Resume Next` indique à VBA que, en cas d'erreur d'exécution, il ne doit ni s'arrêter ni dérouter l'exécution, mais simplement passer à l'instruction suivante. L'erreur n'est pas perdue pour autant : l'objet `Err` est renseigné (numéro, description, source), de sorte que le code peut, **s'il le souhaite**, vérifier ce qui s'est passé.

```vba
Sub Demonstration()
    On Error Resume Next

    Dim x As Integer
    x = 10 / 0                 ' erreur 11 : ignorée, l'exécution continue
    Debug.Print "Cette ligne s'exécute malgré l'erreur précédente"
    Debug.Print "Err.Number vaut : " & Err.Number   ' affiche 11
End Sub
```

L'expression « comme si de rien n'était » est trompeuse : il s'est bien passé quelque chose, mais c'est au développeur d'aller le constater en interrogeant `Err`. Sans cette vérification, l'erreur est purement et simplement avalée.

## Le danger : l'erreur silencieuse

Le mauvais usage de `On Error Resume Next` consiste à l'activer en tête de procédure et à ne plus jamais s'en préoccuper. Le résultat est une procédure dans laquelle **aucune erreur ne sera jamais signalée**, où qu'elle survienne :

```vba
Sub MauvaisExemple()
    On Error Resume Next        ' DANGER : masque tout jusqu'à la fin

    Set db = CurrentDb
    Set rst = db.OpenRecordset("T_Inexistante")   ' échoue silencieusement
    rst.MoveFirst                                  ' échoue : rst est Nothing
    montant = rst!Total                            ' échoue à son tour
    ' ... le traitement « réussit » sans rien avoir fait ...
End Sub
```

Ici, la table n'existe pas, le recordset n'est jamais ouvert, et pourtant la procédure se termine sans le moindre avertissement. Pire : elle a peut-être affiché un message de succès, écrit des valeurs nulles, ou laissé l'utilisateur croire qu'une opération avait abouti alors qu'elle a entièrement échoué. C'est la pire situation possible en informatique de gestion — une **défaillance silencieuse** —, bien plus pernicieuse qu'une erreur visible, car personne n'en a connaissance avant qu'il ne soit trop tard.

Cette utilisation aveugle est si répandue qu'elle mérite d'être désignée explicitement comme une pratique à proscrire. `On Error Resume Next` ne doit jamais couvrir un bloc de code arbitraire : il doit cerner une instruction précise dont on sait, par avance, pourquoi l'on accepte qu'elle échoue.

## La règle d'or : portée minimale et vérification immédiate

L'usage contrôlé de `On Error Resume Next` repose sur un principe simple, qu'il faut s'imposer sans exception : **activer la protection juste avant l'instruction risquée, vérifier `Err.Number` immédiatement après, puis rétablir aussitôt la gestion normale**.

Concrètement, cela se traduit par un encadrement réduit à sa plus stricte expression :

```vba
On Error Resume Next
Set appExcel = CreateObject("Excel.Application")   ' instruction risquée, isolée
On Error GoTo 0                                     ' on referme aussitôt

If appExcel Is Nothing Then
    MsgBox "Excel ne semble pas installé sur ce poste.", vbExclamation
    Exit Sub
End If
```

La zone protégée ne contient qu'une seule instruction. Toute erreur survenant ailleurs dans la procédure restera traitée normalement. Et la vérification qui suit (`If appExcel Is Nothing`) examine immédiatement le résultat, avant qu'une autre opération ne vienne brouiller les pistes. Ces trois temps — activer, exécuter l'unique instruction, refermer et vérifier — forment un bloc indissociable.

## Le pattern de vérification en ligne

La forme la plus courante de cet encadrement consiste à tester explicitement `Err.Number` juste après l'instruction protégée, puis à revenir à la gestion d'erreurs habituelle de la procédure :

```vba
Public Sub OuvrirAvecExcel()
    Dim appExcel As Object

    On Error Resume Next
    Set appExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        MsgBox "Excel ne semble pas installé sur ce poste.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo GestionErreur        ' retour à la gestion normale

    appExcel.Visible = True
    ' ... suite du traitement, désormais sous protection classique ...

    Exit Sub
GestionErreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
End Sub
```

Ce schéma illustre la cohabitation des deux instructions de gestion d'erreurs au sein d'une même procédure : `On Error Resume Next` pour la seule instruction dont l'échec est prévu et géré sur place, puis `On Error GoTo` pour tout le reste. La présence de `Err.Clear` après le traitement garantit que l'erreur ne « traîne » pas dans l'objet `Err` au-delà de l'endroit où elle a été prise en charge.

## L'objet Err après Resume Next : précautions

Une subtilité, souvent à l'origine de diagnostics erronés, mérite une attention particulière : **une instruction qui réussit ne réinitialise pas l'objet `Err`**. Seuls `Err.Clear`, une instruction `Resume`, une sortie de procédure (`Exit Sub`...) ou une nouvelle instruction `On Error` remettent `Err` à zéro ; et une nouvelle erreur en écrase les valeurs.

Cette mécanique a une conséquence directe. Considérons deux instructions protégées dont la première échoue et la seconde réussit :

```vba
On Error Resume Next
Kill "C:\Temp\inexistant.tmp"    ' échoue → Err.Number = 53
x = 5 + 3                        ' réussit, mais ne réinitialise PAS Err
If Err.Number <> 0 Then          ' Err.Number vaut toujours 53 !
    MsgBox "L'addition a échoué"  ' message FAUX : c'est Kill qui avait échoué
End If
```

Le test attribue à tort une erreur à l'addition, alors qu'elle provenait du `Kill` exécuté trois lignes plus haut. C'est la démonstration la plus nette de la règle d'or : il faut tester `Err.Number` **immédiatement** après l'instruction concernée, avant toute autre opération, et nettoyer `Err` (par `Err.Clear`) avant la vérification suivante. Toute distance entre l'instruction risquée et sa vérification ouvre la porte à ce type de confusion.

## Cas d'usage légitimes

Une fois ces règles posées, `On Error Resume Next` se révèle l'outil le plus adapté à plusieurs situations bien identifiées.

### Tester l'existence d'un objet

Lorsque l'on souhaite savoir si un objet existe sans provoquer d'erreur fatale, l'approche la plus simple consiste à tenter d'y accéder sous `On Error Resume Next`, puis à conclure d'après `Err.Number`. C'est notamment le cas pour vérifier la présence d'un contrôle sur un formulaire :

```vba
Public Function ControleExiste(frm As Form, nomControle As String) As Boolean
    Dim ctl As Control

    On Error Resume Next
    Set ctl = frm.Controls(nomControle)
    ControleExiste = (Err.Number = 0)
    On Error GoTo 0
End Function
```

La fonction renvoie `True` si l'accès au contrôle a réussi (aucune erreur), `False` sinon. Le même principe s'applique pour tester l'existence d'un champ dans un recordset, d'une propriété sur un objet, ou d'une table dans la base. Cette technique est concise et fiable, à condition de rester fidèle à l'encadrement strict.

### Opérations facultatives dont l'échec est sans gravité

Certaines opérations sont accessoires : leur échec ne doit pas interrompre le traitement principal. La suppression d'un fichier temporaire qui n'existe peut-être pas en est l'exemple type :

```vba
On Error Resume Next
Kill "C:\Temp\export_provisoire.csv"   ' si absent, l'erreur 53 est sans importance
On Error GoTo 0
```

Plutôt que de tester au préalable l'existence du fichier, on accepte que la suppression échoue silencieusement lorsqu'il n'y a rien à supprimer. L'intention est claire et l'échec, ici, est un non-événement.

### Le nettoyage des ressources

Le bloc de nettoyage présenté dans la section [13.2](/13-gestion-erreurs/02-on-error-goto.md) illustre un troisième usage légitime. Le code de libération des objets (fermeture d'un recordset, remise à `Nothing`) ne doit lui-même jamais provoquer d'erreur — par exemple si l'objet n'a pas pu être créé et vaut déjà `Nothing`. On y place donc un `On Error Resume Next` :

```vba
SortieProcedure:
    On Error Resume Next            ' le nettoyage ne doit jamais échouer
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
```

Ici, l'on ne cherche pas à vérifier les erreurs : on sait que la phase de nettoyage doit aboutir quoi qu'il advienne, et que l'échec d'une libération individuelle est sans conséquence à ce stade.

## Combiner Resume Next et On Error GoTo

Ces exemples montrent que les deux instructions ne s'opposent pas : elles se complètent au sein d'une même procédure. Le schéma le plus robuste utilise `On Error GoTo` comme protection générale, et bascule ponctuellement vers `On Error Resume Next` pour encadrer une instruction précise dont l'échec est anticipé, avant de rétablir la protection générale :

```vba
On Error GoTo GestionErreur        ' protection générale en vigueur
' ... traitement principal ...

On Error Resume Next               ' parenthèse pour une instruction risquée
SupprimerFichierTemporaire
On Error GoTo GestionErreur        ' on referme la parenthèse

' ... suite du traitement, de nouveau sous protection générale ...
```

Ce va-et-vient maîtrisé est la marque d'une gestion d'erreurs réfléchie : chaque zone du code est couverte par le mécanisme le mieux adapté à sa nature.

## Resume Next ne corrige rien

Il faut enfin dissiper un malentendu fréquent : `On Error Resume Next` ne **répare** rien. Il ne fait que poursuivre l'exécution ; l'état du programme après l'instruction fautive reste exactement ce qu'il était au moment de l'échec. Si `Set obj = CreateObject(...)` échoue, `obj` vaut `Nothing` — et si l'on enchaîne sur `obj.Methode` sans vérification, on déclenche une erreur 91, elle aussi silencieusement ignorée si la protection est toujours active. Les défaillances s'enchaînent alors en cascade, chacune masquant la précédente.

C'est pourquoi la vérification n'est pas optionnelle : continuer après une erreur n'a de sens que si l'on prend acte de cette erreur et que l'on adapte la suite en conséquence. Ignorer une erreur et ignorer ses conséquences sont deux choses différentes ; seule la première est parfois acceptable, jamais la seconde.

## Pièges courants

Les erreurs récurrentes dans l'emploi de `On Error Resume Next` découlent presque toutes d'un non-respect de la règle d'or :

- **Activer `On Error Resume Next` en tête de procédure et l'oublier** : toutes les erreurs sont masquées jusqu'à la fin, transformant la procédure en défaillance silencieuse potentielle.
- **Ne pas vérifier `Err.Number` après l'instruction protégée** : le traitement se poursuit avec un état invalide, sans que personne ne le sache.
- **Vérifier `Err.Number` trop tard**, après d'autres instructions : une opération réussie ne réinitialisant pas `Err`, on risque d'attribuer une ancienne erreur à la mauvaise instruction.
- **Oublier `Err.Clear` entre deux vérifications successives** : une erreur antérieure non effacée provoque un faux positif.
- **Confondre « ignorer l'erreur » et « corriger l'erreur »** : poursuivre l'exécution ne rétablit pas un état cohérent.

## En résumé

`On Error Resume Next` poursuit l'exécution à l'instruction suivante en cas d'erreur, sans interruption ni déroutement. Sa simplicité en fait l'instruction la plus dangereuse du langage lorsqu'elle est employée sans cadre : étendue à un bloc entier, elle engendre des défaillances silencieuses, le pire des défauts en informatique de gestion.

Son usage légitime obéit à une règle d'or en deux volets : **portée minimale** (encadrer une seule instruction) et **vérification immédiate** (tester `Err.Number` aussitôt après, avant toute autre opération, et nettoyer `Err`). Dans ce cadre, l'instruction excelle pour tester l'existence d'un objet, tolérer l'échec d'une opération facultative, ou sécuriser un bloc de nettoyage. Elle se combine naturellement avec `On Error GoTo`, qui assure la protection générale tandis que `Resume Next` encadre les instructions au comportement anticipé. Il faut enfin garder à l'esprit qu'elle n'a aucun pouvoir réparateur : continuer après une erreur n'a de valeur que si l'on en tire les conséquences.

Après avoir vu comment intercepter les erreurs avec `On Error GoTo` puis comment les traiter au cas par cas avec `On Error Resume Next`, la section suivante examine en détail l'objet qui décrit chaque erreur : l'objet `Err`.

---


⏭️ [13.4. Objet Err — numéro, description, source](/13-gestion-erreurs/04-objet-err.md)
