🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 19.4. Tests unitaires en VBA — framework léger maison

VBA ne fournit aucun framework de test intégré. Plutôt que de s'en passer, on peut en construire un **léger**, en quelques dizaines de lignes : zéro dépendance, contrôle total, et un dispositif suffisant pour automatiser et rejouer ses vérifications. C'est aussi instructif, car cela met à nu ce qu'est, au fond, un framework de test. (Pour qui souhaite davantage — découverte automatique des tests, explorateur intégré à l'éditeur —, l'extension *Rubberduck* ajoute un véritable cadre de tests à la VBE ; elle est évoquée au [chapitre 24.8](../24-bonnes-pratiques-ressources/08-communaute-ressources.md).)

## Ce dont un framework a besoin

Un cadre de test minimal repose sur quelques éléments : des **assertions** (comparer une valeur attendue à une valeur obtenue), un moyen d'**accumuler les résultats** (compteurs et détail des échecs), un **lanceur** qui exécute les tests, un **compte rendu** lisible, et de quoi gérer la **mise en place et le nettoyage** (setup/teardown). C'est tout — et tout cela tient dans un module.

## Un framework en un module

Le module suivant rassemble le cycle d'exécution, les assertions et la mécanique interne. Il écrit son compte rendu dans la fenêtre Exécution immédiate.

```vba
' ============================================================
'  modTest — framework de test unitaire léger
' ============================================================
Option Compare Database
Option Explicit

Private mNbAssert As Long        ' assertions évaluées
Private mNbEchecs As Long
Private mTestNom  As String      ' nom du test en cours
Private mEchecs   As String      ' détail des échecs accumulés

' ---------- Cycle d'exécution ----------

Public Sub Suite_Debut(ByVal nom As String)
    mNbAssert = 0: mNbEchecs = 0: mEchecs = ""
    Debug.Print String(52, "=")
    Debug.Print "SUITE : " & nom
    Debug.Print String(52, "=")
End Sub

Public Sub Test(ByVal nom As String)
    mTestNom = nom               ' contexte des assertions qui suivent
End Sub

Public Sub Suite_Fin()
    Debug.Print String(52, "-")
    Debug.Print mNbAssert & " assertion(s) — " & mNbEchecs & " échec(s)"
    If Len(mEchecs) > 0 Then Debug.Print vbCrLf & "ÉCHECS :" & vbCrLf & mEchecs
    Debug.Print IIf(mNbEchecs = 0, ">>> SUCCÈS", ">>> ÉCHEC")
End Sub

' ---------- Assertions ----------

Public Sub AssertVrai(ByVal condition As Boolean, Optional ByVal libelle As String = "")
    Verifier condition, "attendu : Vrai", libelle
End Sub

Public Sub AssertFaux(ByVal condition As Boolean, Optional ByVal libelle As String = "")
    Verifier (Not condition), "attendu : Faux", libelle
End Sub

Public Sub AssertEgal(ByVal attendu As Variant, ByVal obtenu As Variant, _
                      Optional ByVal libelle As String = "")
    Verifier SontEgaux(attendu, obtenu), _
             "attendu <" & Aff(attendu) & "> obtenu <" & Aff(obtenu) & ">", libelle
End Sub

Public Sub AssertNull(ByVal valeur As Variant, Optional ByVal libelle As String = "")
    Verifier IsNull(valeur), "valeur non Null : " & Aff(valeur), libelle
End Sub

' ---------- Mécanique interne ----------

Private Sub Verifier(ByVal succes As Boolean, ByVal detail As String, ByVal libelle As String)
    mNbAssert = mNbAssert + 1
    If Not succes Then
        mNbEchecs = mNbEchecs + 1
        mEchecs = mEchecs & " - [" & mTestNom & "] " & _
                  IIf(Len(libelle) > 0, libelle & " : ", "") & detail & vbCrLf
    End If
End Sub

Private Function SontEgaux(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) And IsNull(b) Then
        SontEgaux = True
    ElseIf IsNull(a) Or IsNull(b) Then
        SontEgaux = False
    Else
        SontEgaux = (a = b)
    End If
End Function

Private Function Aff(ByVal v As Variant) As String
    If IsNull(v) Then
        Aff = "Null"
    ElseIf IsEmpty(v) Then
        Aff = "Empty"
    Else
        Aff = CStr(v)
    End If
End Function
```

Deux points de robustesse méritent l'attention. La comparaison d'égalité passe par `SontEgaux`, qui traite le cas `Null` à part : sans cela, comparer deux `Null` avec `=` renverrait `Null` et provoquerait une erreur en l'affectant à un booléen. Et `Aff` produit une représentation lisible de chaque valeur, y compris `Null` et `Empty`, pour des messages d'échec exploitables.

> ℹ️ Le compte rendu s'appuie sur `Debug.Print` et la fenêtre Exécution immédiate, présentés à la [section 19.2](02-debug-print-assert.md).

## Écrire des tests

Le développeur écrit ensuite ses suites de tests, chacune ouverte par `Suite_Debut` et close par `Suite_Fin`, les assertions étant regroupées par `Test`. On suit le schéma *Arranger–Agir–Vérifier*.

```vba
Public Sub TestsCalcul()
    Suite_Debut "Module Calcul"

    Test "Addition simple"
    AssertEgal 5, Additionner(2, 3)

    Test "TVA à 20 %"
    AssertEgal 120, AppliquerTVA(100, 0.2)

    Test "Libellé vide accepté"
    AssertVrai EstValide(""), "une chaîne vide doit être valide"

    Suite_Fin
End Sub
```

Un **lanceur** appelle toutes les suites. VBA n'offrant pas de découverte automatique des procédures de test (sans outil externe), on y inscrit explicitement chaque suite.

```vba
Public Sub LancerTousLesTests()
    TestsCalcul
    TestsClients
    '  ... ajouter ici chaque nouvelle suite ...
End Sub
```

On exécute l'ensemble en tapant `LancerTousLesTests` dans la fenêtre Exécution immédiate (Ctrl+G) ; le compte rendu s'y affiche aussitôt, suite par suite.

## Tester les chemins d'erreur

Vérifier qu'un code lève bien l'erreur attendue fait partie des tests. Comme VBA ne permet pas de passer aisément un bloc de code à exécuter, le test emploie lui-même `On Error Resume Next` puis contrôle `Err.Number`.

```vba
Public Sub TestDivisionParZero()
    Suite_Debut "Division"

    Test "Division par zéro lève l'erreur 11"
    Dim r As Double
    On Error Resume Next
    Err.Clear
    r = Diviser(10, 0)
    AssertEgal 11, Err.Number, "doit lever une division par zéro"
    On Error GoTo 0

    Suite_Fin
End Sub
```

> ℹ️ L'objet `Err` est détaillé au [chapitre 13.4](../13-gestion-erreurs/04-objet-err.md). (Pour automatiser ce motif, on peut bâtir une assertion qui exécute une procédure nommée via `Application.Run` et inspecte l'erreur résultante.)

## Mise en place et nettoyage

Beaucoup de tests requièrent un état initial à préparer puis à défaire. Dans cette approche en module, il suffit d'appeler la mise en place au début de la suite et le nettoyage à la fin. Pour les tests de données, on combine avantageusement le framework avec l'**isolation par transaction** de la [section 19.3](03-tests-modules-donnees.md) : on ouvre une transaction au début de la suite et on l'annule à la fin, ne laissant aucun résidu.

```vba
Public Sub TestsClients()
    Dim ws As DAO.Workspace
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans                        ' mise en place : isolation
    On Error GoTo Fin

    Suite_Debut "Dépôt Clients"

    Test "Ajout puis relecture"
    CurrentDb.Execute "INSERT INTO Clients (Nom) VALUES ('TEST_UNIT')", dbFailOnError
    AssertEgal 1, DCount("*", "Clients", "Nom = 'TEST_UNIT'")

    Suite_Fin

Fin:
    ws.Rollback                          ' nettoyage : la base retrouve son état
End Sub
```

> ℹ️ Les transactions DAO sont traitées au [chapitre 14.2](../14-transactions/02-transactions-dao.md) ; les jeux de données et environnements de test à la [section 19.6](06-simulation-donnees-test.md).

## Conserver le compte rendu

La fenêtre Exécution immédiate suffit en développement. Pour garder une trace des exécutions — ou afficher les résultats dans un formulaire —, on dirige le compte rendu vers une table ou un fichier, en reprenant les techniques de journalisation de la [section 19.5](05-journalisation-evenements.md). Il suffit pour cela de remplacer les `Debug.Print` du module par un appel à une routine de journalisation.

## Concevoir de bons tests

Le framework ne vaut que par la qualité des tests. Quelques principes : un test vérifie **un seul comportement** et porte un nom explicite ; les tests sont **indépendants** les uns des autres (aucune dépendance d'ordre, comme rappelé à la [section 19.3](03-tests-modules-donnees.md)) ; ils sont **rapides** — un test unitaire ne devrait pas toucher le réseau, d'où l'usage de faux dépôts en mémoire ; et ils sont **déterministes**, ne dépendant ni de la date courante, ni de valeurs NuméroAuto, que l'on injecte ou capture plutôt que de coder en dur. Ces tests, accumulés, forment le filet qui autorise à refactorer sans crainte.

> ℹ️ Les faux dépôts isolant la logique de la base sont décrits à la [section 19.3](03-tests-modules-donnees.md) ; le refactoring sécurisé par les tests au [chapitre 24.5](../24-bonnes-pratiques-ressources/05-revue-code-refactoring.md).

## Limites de l'approche maison

Ce framework léger a ses bornes, qu'il vaut mieux connaître. Il n'y a pas de **découverte automatique** : chaque suite doit être inscrite dans le lanceur. L'**isolation** entre tests n'est pas fournie d'office — l'état du module et celui de la base doivent être maîtrisés par conception. Et l'on ne dispose pas d'une **interface graphique** de résultats intégrée à l'éditeur, ni d'exécution test par test. Lorsque ces besoins deviennent pressants, une extension dédiée comme Rubberduck — qui apporte annotations, explorateur de tests et simulacres — constitue l'étape suivante.

## Points clés à retenir

- VBA n'a pas de framework de test natif ; un dispositif maison en un module — assertions, accumulation des résultats, lanceur, compte rendu — suffit et reste sans dépendance.
- Les assertions doivent gérer les cas particuliers (comparaison de `Null`) et produire des messages d'échec lisibles.
- On écrit des suites suivant *Arranger–Agir–Vérifier*, inscrites explicitement dans un lanceur exécuté depuis la fenêtre Exécution immédiate.
- Les chemins d'erreur se testent via `On Error Resume Next` et le contrôle d'`Err.Number`.
- Pour les tests de données, on combine le framework avec l'isolation par transaction (`BeginTrans`/`Rollback`) ; le compte rendu peut être journalisé pour être conservé.
- De bons tests sont uniques, indépendants, rapides et déterministes ; ensemble, ils sécurisent le refactoring.
- L'approche maison ne fournit ni découverte automatique, ni isolation native, ni interface intégrée : Rubberduck répond à ces besoins avancés.

---


⏭️ [19.5. Journalisation d'événements applicatifs](/19-debogage-tests/05-journalisation-evenements.md)
