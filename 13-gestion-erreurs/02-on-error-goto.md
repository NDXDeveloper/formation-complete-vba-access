🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.2. On Error GoTo — structure robuste de gestion d'erreurs

## Introduction

La section précédente a permis de reconnaître les erreurs ; il s'agit maintenant de les intercepter. En VBA, l'instruction `On Error GoTo` est le pilier de cette interception : c'est elle qui permet de détourner le comportement par défaut d'Access — l'arrêt brutal et la boîte de dialogue technique — vers un traitement maîtrisé, écrit par le développeur.

Mais au-delà de l'instruction elle-même, l'enjeu de cette section est la **structure**. Une gestion d'erreurs efficace ne consiste pas à saupoudrer quelques `On Error` au gré des besoins : elle repose sur un schéma reproductible, appliqué de façon systématique à chaque procédure susceptible d'échouer. Ce schéma — corps de procédure, sortie propre, bloc de traitement — constitue le patron que l'on retrouvera, identique, dans toute application Access bien conçue. C'est ce patron, et la logique qui le sous-tend, que cette section détaille.

## Le principe de On Error GoTo

L'instruction `On Error GoTo <étiquette>` indique à VBA que, dès qu'une erreur d'exécution survient, l'exécution doit immédiatement sauter vers l'étiquette désignée, au lieu d'interrompre le programme. Cette étiquette — le **gestionnaire d'erreurs** — est un repère placé dans la même procédure, conventionnellement à la fin, et suivi du code chargé de traiter l'incident.

Voici la forme la plus dépouillée du mécanisme :

```vba
Sub Exemple()
    On Error GoTo GestionErreur

    ' --- Corps de la procédure ---
    Dim resultat As Integer
    resultat = 10 / 0          ' provoque l'erreur 11 (division par zéro)

    Exit Sub                    ' indispensable : voir plus bas

GestionErreur:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description
End Sub
```

Lorsque la division par zéro se produit, VBA ne montre pas la boîte de dialogue système : il saute directement à l'étiquette `GestionErreur:` et exécute le `MsgBox`, qui affiche un message contrôlé. L'instruction `On Error GoTo` reste active à partir du moment où elle est exécutée, et couvre donc l'ensemble des instructions situées en aval, dans la même procédure.

Le nom de l'étiquette est libre. Par convention, on emploie un nom explicite et cohérent dans toute l'application : `GestionErreur`, `Gestion_Erreur`, `ErrHandler` ou encore `TraiterErreur`. La cohérence du nommage facilite la lecture et le copier-coller du patron d'une procédure à l'autre.

## La structure canonique d'une procédure protégée

La forme minimale ci-dessus illustre le mécanisme, mais elle est incomplète : il lui manque un emplacement dédié au **nettoyage** des ressources. La structure réellement utilisée en production comporte trois zones distinctes, séparées par deux étiquettes :

```vba
Public Sub MettreAJourTarifs()
    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    On Error GoTo GestionErreur

    ' ===== 1. CORPS DE LA PROCÉDURE =====
    DoCmd.Hourglass True
    Set db = CurrentDb
    Set rst = db.OpenRecordset("T_Articles", dbOpenDynaset)

    Do While Not rst.EOF
        rst.Edit
        rst!PrixTTC = rst!PrixHT * 1.2
        rst.Update
        rst.MoveNext
    Loop

    MsgBox "Mise à jour terminée.", vbInformation

SortieProcedure:
    ' ===== 2. NETTOYAGE (toujours exécuté) =====
    On Error Resume Next            ' le nettoyage ne doit jamais lever d'erreur
    If Not rst Is Nothing Then
        rst.Close
        Set rst = Nothing
    End If
    Set db = Nothing
    DoCmd.Hourglass False
    Exit Sub                        ' empêche de « tomber » dans le gestionnaire

GestionErreur:
    ' ===== 3. TRAITEMENT DE L'ERREUR =====
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, _
           vbCritical, "Mise à jour des tarifs"
    Resume SortieProcedure
End Sub
```

Cette structure se lit en trois temps. Le **corps** contient le traitement métier, protégé par le `On Error GoTo` placé en tête. Le bloc de **nettoyage**, repéré par l'étiquette `SortieProcedure:`, libère les ressources (ici le recordset et le sablier) ; il est conçu pour être exécuté dans tous les cas, qu'une erreur soit survenue ou non. Le bloc de **traitement de l'erreur**, repéré par `GestionErreur:`, informe l'utilisateur, puis renvoie l'exécution vers le nettoyage via `Resume SortieProcedure`.

Le déroulement normal — sans erreur — suit ce chemin : corps complet, puis `SortieProcedure` (nettoyage), puis `Exit Sub`. Le déroulement en cas d'erreur saute du point fautif vers `GestionErreur`, affiche le message, puis rejoint `SortieProcedure` grâce au `Resume`. Dans les deux cas, **le nettoyage est exécuté** : c'est tout l'intérêt de cette organisation.

## Pourquoi Exit Sub est indispensable

L'instruction `Exit Sub` placée juste avant l'étiquette du gestionnaire est l'élément le plus souvent oublié par les débutants, et son absence provoque un comportement déroutant.

Une étiquette n'est qu'un repère : elle ne crée aucune barrière dans le flux d'exécution. Si l'on omet `Exit Sub`, une procédure qui s'est déroulée **sans aucune erreur** poursuivra naturellement son chemin jusqu'à atteindre l'étiquette `GestionErreur:`, puis exécutera le code de traitement d'erreur — alors même qu'aucune erreur ne s'est produite. L'utilisateur verra ainsi s'afficher un message d'erreur trompeur (généralement « Erreur 0 », l'objet `Err` étant vide) à la fin de chaque opération réussie.

`Exit Sub` (ou `Exit Function` dans une fonction, `Exit Property` dans une propriété) ferme le chemin normal et garantit que le gestionnaire d'erreurs n'est atteint **que** par un saut consécutif à une erreur. C'est une instruction non négociable de la structure.

## Le bloc de nettoyage : l'équivalent de « finally »

Les langages dotés de blocs structurés proposent une clause `finally`, exécutée systématiquement à la sortie d'un bloc protégé, qu'une erreur soit survenue ou non. VBA ne possède pas cette clause, mais l'**étiquette de nettoyage** vers laquelle convergent le flux normal et le flux d'erreur en remplit exactement le rôle.

Ce point de convergence est ce qui distingue une gestion d'erreurs robuste d'une gestion approximative. Si le code de libération des objets (`rst.Close`, `Set ... = Nothing`) était placé uniquement à la fin du corps de la procédure, il ne serait **jamais exécuté en cas d'erreur**, puisque l'erreur fait sauter directement vers le gestionnaire. Les objets resteraient ouverts en mémoire, les verrous non relâchés, le sablier figé à l'écran — autant de fuites de ressources qui, accumulées, dégradent puis déstabilisent l'application.

En faisant converger les deux chemins vers `SortieProcedure`, on garantit que le nettoyage a lieu dans tous les cas. On notera, en tête de ce bloc, l'instruction `On Error Resume Next` : elle protège le nettoyage lui-même contre toute erreur (par exemple si l'objet est déjà fermé), afin que la libération des ressources ne puisse jamais échouer à son tour.

## Les instructions de reprise : la famille Resume

Une fois l'erreur traitée dans le gestionnaire, il faut décider **où** reprendre l'exécution. C'est le rôle des instructions `Resume`, qui n'existent que dans le contexte d'un gestionnaire d'erreurs actif. Il en existe trois variantes, aux comportements très différents.

| Instruction | Reprend l'exécution… | Usage typique |
|---|---|---|
| `Resume` | sur la **même** instruction qui a échoué | Après avoir corrigé la condition à l'origine de l'erreur (dossier créé, ressource rendue disponible) |
| `Resume Next` | sur l'instruction **suivant** celle qui a échoué | Pour ignorer une instruction non critique et continuer |
| `Resume <étiquette>` | au niveau de l'**étiquette** indiquée | Pour converger vers le bloc de nettoyage et sortir proprement (cas le plus fréquent) |

`Resume` seul réessaie l'instruction fautive. Il n'a de sens que si le gestionnaire a effectivement supprimé la cause de l'erreur entre-temps. Dans le cas contraire, l'instruction échouera de nouveau, sera de nouveau déroutée vers le gestionnaire, qui réessaiera… créant une **boucle infinie**. C'est un piège classique, à n'utiliser qu'avec une condition de sortie clairement maîtrisée.

`Resume Next` reprend après l'instruction fautive, ce qui revient à passer outre l'erreur sur cette seule instruction. Il est utile lorsqu'une opération est facultative et que son échec ne doit pas interrompre le traitement.

`Resume <étiquette>` est la forme employée dans le patron canonique : elle redirige vers l'étiquette de nettoyage, assurant une sortie propre et unifiée. C'est, dans la grande majorité des procédures, le choix par défaut.

Un détail important : `Resume` (sous toutes ses formes) **réarme** le gestionnaire d'erreurs et réinitialise l'objet `Err`. Après un `Resume`, une nouvelle erreur sera de nouveau interceptée normalement par le gestionnaire.

## Réagir différemment selon l'erreur

Un gestionnaire n'est pas tenu de traiter toutes les erreurs de la même façon. En examinant `Err.Number`, on peut adapter la réponse à la nature précise de l'incident — message convivial pour les erreurs métier prévisibles, message générique pour les erreurs inattendues. Une structure `Select Case` est l'outil naturel pour cela :

```vba
GestionErreur:
    Select Case Err.Number
        Case 3022                          ' clé primaire en double
            MsgBox "Cet enregistrement existe déjà.", vbExclamation
        Case 3201                          ' enregistrement lié requis
            MsgBox "Sélectionnez d'abord un client valide.", vbExclamation
        Case 2501                          ' action annulée : sans gravité
            ' on ignore silencieusement
        Case Else                          ' toute autre erreur
            MsgBox "Erreur inattendue " & Err.Number & " : " & _
                   Err.Description, vbCritical
    End Select
    Resume SortieProcedure
```

Ce schéma traduit une bonne pratique : distinguer les erreurs **anticipées** (que l'on sait pouvoir se produire et que l'on traite par un message clair) des erreurs **imprévues** (couvertes par le `Case Else`, signalées comme telles). On retrouve ici l'intérêt du catalogue de la section 13.1 : connaître les numéros permet de cibler précisément les erreurs métier.

## Désactiver la gestion d'erreurs : On Error GoTo 0

`On Error GoTo 0` désactive le gestionnaire d'erreurs actuellement actif dans la procédure et rétablit le comportement par défaut. À partir de cette instruction, une erreur n'est plus déroutée : elle provoque de nouveau l'arrêt et l'affichage de la boîte de dialogue système.

Son usage le plus courant est de restreindre volontairement une zone « sans filet », par exemple lors du débogage, ou pour entourer une portion où l'on souhaite que toute erreur remonte immédiatement. Il faut toutefois l'employer en connaissance de cause, car il rend de nouveau visibles les erreurs brutes.

On rencontre parfois une instruction voisine, `On Error GoTo -1`, à ne pas confondre avec la précédente. Là où `On Error GoTo 0` désactive le gestionnaire, `On Error GoTo -1` **réinitialise l'état d'erreur courant** : il signale à VBA que l'erreur en cours est considérée comme « traitée », ce qui autorise l'activation d'un nouveau gestionnaire au sein du même bloc de traitement, sans passer par `Resume`. C'est une nuance avancée, rarement nécessaire ; dans la pratique courante, `Resume` suffit à sortir de l'état d'erreur, et l'on s'en tiendra à `On Error GoTo 0` pour la simple désactivation.

## Portée d'un gestionnaire d'erreurs

Un gestionnaire d'erreurs ne couvre que la procédure dans laquelle il est défini. L'instruction `On Error GoTo` reste active depuis son exécution jusqu'à la fin de la procédure, ou jusqu'à ce qu'une autre instruction `On Error` la remplace.

Une question naturelle se pose alors : que se passe-t-il si une erreur survient dans une procédure **appelée**, qui ne dispose pas de son propre gestionnaire ? Dans ce cas, l'erreur n'est pas perdue : elle « remonte » vers la procédure appelante et y est interceptée par le gestionnaire actif. Ce mécanisme de remontée — la propagation des erreurs le long de la pile des appels — obéit à des règles précises qui font l'objet de la section [13.9](/13-gestion-erreurs/09-propagation-erreurs.md). Pour l'instant, il suffit de retenir qu'une erreur non gérée localement ne disparaît pas : elle cherche un gestionnaire en remontant la chaîne des appels.

## Localiser l'erreur : numérotation de lignes et fonction Erl

Le message « Erreur 11 : Division par zéro » indique la nature de l'erreur, mais pas **l'endroit** où elle s'est produite. Dans une procédure longue, retrouver l'instruction fautive peut s'avérer fastidieux. VBA offre un mécanisme hérité d'anciennes versions de Basic permettant d'y remédier : la **numérotation des lignes**, associée à la fonction `Erl` (*Error Line*), qui renvoie le numéro de la ligne où l'erreur est survenue.

```vba
Sub ExempleNumerote()
10  On Error GoTo GestionErreur
20  Dim x As Integer
30  x = 1 / 0
40  Exit Sub
GestionErreur:
50  MsgBox "Erreur " & Err.Number & " à la ligne " & Erl
End Sub
```

Ici, le message indiquera que l'erreur est survenue à la ligne 30. La numérotation manuelle est rarement pratiquée à la main — elle alourdit le code et doit être maintenue — mais elle prend tout son sens en production lorsqu'elle est ajoutée **automatiquement** par un outil et combinée à une journalisation : on dispose alors, pour chaque incident remonté du terrain, du numéro d'erreur, de la procédure et de la ligne exacte. Cette articulation entre `Erl` et la journalisation est approfondie dans les sections [13.6](/13-gestion-erreurs/06-journalisation-erreurs-table.md) et [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md).

## Pièges courants à éviter

Quelques erreurs récurrentes compromettent la robustesse de la structure :

- **Oublier `Exit Sub`** avant le gestionnaire : la procédure exécute le code de traitement d'erreur même lorsqu'elle réussit, affichant un faux message d'erreur.
- **Utiliser `Resume` au lieu de `Resume Next` ou `Resume <étiquette>`** alors que la cause n'a pas été corrigée : l'instruction fautive est réessayée indéfiniment, créant une boucle infinie.
- **Placer le nettoyage uniquement dans le flux normal**, sans étiquette de convergence : les ressources ne sont pas libérées en cas d'erreur, ce qui provoque des fuites.
- **Croire que le gestionnaire se protège lui-même** : une erreur survenant *à l'intérieur* du bloc de traitement n'est pas interceptée par ce même bloc ; elle se propage. D'où l'intérêt du `On Error Resume Next` en tête du nettoyage.
- **Laisser un `On Error Resume Next` actif** bien au-delà de la zone concernée, masquant ainsi des erreurs réelles plus loin dans la procédure (sujet développé dans la section suivante).

## En résumé

`On Error GoTo <étiquette>` détourne le comportement par défaut d'Access vers un traitement écrit par le développeur. Mais c'est la **structure** qui fait la robustesse : un corps de procédure protégé, un bloc de nettoyage vers lequel convergent le flux normal et le flux d'erreur, et un bloc de traitement qui informe puis renvoie vers ce nettoyage. L'instruction `Exit Sub` ferme le chemin normal pour que le gestionnaire ne soit atteint que sur erreur ; l'étiquette de nettoyage joue le rôle du `finally` absent du langage.

Les instructions `Resume`, `Resume Next` et `Resume <étiquette>` déterminent le point de reprise — la dernière étant la plus courante dans le patron standard. L'examen de `Err.Number` dans le gestionnaire permet d'adapter la réponse à chaque type d'erreur, en distinguant les incidents anticipés des incidents imprévus. Enfin, la portée d'un gestionnaire se limite à sa procédure, les erreurs non gérées remontant la pile des appels.

Ce patron constitue la base sur laquelle s'appuie tout le reste du chapitre. La section suivante examine l'autre grande instruction de gestion d'erreurs, `On Error Resume Next`, et les précautions indispensables à son emploi.

---


⏭️ [13.3. On Error Resume Next — utilisation contrôlée](/13-gestion-erreurs/03-on-error-resume-next.md)
