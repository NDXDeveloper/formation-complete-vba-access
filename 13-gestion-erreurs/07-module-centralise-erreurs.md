🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.7. Module de gestion centralisée des erreurs

## Introduction

Les sections précédentes ont fourni toutes les pièces d'une gestion d'erreurs robuste : la structure `On Error GoTo` (13.2), l'examen de l'objet `Err` (13.4), l'interception des erreurs du moteur de données (13.5) et la journalisation en table (13.6). Reste à les assembler. Car appliquer ces techniques **procédure par procédure**, en réécrivant à chaque fois le même bloc de traitement, conduit vite à une situation intenable : du code dupliqué partout, des comportements légèrement différents d'un endroit à l'autre, et l'impossibilité de modifier la politique de gestion d'erreurs sans rouvrir des centaines de procédures.

La solution consiste à **centraliser** : concentrer toute la logique de traitement — journalisation, message à l'utilisateur, décision de reprise — dans un module unique, exposant une fonction appelée de façon uniforme depuis chaque gestionnaire. Cette section construit ce module, présente le gabarit standard que tout gestionnaire adoptera, et introduit une distinction essentielle entre comportement en développement et comportement en production. C'est l'aboutissement pratique du chapitre : à l'issue de cette section, on disposera d'une infrastructure de gestion d'erreurs complète et réutilisable.

## Le problème de la gestion dispersée

Sans centralisation, chaque gestionnaire d'erreurs ressemble à ses voisins, sans leur être identique :

```vba
' Dans une procédure...
GestionErreur:
    JournaliserErreur Err.Number, Err.Description, Err.Source, "ProcA"
    MsgBox "Erreur : " & Err.Description, vbCritical
    Resume SortieProcedure

' Dans une autre, légèrement différente...
GestionErreur:
    MsgBox "Une erreur est survenue (" & Err.Number & ")"
    ' ... oubli de journaliser ...
    Resume SortieProcedure
```

Cette dispersion engendre trois maux. Le **code est dupliqué** : la même logique se répète à l'identique dans toute l'application. Les **comportements divergent** : ici on journalise, là on oublie ; ici le message est soigné, là il est brut. Et surtout, **toute évolution devient coûteuse** : changer le format des messages, ajouter un niveau de gravité, modifier la destination du journal imposerait de retoucher chaque gestionnaire, un par un. La centralisation supprime ces trois problèmes d'un coup.

## Le principe de la centralisation

L'idée directrice est simple : il existe **un seul endroit** — un module standard dédié, par exemple `modGestionErreurs` — où est définie la manière de traiter une erreur. Chaque gestionnaire de l'application n'y fait plus qu'**appeler une fonction**, en lui transmettant les caractéristiques de l'erreur. Cette fonction décide alors de tout : ce qu'il faut journaliser, ce qu'il faut afficher, et comment la procédure appelante doit poursuivre.

Le gestionnaire local se réduit ainsi à quelques lignes invariables, tandis que l'intelligence du traitement réside, unique et modifiable d'un seul tenant, dans le module central. Mais cette architecture se heurte d'emblée à une contrainte du langage qu'il faut comprendre avant d'écrire la moindre ligne.

## Une contrainte structurante : Resume reste dans l'appelant

En VBA, l'instruction `Resume` — sous toutes ses formes — ne peut être exécutée que **dans la procédure où l'erreur est survenue** et où le gestionnaire est actif. Il est impossible de placer un `Resume` dans une fonction appelée pour qu'il fasse reprendre l'exécution de la procédure appelante. La reprise est indissociable du contexte d'erreur local.

Cette contrainte façonne entièrement la conception du module central. Puisque la fonction centrale ne peut pas exécuter le `Resume` elle-même, elle ne peut que **décider** de ce qu'il convient de faire, puis **retourner cette décision** à l'appelant, à charge pour ce dernier d'exécuter le `Resume` correspondant. La fonction centrale renvoie donc une valeur indiquant la marche à suivre — réessayer, ignorer, ou sortir proprement — et le gestionnaire appelant traduit cette valeur en l'instruction `Resume` adéquate. Ce mécanisme de « décision retournée » est la clé de voûte de toute l'architecture.

## L'énumération des actions de reprise

Pour exprimer clairement les décisions possibles, on définit une énumération. Elle rend le code lisible et évite les valeurs numériques opaques :

```vba
' Indique au gestionnaire appelant comment poursuivre l'exécution
Public Enum ActionErreur
    actReprendre = 0          ' Resume      : réessayer l'instruction fautive
    actReprendreSuite = 1     ' Resume Next : ignorer et continuer
    actQuitter = 2            ' Resume <étiquette de nettoyage>
End Enum
```

Ces trois valeurs couvrent les issues vues à la section [13.2](/13-gestion-erreurs/02-on-error-goto.md) : `actReprendre` correspond à `Resume`, `actReprendreSuite` à `Resume Next`, et `actQuitter` à une sortie propre via l'étiquette de nettoyage. La fonction centrale renverra l'une de ces valeurs, et le gestionnaire appelant la traduira en instruction concrète.

## La fonction centrale de gestion d'erreurs

Voici le cœur du module : une fonction qui journalise systématiquement, informe l'utilisateur de manière adaptée à la nature de l'erreur, et retourne l'action de reprise.

```vba
' =====================================================================
'  Module standard : modGestionErreurs
' =====================================================================
Option Compare Database
Option Explicit

' Mode développement : True en atelier, False avant déploiement
Private Const MODE_DEBUG As Boolean = True

Public Function GererErreur(ByVal numero As Long, _
                            ByVal description As String, _
                            ByVal source As String, _
                            ByVal procedure As String, _
                            Optional ByVal ligne As Long = 0, _
                            Optional ByVal contexte As String = "") _
                            As ActionErreur

    ' 1. Journalisation systématique (procédure de la section 13.6)
    JournaliserErreur numero, description, source, procedure, ligne, contexte

    ' 2. Message à l'utilisateur et décision de reprise
    Select Case numero

        Case 2501                       ' action annulée : sans gravité
            GererErreur = actQuitter

        Case 3022                       ' clé en double
            MsgBox "Cet enregistrement existe déjà.", _
                   vbExclamation, "Doublon"
            GererErreur = actQuitter

        Case 3201                       ' enregistrement lié requis
            MsgBox "Veuillez d'abord sélectionner une valeur valide.", _
                   vbExclamation, "Donnée manquante"
            GererErreur = actQuitter

        Case Else                       ' erreur imprévue
            If MODE_DEBUG Then
                MsgBox "Erreur " & numero & " — " & description & vbCrLf & _
                       "Source : " & source & vbCrLf & _
                       "Procédure : " & procedure & _
                       IIf(ligne > 0, " (ligne " & ligne & ")", ""), _
                       vbCritical, "Erreur [mode développement]"
            Else
                MsgBox "Une erreur inattendue s'est produite." & vbCrLf & _
                       "L'incident a été enregistré." & vbCrLf & _
                       "Référence : " & numero, _
                       vbCritical, "Erreur"
            End If
            GererErreur = actQuitter

    End Select

End Function
```

Cette fonction concentre toute la politique de gestion d'erreurs de l'application. La journalisation y est **systématique** : quelle que soit l'erreur, une trace est conservée. Le traitement est ensuite **différencié** par un `Select Case` sur le numéro, reprenant la distinction établie à la section [13.4](/13-gestion-erreurs/04-objet-err.md) entre erreurs anticipées (messages spécifiques et conviviaux) et erreurs imprévues (regroupées dans le `Case Else`). Modifier le comportement global de l'application — ajouter un message, changer un libellé, traiter un nouveau numéro — ne demande désormais d'intervenir qu'**ici**.

## Le gabarit standard de gestionnaire appelant

Du côté des procédures de l'application, le gestionnaire se réduit à un gabarit invariable, copié à l'identique partout :

```vba
Public Sub TraiterCommande()
    On Error GoTo GestionErreur

    ' ===== corps de la procédure =====
    ' ...

SortieProcedure:
    ' ===== nettoyage =====
    ' ...
    Exit Sub

GestionErreur:
    Select Case GererErreur(Err.Number, Err.Description, Err.Source, _
                            "TraiterCommande", Erl)
        Case actReprendre:      Resume
        Case actReprendreSuite: Resume Next
        Case Else:              Resume SortieProcedure
    End Select
End Sub
```

Ce gabarit appelle `GererErreur` en lui passant les caractéristiques de l'erreur, puis traduit la valeur retournée en l'instruction `Resume` correspondante. On retrouve ici la contrainte évoquée plus haut : les trois `Resume` figurent bien dans la procédure appelante, et non dans `GererErreur`.

Un point mérite d'être souligné concernant le passage des arguments. On transmet `Err.Number`, `Err.Description` et `Err.Source` **directement** à la fonction, sans les capturer au préalable dans des variables locales. Pourquoi est-ce sûr, alors que la section [13.4](/13-gestion-erreurs/04-objet-err.md) insistait sur le risque de réinitialisation de `Err` ? Parce que ces arguments sont **évalués au point d'appel**, c'est-à-dire avant que l'exécution n'entre dans `GererErreur`. Au moment où la fonction commence à s'exécuter — et où son propre code, ou la journalisation qu'elle déclenche, réinitialise `Err` —, les valeurs ont déjà été lues et transmises par copie. La fonction `Erl` est évaluée de la même manière, au point d'appel. Ce passage direct est donc parfaitement fiable, et même plus élégant que la capture en variables locales.

## Le mode développement et le mode production

L'un des grands avantages de la centralisation est de pouvoir faire diverger le comportement de l'application selon le contexte, par le seul jeu d'un indicateur — ici la constante `MODE_DEBUG`.

En **développement**, on souhaite voir l'erreur dans toute sa crudité : numéro, description, source, procédure, ligne. Ces informations techniques, affichées intégralement, permettent de localiser et de corriger le défaut sans détour. En **production**, en revanche, ces détails sont au mieux inutiles, au pire inquiétants pour l'utilisateur. On lui présente alors un message sobre et rassurant, en lui signalant que l'incident a été enregistré, tout en conservant le diagnostic complet dans le journal à l'intention du développeur.

Basculer de l'un à l'autre se résume à modifier une seule constante avant le déploiement. On peut raffiner ce mécanisme : lire l'indicateur depuis une table de paramètres plutôt que depuis une constante (afin de l'activer ponctuellement sur un poste sans recompiler), ou recourir à la **compilation conditionnelle** (`#Const` et `#If`) pour que le code de débogage — par exemple une instruction `Stop` permettant d'interrompre l'exécution en atelier — soit purement et simplement absent de la version livrée. Quelle qu'en soit la forme, le principe reste le même : un comportement adaptable depuis un point unique.

## Centraliser les messages destinés à l'utilisateur

Au-delà de la mécanique, la centralisation présente un bénéfice souvent sous-estimé : elle unifie le **discours** de l'application envers l'utilisateur. Lorsque chaque procédure formulait ses propres messages, une même erreur pouvait être présentée de dix façons différentes, avec des tons et des libellés incohérents. En regroupant dans `GererErreur` la correspondance entre numéros d'erreur et messages conviviaux, on garantit que l'application « parle d'une seule voix » : le doublon de clé donne toujours le même message, où qu'il survienne. Cette cohérence renforce le sentiment de qualité et de fiabilité du logiciel, et facilite par ailleurs une éventuelle traduction, tous les libellés étant rassemblés au même endroit.

## Module standard ou classe ?

La gestion d'erreurs est ici implémentée dans un **module standard**, ce qui constitue le choix le plus courant et le plus simple : la fonction `GererErreur` est un utilitaire sans état, accessible globalement, qui n'a pas besoin d'être instancié. C'est la solution recommandée dans la plupart des applications.

Une approche orientée objet est néanmoins possible : on pourrait encapsuler la gestion d'erreurs dans un **module de classe**, ce qui ouvrirait la voie à des fonctionnalités plus riches — accumulation d'un contexte, niveaux de gravité, plusieurs canaux de sortie configurables. Ces possibilités relèvent de la programmation orientée objet en VBA, traitée au chapitre 16, et dépassent le cadre de ce chapitre. Pour une gestion d'erreurs claire et efficace, le module standard suffit amplement.

## Articulation avec la propagation des erreurs

La fonction centrale décide ici, dans tous les cas, de traiter l'erreur localement (journalisation, message, sortie). Mais une autre stratégie existe : laisser l'erreur **remonter** vers une couche supérieure mieux à même de la traiter, plutôt que de la consommer sur place. Dans une telle conception, le module central pourrait, pour certaines erreurs, choisir de les **propager** au lieu de les afficher — par exemple en signalant à l'appelant qu'il doit lui-même relancer l'erreur. Cette dimension, qui touche à la répartition des responsabilités le long de la pile des appels, fait l'objet de la section [13.9](/13-gestion-erreurs/09-propagation-erreurs.md). Le module présenté ici constitue une base que l'on pourra enrichir de cette logique de propagation une fois ces principes acquis.

## Bénéfices de la centralisation

L'investissement dans un module centralisé se rentabilise immédiatement et durablement :

- **Cohérence** : toutes les erreurs sont traitées selon une même politique, et l'application présente un comportement uniforme.
- **Non-duplication** : la logique de traitement n'existe qu'en un seul exemplaire ; les gestionnaires locaux se réduisent à un gabarit minimal.
- **Maintenabilité** : toute évolution de la politique d'erreurs se fait en un point unique, sans toucher aux procédures.
- **Journalisation systématique** : aucune erreur ne risque plus de passer sous le radar par oubli de journalisation.
- **Adaptabilité** : le basculement entre développement et production, comme la formulation des messages, se pilote depuis le module central.

## En résumé

La gestion centralisée concentre toute la politique de traitement des erreurs dans un module unique, exposant une fonction — `GererErreur` — appelée de façon uniforme par chaque gestionnaire. Cette architecture est dictée par une contrainte fondamentale du langage : `Resume` ne pouvant s'exécuter que dans la procédure d'origine, la fonction centrale ne peut que **décider** de la reprise et **retourner** cette décision, l'appelant exécutant ensuite le `Resume` correspondant. Une énumération (`actReprendre`, `actReprendreSuite`, `actQuitter`) exprime clairement ces décisions.

Le gestionnaire appelant se réduit alors à un gabarit invariable, qui transmet directement `Err.Number`, `Err.Description`, `Err.Source` et `Erl` — un passage sûr, car ces valeurs sont évaluées au point d'appel, avant toute réinitialisation. La fonction centrale journalise systématiquement, différencie les messages selon le numéro d'erreur, et adapte son comportement au contexte grâce à un indicateur de mode développement/production. Au-delà de la mécanique, elle unifie le discours de l'application et concentre en un seul endroit toute évolution future. Le module standard suffit dans la grande majorité des cas, une approche orientée objet restant possible pour des besoins avancés.

Le chapitre a jusqu'ici montré comment intercepter, examiner, journaliser et centraliser le traitement des erreurs *du système*. La section suivante aborde un versant complémentaire : la création d'erreurs *propres à l'application*, déclenchées volontairement pour signaler les violations des règles métier.

---


⏭️ [13.8. Création d'erreurs personnalisées avec Err.Raise](/13-gestion-erreurs/08-erreurs-personnalisees-raise.md)
