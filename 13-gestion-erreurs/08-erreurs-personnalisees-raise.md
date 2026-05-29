🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.8. Création d'erreurs personnalisées avec Err.Raise

## Introduction

Jusqu'à présent, le chapitre a traité des erreurs *subies* : celles que VBA, Access ou le moteur de données déclenchent face à une situation impossible. Mais il est souvent utile de déclencher *soi-même* une erreur, délibérément, pour signaler qu'une **règle métier** a été violée. Un client inactif à qui l'on tente de passer commande, une quantité négative, une opération dans une période comptable close : ces situations ne provoquent aucune erreur technique — le code s'exécuterait sans broncher — et pourtant elles doivent être traitées comme des anomalies.

La méthode `Err.Raise` permet précisément cela : générer une erreur d'exécution exactement comme si VBA l'avait rencontrée. L'intérêt est considérable, car l'erreur ainsi créée emprunte **toute l'infrastructure déjà en place** — interception par `On Error GoTo`, examen de l'objet `Err`, journalisation, traitement par le module centralisé. Les erreurs métier et les erreurs système se gèrent alors de façon parfaitement uniforme. Cette section explique comment déclencher ces erreurs, comment leur attribuer des numéros qui ne risquent pas d'entrer en conflit avec ceux du système, et — point essentiel — quand il convient d'y recourir plutôt que d'employer une simple valeur de retour.

## Pourquoi créer ses propres erreurs

Recourir à `Err.Raise` pour signaler une violation de règle métier procure deux avantages décisifs.

D'abord, l'**uniformité du traitement**. Une erreur métier levée par `Err.Raise` est interceptée par le même `On Error GoTo` que n'importe quelle erreur système, transite par le même module centralisé, et bénéficie de la même journalisation. Il n'est pas nécessaire d'inventer un mécanisme parallèle pour les anomalies métier : elles rejoignent le flux existant.

Ensuite, la **propagation depuis la profondeur du code**. Lorsqu'une règle est vérifiée au fond d'une chaîne d'appels — dans une fonction de validation appelée par une autre, elle-même appelée par un gestionnaire d'événement —, signaler l'anomalie par une valeur de retour obligerait à la faire remonter manuellement à travers chaque niveau. Une erreur levée, au contraire, **remonte d'elle-même** jusqu'au premier gestionnaire actif, sans qu'aucune couche intermédiaire n'ait à s'en préoccuper. C'est un moyen puissant de faire émerger une anomalie là où elle pourra être traitée, depuis l'endroit où elle est détectée.

## La méthode Err.Raise

`Err.Raise` déclenche une erreur d'exécution. Sa signature comporte un argument obligatoire et plusieurs facultatifs :

```vba
Err.Raise Number, [Source], [Description], [HelpFile], [HelpContext]
```

Seul `Number` est requis : c'est le numéro de l'erreur. `Source` identifie le composant à l'origine de l'erreur ; `Description` fournit le libellé qui sera consultable via `Err.Description` ; `HelpFile` et `HelpContext`, rarement employés, associent une rubrique d'aide. En pratique, on renseigne presque toujours `Number`, `Source` et `Description`, qui sont les trois informations exploitées par les gestionnaires.

```vba
Err.Raise vbObjectError + 1000, "modClient", "Le client est inactif."
```

Une fois cette instruction exécutée, tout se passe comme si une erreur système s'était produite : l'exécution est déroutée vers le gestionnaire actif, et l'objet `Err` reflète les valeurs fournies.

## Choisir un numéro d'erreur : vbObjectError

Le choix du numéro n'est pas anodin, et c'est ici qu'intervient la constante `vbObjectError`.

Le problème est celui de la **collision**. Si l'on attribuait à une erreur métier le numéro 5 ou 3022, elle deviendrait indiscernable de l'erreur système portant le même numéro. Un gestionnaire testant `Err.Number = 3022` ne pourrait plus savoir s'il s'agit d'un véritable doublon de clé ou d'une erreur métier mal numérotée. Il faut donc choisir les numéros dans une plage **réservée aux erreurs définies par l'utilisateur**, garantie sans conflit avec les codes du système.

C'est le rôle de `vbObjectError`, une constante (de valeur -2147221504) qui sert de **base** aux numéros personnalisés. On ne l'utilise jamais seule : on lui ajoute un décalage propre à l'application. Les valeurs `vbObjectError + 0` à `vbObjectError + 511` étant elles-mêmes réservées, les numéros personnalisés doivent commencer à `vbObjectError + 512` ; par convention, on choisit souvent un point de départ confortablement au-delà, par exemple `vbObjectError + 1000`, pour disposer d'une plage claire et entièrement dédiée.

Un effet de bord mérite d'être anticipé : `vbObjectError + 1000` est un grand nombre **négatif**, peu lisible et peu commode à manipuler tel quel. La parade, exposée ci-dessous, consiste à ne jamais écrire ces valeurs brutes dans le code, mais à les masquer derrière des **constantes nommées**.

## Définir un catalogue de constantes d'erreur

Plutôt que de répéter `vbObjectError + 1000` à droite et à gauche, on rassemble les erreurs métier de l'application dans un module dédié, sous forme de constantes nommées et explicites :

```vba
' =====================================================================
'  Module standard : modErreursMetier
'  Catalogue des erreurs propres à l'application
' =====================================================================
Option Explicit

Public Const ERR_CLIENT_INACTIF      As Long = vbObjectError + 1000
Public Const ERR_SOLDE_INSUFFISANT   As Long = vbObjectError + 1001
Public Const ERR_QUANTITE_INVALIDE   As Long = vbObjectError + 1002
Public Const ERR_PERIODE_CLOTUREE    As Long = vbObjectError + 1003
```

Ce catalogue présente un double bénéfice. Le code devient **lisible** : on lève `ERR_CLIENT_INACTIF` et l'on teste `Case ERR_CLIENT_INACTIF`, sans jamais voir le nombre négatif sous-jacent. Et les numéros sont **centralisés** : on dispose d'un inventaire unique des anomalies métier de l'application, facile à consulter et à maintenir, dans le même esprit que la centralisation de la section [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md). Cette approche par constantes résout du même coup le problème de lisibilité des valeurs `vbObjectError`.

## Renseigner Source et Description

Au-delà du numéro, les arguments `Source` et `Description` enrichissent considérablement le diagnostic.

Par convention, on renseigne `Source` avec une référence précise à l'origine de l'erreur, typiquement sous la forme `module.procédure` : `"modCommande.ValiderCommande"`. Cette indication, conservée dans le journal, permet de localiser immédiatement le point de levée. Quant à `Description`, elle doit formuler l'anomalie en termes **compréhensibles**, idéalement en incluant les éléments concrets utiles à l'utilisateur ou au support : l'identifiant en cause, la valeur fautive, la contrainte non respectée. Un bon message métier dit non seulement *ce qui* ne va pas, mais *pourquoi*, et si possible *quoi faire*.

## Un exemple complet : valider une règle métier

Voici l'usage type. Une procédure de validation vérifie les règles et lève une erreur dès qu'une condition n'est pas remplie :

```vba
Public Sub ValiderCommande(ByVal idClient As Long, ByVal quantite As Long)

    If quantite <= 0 Then
        Err.Raise ERR_QUANTITE_INVALIDE, "modCommande.ValiderCommande", _
                  "La quantité doit être strictement positive."
    End If

    If Not ClientEstActif(idClient) Then
        Err.Raise ERR_CLIENT_INACTIF, "modCommande.ValiderCommande", _
                  "Le client " & idClient & " est inactif ; commande impossible."
    End If

End Sub
```

Cette procédure de validation ne comporte **aucun gestionnaire local** : elle se contente de lever l'erreur. C'est la procédure appelante, dotée d'un `On Error GoTo`, qui l'interceptera :

```vba
Public Sub EnregistrerCommande()
    On Error GoTo GestionErreur

    ValiderCommande mIdClient, mQuantite     ' peut lever une erreur métier
    ' ... enregistrement effectif de la commande ...

SortieProcedure:
    Exit Sub

GestionErreur:
    Select Case GererErreur(Err.Number, Err.Description, Err.Source, _
                            "EnregistrerCommande", Erl)
        Case actReprendre:      Resume
        Case actReprendreSuite: Resume Next
        Case Else:              Resume SortieProcedure
    End Select
End Sub
```

L'erreur levée dans `ValiderCommande` n'y trouvant pas de gestionnaire, elle remonte naturellement jusqu'à celui de `EnregistrerCommande`, où elle est traitée comme n'importe quelle autre erreur. La validation et la gestion de l'anomalie sont ainsi nettement séparées, chacune dans sa procédure.

## Où atterrit une erreur levée ?

Un point demande clarification, car il détermine le comportement réel de `Err.Raise`. Une erreur levée se comporte exactement comme une erreur système : elle est dirigée vers le **gestionnaire actif de la procédure courante**, s'il en existe un.

La conséquence est importante. Si l'on appelle `Err.Raise` à l'intérieur d'une procédure qui possède son propre `On Error GoTo`, l'erreur sera captée par ce gestionnaire local, et **ne remontera pas** vers l'appelant. Pour qu'une erreur métier se propage jusqu'à un niveau supérieur, on la lève donc dans une procédure dépourvue de gestionnaire local — comme `ValiderCommande` ci-dessus — ou bien on la **relance** explicitement depuis le gestionnaire local. C'est une décision de conception : selon que l'on veut traiter l'anomalie sur place ou la faire remonter, on placera ou non un gestionnaire dans la procédure qui lève l'erreur.

## Intégration avec le module centralisé

Le module centralisé de la section [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md) accueille sans difficulté les erreurs métier : il suffit d'ajouter, dans son `Select Case`, des cas correspondant aux constantes du catalogue.

```vba
Select Case numero

    Case ERR_CLIENT_INACTIF
        MsgBox description, vbExclamation, "Commande refusée"
        GererErreur = actQuitter

    Case ERR_QUANTITE_INVALIDE
        MsgBox description, vbExclamation, "Saisie invalide"
        GererErreur = actReprendreSuite     ' on laisse l'utilisateur corriger

    ' ... cas des erreurs système (2501, 3022, 3201...) ...

End Select
```

C'est l'aboutissement de l'uniformité recherchée : erreurs système et erreurs métier cohabitent dans le même `Select Case`, sont journalisées de la même façon, et donnent lieu au même type de décision de reprise. L'emploi des constantes nommées rend ces cas parfaitement lisibles, en dépit des valeurs négatives qu'elles recouvrent. On notera ici que la décision de reprise peut différer : une erreur de saisie peut justifier un `Resume Next` pour laisser l'utilisateur corriger, tandis qu'un refus métier conduira à une sortie propre.

## Relancer une erreur : un aperçu

`Err.Raise` ne sert pas seulement à créer des erreurs nouvelles : il permet aussi de **relancer** une erreur déjà interceptée, afin de la faire remonter vers une couche supérieure après un traitement partiel local. Le principe consiste à re-déclencher l'erreur courante depuis un gestionnaire :

```vba
GestionErreur:
    ' ... traitement partiel local : nettoyage, journalisation ...
    Err.Raise lngNumero, strSource, strDescription   ' on relance vers l'appelant
```

Cette technique — qui suppose d'avoir au préalable capturé les valeurs de `Err` avant que le nettoyage ne les efface — est au cœur de la répartition des responsabilités entre procédures. Elle fait l'objet d'un traitement détaillé à la section [13.9](/13-gestion-erreurs/09-propagation-erreurs.md), où la propagation des erreurs est étudiée pour elle-même.

## Erreurs ou valeurs de retour ? Quand utiliser Err.Raise

Disposer de `Err.Raise` ne signifie pas qu'il faille y recourir pour signaler chaque condition anormale. Le choix entre **lever une erreur** et **renvoyer une valeur** est une décision de conception qui mérite réflexion, car l'usage abusif des erreurs nuit à la clarté du code.

La règle générale tient en une formule : réserver les erreurs à l'**exceptionnel**, et les valeurs de retour à l'**attendu**. Une erreur convient lorsque la situation traduit une véritable anomalie — la violation d'une invariante qui ne *devrait* jamais se produire si le code est correct, ou une condition détectée au fond de la pile d'appels, là où faire remonter une valeur à travers chaque niveau serait peu pratique. À l'inverse, une validation de saisie courante — un champ obligatoire laissé vide, un format incorrect dans un formulaire — relève du déroulement normal de l'application : l'utilisateur est *censé* commettre ce genre d'erreur, et il vaut souvent mieux renvoyer un résultat (un booléen, ou une structure décrivant le problème) que de lever une exception.

Détourner le mécanisme d'erreurs pour piloter le flux d'exécution ordinaire — lever une erreur pour chaque champ mal rempli, par exemple — alourdit le code, brouille la distinction entre l'attendu et l'exceptionnel, et finit par banaliser les erreurs au point qu'elles n'alertent plus. `Err.Raise` est un outil précieux ; il gagne à être employé là où une situation mérite réellement le statut d'erreur.

## En résumé

`Err.Raise` permet de déclencher délibérément une erreur d'exécution, principalement pour signaler une violation de règle métier. L'erreur ainsi créée emprunte toute l'infrastructure existante — interception, objet `Err`, journalisation, module centralisé —, ce qui unifie le traitement des anomalies métier et des erreurs système. Sa propriété la plus utile est de **remonter d'elle-même** la pile des appels jusqu'au premier gestionnaire actif.

Le choix du numéro passe par la constante `vbObjectError`, à laquelle on ajoute un décalage (à partir de 512, conventionnellement 1000 et au-delà) pour éviter toute collision avec les codes système ; ces numéros, négatifs et peu lisibles, sont masqués derrière des **constantes nommées** rassemblées dans un catalogue dédié. On renseigne `Source` (sous la forme `module.procédure`) et `Description` (un message concret et explicite) pour enrichir le diagnostic. Une erreur levée atterrit dans le gestionnaire de la procédure courante : pour qu'elle se propage, on la lève dans une procédure sans gestionnaire local, ou on la relance explicitement. Enfin, `Err.Raise` se réserve aux situations réellement exceptionnelles ; pour les conditions attendues, comme la validation de saisie courante, une valeur de retour est souvent préférable.

Plusieurs fois évoquée, la propagation des erreurs le long de la pile des appels constitue le dernier maillon de la gestion d'erreurs. La section suivante l'étudie en détail : comment une erreur remonte, comment relancer proprement, et comment répartir les responsabilités entre les procédures appelantes et appelées.

---


⏭️ [13.9. Hiérarchie de propagation des erreurs entre procédures](/13-gestion-erreurs/09-propagation-erreurs.md)
