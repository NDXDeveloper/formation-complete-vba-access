🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 13.9. Hiérarchie de propagation des erreurs entre procédures

## Introduction

Une erreur ne survient presque jamais dans une procédure isolée : elle se produit au sein d'une **chaîne d'appels**, où une procédure en appelle une autre, qui en appelle une troisième. La question devient alors : laquelle de ces procédures doit traiter l'erreur ? Celle où elle s'est produite ? Celle qui l'a appelée ? La toute première, à la racine de la chaîne ?

Les sections précédentes ont plusieurs fois effleuré le sujet — une erreur non gérée localement « remonte » vers l'appelant (13.2), une erreur métier levée au fond du code rejoint le premier gestionnaire actif (13.8). Il est temps de l'étudier pour lui-même. Cette section, qui clôt le chapitre, explique le mécanisme exact de la propagation, montre comment relancer une erreur proprement, et surtout pose le principe architectural qui guide la répartition des responsabilités : où informer l'utilisateur, où nettoyer, où simplement laisser passer. C'est la maîtrise de cette hiérarchie qui distingue une gestion d'erreurs cohérente à l'échelle d'une application d'une simple accumulation de gestionnaires locaux.

## La pile des appels

Lorsqu'une procédure A appelle une procédure B, qui appelle elle-même une procédure C, ces appels s'empilent : C s'exécute « au-dessus » de B, qui s'exécute « au-dessus » de A. Cette structure, appelée **pile des appels** (*call stack*), reflète l'ordre dans lequel les procédures ont été imbriquées. Tant que C s'exécute, B et A sont en attente, suspendues à l'instruction d'appel.

La propagation des erreurs suit cette pile, mais en sens inverse : du sommet (C, la procédure la plus profonde) vers la base (A, la procédure initiale). Comprendre ce parcours est la clé de tout ce qui suit.

## Le mécanisme de propagation

Lorsqu'une erreur survient dans une procédure, VBA cherche un gestionnaire actif selon une règle précise, en remontant la pile :

1. Si la **procédure courante** dispose d'un gestionnaire `On Error GoTo` actif, l'erreur y est dirigée. La recherche s'arrête.
2. Sinon, la procédure courante est **abandonnée**, et l'erreur est transmise à la **procédure appelante**. La même question se pose alors à ce niveau : dispose-t-elle d'un gestionnaire actif ?
3. Ce processus se répète, de proche en proche, en remontant la pile.
4. Si **aucune** procédure de la chaîne ne possède de gestionnaire, le comportement par défaut s'applique : l'exécution s'arrête et la boîte de dialogue système s'affiche.

Ainsi, une erreur survenue dans C, si C n'a pas de gestionnaire, remonte vers B ; si B n'en a pas davantage, elle remonte vers A ; et si A en possède un, c'est lui qui la traite. L'erreur « cherche » donc le premier gestionnaire disponible en descendant la pile, et c'est à ce niveau qu'elle est finalement prise en charge.

## Le déroulement de la pile et le piège du nettoyage ignoré

Cette remontée a une conséquence trop souvent négligée. Lorsqu'une procédure sans gestionnaire est traversée par une erreur, elle est **abandonnée sur-le-champ** : l'exécution la quitte immédiatement, sans atteindre la fin de son corps. Tout code de nettoyage qu'elle contiendrait — placé, comme le veut le patron de la section [13.2](/13-gestion-erreurs/02-on-error-goto.md), après une étiquette atteinte par le flux normal — ne sera **pas exécuté**.

L'effet est direct : si la procédure C, dépourvue de gestionnaire, avait ouvert un recordset ou posé un verrou, ces ressources ne seront **jamais libérées** lorsqu'une erreur la traverse. Elles fuient, car le chemin qui les aurait fermées a été court-circuité par la remontée.

De là découle une règle impérative : **toute procédure qui alloue des ressources doit posséder son propre gestionnaire**, ne serait-ce que pour garantir leur libération avant que l'erreur ne poursuive sa remontée. À l'inverse, une procédure qui n'alloue aucune ressource — une simple fonction de calcul ou de validation — peut sans dommage se passer de gestionnaire et laisser l'erreur passer au travers. C'est cette distinction qui guide le choix entre les stratégies présentées plus loin.

## Une subtilité : Erl pointe sur le site d'appel

Lorsqu'un gestionnaire situé dans une procédure appelante intercepte une erreur survenue dans une procédure appelée, une subtilité affecte le diagnostic. Les propriétés `Err.Number` et `Err.Description` reflètent fidèlement l'**erreur d'origine** — celle qui s'est réellement produite dans la procédure profonde. En revanche, la fonction `Erl`, qui renvoie un numéro de ligne, indique la ligne de la **procédure qui intercepte**, c'est-à-dire celle de l'**appel** à la procédure défaillante — et non la ligne fautive à l'intérieur de celle-ci.

Autrement dit, si A intercepte une erreur survenue dans C (C et B n'ayant pas de gestionnaire), `Erl` dans A désignera la ligne où A a appelé B, pas la ligne précise où l'erreur s'est produite dans C. On sait donc *quel appel* a échoué, mais pas *où exactement* à l'intérieur. Cette limite a une implication pratique importante pour la journalisation, sur laquelle on reviendra : pour connaître la ligne exacte d'origine, il faut que la procédure profonde dispose de son propre gestionnaire capable de relever sa propre valeur de `Erl`.

## Resume après propagation

La propagation modifie aussi le comportement de `Resume`. On l'a établi à la section [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md) : `Resume` ne s'exécute que dans la procédure dont le gestionnaire a capté l'erreur. Après une propagation, ce gestionnaire n'est pas celui de la procédure où l'erreur est née, mais celui de la procédure qui l'a interceptée plus bas dans la pile.

La conséquence est qu'un `Resume` exécuté dans le gestionnaire de A reprendra l'exécution **au niveau de A** — en réessayant l'instruction de A qui appelait B (`Resume`), ou en passant à la suivante dans A (`Resume Next`). Les procédures intermédiaires B et C, abandonnées lors de la remontée, ont disparu : il est **impossible** de reprendre l'exécution à l'intérieur de C. La reprise opère toujours au niveau du gestionnaire qui a capté l'erreur, jamais au niveau où elle s'est produite. C'est une raison supplémentaire de placer un gestionnaire au plus près des opérations délicates si l'on souhaite pouvoir les réessayer finement.

## Trois stratégies de répartition des responsabilités

Le mécanisme de propagation ouvre trois manières de répartir le traitement des erreurs entre les procédures. Le choix se fait procédure par procédure, selon sa nature.

### Traiter localement

La procédure gère entièrement ses propres erreurs : elle les intercepte, les journalise, informe l'utilisateur et décide de la reprise. C'est le modèle autonome décrit aux sections [13.2](/13-gestion-erreurs/02-on-error-goto.md) et [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md). Il convient aux procédures de haut niveau, en particulier aux gestionnaires d'événements, qui constituent souvent le point d'entrée d'une chaîne d'appels.

### Laisser propager

La procédure ne comporte **aucun gestionnaire** et laisse les erreurs remonter telles quelles vers un niveau supérieur. C'est le choix indiqué pour les petites procédures sans ressources à libérer — fonctions de calcul, de validation, utilitaires —, comme la procédure `ValiderCommande` de la section [13.8](/13-gestion-erreurs/08-erreurs-personnalisees-raise.md). On évite ainsi de dupliquer du code de gestion là où il n'apporte rien, et l'on concentre le traitement à un niveau pertinent.

### Traiter partiellement puis relancer

La procédure intercepte l'erreur pour accomplir un traitement **local et limité** — typiquement libérer ses ressources, éventuellement journaliser —, puis **relance** l'erreur afin qu'une couche supérieure prenne la décision finale. C'est la stratégie la plus élaborée, indispensable aux procédures qui détiennent des ressources mais n'ont pas vocation à décider de l'interaction avec l'utilisateur.

## Relancer une erreur proprement

La troisième stratégie repose sur la relance, dont la technique a été introduite à la section [13.8](/13-gestion-erreurs/08-erreurs-personnalisees-raise.md). Sa mise en œuvre exige un ordre précis, dicté par le cycle de vie de l'objet `Err` :

```vba
Public Sub ChargerDonnees()
    Dim rst As DAO.Recordset
    Dim lngNum As Long, strSrc As String, strDesc As String

    On Error GoTo GestionErreur

    Set rst = CurrentDb.OpenRecordset("R_Synthese", dbOpenSnapshot)
    ' ... traitement ...

SortieProcedure:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Exit Sub

GestionErreur:
    ' 1. Capturer AVANT que le nettoyage ne réinitialise Err
    lngNum = Err.Number
    strSrc = Err.Source
    strDesc = Err.Description

    ' 2. Nettoyage local des ressources
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing

    ' 3. Relancer vers l'appelant, qui décidera de la suite
    Err.Raise lngNum, strSrc, strDesc
End Sub
```

Trois temps se succèdent. On **capture** d'abord les propriétés de `Err` dans des variables locales : c'est impératif, car le nettoyage qui suit — avec son `On Error` et ses opérations — réinitialiserait l'objet `Err`, et l'on perdrait l'information à relancer (rappel de la section [13.4](/13-gestion-erreurs/04-objet-err.md)). On **nettoie** ensuite les ressources de la procédure. On **relance** enfin l'erreur d'origine, à l'identique, au moyen des valeurs capturées.

Cette relance fonctionne précisément parce qu'un gestionnaire ne capte pas les erreurs survenant en son sein : l'erreur relancée depuis le gestionnaire ne lui revient pas, mais propage vers l'appelant — comportement déjà signalé parmi les pièges de la section [13.2](/13-gestion-erreurs/02-on-error-goto.md). La procédure a ainsi rempli sa part (libérer ses ressources) avant de déléguer la décision finale à un niveau plus apte à la prendre.

## Le principe architectural : informer en haut, propager en bas

Au-delà de la mécanique, la propagation répond à une question de conception : **où l'utilisateur doit-il être informé de l'erreur ?** La réponse, dans une application bien structurée, est claire : **au sommet de la pile**, dans la couche d'interface, et non dans les couches profondes d'accès aux données ou de logique métier.

Plusieurs raisons motivent ce principe. Les couches basses sont conçues pour être **réutilisables** : une fonction d'accès aux données peut être appelée depuis un formulaire, mais aussi depuis un traitement par lots non interactif, où un `MsgBox` n'aurait aucun sens — voire bloquerait le traitement en attendant un clic qui ne viendra jamais. Une couche basse qui afficherait des messages s'arrogerait par ailleurs une responsabilité — l'interaction avec l'utilisateur — qui ne lui revient pas, et perdrait sa neutralité. Enfin, centraliser l'information de l'utilisateur au sommet garantit la cohérence du discours de l'application, dans l'esprit du module centralisé de la section [13.7](/13-gestion-erreurs/07-module-centralise-erreurs.md).

La règle pratique est donc : les couches basses **nettoient leurs ressources et propagent** (ou relancent), sans afficher de message ; la couche d'interface **intercepte, informe et décide** de la reprise. Chaque niveau assume la responsabilité qui correspond à sa nature.

## Un exemple en couches

L'exemple suivant illustre ce principe sur trois niveaux : accès aux données, logique métier, interface.

```vba
' ----- Couche d'accès aux données : nettoie puis propage, sans message -----
Public Function LireSoldeClient(ByVal idClient As Long) As Currency
    Dim rst As DAO.Recordset
    Dim lngNum As Long, strSrc As String, strDesc As String

    On Error GoTo GestionErreur

    Set rst = CurrentDb.OpenRecordset( _
        "SELECT Solde FROM T_Clients WHERE ID=" & idClient, dbOpenSnapshot)
    If Not rst.EOF Then LireSoldeClient = Nz(rst!Solde, 0)

SortieProcedure:
    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Exit Function

GestionErreur:
    lngNum = Err.Number: strSrc = Err.Source: strDesc = Err.Description
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Err.Raise lngNum, strSrc, strDesc      ' propage, sans informer l'utilisateur
End Function
```

```vba
' ----- Couche métier : pas de ressource, laisse propager (et peut lever) -----
Public Function PeutPasserCommande(ByVal idClient As Long, _
                                   ByVal montant As Currency) As Boolean
    ' Aucun gestionnaire local : les erreurs d'accès remontent telles quelles.
    If LireSoldeClient(idClient) < montant Then
        Err.Raise ERR_SOLDE_INSUFFISANT, "modCommande.PeutPasserCommande", _
                  "Solde insuffisant pour cette commande."
    End If
    PeutPasserCommande = True
End Function
```

```vba
' ----- Couche interface : c'est ICI que l'utilisateur est informé -----
Private Sub cmdValider_Click()
    On Error GoTo GestionErreur

    If PeutPasserCommande(Me!IDClient, Me!Montant) Then
        ' ... enregistrement de la commande ...
        MsgBox "Commande enregistrée.", vbInformation
    End If

SortieProcedure:
    Exit Sub

GestionErreur:
    Select Case GererErreur(Err.Number, Err.Description, Err.Source, _
                            "cmdValider_Click", Erl)
        Case actReprendre:      Resume
        Case actReprendreSuite: Resume Next
        Case Else:              Resume SortieProcedure
    End Select
End Sub
```

Le déroulement se lit ainsi : une erreur d'accès survenant dans `LireSoldeClient` y est interceptée le temps de fermer le recordset, puis relancée ; elle traverse `PeutPasserCommande` (sans gestionnaire) sans s'y arrêter ; elle atteint enfin le gestionnaire de `cmdValider_Click`, où l'utilisateur est informé via le module centralisé. Une erreur métier levée dans `PeutPasserCommande` suit le même chemin jusqu'au sommet. Chaque couche a joué exactement le rôle qui lui revient. Cette organisation est l'application concrète de l'architecture en couches étudiée au chapitre 16.

## Journaliser au bon endroit

La subtilité de `Erl` évoquée plus haut a une incidence sur la journalisation, qu'il faut arbitrer. Si l'on journalise **uniquement au sommet**, dans le gestionnaire centralisé, on bénéficie d'un point de consignation unique et simple, mais la ligne enregistrée sera celle du site d'appel au sommet, et non la ligne exacte d'origine ; le numéro et la description, eux, restent fidèles à l'erreur réelle. Si l'on souhaite la **ligne précise d'origine**, il faut journaliser dans la couche profonde, là où l'erreur naît et où `Erl` désigne la bonne ligne, avant de relancer.

Le risque, dans ce second cas, est la **double journalisation** : une fois en bas, une fois au sommet. On l'évite en choisissant un point de consignation unique. L'approche la plus simple et la plus répandue consiste à journaliser une seule fois, au sommet, en acceptant la granularité du site d'appel. Lorsque la précision de la ligne importe pour une opération critique, on ajoute alors la numérotation des lignes et un gestionnaire de journalisation dans la procédure profonde concernée, en veillant à ne pas re-journaliser plus haut. C'est un arbitrage entre simplicité et précision, à trancher selon les besoins de diagnostic de l'application.

## Bonnes pratiques

La gestion de la propagation se résume à quelques principes complémentaires :

- **Doter d'un gestionnaire toute procédure qui alloue des ressources**, afin de garantir leur libération avant que l'erreur ne poursuive sa remontée — faute de quoi le nettoyage est court-circuité.
- **Laisser propager les procédures sans ressources** (calcul, validation), pour ne pas dupliquer inutilement du code de gestion.
- **Pour relancer, respecter l'ordre** : capturer `Err` d'abord, nettoyer ensuite, relancer enfin avec les valeurs capturées.
- **Informer l'utilisateur au sommet**, dans la couche d'interface, et jamais dans les couches basses réutilisables.
- **Choisir un point de journalisation unique** pour éviter les doublons, en recourant à `Erl` et à un gestionnaire profond seulement lorsque la ligne exacte est nécessaire.
- **Se rappeler que `Resume` opère au niveau du gestionnaire qui a capté l'erreur**, jamais au niveau où elle s'est produite.

## En résumé

Une erreur se propage le long de la pile des appels, du point où elle survient vers la base : elle est dirigée vers le premier gestionnaire actif rencontré en remontant, et provoque le comportement par défaut si aucun n'existe. Cette remontée abandonne au passage les procédures sans gestionnaire, dont le code de nettoyage est alors ignoré — d'où la nécessité de doter d'un gestionnaire toute procédure détenant des ressources. Deux subtilités accompagnent la propagation : `Erl` désigne, dans la procédure qui intercepte, la ligne de l'appel et non celle de l'erreur d'origine ; et `Resume` reprend toujours au niveau du gestionnaire qui a capté l'erreur, jamais dans les procédures abandonnées.

Trois stratégies se combinent selon la nature de chaque procédure : traiter localement (couches hautes), laisser propager (procédures sans ressources), ou nettoyer puis relancer (couches détenant des ressources). La relance suit un ordre strict — capturer, nettoyer, relancer — dicté par le cycle de vie de `Err`. Surtout, un principe architectural gouverne l'ensemble : **informer l'utilisateur au sommet, propager dans les couches basses**, afin de préserver la réutilisabilité des couches profondes et la cohérence du discours applicatif.

Ainsi s'achève le chapitre consacré à la gestion des erreurs. Du catalogue des erreurs courantes jusqu'à leur propagation hiérarchique, en passant par la structure `On Error GoTo`, l'objet `Err`, l'interception spécifique des erreurs de données, la journalisation et la centralisation, il a dessiné une infrastructure complète. Gérer les erreurs ne se réduit jamais à empêcher les plantages : c'est faire de l'erreur un événement prévu, tracé et maîtrisé, à tous les niveaux de l'application. Cette maîtrise prend toute sa dimension lorsqu'elle se conjugue à la préservation de l'intégrité des données — précisément l'objet du chapitre suivant, consacré aux transactions, où la gestion d'une erreur déclenche l'annulation propre d'un traitement inachevé.

---


⏭️ [14. Transactions et intégrité des données](/14-transactions/README.md)
