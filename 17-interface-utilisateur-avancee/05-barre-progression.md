🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.5. Barre de progression personnalisée dans un formulaire

Lorsqu'une application exécute une opération longue — l'import d'un fichier volumineux, le traitement de milliers d'enregistrements, la génération d'un état complexe — l'utilisateur a besoin de savoir qu'elle travaille, et d'estimer le temps restant. Sans retour visuel, une application qui semble figée pendant trente secondes est perçue comme bloquée ou défaillante, alors même qu'elle fonctionne parfaitement. La barre de progression répond à ce besoin : elle matérialise l'avancement, transforme une attente anxieuse en attente comprise, et confère à l'application le sérieux d'un logiciel abouti.

Cette section est le prolongement naturel de la précédente (17.4) : la barre de progression est l'élément qui, à l'intérieur d'un écran de chargement ou d'une boîte de dialogue dédiée, donne une mesure *chiffrée* de la progression, là où un simple libellé d'état se contente de nommer l'étape en cours. On y verra comment construire une barre entièrement personnalisée à même un formulaire, comment la piloter sans pénaliser les performances, et comment l'encapsuler pour la réutiliser.

## Trois approches possibles

Trois voies permettent d'afficher une progression en Access, dont l'intitulé de cette section désigne clairement la principale.

La première est le **compteur intégré de `SysCmd`**, qui affiche une barre dans la barre d'état d'Access, en bas de la fenêtre. Sans aucune conception, c'est l'option la plus économique — l'équivalent, pour la progression, de l'image splash intégrée vue en 17.4.

La deuxième, celle que développe cette section, est la **barre personnalisée construite dans un formulaire** à l'aide de rectangles dont on fait varier la largeur. Elle n'exige aucune dépendance, s'intègre où on le souhaite (notamment sur l'écran de chargement) et se personnalise librement. C'est l'approche recommandée.

La troisième est le **contrôle ActiveX ProgressBar** des composants communs Windows. D'apparence native, il présente l'inconvénient majeur d'une dépendance externe à déployer et entretenir, ce qui le rend fragile en distribution ; il est abordé en fin de section, avec ses réserves.

## La barre intégrée : le compteur SysCmd

L'objet `SysCmd` (présenté à la section 4.7) expose un compteur de progression prêt à l'emploi, piloté par trois appels : l'initialisation avec un texte et une valeur maximale, la mise à jour avec la valeur courante, et le retrait final.

```vba
Public Sub TraiterAvecSysCmd()
    Dim total As Long, n As Long
    total = 5000

    ' 1. Initialiser le compteur (texte + maximum).
    SysCmd acSysCmdInitMeter, "Traitement en cours...", total

    Do While n < total
        n = n + 1
        ' ... traitement ...

        ' 2. Mettre à jour la position.
        SysCmd acSysCmdUpdateMeter, n
    Loop

    ' 3. Retirer le compteur.
    SysCmd acSysCmdRemoveMeter
End Sub
```

Cette solution est imbattable de simplicité, mais ses limites sont réelles. La barre s'affiche **dans la barre d'état**, en bas de l'écran, où elle passe facilement inaperçue — et où elle est invisible si la barre d'état a été masquée. Son apparence n'est pas personnalisable et sa position est imposée. Pour une progression discrète d'une opération secondaire, elle convient ; pour un retour visuel central et soigné, on lui préférera la barre personnalisée.

## La barre personnalisée par rectangles

Le principe de la barre personnalisée est d'une élégante simplicité : on superpose deux rectangles sur un formulaire. Un rectangle de **fond** (la « piste »), de largeur fixe, représente les 100 %. Un rectangle de **remplissage**, placé exactement au même endroit (mêmes `Left` et `Top`) et doté d'une couleur de remplissage, voit sa **largeur varier proportionnellement à l'avancement**, de zéro jusqu'à la largeur totale de la piste. Un libellé par-dessus affiche le pourcentage chiffré.

Un point technique conditionne le calcul : les dimensions des contrôles (`Width`, `Height`, `Left`, `Top`) s'expriment en **twips** — l'unité interne d'Access, où 1 440 twips valent un pouce (2,54 cm) et 567 twips environ un centimètre. La largeur du rectangle de remplissage se calcule donc en twips, comme une fraction de la largeur de la piste. Modifier la largeur d'un contrôle à l'exécution relève des techniques de manipulation dynamique des contrôles (section 17.11), parfaitement applicables sur un formulaire de progression en mode simple.

Sur un formulaire `frmProgression` contenant un rectangle de fond `boxFond`, un rectangle de remplissage `boxRemplissage` (superposé) et un libellé `lblPourcent`, une procédure de mise à jour s'écrit ainsi :

```vba
' --- Dans le module du formulaire frmProgression ---
Option Compare Database
Option Explicit

' Remet la barre à zéro à l'ouverture.
Private Sub Form_Load()
    Me!boxRemplissage.Width = 0
    Me!lblPourcent.Caption = "0 %"
End Sub

' Met la barre à jour pour une valeur sur un total donné.
Public Sub MajProgression(ByVal valeur As Long, ByVal total As Long)
    Dim ratio As Double
    If total <= 0 Then Exit Sub

    ratio = valeur / total
    If ratio < 0 Then ratio = 0
    If ratio > 1 Then ratio = 1

    ' La largeur du remplissage = fraction de la largeur de la piste.
    Me!boxRemplissage.Width = CLng(Me!boxFond.Width * ratio)
    Me!lblPourcent.Caption = Format(ratio, "0%")
    Me.Repaint                          ' indispensable : voir ci-dessous
End Sub
```

## Le rafraîchissement, à nouveau

On retrouve ici, exactement, le problème déjà rencontré pour le libellé d'état de l'écran de chargement (section 17.4) : pendant qu'une procédure VBA s'exécute de façon synchrone, **Access ne repeint pas l'écran**. Modifier la largeur du rectangle de remplissage ne suffit donc pas à voir la barre progresser ; sans intervention, elle resterait figée jusqu'à la fin du traitement, puis sauterait d'un coup à 100 %.

La solution est identique : forcer le repeint après chaque mise à jour, par la méthode `Repaint` du formulaire — qui applique immédiatement les modifications d'affichage en attente — éventuellement complétée par `DoEvents` lorsqu'on souhaite aussi que l'interface reste réactive (notamment pour qu'un bouton d'annulation puisse fonctionner, voir plus bas). Cette répétition n'est pas un hasard : le rafraîchissement forcé est le mécanisme commun à tous les retours visuels pilotés depuis une boucle de traitement.

## La performance : ne pas rafraîchir trop souvent

Voici le point le plus important de la section, et l'erreur la plus fréquente. Repeindre l'écran a un coût. Si l'on appelle `MajProgression` — donc `Repaint` — à **chaque itération** d'une boucle qui traite cent mille enregistrements, les repeints finissent par dominer le temps d'exécution : l'opération devient *plus lente avec la barre que sans elle*, parfois dans des proportions spectaculaires. Paradoxalement, la barre censée informer l'utilisateur de l'avancement ralentit considérablement ce qu'elle mesure.

La parade consiste à **ne rafraîchir que lorsque l'affichage changerait réellement**. Puisque la barre ne progresse visuellement que par pas de 1 %, il est inutile de la repeindre cent fois entre 50 % et 51 % : on ne met à jour que lorsque le **pourcentage entier change**. Cette simple condition réduit le nombre de repeints à cent au maximum pour toute l'opération, quel que soit le nombre d'enregistrements, et élimine la pénalité de performance :

```vba
' Exemple d'utilisation : traitement d'un jeu d'enregistrements,
' avec une barre rafraîchie seulement au changement de pourcentage.
Public Sub TraiterClients()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim total As Long, n As Long
    Dim dernierPourcent As Integer

    On Error GoTo Nettoyage

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblClients", dbOpenSnapshot)

    ' Connaître le total est indispensable pour une barre déterminée.
    If Not rs.EOF Then
        rs.MoveLast
        total = rs.RecordCount
        rs.MoveFirst
    End If

    DoCmd.OpenForm "frmProgression"
    dernierPourcent = -1

    Do While Not rs.EOF
        ' ... traitement de l'enregistrement courant ...

        n = n + 1
        ' Ne rafraîchir QUE si le pourcentage entier a changé.
        If Int(n / total * 100) <> dernierPourcent Then
            dernierPourcent = Int(n / total * 100)
            Forms!frmProgression.MajProgression n, total
            DoEvents
        End If

        rs.MoveNext
    Loop

Nettoyage:
    ' Fermer la barre quoi qu'il arrive (cf. 17.4).
    If CurrentProject.AllForms("frmProgression").IsLoaded Then
        DoCmd.Close acForm, "frmProgression"
    End If
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    If Err.Number <> 0 Then
        MsgBox "Erreur durant le traitement : " & Err.Description, vbCritical
    End If
End Sub
```

On note au passage la fermeture systématique de la barre dans la section de nettoyage, dans le même esprit que la fermeture garantie du splash vue en 17.4 : une opération longue qui échoue ne doit jamais laisser une barre de progression suspendue à l'écran. Plus largement, cette logique de mesure et d'optimisation rejoint les techniques du chapitre 18 sur la performance (sections 18.1 et 18.2).

## Encapsuler dans une classe réutilisable

Recopier la logique de pilotage dans chaque traitement long serait une duplication regrettable. Comme la barre de progression est un objet réutilisable au comportement bien défini, elle se prête idéalement à un **module de classe** (chapitre 16) qui en encapsule entièrement la mécanique — calcul du ratio, gestion des twips, rafraîchissement conditionnel — derrière une interface simple. Le code appelant n'a alors plus qu'à initialiser la barre puis à lui transmettre l'avancement, sans se soucier ni des rectangles ni des repeints.

```vba
' ============================================
' Module de classe : clsBarreProgression
' Encapsule le pilotage d'une barre de progression
' construite sur un formulaire (rectangles + libellé).
' Le rafraîchissement conditionnel est géré en interne.
' ============================================
Option Compare Database
Option Explicit

Private mForm As Access.Form
Private mFond As Access.Control
Private mRemplissage As Access.Control
Private mLibelle As Access.Control
Private mTotal As Long
Private mLargeurMax As Long
Private mDernierPourcent As Integer

' Initialise la barre à partir du formulaire hôte et d'un total.
' (Les contrôles suivent une convention de nommage standard.)
Public Sub Initialiser(ByRef frm As Access.Form, ByVal total As Long)
    Set mForm = frm
    Set mFond = frm.Controls("boxFond")
    Set mRemplissage = frm.Controls("boxRemplissage")
    Set mLibelle = frm.Controls("lblPourcent")

    mLargeurMax = mFond.Width
    mTotal = total
    mDernierPourcent = -1
    mRemplissage.Width = 0
    mLibelle.Caption = "0 %"
End Sub

' Propriété Valeur : fixe l'avancement et ne rafraîchit
' que si le pourcentage entier a changé.
Public Property Let Valeur(ByVal v As Long)
    Dim pourcent As Integer
    If mTotal <= 0 Then Exit Property

    pourcent = Int(v / mTotal * 100)
    If pourcent < 0 Then pourcent = 0
    If pourcent > 100 Then pourcent = 100

    If pourcent <> mDernierPourcent Then
        mDernierPourcent = pourcent
        mRemplissage.Width = CLng(mLargeurMax * pourcent / 100)
        mLibelle.Caption = pourcent & " %"
        mForm.Repaint
        DoEvents
    End If
End Property
```

L'utilisation devient limpide, et toute la complexité disparaît du code métier :

```vba
Dim barre As clsBarreProgression
Set barre = New clsBarreProgression

DoCmd.OpenForm "frmProgression"
barre.Initialiser Forms!frmProgression, total

Do While Not rs.EOF
    ' ... traitement ...
    n = n + 1
    barre.Valeur = n          ' le rafraîchissement conditionnel est automatique
    rs.MoveNext
Loop
```

Cette version illustre concrètement l'apport de la programmation orientée objet à l'interface : un composant autonome, testable, réutilisable d'un traitement à l'autre, dont l'optimisation du rafraîchissement est garantie une fois pour toutes à l'intérieur de la classe.

## Connaître le total : barres déterminées et indéterminées

Une barre de progression chiffrée suppose un **total connu à l'avance**. Lorsqu'on traite un jeu d'enregistrements, ce total s'obtient par `rs.RecordCount` après un `MoveLast` (le déplacement en fin de jeu étant nécessaire pour qu'Access compte tous les enregistrements, comme le rappelle le chapitre 9) ; lorsqu'on traite une liste de fichiers ou une séquence d'étapes, il s'agit simplement de leur nombre. Tant que ce dénominateur existe, la barre affiche une fraction fidèle.

Mais certaines opérations ont une **durée indéterminée** : une requête dont on ignore le temps d'exécution, l'attente d'une réponse réseau, un traitement dont le volume n'est pas connu d'avance. Une barre de progression classique ne peut alors rien afficher de pertinent — quel pourcentage montrer sans dénominateur ? Dans ce cas, on renonce à la barre chiffrée au profit d'un **indicateur d'activité** : un libellé « Traitement en cours… » animé, une barre dite *indéterminée* (un bloc qui va et vient), ou simplement le retour de statut de l'écran de chargement (section 17.4). L'important est de distinguer dès la conception ce qui est *mesurable* (barre chiffrée) de ce qui ne l'est pas (indicateur d'activité), pour ne pas chercher à quantifier l'inquantifiable.

## Un bouton d'annulation

Un raffinement appréciable consiste à permettre à l'utilisateur d'**interrompre** une opération longue. Le mécanisme s'appuie sur le `DoEvents` déjà présent : on place un bouton « Annuler » sur le formulaire de progression, dont le clic positionne un indicateur ; la boucle de traitement consulte cet indicateur à chaque tour et s'arrête proprement si l'annulation a été demandée. Sans `DoEvents`, le clic ne serait jamais traité pendant la boucle — c'est lui qui rend le bouton réactif.

```vba
' --- Sur le formulaire frmProgression ---
Private mDemandeAnnulation As Boolean

Private Sub cmdAnnuler_Click()
    mDemandeAnnulation = True
End Sub

Public Property Get AnnulationDemandee() As Boolean
    AnnulationDemandee = mDemandeAnnulation
End Property
```

```vba
' --- Dans la boucle de traitement ---
Do While Not rs.EOF
    If Forms!frmProgression.AnnulationDemandee Then
        Exit Do                ' arrêt propre demandé par l'utilisateur
    End If
    ' ... traitement ...
    rs.MoveNext
Loop
```

Ce motif transforme une opération « tout ou rien » en un traitement maîtrisable, ce que les utilisateurs apprécient particulièrement pour les imports et exports de longue haleine.

## Le contrôle ActiveX ProgressBar

Pour mémoire, les composants communs Windows (`mscomctl.ocx`) fournissent un contrôle ActiveX *ProgressBar* qu'on peut déposer sur un formulaire et piloter par ses propriétés `Min`, `Max` et `Value`. Son apparence est native et lisse, ce qui peut séduire.

Cet attrait est cependant largement contrebalancé par les inconvénients propres aux contrôles ActiveX (section 8.5). Le composant introduit une **dépendance externe** qui doit être installée et enregistrée sur **chaque poste** utilisateur, sous peine d'une application qui ne s'ouvre plus. Il pose des **problèmes de compatibilité 32/64 bits** délicats en environnement Office 64 bits, et peut être bloqué par les paramètres de sécurité. Ces fragilités en font un mauvais choix pour une application destinée à être déployée largement. La barre personnalisée par rectangles, sans aucune dépendance et pleinement maîtrisée, lui est presque toujours préférable — raison pour laquelle cette section en a fait son sujet principal.

## Où placer la barre de progression

La barre s'intègre à deux endroits selon le moment du traitement. Au **démarrage**, elle prend place sur l'écran de chargement (section 17.4) lorsque l'initialisation comporte des étapes nombreuses et quantifiables, en lieu et place — ou en complément — du libellé d'état. En **cours d'utilisation**, elle s'affiche dans une **boîte de dialogue de progression dédiée**, un petit formulaire `PopUp` ouvert au lancement de l'opération longue et fermé à son terme, exactement sur le modèle d'organisation décrit en 17.4.

Dans les deux cas, le formulaire hôte adopte les caractéristiques d'un écran épuré (sans bordure superflue, centré, flottant) et — point déjà souligné — ne doit jamais être ouvert en mode `acDialog`, qui suspendrait le code chargé de faire avancer la barre.

## Pièges courants

Le piège de performance domine : **rafraîchir à chaque itération** ralentit le traitement au lieu de l'accompagner. Le rafraîchissement conditionnel au changement de pourcentage entier est la réponse à connaître et à appliquer systématiquement.

Vient ensuite le **rafraîchissement oublié** : modifier la largeur du rectangle sans appeler `Repaint` laisse la barre figée jusqu'à la fin, donnant l'illusion d'un blocage — l'inverse exact de l'effet recherché.

L'**absence de total** conduit à vouloir afficher une progression chiffrée là où aucun dénominateur n'existe : il faut alors basculer vers un indicateur d'activité plutôt que de forcer une barre déterminée.

La **barre qui ne se ferme pas** après une erreur reprend l'écueil de 17.4 : seule une gestion d'erreurs fermant le formulaire de progression en toute circonstance l'évite.

Enfin, l'ouverture en **`acDialog`** du formulaire de progression bloquerait l'exécution du traitement, exactement comme pour le splash : la barre ne progresserait jamais.

## Synthèse

La barre de progression transforme une attente opaque en avancement compris, et constitue le pendant chiffré de l'écran de chargement de la section précédente. Parmi les trois approches — compteur intégré `SysCmd` dans la barre d'état, barre personnalisée par rectangles, contrôle ActiveX —, la **barre par rectangles** s'impose : sans dépendance, librement stylable, intégrable partout, elle consiste à faire varier la largeur d'un rectangle de remplissage (en twips) proportionnellement à l'avancement, par-dessus un rectangle de fond.

Deux impératifs techniques en gouvernent la réussite : forcer le rafraîchissement par `Repaint`/`DoEvents`, et surtout **ne rafraîchir qu'au changement de pourcentage entier** pour ne pas pénaliser les performances. Encapsulée dans une classe réutilisable, dotée éventuellement d'un bouton d'annulation s'appuyant sur `DoEvents`, et fermée en toute circonstance par une gestion d'erreurs soignée, la barre devient un composant fiable et élégant de l'application. Elle suppose un total connu ; à défaut, un indicateur d'activité prend le relais.

---

> 🔗 **Pour aller plus loin :** la barre prolonge l'écran de chargement (section 17.4) et en partage la logique de rafraîchissement et de fermeture garantie ; la modification de la largeur des rectangles à l'exécution relève de la manipulation dynamique des contrôles (section 17.11) ; le compteur intégré utilise l'objet `SysCmd` (section 4.7) ; l'encapsulation dans une classe illustre la programmation orientée objet (chapitre 16) ; l'obtention du total via `RecordCount` renvoie au chapitre 9 ; les enjeux de performance rejoignent le chapitre 18 (sections 18.1 et 18.2) ; et le contrôle ActiveX relève des réserves de la section 8.5.

⏭️ [17.6. Gestion du focus et navigation clavier avancée](/17-interface-utilisateur-avancee/06-focus-navigation-clavier.md)
