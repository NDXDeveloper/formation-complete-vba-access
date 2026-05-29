🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.2. Menus contextuels personnalisés (CommandBars)

Le menu contextuel — celui qui surgit au clic droit — est un point d'accès aux commandes que les utilisateurs expérimentés sollicitent en permanence. Par défaut, Access affiche un menu générique proposant copier-coller, tri et filtres, sans le moindre rapport avec la logique de votre application. Le remplacer par un menu sur mesure, proposant les actions réellement pertinentes au contexte (« Valider la commande », « Exporter cette fiche », « Voir l'historique »), améliore sensiblement l'ergonomie et renforce l'impression d'un logiciel pensé pour ses utilisateurs.

Contrairement au ruban principal, qui se définit en XML (section 17.1), les menus contextuels d'Access se construisent essentiellement par programmation, à travers un modèle objet ancien mais toujours d'actualité : **CommandBars**. Cette section explique ce modèle, la manière de bâtir des menus contextuels statiques ou dynamiques, de les rattacher à l'application ou à des objets précis, et de relier leurs commandes à votre code VBA.

## CommandBars dans l'Access moderne : où en est-on ?

Une clarification s'impose d'emblée, car le sujet prête à confusion. Le modèle CommandBars était, jusqu'à Office 2003, le mécanisme universel de personnalisation de *toute* l'interface : barres d'outils, barres de menus et menus contextuels. À partir d'Access 2007, le **ruban** a remplacé les barres d'outils et les barres de menus. Les anciennes barres d'outils créées en CommandBars continuent de fonctionner, mais sont reléguées dans un onglet « Compléments » du ruban : on ne les utilise plus pour bâtir une interface moderne.

En revanche — et c'est le point décisif — **les menus contextuels (menus de raccourci) restent gérés par CommandBars** dans Access. Le ruban ne les prend pas en charge de la même façon que dans d'autres applications Office. Construire un menu de clic droit personnalisé en Access passe donc encore aujourd'hui, et de façon parfaitement supportée, par le modèle CommandBars. Ce qui est obsolète, ce sont les *barres d'outils* en CommandBars ; les *menus contextuels* en CommandBars demeurent la voie standard. C'est pourquoi cette technique, malgré son âge, n'a rien d'un vestige : elle est la bonne réponse au besoin précis traité ici.

Il existe bien une alternative déclarative en XML (l'élément `<contextMenus>` de l'espace de noms RibbonX 2009), évoquée en fin de section, mais elle est moins souple pour les menus dynamiques et CommandBars reste l'approche de référence.

## Le modèle objet CommandBars

La collection `CommandBars`, accessible directement en Access (`Application.CommandBars` ou simplement `CommandBars`), regroupe toutes les barres de commande de l'application. Pour un menu contextuel, trois types d'objets entrent en jeu.

L'objet **`CommandBar`** représente la barre elle-même. Pour un menu contextuel, c'est une barre de type *popup* (`msoBarPopup`), c'est-à-dire un menu flottant déclenché à la demande plutôt qu'une barre fixe.

L'objet **`CommandBarControl`** représente un élément du menu. Il se décline en plusieurs types selon le rôle, les deux principaux étant le **`CommandBarButton`** (`msoControlButton`), un élément cliquable, et le **`CommandBarPopup`** (`msoControlPopup`), un sous-menu contenant lui-même des contrôles.

Ces objets et leurs constantes (`msoBarPopup`, `msoControlButton`, `msoControlPopup`…) proviennent de la **bibliothèque Microsoft Office Object Library**, normalement déjà référencée (section 2.5). À défaut de cette référence, on peut recourir aux valeurs numériques et à la liaison tardive (section 2.6), au prix de l'autocomplétion.

Les propriétés les plus utiles d'un contrôle de menu sont les suivantes :

| Propriété | Rôle |
|-----------|------|
| `.Caption` | Le texte affiché de l'élément |
| `.OnAction` | Le nom de la procédure VBA à exécuter au clic |
| `.FaceId` | L'icône intégrée (identifiant numérique) |
| `.BeginGroup` | À `True`, insère une ligne de séparation au-dessus de l'élément |
| `.Enabled` | Active ou grise l'élément |
| `.Visible` | Affiche ou masque l'élément |
| `.State` | État coché/décoché (`msoButtonDown` / `msoButtonUp`) |
| `.Tag` / `.Parameter` | Données libres, utiles pour identifier un contrôle ou transmettre un contexte |

## Créer un menu contextuel

La construction d'un menu suit toujours la même séquence : créer la barre popup, puis y ajouter les contrôles un à un. Le code ci-dessous bâtit un menu contextuel pour une liste de clients, dans la lignée des exemples des sections précédentes :

```vba
' ============================================
' Module standard : modMenusContextuels
' Construit les menus contextuels de l'application.
' À appeler UNE FOIS au démarrage (macro AutoExec ou
' formulaire de démarrage), car les barres temporaires
' sont recréées à chaque session.
' ============================================
Option Compare Database
Option Explicit

Private Const NOM_MENU As String = "mnuClients"

Public Sub ConstruireMenuClients()
    Dim cmb As Office.CommandBar
    Dim btn As Office.CommandBarButton
    Dim sousMenu As Office.CommandBarPopup
    Dim btnSub As Office.CommandBarButton

    ' Supprimer une version antérieure pour éviter un doublon de nom.
    SupprimerMenu NOM_MENU

    ' Créer le menu contextuel : barre de type popup, temporaire.
    Set cmb = CommandBars.Add(Name:=NOM_MENU, _
                              Position:=msoBarPopup, _
                              Temporary:=True)

    ' --- Élément « Actualiser » ---
    Set btn = cmb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Actualiser"
    btn.FaceId = 459              ' icône intégrée (numéro découvert via galerie)
    btn.OnAction = "MenuActualiser"

    ' --- Élément « Nouveau client », précédé d'un séparateur ---
    Set btn = cmb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Nouveau client"
    btn.FaceId = 2174
    btn.BeginGroup = True
    btn.OnAction = "MenuNouveauClient"

    ' --- Élément « Supprimer » (repéré par .Tag pour activation dynamique) ---
    Set btn = cmb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Supprimer le client"
    btn.FaceId = 478
    btn.Tag = "SUPPR"
    btn.OnAction = "MenuSupprimer"

    ' --- Sous-menu « Exporter vers » ---
    Set sousMenu = cmb.Controls.Add(Type:=msoControlPopup)
    sousMenu.Caption = "Exporter vers"
    sousMenu.BeginGroup = True

    Set btnSub = sousMenu.Controls.Add(Type:=msoControlButton)
    btnSub.Caption = "PDF"
    btnSub.OnAction = "MenuExporter"
    btnSub.Parameter = "PDF"      ' contexte lu par le handler

    Set btnSub = sousMenu.Controls.Add(Type:=msoControlButton)
    btnSub.Caption = "Excel"
    btnSub.OnAction = "MenuExporter"
    btnSub.Parameter = "XLSX"
End Sub

' Supprime une barre si elle existe (sans erreur sinon).
Private Sub SupprimerMenu(ByVal nom As String)
    On Error Resume Next
    CommandBars(nom).Delete
End Sub
```

Trois points méritent l'attention. L'option `Temporary:=True` crée une barre non persistante, automatiquement supprimée à la fermeture d'Access ; c'est le comportement recommandé pour ne pas polluer durablement l'environnement. La suppression préalable (`SupprimerMenu`) évite l'erreur « ce nom existe déjà » si la procédure est rappelée dans la même session. Enfin, le sous-menu `msoControlPopup` se construit exactement comme la barre principale : on lui ajoute des contrôles via sa propre collection `.Controls`.

## Le mécanisme OnAction : relier les clics au code VBA

À l'image des callbacks du ruban, la propriété `.OnAction` ne contient pas de code mais le **nom d'une procédure VBA**, qu'Access exécute lorsque l'utilisateur clique sur l'élément. Cette procédure doit être **publique** et résider dans un **module standard**. Sa signature peut être un simple `Sub` sans paramètre.

Lorsqu'un même handler doit servir plusieurs éléments (les deux entrées du sous-menu d'export, par exemple), on n'écrit pas une procédure par élément : on utilise `CommandBars.ActionControl`, qui, à l'intérieur du handler, renvoie le contrôle effectivement cliqué. On lit alors sa propriété `.Parameter` (ou `.Tag`, ou `.Caption`) pour adapter le comportement :

```vba
' --- Handlers OnAction (publics, dans un module standard) ---

Public Sub MenuActualiser()
    On Error Resume Next
    If Not Screen.ActiveForm Is Nothing Then Screen.ActiveForm.Requery
End Sub

Public Sub MenuNouveauClient()
    On Error GoTo Erreur
    DoCmd.OpenForm "frmClient", , , , acFormAdd
    Exit Sub
Erreur:
    MsgBox "Ouverture impossible : " & Err.Description, vbExclamation
End Sub

Public Sub MenuSupprimer()
    On Error Resume Next
    If Not Screen.ActiveForm Is Nothing Then
        If MsgBox("Supprimer ce client ?", vbYesNo + vbQuestion) = vbYes Then
            DoCmd.RunCommand acCmdDeleteRecord
        End If
    End If
End Sub

' UN SEUL handler pour les deux éléments du sous-menu :
' on lit .Parameter du contrôle cliqué via ActionControl.
Public Sub MenuExporter()
    Dim ctl As Office.CommandBarControl
    Set ctl = CommandBars.ActionControl    ' le contrôle à l'origine du clic
    Select Case ctl.Parameter
        Case "PDF":  MsgBox "Export PDF demandé."
        Case "XLSX": MsgBox "Export Excel demandé."
    End Select
End Sub
```

Comme pour le ruban, on constate que les handlers délèguent l'essentiel à l'objet **DoCmd** (chapitre 5) ou au code applicatif : le menu n'est qu'un déclencheur. La propriété `.OnAction` accepte aussi la syntaxe d'appel de fonction avec arguments (`=MaFonction("arg")`), mais le recours à `ActionControl` et à `.Parameter` reste plus lisible et plus robuste pour mutualiser un handler.

## Affecter le menu à l'application ou à un objet

Une fois la barre popup construite, elle doit être *rattachée* pour s'afficher au clic droit. Trois niveaux sont possibles.

Au niveau de **l'application entière** : **Fichier → Options → Base de données active → champ « Barre de menus contextuels »**, où l'on indique le nom de la barre popup. Ce menu remplace alors le menu contextuel par défaut partout dans l'application.

Au niveau d'un **formulaire, état ou contrôle particulier** : ces objets possèdent une propriété **« Barre de menus contextuels »** (onglet *Autres* de la feuille de propriétés), à laquelle on affecte le nom de la barre. Le clic droit sur cet objet affiche alors ce menu spécifique. Une seconde propriété, **« Menu contextuel »** (Oui/Non), permet d'activer ou de désactiver entièrement l'affichage du menu contextuel sur l'objet.

À la demande, par code, via la méthode **`ShowPopup`** : `CommandBars("mnuClients").ShowPopup` affiche le menu à l'emplacement du curseur. C'est l'approche la plus souple, indispensable pour les menus dynamiques décrits plus bas.

Un impératif de **chronologie** découle de tout cela : comme les barres temporaires sont recréées à chaque session, elles doivent exister *avant* d'être référencées. Si la propriété « Barre de menus contextuels » d'un formulaire pointe vers un menu non encore construit, l'affectation échoue. On bâtit donc les menus tôt dans le démarrage de l'application — typiquement depuis une macro **AutoExec** ou l'événement d'ouverture du formulaire de démarrage (en lien avec les options de démarrage, section 17.7).

## Menus contextuels dynamiques

Un menu contextuel gagne souvent à refléter le contexte : griser « Supprimer » en l'absence d'enregistrement, masquer une commande réservée à certains droits, ou même construire des éléments à partir des données. Deux stratégies répondent à ce besoin.

Pour des ajustements simples (activer/désactiver, afficher/masquer), on modifie les propriétés `.Enabled` ou `.Visible` des contrôles lorsque le contexte évolue — par exemple dans l'événement `Current` du formulaire (section 8.8).

Pour un contrôle complet du moment d'affichage, on combine la méthode `ShowPopup` avec l'événement `MouseUp` du formulaire ou du contrôle (section 8.9). On désactive alors le menu automatique (propriété « Menu contextuel » à *Non* pour supprimer l'affichage par défaut), on intercepte le clic droit, on ajuste les éléments, puis on affiche le menu soi-même :

```vba
' --- Dans le module du formulaire ---
' Prérequis : propriété « Menu contextuel » du formulaire = Non,
' afin de supprimer le menu automatique et de tout piloter ici.
Private Sub Detail_MouseUp(Button As Integer, Shift As Integer, _
                           X As Single, Y As Single)
    If Button = acRightButton Then
        AjusterEtAfficherMenuClients
    End If
End Sub
```

```vba
' --- Dans modMenusContextuels ---
Public Sub AjusterEtAfficherMenuClients()
    Dim cmb As Office.CommandBar
    Dim ctl As Office.CommandBarControl

    ' S'assurer que le menu existe (le reconstruire au besoin).
    On Error Resume Next
    Set cmb = CommandBars(NOM_MENU)
    On Error GoTo 0
    If cmb Is Nothing Then
        ConstruireMenuClients
        Set cmb = CommandBars(NOM_MENU)
    End If

    ' Activer « Supprimer » uniquement si un enregistrement est présent.
    For Each ctl In cmb.Controls
        If ctl.Tag = "SUPPR" Then
            ctl.Enabled = ContexteSuppressionValide()
        End If
    Next ctl

    ' Afficher le menu à l'emplacement du curseur.
    cmb.ShowPopup
End Sub

Private Function ContexteSuppressionValide() As Boolean
    On Error Resume Next
    ContexteSuppressionValide = _
        (Not Screen.ActiveForm Is Nothing) And _
        (Not Screen.ActiveForm.NewRecord)   ' pas sur la ligne « nouvelle »
End Function
```

C'est ici que `.Tag` prend tout son sens : il sert d'étiquette stable pour retrouver un contrôle précis dans la collection, indépendamment de sa position ou de son libellé.

## Icônes des éléments

L'icône d'un élément se définit le plus simplement via la propriété **`.FaceId`**, qui réutilise les icônes intégrées d'Office à partir d'un identifiant numérique. Ces numéros constituent un référentiel ancien, distinct des noms `imageMso` du ruban (section 17.1) ; ils se découvrent via les galeries publiées ou des utilitaires dédiés, et aucune liste n'est reproduite ici.

Pour une **icône personnalisée**, les propriétés `.Picture` et `.Mask` acceptent un objet `IPictureDisp` (la première pour l'image, la seconde pour la transparence). La fabrication d'un tel objet à partir d'un fichier relève du même utilitaire que celui évoqué pour le ruban, dont un modèle réutilisable figure à l'annexe K. Dans la pratique, les icônes `FaceId` intégrées couvrent l'essentiel des besoins et évitent cette complexité.

## Nettoyage et cycle de vie

La gestion du cycle de vie des barres conditionne la fiabilité de l'ensemble. L'option `Temporary:=True` règle la majeure partie du problème : les barres temporaires disparaissent à la fermeture d'Access, sans accumulation d'une session à l'autre. Au sein d'une même session, la **suppression préalable avant reconstruction** (le réflexe `On Error Resume Next : CommandBars(nom).Delete`) évite les doublons de nom. À l'inverse, créer des barres **non** temporaires sans jamais les supprimer conduit, à terme, à une prolifération de menus fantômes dans l'environnement — une mauvaise pratique à proscrire.

Le bon schéma consiste donc à concentrer la construction des menus dans une procédure unique, appelée une fois au démarrage, et à reconstruire à la demande si nécessaire. Cette procédure relève d'un module standard dédié, nommé selon les conventions du chapitre 24.

## Gérer les erreurs dans les handlers

Comme pour les callbacks du ruban, un handler `OnAction` qui laisse remonter une erreur non interceptée produit un comportement peu maîtrisé. La règle est d'**encadrer les handlers par une gestion d'erreurs** (chapitre 13) : un déroutement vers un gestionnaire pour les actions susceptibles d'échouer (ouverture de formulaire, exécution de commande), un simple `On Error Resume Next` pour les actions anodines. Les procédures qui interrogent le contexte (`Screen.ActiveForm`, par exemple) doivent par ailleurs tolérer l'absence de formulaire actif sans planter, d'où la protection systématique des accès à `Screen` observée dans les exemples.

## L'alternative déclarative : RibbonX `<contextMenus>`

Pour ceux qui préfèrent une approche cohérente avec le ruban, l'espace de noms RibbonX 2009 (`http://schemas.microsoft.com/office/2009/07/customui`) introduit l'élément `<contextMenus>`, qui permet d'**enrichir les menus contextuels intégrés** en y ajoutant ses propres commandes, déclarées en XML et reliées à des callbacks identiques à ceux du ruban (section 17.1). On référence le menu intégré à compléter par son identifiant interne (`idMso`), puis on y insère boutons et sous-menus.

Cette voie a l'avantage d'unifier toute la définition de l'interface dans un seul document XML stocké dans `USysRibbons`, et de réutiliser le modèle de callbacks déjà maîtrisé. Elle est cependant **moins souple que CommandBars pour les menus pleinement dynamiques** (construction d'éléments à partir de données, par exemple), où elle impose de passer par des callbacks `getEnabled`/`getVisible` et l'invalidation. Pour des compléments ponctuels à un menu existant, elle est élégante ; pour un menu contextuel entièrement personnalisé et dynamique, CommandBars conserve l'avantage. Le choix dépend donc de la nature du besoin : enrichissement déclaratif d'un menu standard d'un côté, menu sur mesure et réactif de l'autre.

## Pièges courants

Le piège le plus fréquent tient à la **chronologie de construction** : référencer un menu (via une propriété « Barre de menus contextuels » ou un `ShowPopup`) avant qu'il n'ait été bâti. Comme les barres temporaires ne survivent pas à la fermeture, oublier de les reconstruire au démarrage les rend introuvables. La parade est de toujours construire les menus tôt, depuis AutoExec ou le formulaire de démarrage.

Vient ensuite la **prolifération de barres** : créer des barres non temporaires, ou omettre la suppression préalable avant reconstruction, laisse s'accumuler des menus en double ou résiduels. L'option `Temporary:=True` assortie d'un `Delete` préalable règle ce point.

Côté handlers, les erreurs classiques rejoignent celles du ruban : une procédure `OnAction` placée par mégarde dans un module de classe ou de formulaire au lieu d'un module standard, ou un nom cité dans `.OnAction` ne correspondant à aucune procédure (élément alors inactif). Tout refactoring (section 24.5) doit préserver la correspondance entre les noms cités et les procédures présentes.

Enfin, un **double affichage** survient lorsqu'on combine maladroitement l'affectation par propriété et l'appel à `ShowPopup` : le menu automatique et le menu déclenché manuellement s'ajoutent. Pour un menu dynamique piloté par `ShowPopup`, il faut désactiver le menu automatique en réglant la propriété « Menu contextuel » de l'objet sur *Non*.

## Synthèse

Les menus contextuels d'Access se construisent par programmation via le modèle **CommandBars**, qui — contrairement aux barres d'outils, désormais remplacées par le ruban — demeure le mécanisme standard et supporté pour personnaliser le clic droit. On crée une barre de type `msoBarPopup`, on y ajoute des boutons (`msoControlButton`) et des sous-menus (`msoControlPopup`), et l'on relie chaque élément à une procédure VBA publique via sa propriété `.OnAction`, en mutualisant les handlers grâce à `CommandBars.ActionControl` et aux propriétés `.Parameter`/`.Tag`.

Le menu se rattache au niveau de l'application, d'un objet précis, ou s'affiche à la demande par `ShowPopup` — cette dernière méthode, combinée à l'événement `MouseUp`, ouvrant la voie aux menus dynamiques dont le contenu et l'activation suivent le contexte. La maîtrise du cycle de vie (barres temporaires, suppression avant reconstruction, construction au démarrage) garantit la fiabilité de l'ensemble. Pour un simple enrichissement de menus intégrés, l'élément `<contextMenus>` de RibbonX offre une alternative déclarative ; pour un menu sur mesure et réactif, CommandBars reste l'outil de choix.

---

> 🔗 **Pour aller plus loin :** les handlers `OnAction` s'appuient sur l'objet DoCmd (chapitre 5) ; les types `CommandBar`/`CommandBarControl` relèvent de la bibliothèque Office (section 2.5) et de la liaison précoce/tardive (section 2.6) ; l'affichage dynamique mobilise les événements `MouseUp` et `Current` (sections 8.9 et 8.8) ; la construction au démarrage se coordonne avec les options de démarrage (section 17.7) ; l'alternative `<contextMenus>` réutilise le modèle de callbacks du ruban (section 17.1) ; et un chargeur d'image `IPictureDisp` réutilisable figure à l'annexe K.

⏭️ [17.3. Formulaire de navigation (Navigation Form) contrôlé par VBA](/17-interface-utilisateur-avancee/03-formulaire-navigation.md)
