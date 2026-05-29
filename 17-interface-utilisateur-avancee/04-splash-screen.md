🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.4. Splash screen et écran de chargement

L'écran d'accueil — le *splash screen* — est cette fenêtre qui s'affiche brièvement au lancement d'une application, présentant son nom, son logo, sa version et parfois un message de chargement. Il remplit deux fonctions distinctes mais complémentaires. La première est **esthétique et identitaire** : il donne immédiatement à l'application l'allure d'un logiciel professionnel, affiche la marque et rassure l'utilisateur sur le fait que le lancement est bien en cours. La seconde est **fonctionnelle** : il masque la phase d'initialisation, souvent inélégante, pendant laquelle l'application vérifie ses liaisons, construit ses menus, charge ses paramètres et ouvre son interface.

Cette double nature explique la distinction entre les deux termes du titre. Un *splash screen* désigne couramment l'écran d'accueil affiché au démarrage ; un *écran de chargement* (*loading screen*) met l'accent sur sa fonction de couverture d'un traitement long, qu'il s'agisse de l'initialisation au lancement ou d'une opération coûteuse en cours d'utilisation. La même technique sous-tend les deux : une fenêtre épurée, maintenue affichée pendant le travail, qui informe l'utilisateur de la progression. Cette section couvre l'option intégrée d'Access, puis la construction et le pilotage d'un écran d'accueil personnalisé par VBA.

## L'option la plus simple : l'image splash intégrée

Access offre un mécanisme de splash screen entièrement gratuit, sans la moindre ligne de code. Il suffit de placer une **image bitmap (`.bmp`) portant exactement le même nom que le fichier de la base de données, dans le même dossier**. Ainsi, pour une base `MonAppli.accdb`, une image `MonAppli.bmp` située à ses côtés sera affichée automatiquement au démarrage, puis disparaîtra une fois l'application chargée.

Cette approche a le mérite de la simplicité absolue : aucune programmation, aucun formulaire à concevoir. Ses limites en font toutefois une solution d'entrée de gamme. L'image est **statique** : elle ne peut afficher ni progression, ni message de chargement évolutif. Sa **durée d'affichage n'est pas réglable** et reste brève. Le format `.bmp` produit par ailleurs des fichiers volumineux comparés aux formats modernes. Pour un outil personnel ou une application modeste, cette image suffit ; mais dès qu'on souhaite un écran d'accueil vivant, informatif ou stylé, le formulaire personnalisé s'impose.

## Le formulaire splash personnalisé

L'écran d'accueil sur mesure est un formulaire Access ordinaire, mais dépouillé de tous ses attributs d'édition pour ne laisser apparaître qu'un contenu visuel propre. Plusieurs propriétés du formulaire doivent être réglées en conséquence :

- **`BorderStyle`** à *Aucune* (ou *Fine*) pour supprimer le cadre redimensionnable.
- **`ControlBox`**, **`CloseButton`**, **`MinMaxButtons`** désactivés pour retirer les boutons système et la barre de titre.
- **`RecordSelectors`**, **`NavigationButtons`**, **`ScrollBars`**, **`DividingLines`** sur *Non*/*Aucune* : un splash n'affiche pas de données navigables.
- **`AutoCenter`** à *Oui* pour centrer la fenêtre à l'écran.
- **`PopUp`** à *Oui* pour qu'il flotte au premier plan, indépendamment de la fenêtre Access.

Le formulaire reste **indépendant** (*unbound*, section 6.8) : il n'est lié à aucune table. On y dispose un logo (contrôle image), un libellé pour le nom de l'application, un autre pour la version et le copyright, et — élément clé pour un écran de chargement — un **libellé d'état** (par exemple `lblStatut`) destiné à afficher l'étape d'initialisation en cours. C'est ce dernier qui transforme un simple splash décoratif en véritable écran de chargement informatif.

## Afficher l'écran d'accueil au démarrage

Deux schémas permettent d'afficher le splash au lancement, selon qu'il sert un objectif purement esthétique ou qu'il accompagne une réelle initialisation.

Le premier, le plus simple, convient à un **splash purement décoratif** sans travail à couvrir. On désigne le formulaire splash comme « Formulaire à afficher » dans les options de la base courante (section 17.7) ; il s'affiche au lancement, puis se ferme de lui-même après un délai géré par son propre minuteur (*timer*), en ouvrant au passage le formulaire principal :

```vba
' --- Dans le module du formulaire frmSplash ---
' Propriété TimerInterval du formulaire réglée à 3000 (3 secondes).
Private Sub Form_Timer()
    Me.TimerInterval = 0                  ' empêche un nouveau déclenchement
    DoCmd.OpenForm "frmAccueil"           ' ouvre l'interface principale
    DoCmd.Close acForm, Me.Name           ' puis se ferme
End Sub
```

Le second schéma, recommandé dès qu'une **initialisation réelle** a lieu, sépare le pilotage du formulaire. On confie le lancement à une **macro `AutoExec`** — une macro portant exactement ce nom, exécutée automatiquement à l'ouverture de la base — qui appelle, via son action `RunCode`, une fonction VBA d'orchestration. Cette fonction ouvre le splash, déroule l'initialisation, ouvre le formulaire principal, puis ferme le splash. C'est ce modèle, le plus contrôlable, que développe la suite.

## L'écran de chargement piloté par l'initialisation

Voici le cœur de la section : un écran d'accueil qui reste affiché pendant que l'application s'initialise réellement, en informant l'utilisateur de chaque étape. La fonction d'orchestration centralise la séquence de démarrage — et c'est précisément à cet endroit que prennent place les opérations vues dans les sections voisines : construction des menus contextuels (section 17.2) ou du ruban (section 17.1), vérification des tables liées (section 21.4), chargement des paramètres.

```vba
' ============================================
' Module standard : modDemarrage
' Orchestre le lancement de l'application.
' Appelé par la macro AutoExec via : RunCode  →  Demarrage()
' ============================================
Option Compare Database
Option Explicit

Private Const DUREE_MINI As Long = 2     ' durée d'affichage minimale (secondes)

Public Function Demarrage()
    Dim debut As Single
    debut = Timer

    On Error GoTo Erreur

    ' 1. Afficher l'écran d'accueil.
    DoCmd.OpenForm "frmSplash"
    AfficherStatut "Démarrage en cours..."

    ' 2. Initialisation, en informant l'utilisateur à chaque étape.
    AfficherStatut "Vérification des liaisons..."
    VerifierTablesLiees                  ' section 21.4

    AfficherStatut "Construction des menus..."
    ConstruireMenuClients                ' section 17.2

    AfficherStatut "Chargement des paramètres..."
    ChargerParametres

    ' 3. Garantir une durée d'affichage minimale (voir plus bas).
    Do While (Timer - debut) < DUREE_MINI
        DoEvents
    Loop

    ' 4. Ouvrir l'interface principale, puis fermer le splash.
    DoCmd.OpenForm "frmAccueil"
    FermerSplash
    Exit Function

Erreur:
    ' En cas d'échec, on ferme toujours le splash avant de signaler l'erreur.
    FermerSplash
    MsgBox "Une erreur est survenue au démarrage :" & vbCrLf & _
           Err.Description, vbCritical, "Démarrage"
End Function
```

Cette fonction concentre toute la logique de lancement en un seul point, ce qui en facilite la maintenance et le débogage. Deux procédures de soutien la complètent — l'affichage du statut et la fermeture sûre du splash — détaillées ci-dessous, car chacune répond à un piège classique.

## Mettre à jour le statut : le problème du rafraîchissement

Un écueil non évident attend ici. Pendant qu'une procédure VBA s'exécute de façon synchrone, **Access ne rafraîchit pas l'affichage du formulaire** : modifier la légende du libellé d'état ne suffit pas à la voir changer à l'écran. Sans précaution, l'utilisateur ne verrait qu'un seul message figé pendant toute l'initialisation.

La parade consiste à **forcer le repeint** du formulaire après chaque mise à jour, en combinant deux mécanismes : la méthode `Repaint` du formulaire, qui applique immédiatement les modifications d'affichage en attente, et l'instruction `DoEvents`, qui rend la main au système le temps de traiter la file des messages d'interface.

```vba
' Met à jour le libellé d'état du splash et force son affichage immédiat.
Private Sub AfficherStatut(ByVal texte As String)
    On Error Resume Next
    Forms!frmSplash!lblStatut.Caption = texte
    Forms!frmSplash.Repaint              ' repeint sans attendre
    DoEvents                              ' laisse l'interface se mettre à jour
End Sub
```

Le `On Error Resume Next` protège l'appel au cas où le splash aurait déjà été fermé. C'est cette procédure, appelée à chaque jalon de l'initialisation, qui donne à l'écran de chargement son caractère vivant et rassurant. Lorsque la progression est quantifiable (un nombre d'étapes ou d'enregistrements connu), on remplacera avantageusement ce libellé par une véritable barre de progression, dont la mise en œuvre fait l'objet de la section suivante (17.5).

## Garantir une durée minimale d'affichage

Sur une machine rapide, l'initialisation peut s'achever en une fraction de seconde, ce qui fait *clignoter* le splash : il apparaît et disparaît si vite que l'utilisateur le perçoit comme un défaut plutôt que comme un écran d'accueil. Pour éviter cet effet, on impose une **durée d'affichage minimale**.

Le principe, visible dans la fonction `Demarrage`, consiste à mémoriser l'heure de début, puis, une fois l'initialisation terminée, à patienter jusqu'à atteindre la durée minimale *seulement si* le travail s'est révélé trop rapide. La fonction `Timer` renvoyant le nombre de secondes écoulées depuis minuit, une simple boucle assortie de `DoEvents` suffit :

```vba
Do While (Timer - debut) < DUREE_MINI
    DoEvents
Loop
```

Cette attente n'a lieu que lorsqu'elle est nécessaire : si l'initialisation a duré plus longtemps que `DUREE_MINI`, la boucle ne s'exécute pas et le splash se ferme immédiatement. La boucle avec `DoEvents` reste une approche simple et acceptable pour une attente de l'ordre de la seconde ; on évite ainsi de recourir à l'API `Sleep` qui, elle, bloquerait tout repeint.

## Fermer le splash en toute circonstance

Le défaut le plus visible — et le plus embarrassant — est un écran d'accueil qui **ne se ferme jamais**, restant suspendu au premier plan parce qu'une erreur est survenue pendant l'initialisation avant l'instruction de fermeture. La règle absolue est donc de garantir la fermeture du splash *quoi qu'il arrive*, ce que réalise le gestionnaire d'erreurs de la fonction `Demarrage` : qu'elle réussisse ou échoue, elle passe par `FermerSplash`.

```vba
' Ferme le splash sans provoquer d'erreur s'il est déjà fermé.
Private Sub FermerSplash()
    On Error Resume Next
    DoCmd.Close acForm, "frmSplash"
End Sub
```

Encadrer l'initialisation par une gestion d'erreurs robuste (chapitre 13) n'est donc pas optionnel : c'est ce qui empêche un démarrage défaillant de laisser l'utilisateur face à un écran bloqué, et qui lui permet au contraire de recevoir un message explicite sur la cause du problème.

## Masquer la fenêtre Access pour un effet pleine page

Aussi soigné soit-il, le formulaire splash flotte par défaut au-dessus de la grande fenêtre grise de l'application Access, qui reste visible derrière lui et trahit son origine. Pour obtenir l'effet d'un écran d'accueil véritablement autonome, on masque la fenêtre principale d'Access pendant l'affichage du splash.

Cette opération relève de l'**API Windows** : on récupère le descripteur de la fenêtre Access via `Application.hWndAccessApp`, puis on la masque et on la réaffiche à l'aide de la fonction `ShowWindow`. Les techniques de déclaration et d'appel d'API depuis Access — avec les précautions de compatibilité 32/64 bits (`PtrSafe`) — sont traitées au chapitre 22 (sections 22.1 et 22.2). Ce masquage se combine naturellement avec l'ensemble des mesures de verrouillage de l'interface (masquage du volet de navigation, options de démarrage) décrites à la section 17.7, dont l'écran d'accueil constitue souvent la première pièce visible.

## Écran d'accueil, écran de chargement et barre de progression

Ces trois notions, parfois confondues, se complètent. L'**écran d'accueil** (splash) est avant tout identitaire : il affiche la marque au lancement et couvre l'initialisation. L'**écran de chargement** désigne ce même type de fenêtre épurée lorsqu'on l'emploie pour masquer un traitement long et en signaler la progression — au démarrage comme en cours d'utilisation (un import volumineux, un calcul lourd). La **barre de progression** (section 17.5) est l'élément visuel qui, à l'intérieur de l'un ou l'autre, matérialise une progression *quantifiée*, là où le libellé d'état se contente de nommer l'étape courante.

En pratique, on superpose ces ingrédients selon le besoin : un splash purement décoratif avec minuteur pour une petite application ; un écran de chargement à libellé d'état pour couvrir une initialisation à étapes nommées mais non chiffrées ; un écran de chargement doté d'une barre de progression lorsque le volume de travail est connu et mérite d'être visualisé précisément.

## Pièges courants

Le piège technique le plus retors est d'ouvrir le splash en mode **dialogue** (`DoCmd.OpenForm "frmSplash", , , , , acDialog`). Ce mode **suspend l'exécution du code** jusqu'à la fermeture du formulaire : l'initialisation qui suit ne démarrerait donc jamais. Un écran d'accueil destiné à rester affiché *pendant* que le code travaille ne doit jamais être ouvert en `acDialog` ; on règle son comportement flottant via la propriété `PopUp` en conception, pas via le mode fenêtre à l'ouverture.

Le piège le plus visible, déjà évoqué, est le **splash qui ne se referme pas** à la suite d'une erreur : seule une gestion d'erreurs garantissant la fermeture l'évite. Vient ensuite le **clignotement** sur machine rapide, corrigé par la durée d'affichage minimale, et le **libellé d'état figé**, conséquence d'un repeint oublié (`Repaint` + `DoEvents`).

Un dernier point concerne le **contournement par la touche Maj**. La macro `AutoExec`, comme l'ensemble des options de démarrage, peut être neutralisée en maintenant la touche Maj enfoncée à l'ouverture de la base : dans ce cas, ni le splash ni l'initialisation ne s'exécutent. C'est un comportement utile au développeur pour reprendre la main, mais qu'il faudra désactiver pour une application livrée — sujet abordé avec les options de démarrage et la sécurité (sections 17.7 et chapitre 20).

## Synthèse

L'écran d'accueil remplit un double rôle, identitaire et fonctionnel, en présentant la marque de l'application tout en masquant son initialisation. L'option la plus économique est l'**image splash intégrée** (un `.bmp` du même nom que la base), suffisante pour les besoins simples mais statique. Le **formulaire splash personnalisé** — indépendant, sans bordure ni contrôle d'édition, flottant via `PopUp` et centré — offre en revanche un écran vivant : maintenu affiché par une fonction d'orchestration lancée depuis une macro `AutoExec`, il informe l'utilisateur de chaque étape grâce à un libellé d'état.

Trois réflexes techniques conditionnent sa réussite : forcer le rafraîchissement du libellé par `Repaint` et `DoEvents`, imposer une durée d'affichage minimale pour éviter le clignotement, et garantir la fermeture du splash en toute circonstance via une gestion d'erreurs. Pour un effet pleinement autonome, on masque la fenêtre Access par l'API Windows (chapitre 22). Lorsque la progression est quantifiable, l'écran de chargement accueille une barre de progression, objet de la section suivante.

---

> 🔗 **Pour aller plus loin :** le formulaire splash est un formulaire indépendant (section 6.8) ; son affichage au démarrage et le contournement par la touche Maj relèvent des options de démarrage (section 17.7) ; les étapes d'initialisation reprennent la construction des menus (sections 17.1 et 17.2) et la vérification des tables liées (section 21.4) ; la visualisation chiffrée de la progression fait l'objet de la section 17.5 ; le masquage de la fenêtre Access mobilise l'API Windows (sections 22.1 et 22.2) ; et la gestion d'erreurs garantissant la fermeture s'appuie sur le chapitre 13.

⏭️ [17.5. Barre de progression personnalisée dans un formulaire](/17-interface-utilisateur-avancee/05-barre-progression.md)
