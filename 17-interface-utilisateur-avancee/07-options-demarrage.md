🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.7. Options de démarrage et verrouillage de l'interface

Toutes les techniques de finition vues jusqu'ici — ruban personnalisé, menus contextuels, formulaire de navigation, écran d'accueil — convergent vers une même intention : faire en sorte que l'utilisateur perçoive un **logiciel** et non une base de données ouverte dans Access. Cette section en est l'aboutissement. Elle traite de deux facettes complémentaires : la **configuration du démarrage** — ce qui se produit et s'affiche à l'ouverture du fichier — et le **verrouillage de l'interface** — le masquage des outils de développement d'Access pour que l'utilisateur ne puisse ni voir ni modifier la structure de l'application.

Ces réglages se font en grande partie dans une boîte de dialogue d'options, mais — c'est l'angle de cette formation — ils sont aussi accessibles **par code**, ce qui permet de les appliquer automatiquement au déploiement. Deux avertissements majeurs structureront le propos, car ils correspondent aux erreurs les plus coûteuses du sujet : le piège du verrouillage de la touche Maj, et la confusion entre verrouillage de l'interface et sécurité réelle.

## Les options de la base courante

L'essentiel de la configuration du démarrage se trouve dans **Fichier → Options → Base de données active**. Ces options s'appliquent au **fichier courant** — concrètement, au front-end de l'application (architecture front-end/back-end, section 15.1) — et non à l'installation d'Access. Les principales sont les suivantes.

Le **titre de l'application** remplace « Microsoft Access » dans la barre de titre, et l'**icône de l'application** y substitue l'icône Access — deux touches identitaires immédiates. Le **formulaire à afficher** désigne le formulaire ouvert automatiquement au lancement (souvent l'écran de navigation de la section 17.3 ou, indirectement, l'écran d'accueil de la section 17.4). Le **nom du ruban** (section 17.1) et la **barre de menus contextuels** (section 17.2) raccordent ici les rubans et menus personnalisés construits précédemment.

Viennent ensuite les options de verrouillage proprement dites : **afficher le volet de navigation** (à décocher pour masquer la liste des objets), **autoriser les menus complets** (pour restreindre le ruban à un jeu limité), **autoriser les menus contextuels par défaut**, et **utiliser les touches spéciales d'Access** (qui gouverne F11, Ctrl+G et les autres raccourcis de développement). Une dernière option, **compacter lors de la fermeture**, déclenche un compactage automatique à chaque sortie (à manier avec prudence en multi-utilisateur).

Un point essentiel : **la plupart de ces options ne prennent effet qu'à la prochaine ouverture** de la base. C'est une source fréquente de confusion lorsqu'on les modifie et qu'on ne voit aucun changement immédiat.

## Configurer les options par code

Régler ces options manuellement dans la boîte de dialogue convient au développement, mais une application déployée gagne à les **appliquer par programmation** : cela rend la configuration reproductible, automatisable lors d'une installation (section 21.8), et réparable si un utilisateur l'a altérée. Ces options de démarrage sont en effet stockées comme des **propriétés de la base de données**, que l'on peut créer et modifier en DAO (propriétés personnalisées, section 12.8).

Une difficulté technique se présente : **ces propriétés n'existent pas tant qu'elles n'ont pas été définies au moins une fois**. Tenter de modifier une propriété absente provoque une erreur ; il faut alors la créer. D'où un motif canonique, à connaître, qui tente d'abord la modification et, en cas d'échec, crée la propriété :

```vba
' ============================================
' Module standard : modDemarrageOptions
' Configure les options de démarrage et le verrouillage
' de l'interface par code.
' ============================================
Option Compare Database
Option Explicit

' Crée ou modifie une propriété de la base (motif canonique).
Private Sub DefinirPropriete(ByVal nom As String, _
                             ByVal typeProp As Integer, _
                             ByVal valeur As Variant)
    Dim db As DAO.Database
    Dim prop As DAO.Property

    Set db = CurrentDb
    On Error Resume Next
    db.Properties(nom) = valeur                ' tente la modification
    If Err.Number <> 0 Then                     ' la propriété n'existe pas...
        Err.Clear
        Set prop = db.CreateProperty(nom, typeProp, valeur)
        db.Properties.Append prop               ' ...on la crée
    End If
    On Error GoTo 0
    Set prop = Nothing
    Set db = Nothing
End Sub
```

Les noms et types des propriétés de démarrage les plus utiles sont les suivants : `AppTitle` (texte, titre), `AppIcon` (texte, chemin de l'icône), `StartupForm` (texte, formulaire de démarrage), `StartupShowDBWindow` (booléen, volet de navigation), `StartupShowStatusBar` (booléen, barre d'état), `AllowFullMenus` (booléen), `AllowBuiltInToolbars` (booléen), `AllowShortcutMenus` (booléen, menus contextuels par défaut), `AllowSpecialKeys` (booléen, touches spéciales), `StartupMenuBar` et `StartupShortcutMenuBar` (texte, barres personnalisées).

Une procédure unique applique alors toute la configuration :

```vba
Public Sub ConfigurerDemarrage()
    DefinirPropriete "AppTitle", dbText, "Gestion Clients"
    DefinirPropriete "StartupForm", dbText, "frmAccueil"
    DefinirPropriete "StartupShowStatusBar", dbBoolean, True
    DefinirPropriete "StartupShortcutMenuBar", dbText, "mnuClients"  ' section 17.2

    Application.RefreshTitleBar         ' applique immédiatement titre et icône
    MsgBox "Configuration appliquée. Effet complet à la prochaine ouverture.", _
           vbInformation
End Sub
```

On notera l'appel à `Application.RefreshTitleBar` : la modification du titre et de l'icône ne se reflète dans la barre de titre qu'après cet appel, les autres options attendant la réouverture.

Il convient de distinguer ces propriétés de base de données des **options propres à l'installation** d'Access (comportement de la touche Entrée évoqué en 17.6, affichage des erreurs d'interface vu en 17.1). Ces dernières, qui valent pour tous les fichiers ouverts sur le poste, s'obtiennent et se règlent plutôt via les méthodes `Application.GetOption` et `Application.SetOption` (objet `Application`, section 4.2), tandis que les options de démarrage ci-dessus sont attachées au fichier lui-même.

## Le formulaire de démarrage et la macro AutoExec

Deux mécanismes pilotent ce qui s'exécute au lancement, et il faut comprendre leur articulation. Le **formulaire de démarrage** (`StartupForm`) est ouvert automatiquement par Access. La **macro `AutoExec`** — une macro portant exactement ce nom — s'exécute, elle aussi, automatiquement à l'ouverture, et peut lancer du code VBA via son action `RunCode`.

Plutôt que de chercher à coordonner finement ces deux dispositifs, l'approche la plus claire et la plus contrôlable consiste à **n'en retenir qu'un seul comme point d'entrée**. La voie recommandée, déjà adoptée pour l'écran de chargement (section 17.4), est de laisser le formulaire de démarrage vide et de confier le lancement à une macro `AutoExec` qui appelle une **fonction d'orchestration** unique. Cette fonction maîtrise alors entièrement la séquence : application des options, affichage de l'écran d'accueil, initialisation, ouverture du formulaire principal. On obtient ainsi un démarrage entièrement scénarisé en VBA, sans dépendre de l'ordre exact entre formulaire de démarrage et macro.

## Verrouiller l'interface

Le verrouillage proprement dit empile plusieurs réglages dont l'effet conjugué fait disparaître l'environnement de développement aux yeux de l'utilisateur.

**Masquer le volet de navigation** (`StartupShowDBWindow = False`) retire la liste des tables, requêtes, formulaires et états — la porte d'entrée vers la structure de l'application. On peut aussi le masquer par code à l'exécution, en lui donnant le focus puis en masquant la fenêtre active :

```vba
' Masquer le volet de navigation à l'exécution.
DoCmd.NavigateTo "acNavigationCategoryObjectType"
DoCmd.RunCommand acCmdWindowHide
```

**Désactiver les touches spéciales** (`AllowSpecialKeys = False`) neutralise les raccourcis de développement, au premier rang desquels **F11** (affichage du volet de navigation) et **Ctrl+G** (fenêtre d'exécution immédiate). Sans cette désactivation, un utilisateur peut rouvrir d'une touche ce que les autres options ont masqué.

**Restreindre les menus** (`AllowFullMenus = False`, `AllowShortcutMenus = False`) limite le ruban à un jeu réduit et supprime les menus contextuels par défaut — utilement remplacés par les menus personnalisés de la section 17.2.

L'effet de ces réglages est **cumulatif** : pris isolément, chacun se contourne ; combinés, ils ferment l'essentiel des accès à la structure. C'est en les appliquant ensemble — volet masqué, touches spéciales désactivées, menus restreints, ruban et menus personnalisés substitués — qu'on obtient une application qui se présente comme un logiciel clos.

## La touche Maj et AllowBypassKey : le piège majeur

Tout l'édifice précédent repose pourtant sur une faille béante, qu'il faut absolument connaître : **maintenir la touche Maj enfoncée à l'ouverture de la base contourne l'intégralité des options de démarrage et de la macro `AutoExec`**. L'utilisateur se retrouve alors dans la base brute, volet de navigation visible, menus complets, comme si aucun verrouillage n'existait. Ce contournement est volontaire et précieux pour le développeur — c'est ainsi qu'on reprend la main sur une base verrouillée — mais il anéantit le verrouillage pour quiconque le connaît.

Pour le neutraliser, on désactive la propriété **`AllowBypassKey`** (booléen), qui — comme les autres — doit être créée par code :

```vba
' DÉSACTIVER le contournement par Maj.  ⚠ À MANIER AVEC LA PLUS GRANDE PRUDENCE.
Public Sub DesactiverContournementMaj()
    DefinirPropriete "AllowBypassKey", dbBoolean, False
End Sub

' RÉACTIVER le contournement — la « porte de secours » INDISPENSABLE.
Public Sub ReactiverContournementMaj()
    DefinirPropriete "AllowBypassKey", dbBoolean, True
    MsgBox "Contournement par Maj réactivé.", vbInformation
End Sub
```

**Voici l'avertissement le plus important de cette section.** Une fois `AllowBypassKey` désactivée, **vous non plus ne pouvez plus contourner le démarrage** par la touche Maj. Si, dans le même temps, le verrouillage de l'interface vous empêche d'exécuter du code pour la réactiver, vous êtes **enfermé hors de votre propre application** — et la seule issue peut être de repartir d'une copie non verrouillée. Désactiver le contournement sans avoir prévu de **porte de secours** est l'erreur qui a coûté le plus de bases Access à leurs auteurs.

La porte de secours peut prendre plusieurs formes : une commande cachée dans un menu d'administration réservé (accessible après authentification, section 20.6), un raccourci clavier secret capté par `KeyPreview` (section 17.6) exécutant `ReactiverContournementMaj`, ou — le plus sûr — une **base d'administration distincte** qui ouvre la base verrouillée et y rétablit la propriété. Quelle que soit la solution, la règle est absolue : **ne jamais désactiver `AllowBypassKey` sans disposer d'un moyen éprouvé de le réactiver**, et toujours conserver une copie maîtresse non verrouillée à l'abri.

## Niveaux de protection : obscurité, ACCDE, sécurité réelle

Il faut être parfaitement lucide sur ce que ces options protègent — et ce qu'elles ne protègent pas. **Le verrouillage de l'interface relève de l'obscurité, non de la sécurité.** Masquer le volet de navigation et désactiver les touches spéciales décourage la manipulation accidentelle ou opportuniste par un utilisateur ordinaire ; cela n'oppose **aucune résistance** à quelqu'un de déterminé et un tant soit peu informé, qui connaît la touche Maj ou sait ouvrir le fichier autrement. Ces réglages visent à « empêcher les honnêtes gens de faire des bêtises », pas à contrer une attaque.

La protection véritable se situe à deux autres niveaux, traités au chapitre 20. La **compilation en ACCDE** (section 20.3) supprime le code source VBA et interdit toute modification de la conception des formulaires, états et modules : c'est elle qui protège réellement la **logique** de l'application. La **sécurité des données** — chiffrement, mot de passe de base, gestion applicative des droits (sections 20.2, 20.6, 20.8) — protège quant à elle les **données**. Le verrouillage de l'interface vient **par-dessus** ces mesures, non à leur place : il soigne l'expérience et écarte les tentations, tandis qu'ACCDE et la sécurité des données constituent les remparts.

Cette hiérarchie se complète au déploiement. Une application distribuée via **Access Runtime** (section 21.2) ne dispose d'aucun environnement de conception : le verrouillage y est en partie naturel, puisqu'il n'y a ni ruban de création ni éditeur VBA accessibles. Les couches s'empilent alors : Runtime supprime l'environnement de développement, l'ACCDE protège le code, les options de démarrage soignent la présentation, et la sécurité des données garde l'information.

## Méthodologie de développement

Le verrouillage complique, par nature, le développement : une fois les options actives, modifier l'application suppose de les contourner. D'où quelques règles de méthode. On **développe avec les options désactivées**, en n'appliquant le verrouillage qu'en phase finale, juste avant la livraison. On **conserve systématiquement une copie maîtresse non verrouillée**, qui reste l'exemplaire de travail et la source des évolutions futures. Tant que `AllowBypassKey` est active, la touche Maj reste le moyen normal de reprendre la main pour le développeur. Et l'on teste le verrouillage **sur une copie**, jamais sur l'unique exemplaire, précisément à cause du risque d'enfermement décrit plus haut.

## Pièges courants

Le piège de loin le plus grave, longuement développé, est l'**enfermement** consécutif à la désactivation de `AllowBypassKey` sans porte de secours : il impose la discipline de la copie maîtresse et d'un mécanisme de réactivation.

Vient ensuite l'**effet différé** : la plupart des options ne s'appliquant qu'à la réouverture, on croit à tort qu'elles n'ont pas fonctionné. Seuls le titre et l'icône réagissent immédiatement, et encore après `Application.RefreshTitleBar`.

Côté code, l'erreur classique est de tenter de **modifier une propriété inexistante** sans gérer le cas de sa création : le motif « tenter, puis créer si erreur » est indispensable.

Une **confusion fréquente** oppose les options de la base courante (propriétés du fichier) aux options de l'installation (`SetOption`/`GetOption`) : régler les unes là où il faudrait régler les autres mène à des effets absents ou inattendus.

Enfin, le **faux sentiment de sécurité** : croire qu'un volet masqué et des touches désactivées protègent l'application ou ses données. Ces réglages n'ont aucune valeur de sécurité ; seules la compilation ACCDE et la sécurité des données (chapitre 20) protègent réellement.

## Synthèse

Les options de démarrage et le verrouillage de l'interface transforment une base ouverte dans Access en application d'apparence autonome. La configuration — titre, icône, formulaire de démarrage, ruban et menus personnalisés — se règle dans les options de la base courante, mais aussi **par code** via les propriétés de la base de données, selon le motif « créer ou modifier » incontournable. L'orchestration du lancement gagne à être confiée à une fonction unique appelée par `AutoExec`, dans le prolongement de l'écran de chargement.

Le verrouillage empile, de façon cumulative, le masquage du volet de navigation, la désactivation des touches spéciales et la restriction des menus. Mais deux vérités le surplombent : la désactivation du contournement par la touche Maj (`AllowBypassKey`) ne se conçoit **jamais sans porte de secours**, sous peine d'enfermement définitif ; et l'ensemble de ces réglages relève de l'**obscurité, non de la sécurité** — la protection réelle du code et des données passant par la compilation ACCDE et la sécurité des données du chapitre 20, par-dessus lesquelles ces options ne font qu'apporter la dernière touche de présentation.

---

> 🔗 **Pour aller plus loin :** le ruban et les menus rattachés au démarrage relèvent des sections 17.1 et 17.2 ; le formulaire de démarrage et l'écran d'accueil des sections 17.3 et 17.4 ; les propriétés de base de données et leur création de la section 12.8 ; l'objet `Application` (`RefreshTitleBar`, `SetOption`) de la section 4.2 ; la protection réelle du code par compilation ACCDE de la section 20.3 et la sécurité des données du chapitre 20 ; le déploiement via Runtime et l'installation automatisée des sections 21.2 et 21.8 ; et la portée front-end de ces options de l'architecture de la section 15.1.

⏭️ [17.8. Mise en forme conditionnelle par code](/17-interface-utilisateur-avancee/08-mise-en-forme-conditionnelle.md)
