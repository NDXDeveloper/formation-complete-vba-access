🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.3. Formulaire de navigation (Navigation Form) contrôlé par VBA

Le formulaire de navigation est un type de formulaire spécial, introduit avec Access 2010, qui fournit clés en main une interface de type « tableau de bord » : une série de boutons disposés en onglets qui, lorsqu'on les clique, chargent un formulaire ou un état dans une zone d'affichage. Il constitue la réponse intégrée d'Access au besoin d'un point d'entrée unique d'où l'utilisateur atteint chaque fonction de l'application — ce qu'on appelait autrefois un « menu général » (*switchboard*).

Son principal attrait tient à sa simplicité : en mode création, il suffit de faire glisser des formulaires sur les boutons pour obtenir une navigation fonctionnelle, sans écrire une seule ligne de code. Mais cette facilité a une contrepartie : le formulaire de navigation par défaut est rigide, et dès qu'on veut le piloter finement — naviguer par programmation, lire ou alimenter le formulaire chargé, passer des paramètres, adapter les boutons aux droits de l'utilisateur — il faut maîtriser un modèle objet dont l'accès est notoirement déroutant. C'est ce contrôle par VBA qui fait l'objet de cette section.

## Anatomie d'un formulaire de navigation

Un formulaire de navigation s'articule autour de trois éléments imbriqués.

Le **`NavigationControl`** est le conteneur des boutons de navigation. C'est lui qui matérialise la barre d'onglets, dans la disposition choisie.

Les **`NavigationButton`** sont les boutons individuels de cette barre. Chacun connaît, via sa propriété `NavigationTargetName`, le nom du formulaire ou de l'état qu'il doit charger. Ce sont des contrôles à part entière, dotés des propriétés habituelles : `Caption`, `Visible`, `Enabled`.

Le **`NavigationSubform`** est un contrôle de sous-formulaire — une zone d'affichage — dans laquelle la cible du bouton cliqué vient se charger. Quand l'utilisateur change d'onglet, c'est en réalité la propriété `SourceObject` de ce sous-formulaire qui est réaffectée. **Le nom de ce contrôle, « NavigationSubform », est celui attribué par défaut**, mais il peut différer : comme tout le code de pilotage en dépend, il faut impérativement vérifier et, idéalement, standardiser ce nom dès la création.

Plusieurs **dispositions** sont proposées : onglets horizontaux (en haut), onglets verticaux (à gauche ou à droite), et des variantes à deux niveaux combinant des onglets parents et des sous-onglets. Le choix est purement ergonomique et n'a pas d'incidence sur le pilotage VBA.

## Création : un modèle, pas du code pur

Un point pratique conditionne toute la suite : **le formulaire de navigation se crée en mode conception, pas par programmation**. On le génère via le ruban (**Créer → Navigation**, puis choix de la disposition), ce qui produit un formulaire contenant le `NavigationControl`. On y associe ensuite les cibles en faisant glisser les formulaires et états depuis le volet de navigation sur les emplacements de boutons, ou en renseignant leur propriété `NavigationTargetName`.

Le `NavigationControl` est un composant complexe qu'on ne reconstruit pas aisément par code ; l'approche réaliste consiste donc à **bâtir la structure en conception, puis à la piloter par VBA à l'exécution**. C'est exactement ce que reflète l'intitulé de cette section : il ne s'agit pas de créer le formulaire de navigation par code, mais de le commander une fois en place.

## Le défi central : référencer le formulaire chargé

Voici la difficulté qui déroute la plupart des développeurs. Le formulaire affiché à l'écran n'est pas le formulaire de navigation lui-même, mais un formulaire **chargé à l'intérieur** de son `NavigationSubform`. Pour accéder à ses contrôles, il faut donc traverser cette imbrication, exactement comme pour un sous-formulaire classique (section 6.4), via la propriété `.Form` du contrôle sous-formulaire :

```vba
' Depuis le formulaire de navigation (l'hôte) :
' accéder au formulaire actuellement chargé dans la sous-zone.
Private Sub ExempleAccesPageCourante()
    Dim frm As Access.Form

    On Error Resume Next            ' rien n'est peut-être chargé
    Set frm = Me.NavigationSubform.Form
    On Error GoTo 0

    If frm Is Nothing Then
        MsgBox "Aucune page n'est actuellement chargée."
    Else
        ' On manipule alors le formulaire chargé comme n'importe quel autre.
        MsgBox "Page : " & Me.NavigationSubform.SourceObject & vbCrLf & _
               "Champ Nom : " & frm!txtNom
    End If
End Sub
```

Trois précautions ressortent de cet exemple. D'abord, le `On Error Resume Next` est nécessaire car l'accès à `.Form` échoue lorsque aucune cible n'est chargée. Ensuite, la propriété `SourceObject` du sous-formulaire renvoie le nom de la cible sous une forme préfixée, du type `"Form.frmClients"`. Enfin — et c'est le piège le plus courant — toute cette chaîne repose sur le nom exact du contrôle sous-formulaire (`NavigationSubform` ici) : une faute de frappe ou un nom non standardisé fait silencieusement échouer l'ensemble.

Pour isoler cette fragilité, il est judicieux d'encapsuler la lecture de la page courante dans une fonction utilitaire du formulaire de navigation :

```vba
' Renvoie le nom du formulaire chargé (sans le préfixe), ou "" si aucun.
Public Function PageCourante() As String
    Dim src As String
    src = Me.NavigationSubform.SourceObject       ' ex. "Form.frmClients"
    If Len(src) > 0 And InStr(src, ".") > 0 Then
        PageCourante = Mid$(src, InStr(src, ".") + 1)
    End If
End Function
```

## Naviguer par programmation

Charger une page précise par code se fait en réaffectant le `SourceObject` du sous-formulaire :

```vba
' Charge une page précise (méthode directe).
Public Sub AllerVersPage(ByVal nomFormulaire As String)
    Me.NavigationSubform.SourceObject = "Form." & nomFormulaire
End Sub
```

Cette méthode est simple et fiable pour *afficher* la cible, mais elle présente une limite à connaître : **elle ne met pas à jour la mise en évidence du bouton de navigation correspondant**. L'utilisateur voit le bon formulaire, mais l'onglet « actif » de la barre ne suit pas, ce qui peut prêter à confusion. Pour préserver la cohérence visuelle de la barre d'onglets, l'alternative consiste à donner le focus au bouton de navigation lui-même, ce qui déclenche le chargement *et* l'activation visuelle du bouton :

```vba
' Navigation préservant la mise en évidence du bouton.
Me!btnClients.SetFocus      ' où btnClients est le NavigationButton ciblé
```

Le choix entre les deux approches dépend du besoin : la réaffectation de `SourceObject` pour un simple chargement programmatique discret, le passage par le focus du bouton lorsque la synchronisation de l'onglet actif importe. Ce décalage entre chargement et mise en évidence est une caractéristique connue du composant, qu'il vaut mieux anticiper que découvrir.

## Communication entre l'hôte et la page chargée

La navigation soulève inévitablement la question des échanges entre le formulaire hôte et les formulaires qu'il héberge. Deux directions sont à considérer.

De **l'hôte vers la page**, l'accès se fait par `Me.NavigationSubform.Form`, comme vu plus haut : l'hôte lit et écrit les contrôles du formulaire chargé.

De **la page vers l'hôte**, le mécanisme repose sur la propriété `Parent`. Depuis le code d'un formulaire chargé dans la sous-zone, `Me.Parent` renvoie le **formulaire de navigation hôte**. La page peut ainsi appeler les méthodes publiques que l'hôte expose, ce qui ouvre la voie à une navigation déclenchée depuis une page :

```vba
' DANS un formulaire chargé (ex. frmClients) :
' un bouton qui demande à l'hôte de basculer vers une autre page.
Private Sub cmdVoirCommandes_Click()
    On Error Resume Next
    ' Me.Parent = le formulaire de navigation hôte ;
    ' AllerVersPage est une procédure Public de son module.
    Me.Parent.AllerVersPage "frmCommandes"
End Sub
```

Ce schéma — l'hôte exposant des procédures publiques que les pages invoquent via `Me.Parent` — est le moyen propre de faire dialoguer les deux niveaux sans couplage rigide. Il rejoint les principes de communication entre formulaires de la section 6.9 : la page ne connaît pas le détail de l'hôte, elle se contente d'appeler un service public bien défini.

## Passer des paramètres à une page

Un manque du formulaire de navigation déroute souvent : **on ne peut pas transmettre d'`OpenArgs` à la cible**. Le mécanisme habituel de passage de paramètres à l'ouverture d'un formulaire (la propriété `OpenArgs`, section 6.6) repose sur `DoCmd.OpenForm` ; or la navigation ne passe pas par `DoCmd.OpenForm` mais par la réaffectation de `SourceObject`, qui n'offre aucun équivalent.

La solution idiomatique consiste à recourir aux **variables de session `TempVars`** (section 15.7) : l'appelant y dépose le contexte juste avant de naviguer, et la page le lit à son chargement.

```vba
' --- Sur le formulaire de navigation ---
Public Sub AllerVersFicheClient(ByVal idClient As Long)
    TempVars!ClientCourant = idClient            ' on dépose le contexte
    Me.NavigationSubform.SourceObject = "Form.frmClient"
End Sub
```

```vba
' --- DANS frmClient (la page chargée) ---
Private Sub Form_Load()
    If Not IsNull(TempVars!ClientCourant) Then
        ' La page exploite le contexte déposé par l'appelant.
        Me.RecordSource = "SELECT * FROM tblClients WHERE ID = " & _
                          CLng(TempVars!ClientCourant)
    End If
End Sub
```

Une variante consiste à laisser l'hôte alimenter directement les contrôles de la page *après* l'avoir chargée, en accédant à `Me.NavigationSubform.Form`. Mais l'approche par `TempVars` reste la plus robuste, car elle ne dépend ni de l'ordre exact des événements ni du nom du sous-formulaire.

## Piloter les boutons dynamiquement

Le pilotage par VBA permet d'adapter la barre de navigation au contexte ou au profil de l'utilisateur. Les `NavigationButton` exposant les propriétés `Visible`, `Enabled` et `Caption`, on peut masquer un onglet réservé aux administrateurs, désactiver une fonction temporairement indisponible, ou renommer un bouton à la volée :

```vba
' Adapter la barre de navigation aux droits de l'utilisateur, au chargement.
Private Sub Form_Load()
    ' Masquer l'onglet d'administration aux utilisateurs non habilités
    ' (gestion des droits applicatifs, section 20.6).
    If Not UtilisateurEstAdmin() Then
        Me!btnAdministration.Visible = False
    End If

    ' Personnaliser un libellé selon le contexte.
    Me!btnAccueil.Caption = "Bonjour " & TempVars!NomUtilisateur
End Sub
```

Cette capacité à moduler l'interface selon le profil connecté est l'un des apports concrets du pilotage VBA, impossible avec un formulaire de navigation purement déclaratif.

## Réagir aux changements de page

Le formulaire de navigation n'expose pas d'événement clair et documenté signalant « l'utilisateur a changé d'onglet » au niveau de l'hôte. La détection d'un changement de page passe donc, en pratique, par les **événements du formulaire chargé** : son événement `Load` (déclenché à chaque chargement) ou `Activate` (section 8.7) marque son apparition à l'écran. C'est dans ces événements que la page initialise son contenu, lit ses paramètres ou prévient l'hôte de son arrivée via `Me.Parent`.

Cette absence d'événement hôte centralisé est une contrainte structurante : la logique « ce qui se passe quand telle page s'affiche » se loge naturellement dans chaque page, et non dans le formulaire de navigation. Concevoir l'application en conséquence — chaque page responsable de sa propre initialisation — évite de chercher en vain un point d'accroche côté hôte.

## Le rechargement à chaque navigation

Un comportement important doit être intégré dès la conception : **chaque navigation recharge la page cible**. Lorsqu'on quitte un onglet puis qu'on y revient, le formulaire est rechargé à neuf, et non restauré dans l'état où on l'avait laissé. Cela a deux conséquences directes.

D'une part, **l'état n'est pas préservé** : une saisie en cours non enregistrée, une position dans une liste, un filtre appliqué manuellement disparaissent au changement d'onglet. Si la conservation de cet état importe, c'est au développeur de la prendre en charge — par exemple en mémorisant les valeurs dans des `TempVars` et en les restaurant au `Load` de la page.

D'autre part, le rechargement systématique a un **coût de performance** pour les pages lourdes (gros jeux d'enregistrements, nombreux contrôles). Les techniques d'optimisation des formulaires — limitation du `RecordSource`, chargement différé — abordées au chapitre 18 (section 18.5) prennent ici tout leur intérêt.

## Définir le formulaire de navigation au démarrage

Pour qu'il serve effectivement de point d'entrée, le formulaire de navigation s'affiche au lancement de l'application via l'option **Fichier → Options → Base de données active → « Formulaire à afficher »**. Combinée au masquage du volet de navigation et au verrouillage de l'interface (section 17.7), cette configuration fait du formulaire de navigation l'unique porte d'entrée visible, conférant à l'application son allure de logiciel autonome.

## Formulaire de navigation ou menu général maison ?

Le formulaire de navigation n'est pas l'unique manière de structurer la navigation d'une application, et il convient d'en peser le choix avec lucidité. Son alternative est le **menu général construit à la main** : un formulaire principal hébergeant lui-même un contrôle sous-formulaire ordinaire, dont on pilote le `SourceObject` à l'aide de ses propres boutons de commande.

Le formulaire de navigation l'emporte par sa **rapidité de mise en œuvre** et la **cohérence visuelle immédiate** de sa barre d'onglets, sans code pour les cas simples. En contrepartie, il impose sa structure, son apparence est peu personnalisable, l'accès au formulaire chargé est verbeux, et le décalage de mise en évidence évoqué plus haut peut gêner.

Le menu général maison demande **davantage de code** et le soin de gérer soi-même la mise en page et la mise en évidence du bouton actif, mais il offre en retour un **contrôle VBA total** : maîtrise complète des événements, liberté de style, conservation possible de l'état des pages, intégration de barres de progression ou de transitions personnalisées. Il faut noter qu'il ne supprime pas pour autant la limitation de l'`OpenArgs` : comme le formulaire de navigation, il charge ses cibles via `SourceObject` et recourt donc aux mêmes `TempVars` pour le passage de paramètres.

En synthèse, le formulaire de navigation convient aux applications où la rapidité prime et où la navigation reste standard ; le menu général maison s'impose dès que l'on recherche une personnalisation poussée ou un contrôle VBA fin. Beaucoup de développeurs expérimentés privilégient d'ailleurs la seconde voie pour les applications ambitieuses, précisément pour la liberté qu'elle procure.

## Pièges courants

Le piège numéro un est le **nom du contrôle sous-formulaire** : tout le code de pilotage dépend de `NavigationSubform`, et un nom différent ou mal orthographié fait échouer silencieusement chaque accès. La première chose à faire après création est de vérifier ce nom dans la feuille de propriétés.

Vient ensuite l'**accès à `.Form` sans protection** : tenter de lire `Me.NavigationSubform.Form` alors qu'aucune page n'est chargée provoque une erreur ; d'où la nécessité d'un `On Error Resume Next` ou d'un test préalable.

Le **décalage de mise en évidence** lors d'une navigation par `SourceObject` surprend régulièrement : la page change mais l'onglet actif reste l'ancien. On le contourne en navigant par le focus du bouton lorsque la synchronisation visuelle est requise.

La **tentation de passer un `OpenArgs`** se solde par un échec, faute de support : il faut adopter les `TempVars` dès le départ pour le passage de contexte.

Enfin, **compter sur la persistance de l'état** d'une page entre deux visites mène à des comportements inattendus, puisque chaque navigation recharge la cible : la conservation d'état, si elle est nécessaire, doit être explicitement programmée.

## Synthèse

Le formulaire de navigation offre une interface de tableau de bord clés en main, où des `NavigationButton` chargent des formulaires dans un contrôle `NavigationSubform`. Sa simplicité de création en conception contraste avec la subtilité de son pilotage par VBA, dont la clé est de comprendre que la page visible est un formulaire **imbriqué** : on y accède par `Me.NavigationSubform.Form`, et la page communique en retour avec l'hôte par `Me.Parent`. La navigation programmatique s'effectue en réaffectant `SourceObject` (au prix de la mise en évidence du bouton), le passage de paramètres par les `TempVars` faute d'`OpenArgs`, et l'adaptation de la barre aux droits par les propriétés `Visible`/`Enabled`/`Caption` des boutons.

Deux réalités structurent la conception : l'absence d'événement hôte signalant un changement de page (la logique d'initialisation se loge donc dans chaque page) et le rechargement systématique de la cible à chaque navigation (l'état n'est pas conservé sans intervention). Lorsque ces contraintes pèsent trop, le menu général construit à la main, plus exigeant en code mais offrant un contrôle VBA total, constitue l'alternative de référence. Dans les deux cas, ce point d'entrée s'affiche au démarrage et s'intègre au verrouillage de l'interface traité à la section suivante.

---

> 🔗 **Pour aller plus loin :** l'imbrication du formulaire chargé relève des sous-formulaires (section 6.4) ; le passage de contexte mobilise les `TempVars` (section 15.7) faute d'`OpenArgs` (section 6.6) ; la communication entre niveaux prolonge la section 6.9 ; la réaction aux affichages s'appuie sur les événements `Load`/`Activate` (section 8.7) ; l'adaptation aux droits renvoie à la section 20.6 ; l'optimisation des pages rechargées au chapitre 18 (section 18.5) ; et l'affichage au démarrage ainsi que le verrouillage de l'interface sont traités à la section 17.7.

⏭️ [17.4. Splash screen et écran de chargement](/17-interface-utilisateur-avancee/04-splash-screen.md)
