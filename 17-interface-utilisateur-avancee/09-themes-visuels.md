🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 17.9. Thèmes visuels et personnalisation cohérente de l'application

Tout au long de ce chapitre, l'application s'est dotée d'un ruban, de menus, d'une navigation, d'un écran d'accueil, de barres de progression et de raccourcis clavier. Cette dernière section assure leur **cohérence visuelle** : faire en sorte que tous les écrans partagent une même identité — mêmes couleurs, mêmes polices, mêmes styles de contrôles — au lieu de former un patchwork hétéroclite où chaque formulaire trahit la date à laquelle il a été conçu. C'est la finition qui transforme une collection de formulaires fonctionnels en un produit unifié.

Le mot « cohérente » est ici le maître mot. Une application Access devient presque inévitablement disparate avec le temps : un formulaire arbore une police, un autre une couleur de bouton différente, un troisième un alignement particulier. La cohérence ne s'obtient pas en retouchant chaque écran un par un — démarche vouée à l'oubli et à l'incohérence — mais en **centralisant les décisions visuelles** et en les appliquant systématiquement. C'est précisément ce que permet le pilotage par code, sujet principal de cette section. Une mise au point honnête s'impose toutefois d'emblée : Access n'est pas un outil de design, et ses possibilités stylistiques sont bornées ; l'objectif réaliste est un rendu **propre et homogène**, non un habillage spectaculaire.

## Les fondations en conception : modèles et propriétés par défaut

Avant tout recours au code, Access offre des mécanismes de conception qui posent la base visuelle sans effort de programmation. On peut définir un **modèle de formulaire** et un **modèle d'état** (dans les options des concepteurs d'objets) : tout nouveau formulaire ou état créé hérite alors des propriétés de ce modèle — couleurs de section, polices, styles. De même, après avoir mis en forme un contrôle d'une certaine manière, on peut le **définir comme valeur par défaut** pour ce type de contrôle, de sorte que les contrôles ajoutés ensuite reprennent automatiquement cette apparence.

Ces dispositifs de conception constituent la **ligne de base** de la cohérence et présentent un avantage majeur : étant appliqués à la conception, ils n'entraînent **aucun coût à l'exécution**. Ils ne suffisent cependant pas dès qu'on souhaite faire évoluer le thème globalement ou le changer dynamiquement — c'est là qu'intervient le code.

## Les thèmes Office intégrés

Access prend en charge, comme les autres applications Office, la notion de **thème Office** : un jeu coordonné de couleurs et de polices de thème. Les contrôles peuvent référencer une **couleur de thème** plutôt qu'une couleur fixe ; changer le thème du document recolore alors l'ensemble de façon cohérente.

En pratique, ce mécanisme se révèle **moins souple et plus capricieux en Access** que dans Word ou Excel, et son comportement n'est pas toujours intuitif. Pour une maîtrise fiable et portable de l'identité visuelle, l'approche par **palette centralisée en code**, décrite ci-dessous, donne de meilleurs résultats et s'avère plus prévisible — c'est elle que privilégie cette section.

## Centraliser la palette et les polices

Voici le cœur de l'angle VBA. Le principe fondateur de toute cohérence est l'**unicité de la source** : les couleurs et les polices de l'application ne doivent être définies **qu'à un seul endroit**, et non éparpillées en valeurs `RGB` littérales d'un formulaire à l'autre. Lorsqu'une couleur est répétée dans trente formulaires, en changer suppose trente modifications — et l'on en oublie toujours. Définie en un point unique, elle se change une fois pour toutes.

Cette centralisation se décline selon trois niveaux d'ambition. Le plus simple est un **module de constantes** rassemblant la palette et les polices :

```vba
' ============================================
' Module standard : modTheme
' Source unique de l'identité visuelle de l'application.
' ============================================
Option Compare Database
Option Explicit

' --- Palette ---
Public Const CLR_PRIMAIRE As Long = 10510624     ' bleu (valeur Long)
Public Const CLR_FOND As Long = 16448250          ' gris très clair
Public Const CLR_TEXTE As Long = 3355443          ' gris foncé
Public Const CLR_ACCENT As Long = 2003199         ' orange d'accentuation

' --- Typographie ---
Public Const POLICE_NOM As String = "Segoe UI"
Public Const POLICE_TAILLE As Integer = 10
Public Const POLICE_TITRE_TAILLE As Integer = 16
```

Pour les applications plus ambitieuses, on encapsule ces valeurs dans une **classe de thème** (`clsTheme`), exposant la palette par des propriétés (chapitre 16) — ce qui permet de charger différents thèmes et d'en changer à l'exécution. Une instance unique de cette classe, fournie par un accesseur global selon le patron Singleton (section 16.7), devient alors la référence consultée partout. Enfin, le niveau le plus souple stocke les valeurs du thème dans une **table de configuration**, lue à l'exécution — voie détaillée plus bas, qui autorise la modification du thème sans recompilation.

## Appliquer le thème par code

Une fois la source centralisée, une procédure unique applique le thème à un formulaire en **parcourant sa collection de contrôles** et en stylant chacun selon son type. Cette routine, appelée depuis l'événement `Form_Load` de chaque formulaire, garantit qu'aucun écran n'échappe à l'identité commune :

```vba
Public Sub AppliquerTheme(ByRef frm As Access.Form)
    Dim ctl As Access.Control

    ' Tolère les propriétés non supportées par certains types de contrôle.
    On Error Resume Next

    ' Fond de la section Détail.
    frm.Section(acDetail).BackColor = CLR_FOND

    For Each ctl In frm.Controls
        Select Case ctl.ControlType
            Case acCommandButton
                ctl.BackColor = CLR_PRIMAIRE
                ctl.ForeColor = vbWhite
                ctl.FontName = POLICE_NOM
                ctl.FontSize = POLICE_TAILLE

            Case acLabel
                ctl.ForeColor = CLR_TEXTE
                ctl.FontName = POLICE_NOM

            Case acTextBox, acComboBox, acListBox
                ctl.BackColor = vbWhite
                ctl.ForeColor = CLR_TEXTE
                ctl.FontName = POLICE_NOM
                ctl.FontSize = POLICE_TAILLE
        End Select
    Next ctl

    On Error GoTo 0
End Sub
```

Deux points appellent un commentaire. Le `On Error Resume Next` est ici **nécessaire et non négligent** : toutes les propriétés ne s'appliquent pas à tous les types de contrôle, et tenter de fixer une couleur de fond sur un contrôle qui n'en possède pas déclencherait une erreur ; le `Select Case` par type limite déjà ce risque, la tolérance d'erreur le couvre entièrement. Ce parcours de contrôles relève par ailleurs des techniques de manipulation dynamique de la section 17.11.

Un rapprochement éclairant s'impose avec la section précédente. En 17.8, le fait qu'une propriété posée par code sur un formulaire continu colore **toutes** les lignes était un piège. Pour le thème, ce même comportement est exactement ce que l'on recherche : on *veut* que toutes les lignes partagent la même apparence. Le mécanisme identique est donc un défaut ou une qualité selon l'intention — uniformité souhaitée pour le thème, distinction par ligne recherchée pour la mise en forme conditionnelle.

## Le thème par les données : changer sans recompiler

Le niveau le plus souple consiste à stocker les valeurs du thème dans une **table de configuration** et à les lire à l'exécution. L'intérêt est double : un administrateur peut alors **modifier l'identité visuelle sans rouvrir l'application en conception** ni recompiler, et l'on peut servir une **charte différente selon le client** ou le déploiement à partir d'un simple enregistrement.

```vba
' Charger une couleur depuis la table de configuration du thème.
Dim clrPrimaire As Long
clrPrimaire = Nz(DLookup("Valeur", "tblTheme", _
                        "Cle = 'CouleurPrimaire'"), CLR_PRIMAIRE)
```

On stocke alors les couleurs sous forme de valeurs `Long` (ou de composantes RGB) dans la table, lues via les fonctions de domaine (section 11.10), avec une valeur de repli en cas d'absence. Combinée à la classe de thème évoquée plus haut — qui charge la palette depuis cette table à son initialisation —, cette approche offre une personnalisation entièrement pilotée par les données.

## Basculer de thème à l'exécution

C'est le bénéfice qui justifie tout l'effort de centralisation : pouvoir **changer de thème en cours d'exécution**. Un mode sombre, une charte saisonnière, un habillage propre à chaque client deviennent réalisables, parce que la source des couleurs est unique et que l'application sait se re-styler à la demande. Il suffit de changer la source du thème, puis de **réappliquer le thème à tous les formulaires ouverts** :

```vba
' Réappliquer le thème courant à l'ensemble des formulaires ouverts.
Public Sub RappliquerThemePartout()
    Dim f As Access.Form
    For Each f In Forms              ' la collection Forms = formulaires ouverts
        AppliquerTheme f
    Next f
End Sub
```

La collection `Forms` ne contenant que les formulaires actuellement ouverts, cette boucle les parcourt et leur réapplique l'apparence — un changement de thème instantané et global. Sans la centralisation, un tel basculement serait tout simplement impossible : il faudrait modifier chaque contrôle de chaque formulaire à la main.

## La cohérence au-delà de la couleur

L'identité visuelle ne se réduit pas aux couleurs. La **typographie** y contribue tout autant : une famille de police unique et un jeu cohérent de tailles (titre, corps, étiquette) unifient l'ensemble bien plus que la couleur seule. Un point de vigilance s'impose ici, qui touche au déploiement (chapitre 21) : **une police utilisée doit être installée sur chaque poste utilisateur**, faute de quoi Access lui substitue une police de repli et la cohérence vole en éclats. On s'en tient donc, pour une application distribuée, à des polices universellement présentes (Segoe UI, Calibri, Arial).

La cohérence se joue aussi dans des choix que le code ne pilote pas toujours : des **tailles et un alignement réguliers** des contrôles, un **placement constant** des éléments (les boutons Valider/Annuler toujours dans le même ordre, les étiquettes toujours du même côté), une **terminologie homogène** d'un écran à l'autre, et une **iconographie unifiée**. Ces principes relèvent de la discipline de conception — l'équivalent d'un « système de design » appliqué à Access — et rejoignent les bonnes pratiques d'architecture applicative du chapitre 24. Un raffinement courant consiste à doter chaque formulaire d'un **bandeau d'en-tête identique** (un rectangle coloré et un libellé de titre), qui ancre l'identité visuelle sur tous les écrans.

## Approche hybride et dosage

La meilleure stratégie est le plus souvent **hybride** : poser la ligne de base par les modèles et les propriétés par défaut à la conception — sans coût d'exécution —, et ne recourir au code que pour ce qui doit être **dynamique** (basculement de thème, charte par client lue dans une table). On évite ainsi d'appliquer inutilement le thème par code là où une valeur par défaut de conception suffirait.

Comme toute finition de ce chapitre, la personnalisation visuelle s'applique avec **discernement**. Centraliser la palette dans un module de constantes est peu coûteux et profite même aux petites applications : c'est une bonne pratique quasi inconditionnelle. En revanche, une classe de thème pilotée par une table et un moteur de basculement à l'exécution ne se justifient que pour des applications d'envergure, multi-clients ou à longue durée de vie. Sur le plan des performances, l'application du thème par parcours des contrôles à chaque ouverture de formulaire ajoute un coût négligeable, sauf sur des formulaires comportant des centaines de contrôles — cas où l'on privilégiera les valeurs par défaut de conception.

## Les limites d'Access en matière de style

L'honnêteté commande de fixer des attentes réalistes. Access n'est **pas** un environnement de design moderne : ses contrôles ne gèrent pas nativement les coins arrondis, les dégradés, les ombres portées ou les animations fines des interfaces web et mobiles contemporaines. Certains de ces effets se simulent péniblement à coups d'images ou de superpositions, au prix d'une complexité disproportionnée. Le rendu accessible est celui d'une **application de gestion soignée** — claire, lisible, homogène — et non d'une interface spectaculaire.

Cette limite n'est pas un échec : une application Access cohérente, à la palette maîtrisée et à la typographie régulière, inspire confiance et se révèle agréable à utiliser, ce qui est l'essentiel. Chercher à lui faire imiter une application web aboutirait le plus souvent à un résultat artificiel et fragile. Viser le propre et le cohérent, plutôt que l'esbroufe, est ici la bonne ambition.

## Pièges courants

Le piège fondamental, à l'origine de toute incohérence, est la **dispersion des valeurs visuelles** : des couleurs et polices répétées en littéraux d'un formulaire à l'autre, impossibles à faire évoluer d'ensemble. La centralisation en une source unique est la réponse, et le réflexe à acquérir.

Vient ensuite l'**erreur de propriété non supportée** : appliquer une couleur de fond à un contrôle qui n'en possède pas interrompt le traitement — d'où le `Select Case` par type assorti de la tolérance d'erreur dans la routine d'application.

La **police non installée** sur un poste, substituée silencieusement par Access, ruine la cohérence à l'insu du développeur : seules des polices universelles l'évitent en déploiement.

Lors d'un basculement de thème, l'**oubli de réappliquer** l'apparence aux formulaires déjà ouverts laisse coexister deux thèmes à l'écran ; le parcours de la collection `Forms` règle ce point.

Enfin, deux excès opposés : la **sur-ingénierie** d'un moteur de thème pour une application qui n'en a pas besoin, et la tentation de **forcer Access au-delà de ses capacités** stylistiques, qui produit des résultats laborieux et fragiles plutôt qu'une interface réellement aboutie.

## Synthèse

La cohérence visuelle est la finition qui unifie tous les éléments d'interface du chapitre en un produit homogène. Elle repose sur un principe unique : **centraliser les décisions visuelles** — palette et typographie — en une seule source, qu'il s'agisse d'un module de constantes, d'une classe de thème ou d'une table de configuration, plutôt que de les disperser en valeurs littérales. Une routine d'application parcourt alors les contrôles de chaque formulaire pour leur imposer cette identité, en tolérant les propriétés que certains types ne supportent pas. Lorsque la palette provient d'une table, le thème devient modifiable sans recompilation et adaptable par client ; centralisé, il autorise le **basculement à l'exécution**, en réappliquant l'apparence à tous les formulaires ouverts.

La cohérence dépasse la couleur : typographie homogène (en veillant à n'employer que des polices installées partout), régularité des tailles, des alignements et des placements, terminologie et iconographie unifiées. L'approche la plus saine est hybride — ligne de base posée à la conception, code réservé au dynamique — et dosée selon l'ampleur de l'application. Surtout, elle s'inscrit dans les **limites réelles d'Access**, qui permet un rendu propre et professionnel d'application de gestion, non un habillage spectaculaire. C'est sur cette note que se referme le chapitre : une application qui, au terme de ces neuf sections, ne fonctionne pas seulement, mais se présente, se navigue et s'utilise comme un logiciel pensé pour ceux qui s'en servent.

---

> 🔗 **Pour aller plus loin :** la centralisation par classe et l'accesseur unique s'appuient sur la programmation orientée objet et le patron Singleton (chapitre 16, section 16.7) ; le thème piloté par les données mobilise une table de configuration et les fonctions de domaine (section 11.10) ; le parcours et le stylage des contrôles relèvent de la manipulation dynamique (section 17.11) ; la cohérence des couleurs avec la mise en forme conditionnelle renvoie à la section 17.8 ; la disponibilité des polices et la distribution du thème touchent au déploiement (chapitre 21) ; et la discipline de cohérence rejoint les bonnes pratiques d'architecture du chapitre 24.

⏭️ [18. Optimisation et performance](/18-optimisation-performance/README.md)
