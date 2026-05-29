🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.5. Déploiement en réseau local et gestion des chemins UNC

Le déploiement d'une application Access en architecture front-end / back-end est fondamentalement un déploiement **en réseau local**. C'est le réseau qui héberge les données partagées, accessibles à l'ensemble des utilisateurs, et c'est sur lui que reposent les mécanismes d'auto-update et de reliaison vus aux sections précédentes. La fiabilité de l'ensemble dépend en grande partie d'un détail souvent négligé : la façon dont les chemins d'accès réseau sont exprimés. Privilégier les chemins **UNC** plutôt que les lecteurs mappés est l'une des décisions les plus déterminantes pour la robustesse d'un déploiement.

Cette section décrit la topologie type d'un déploiement réseau, explique la supériorité des chemins UNC, montre comment les manipuler en VBA, et aborde les questions de permissions et de performance propres à l'environnement réseau.

## La topologie d'un déploiement réseau Access

Un déploiement réseau correct répartit les éléments de l'application en trois emplacements distincts, chacun avec un rôle précis :

- Le **back-end** (les données) réside dans un dossier partagé sur le serveur, accessible en lecture et en écriture par tous les utilisateurs. Il est unique et centralisé.
- Le **front-end de référence** (le « master » utilisé par l'auto-update) réside dans un dossier de déploiement sur le réseau, en lecture pour les utilisateurs et en écriture pour l'administrateur qui publie les mises à jour.
- La **copie locale du front-end** réside sur le poste de chaque utilisateur, dans un dossier inscriptible (section 21.3), et c'est elle qui est réellement exécutée.

### La règle d'or : ne jamais exécuter le front-end depuis le réseau

La règle la plus importante du déploiement réseau tient en une phrase : **seul le back-end est partagé ; le front-end s'exécute toujours localement.**

Faire ouvrir directement par tous les utilisateurs un même fichier front-end posé sur le réseau — voire un fichier unique non scindé contenant à la fois données et interface — est une erreur lourde de conséquences. Elle dégrade drastiquement les performances, puisque l'intégralité de l'application (formulaires, états, code) transite alors par le réseau à chaque opération. Surtout, elle multiplie les risques de **corruption** du fichier, particulièrement en accès concurrent. La séparation front-end / back-end et la distribution locale du front-end (sections 21.3 et 21.4) ne sont donc pas de simples commodités : elles conditionnent la stabilité même de l'application en réseau.

Le front-end local minimise le trafic réseau : seules les requêtes de données franchissent le réseau pour atteindre le back-end, et non l'application elle-même.

### Une organisation type des dossiers réseau

Une structure de déploiement claire facilite à la fois l'administration et la configuration des chemins. Une organisation courante ressemble à ceci :

```
\\Serveur\AppPartage\
├── Donnees\
│   └── MonApp_be.accdb        ← back-end (lecture/écriture pour les utilisateurs)
└── Deploiement\
    ├── MonApp.accde            ← front-end de référence (auto-update)
    ├── version.txt             ← version de référence
    └── MAJ_Lanceur.vbs         ← lanceur distribué aux postes
```

Les chemins de ces différents éléments — back-end à relier, master à copier — sont précisément ceux qui doivent être exprimés en UNC dans les chaînes de connexion, les fichiers de configuration et les scripts de lancement.

## Chemins UNC contre lecteurs mappés

Deux façons coexistent pour désigner un emplacement réseau :

- Un **chemin UNC** (Universal Naming Convention) désigne directement le serveur et le partage : `\\Serveur\AppPartage\Donnees\MonApp_be.accdb`. Il est indépendant de toute configuration locale.
- Un **lecteur mappé** associe une lettre au partage : `S:\Donnees\MonApp_be.accdb`, où `S:` est censé pointer vers `\\Serveur\AppPartage`.

### Pourquoi privilégier l'UNC

Les chemins UNC sont nettement plus fiables pour plusieurs raisons concrètes :

- **Les lettres de lecteur varient d'un poste à l'autre.** Un utilisateur mappe `S:`, un autre `M:`, un troisième n'a aucun mappage. Un lien stocké sous `S:\...` se casse sur tout poste où `S:` n'existe pas ou pointe ailleurs. L'UNC, lui, désigne le même emplacement partout.
- **Les lecteurs mappés ne sont pas toujours disponibles au démarrage.** Restaurés à l'ouverture de session, ils peuvent n'être réellement connectés qu'au premier accès, créant des problèmes de disponibilité au lancement de l'application. L'UNC ne dépend d'aucune reconnexion préalable.
- **L'UNC lève toute ambiguïté.** Il identifie sans équivoque le serveur et le partage, ce qui simplifie le diagnostic et la documentation du déploiement.

Pour ces raisons, l'UNC est le format à adopter par défaut dans tout déploiement réseau Access.

### Où l'UNC doit être utilisé

L'usage de l'UNC doit être systématique partout où un chemin réseau est mémorisé ou transmis :

- Dans la chaîne de connexion `Connect` des **tables liées** (section 21.4), afin que la reliaison reste valable sur tous les postes.
- Dans les **fichiers de configuration** ou le **registre** où l'on mémorise le chemin du back-end.
- Dans le **lanceur** et les scripts d'auto-update qui pointent vers le front-end de référence (section 21.3).
- Dans les **raccourcis** et paramètres d'installation (section 21.8).

## Manipuler les chemins UNC en VBA

Les fonctions de manipulation de fichiers de VBA — `Dir`, `FileDateTime`, ainsi que le `FileSystemObject` (section 22.10) — fonctionnent indifféremment sur les chemins UNC et sur les chemins locaux. Aucune adaptation particulière n'est nécessaire pour lire ou copier un fichier désigné en UNC.

La difficulté pratique survient plutôt lors de la **saisie d'un chemin par l'utilisateur**. Lorsqu'un administrateur localise le back-end via la boîte de dialogue de sélection (section 21.4), il désigne souvent un chemin sur lecteur mappé, par exemple `S:\Donnees\MonApp_be.accdb`. Si ce chemin est mémorisé tel quel, il sera inutilisable sur les postes ayant un mappage différent. La bonne pratique consiste à **convertir systématiquement un chemin mappé en son équivalent UNC avant de le mémoriser**.

### Convertir un lecteur mappé en chemin UNC

La conversion s'appuie sur la fonction d'API Windows `WNetGetConnection`, qui retrouve le partage réseau associé à une lettre de lecteur. La déclaration suivante est compatible 64 bits grâce au mot-clé `PtrSafe` ; les conventions complètes de déclaration d'API et la compatibilité 32/64 bits sont détaillées aux sections 22.1 et 21.7 :

```vba
Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" _
    Alias "WNetGetConnectionA" ( _
    ByVal lpLocalName As String, _
    ByVal lpRemoteName As String, _
    ByRef lpnLength As Long) As Long
```

La fonction d'enveloppe ci-dessous renvoie l'équivalent UNC d'un chemin sur lecteur mappé, et laisse inchangé tout chemin déjà exprimé en UNC ou réellement local :

```vba
' Convertit un chemin sur lecteur mappé en chemin UNC.
' Renvoie le chemin inchangé s'il est déjà UNC ou s'il n'est pas mappé.
Public Function CheminVersUNC(ByVal chemin As String) As String
    Dim lettre As String
    Dim reste As String
    Dim distant As String
    Dim longueur As Long

    CheminVersUNC = chemin

    ' Déjà un chemin UNC ?
    If Left$(chemin, 2) = "\\" Then Exit Function
    ' Doit commencer par "X:" pour être un lecteur
    If Mid$(chemin, 2, 1) <> ":" Then Exit Function

    lettre = Left$(chemin, 2)               ' ex. "S:"
    reste = Mid$(chemin, 3)                 ' ex. "\Donnees\MonApp_be.accdb"

    distant = String$(260, vbNullChar)
    longueur = 260

    If WNetGetConnection(lettre, distant, longueur) = 0 Then   ' 0 = succès
        distant = Left$(distant, InStr(distant, vbNullChar) - 1)
        If Len(distant) > 0 Then
            CheminVersUNC = distant & reste
        End If
    End If
End Function
```

Appliquée au chemin retourné par la boîte de dialogue avant mémorisation, cette fonction garantit que l'application stocke toujours une forme UNC exploitable par l'ensemble des postes.

## Permissions et partage réseau

Un déploiement réseau ne fonctionne que si les droits d'accès sont correctement configurés sur les partages, et c'est une source fréquente de dysfonctionnements.

Le **dossier du back-end** doit accorder aux utilisateurs des droits de **modification** : lecture et écriture sur le fichier de données, mais aussi création et suppression de fichiers dans le dossier. Ce dernier point est essentiel et souvent négligé. Accorder un droit de lecture seule sur le back-end conduit à des erreurs dès que l'utilisateur tente d'écrire, ou empêche la création du fichier de verrouillage décrit ci-dessous.

Le **dossier de déploiement** (front-end de référence) n'exige, lui, qu'un droit de **lecture** pour les utilisateurs — qui s'y contentent de copier le master — et un droit d'écriture réservé à l'administrateur chargé de publier les mises à jour.

### Le fichier de verrouillage .laccdb

Lorsqu'une base est ouverte en mode partagé, Access crée un fichier de verrouillage portant l'extension `.laccdb` (ou `.ldb` pour l'ancien format `.mdb`), **dans le même dossier que la base**. Ce fichier enregistre les utilisateurs ayant la base ouverte et joue un rôle central dans la gestion de l'accès concurrent.

Sa création et sa suppression imposent que le dossier du back-end soit **inscriptible** par les utilisateurs : un droit limité à la lecture sur ce dossier empêche la création du fichier de verrouillage et bloque l'ouverture partagée. Le rôle du fichier de verrouillage et la gestion multi-utilisateur sont approfondis au chapitre 15.

## Performance et fiabilité sur le réseau

Access en architecture front-end / back-end est une technologie conçue pour le **réseau local**. Dans ce cadre, plusieurs leviers améliorent la performance et la stabilité.

Le premier est la **persistance de la connexion** au back-end : maintenir une connexion ouverte vers le fichier de données pendant toute la session, plutôt que de l'ouvrir et de la fermer à répétition, réduit la latence et évite la recréation incessante du fichier de verrouillage. Cette technique d'optimisation réseau est détaillée à la section 18.8.

Il faut en revanche être conscient des **limites de l'environnement**. Access n'est pas conçu pour fonctionner de manière fiable sur des liaisons étendues, instables ou à forte latence : un back-end accédé à travers un VPN, une connexion WAN ou une liaison sans fil expose à des lenteurs importantes et à un risque élevé de corruption du fichier. Pour les besoins d'accès distant, les approches recommandées consistent à exécuter l'application sur un serveur de bureau à distance (les utilisateurs distants se connectant à une session proche du back-end), ou à migrer les données vers un véritable serveur de base de données. La migration vers SQL Server et l'architecture hybride correspondante font l'objet du chapitre 23.

## Articulation avec le reste du déploiement

La gestion des chemins réseau irrigue l'ensemble de la chaîne de déploiement :

- L'**auto-update** (section 21.3) copie le front-end de référence désigné par un chemin UNC.
- La **reliaison automatique** (section 21.4) doit cibler le back-end par un chemin UNC pour rester valable sur tous les postes.
- L'**installation automatisée** (section 21.8) configure les raccourcis et les chemins initiaux, là encore en UNC.
- L'**optimisation réseau** par persistance de connexion relève de la section 18.8, et la **gestion multi-utilisateur** du fichier de verrouillage du chapitre 15.
- La conversion de chemin s'appuie sur une **API Windows** (chapitre 22) et doit respecter la compatibilité **32/64 bits** (section 21.7).

## Points de vigilance

Plusieurs erreurs reviennent fréquemment dans un déploiement réseau :

- **Exécuter le front-end depuis le réseau** au lieu de le copier localement, ce qui effondre les performances et favorise la corruption.
- **Stocker des chemins sur lecteur mappé** (`S:\...`) plutôt qu'en UNC, fragilisant liens et configuration d'un poste à l'autre.
- **Mémoriser tel quel un chemin saisi par l'utilisateur** sans le convertir en UNC au préalable.
- **Accorder un droit de lecture seule sur le dossier du back-end**, empêchant l'écriture des données et la création du fichier de verrouillage `.laccdb`.
- **Faire reposer un back-end Access sur une liaison WAN, VPN ou sans fil**, au mépris des limites de la technologie, avec un risque sérieux de corruption.
- **Confondre droits sur le fichier et droits sur le dossier** : c'est bien le dossier qui doit être inscriptible pour permettre la gestion du verrouillage.

⏭️ [21.6. Compatibilité entre versions d'Access (2016, 2019, 2021, M365)](/21-deploiement-distribution/06-compatibilite-versions-access.md)
