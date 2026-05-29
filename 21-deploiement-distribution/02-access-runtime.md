🔝 Retour au [Sommaire](/SOMMAIRE.md)

# 21.2. Access Runtime — déploiement sans licence complète

Le coût des licences est l'un des principaux freins au déploiement d'une application Access à grande échelle. Si chaque utilisateur devait disposer d'une licence Access complète pour faire tourner votre application, l'addition deviendrait vite prohibitive pour une organisation comptant des dizaines ou des centaines de postes. L'**Access Runtime** répond précisément à ce problème : il s'agit d'une version gratuite et redistribuable d'Access qui permet d'exécuter une application sur un poste dépourvu de licence Access complète.

Cette section présente le moteur Runtime lui-même, ses possibilités et ses limites, ainsi que les contraintes pratiques de son déploiement. La préparation du livrable (ACCDE, mode runtime) ayant été traitée à la section 21.1, on s'intéresse ici à l'environnement d'exécution mis à disposition de l'utilisateur final.

## Qu'est-ce que l'Access Runtime ?

L'Access Runtime est un moteur d'exécution fourni gratuitement par Microsoft. Il contient tout ce qui est nécessaire pour **faire fonctionner** une application Access — le moteur de base de données ACE, l'interface d'exécution des formulaires et des états — mais il est dépourvu de tout outil de **conception et de développement**.

Concrètement, un poste équipé du seul Runtime peut ouvrir et utiliser normalement une application (`.accdb`, `.accde` ou `.accdr`) : saisir des données, naviguer dans les formulaires, imprimer des états, exécuter le code VBA. En revanche, il est impossible d'y créer une nouvelle base, de modifier la conception d'un objet, d'éditer du code ou d'inspecter la structure des tables. Le Runtime ouvre par ailleurs **toujours** l'application en mode runtime : le volet de navigation et les rubans de conception sont absents par construction.

Le Runtime est, en somme, un lecteur-exécuteur d'applications Access, comparable dans son esprit à un lecteur PDF gratuit qui permet de consulter un document sans pouvoir le créer ni le modifier.

## Runtime mode et moteur Runtime : ne pas confondre

La distinction a déjà été esquissée à la section 21.1, mais elle mérite d'être rappelée car elle est au cœur des malentendus sur le sujet :

- Le **mode runtime** est un *comportement d'affichage* dans lequel les surfaces de conception sont masquées. Une installation **complète** d'Access peut elle aussi entrer dans ce mode (via l'extension `.accdr` ou le commutateur `/runtime`).
- Le **moteur Runtime** est un *produit distinct*, une version allégée et gratuite d'Access, qui n'autorise que l'exécution et qui fonctionne *exclusivement* en mode runtime.

Autrement dit : tout poste équipé du moteur Runtime s'exécute en mode runtime, mais tout poste s'exécutant en mode runtime n'est pas nécessairement équipé du moteur Runtime. Un développeur disposant d'Access complet peut tester son application en mode runtime sans jamais installer le moteur Runtime.

## L'avantage déterminant : un déploiement gratuit et sans licence

L'intérêt majeur du Runtime est économique et juridique : il est **gratuit** et **redistribuable sans redevance**. Vous pouvez le télécharger auprès de Microsoft et le diffuser librement avec votre application sur autant de postes que nécessaire.

La conséquence sur le modèle de déploiement est considérable : **seul le développeur a besoin d'une licence Access complète** (pour concevoir, coder et générer le livrable). Les utilisateurs finaux, eux, n'ont besoin que du Runtime gratuit. Pour une application interne destinée à de nombreux postes, l'économie réalisée est substantielle.

La contrepartie est tout aussi nette : le Runtime ne permet pas de développer. Toute modification de l'application doit être réalisée sur le poste du développeur, à partir du fichier source `.accdb`, puis redéployée. Le Runtime est un environnement de consommation, jamais de production de code.

## Ce que le Runtime ne permet pas

Déployer sur le Runtime suppose d'accepter un environnement volontairement restreint. Les limitations principales sont les suivantes :

- **Aucun mode Création ni mode Page** n'est disponible, pour aucun type d'objet (formulaires, états, requêtes, tables, modules).
- **Le volet de navigation est absent**, ainsi que les onglets de ruban liés à la conception.
- **De nombreuses boîtes de dialogue d'erreur natives d'Access sont supprimées ou indisponibles.** C'est le point le plus lourd de conséquences, traité en détail ci-dessous.
- **Certains raccourcis clavier** qui exposeraient l'environnement de conception sont désactivés.
- **Il n'existe aucun « repli » vers l'interface Access** : si votre application ne fournit pas elle-même un moyen de naviguer ou d'accomplir une action, l'utilisateur n'a aucune solution de secours.

Ces restrictions ne sont pas des défauts mais la finalité même du Runtime : offrir un environnement verrouillé et sûr pour l'utilisateur final.

## L'impératif absolu : une gestion d'erreurs robuste

C'est la conséquence la plus importante du déploiement sur le Runtime, et celle que les développeurs sous-estiment le plus souvent.

En environnement de développement, lorsqu'une erreur VBA non gérée survient, Access affiche une boîte de dialogue permettant d'arrêter le code, de le déboguer ou de poursuivre. Cette possibilité **n'existe pas** sous le Runtime. Une erreur non interceptée n'y trouve aucun mécanisme de récupération : le plus souvent, **l'application se ferme brutalement**, sans message explicite, laissant l'utilisateur démuni et donnant l'impression d'un logiciel instable.

Il en découle une règle non négociable : **une application destinée au Runtime doit comporter une gestion d'erreurs exhaustive avant tout déploiement.** Chaque procédure susceptible de générer une erreur doit disposer de son propre gestionnaire, et l'application doit s'appuyer sur un dispositif centralisé qui intercepte, journalise et présente les erreurs de façon contrôlée. Ces techniques — `On Error GoTo`, objet `Err`, module de gestion centralisée des erreurs, journalisation dans une table — font l'objet du chapitre 13, dont la maîtrise est ici un prérequis de fait au déploiement.

Une application qui « pardonne » ses propres erreurs en mode développement peut se révéler totalement inutilisable une fois livrée sur le Runtime si cette discipline n'a pas été respectée.

## Fournir toute l'interface : l'application doit être autonome

L'absence de tout repli vers l'interface Access impose une seconde exigence : **l'application doit fournir par elle-même l'intégralité de sa navigation et de ses fonctions.**

L'utilisateur du Runtime ne disposant ni du volet de navigation, ni des rubans standard, c'est à vous de mettre à sa disposition tous les points d'entrée : un menu principal ou un formulaire de navigation pour accéder aux différentes fonctionnalités, un ruban personnalisé si nécessaire, et des options de démarrage cohérentes. Ces éléments relèvent du chapitre 17, en particulier du ruban personnalisé (section 17.1), du formulaire de navigation (section 17.3) et des options de démarrage (section 17.7).

Concevoir une application « comme si l'environnement Access n'existait pas » est la bonne posture mentale : tout ce qui n'est pas explicitement fourni par votre application sera, pour l'utilisateur du Runtime, tout simplement inaccessible.

## Obtenir et installer le Runtime

### Téléchargement

L'Access Runtime se télécharge gratuitement depuis le centre de téléchargement officiel de Microsoft. Microsoft publie un Runtime distinct pour chaque version d'Access ; pour les versions Microsoft 365 et récentes, il prend la forme d'un installeur Click-to-Run. Il convient de récupérer la version correspondant à votre cible de déploiement.

### Correspondance de version et de bitness

Le choix de la version du Runtime n'est pas libre : il doit être cohérent avec le livrable que vous distribuez.

La **version** du Runtime doit être compatible avec celle utilisée pour générer l'ACCDE. La règle de prudence consiste à déployer un Runtime de version au moins égale à celle de génération du livrable, en gardant à l'esprit les recommandations de la section 21.6 sur la compatibilité entre versions d'Access.

La **bitness** (32 ou 64 bits) du Runtime doit impérativement correspondre à celle de l'ACCDE. Un ACCDE 32 bits exige un Runtime 32 bits, et réciproquement. Cette contrainte prolonge celle déjà évoquée pour l'ACCDE à la section 21.1, et s'inscrit dans la problématique plus large des différences 32/64 bits traitée à la section 21.7.

### Coexistence avec une installation complète d'Office

Un point de friction fréquent concerne la cohabitation du Runtime avec une installation Office existante.

D'une part, **un poste déjà équipé d'une version complète et compatible d'Access n'a pas besoin du Runtime** : il peut exécuter votre application directement. Installer le Runtime sur un tel poste est non seulement inutile, mais potentiellement source de conflits.

D'autre part, les règles de cohabitation des produits Office s'appliquent : deux produits Office de bitness différentes (32 et 64 bits) ne cohabitent généralement pas sur une même machine, pas plus que des installations reposant sur des technologies différentes (MSI et Click-to-Run). Tenter d'installer un Runtime dont la bitness ou la technologie entre en conflit avec un Office déjà présent conduit à des échecs d'installation ou à des comportements erratiques.

La conséquence pratique est qu'il faut **réserver l'installation du Runtime aux postes qui ne disposent pas déjà d'un Access complet compatible**, et veiller à la cohérence de bitness avec l'environnement existant.

## Déterminer si un poste a besoin du Runtime

La logique de déploiement découle directement de ce qui précède. Avant d'installer quoi que ce soit, il s'agit de déterminer l'état du poste cible :

- Si le poste dispose déjà d'une version **complète** et **compatible** d'Access (version et bitness adéquates), aucun Runtime n'est nécessaire ; l'application peut être déployée et lancée directement, de préférence en mode runtime via `.accdr` ou le commutateur `/runtime`.
- Si le poste ne dispose d'**aucun** Access, ou d'une version incompatible, il faut installer le **Runtime** correspondant avant de déployer l'application.

Cette détection peut être automatisée lors de l'installation, en s'appuyant notamment sur l'inspection du registre Windows pour repérer une installation Access existante (lecture du registre traitée à la section 22.9). L'industrialisation complète de ce processus relève de l'installation automatisée (section 21.8).

## Tester son application en conditions Runtime

Le commutateur `/runtime` permet de simuler le mode runtime sur le poste du développeur et constitue un premier filet de sécurité utile pour repérer les comportements liés à l'absence des surfaces de conception. Il ne reproduit toutefois pas tout : certains problèmes ne se manifestent que sous le **véritable** moteur Runtime, notamment des références manquantes que le poste de développement masque par sa propre installation complète, ou des comportements d'erreur que seul le Runtime révèle.

La bonne pratique consiste donc à valider l'application sur un poste — physique ou virtuel — équipé du seul Runtime, dans une configuration aussi proche que possible de celle des utilisateurs finaux. C'est le seul moyen de garantir qu'aucune dépendance présente uniquement sur le poste de développement ne viendra manquer en production.

## Le Runtime dans la chaîne de déploiement

Le moteur Runtime est une pièce du dispositif d'ensemble du déploiement, et s'articule avec les autres briques du chapitre :

- Le livrable exécuté par le Runtime est l'**ACCDE** préparé selon les principes de la section 21.1.
- La diffusion des nouvelles versions de ce livrable relève des **stratégies de mise à jour du front-end** (section 21.3).
- L'installation conjointe du Runtime et de l'application sur les postes, raccourcis et paramètres de lancement compris, est traitée par l'**installation automatisée** (section 21.8).

## Points de vigilance

Plusieurs erreurs reviennent fréquemment lors d'un déploiement sur le Runtime :

- **Négliger la gestion d'erreurs.** C'est l'écueil majeur : une application sans gestion d'erreurs exhaustive se ferme brutalement à la première erreur non interceptée, sans message pour l'utilisateur. Le chapitre 13 est un prérequis.
- **Supposer que l'interface Access servira de repli.** Sous le Runtime, rien n'est fourni en dehors de ce que contient l'application : toute la navigation doit être prévue (chapitre 17).
- **Installer un Runtime de bitness incompatible** avec l'ACCDE ou avec un Office déjà présent sur le poste, provoquant échecs ou conflits.
- **Installer le Runtime sur un poste disposant déjà d'un Access complet compatible**, ce qui est inutile et risque de créer des conflits.
- **Ne tester qu'avec `/runtime` sur le poste de développement**, sans valider sur un poste réellement équipé du seul Runtime, et découvrir trop tard des dépendances manquantes en production.

⏭️ [21.3. Mise à jour du front-end — stratégies d'auto-update](/21-deploiement-distribution/03-mise-a-jour-front-end.md)
